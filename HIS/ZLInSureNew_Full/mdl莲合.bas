Attribute VB_Name = "mdl莲合"
Option Explicit
Private mblnInit As Boolean
Public gcurBanlance As Currency                '莲合专用,保存个人帐户余额
Public gintLen As Integer

Public Function 医保设置_莲合() As Boolean
'功能： 该方法用于供相关应用部件调用配置连接医保数据服务器的连接串
'返回：接口配置成功，返回true；否则，返回false
    
    Dim strConn As String
    
    If frmSet成都.ShowSet(TYPE_成都莲合) = False Then
        Exit Function
    End If
    
    strConn = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LHConnectionStrINg"), "dsn=lhyb;uid=sa;pwd=;")
    '重新建立到医保服务器的公共连接
    If gcnSybase.State = adStateClosed Then
        On Error Resume Next
        gcnSybase.Open strConn
        If Err = 0 Then
            医保设置_莲合 = True
        Else
            Err.Clear
        End If
    Else
        医保设置_莲合 = True
    End If

End Function


Public Function 医保初始化_莲合() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false

    '建立到医保服务器的公共连接
    Dim strCnn As String
    
    If mblnInit Then
        医保初始化_莲合 = mblnInit
        Exit Function
    End If
    
    strCnn = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LHConnectionStrINg"), "")
    Err = 0
    On Error Resume Next
    With gcnSybase
        If .State = adStateOpen Then .Close
        .ConnectionString = strCnn
        .Open
        If Err <> 0 Then
            MsgBox "不能建立到医保服务器的连接，无法执行医保交易", vbExclamation, gstrSysName
            Exit Function
        End If
    End With
    gintLen = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("CardNOLength"), 10)
'    '检查联合医保所需的表是否建立
'    gstrSQL = "select * from RCPT_TAB,DIAG_REC "
'    gcnSybase.Execute gstrSQL, 1
'    If Err <> 0 Then
'        MsgBox "RCPT_TAB表和DIAG_REC表没有建立，无法执行医保交易", vbExclamation, gstrSysName
'        Exit Function
'    End If
    
    mblnInit = True
    医保初始化_莲合 = True
End Function


Public Function 身份标识_莲合(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：strSelfNO-个人编号，刷卡得到；strSelfPwd-病人密码；bytType-识别类型，0-门诊，1-住院
'返回：空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmlhIDentified As frmIdentify莲合
     
    Set frmlhIDentified = New frmIdentify莲合
    With frmlhIDentified
        .mlng病人ID = lng病人ID
        .Tag = bytType
        .Show 1
        'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
        身份标识_莲合 = .strPatiInfo
        
        If 身份标识_莲合 <> "" Then
            '建立病人档案信息，传入格式：
            '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
            '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
            '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)

            lng病人ID = BuildPatiInfo(bytType, 身份标识_莲合 & ";;;;;;;;;;;;;;;;", .mlng病人ID, TYPE_成都莲合)
            '返回格式:中间插入病人ID
            身份标识_莲合 = 身份标识_莲合 & ";" & lng病人ID & ";;;;;;;;;;;;;;;;"
        End If
        
    End With
    Unload frmlhIDentified
    
End Function

Public Function 个人余额_莲合() As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额的金额
    个人余额_莲合 = gcurBanlance
End Function

Public Function 门诊结算_莲合(lng结帐ID As Long) As Boolean
'该过程目前未使用，门诊结算时通过调用传输明细达到
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    
    Dim rsPay As New Recordset
    Dim strReptNo As String
    Dim strInterCode As String
    Dim rsList As New ADODB.Recordset
    Dim lngCount As Long, lng病人ID As Long
    
    Dim cur个帐支付 As Currency, cur发生费用 As Currency
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date
On Error GoTo ErrH
    '此时所有收费细目必然有对应的医保编码
    门诊结算_莲合 = False
    
    gstrSQL = _
        "Select NO,登记时间,开单人 as 医生,姓名,病人ID,Sum(结帐金额) as 结帐金额 " & _
        " From 门诊费用记录" & _
        " Where Nvl(附加标志,0)<>9 And Nvl(实收金额,0)<>0 And 结帐ID=[1]" & _
        " Group by NO,登记时间,开单人,姓名,病人ID"
    Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, "莲合医保", lng结帐ID)
    
    With rsList
        If .RecordCount = 0 Then
            Err.Raise 9000, gstrSysName, "没有填写收费记录", vbExclamation, gstrSysName
            Exit Function
        End If

        strReptNo = !NO
        strInterCode = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("intercode"), 713)
        strInterCode = IIf(IsNumeric(strInterCode), strInterCode, "0")
        lng病人ID = !病人ID

        lngCount = 0
        Do While Not .EOF
            lngCount = lngCount + 1
            gstrSQL = "insert into rcpt_tab(hosp_id,rcpt_no,sno,p_name,inter_id,class,amount,doctor_id,r_date,hosp_price)" _
                & " values('0','" & !NO & "'," & lngCount & ",'" & !姓名 & "'," & strInterCode & ",'01'," & !结帐金额 & ",'" & Trim(!医生) & _
                "',to_date('" & Format(!登记时间, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'),1)"
            gcnSybase.Execute gstrSQL
            .MoveNext
        Loop

'        '填写结算表
'        curDate = zlDatabase.Currentdate
'
'        '求个人帐户支付金额
'        gstrSQL = "Select 冲预交 From 病人预交记录 Where 结算方式='个人帐户' And 结帐ID=" & lng结帐ID
'        If .State = adStateOpen Then .Close
'        .Open gstrSQL, gcnOracle, adOpenKeyset
'        If Not .EOF Then cur个帐支付 = IIf(IsNull(!冲预交), 0, !冲预交)
'
'        '帐户年度信息
'        Call Get帐户信息(lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
'
'        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_成都莲合 & "," & Year(curDate) & "," & _
'            cur帐户增加累计 & "," & cur帐户支出累计 + cur个帐支付 & "," & cur进入统筹累计 & "," & _
'            cur统筹报销累计 & "," & int住院次数累计 & ")"
'        gcnOracle.Execute gstrSQL, , adCmdStoredProc
'
'        '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
'        gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_成都莲合 & "," & lng病人ID & "," & _
'            Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
'            cur统筹报销累计 & "," & int住院次数累计 & ",NULL,NULL,NULL," & cur发生费用 & ",NULL,NULL," & _
'            cur个帐支付 & ",NULL)"
'        gcnOracle.Execute gstrSQL, , adCmdStoredProc
'        '---------------------------------------------------------------------------------------------
        门诊结算_莲合 = True
    End With
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 个人帐户转预交_莲合(lng预交ID As Long, curMoney As Currency) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    
End Function

Public Function 入院登记_莲合(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_成都莲合 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "莲合医保")
    
    入院登记_莲合 = True
End Function

Public Function 出院登记_莲合(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_成都莲合 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "莲合医保")
    
    出院登记_莲合 = True
End Function

Public Function 住院虚拟结算_莲合(rsExse As Recordset, strSelfNo As String, strSelfPwd As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim str住院号 As String
    Dim STR姓名 As String
    Dim strReptNo As String
    Dim str就诊类别 As String
    Dim dbl自付金额 As Double
    Dim dbl统筹资金 As Double
    Dim dbl原始金额 As Double
    Dim dblAccount As Double
    Dim intWait As Integer
    Dim sngBegin As Single
    
    Dim rsTmp As New ADODB.Recordset
    Dim rsExpen As New ADODB.Recordset
    
    gstrSQL = "select 住院号,姓名 from 病人信息 where 病人id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "莲合医保", CLng(rsExse!病人ID))
    
    str住院号 = IIf(IsNull(rsTmp!住院号), "", rsTmp!住院号)
    STR姓名 = rsTmp!姓名
    rsTmp.Close
    
    With rsExse
        dbl原始金额 = 0
        .MoveFirst
        Do While Not .EOF
            dbl原始金额 = dbl原始金额 + !金额
            .MoveNext
        Loop
        .MoveFirst
    End With
    
    gstrSQL = "select a.id, A.NO,A.序号,B.名称 as 开单部门,C.项目编码 as 医保项目编码,d.名称 as 项目," & _
        " A.发生时间,A.开单人 as 医生,decode(d.是否变价,1,a.实收金额,Nvl(A.付数,1)*A.数次) as 数量,decode(d.是否变价,1,1,a.实收金额/(Nvl(A.付数,1)*A.数次)) 单价" & _
        " from 住院费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=[2]) C,收费细目 d " & _
        " where A.开单部门ID=B.ID(+) And A.收费细目ID=C.收费细目ID(+) and a.收费细目id=d.id " & _
        " And A.病人ID=[1] And A.记帐费用=1 And Nvl(A.是否上传,0)=0 And Nvl(A.实收金额,0)<>0 And Nvl(A.附加标志,0)<>9 And A.记录状态<>0"
    Set rsExpen = zlDatabase.OpenSQLRecord(gstrSQL, "莲合医保", rsExse!病人ID, CLng(TYPE_成都莲合))
    
    With rsExpen
        str就诊类别 = "02"
        Do While Not .EOF
            '删除以前未保存功的
            If IsNull(!医保项目编码) Then
                MsgBox "HIS中的项目“" & !项目 & "”未设置医保对应的编码," & vbCrLf & "不能报销医保基金,请检查！", vbExclamation + vbOKOnly, gstrSysName
                Exit Function
            End If
            '下面已经设置了上传标志，但这里任要删除的目的是避免以前上传的没有设置是否上传标志
            'gstrSQL = "delete from rcpt_tab where LPAD(RTrim(hosp_id),8,'0')='" & Format(str住院号, "0000000000") & "' and rcpt_no='" & !no & "' and sno=" & !序号 & " and class='02' and to_char(r_date,'yyyy-mm-dd HH24:MI:SS')='" & Format(!发生时间, "yyyy-mm-dd hh:mm:ss") & "'"
            'gcnSybase.Execute gstrSQL
                
            gstrSQL = "insert into rcpt_tab(hosp_id,rcpt_no,sno,p_name,inter_id,class,amount,doctor_id,r_date,dept_id,exe_id,hosp_price)" _
                & " values('" & Format(str住院号, String(gintLen, "0")) & "','" & !NO & "'," & !序号 & ",'" & STR姓名 & "'," & !医保项目编码 & ",'02'," & !数量 & ",'" & !医生 & _
                "',to_date('" & Format(!发生时间, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'),'',''," & !单价 & ")"
            gcnSybase.Execute gstrSQL
            
            '上传后就不再上传
            gstrSQL = "Update 住院费用记录 set 是否上传=1 where id=" & !ID
            gcnOracle.Execute gstrSQL
            
            .MoveNext
        Loop
        
        Do While True
            dbl自付金额 = 0
            dbl统筹资金 = 0
            gstrSQL = "select acct_pay,self_pay from diag_rec where LPAD(RTrim(hosp_id)," & gintLen & ",'0')='" & Format(str住院号, String(gintLen, "0")) & "' and rcpt_no is null and class='02'" _
                    & " and sno is null and p_name='" & STR姓名 & "' and inter_id is null "
            If rsTmp.State = adStateOpen Then rsTmp.Close
            rsTmp.Open gstrSQL, gcnSybase
            If Not rsTmp.EOF Then
                dbl自付金额 = dbl自付金额 + IIf(IsNull(rsTmp!self_pay), 0, rsTmp!self_pay)
                dbl统筹资金 = dbl统筹资金 + IIf(IsNull(rsTmp!acct_pay), 0, rsTmp!acct_pay)
            End If
            
            If dbl自付金额 + dbl统筹资金 > 0 Then '= dbl原始金额
                住院虚拟结算_莲合 = "医保基金;" & dbl统筹资金 & ";0"
                Exit Do
            End If
            
            '无结果也允许按普通病人方式结帐
            If MsgBox("没有得到医保结果，你继续等待吗？", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                住院虚拟结算_莲合 = "医保基金;0;0"
                Exit Function
            End If
        Loop
    End With
End Function

Public Function 住院结算_莲合(lng结帐ID As Long, rs帐户 As ADODB.Recordset) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    
    Dim str住院号 As String
    Dim STR姓名 As String
    Dim strReptNo As String
    Dim str就诊类别 As String
    Dim dbl自付金额 As Double
    Dim dbl统筹资金 As Double
    Dim dbl原始金额 As Double
    Dim lng病人ID As Long
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date
    
    Dim cur住院基数 As Currency, cur发生费用 As Currency
    Dim cur进入统筹 As Currency, cur统筹支付 As Currency
    Dim cur首先自付 As Currency, cur全自付 As Currency
    
    Dim rsTmp As New ADODB.Recordset
On Error GoTo ErrH
    住院结算_莲合 = False
    
    gstrSQL = "select 住院号,姓名 from 病人信息 where 病人id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "莲合医保", CLng(rs帐户!病人ID))
    
    str住院号 = rsTmp!住院号
    STR姓名 = rsTmp!姓名
    rsTmp.Close
    
    dbl自付金额 = 0
    dbl统筹资金 = 0
    lng病人ID = rs帐户!病人ID
    
    gstrSQL = "select acct_pay,self_pay from diag_rec where LPAD(RTrim(hosp_id)," & gintLen & ",'0')='" & Format(str住院号, String(gintLen, "0")) & "' and rcpt_no is null and class='02'" _
            & " and sno is null and p_name='" & STR姓名 & "' and inter_id is null "
            
    If rsTmp.State = adStateOpen Then rsTmp.Close
    rsTmp.Open gstrSQL, gcnSybase
    If Not rsTmp.EOF Then
        dbl自付金额 = dbl自付金额 + IIf(IsNull(rsTmp!self_pay), 0, rsTmp!self_pay)
        dbl统筹资金 = dbl统筹资金 + IIf(IsNull(rsTmp!acct_pay), 0, rsTmp!acct_pay)
    End If

    gstrSQL = "Select Sum(结帐金额) as 结帐金额 From 住院费用记录 Where Nvl(附加标志,0)<>9 And 结帐ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "莲合医保", lng结帐ID)
    
    dbl原始金额 = rsTmp.Fields(0)
    
    
'    If dbl自付金额 + dbl统筹资金 = dbl原始金额 Then
        
        '填写结算表
        curDate = zlDatabase.Currentdate
        
        With rsTmp
            '住院基数,费用总额,进入统筹部分,统筹支付部份
            '由于对方不提供，所以不能提取住院基数和进入统筹金额
            
            cur住院基数 = 0
            cur发生费用 = dbl原始金额
            cur进入统筹 = 0
            cur统筹支付 = dbl统筹资金
            cur全自付 = 0
            cur首先自付 = cur发生费用 - cur全自付 - cur进入统筹
            
            '帐户年度信息
            Call Get帐户信息(TYPE_成都莲合, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
                    
            gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_成都莲合 & "," & Year(curDate) & "," & _
                cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 + cur进入统筹 & "," & _
                cur统筹报销累计 + cur统筹支付 & "," & int住院次数累计 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "莲合医保")
            
            '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
            gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_成都莲合 & "," & lng病人ID & "," & _
                Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
                cur统筹报销累计 & "," & int住院次数累计 & "," & cur住院基数 & ",NULL," & cur住院基数 & "," & _
                cur发生费用 & "," & cur全自付 & "," & cur首先自付 & "," & cur进入统筹 & "," & cur统筹支付 & ",NULL,NULL,NULL,NULL)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "莲合医保")
            
            '保险结算计算
            
            gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",1," & cur进入统筹 & "," & cur统筹支付 & ",NULL)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "莲合医保")
        End With
        '-------------------------------------------
        
        '删除中间数据库的结算数据
        gstrSQL = "delete from diag_rec where LPAD(RTrim(hosp_id)," & gintLen & ",'0')='" & Format(str住院号, String(gintLen, "0")) & "' and rcpt_no is null and class='02'" _
            & " and sno is null and p_name='" & STR姓名 & "' and inter_id is null "
        gcnSybase.Execute gstrSQL
        
        
        住院结算_莲合 = True
'    End If
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_莲合(lng结帐ID As Long, rs帐户 As ADODB.Recordset) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    Dim lng病人ID As Long
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date, lng新ID As Long
    
    
    Dim cur住院基数 As Currency, cur发生费用 As Currency
    Dim cur进入统筹 As Currency, cur统筹支付 As Currency
    Dim dbl自付金额  As Currency, dbl统筹资金  As Currency
    Dim cur首先自付 As Currency, cur全自付 As Currency
    
    Dim rsTmp As New ADODB.Recordset
    Dim str住院号 As String, STR姓名 As String
On Error GoTo ErrH
    住院结算冲销_莲合 = False
    gstrSQL = "select 住院号,姓名 from 病人信息 where 病人id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "莲合医保", CLng(rs帐户!病人ID))
    
    str住院号 = rsTmp!住院号
    STR姓名 = rsTmp!姓名
    rsTmp.Close
    
    dbl自付金额 = 0
    dbl统筹资金 = 0
    lng病人ID = rs帐户!病人ID
    
    gstrSQL = "delete from diag_rec where LPAD(RTrim(hosp_id)," & gintLen & ",'0')='" & Format(str住院号, String(gintLen, "0")) & "' and class='02'" _
            & " and p_name='" & STR姓名 & "' "
    gcnSybase.Execute gstrSQL
    
    curDate = zlDatabase.Currentdate
    '获取作废后的结帐ID
    gstrSQL = "Select A.ID From 病人结帐记录 A,病人结帐记录 B" & _
        " Where A.NO=B.NO And A.记录状态=2 And B.记录状态=3" & _
        " And B.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "莲合医保", lng结帐ID)
    If rsTmp.EOF Then
        Err.Raise 9000, gstrSysName, "未发现作废的结算数据！", vbInformation, gstrSysName
        Exit Function: 住院结算冲销_莲合 = False
    End If
    
    With rsTmp
        lng新ID = .Fields("ID").Value
        
        '帐户年度信息
        Call Get帐户信息(TYPE_成都莲合, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
        If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
        
        gstrSQL = "Select * From 保险结算计算 Where Nvl(档次,0)=0 And 结帐ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "莲合医保", lng结帐ID)
        
        If Not .EOF Then
            cur进入统筹 = IIf(IsNull(!进入统筹金额), 0, !进入统筹金额)
            cur统筹支付 = IIf(IsNull(!统筹报销金额), 0, !统筹报销金额)
        End If
    End With
    
    gstrSQL = "Select * From 保险结算记录 Where 性质=2 And 记录ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "莲合医保", lng结帐ID)
            
    With rsTmp
        If Not .EOF Then
            cur住院基数 = IIf(IsNull(!实际起付线), 0, !实际起付线)
            cur发生费用 = IIf(IsNull(!发生费用金额), 0, !发生费用金额)
            cur首先自付 = IIf(IsNull(!首先自付金额), 0, !首先自付金额)
            If cur进入统筹 = 0 Then cur进入统筹 = IIf(IsNull(!进入统筹金额), 0, !进入统筹金额)
            If cur统筹支付 = 0 Then cur统筹支付 = IIf(IsNull(!统筹报销金额), 0, !统筹报销金额)
            cur全自付 = IIf(IsNull(!全自付金额), 0, !全自付金额)
        End If
        
        '插入新的作废记录
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_成都莲合 & "," & Year(curDate) & "," & _
            cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 - cur进入统筹 & "," & _
            cur统筹报销累计 - cur统筹支付 & "," & int住院次数累计 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "莲合医保")
        
        '保险结算计算
        gstrSQL = "zl_保险结算计算_insert(" & lng新ID & ",1," & -1 * cur进入统筹 & "," & -1 * cur统筹支付 & ",NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "莲合医保")
        
        '保险结算记录
        gstrSQL = "zl_保险结算记录_insert(2," & lng新ID & "," & TYPE_成都莲合 & "," & lng病人ID & "," & Year(curDate) & "," & _
            cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & cur统筹报销累计 & "," & _
            int住院次数累计 & "," & cur住院基数 & ",NULL," & cur住院基数 & "," & -1 * cur发生费用 & "," & _
             -1 * cur全自付 & "," & -1 * cur首先自付 & "," & _
            -1 * cur进入统筹 & "," & -1 * cur统筹支付 & ",NULL,NULL,NULL,NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "莲合医保")
    End With
    住院结算冲销_莲合 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 错误信息_莲合(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息

End Function

Public Function 传输明细_莲合(ByVal str单据号 As String, ByVal int性质 As Integer, ByVal int状态 As Integer, ByVal intClinic As Integer) As Boolean
'功能: 传输门诊费用明细(划价单)。仅门诊使用。
'说明：因为ZLHIS9/10对收费划价单的记录方式不同，所以必须使用记录性质，记录状态参数。
'------------------------------------------------------------------------------------------------------------------
'调用模块：1121-门诊收费
    On Error GoTo errHand
    Dim rsExse As New ADODB.Recordset
    Dim 病人姓名 As String  '(10)
    Dim 单据号 As String '(8)
    Dim 序号 As Long   '(4,0)、
    Dim 医保项目编码 As Long  '(6,0)
    Dim 数量 As Double   '(8,2)
    Dim 金额 As Currency
    Dim 单价 As Currency
    Dim 开单医生 As String  '(6)、
    Dim 开单部门 As String  '(10)
    Dim 操作员部门 As String '(4)
    Dim 发生时间 As String
    Dim 住院号 As String
    
    传输明细_莲合 = False
    '删除以前未保存功的
    gcnSybase.BeginTrans
    gstrSQL = "delete from rcpt_tab " _
                     & " where LPAD(RTrim(hosp_id)," & gintLen & ",'0')='" & IIf(intClinic = 1, String(gintLen, "0"), Format(住院号, String(gintLen, "0"))) & "'" _
                     & "   and rcpt_no='" & str单据号 & "' " _
                     & "   and class='" & IIf(intClinic = 1, "01", "02") & "'"
    gcnSybase.Execute gstrSQL
    
    '门诊：病人id、病人姓名、单据号、序号、医保项目编码、数量、单价、金额、开单医生、开单部门、操作员、发生时间
    gstrSQL = "Select A.病人ID,A.姓名 As 病人姓名,A.No As 单据号,Nvl(A.价格父号,A.序号) As 序号," _
        & " C.项目编码 As 医保项目编码,decode(d.是否变价,1,Sum(A.实收金额),Avg(A.数次*Nvl(A.付数,1))) As 数量,decode(d.是否变价,1,1,Sum(A.标准单价)) As 单价," _
        & " Sum(A.实收金额) As 金额,A.开单人 As 开单医生,B.名称 As 开单部门,A.划价人 As 操作员,a.发生时间,d.名称 as 项目名称 " _
        & " From 住院费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=[4]) C,收费细目 D,保险帐户 F " _
        & " Where Nvl(A.实收金额,0)<>0 And Nvl(A.附加标志,0)<>9 And A.开单部门ID=B.ID(+) And A.收费细目ID=C.收费细目ID(+) and A.收费细目ID=d.id " _
        & " And A.病人ID=F.病人ID And F.险类=[4]" _
        & " And A.记录性质=[1] And A.记录状态=[2] And A.NO=[3]" _
        & " Group By A.No,Nvl(A.价格父号,A.序号),A.病人ID,A.姓名,C.项目编码,A.开单人,B.名称,A.划价人,d.名称,a.发生时间,d.是否变价 "
    Set rsExse = zlDatabase.OpenSQLRecord(gstrSQL, "莲合医保", int性质, int状态, str单据号, TYPE_成都莲合)
    
    With rsExse
        If .EOF Then
            MsgBox "没有一条发生的明细数据，可能是没有设置医保内码，请检查！", vbInformation, gstrSysName
            gcnSybase.RollbackTrans
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            If IsNull(!医保项目编码) = True Then
                MsgBox "费用中包含未设置保险支付项目的收费项目（" & !项目名称 & "）," & vbCrLf & "不能执行医保交易！", vbInformation, gstrSysName
                gcnSybase.RollbackTrans
                Exit Function
            End If
            医保项目编码 = !医保项目编码
            病人姓名 = !病人姓名
            单据号 = str单据号
            序号 = !序号
            数量 = !数量
            单价 = !单价
            金额 = !金额
            开单医生 = !开单医生            '(6)
            开单部门 = !开单部门   '         (10)
            '操作员部门 = StrConv(Mid(StrConv(!操作员部门, vbFromUnicode), 1, 4), vbUnicode) '(4)
            发生时间 = Format(!发生时间, "yyyy-mm-dd")
                
            gstrSQL = "insert into rcpt_tab" _
                    & "(hosp_id,rcpt_no,sno,p_name,inter_id,class,amount,doctor_id,r_date,dept_id,exe_id,hosp_price)" _
             & " values('" & IIf(intClinic = 1, "0", Format(住院号, String(gintLen, "0"))) & "','" & 单据号 & "'," & 序号 & ",'" & 病人姓名 & "'," & 医保项目编码 & ",'" & IIf(intClinic = 1, "01", "02") & "'," _
                      & 数量 & ",'" & 开单医生 & "',to_date('" & 发生时间 & "','yyyy-mm-dd'),'','" & 开单部门 & "'," & 单价 & ")"
            gcnSybase.Execute gstrSQL
            
            .MoveNext
        Loop
    End With
    传输明细_莲合 = True
    gcnSybase.CommitTrans
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    gcnSybase.RollbackTrans
End Function
