Attribute VB_Name = "mdl昭通"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)

Public blnConnIsOpen As Boolean, gcn昭通 As New ADODB.Connection
Private mstr卡号 As String, mstr入院状态 As String, mcur个帐 As Currency, mcur统筹 As Currency
'候国君要求：自费药品也需对码，然后上传到中心等待审核，在录入该项目时也需检查是否通过中心审批，且自费药品需操作员选择普通/病种需要等信息
'修改号：5825
'内容:
'1.修改I100-3标准药品目录结构以符合省统一标准(p4).
'2.I110-1申报药品中,若属于省标准目录的甲乙类,申报后不再需要经过中心审批即可使用.非目录药品仍然按原来的方法处理(需要申报、审批过程)
'3.增加I110-4机构前台修改正常记录

'调试时使用
Dim gArrayTest() As String

Public Function 医保初始化_昭通(Optional ByVal bln操作员检查 As Boolean = True) As Boolean
    Dim rsTemp As New ADODB.Recordset, str参数值 As String, lngPort As Long, strServer As String, _
        strSN As String, strDataSource As String
    On Error GoTo errHandle
    If blnConnIsOpen Then
        医保初始化_昭通 = True
        Exit Function
    End If
    
    strDataSource = Mid(gcnOracle.ConnectionString, InStr(UCase(gcnOracle.ConnectionString), "SERVER=") + 7)
    strDataSource = Left(strDataSource, InStr(strDataSource, """;") - 1)
    
    On Error Resume Next
    If gcn昭通.State = 1 Then gcn昭通.Close
    gcn昭通.ConnectionString = "Provider=MSDAORA.1;Password=his;User ID=ybuser;Data Source=" & strDataSource & ";Persist Security Info=True"
    gcn昭通.CursorLocation = adUseClient
    gcn昭通.Open
    If Err.Number <> 0 Then
        MsgBox "连接中间数据库失败", vbInformation, "医保初始化"
        Exit Function
    End If
    On Error GoTo errHandle
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_昭通)
        
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "昭通许可证"
                strSN = str参数值
            Case "昭通服务器"
                strServer = str参数值
            Case "昭通端口号"
                lngPort = CLng(str参数值)
        End Select
        rsTemp.MoveNext
    Loop
    
    If strSN = "" Or strServer = "" Or lngPort = 0 Then
        MsgBox "保险参数设置不完整，不能连接到医保", vbInformation, "医保初始化"
        Exit Function
    End If
    
    If frmConn昭通.ConnCenter(strServer, lngPort, strSN, IIf(bln操作员检查, UserInfo.ID, 0)) = False Then Exit Function
    
    gstrSQL = "Select * From 保险类别 Where 序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保初始化", TYPE_昭通)
    gstr医院编码 = Trim(rsTemp!医院编码)
    
    blnConnIsOpen = True
    医保初始化_昭通 = True
    Exit Function
errHandle:
    WriteInfo "初始化发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 医保终止_昭通() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    On Error GoTo errHandle
    
    Call frmConn昭通.ConnClose
    医保终止_昭通 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_昭通(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select 帐户余额 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取个人帐户余额", lng病人ID, TYPE_昭通)
    
    If rsTemp.EOF Then
        个人余额_昭通 = 0
    Else
        个人余额_昭通 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
    End If
End Function

Public Function 身份标识_昭通(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim strPatiInfo As String, cur余额 As Currency
    Dim arr, datCurr As Date, str门诊号 As String
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    If bytType = 1 Then
        strPatiInfo = frmIdentify昭通住院.GetPatient(bytType, mstr入院状态)
    Else
        strPatiInfo = frmIdentify昭通.GetPatient(bytType)
    End If

    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '建立病人档案信息，传入格式：
        '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
        '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
        '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
        lng病人ID = BuildPatiInfo(bytType, strPatiInfo, lng病人ID, TYPE_昭通)

        '返回格式:中间插入病人ID
        If bytType = 1 Then
            strPatiInfo = frmIdentify昭通住院.mstrPatient & lng病人ID & ";" & frmIdentify昭通住院.mstrOther
        Else
            strPatiInfo = frmIdentify昭通.mstrPatient & lng病人ID & ";" & frmIdentify昭通.mstrOther
        End If
    Else
        身份标识_昭通 = ""
        MsgBox "医保病人信息提取失败", vbInformation, gstrSysName
        Exit Function
    End If
    身份标识_昭通 = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_昭通 = ""
End Function

Public Function 门诊结算_昭通(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency, Optional ByRef strAdvance As String) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
'        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset, rsCheck As New ADODB.Recordset
    Dim str操作员 As String, datCurr As Date, str就诊编号 As String, str结算方式 As String
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, curCount As Currency
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur起付线 As Currency, cur基本统筹限额 As Currency, strTemp As String
    Dim cur大额统筹限额 As Currency, cur基数自付 As Currency, cur余额 As Currency
    Dim cur发生费用 As Currency, cur先自付 As Currency, strPara As String
    Dim blnOld As Boolean
    Dim blnBalance As Boolean
    
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "Select * From 门诊费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng病人ID = rs明细!病人ID
    
    If rs明细.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "没有填写收费记录"
        Exit Function
    End If
    curCount = 0
    
    WriteInfo vbCrLf & "开始门诊结算"
    '组织费用明细
    While Not rs明细.EOF
        gstrSQL = "Select * From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        strTemp = rsTemp!类别
        
        gstrSQL = "Select * From 保险支付项目 Where 险类=103 And 收费细目ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险支付项目", CLng(rs明细!收费细目ID))
        If rsTemp.EOF Then             '如果没有对码则不允许使用
            gstrSQL = "Select * From 收费细目 Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取收费细目", CLng(rs明细!收费细目ID))
            Err.Raise 9000, gstrSysName, "项目[" & rsTemp!名称 & "]没有对应的医保项目,请先进行对码"
            WriteInfo "项目[" & rsTemp!名称 & "]没有对应的医保项目,退出门诊结算"
            Exit Function
        ElseIf Nvl(rsTemp!附注) <> "已审批" And InStr(" 5 6 7 ", " " & strTemp & " ") > 0 Then
            strTemp = Nvl(rsTemp!附注, "未审批")
            
            '如果是省目录的甲、乙类项目，不需要检查审批标志
            gstrSQL = "Select lb From tab_syml where dm='" & rsTemp!项目编码 & "'"
            Call OpenRecordset_OtherBase(rsCheck, "如果是省目录的甲、乙类项目，不需要检查审批标志", gstrSQL, gcn昭通)
            If rsCheck.RecordCount <> 0 Then
                If rsCheck!lb = 15 Then
                    gstrSQL = "Select * From 收费细目 Where ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取收费细目", CLng(rs明细!收费细目ID))
                    Err.Raise 9000, gstrSysName, "项目[" & rsTemp!名称 & "]" & strTemp & "，不能使用"
                    WriteInfo "项目[" & rsTemp!名称 & "]" & strTemp & "，不能使用"
                    Exit Function
                End If
            End If
        End If
        
        '分号分隔行,逗号分隔列
        If InStr(" 5 6 7 ", " " & strTemp & " ") > 0 Then
            strTemp = rs明细!收费细目ID
'        ElseIf strTemp = "7" Then
'            '获取中草药的费用类型
'            If Trim(Nvl(rs明细!摘要)) = "" Or InStr(1, ",UZY01,UZY02,UZY03,", "," & UCase(Nvl(rs明细!摘要)) & ",") = 0 Then
'                strTemp = GetItemInfo_昭通(1, lng病人ID, rs明细!收费细目ID, Nvl(rs明细!摘要), rs明细!NO)
'            Else
'                strTemp = Nvl(rs明细!摘要)
'            End If
        Else
            strTemp = rsTemp!项目编码
        End If
        strPara = strPara & ";" & strTemp & "," & rs明细!数次 * rs明细!付数 & "," & _
            Round(rs明细!实收金额 / (rs明细!数次 * rs明细!付数), 4) & "," & rs明细!实收金额
        
        curCount = curCount + rs明细!实收金额
        rs明细.MoveNext
    Wend
    
    If strPara <> "" Then strPara = Mid(strPara, 2)             '去掉开头的分号
    
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取医保病人信息", lng病人ID)
    
    strPara = rsTemp!卡号 & vbTab & Nvl(rsTemp!密码, " ") & vbTab & strPara
    WriteInfo "交易传递参数:" & strPara
    If frmConn昭通.Execute("I200", 1, strPara, "正在进行医保交易,请稍候......") = False Then Exit Function
    If frmConn昭通.Query(0, 1) = False Then Exit Function
    strPara = frmConn昭通.strReturnInfo
    WriteInfo "交易返回数据:" & strPara
    If strPara = "" Then
        Err.Raise 9000, gstrSysName, "返回数据格式错误", vbInformation, "门诊结算"
        Exit Function
    End If
    
    blnBalance = True
    str就诊编号 = Split(strPara, vbTab)(0)
    cur个人帐户 = Split(strPara, vbTab)(2)
        
    If cur个人帐户 <> 0 Then
        str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户
    End If
    
    WriteInfo "交易结果:" & Mid(str结算方式, 3)
    '如果存在
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        #If gverControl < 2 Then
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
        #Else
            strAdvance = str结算方式
            gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
    End If
    #If gverControl < 2 Then
        blnOld = True
        frm结算信息.ShowME lng结帐ID
    #End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_昭通, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
            "," & TYPE_昭通 & "," & Year(datCurr) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & _
            cur起付线 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_昭通 & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & ",NULL,NULL," & _
            cur发生费用 & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
            cur个人帐户 & ",NULL,NULL,NULL,'" & str就诊编号 & "'" & IIf(blnOld, "", ",1") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    门诊结算_昭通 = True
    
    WriteInfo "保存门诊发票数据"
    Call SaveOutExse(str就诊编号)
    
    WriteInfo "完成门诊交易"
    Exit Function
errHandle:
    If blnBalance Then
        ErrMsgBox "请通过医保工具将当前这笔门诊收费单据冲销，以下信息请记录：" & vbCrLf & _
            "就诊编号：" & str就诊编号 & "；个人帐户：" & cur个人帐户, vbInformation, "门诊结算"
    Else
        ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
        Err.Clear
        Exit Function
    End If
    WriteInfo "发生错误:" & Err.Description
End Function

Public Function 门诊结算冲销_昭通(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset
    Dim lng冲销ID As Long, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, strPara As String
    Dim curCount As Currency, str操作员 As String
    Dim datCurr As Date
    
    WriteInfo vbCrLf & "开始门诊冲销"
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    str操作员 = UserInfo.编号
    
    gstrSQL = "Select 病人ID,结帐金额,操作员编号 From 门诊费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "找不到单据的明细记录,不能进行冲销"
        Exit Function
    End If
    lng病人ID = rsTemp!病人ID
    If rsTemp!操作员编号 <> str操作员 Then
        Err.Raise 9000, gstrSysName, "医保规定必须由执行本单据门诊结算的操作员进行冲销"
        Exit Function
    End If
    
    Do Until rsTemp.EOF
        curCount = curCount + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_昭通, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    If IsNull(rsTemp!备注) Then
        Err.Raise 9000, gstrSysName, "该单据的就诊编号丢失，不能作废。"
        Exit Function
    End If
    str就诊编号 = rsTemp!备注
    
    strPara = str就诊编号 & vbTab & rsTemp!个人帐户支付
    WriteInfo "交易传递参数:" & strPara
    
    '调用接口数冲销
    If Not frmConn昭通.Execute("I220", 0, strPara, "正在进行医保交易,请稍候......") Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(TYPE_昭通, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_昭通 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_昭通 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & curCount * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        cur个人帐户 * -1 & ",NULL,NULL,NULL,'" & str就诊编号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    门诊结算冲销_昭通 = True
    WriteInfo "完成门诊冲销"
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
    WriteInfo "发生错误:" & Err.Description
End Function

Public Function 入院登记_昭通(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset, str卡号 As String, datCurr As Date, strInNote As String, _
        strPara As String, str住院序号 As String

    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    WriteInfo vbCrLf & "开始入院登记"
    
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,D.编码 as 科室编码,A.入院病床,A.住院医师,C.卡号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        入院登记_昭通 = False
        Exit Function
    End If
    
    '获取入院诊断（病种编码）
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID)  '入院诊断
'    If strInNote <> "" Then
'        strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
'    End If
    WriteInfo "取得入院诊断：" & strInNote
    
    strPara = rsTemp!卡号 & vbTab & rsTemp!密码 & vbTab & mstr入院状态 & vbTab & _
              Nvl(rsTemp!住院科室, " ") & vbTab & ToVarchar(Nvl(rsTemp!入院病床, "0"), 10) & vbTab & strInNote
    
'    gstrSQL = "Select sum(金额) From 病人预交记录 Where 病人ID=" & lng病人id & " And 主页ID=" & lng主页ID
'    Call OpenRecordset(rsTemp, gstrSysName)
    strPara = strPara & vbTab & "0"         ' Nvl(rsTemp(0), 0)
    WriteInfo "交易传递参数:" & strPara
    
    '调用接口进行登记
    If Not frmConn昭通.Execute("I300", 1, Replace(strPara, vbTab & vbTab, vbTab & " " & vbTab), "正在进行医保交易,请稍候......") Then Exit Function
    If frmConn昭通.Query(0, 1) = False Then Exit Function
    str住院序号 = Replace(frmConn昭通.strReturnInfo, " ", "")
    
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_昭通 & ",'顺序号','''" & str住院序号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新住院序号")
    
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_昭通 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_昭通 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_昭通 = False
End Function

Public Function 入院登记撤消_昭通(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将入院登记撤消信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim str顺序号 As String
    Dim rsTemp As New ADODB.Recordset, str卡号 As String, datCurr As Date, strInNote As String

    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    WriteInfo vbCrLf & "开始撤消入院"
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的医保信息。", vbInformation, gstrSysName
        入院登记撤消_昭通 = False
        Exit Function
    End If
    If Nvl(Replace(rsTemp!顺序号, " ", ""), "") = "" Then
        MsgBox "取医保病人住院序号错误", vbInformation, gstrSysName
        Exit Function
    End If
    str顺序号 = Nvl(Replace(rsTemp!顺序号, " ", ""))
    
    '只要存在病人费用记录则不允许办理撤销入院,只能出院登记
    gstrSQL = "SELECT 1 FROM 住院费用记录 WHERE 病人ID=[1] AND 主页ID=[2] AND ROWNUM<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "只要存在病人费用记录则不允许办理撤销入院,只能出院登记", lng病人ID, lng主页ID)
    If rsTemp.RecordCount <> 0 Then
        MsgBox "该病人已发生费用,不能办理撤销入院,只能办理出院!", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "交易传递参数:" & str顺序号
    '调用接口进行登记
    If Not frmConn昭通.Execute("I305", 0, str顺序号, "正在进行医保交易,请稍候......") Then Exit Function

     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_昭通 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记撤消_昭通 = True
    WriteInfo "撤消入院完成"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记撤消_昭通 = False
End Function

Public Function 转科转床_昭通(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将转科转床信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim strInNote As String, rsTemp As New ADODB.Recordset, strPara As String, str住院序号 As String
    
    '求出病人的相关信息
    On Error GoTo errHandle
    WriteInfo vbCrLf & "开始入院信息变动"
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    If Nvl(Replace(rsTemp!顺序号, " ", ""), "") = "" Then
        MsgBox "取病人入院序号错误", vbInformation, gstrSysName
        Exit Function
    End If
    str住院序号 = Replace(rsTemp!顺序号, " ", "")
    
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,D.编码 as 科室编码,A.入院病床,A.住院医师,C.顺序号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        转科转床_昭通 = False
        Exit Function
    End If
    
    '获取入院诊断
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID, True, True, False) '入院诊断
    If strInNote = "" Then strInNote = "普通"
    WriteInfo "获取入院诊断：" & strInNote
    
    strPara = str住院序号 & vbTab & Nvl(rsTemp!科室编码, "0") & vbTab & ToVarchar(Nvl(rsTemp!入院病床, "0"), 10) & vbTab & strInNote
    WriteInfo "交易传递参数:" & strPara
    
    '调用接口进行登记
    If Not frmConn昭通.Execute("I309", 1, strPara, "正在进行医保交易,请稍候......") Then Exit Function
     
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_昭通 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    转科转床_昭通 = True
    WriteInfo "入院信息变更完成"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    转科转床_昭通 = False
End Function

Public Function 住院虚拟结算_昭通(rs明细 As Recordset, lng病人ID As Long, str医保号 As String) As String
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
        strTemp As String, cur个帐支付 As Currency, cur现金 As Currency, curTemp As Currency, _
        cur公务员补助 As Currency, lng主页ID As Long, cur大病统筹 As Currency, cur基本统筹 As Currency
    Dim cur费用总额 As Currency, str住院号 As String, str结算方式 As String, strReturn() As String
    Dim strMessage As String
    
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
    
    If 记帐传输_昭通("", 0, strMessage, lng病人ID) = False Then
        MsgBox strMessage, vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select * From 保险帐户 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    str住院号 = Nvl(Replace(rsTemp!顺序号, " ", ""), "")
    If str住院号 = "" Then
        MsgBox "不能获取病人住院顺序号，不能进行结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "交易传递参数：" & str住院号
    If frmConn昭通.Execute("I361", 5, str住院号, "正在读取病人费用信息......") = False Then Exit Function
    If frmConn昭通.Query(0, 1) = False Then Exit Function
    WriteInfo "返回：" & frmConn昭通.strReturnInfo
    strReturn = Split(frmConn昭通.strReturnInfo, vbTab)
    cur基本统筹 = CCur(strReturn(0))
    cur大病统筹 = CCur(strReturn(1))
    
    mcur统筹 = cur基本统筹 + cur大病统筹
    
    curTemp = 个人余额_昭通(lng病人ID)
    gstrSQL = "Select nvl(sum(金额),0) From 病人预交记录 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If cur费用总额 - mcur统筹 > rsTemp(0) Then
        cur个帐支付 = cur费用总额 - mcur统筹 - rsTemp(0)
    Else
        cur个帐支付 = 0
    End If
    If cur个帐支付 > curTemp Then cur个帐支付 = curTemp
    If cur个帐支付 < 0 Then cur个帐支付 = 0
    mcur个帐 = cur个帐支付
    str结算方式 = "个人帐户;" & cur个帐支付 & ";1"
    If cur基本统筹 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 <> "", "|", "") & "基本统筹;" & cur基本统筹 & ";0"
    If cur大病统筹 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 <> "", "|", "") & "大病统筹;" & cur大病统筹 & ";0"
    
    住院虚拟结算_昭通 = str结算方式
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_昭通(lng结帐ID As Long, lng病人ID As Long) As Boolean
'功能：对住院费用进行明细传递并且进行结算
'如果住院费用明细传递失败，就直接结束函数，返回函数失败
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim rs昭通 As New ADODB.Recordset, strTemp As String, lng主页ID As Long, str住院号 As String
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, datCurr As Date, strPara As String
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur个帐支付 As Currency, cur余额 As Currency, cur发生费用 As Currency, cur统筹支付 As Currency
    Dim cur救助金支付 As Currency, cur公务员补助 As Currency, str交易流水号 As String, str卡号 As String
    Dim str出院病床 As String, str出院诊断 As String
    On Error GoTo errHandle
    WriteInfo vbCrLf & "开始出院结算"
    gstrSQL = "Select * from 保险帐户 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    str住院号 = Replace(rsTemp!顺序号, " ", "")
    
    gstrSQL = "Select max(主页ID) From 病案主页 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    lng主页ID = rsTemp(0)
    gstrSQL = "Select A.出院病床,B.名称 From 病案主页 A,部门表 B Where A.出院科室ID=B.ID And 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    
    str出院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, False, True) '入院诊断
    If str出院诊断 = "" Then str出院诊断 = " "
    
    WriteInfo "修改出院信息"
    strTemp = str住院号 & vbTab & rsTemp!名称 & vbTab & ToVarchar(rsTemp!出院病床, 10) & vbTab & str出院诊断
    WriteInfo "交易传递参数：" & strTemp
    frmConn昭通.Execute "I309", 1, strTemp, "正在进行医保交易......"
    
    gstrSQL = "Select * From 病人预交记录 Where 结帐id=[1] And 结算方式='个人帐户'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    If rsTemp.EOF Then
        cur个帐支付 = 0
    Else
        cur个帐支付 = Nvl(rsTemp!冲预交, 0)
    End If
    
    Screen.MousePointer = 0
    If cur个帐支付 <> 0 Then
        If frmIdentify昭通.GetPatient(0) = "" Then Exit Function
        str卡号 = Split(frmIdentify昭通.mstrPatient, ";")(0)
        Unload frmIdentify昭通
    End If
    
    gstrSQL = "Select * from 保险帐户 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    str住院号 = Replace(rsTemp!顺序号, " ", "")
    
    If cur个帐支付 <> 0 Then
        strPara = str住院号 & vbTab & str卡号 & vbTab & IIf(Nvl(rsTemp!密码, "") = "", " ", rsTemp!密码) & vbTab & cur个帐支付
    Else
        strPara = str住院号 & vbTab & " " & vbTab & IIf(Nvl(rsTemp!密码, "") = "", " ", rsTemp!密码) & vbTab & cur个帐支付
    End If
    WriteInfo "交易传递参数：" & strPara
    If frmConn昭通.Execute("I340", 1, strPara, "正在进行出院结算......") = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(TYPE_昭通, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
            "," & TYPE_昭通 & "," & Year(datCurr) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + cur个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + mcur统筹 & "," & int住院次数累计 & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "昭通医保")
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_昭通 & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 + cur个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + mcur统筹 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            cur发生费用 & ",0,0,NULL," & mcur统筹 & ",NULL,NULL," & _
            cur个帐支付 & ",NULL,NULL,NULL,'" & str住院号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "昭通医保")
    
    住院结算_昭通 = True
    
    Call SaveInExse(str住院号)
    
    WriteInfo "出院结算成功"
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_昭通(lng结帐ID As Long) As Boolean
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim rs昭通 As New ADODB.Recordset, strTemp As String, lng主页ID As Long, str住院号 As String
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, datCurr As Date
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur个帐支付 As Currency, cur余额 As Currency, cur发生费用 As Currency, cur统筹支付 As Currency
    Dim cur救助金支付 As Currency, cur公务员补助 As Currency, str交易流水号 As String, str卡号 As String
    Dim lng病人ID As Long
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select * From 病人预交记录 Where 结帐id=[1] And 结算方式='个人帐户'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    If rsTemp.EOF Then
        cur个帐支付 = 0
    Else
        cur个帐支付 = Nvl(rsTemp!冲预交, 0)
    End If
    
    gstrSQL = "Select * from 保险结算记录 Where 记录id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    str住院号 = rsTemp!备注
    lng病人ID = rsTemp!病人ID
    
    If frmConn昭通.Execute("I345", 0, str住院号, "正在取消出院结算......") = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(TYPE_昭通, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
            "," & TYPE_昭通 & "," & Year(datCurr) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 - cur个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 - rsTemp!统筹报销金额 & "," & int住院次数累计 & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "昭通医保")
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_昭通 & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 - cur个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 - rsTemp!统筹报销金额 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            cur发生费用 & ",0,0,NULL," & 0 - rsTemp!统筹报销金额 & ",NULL,NULL," & _
            0 - cur个帐支付 & ",NULL,NULL,NULL,'" & str住院号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "昭通医保")
    
    住院结算冲销_昭通 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 记帐传输_昭通(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset, lng主页ID As Long, str住院序号 As String, strPara As String, _
        rs明细 As New ADODB.Recordset, strID() As String, lngLoop As Long, strRetu() As String, _
        str批号 As String, blnAll As Boolean, cur单价 As Currency
    Dim strTemp As String, strMessage As String
    Dim int用药标志 As Integer
    
    On Error GoTo errHandle
    blnAll = True
    If lng病人ID <> 0 Then
        gstrSQL = "Select Max(主页ID) From 病案主页 Where 病人id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
        lng主页ID = rsTemp(0)
    End If
    
    If str单据号 <> "" Then
        gstrSQL = " Select A.* From 住院费用记录 A,保险帐户 B " & _
                  " Where A.门诊标志=2 And A.记录状态<>0 And A.记录状态<>3 And A.记录状态<>2 And nvl(A.附加标志,0)<>9 " & _
                  " and nvl(A.实收金额,0)<>0 and A.记录性质=[1] and A.NO=[2]" & _
                  " and A.病人ID=B.病人ID and B.险类=[2]" & _
                  " order by A.病人ID,A.主页ID,A.序号"
        Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, int性质, str单据号, TYPE_昭通)
    Else
        gstrSQL = " Select * From 住院费用记录 " & _
                  " Where 门诊标志=2 And 记录状态<>0 And 记录状态<>3 And 记录状态<>2 And nvl(附加标志,0)<>9 " & _
                  " and nvl(实收金额,0)<>0 and 病人id=[1] And 主页id=[2]" & _
                  " order by 主页ID,序号"
        Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    End If
    
    If rs明细.EOF Then
        MsgBox "没有需要上传的病人费用", vbInformation, gstrSysName
        记帐传输_昭通 = True
        Exit Function
    End If
    
    lng病人ID = 0
    strPara = ""
    ReDim strID(rs明细.RecordCount)
    lngLoop = 0
    While Not rs明细.EOF
        If lng病人ID <> rs明细!病人ID Then
            lng病人ID = rs明细!病人ID
            gstrSQL = "Select * From 保险帐户 Where 病人id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
            If rsTemp.EOF Then
                str消息 = "取病人医保信息时失败"
                Exit Function
            End If
            If Nvl(Replace(rsTemp!顺序号, " ", ""), "") = "" Then
                str消息 = "取医保病人住院序号失败"
                Exit Function
            End If
            str住院序号 = Replace(rsTemp!顺序号, " ", "")
            '调用接口进行费用登记预读
'            If Not frmConn昭通.Execute("I320", 0, str住院序号, "正在进行医保交易,请稍候......") Then Exit Function
            
            If strPara <> "" Then   '如果有参数表示该病人有数据，调用接口上传
                strPara = Left(strPara, Len(strPara) - 1)
                frmConn昭通.Execute "I320", 1, strPara, "正在进行医保交易,请稍候......"
            
                If frmConn昭通.Query(0, 1, "正在读取明细上传返回的结果集") = False Then Exit Function
                frmConn昭通.strReturnInfo = Mid(frmConn昭通.strReturnInfo, InStr(1, frmConn昭通.strReturnInfo, vbTab) + 1)
                frmConn昭通.mlngRows = UBound(Split(frmConn昭通.strReturnInfo, ";"))
                For lngLoop = 1 To frmConn昭通.mlngRows
                    If Split(Split(frmConn昭通.strReturnInfo, ";")(lngLoop - 1), ",")(2) = "0" Then
                        gstrSQL = "zl_病人记帐记录_上传 ('" & strID(lngLoop - 1) & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
                    Else
                        gstrSQL = "Select B.名称 From 住院费用记录 A,收费细目 B Where A.收费细目ID=B.ID And A.ID=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(strID(lngLoop - 1)))
                        If rsTemp.RecordCount <> 0 Then
                            strMessage = strMessage & "项目[" & rsTemp(0) & "]上传失败，返回信息：" & Split(Split(frmConn昭通.strReturnInfo, ";")(lngLoop - 1), ",")(3) & Chr(13) & Chr(10)
                        Else
                            MsgBox "没有找到记录,ID为:" & strID(lngLoop - 1), vbInformation, gstrSysName
                            Exit Function
                        End If
                        blnAll = False
                    End If
                Next
            
            End If
            strPara = str住院序号 & vbTab & "0" & vbTab
            lngLoop = 0
        End If
        
        gstrSQL = "Select * From 保险支付项目 Where 收费细目ID=[1] And 险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID), TYPE_昭通)
        If rsTemp.EOF Then
            gstrSQL = "Select 编码||名称 From 收费细目 Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
            str消息 = "记录中有未对码的项目[" & rsTemp.Fields(0).Value & "，不能上传到医保中心！"
            Exit Function
        End If
        
        '如果单价大于限价，则按限价执行
        cur单价 = Format(rs明细!实收金额 / (rs明细!付数 * rs明细!数次), "0.##")
'        If Nvl(rsTemp!附注, "0") <> "0" Then
'            If cur单价 > CCur(rsTemp!附注) Then
'                cur单价 = CCur(rsTemp!附注)
'            End If
'        End If
        int用药标志 = 0
        strTemp = rsTemp!项目编码
        gstrSQL = "Select * From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        Select Case rsTemp!类别
        Case "5", "6", "7"
            '获取非目录内药品用药标志
            If Get自费药品(UCase(strTemp)) Then
                If Trim(Nvl(rs明细!摘要)) = "" Or InStr(1, ",普通,抢救,病情需要（大病）,", "," & UCase(Nvl(rs明细!摘要)) & ",") = 0 Then
                    strTemp = GetItemInfo_昭通(2, lng病人ID, rs明细!收费细目ID, Nvl(rs明细!摘要), rs明细!NO)
                Else
                    strTemp = Nvl(rs明细!摘要)
                End If
                int用药标志 = IIf(strTemp = "抢救", 1, IIf(strTemp = "病情需要（大病）", 2, 0))
            End If
            strTemp = rsTemp!ID
'        Case "7"
'            '获取中草药的费用类型
'            If Trim(Nvl(rs明细!摘要)) = "" Or InStr(1, ",UZY01,UZY02,UZY03,", "," & UCase(Nvl(rs明细!摘要)) & ",") = 0 Then
'                strTemp = GetItemInfo_昭通(2, lng病人ID, rs明细!收费细目ID, Nvl(rs明细!摘要), rs明细!NO)
'            Else
'                strTemp = Nvl(rs明细!摘要)
'            End If
        End Select
            
        strPara = strPara & Format(rs明细!发生时间, "yyyymmdd") & "," & strTemp & "," & _
            Format(rs明细!付数 * rs明细!数次, "0.##") & "," & Format(rs明细!实收金额 / (rs明细!付数 * _
            rs明细!数次), "0.##") & "," & int用药标志 & ";"
        
        strID(lngLoop) = rs明细!ID
        lngLoop = lngLoop + 1
        rs明细.MoveNext
    Wend
    
    If strPara <> "" Then           '处理最后一次循环的结果
        gstrSQL = "Select * From 保险帐户 Where 病人id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
        If rsTemp.EOF Then
            str消息 = "取病人医保信息时失败"
            Exit Function
        End If
        If Nvl(Replace(rsTemp!顺序号, " ", ""), "") = "" Then
            str消息 = "取医保病人住院序号失败"
            Exit Function
        End If
        str住院序号 = Replace(rsTemp!顺序号, " ", "")
        '调用接口进行费用登记预读
'        If Not frmConn昭通.Execute("I320", 0, str住院序号, "正在进行医保交易,请稍候......") Then Exit Function
        
        If strPara <> "" Then   '如果有参数表示该病人有数据，调用接口上传
            strPara = Left(strPara, Len(strPara) - 1)
            frmConn昭通.Execute "I320", 1, strPara, "正在进行医保交易,请稍候......"
            
            If frmConn昭通.Query(0, 1, "正在读取明细上传返回的结果集") = False Then Exit Function
            frmConn昭通.strReturnInfo = Mid(frmConn昭通.strReturnInfo, InStr(1, frmConn昭通.strReturnInfo, vbTab) + 1)
            frmConn昭通.mlngRows = UBound(Split(frmConn昭通.strReturnInfo, ";"))
            For lngLoop = 1 To frmConn昭通.mlngRows
                If Split(Split(frmConn昭通.strReturnInfo, ";")(lngLoop - 1), ",")(2) = "0" Then
                    gstrSQL = "zl_病人记帐记录_上传 ('" & strID(lngLoop - 1) & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
                Else
                    gstrSQL = "Select B.名称 From 住院费用记录 A,收费细目 B Where A.收费细目ID=B.ID And A.ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(strID(lngLoop - 1)))
                    If rsTemp.RecordCount <> 0 Then
                        strMessage = strMessage & "项目[" & rsTemp(0) & "]上传失败，返回信息：" & Split(Split(frmConn昭通.strReturnInfo, ";")(lngLoop - 1), ",")(3) & Chr(13) & Chr(10)
                    Else
                        MsgBox "没有找到记录,ID为:" & strID(lngLoop - 1), vbInformation, gstrSysName
                        Exit Function
                    End If
                    blnAll = False
                End If
            Next
        End If
        strPara = str住院序号 & vbTab & "0" & vbTab
    End If
    Screen.MousePointer = vbDefault
    If strMessage <> "" Then
        str消息 = strMessage
        Exit Function
    End If
    
    记帐传输_昭通 = True
    Exit Function
errHandle:
    If ErrCenter() Then
        Resume
    End If
End Function

Public Function 出院登记撤消_昭通(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHandle
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_昭通 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记撤消_昭通 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    出院登记撤消_昭通 = False
End Function

Public Function 出院登记_昭通(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHandle
    Dim BLN无费退院 As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    '检查是否存在未结费用
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        '不存在未结费用,再检查是否结算过
        gstrSQL = "SELECT 1 FROM 住院费用记录 WHERE 结帐ID IS NOT NULL AND  病人ID=[1] AND 主页ID=[2] AND ROWNUM<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否结算过", lng病人ID, lng主页ID)
        If rsTemp.RecordCount = 0 Then
            '说明该病人是无费退院,没结算过也没有未结费用
            gstrSQL = "Select * From 保险帐户 Where 病人ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
            If Nvl(Replace(rsTemp!顺序号, " ", "")) = "" Then
                MsgBox "取医保病人住院序号错误", vbInformation, gstrSysName
                Exit Function
            End If
            
            WriteInfo "交易传递参数:" & Replace(rsTemp!顺序号, " ", "")
            '调用接口进行登记
            If Not frmConn昭通.Execute("I305", 0, rsTemp!顺序号, "正在进行医保交易,请稍候......") Then Exit Function
        End If
    End If
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_昭通 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_昭通 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    出院登记_昭通 = False
End Function

Public Function GetItemInfo_昭通(ByVal intType As Integer, ByVal lng病人ID As Long, ByVal lng细目ID As Long, _
    ByVal str摘要 As String, Optional ByVal str备注 As String = "") As String
    '中草药进入，需选择费用类型
    '非中草药类，且非医保药品，需选择是否大病或抢救用药
    'intType-调用类型(0-医嘱,1-门诊收费,2-住院记帐)
    Dim bln中草药 As Boolean
    Dim rsTemp As New ADODB.Recordset
    '针处理中草药:甲类：uzy01、乙类：uzy02、非医保：uzy03
    gstrSQL = "Select 类别 From 收费细目 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否是中草药", lng细目ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    If InStr(1, "5,6", rsTemp!类别) <> 0 Then
        gstrSQL = "Select 项目编码,附注 From 保险支付项目 Where 收费细目ID=[1] And 险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否对码", lng细目ID, TYPE_昭通)
        If rsTemp.RecordCount = 0 Then Exit Function
        If Not Get自费药品(UCase(rsTemp!项目编码)) Then Exit Function
        If Nvl(rsTemp!附注) <> "已审批" Then Exit Function
'    ElseIf rsTemp!类别 = "7" Then
'        bln中草药 = True
    Else
        Exit Function
    End If
    
    GetItemInfo_昭通 = frm昭通_项目信息.ShowME(intType, lng病人ID, lng细目ID, str摘要, str备注, bln中草药)
End Function

Public Function CheckInsureItem_昭通(ByVal lng收费细目ID As Long) As Boolean
    '检查项目是否通过中心的审批
    '如果是药品项目，且在省目录中是甲、乙类，则不必进行审批检查，可直接使用
    Dim str附注 As String, str项目编码 As String
    Dim rsTemp As New ADODB.Recordset

    gstrSQL = "Select 项目编码,附注 From 保险支付项目 Where 险类=[1] And 收费细目ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查项目是否通过中心的审批", TYPE_昭通, lng收费细目ID)
    If rsTemp.RecordCount <> 0 Then
        str附注 = Nvl(rsTemp!附注)
        str项目编码 = rsTemp!项目编码

        If Not Get自费药品(UCase(str项目编码)) Then Exit Function
        If str附注 <> "已审批" Then
            MsgBox "该项目还未通过中心审批，不能使用！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
End Function

Public Sub SaveOutExse(ByVal str流水号 As String)
    Dim str门诊资料 As String, str门诊明细 As String
    Dim arrData
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    '删除该流水号的相关数据
    WriteInfo "删除该流水号的相关数据"
    gcn昭通.Execute "Delete 门诊明细 Where 流水号='" & str流水号 & "'"
    gcn昭通.Execute "Delete 门诊资料 Where 流水号='" & str流水号 & "'"
    
    '读取门诊资料数据
    str门诊资料 = str流水号
    WriteInfo "(读取门诊资料数据)交易传递参数:" & str门诊资料
    If frmConn昭通.Execute("I280", 0, str门诊资料, "正在进行医保交易,请稍候......") = False Then Exit Sub
    If frmConn昭通.Query(0, 1) = False Then Exit Sub
    str门诊资料 = frmConn昭通.strReturnInfo
    WriteInfo "交易返回数据:" & str门诊资料
    If str门诊资料 = "" Then
        MsgBox "返回数据格式错误", vbInformation, "门诊结算"
        Exit Sub
    End If
    '保存门诊资料数据,少了支付后余额,在插入值中加入了收费时间(佴云明)
    arrData = Split(str门诊资料, vbTab)
    gstrSQL = " Insert Into 门诊资料(流水号,保险证号,姓名,支付前余额,帐户支付,支付后余额,现金支付,交易时间,收费地址,电话,交易金额,交易时间1)" & _
        " Values ('" & str流水号 & "','" & arrData(0) & "','" & arrData(1) & "'," & Val(arrData(2)) & "," & Val(arrData(3)) & "," & Val(arrData(4)) & "," & _
        "'" & arrData(5) & "','" & arrData(6) & "','" & arrData(7) & "','" & Val(arrData(8)) & "','" & arrData(9) & "','" & arrData(6) & "')"
    gcn昭通.Execute gstrSQL
    
    '读取门诊明细数据
    str门诊明细 = str流水号
    WriteInfo "(读取门诊明细数据)交易传递参数:" & str门诊明细
    If frmConn昭通.Execute("I280", 1, str门诊明细, "正在进行医保交易,请稍候......") = False Then Exit Sub
    For lngLoop = 1 To frmConn昭通.mlngRows
        If frmConn昭通.Query(lngLoop - 1, 1, "正在更新数据(" & lngLoop & "/" & (frmConn昭通.mlngRows) & ")......") = False Then Exit Sub
        str门诊明细 = frmConn昭通.strReturnInfo
        arrData = Split(str门诊明细, vbTab)
        gstrSQL = " Insert Into 门诊明细(流水号,名称,单位,规格,产地,数量,单价,金额)" & _
            " Values ('" & str流水号 & "','" & arrData(0) & "','" & arrData(1) & "','" & arrData(2) & "','" & arrData(3) & "'," & _
            Val(arrData(4)) & "," & Val(arrData(5)) & "," & Val(arrData(6)) & ")"
        gcn昭通.Execute gstrSQL
    Next
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub SaveInExse(ByVal str流水号 As String)
    Dim str住院资料 As String, str住院明细 As String
    Dim arrData
    Dim lngLoop As Long
    
    '删除该流水号的相关数据
    WriteInfo "删除该流水号的相关数据"
    On Error Resume Next
    gcn昭通.Execute "Delete 住院明细 Where 流水号='" & str流水号 & "'"
    gcn昭通.Execute "Delete 住院资料 Where 流水号='" & str流水号 & "'"
    
    On Error GoTo errHand
    
    '读取门诊资料数据
    str住院资料 = str流水号
    WriteInfo "(读取门诊资料数据)交易传递参数:" & str住院资料
    If frmConn昭通.Execute("I348", 0, str住院资料, "正在进行医保交易,请稍候......") = False Then Exit Sub
    If frmConn昭通.Query(0, 1) = False Then Exit Sub
    str住院资料 = frmConn昭通.strReturnInfo
    WriteInfo "交易返回数据:" & str住院资料
    If str住院资料 = "" Then
        MsgBox "返回数据格式错误", vbInformation, "住院结算"
        Exit Sub
    End If
    '保存门诊资料数据
    arrData = Split(str住院资料, vbTab)
    gstrSQL = " Insert Into 住院资料(流水号,单位名称,类别,人员状态,保险证号,姓名,性别,年龄,入院日期,出院日期,科别,床号,疾病名称,押金,帐户支付,缴费日期)" & _
        " Values ('" & str流水号 & "','" & arrData(0) & "','" & arrData(1) & "','" & arrData(2) & "','" & arrData(3) & "','" & arrData(4) & "'," & _
        "'" & arrData(5) & "'," & Val(arrData(6)) & ",'" & arrData(7) & "','" & arrData(8) & "','" & arrData(9) & "'," & _
        "'" & arrData(10) & "','" & arrData(11) & "'," & Val(arrData(12)) & "," & Val(arrData(13)) & ",'" & arrData(14) & "')"
    gcn昭通.Execute gstrSQL
    
    '读取门诊明细数据
    str住院明细 = str流水号 & vbTab & "1"
    WriteInfo "(读取门诊明细数据)交易传递参数:" & str住院明细
    If frmConn昭通.Execute("I348", 1, str住院明细, "正在进行医保交易,请稍候......") = False Then Exit Sub
    For lngLoop = 1 To frmConn昭通.mlngRows
        If frmConn昭通.Query(lngLoop - 1, 1, "正在更新数据(" & lngLoop & "/" & (frmConn昭通.mlngRows) & ")......") = False Then Exit Sub
        str住院明细 = frmConn昭通.strReturnInfo
        arrData = Split(str住院明细, vbTab)
        gstrSQL = " Insert Into 住院明细(流水号,名称,费用,个人负担先付比例,个人负担标准比例,个人负担金额,统筹基金负担)" & _
            " Values ('" & str流水号 & "','" & arrData(0) & "'," & Val(arrData(1)) & "," & Val(arrData(2)) & "," & Val(arrData(3)) & "," & _
            Val(arrData(4)) & "," & Val(arrData(5)) & ")"
        gcn昭通.Execute gstrSQL
    Next
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function Get自费药品(ByVal strCode As String) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '诊疗、甲类或乙类药品均返回假
    
    gstrSQL = "Select lb From tab_syml where upper(dm)='" & strCode & "'"
    Call OpenRecordset_OtherBase(rsCheck, "检查药品类型", gstrSQL, gcn昭通)
    If rsCheck.RecordCount = 0 Then Exit Function   '说明是诊疗项目
    If rsCheck!lb <> 15 Then Exit Function           '说明是甲类、乙类药品
    Get自费药品 = True
End Function
