Attribute VB_Name = "mdl徐州农保"
Option Explicit

Public gcn徐州农保 As New ADODB.Connection

Public Function openConn徐州农保() As Boolean
    Dim rsTemp As New ADODB.Recordset, str参数值 As String, strUser As String, strServer As String, _
        strPass As String, strDatabase As String
    On Error GoTo errHandle
    If gcn徐州农保.State <> adStateOpen Then
        gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_徐州农保)
        
        Do Until rsTemp.EOF
            str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Select Case rsTemp("参数名")
                Case "徐州农保用户名"
                    strUser = str参数值
                Case "徐州农保服务器"
                    strServer = str参数值
                Case "徐州农保用户密码"
                    strPass = str参数值
                Case "徐州农保数据库"
                    strDatabase = str参数值
            End Select
            rsTemp.MoveNext
        Loop
        
        On Error Resume Next
        gcn徐州农保.ConnectionString = "Provider=SQLOLEDB.1;Initial Catalog=" & strDatabase & ";Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
        gcn徐州农保.CursorLocation = adUseClient
        gcn徐州农保.Open

        
        If Err <> 0 Then
            MsgBox "医保前置服务器连接失败！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    openConn徐州农保 = True
    Exit Function

errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 医保初始化_徐州农保() As Boolean
    Dim rsTemp As New ADODB.Recordset, str参数值 As String, strUser As String, strServer As String, _
        strPass As String, strDatabase As String
    
    If openConn徐州农保() = False Then Exit Function
    
    gstrSQL = "Select * From 保险类别 Where 序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保初始化", TYPE_徐州农保)
    gstr医院编码 = Trim(rsTemp!医院编码)
    
    医保初始化_徐州农保 = True
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 身份标识_徐州农保(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能:识别指定人员是否为参保病人，返回病人的信息
'参数:bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回:空或信息串
'注意:1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmIDentified As Object, strPatiInfo As String
    
    WriteInfo vbCrLf & "开始身份验证"
    
    If bytType = 0 Then
        Set frmIDentified = New frmIdentify徐州农保_门诊
    Else
        Set frmIDentified = New frmIdentify徐州农保
    End If
    
    strPatiInfo = frmIDentified.GetPatient(bytType)
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '建立病人档案信息
        lng病人ID = BuildPatiInfo(bytType, strPatiInfo, lng病人ID, TYPE_徐州农保)
        
        '返回格式:中间插入病人ID
        strPatiInfo = frmIDentified.mstrPatient & lng病人ID & ";" & frmIDentified.mstrOther
        Unload frmIDentified
    Else
        身份标识_徐州农保 = ""
        MsgBox "医保病人信息提取失败", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    
    WriteInfo "结束身份验证"
    
    身份标识_徐州农保 = strPatiInfo
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_徐州农保 = ""
End Function

Public Function 个人余额_徐州农保(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select 帐户余额 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取个人帐户余额", lng病人ID, TYPE_徐州农保)
    
    If rsTemp.EOF Then
        个人余额_徐州农保 = 0
    Else
        个人余额_徐州农保 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
    End If
End Function

Public Function 入院登记_徐州农保(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
     '将病人的状态进行修改
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select * From 保险帐户 Where 病人id=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_徐州农保)
    If rsTemp.EOF Then
        MsgBox "保险帐户中没有医保病人信息，不能办理入院登记", vbInformation, gstrSysName
        Exit Function
    ElseIf IsNull(rsTemp!顺序号) Then
        MsgBox "医保病人信息中不能确定病人ID，不能办理入院登记", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_徐州农保 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_徐州农保 = True
End Function

Public Function 入院登记撤消_徐州农保(lng病人ID As Long, lng主页ID As Long) As Boolean
    '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_徐州农保 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记撤消_徐州农保 = True
End Function

Public Function 出院登记_徐州农保(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, strSQL As String, datCurr As Date
    
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select * From 保险帐户 Where 病人id=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_徐州农保)
    
    strSQL = "Update inpatient Set outdate='" & Format(datCurr, "yyyy-mm-dd") & "',mark=1 Where id=" & rsTemp!顺序号
    gcn徐州农保.Execute strSQL
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_徐州农保 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    出院登记_徐州农保 = True
End Function

Public Function 出院登记撤消_徐州农保(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, strSQL As String, datCurr As Date
    
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select * From 保险帐户 Where 病人id=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_徐州农保)
    
    strSQL = "Update inpatient Set outdate=NULL,mark=0 Where id=" & rsTemp!顺序号
    gcn徐州农保.Execute strSQL
    '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_徐州农保 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    出院登记撤消_徐州农保 = True
End Function

Public Function 住院虚拟结算_徐州农保(rs明细 As ADODB.Recordset, lng病人ID As Long) As String
    Dim rsTemp As New ADODB.Recordset, strSQL As String, lng医保ID As Long, str入院 As String, lng主页ID As Long, _
        rs项目 As New ADODB.Recordset, str结帐 As String, datCurr As Date, cur总额 As Currency, int天数 As Integer
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    WriteInfo vbCrLf & "开始上传费用明细"
    
    gstrSQL = "Select * From 保险帐户 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    lng医保ID = rsTemp!顺序号
    str结帐 = Nvl(rsTemp!退休证号, "")
    
    gstrSQL = "Select Max(主页ID) From 病案主页 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    lng主页ID = rsTemp(0)
    
    gstrSQL = "Select 入院日期 From 病案主页 Where 病人id=[1] And 主页id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    str入院 = Format(rsTemp(0), "yyyy-mm-dd")
    int天数 = CDate(Format(datCurr, "yyyy-mm-dd")) - CDate(str入院)
    
    gstrSQL = "Select * From 住院费用记录 Where 门诊标志=2 And 记录状态<>0 And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 and nvl(实收金额,0)<>0 and" & _
        " 病人id=[1] And 主页id=[2] order by 主页ID,序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    
    gcn徐州农保.BeginTrans
    
    While Not rsTemp.EOF
        gstrSQL = "Select nvl(附注,0) As 附注 From 保险支付项目 Where 险类=[1] And 收费细目ID=[2]"
        Set rs项目 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_徐州农保, CLng(rsTemp!收费细目ID))
        If rs项目.EOF Then          '注意询问项目不是医保项目时如何处理
            gstrSQL = "Select * From 收费细目 Where ID=[1]"
            Set rs项目 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rsTemp!收费细目ID))
            MsgBox "项目[" & rs项目!名称 & "]没有对应的医保编码，不能上传费用", vbInformation, gstrSysName
            gcn徐州农保.RollbackTrans
            Exit Function
        Else
            strSQL = "Insert Into infee_mx (id,times,[Date],yp_id,sl,je) values (" & lng医保ID & ",1,'" & _
                Format(rsTemp!发生时间, "yyyy-MM-dd HH:MM:SS") & "'," & rs项目!附注 & "," & rsTemp!数次 * rsTemp!付数 & _
                "," & rsTemp!实收金额 & ")"
            WriteInfo "写费用表:" & strSQL
            gcn徐州农保.Execute strSQL
        End If
        
        rsTemp.MoveNext
    Wend
    
    cur总额 = 0
    While Not rs明细.EOF
        cur总额 = cur总额 + rs明细!金额
        rs明细.MoveNext
    Wend
    strSQL = "Delete From infee Where id=" & lng医保ID
    gcn徐州农保.Execute strSQL
    
    strSQL = "Insert Into infee (ID,times,[jzdate],Days,fee_sum) values (" & lng医保ID & ",1,'" & _
        Format(datCurr, "yyyy-mm-dd HH:MM:SS") & "'," & int天数 & "," & cur总额 & ")"
    WriteInfo "写入汇总:" & strSQL
    
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_徐州农保 & ",'退休证号','''" & Format(datCurr, "yyyy-mm-dd") & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "徐州医保ID")
    
    gcn徐州农保.Execute strSQL
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            gstrSQL = "zl_病人记帐记录_上传 ('" & rsTemp("ID") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            rsTemp.MoveNext
        Wend
    End If
    
    WriteInfo "完成费用明细传递"
    gcn徐州农保.CommitTrans
    
    Call UpdateClass(lng病人ID, lng主页ID)
    
    住院虚拟结算_徐州农保 = "统筹基金;0;0"
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    gcn徐州农保.RollbackTrans
End Function

Public Function 住院结算_徐州农保(lng结帐ID As Long, lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, strSQL As String, datCurr As Date
On Error GoTo ErrH
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select * From 保险帐户 Where 病人id=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_徐州农保)
    
    strSQL = "Update inpatient Set outdate='" & Format(datCurr, "yyyy-mm-dd") & "',mark=2 Where id=" & rsTemp!顺序号
    gcn徐州农保.Execute strSQL
    
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_徐州农保 & "," & lng病人ID & "," & _
        Year(datCurr) & ",0,0,0,0,0,NULL,NULL,NULL,0,0,0,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL,'" & rsTemp!顺序号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "徐州医保")
    
    住院结算_徐州农保 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_徐州农保(lng结帐ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, strSQL As String, datCurr As Date
On Error GoTo ErrH
    
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select 备注,病人ID From 保险结算记录 Where 性质=2 And 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    If IsNull(rsTemp(0)) Then
        MsgBox "结算记录中找不到病人的医保ID，不能进行退结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    strSQL = "Update inpatient Set outdate='" & Format(datCurr, "yyyy-mm-dd") & "',mark=2 Where id=" & rsTemp(0)
    gcn徐州农保.Execute strSQL
    
    gstrSQL = "ZL_保险帐户_更新信息(" & rsTemp!病人ID & "," & TYPE_徐州农保 & ",'顺序号','''" & rsTemp(0) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "徐州医保ID")
    
    住院结算冲销_徐州农保 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function GetItemInfo_徐州(ByVal lngPatiID As Long, ByVal lngItemID As Long, Optional ByVal str摘要 As String, Optional intType As Integer = 0, Optional ByVal blnMsg As Boolean = False) As String
    Dim rsTemp As New ADODB.Recordset, strTemp As String, int险类 As Integer
    
    Dim str类型 As String
    
    '获取当前病人的险类
    gstrSQL = "Select * From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人信息", lngPatiID)
    int险类 = Nvl(rsTemp!险类, 0)
    If int险类 = 0 Then Exit Function
    
    WriteInfo "开始取项目信息(" & int险类 & ")"
    
    gstrSQL = "Select * From 保险支付项目 Where 险类=[1] And 是否医保=1 And 项目名称 Is Not Null And 收费细目ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取保险项目信息", int险类, lngItemID)
    If rsTemp.EOF Then
        MsgBox "该项目没有对码,将按自费项目处理,请注意使用", vbInformation, gstrSysName
        str类型 = "自费"
        GetItemInfo_徐州 = str类型
        Exit Function
    End If
    
    
    Select Case int险类
        Case TYPE_徐州
            strTemp = Nvl(rsTemp!附注, "")
            If InStr(strTemp, "甲") > 0 Or strTemp = "A类诊疗" Then Exit Function
            If strTemp = "" Then
                GetItemInfo_徐州 = "不能确定该项目的医保类别，请注意使用"
            Else
                GetItemInfo_徐州 = "该项目的医保类别为“" & strTemp & "”，请注意使用"
                str类型 = strTemp
            End If
        Case TYPE_徐州农保
            str类型 = "自费"
            Call openConn徐州农保
            strTemp = Nvl(rsTemp!附注, "")
            If strTemp = "" Then
                MsgBox "不能确定该项目的医保类别，请注意使用", vbInformation, gstrSysName
                str类型 = "自费"
                GetItemInfo_徐州 = str类型
                Exit Function
            End If
            WriteInfo "Select * From price_item Where id=" & strTemp
            Set rsTemp = gcn徐州农保.Execute("Select * From price_item Where id=" & strTemp)
            If rsTemp!yp_bz = True Then
                strTemp = rsTemp!GeneralID
                WriteInfo "Select * From GeneralDrug Where General_ID=" & strTemp
                Set rsTemp = gcn徐州农保.Execute("Select CompenSateMark From GeneralDrug Where General_ID=" & strTemp)
                If rsTemp.EOF Then
                    GetItemInfo_徐州 = "请注意：前置服务器中没有找到该项目，不能确定医保类别"
                    str类型 = "自费"
                Else
                    Select Case rsTemp(0)
                        Case 1
                            GetItemInfo_徐州 = "该项目的医保类别为“自费”，请注意使用"
                            str类型 = "自费"
                        Case 2
                            'GetItemInfo_徐州 = "该项目的补偿范围为“村级”，请注意使用"
                            str类型 = "村级"
                        Case 3
                            'GetItemInfo_徐州 = "该项目的补偿范围为“乡级”，请注意使用"
                            str类型 = "乡级"
                        Case 4
                            'GetItemInfo_徐州 = "该项目的补偿范围为“县级”，请注意使用"
                            str类型 = "县级"
                        Case 5
                            'GetItemInfo_徐州 = "该项目的补偿范围为“市级”，请注意使用"
                            str类型 = "市级"
                        Case 6
                            'GetItemInfo_徐州 = "该项目的补偿范围为“省级”，请注意使用"
                            str类型 = "省级"
                    End Select
                End If
            Else
                strTemp = rsTemp!CenterID
                Set rsTemp = gcn徐州农保.Execute("Select CompenSationMark From FeeItemList Where ID=" & strTemp)
                If rsTemp.EOF Then
                    'GetItemInfo_徐州 = "请注意：前置服务器中没有找到该项目，不能确定医保类别"
                    str类型 = "自费"
                Else
                    Select Case rsTemp(0)
                        Case 1
                            GetItemInfo_徐州 = "该项目的医保类别为“自费”，请注意使用"
                            str类型 = "自费"
                        Case 2
                            str类型 = "甲类"
                        Case 3
                            GetItemInfo_徐州 = "该项目的医保类别为“乙类”，请注意使用"
                            str类型 = "乙类"
                    End Select
                End If
            End If
        Case TYPE_徐州市
        
    End Select
    If blnMsg = True Then
        If GetItemInfo_徐州 <> "" Then MsgBox GetItemInfo_徐州, vbInformation, gstrSysName
    End If
    GetItemInfo_徐州 = str类型
End Function

Public Function 门诊虚拟结算_徐州农保(rs明细 As ADODB.Recordset, str结算方式 As String, Optional ByRef strAdvance As String = "") As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '检查是否存在未对码的项目
    
    With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            gstrSQL = "Select nvl(附注,0) As 附注 From 保险支付项目 Where 险类=[1] And 收费细目ID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保项目ID", TYPE_徐州农保, CLng(!收费细目ID))
            If rsTemp.RecordCount = 0 Then         '注意询问项目不是医保项目时如何处理
                gstrSQL = "Select * From 收费细目 Where ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取HIS项目编码与名称", CLng(!收费细目ID))
                MsgBox "项目[" & rsTemp!名称 & "]没有对应的医保编码，不能上传费用", vbInformation, gstrSysName
                Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    '什么都不做，直接返回
    str结算方式 = "个人帐户;0;0"
    门诊虚拟结算_徐州农保 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_徐州农保(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, Optional ByRef strAdvance As String = "") As Boolean
    Dim lng病人ID As Long
    Dim dbl总额 As Double
    Dim str医保序号 As String
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    Call DebugTool("门诊结算")
    '提取所有费用明细,供产生明细表用
    gstrSQL = "Select * From 门诊费用记录 Where 结帐ID=[1]"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人ID", lng结帐ID)
    lng病人ID = rs明细!病人ID
    '2006-4-27 陈玉强反馈,不需要此票号,要求取消
'    str发票号 = Nvl(rs明细!实际票号)
'    If str发票号 = "" Then
'        Err.Raise 9000,gstrSysName, "发票号不能为空！"
'        Exit Function
'    End If
    
    gcn徐州农保.BeginTrans
    Call DebugTool("准备上传...")
    '上传处方明细
    With rs明细
        Call DebugTool("计算总额")
        dbl总额 = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dbl总额 = dbl总额 + Val(Format(Nvl(!实收金额, 0), "#0.00;-#0.00;0;"))
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        Call DebugTool("插入门诊总表")
        '    先插入门诊总表
        gstrSQL = "Insert into MZ11(id,xm,[date],tp_bz,je) " & _
            " Values (" & !结帐ID & ",'" & Nvl(!姓名) & "','" & _
            Format(!发生时间, "yyyy-MM-dd HH:MM:SS") & "',0" & _
            "," & dbl总额 & ")"
        gcn徐州农保.Execute gstrSQL
    End With
    
    With rs明细
        Call DebugTool("插入门诊明细表")
        Do While Not .EOF
            gstrSQL = "Select 附注 From 保险支付项目 Where 险类=[1] And 收费细目ID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取医保项目序号", TYPE_徐州农保, CLng(!收费细目ID))
            str医保序号 = Nvl(rsTemp!附注)
            
            '    再插入门诊明细表
            gstrSQL = "Insert Into MZ22(mz11_id,yp_id,sl,je) " & _
                " Values (" & lng结帐ID & "," & str医保序号 & "," & _
                !数次 * Nvl(!付数, 1) & "," & Format(Nvl(!实收金额, 0), "#0.00;-#0.00;0;") & ")"
            gcn徐州农保.Execute gstrSQL
            .MoveNext
        Loop
    End With
    
    Call DebugTool("保存保险结算记录")
    '保存保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_徐州农保 & "," & lng病人ID & "," & _
        Year(zlDatabase.Currentdate) & ",0,0,0,0,0,NULL,NULL,NULL,0,0,0,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL,'" & lng结帐ID & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "徐州医保")
    
    gcn徐州农保.CommitTrans
    门诊结算_徐州农保 = True
    Exit Function
errHand:
    gcn徐州农保.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_徐州农保(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, Optional ByRef strAdvance As String = "") As Boolean
    Dim lng冲销ID As Long
    Dim dbl总额 As Double
    Dim str医保序号 As String
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
On Error GoTo errHand
    
    Call DebugTool("门诊退费")
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!结帐ID
    
    Call DebugTool("提取所有费用明细")
    '提取所有费用明细,供产生明细表用
    gstrSQL = "Select * From 门诊费用记录 Where 结帐ID=[1]"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人ID", lng冲销ID)
'    str发票号 = Nvl(rs明细!实际票号)
    
    gcn徐州农保.BeginTrans
    Call DebugTool("准备上传明细...")
    '上传处方明细
    With rs明细
        Call DebugTool("计算总额")
        dbl总额 = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dbl总额 = dbl总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        Call DebugTool("插入门诊总表")
        '    先插入门诊总表
        gstrSQL = "Insert into MZ11(id,xm,[date],tp_bz,je) " & _
            " Values (" & lng冲销ID & ",'" & Nvl(!姓名) & "','" & _
            Format(!发生时间, "yyyy-MM-dd HH:MM:SS") & "',0" & _
            "," & dbl总额 & ")"
        gcn徐州农保.Execute gstrSQL
    End With
    
    Call DebugTool("插入门诊明细表")
    With rs明细
        Do While Not .EOF
            gstrSQL = "Select 附注 From 保险支付项目 Where 险类=[1] And 收费细目ID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取医保项目序号", TYPE_徐州农保, CLng(!收费细目ID))
            str医保序号 = Nvl(rsTemp!附注)
            
            '    再插入门诊明细表
            gstrSQL = "Insert Into MZ22(mz11_id,yp_id,sl,je) " & _
                " Values (" & lng冲销ID & "," & str医保序号 & "," & _
                !数次 * Nvl(!付数, 1) & "," & !实收金额 & ")"
            gcn徐州农保.Execute gstrSQL
            .MoveNext
        Loop
    End With
    
    Call DebugTool("保存保险结算记录")
    '保存保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_徐州农保 & "," & lng病人ID & "," & _
        Year(zlDatabase.Currentdate) & ",0,0,0,0,0,NULL,NULL,NULL,0,0,0,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL,'" & lng冲销ID & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "徐州医保")
    
    gcn徐州农保.CommitTrans
    门诊结算冲销_徐州农保 = True
    Exit Function
errHand:
    gcn徐州农保.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Private Sub UpdateClass(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    Dim str费用类型 As String
    Dim rsTemp As New ADODB.Recordset
    '循环更新所有项目的费用类型
    gstrSQL = "Select ID,病人ID,收费细目ID,费用类型 From 住院费用记录" & _
        " Where 病人ID=[1] And 主页ID=[2]" & _
        " And Nvl(是否上传,0)=1 And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0 And Nvl(实收金额,0)<>0 And 费用类型 is null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "循环更新所有项目的费用类型", lng病人ID, lng主页ID)
    
    With rsTemp
        Do While Not .EOF
            str费用类型 = GetItemInfo_徐州(!病人ID, !收费细目ID, "", 0, False)
            gstrSQL = "ZL_病人记帐记录_上传(" & !ID & ",NULL,NULL,'" & str费用类型 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新费用类型")
            .MoveNext
        Loop
    End With
End Sub
