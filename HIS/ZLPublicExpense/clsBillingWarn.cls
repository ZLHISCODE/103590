VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBillingWarn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mrsWarn As ADODB.Recordset
Private mrsFeeType As ADODB.Recordset
Public Enum gEM_WarnType
    EM_NotWarn = 0  '没有报警,继续
    EM_Warn_Qury_OK = 1  '报警提示后用户选择继续
    EM_Warn_Qury_Cancel = 2  '报警提示后用户选择中断
    EM_Warn_Cancel = 3  '报警提示必须中断
    EM_Warn_Billing = 4  '强制记帐报警,继续
    EM_Warn_SavePrice = 5  '报警提示后用户选择继续,但只允许保存存为划价单
End Enum
Private Function GetRsWarng() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取报警线
    '返回:返回报警线记录集
    '编制:刘兴洪
    '日期:2014-03-19 11:42:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errH
    If Not mrsWarn Is Nothing Then
        If mrsWarn.State = 1 Then Set GetRsWarng = mrsWarn: Exit Function
    End If
    
    strSql = "" & _
    "   Select Nvl(病区ID,0) 病区ID,适用病人,Nvl(报警方法,1) as 报警方法," & _
    "           报警值,报警标志1,报警标志2,报警标志3" & _
    "   From 记帐报警线 "
    Set mrsWarn = gobjDatabase.OpenSQLRecord(strSql, "缓存记帐报警线的相关设置")
    Set GetRsWarng = mrsWarn
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Private Function GetRsFeeType() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费类别
    '编制:刘兴洪
    '日期:2014-03-25 09:40:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errHandle
    If mrsFeeType Is Nothing Then
        strSql = "Select 编码,名称,简码 From 收费项目类别"
    ElseIf mrsFeeType.State <> 1 Then
        strSql = "Select 编码,名称,简码 From 收费项目类别"
    Else
        Set GetRsFeeType = mrsFeeType
        Exit Function
    End If
    Set mrsFeeType = gobjDatabase.OpenSQLRecord(strSql, "缓存收费类别")
    Set GetRsFeeType = mrsFeeType
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetPati(ByVal lng病人ID As Long, ByVal lng主页Id As Long) As clsPatiInfor
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息给病人信息对象
    '入参:lng病人ID-病人ID
    '     lng主页ID-主页ID
    '返回:获取病人信息给病人信息对象
    '编制:刘兴洪
    '日期:2014-03-25 10:48:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim objPati As New clsPatiInfor
    Dim objPatiFeeInfor As New clsPatiFeeinfor
    On Error GoTo errHandle
    strSql = "" & _
    "   Select 病人ID,姓名,性别,年龄,住院号,当前床号, " & _
    "          Zl_Patiwarnscheme([1], [2]) as 报警方案," & _
    "          Zl_Patidaycharge(病人ID) as 当日额 " & _
    "   From 病人信息 " & _
    "   Where 病人ID=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取病人基本信息", lng病人ID, lng主页Id)
    If Not rsTemp.EOF Then
        With objPati
            .病人ID = Val(Nvl(rsTemp!病人ID))
            .报警方案 = Nvl(rsTemp!报警方案)
            .年龄 = Nvl(rsTemp!年龄)
            .性别 = Nvl(rsTemp!性别)
            .姓名 = Nvl(rsTemp!姓名)
            .住院号 = Nvl(rsTemp!住院号)
            .床号 = Nvl(rsTemp!当前床号)
            .当日记帐总额 = FormatEx(Val(Nvl(rsTemp!当日额)), 4)
            Call GetPatiMoney(IIf(lng主页Id <> 0, 1, 0), lng病人ID, objPatiFeeInfor)
            If objPatiFeeInfor Is Nothing Then
                Set .Obj费用概况 = New clsPatiFeeinfor
            Else
                Set .Obj费用概况 = objPatiFeeInfor
            End If
        End With
    End If
    Set GetPati = objPati
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetFeeTypeName(ByVal str类别编码 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据类别编码获取类别名称
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-25 09:51:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    Set rsTemp = GetRsFeeType
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    rsTemp.Filter = "编码='" & str类别编码 & "'"
    If Not rsTemp.EOF Then
        GetFeeTypeName = Nvl(rsTemp!名称)
    End If
    rsTemp.Filter = 0
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetPatiDayCharge(ByVal lng病人ID As Long, _
    ByVal dtDate As Date, ByVal bln含划单 As Boolean, _
    ByRef dblOutDayTotal As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定日期的记帐费用总额(获取日费用)
    '入参:lng病人ID-指定的病人ID
    '     bln含划单-是否包含划价单费用
    '     dtDate-指定日期
    '出参:dblDayTotal-返回费用总额
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-19 10:16:10
    '说明:一般情况下,主要应用于报警
    '   1.住院医生站
    '   2.医技工作站
    '   3.住院护士站 (医嘱发送)
    '   4.费用所有记帐模块
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset, int划价 As Integer
    On Error GoTo errHandle
    dblOutDayTotal = 0
    int划价 = IIf(bln含划单, 1, 0)
    strSql = "Select Zl_Patidaycharge([1],[2],[3]) as 费用额 From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取指定日期的费用总额", lng病人ID, dtDate, int划价)
    dblOutDayTotal = Val(gobjCommFun.Nvl(rsTemp!费用额))
    zlGetPatiDayCharge = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiWarnScheme(ByVal lng病人ID As Long, _
    ByVal lng主页Id As Long, _
    ByRef strOut适用病人 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取报警方式的适用病人
    '入参:lng病人ID-指定的病人ID
    '     lng主页ID-主页ID,允许传入零,表示门诊;如果>0时，表示第几次住院
    '出参:strOut适用病人-返回适用病人
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-19 10:16:10
    '说明:一般情况下,主要应用于报警
    '调用模块:
    '   1.住院医生站
    '   2.医技工作站
    '   3.住院护士站 (医嘱发送)
    '   4.费用所有记帐模块
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strOut适用病人 = ""
    strSql = "Select Zl_Patiwarnscheme([1], [2]) as 适用病人 From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取报警方式的适用病人", lng病人ID, lng主页Id)
    strOut适用病人 = gobjCommFun.Nvl(rsTemp!适用病人)
    zlGetPatiWarnScheme = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlBillingVerfyWarnCheck(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal str医嘱IDs As String, _
    Optional ByVal strNos As String = "", _
    Optional ByVal strPrivs As String = "", _
    Optional lng病区ID As Long = -1) As Boolean
    
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记帐审核时报警检查
    '入参: frmMain-调用的父窗口
    '      lngModule-模块号
    '      str医嘱IDs-允许传入多个,用逗号分隔,如:1,2,3,4(目前只针对临嘱),如果为长嘱,哪只能传NO来处理
    '      Nos-单据号,格式为:性质(1-门诊;2-住院)|NO1,…NOn),如:1|A00001,A0002,...,A000n
    '      lng病区ID-传入病区ID(如果未传,则以住院费用记录.病区ID为主):-1表示传入
    '      strPrivs-如果是门诊时,需要传入(否则直接取1150模块的权限)
    '               但必须保证该模块有"欠费强制记帐"的权限
    '出参:
    '返回:接口返回true或False
    '       1.true时:数据检查合法(含提示选择继续)
    '       2.False时检查不合法(包含:提示选择否,或禁止提示或内部出错)
    '编制:刘兴洪
    '日期:2014-03-19 11:30:48
    '说明:对病人记帐费用进行审核前的费用报警检查(含提示,禁止等)
    '调用模块:
    '    1.采集工作站
    '    2.技师工作站
    '    3.标本签收
    '    3.实验室管理
    '    4.Pacs
    '    5.药品处方发药 (发药)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim strSql As String, varTemp As Variant, objPatiFee As clsPatiFeeinfor
    Dim rsTemp As ADODB.Recordset, strTempNos As String, strTittle As String
    Dim lng病人ID As Long, lng主页Id As Long, dblWarnMoney As Double
    Dim str姓名 As String, str适用病人 As String, strDec As String
    Dim rsWarn As ADODB.Recordset, rsPaiti As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '获取报警集
    Set rsWarn = GetRsWarng
    If rsWarn Is Nothing Then Exit Function
    If rsWarn.State <> 1 Then Exit Function
    
    strDec = gSysPara.Money_Decimal.strFormt_VB
    
    If strPrivs = "" Then strPrivs = ";" & gobjComlib.GetPrivFunc(glngSys, 1150) & ";"
    
    '报警检查的医嘱ID或NOs应该不会造成4000个字符,所以不用作分次处理
    If str医嘱IDs <> "" Then
        strSql = "" & _
        "   Select /*+ RULE */  M.类别,J.类别 as 名称,M.病区ID Sum(m.实收金额) As 金额,m.病人ID,m.主页ID " & _
        "    From (With 医嘱数据 As (Select Column_Value As 医嘱id From Table(f_Num2list([1]))) " & _
        "        Select A.病人ID,0 as 主页ID,a.收费类别 as 类别,0 as 病区ID, Nvl(Sum(A.实收金额), 0) As 实收金额 " & _
        "        From 门诊费用记录 A,  医嘱数据 B " & _
        "        Where a.医嘱序号 = b.医嘱id And 记帐费用 = 1 And 记录状态 = 0 " & _
        "        Group By a.收费类别,A.病人ID " & _
        "        Union All " & _
        "        Select A.病人ID,A.主页ID,a.收费类别 as 类别,A.病人病区ID as 病区ID, Nvl(Sum(实收金额), 0) As 实收金额 " & _
        "        From 住院费用记录 A, 医嘱数据 B " & _
        "        Where a.医嘱序号 = b.医嘱id And 记帐费用 = 1 And 记录状态 = 0 " & _
        "        Group By a.收费类别,A.病人ID,A.主页ID,A.病人病区ID  ) M,收费类别 J  " & _
        " Where m.类别=J.编码" & _
        " Group By m.病人ID,m.主页ID,M.类别,J.类别,m.病区ID"
        strTittle = "医嘱"
    ElseIf InStr(strNos, "|") = 0 Then
        If strNos = "" Then Exit Function
        strSql = "" & _
        "   Select /*+ RULE */  M.类别,J.类别 as 名称,M.病区ID, Sum(m.实收金额) As 金额,m.病人ID,m.主页ID " & _
        "    From ( " & _
        "        Select A.病人ID,0 as 主页ID,a.收费类别 as 类别,0 as 病区ID, Nvl(Sum(A.实收金额), 0) As 实收金额 " & _
        "        From 门诊费用记录 A,  table(f_str2List([2])) B " & _
        "        Where a.NO = b.Column_Value And 记帐费用 = 1  And 记录状态 = 0 And 记录性质=2 " & _
        "        Group By a.收费类别,A.病人ID " & _
        "        Union All " & _
        "        Select A.病人ID,A.主页ID,a.收费类别 as 类别,A.病人病区ID as 病区ID, Nvl(Sum(A.实收金额), 0) As 实收金额 " & _
        "        From 住院费用记录 A,  table(f_str2List([2])) B  " & _
        "        Where a.NO = b.Column_Value And 记帐费用 = 1  And 记录状态 = 0 And 记录性质=2 " & _
        "        Group By a.收费类别,A.病人ID,A.主页ID,A.病人病区ID ) M,收费类别 J  " & _
        " Where m.类别=J.编码" & _
        " Group By m.病人ID,m.主页ID,M.类别,J.类别,M.病区ID"
        strTittle = "记帐单"
        strTempNos = strNos
    Else
        varTemp = Split(strNos, "|")
        strTempNos = varTemp(1)
        If Val(varTemp(0)) = 1 Then
            '门诊
            strSql = "" & _
            "        Select  /*+ RULE */ A.病人ID,0 as 主页ID,a.收费类别 as 类别,J.类别 as 名称,0 as 病区ID, " & _
            "               Nvl(Sum(A.实收金额), 0) As 金额 " & _
            "        From 门诊费用记录 A,  table(f_str2List([3])) B,收费类别 J " & _
            "        Where a.NO = b.Column_Value And 记帐费用 = 1  And 记录状态 = 0 And 记录性质=2 " & _
            "        Group By a.收费类别,A.病人ID,J.类别 "
            strTittle = "门诊记帐单"
        Else
            strSql = "" & _
            "        Select  /*+ RULE */ A.病人ID,0 as 主页ID,a.收费类别 as 类别,J.类别 as 名称,A.病人病区ID as 病区ID, " & _
            "               Nvl(Sum(A.实收金额), 0) As 金额 " & _
            "        From 住院费用记录 A,  table(f_str2List([3])) B,收费类别 J " & _
            "        Where a.NO = b.Column_Value And 记帐费用 = 1  And 记录状态 = 0 And 记录性质=2 " & _
            "        Group By a.收费类别,A.病人ID,J.类别,A.病人病区ID"
            strTittle = "住院记帐单"
        End If
    End If
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取相关报警类别金额", str医嘱IDs, strNos, strTempNos)
    If rsTemp.EOF Then
        MsgBox "未找到指定的" & strTittle & "数据,请检查传入的医嘱是否正确!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    lng病人ID = Val(Nvl(rsTemp!病人ID)): lng主页Id = Val(Nvl(rsTemp!主页ID))
    
    strSql = "Select 姓名 From 病人信息 Where 病人ID=[1]"
    Set rsPaiti = gobjDatabase.OpenSQLRecord(strSql, "获取病人姓名", lng病人ID)
    If rsPaiti.EOF Then
        MsgBox "未找到指定的病人信息", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    str姓名 = Nvl(rsPaiti!姓名)
    
    
    '获取病人的应用类别
    If zlGetPatiWarnScheme(lng病人ID, lng主页Id, str适用病人) = False Then Exit Function
    
    If lng病区ID < 0 Then lng病区ID = rsTemp!病区ID
    
    rsWarn.Filter = "适用病人='" & str适用病人 & "' And 病区ID=" & lng病区ID
    If rsWarn.EOF Then zlBillingVerfyWarnCheck = True: Exit Function   '未设置,直接返回true
    If IsNull(rsWarn!报警值) Then zlBillingVerfyWarnCheck = True: Exit Function
    
    Dim dblMoney As Double, dbl余额 As Double
    Dim byt标志 As Integer, str类别名称 As String
    
 
    '对应类别定位有效报警设置
    With rsTemp
        byt标志 = -1
        Do While Not .EOF
            If byt标志 = -1 Then
                If Not IsNull(rsWarn!报警标志1) Then
                    If Nvl(rsWarn!报警标志1) = "-" Or InStr(rsWarn!报警标志1, Nvl(!类别)) > 0 Then byt标志 = 1
                    If Nvl(rsWarn!报警标志1) = "-" Then str类别名称 = ""
                    If InStr(rsWarn!报警标志1, Nvl(!类别)) > 0 Then str类别名称 = Nvl(!名称)
                End If
                If byt标志 = -1 And Not IsNull(rsWarn!报警标志2) Then
                    If Nvl(rsWarn!报警标志2) = "-" Or InStr(rsWarn!报警标志2, Nvl(!类别)) > 0 Then byt标志 = 2
                    If Nvl(rsWarn!报警标志2) = "-" Then str类别名称 = ""
                    If InStr(rsWarn!报警标志2, Nvl(!类别)) > 0 Then str类别名称 = Nvl(!名称)
                End If
                If byt标志 = -1 And Not IsNull(rsWarn!报警标志3) Then
                    If Nvl(rsWarn!报警标志3) = "-" Or InStr(rsWarn!报警标志3, Nvl(!类别)) > 0 Then byt标志 = 3
                    If Nvl(rsWarn!报警标志3) = "-" Then str类别名称 = ""
                    If InStr(rsWarn!报警标志3, Nvl(!类别)) > 0 Then str类别名称 = Nvl(!名称)
                End If
            End If
            dblMoney = dblMoney + Val(Nvl(!金额))
            .MoveNext
        Loop
    End With
    If byt标志 = 0 Then Exit Function '无有效设置
    If str类别名称 <> "" Then str类别名称 = """" & str类别名称 & """费用"
    
    If GetPatiMoney(IIf(lng主页Id = 0 And lng病区ID = 0, 0, 1), lng病人ID, objPatiFee) = False Then Exit Function
    
    '报警处理
    dblWarnMoney = FormatEx(Val(Nvl(rsWarn!报警值)), 6)
    dbl余额 = objPatiFee.剩余款
    
    If gSysPara.bln报警包含划价费用 Then
        '需要统计划价单总额,但需要排除本次检查的划价单额
        dbl余额 = dbl余额 + dblMoney - GetPriceMoneyTotal(IIf(lng主页Id = 0 And lng病区ID = 0, 0, 1), lng病人ID)
    End If

    Select Case Val(Nvl(rsWarn!报警方法))
    Case 1 '累计费用报警:指按照病人当前所有未结费用累计进行报警
        Select Case byt标志
        Case 1   '报警标志1:低于报警值(包括预交款耗尽)提示询问记帐
            If objPatiFee.剩余款 + objPatiFee.担保额 - dblMoney < dblWarnMoney Then
                If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                   If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(objPatiFee.担保额, "0.00") & "):" & Format(objPatiFee.剩余款 + objPatiFee.担保额 - dblMoney, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(dblWarnMoney, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                       Exit Function
                   End If
                Else
                   MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(objPatiFee.担保额, "0.00") & "):" & Format(objPatiFee.剩余款 + objPatiFee.担保额 - dblMoney, "0.00") & " 低于" & str类别名称 & "报警值:" & Format(dblWarnMoney, "0.00") & "。", vbInformation, gstrSysName
                End If
            End If
        Case 2   '报警标志2:低于报警值提示询问记帐,预交款耗尽时禁止记帐
            If objPatiFee.剩余款 + objPatiFee.担保额 - dblMoney < 0 Then
                If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                    MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(objPatiFee.担保额, "0.00") & ")已经耗尽," & str类别名称 & "禁止记帐。", vbInformation, gstrSysName
                    Exit Function
                End If
                MsgBox str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(objPatiFee.担保额, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
            ElseIf objPatiFee.剩余款 + objPatiFee.担保额 - dblMoney < dblWarnMoney Then
                If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                    If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(objPatiFee.担保额, "0.00") & "):" & Format(objPatiFee.剩余款 + objPatiFee.担保额 - dblMoney, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(dblWarnMoney, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                Else
                    MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(objPatiFee.担保额, "0.00") & "):" & Format(objPatiFee.剩余款 + objPatiFee.担保额 - dblMoney, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(dblWarnMoney, "0.00") & "。", vbInformation, gstrSysName
                End If
            End If
        Case 3   '报警标志3:低于报警值禁止记帐
            If objPatiFee.剩余款 + objPatiFee.担保额 - dblMoney < dblWarnMoney Then
                If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                    MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(objPatiFee.担保额, "0.00") & "):" & Format(objPatiFee.剩余款 + objPatiFee.担保额 - dblMoney, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(dblWarnMoney, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                    Exit Function
                End If
                MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(objPatiFee.担保额, "0.00") & "):" & Format(objPatiFee.剩余款 + objPatiFee.担保额 - dblMoney, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(dblWarnMoney, "0.00") & "。", vbInformation, gstrSysName
            End If
        End Select
    Case 2  '每日费用报警:按照费用发生时间每日报警
        Dim dblDayMoney As Double
        If zlGetPatiDayCharge(lng病区ID, gobjDatabase.Currentdate, gSysPara.bln报警包含划价费用, dblDayMoney) = False Then Exit Function
        
        
        Select Case byt标志
        Case 1   '报警标志1:高于报警值提示询问记帐
            If dblDayMoney + dblMoney > dblWarnMoney Then
                If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                    If MsgBox(str姓名 & " 当日费用:" & Format(dblDayMoney + dblMoney, strDec) & ",高于" & str类别名称 & "报警值:" & Format(dblWarnMoney, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                Else
                    MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(dblDayMoney + dblMoney, strDec) & ",高于" & str类别名称 & "报警值:" & Format(dblWarnMoney, "0.00") & "。", vbInformation, gstrSysName
                End If
            End If
        Case 2   '报警标志2:不处理
        Case 3   '报警标志3:高于报警值禁止记帐
            If dblDayMoney + dblMoney > dblWarnMoney Then
                If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                    MsgBox str姓名 & " 当日费用:" & Format(dblDayMoney + dblMoney, strDec) & ",高于" & str类别名称 & "报警值:" & Format(dblWarnMoney, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                    Exit Function
                End If
                MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(dblDayMoney + dblMoney, strDec) & ",高于" & str类别名称 & "报警值:" & Format(dblWarnMoney, "0.00") & "。", vbInformation, gstrSysName
            End If
        End Select
    End Select
    zlBillingVerfyWarnCheck = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlBillingWarnCheck(ByVal frmMain As Object, ByVal bytType As Byte, ByVal bytBillingType As Byte, _
    ByVal lng病人ID As Long, ByVal lng主页Id As Long, ByVal lng病区ID As Long, _
    ByVal str类别编码 As String, ByVal dbl单据金额 As Double, ByVal bln允许强制记帐 As Boolean, _
    ByVal bln允许保存划价单 As Boolean, ByVal blnClearPatiCache As Boolean, _
    ByRef intOutType As Integer, Optional ByVal bln多病人 As Boolean, _
    Optional ByVal dblItemMoney As Double = 0, _
    Optional ByVal blnNotCheck类别 As Boolean = False, _
    Optional ByVal bln报警询问 As Boolean = False, _
    Optional ByVal blnMsgShowChkSkip As Boolean = False, _
    Optional ByRef blnOutSkip As Boolean = False, _
    Optional ByRef intPreMsgButton As Integer = -1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对病人记帐进行报警提示
    '入参:frmMain-调用的主窗体
    '     bytType-0-门诊;1-住院
    '     bytBillingType - 0-记帐,1-划价,2-审核
    '     lng主页ID-主页ID:发送为门诊的,传为零,否则指定次数
    '     lng病区ID=病人病区ID,用于选择适用的病区报警设置，0表示没有确定病区，仅所有病区的报警设置适用
    '     dbl单据金额=病人单据中输入的费用
    '     str类别编码=当前要检查的类别,用于分类报警
    '     bln允许保存划价单-主要是确定当不存在欠费强制权限时，是否允许将记帐单保存为划价单( 即:欠费时是否允许强制保存为划价单,用于记帐或划价有)
    '     blnClearPatiCache-清除缓存数据(改变病人时传入,主要是清除缓存中的病人的预额，当日额，担保额等)
    '     dblItemMoney-当笔金额(如果传入<>0 ,则需要判断当笔情况,如果超出金额,则允许用户继续,否则根据报警方式进行):刘兴洪:24491
    '     blnNotCheck类别:不对类别进行检查(主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
    '     bln报警询问-在门诊（没有权限控制）或住院具备强制记账时，采用询问方式提示
    '     blnMsgShowChkSkip-是否以ShowMsgbox 显示,主要是临床批量记帐有效，其他传为False)
    '     intPreMsgButton-参数blnMsgShowChkSkip=true有效(主要是临床批量记帐判断)：-1表示未曾选择过.>=0时，表示不再提示，发生相关错误时，按传入的值返回

    '出参:
    '     blnOutSkip-blnMsgShowChkSkip=true时有效，选择的是"按相同方式处理"
    '     str已报类别-返回已经报警的类别
    '     intOutType-报警返回的类别
    '       0;没有报警,继续
    '       1:报警提示后用户选择继续
    '       2:报警提示后用户选择中断
    '       3:报警提示必须中断
    '       4:强制记帐报警,继续
    '       5.报警提示后用户选择继续,但只允许保存存为划价单
    '    intPreMsgButton-发生提醒时，返回当前提示框选中的按钮值
    '返回:检查成功，返回true,否则返回失败(请结合出参:intOutType来进行判断)
    '     str报警类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
    '编制:刘兴洪
    '日期:2017-12-12 16:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, byt标志 As Byte, bln已报警 As Boolean
    Dim byt方式 As Byte, byt已报方式 As Byte, arrTmp As Variant
    Dim str类别名称 As String, dbl余额 As Double, strMsgInfor As String
    Dim rsWarn As ADODB.Recordset
    Dim strDec As String, intMsgButton As Integer
    Static str已报类别 As String
    Static objPati As clsPatiInfor
    
    On Error GoTo errH
    
    intOutType = EM_NotWarn
    Set rsWarn = GetRsWarng
    If rsWarn Is Nothing Then Exit Function
    If rsWarn.State <> 1 Then Exit Function
    
    strDec = gSysPara.Money_Decimal.strFormt_VB
    
    '主要缓存上次的病人信息,避免重复读取
    If objPati Is Nothing Or blnClearPatiCache Then
        Set objPati = New clsPatiInfor
        str已报类别 = ""
    End If
    If objPati.病人ID = 0 Or objPati.病人ID <> lng病人ID Then
        '重新读取病人信息
        Set objPati = GetPati(lng病人ID, lng主页Id)
        If Not bln多病人 Then str已报类别 = ""
        If gSysPara.bln报警包含划价费用 Then
            '报警是否含划价单
            With objPati.Obj费用概况
                .剩余款 = .剩余款 - GetPriceMoneyTotal(bytType, lng病人ID)
            End With
        End If
    End If
    
    
    '报警参数检查
    rsWarn.Filter = "适用病人='" & objPati.报警方案 & "' And 病区ID=" & IIf(bytType = 0, 0, lng病区ID)
    If rsWarn.EOF Then zlBillingWarnCheck = True: Exit Function
    
    If IsNull(rsWarn!报警值) Then zlBillingWarnCheck = True: Exit Function
    str类别名称 = GetFeeTypeName(str类别编码)
         
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str类别编码) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志1 <> "-" And blnNotCheck类别 Then zlBillingWarnCheck = True: Exit Function
    End If
    
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str类别编码) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志2 <> "-" And blnNotCheck类别 Then zlBillingWarnCheck = True:    Exit Function
    End If
    
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str类别编码) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志3 <> "-" And blnNotCheck类别 Then zlBillingWarnCheck = True:   Exit Function
    End If
    If byt标志 = 0 Then zlBillingWarnCheck = True:   Exit Function '无有效设置
    
    '报警标志2实际上是两种判断①②,其它只有一种判断①
    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
    If bln多病人 Then
        '示例：",周:-,张:DEF,李:567,张567"
        '报警标志2示例：",周:-①,张:DEF①,李:567①,张567②"
        bln已报警 = str已报类别 & "," Like "*," & objPati.姓名 & ":-*,*" _
            Or str已报类别 & "," Like "*," & objPati.姓名 & ":*" & str类别编码 & "*,*"
    Else
        '示例："-" 或 ",ABC,567,DEF"
        '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
        bln已报警 = InStr(str已报类别, str类别编码) > 0 Or str已报类别 Like "-*"
    End If
    
    If bln已报警 Then
        If byt标志 <> 2 Then zlBillingWarnCheck = True: Exit Function
        
        If bln多病人 Then
            arrTmp = Split(str已报类别, ",")
            For i = 0 To UBound(arrTmp)
                If "," & arrTmp(i) & "," Like "*," & objPati.姓名 & ":-*,*" _
                    Or "," & arrTmp(i) & "," Like "*," & objPati.病人ID & ":*" & str类别编码 & "*,*" Then
                    byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                    '如果已报方式存在两种:低于报警值时,预交耗尽时,则只需最后一种,因为最后一种总是最后发生
                    'Exit For
                End If
            Next
        Else
            If str已报类别 Like "-*" Then
                byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
            Else
                arrTmp = Split(str已报类别, ",")
                For i = 0 To UBound(arrTmp)
                    If InStr(arrTmp(i), str类别编码) > 0 Then
                        byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                        '如果已报类别存在两种:低于报警值时,预交耗尽时,则只需最后一种,因为最后一种总是最后发生
                        'Exit For
                    End If
                Next
            End If
        End If
    End If
    If str类别名称 <> "" Then str类别名称 = """" & str类别名称 & """费用"

    '---------------------------------------------------------------------
    dbl余额 = objPati.Obj费用概况.剩余款 + objPati.Obj费用概况.担保额
    
    Select Case Val(Nvl(rsWarn!报警方法))
    Case 1 '累计费用报警(低于)
        Select Case byt标志
            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                
                If dbl余额 - dbl单据金额 < rsWarn!报警值 Then
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If dblItemMoney <> 0 And dbl余额 - (dbl单据金额 - dblItemMoney) > Val(Nvl(rsWarn!报警值)) Then
                         '只提示: bytBillingType As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        
                        strMsgInfor = "注意" & vbCrLf & _
                                       "    病人“" & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & "” 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ", 是否继续录入该项目？"
                                       
                        If blnMsgShowChkSkip Then
                            If intPreMsgButton = -1 Then
                                intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                            Else
                                intMsgButton = intPreMsgButton
                            End If
                            intPreMsgButton = intMsgButton
                            
                            blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' '选择了CheckBox的,不再提示
                            
                            Select Case intMsgButton
                            Case vbIgnore, vbYes
                                intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                            Case vbCancel, vbNo
                                intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                            End Select
                            
                        Else
                            If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                            Else
                                intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                            End If
                        End If
                        zlBillingWarnCheck = True
                        Exit Function
                     End If
                     
                    If Not bln允许强制记帐 Then
                        If bytType = 0 Then
                            
                            '门诊记帐
                            If bln报警询问 Then
                                strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & " 低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗?"
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' '选择了CheckBox的,不再提示
                                    Select Case intMsgButton
                                    Case vbIgnore, vbYes
                                        intOutType = EM_Warn_Billing  '报警提示后用户选择继续
                                    Case vbCancel, vbNo
                                        intOutType = EM_Warn_Cancel  '报警提示后用户选择中断
                                    End Select
                                Else
                                    If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                        intOutType = EM_Warn_Billing '强制记帐报警,继续
                                    Else
                                        intOutType = EM_Warn_Cancel '强制记帐报警,取消
                                    End If
                                End If
                                
                            Else
                                strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止门诊记帐！"
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore    '选择了CheckBox的,不再提示
                                Else
                                    Call MsgBox(strMsgInfor, vbInformation + vbOKOnly, gstrSysName)
                                End If
                                intOutType = EM_Warn_Cancel '报警提示必须中断
                            End If
                        Else
                            strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？" & _
                                    IIf(bln允许保存划价单 And bytBillingType = 0, vbCrLf & vbCrLf & "提示:你可以选择继续记帐,并将当前单据保存为划价单,等病人缴费后再审核。", "")
                                    
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' '选择了CheckBox的,不再提示
                                
                                Select Case intMsgButton
                                Case vbIgnore, vbYes
                                    If bln允许保存划价单 And bytBillingType = 0 Then
                                        intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                    Else
                                        intOutType = EM_Warn_SavePrice '报警提示后用户选择继续,但只允许保存为划价单
                                    End If
                                Case vbCancel, vbNo
                                    intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                                End Select
                            Else
                                If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    intOutType = EM_Warn_Qury_Cancel    '报警提示后用户选择中断
                                ElseIf bln允许保存划价单 And bytBillingType = 0 Then
                                    intOutType = EM_Warn_Qury_OK '报警提示后用户选择继续
                                Else
                                    intOutType = EM_Warn_SavePrice '报警提示后用户选择继续,但只允许保存为划价单
                                End If
                            End If
                        End If
                    Else
                        If bln允许保存划价单 And bytBillingType = 0 And bytType <> 0 Then
                        
                            strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & _
                                        Format(dbl余额 - dbl单据金额, "0.00") & " 低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。" & _
                                        vbCrLf & vbCrLf & "提示:你可以选择将当前单据保存为划价单,等病人缴费后再审核。"
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                            Else
                                MsgBox strMsgInfor, vbInformation, gstrSysName
                            End If
                            intOutType = EM_Warn_Billing '强制记帐报警,继续
                        Else
                            If bln报警询问 Then
                                strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & " 低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗?"
                                
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                                    Select Case intMsgButton
                                    Case vbIgnore, vbYes
                                        intOutType = EM_Warn_Billing  '报警提示后用户选择继续
                                    Case vbCancel, vbNo
                                        intOutType = EM_Warn_Cancel  '报警提示后用户选择中断
                                    End Select
                                    
                                Else
                                    If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                        intOutType = EM_Warn_Billing '强制记帐报警,继续
                                    Else
                                        intOutType = EM_Warn_Cancel '强制记帐报警,取消
                                    End If
                                End If
                                
                            Else
                                strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & " 低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。"
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                Else
                                    MsgBox strMsgInfor, vbInformation, gstrSysName
                                End If
                                intOutType = EM_Warn_Billing '强制记帐报警,继续
                            End If
                        End If
                        
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If dbl余额 - dbl单据金额 < 0 Then
                        '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                         If dblItemMoney <> 0 And dbl余额 - (dbl单据金额 - dblItemMoney) > 0 Then
                             '只提示: bytBillingType As Byte '0-记帐,1-划价,2-审核
                             '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                            
                            strMsgInfor = "注意" & vbCrLf & _
                                      "    病人“" & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & "” 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽,是否继续录入该项目？"
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示
                                Select Case intMsgButton
                                Case vbIgnore, vbYes
                                    intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                Case vbCancel, vbNo
                                    intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                                End Select
                            Else
                                If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                    intOutType = EM_Warn_Qury_Cancel    '报警提示后用户选择中断
                                Else
                                    intOutType = EM_Warn_Qury_OK '  报警提示后用户选择继续
                                End If
                            End If
                            zlBillingWarnCheck = True
                            Exit Function
                         End If

                        byt方式 = 2
                        If Not bln允许强制记帐 Then
                            If bln允许保存划价单 Then
                            
                                strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽," & str类别名称 & "禁止记帐。" & _
                                        vbCrLf & vbCrLf & IIf(bytBillingType = 0, "要将当前单据保存为划价单吗？", "要强制保存划价单吗？")
                                        
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示
                                    
                                    Select Case intMsgButton
                                    Case vbIgnore, vbYes
                                        If bytBillingType = 0 Then
                                            intOutType = EM_Warn_SavePrice '报警提示后用户选择继续,但只允许保存存为划价单
                                        Else
                                            intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                        End If
                                    Case vbCancel, vbNo
                                        intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                                    End Select
                                    
                                Else
                                    If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        intOutType = EM_Warn_Qury_Cancel '报警提示后用户选择中断
                                    ElseIf bytBillingType = 0 Then
                                        intOutType = EM_Warn_SavePrice '报警提示后用户选择继续,但只允许保存存为划价单
                                    Else
                                        intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                    End If
                                End If
                            Else
                                strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽," & str类别名称 & "禁止记帐。"
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                Else
                                    MsgBox strMsgInfor, vbInformation, gstrSysName
                                End If
                                intOutType = EM_Warn_Cancel '报警提示必须中断
                            End If
                            
                        Else
                            If bytBillingType = 0 And bln允许保存划价单 Then
                                strMsgInfor = str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽。" & _
                                        vbCrLf & vbCrLf & "提示:你可以选择将当前单据保存为划价单,等病人缴费后再审核。"
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                Else
                                    MsgBox strMsgInfor, vbInformation, gstrSysName
                                End If
                                intOutType = EM_Warn_Billing '强制记帐报警,继续
                            Else
                                
                                If bln报警询问 Then
                                    strMsgInfor = str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽" & ",继续记帐吗?"
                                    If blnMsgShowChkSkip Then
                                        If intPreMsgButton = -1 Then
                                            intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                        Else
                                            intMsgButton = intPreMsgButton
                                        End If
                                        intPreMsgButton = intMsgButton
                                        blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                                        Select Case intMsgButton
                                        Case vbIgnore, vbYes
                                            intOutType = EM_Warn_Billing  '报警提示后用户选择继续
                                        Case vbCancel, vbNo
                                            intOutType = EM_Warn_Cancel  '报警提示后用户选择中断
                                        End Select
                                    Else
                                        If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                            intOutType = EM_Warn_Billing '强制记帐报警,继续
                                        Else
                                            intOutType = EM_Warn_Cancel '强制记帐报警,取消
                                        End If
                                    End If
                                    
                                Else
                                    strMsgInfor = str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽。"
                                    If blnMsgShowChkSkip Then
                                        If intPreMsgButton = -1 Then
                                            intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                        Else
                                            intMsgButton = intPreMsgButton
                                        End If
                                        intPreMsgButton = intMsgButton
                                        blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                    Else
                                        MsgBox strMsgInfor, vbInformation, gstrSysName
                                    End If
                                    intOutType = EM_Warn_Billing '强制记帐报警,继续
                                End If
                            End If
                        End If
                        
                    ElseIf dbl余额 - dbl单据金额 < rsWarn!报警值 Then
                        '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                         If dblItemMoney <> 0 And dbl余额 - (dbl单据金额 - dblItemMoney) > Val(Nvl(rsWarn!报警值)) Then
                             '只提示: bytBillingType As Byte '0-记帐,1-划价,2-审核
                             '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                             strMsgInfor = "注意" & vbCrLf & _
                                           "    病人“" & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & "” 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ", 是否继续录入该项目？"
                                           
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                                Select Case intMsgButton
                                Case vbIgnore, vbYes
                                    intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                Case vbCancel, vbNo
                                    intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                               End Select
                            Else
                                If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                    intOutType = EM_Warn_Qury_Cancel    '报警提示后用户选择中断
                                Else
                                    intOutType = EM_Warn_Qury_OK '报警提示后用户选择继续
                                End If
                            End If
                            zlBillingWarnCheck = True
                            Exit Function
                         End If

                        byt方式 = 1
                        If Not bln允许强制记帐 Then
                            strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？" & _
                                    IIf(bytBillingType = 0 And bln允许保存划价单, vbCrLf & vbCrLf & "提示:你可以选择继续记帐,并将当前单据保存为划价单,等病人缴费后再审核。", "")
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                                Select Case intMsgButton
                                Case vbIgnore, vbYes
                                    intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                Case vbCancel, vbNo
                                    intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                                End Select
                                
                            Else
                                If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    intOutType = EM_Warn_Qury_Cancel '报警提示后用户选择中断
                                Else
                                    intOutType = EM_Warn_Qury_OK    '报警提示后用户选择继续
                                End If
                            End If
                        Else
                            If bytBillingType = 0 And bln允许保存划价单 Then
                            
                                strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & _
                                        Format(dbl余额 - dbl单据金额, "0.00") & " 低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。" & _
                                        vbCrLf & vbCrLf & "提示:你可以选择将当前单据保存为划价单,等病人缴费后再审核。"
                                
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                Else
                                    MsgBox strMsgInfor, vbInformation, gstrSysName
                                End If
                                intOutType = EM_Warn_Billing '强制记帐报警,继续
                            Else
                                If bln报警询问 Then
                                    strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗?"
                                    If blnMsgShowChkSkip Then
                                        If intPreMsgButton = -1 Then
                                            intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                        Else
                                            intMsgButton = intPreMsgButton
                                        End If
                                        
                                        intPreMsgButton = intMsgButton
                                        blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示
                                        Select Case intMsgButton
                                        Case vbIgnore, vbYes
                                            intOutType = EM_Warn_Billing  '报警提示后用户选择继续
                                        Case vbCancel, vbNo
                                            intOutType = EM_Warn_Cancel  '报警提示后用户选择中断
                                        End Select
                                    Else
                                        If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                            intOutType = EM_Warn_Billing '强制记帐报警,继续
                                        Else
                                            intOutType = EM_Warn_Cancel '强制记帐报警,取消
                                        End If
                                    End If
                                Else
                                    strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。"
                                    If blnMsgShowChkSkip Then
                                        If intPreMsgButton = -1 Then
                                            intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                        Else
                                            intMsgButton = intPreMsgButton
                                        End If
                                        intPreMsgButton = intMsgButton
                                        blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                    Else
                                        MsgBox strMsgInfor, vbInformation, gstrSysName
                                    End If
                                    intOutType = EM_Warn_Billing '强制记帐报警,继续
                                End If
                            End If
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If dbl余额 - dbl单据金额 < 0 Then
                            '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                             If dblItemMoney <> 0 And dbl余额 - (dbl单据金额 - dblItemMoney) > 0 Then
                                 '只提示: bytBillingType As Byte '0-记帐,1-划价,2-审核
                                 '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                                
                                strMsgInfor = "注意" & vbCrLf & _
                                               "    病人“" & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & "” 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽,是否继续录入该项目？"
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                                    Select Case intMsgButton
                                    Case vbIgnore, vbYes
                                        intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                    Case vbCancel, vbNo
                                        intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                                    End Select
                                    
                                Else
                                    If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                        intOutType = EM_Warn_Qury_Cancel    '报警提示后用户选择中断
                                    Else
                                        intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                    End If
                                End If
                                zlBillingWarnCheck = True
                                Exit Function
                             End If

                            byt方式 = 2
                            If Not bln允许强制记帐 Then
                                If bln允许保存划价单 Then
                                    strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽," & str类别名称 & "禁止记帐。" & _
                                            vbCrLf & vbCrLf & IIf(bytBillingType = 0, "要将当前单据保存为划价单吗？", "要强制保存划价单吗？")
                                    If blnMsgShowChkSkip Then
                                        If intPreMsgButton = -1 Then
                                            intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                        Else
                                            intMsgButton = intPreMsgButton
                                        End If
                                        intPreMsgButton = intMsgButton
                                        blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                                        Select Case intMsgButton
                                        Case vbIgnore, vbYes
                                            If bytBillingType = 0 Then
                                                intOutType = EM_Warn_SavePrice '报警提示后用户选择继续,但只允许保存存为划价单
                                            Else
                                                intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                            End If
                                        Case vbCancel, vbNo
                                            intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                                        End Select
                                    Else
                                        If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                            intOutType = EM_Warn_Qury_Cancel '报警提示后用户选择中断
                                        ElseIf bytBillingType = 0 Then
                                            intOutType = EM_Warn_SavePrice '报警提示后用户选择继续,但只允许保存存为划价单
                                        Else
                                            intOutType = EM_Warn_Qury_OK '报警提示后用户选择继续
                                        End If
                                    End If
                                    
                                Else
                                    strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽," & str类别名称 & "禁止记帐。"
                                    If blnMsgShowChkSkip Then
                                        If intPreMsgButton = -1 Then
                                            intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                        Else
                                            intMsgButton = intPreMsgButton
                                        End If
                                        intPreMsgButton = intMsgButton
                                        blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                    Else
                                        MsgBox strMsgInfor, vbInformation, gstrSysName
                                    End If
                                    intOutType = EM_Warn_Cancel '报警提示必须中断
                                End If
                            Else
                                If bytBillingType = 0 And bln允许保存划价单 Then
                                    strMsgInfor = str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽。" & _
                                            vbCrLf & vbCrLf & "提示:你可以选择将当前单据保存为划价单,等病人缴费后再审核。"
                                    If blnMsgShowChkSkip Then
                                        If intPreMsgButton = -1 Then
                                            intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                        Else
                                            intMsgButton = intPreMsgButton
                                        End If
                                        intPreMsgButton = intMsgButton
                                        blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                    Else
                                        MsgBox strMsgInfor, vbInformation, gstrSysName
                                    End If
                                    intOutType = EM_Warn_Billing  '强制记帐报警,继续
                                Else
                                    If bln报警询问 Then
                                        strMsgInfor = str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽" & ",继续记帐吗?"
                                        If blnMsgShowChkSkip Then
                                            If intPreMsgButton = -1 Then
                                                intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                            Else
                                                intMsgButton = intPreMsgButton
                                            End If
                                            intPreMsgButton = intMsgButton
                                            blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                                            Select Case intMsgButton
                                            Case vbIgnore, vbYes
                                                intOutType = EM_Warn_Billing  '报警提示后用户选择继续
                                            Case vbCancel, vbNo
                                                intOutType = EM_Warn_Cancel  '报警提示后用户选择中断
                                            End Select
                                        Else
                                            If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                intOutType = EM_Warn_Billing '强制记帐报警,继续
                                            Else
                                                intOutType = EM_Warn_Cancel '强制记帐报警,取消
                                            End If
                                        End If
                                    Else
                                        strMsgInfor = str类别名称 & "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽。"
                                        If blnMsgShowChkSkip Then
                                            If intPreMsgButton = -1 Then
                                                intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                            Else
                                                intMsgButton = intPreMsgButton
                                            End If
                                            intPreMsgButton = intMsgButton
                                            blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                        Else
                                            MsgBox strMsgInfor, vbInformation, gstrSysName
                                        End If
                                        intOutType = EM_Warn_Billing  '强制记帐报警,继续
                                    End If
                                End If
                            End If
                        End If
                    ElseIf byt已报方式 = 2 Then
                        '上次预交款已经耗尽并强制继续,不再处理
                        zlBillingWarnCheck = True
                        Exit Function
                    End If
                End If
            Case 3 '低于报警值禁止记帐
                If dbl余额 - dbl单据金额 < rsWarn!报警值 Then
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If dblItemMoney <> 0 And dbl余额 - (dbl单据金额 - dblItemMoney) > Val(Nvl(rsWarn!报警值)) Then
                         '只提示: bytBillingType As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        strMsgInfor = "注意" & vbCrLf & _
                                       "    病人“" & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & "” 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & ")已经耗尽,是否继续录入该项目？"
                        
                        If blnMsgShowChkSkip Then
                            If intPreMsgButton = -1 Then
                                intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                            Else
                                intMsgButton = intPreMsgButton
                            End If
                            intPreMsgButton = intMsgButton
                            blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                            Select Case intMsgButton
                            Case vbIgnore, vbYes
                                intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                            Case vbCancel, vbNo
                                intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                            End Select
                        Else
                            If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                            Else
                                intOutType = EM_Warn_Qury_OK  ' 报警提示后用户选择继续
                            End If
                        End If
                        zlBillingWarnCheck = True
                        Exit Function
                     End If

                    If Not bln允许强制记帐 Then
                        If bln允许保存划价单 Then
                            strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。" & _
                                    vbCrLf & vbCrLf & IIf(bytBillingType = 0, "要将当前单据保存为划价单吗？", "要强制保存划价单吗？")
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                                Select Case intMsgButton
                                Case vbIgnore, vbYes
                                    If bytBillingType = 0 Then
                                        intOutType = EM_Warn_SavePrice '报警提示后用户选择继续,但只允许保存存为划价单
                                    Else
                                        intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                    End If
                                Case vbCancel, vbNo
                                    intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                                End Select
                            Else
                                If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    intOutType = EM_Warn_Qury_Cancel '报警提示后用户选择中断
                                ElseIf bytBillingType = 0 Then
                                    intOutType = EM_Warn_SavePrice '报警提示后用户选择继续,但只允许保存存为划价单
                                Else
                                    intOutType = EM_Warn_Qury_OK '报警提示后用户选择继续
                                End If
                            End If
                        Else
                            strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。"
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                            Else
                                MsgBox strMsgInfor, vbInformation, gstrSysName
                            End If
                            intOutType = EM_Warn_Cancel '报警提示必须中断
                        End If
                    Else
                        If bytBillingType = 0 And bln允许保存划价单 Then
                            strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。" & _
                                    vbCrLf & vbCrLf & "提示:你可以选择将当前单据保存为划价单,等病人缴费后再审核。"
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                            Else
                                MsgBox strMsgInfor, vbInformation, gstrSysName
                            End If
                            intOutType = EM_Warn_Billing '强制记帐报警,继续
                        Else
                            If bln报警询问 Then
                                strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗?"
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示
                                
                                    Select Case intMsgButton
                                    Case vbIgnore, vbYes
                                        intOutType = EM_Warn_Billing  '报警提示后用户选择继续
                                    Case vbCancel, vbNo
                                        intOutType = EM_Warn_Cancel  '报警提示后用户选择中断
                                    End Select
                                Else
                                    If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                        intOutType = EM_Warn_Billing '强制记帐报警,继续
                                    Else
                                        intOutType = EM_Warn_Cancel '强制记帐报警,取消
                                    End If
                                End If
                            Else
                                strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当前剩余款(含担保额:" & Format(objPati.Obj费用概况.担保额, "0.00") & "):" & Format(dbl余额 - dbl单据金额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。"
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                Else
                                    MsgBox strMsgInfor, vbInformation, gstrSysName
                                End If
                                intOutType = EM_Warn_Billing '强制记帐报警,继续
                            End If
                        End If
                    End If
                End If
        End Select
    Case 2    '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If objPati.当日记帐总额 + dbl单据金额 > rsWarn!报警值 Then
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If dblItemMoney <> 0 And objPati.当日记帐总额 + dbl单据金额 - dblItemMoney < Val(Nvl(rsWarn!报警值)) Then
                         '只提示: bytBillingType As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        strMsgInfor = "注意" & vbCrLf & _
                                       "    病人“" & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & "” 当日费用:" & Format(objPati.当日记帐总额 + dbl单据金额, strDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",是否继续录入该项目？"
                        
                        If blnMsgShowChkSkip Then
                            If intPreMsgButton = -1 Then
                                intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                            Else
                                intMsgButton = intPreMsgButton
                            End If
                            intPreMsgButton = intMsgButton
                            blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                            Select Case intMsgButton
                            Case vbIgnore, vbYes
                                intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                            Case vbCancel, vbNo
                                intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                            End Select
                        Else
                            If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                intOutType = EM_Warn_Qury_Cancel '报警提示后用户选择中断
                            Else
                                intOutType = EM_Warn_Qury_OK '报警提示后用户选择继续
                            End If
                        End If
                        zlBillingWarnCheck = True
                        Exit Function
                     End If
                    If Not bln允许强制记帐 Then
                        If bytType = 0 Then
                            If bln报警询问 Then
                                strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当日费用:" & Format(objPati.当日记帐总额 + dbl单据金额, strDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？"
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                                    Select Case intMsgButton
                                    Case vbIgnore, vbYes
                                        intOutType = EM_Warn_Billing  '报警提示后用户选择继续
                                    Case vbCancel, vbNo
                                        intOutType = EM_Warn_Cancel  '报警提示后用户选择中断
                                    End Select
                                    
                                Else
                                    If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                        intOutType = EM_Warn_Billing '强制记帐报警,继续
                                    Else
                                        intOutType = EM_Warn_Cancel '强制记帐报警,继续
                                    End If
                                End If
                            Else
                                strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当日费用:" & Format(objPati.当日记帐总额 + dbl单据金额, strDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止门诊记帐！"
                                If blnMsgShowChkSkip Then
                                    If intPreMsgButton = -1 Then
                                        intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                    Else
                                        intMsgButton = intPreMsgButton
                                    End If
                                    intPreMsgButton = intMsgButton
                                    blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                                Else
                                    Call MsgBox(strMsgInfor, vbOKOnly + vbInformation, gstrSysName)
                                End If
                                intOutType = EM_Warn_Cancel '报警提示必须中断
                            End If
                        Else
                            strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当日费用:" & Format(objPati.当日记帐总额 + dbl单据金额, strDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？"
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示

                                Select Case intMsgButton
                                Case vbIgnore, vbYes
                                    intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                Case vbCancel, vbNo
                                    intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                                End Select
                            Else
                                If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    intOutType = EM_Warn_Qury_Cancel    '报警提示后用户选择中断
                                Else
                                    intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                                End If
                            End If
                        End If
                    Else
                        If bln报警询问 Then
                            strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当日费用:" & Format(objPati.当日记帐总额 + dbl单据金额, strDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗?"
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示
                                
                                Select Case intMsgButton
                                Case vbIgnore, vbYes
                                    intOutType = EM_Warn_Billing  '报警提示后用户选择继续
                                Case vbCancel, vbNo
                                    intOutType = EM_Warn_Cancel  '报警提示后用户选择中断
                                End Select
                            Else
                                If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                    intOutType = EM_Warn_Billing '强制记帐报警,继续
                                Else
                                    intOutType = EM_Warn_Cancel '强制记帐报警,取消
                                End If
                            End If
                        Else
                            strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当日费用:" & Format(objPati.当日记帐总额 + dbl单据金额, strDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。"
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                            Else
                                MsgBox strMsgInfor, vbInformation, gstrSysName
                            End If
                            intOutType = EM_Warn_Billing  '强制记帐报警,继续
                        End If
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If objPati.当日记帐总额 + dbl单据金额 > rsWarn!报警值 Then
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If dblItemMoney <> 0 And objPati.当日记帐总额 + dbl单据金额 - dblItemMoney < Val(Nvl(rsWarn!报警值)) Then
                         '只提示: bytBillingType As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        
                        strMsgInfor = "注意" & vbCrLf & _
                                       "    病人“" & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & "” 当日费用:" & Format(objPati.当日记帐总额 + dbl单据金额, strDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",是否继续录入该项目？"
                        If blnMsgShowChkSkip Then
                            If intPreMsgButton = -1 Then
                                intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                            Else
                                intMsgButton = intPreMsgButton
                            End If
                            intPreMsgButton = intMsgButton
                            blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示
                            Select Case intMsgButton
                            Case vbIgnore, vbYes
                                intOutType = EM_Warn_Qury_OK  '报警提示后用户选择继续
                            Case vbCancel, vbNo
                                intOutType = EM_Warn_Qury_Cancel  '报警提示后用户选择中断
                            End Select
                            
                        Else
                            If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                intOutType = EM_Warn_Qury_Cancel '报警提示后用户选择中断
                            Else
                                intOutType = EM_Warn_Qury_OK ' 报警提示后用户选择继续
                            End If
                        End If
                        zlBillingWarnCheck = True
                        Exit Function
                     End If

                    If Not bln允许强制记帐 Then
                        strMsgInfor = objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当日费用:" & Format(objPati.当日记帐总额 + dbl单据金额, strDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。"
                        If blnMsgShowChkSkip Then
                            If intPreMsgButton = -1 Then
                                intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                            Else
                                intMsgButton = intPreMsgButton
                            End If
                            intPreMsgButton = intMsgButton
                            blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                        Else
                            MsgBox strMsgInfor, vbInformation, gstrSysName
                        End If
                        
                        intOutType = EM_Warn_Cancel '报警提示必须中断
                    Else
                        If bln报警询问 Then
                            strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当日费用:" & Format(objPati.当日记帐总额 + dbl单据金额, strDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗?"
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore Or intMsgButton = vbCancel   ' 选择了CheckBox的,不再提示
                                
                                Select Case intMsgButton
                                Case vbIgnore, vbYes
                                    intOutType = EM_Warn_Billing  '报警提示后用户选择继续
                                Case vbCancel, vbNo
                                    intOutType = EM_Warn_Cancel  '报警提示后用户选择中断
                                End Select
                            Else
                                If MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                    intOutType = EM_Warn_Billing '强制记帐报警,继续
                                Else
                                    intOutType = EM_Warn_Cancel '强制记帐报警,取消
                                End If
                            End If
                        Else
                            strMsgInfor = "强制记帐提醒:" & vbCrLf & vbCrLf & objPati.姓名 & IIf(objPati.住院号 = "", "", "(住院号:" & objPati.住院号 & " 床号:" & objPati.床号 & ")") & " 当日费用:" & Format(objPati.当日记帐总额 + dbl单据金额, strDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。"
                            
                            If blnMsgShowChkSkip Then
                                If intPreMsgButton = -1 Then
                                    intMsgButton = frmMsgBox.ShowMsgBox(strMsgInfor, frmMain, True)
                                Else
                                    intMsgButton = intPreMsgButton
                                End If
                                intPreMsgButton = intMsgButton
                                blnOutSkip = intMsgButton = vbIgnore    ' 选择了CheckBox的,不再提示,不包含vbCancel,主要是临床批量记帐时，点击窗体中的X时，表示未选择Cehckbox控件（不再提示)
                            Else
                                MsgBox strMsgInfor, vbInformation, gstrSysName
                            End If
                            intOutType = EM_Warn_Billing    '强制记帐报警,继续
                        End If
                    End If
                End If
        End Select
    End Select

    '对于继续类的操作,返回已报警类别
    If intOutType = EM_Warn_Qury_OK Or intOutType = EM_Warn_Billing Or intOutType = EM_Warn_SavePrice Then
        '报警提示后用户选择继续;强制记帐报警,继续;报警提示后用户选择继续,但只允许保存存为划价单
        
        If byt标志 = 1 Then
            If rsWarn!报警标志1 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & objPati.病人ID & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & objPati.病人ID & ":", ",") & rsWarn!报警标志1
            End If
        ElseIf byt标志 = 2 Then
            If rsWarn!报警标志2 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & objPati.病人ID & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & objPati.病人ID & ":", ",") & rsWarn!报警标志2
            End If
            '附加标注以判断已报警的具体方式
            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
        ElseIf byt标志 = 3 Then
            If rsWarn!报警标志3 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & objPati.病人ID & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & objPati.病人ID & ":", ",") & rsWarn!报警标志3
            End If
        End If
    End If
    zlBillingWarnCheck = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Private Sub Class_Initialize()
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mrsWarn Is Nothing Then Set mrsWarn = Nothing
    If Not mrsFeeType Is Nothing Then Set mrsFeeType = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
