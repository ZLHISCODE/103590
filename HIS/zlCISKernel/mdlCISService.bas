Attribute VB_Name = "mdlCISService"
Option Explicit
        
Public Function GetColVal(ByVal colData As Collection, ByVal strKey As String, Optional ByVal strType As String, Optional ByVal strDef As String, Optional ByRef lngExist As Long) As String
'功能:通集合关键字获取集合的值,基本数据类型,数字或字符
'入参：strType  N/n  表示数字类型，c表示字符串
'      strDef  缺省值，当出错时以这个值为缺省值返回
'出参:lngExist 集合中是否存在这个结点值,0-存在,-1不存在
    Dim strValue As String
    
    On Error GoTo errH
    
    If IsNull(colData(strKey)) Then
        strValue = ""
    Else
        strValue = colData(strKey)
    End If
     
    If UCase(strType) = "N" Then
        strValue = Val(strValue)
    End If
    
    GetColVal = strValue
    
    Exit Function
errH:
    err.Clear
    lngExist = -1
    '集合访问不到不提示继续处理
    If strDef <> "" Then
        strValue = strDef
    Else
        If UCase(strType) = "N" Then
            strValue = 0
        Else
            strValue = ""
        End If
    End If
    GetColVal = strValue
End Function

Public Function GetColObj(ByVal colData As Collection, ByVal strKey As String) As Collection
'功能:通集合关键字获取集合中的集合对象
    On Error GoTo errH
    Set GetColObj = colData(strKey)
    Exit Function
errH:
    err.Clear
    Set GetColObj = New Collection
End Function

Public Function MergeStr(ByVal strIn1 As String, ByVal strIn2 As String, Optional ByVal strTag As String = ",") As String
'功能:将两个字符串按指定连接符合拼接到一起
'参入:strTag 连接符号
    Dim i As Long
    Dim varTmp As Variant
    Dim strTmp As String
    
    On Error GoTo errH
    
    If strIn1 <> "" Or strIn2 <> "" Then
        If strIn1 = "" Then
            MergeStr = strIn2
        ElseIf strIn2 = "" Then
            MergeStr = strIn1
        Else
            '判断是否互相包含
            strTmp = strIn1 & strTag & strIn2
            varTmp = Split(strTmp, strTag)
            strTmp = ""
            For i = 0 To UBound(varTmp)
                If varTmp(i) & "" <> "" Then
                    If InStr(strTag & strTmp & strTag, strTag & varTmp(i) & strTag) = 0 Then
                        strTmp = strTmp & strTag & varTmp(i)
                    End If
                End If
            Next
            strTmp = Mid(strTmp, Len(strTag) + 1)
            MergeStr = strTmp
        
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetJsonNum(ByVal strNum As String) As String
'功能:json拼串时带小数点的要特殊处理
    Dim dblTmp As Double
    Dim strTmp As String
    dblTmp = Val(strNum)
    strTmp = dblTmp
    If Mid(strTmp, 1, 1) = "." Then strTmp = "0" & strTmp
    GetJsonNum = strTmp
End Function

Public Sub InitSQLSend(rsSQL As ADODB.Recordset)
'功能:医嘱发送窗体初始化记录集
    'SQL记录集
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "类型", adInteger '1-费用记录,2-医嘱记录,3-发送记录,4-发料记录
                                          '1-计价,2-发送,3-签名,4-费用,5-发料 门诊医嘱送
                                          '1-计价,2-签名,3-校对,4-发送,5-费用,6-发料,,8-婴儿转科
                                          '门诊医嘱发送sql,类型=2
                                          
    rsSQL.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsSQL.Fields.Append "项目ID", adBigInt '收费细目ID
    rsSQL.Fields.Append "序号", adBigInt '用于排序
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.Fields.Append "NO", adVarChar, 30, adFldIsNullable '用于NO替换处理时排序
    rsSQL.Fields.Append "费用ID", adVarChar, 300 '费用ID 药品卫材,调用服务的时间要用需提前生成,要用字符串方便在sql中进行替换
    rsSQL.Fields.Append "科室ID", adBigInt '药房ID 药品卫材,调用服务的时间要用需提前生成,用于生成发药窗口
    rsSQL.Fields.Append "收费类别", adVarChar, 30 '用于区分药品卫材,调用服务的时间要用,卫材='4' 药品='5'
    rsSQL.Fields.Append "主医嘱IDs", adVarChar, 50000 '用于输液配药记录的生成时使用
    rsSQL.Fields.Append "费用来源", adBigInt '费用来源,1-门诊费用记录,2-住院费用记录
    rsSQL.Fields.Append "单据类型", adBigInt '单据类型,1-收费,2-记帐
    rsSQL.Fields.Append "SQLEX", adVarChar, 5000 '费用数据生成需要信息拼串
    rsSQL.Fields.Append "药品卫材", adVarChar, 5000 '药品卫材数据生需要信息拼串,要配合前面的费用数据一起完成
    rsSQL.Fields.Append "同步标记", adVarChar, 5000  '1-药品,2-卫材,格式:"医嘱ID:发送号:1,医嘱ID:发送号:2",一条医嘱只可能有3种情况,药品/卫材/药品+卫材
    rsSQL.Fields.Append "数量", adBigInt '数量
    rsSQL.Fields.Append "单价", adDouble '标准单价
    rsSQL.Fields.Append "批次", adBigInt '卫材批次
    rsSQL.Fields.Append "开单科室ID", adBigInt '开单科室ID
    rsSQL.Fields.Append "执行完成", adBigInt '自动执行完成,0-不执行完成,1-要执行完成
    
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
End Sub

Public Function GetTakeNos(ByVal lng部门ID As Long, ByVal strTitle As String) As String
'功能:获取领药号
    Dim strJsonIn As String
    On Error GoTo errH
    strJsonIn = "{""input"":{""dept_id"":" & lng部门ID & "}}"
    Call CallService("Zl_Drugsvr_GetTakeNo", strJsonIn, , strTitle, , False, , , , True)
    GetTakeNos = gobjService.GetJsonNodeValue("output.data")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetStockCheck(ByVal bytType As Byte) As Collection
'功能：获取药品或卫材出库检查的集合
'参数：bytType:0-药品，1-卫材

    Dim colStock As Collection, i As Long
    Dim strJsonOut As String
    Dim colList As Collection
    Dim strName As String
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '避免出错
    If bytType = 0 Then
        strName = "zl_DrugSvr_Getstockcheck"
    Else
        strName = "Zl_Stuffsvr_Getstockcheck"
    End If
    
    If CallService(strName, "", strJsonOut, "mdlCisService.GetStockCheck") Then
       Set colList = gobjService.GetJsonListValue("output.item_list")
    End If
    If Not colList Is Nothing Then
        If colList.Count > 0 Then
            For i = 1 To colList.Count
                If bytType = 0 Then
                    colStock.Add Val("" & colList(i)("_check_type")), "_" & colList(i)("_pharmacy_id")
                Else
                    colStock.Add Val("" & colList(i)("_check_type")), "_" & colList(i)("_warehouse_id")
                End If
            Next
        End If
    End If
    Set GetStockCheck = colStock
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function GetPati剩余款(ByVal strTitle As String, ByVal lngMod As Long, ByRef colIn As Collection) As Double
'功能：获取住院病人剩余款
'入参：colIn 可做入参和出参 4个元素，病人id,主页id,病人性质，险类,查询方式
'       query_type 查询方式
'               0-门诊病人信息
'               1-门诊医嘱下达,记帐报警用,门诊医嘱清单显示医嘱金额
'               2-住院医嘱下达,记帐报警用
'               3-住院编辑状态栏显示
'               4-获取 病人余额.费用余额，记帐权限控制检查提示用
'               5-门诊医嘱发送状态标显示
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim strTmp As String
    
    On Error GoTo errH
    
    strJsonIn = "{""input"":{" & _
        """pati_id"":" & colIn("_pati_id") & "," & _
        """pati_pageid"":" & GetColVal(colIn, "_pati_pageid", "N", 0) & "," & _
        """pati_character"":" & GetColVal(colIn, "_pati_character", "N", 0) & "," & _
        """insurance_type"":" & GetColVal(colIn, "_insurance_type", "N", 0) & "," & _
        """query_type"":" & GetColVal(colIn, "_query_type", "N", 0) & _
        "}}"
        
    If CallService("Zl_Exsesvr_Getexcessfunds", strJsonIn, strJsonOut, strTitle, lngMod, True) Then
        Set colIn = New Collection: strTmp = "" & gobjService.GetJsonNodeValue("output.excess_funds")
        colIn.Add strTmp, "余额": strTmp = "" & gobjService.GetJsonNodeValue("output.prepay_funds")
        colIn.Add strTmp, "预交余额": strTmp = "" & gobjService.GetJsonNodeValue("output.prepaid_expenses")
        colIn.Add strTmp, "预结费用": strTmp = "" & gobjService.GetJsonNodeValue("output.guarantee_amount")
        colIn.Add strTmp, "担保额": strTmp = ""
        GetPati剩余款 = Val("" & gobjService.GetJsonNodeValue("output.excess_funds"))
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPati家属(ByVal strTitle As String, ByVal lngMod As Long, ByVal lng病人ID As Long) As ADODB.Recordset
'功能：获取住院病人剩余款
'入参：colIn 4个元素，病人id,主页id,病人性质，险类
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim rsTmp As New ADODB.Recordset
    Dim colList As New Collection
    Dim i As Long
    
    On Error GoTo errH
   
    
    rsTmp.Fields.Append "关系", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "姓名", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "性别", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "年龄", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "就诊卡号", adVarChar, 500, adFldIsNullable
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    
    strJsonIn = "{""input"":{""pati_id"":" & lng病人ID & "}}"
    If CallService("zl_patisvr_getfamilymembers", strJsonIn, strJsonOut, strTitle, lngMod, True) Then
        Set colList = gobjService.GetJsonListValue("output.item_list")
        If colList.Count > 0 Then
            For i = 1 To colList.Count
                rsTmp.AddNew
                rsTmp!关系 = colList(i)("_relationship")
                rsTmp!姓名 = colList(i)("_pati_name")
                rsTmp!性别 = colList(i)("_pati_sex")
                rsTmp!年龄 = colList(i)("_pati_age")
                rsTmp!就诊卡号 = colList(i)("_visit_card_no")
                rsTmp.Update
            Next
            rsTmp.MoveFirst
        End If
    End If
    Set GetPati家属 = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get适用病人(ByVal strTitle As String, ByVal lngMod As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：获取 适用病人
'入参：colIn 可做入参和出参 4个元素，病人id,主页id,病人性质，险类
'
    Dim strJsonIn As String
    Dim strJsonOut As String
   
    
    On Error GoTo errH
    
    strJsonIn = "{""input"":{" & _
        """pati_id"":" & lng病人ID & "," & _
        """pati_pageid"":" & lng主页ID & _
        "}}"
        
    If CallService("Zl_Patisvr_Patiwarnscheme", strJsonIn, strJsonOut, strTitle, lngMod, True) Then
        Get适用病人 = "" & gobjService.GetJsonNodeValue("output.pati_scheme")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get报警线(ByVal strTitle As String, ByVal lngMod As Long, ByRef colIn As Collection) As ADODB.Recordset
'功能:查询记帐报警线
    '入参：colIn 可做入参和出参 4个元素，病人id,主页id,病人性质，险类
'  --     query_type   N 1 查询方式
'  --                     0-仅根据 病区id / 适用病人 查找，返回一个值,用于记帐报警提示
'  --                     1-按病区id 查找，返回列表,过滤病人列表排开欠费病人
    Dim strJsonIn As String
    Dim strJsonOut As String
 
    Dim rsTmp As New ADODB.Recordset
    Dim colList As New Collection
    Dim i As Long
    
    On Error GoTo errH
     
    rsTmp.Fields.Append "适用病人", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "报警方法", adInteger, , adFldIsNullable
    rsTmp.Fields.Append "报警值", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "报警标志1", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "报警标志2", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "报警标志3", adVarChar, 500, adFldIsNullable
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    
    strJsonIn = "{""input"":{" & _
        """pati_scheme"":""" & GetColVal(colIn, "_pati_scheme") & """," & _
        """wardarea_id"":" & GetColVal(colIn, "_wardarea_id", "N", 0) & "," & _
        """query_type"":" & GetColVal(colIn, "_query_type", "N", 0) & _
        "}}"
        

    If CallService("Zl_Exsesvr_Getwarnline", strJsonIn, strJsonOut, strTitle, lngMod, True) Then
    
        If Val("" & gobjService.GetJsonNodeValue("output.alarm_value")) > 0 Then
            rsTmp.AddNew
            rsTmp!报警值 = Val("" & gobjService.GetJsonNodeValue("output.alarm_value"))
            rsTmp.Update
        End If
        
        
        Set colList = gobjService.GetJsonListValue("output.item_list")
        If colList.Count > 0 Then
            For i = 1 To colList.Count
                rsTmp.AddNew
                rsTmp!适用病人 = "" & colList(i)("_pati_scheme")
                rsTmp!报警方法 = Val("" & colList(i)("_alarm_way"))
                rsTmp!报警值 = Val("" & colList(i)("_alarm_value"))
                rsTmp!报警标志1 = Val("" & colList(i)("_alarm_one"))
                rsTmp!报警标志2 = Val("" & colList(i)("_alarm_two"))
                rsTmp!报警标志3 = Val("" & colList(i)("_alarm_three"))
                rsTmp.Update
            Next
            rsTmp.MoveFirst
        End If
    End If
    Set Get报警线 = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetFeeJsonIn(ByVal colData As Collection) As String
'功能:返回费用生成入参,一个病人一调用一次
'Zl_Exsesvr_Newbill
    Dim i As Long
    Dim colItem As Collection
    Dim strPati As String
    Dim strItems As String
    Dim strOut As String
    
    On Error GoTo errH
    
    If colData.Count = 0 Then
        Exit Function
    End If
    
    For i = 1 To colData.Count
        Set colItem = colData(i)
        
        If strPati = "" Then
            strPati = GetJsonStrNode("billtype", "单据类型", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_source", "费用来源", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_id", "病人ID", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_pageid", "主页ID", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("baby_num", "婴儿费", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("sgin_no", "标识号", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("bed_num", "床号", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_name", "姓名", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_sex", "性别", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_age", "年龄", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("fee_category", "费别", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("overtime_sign", "加班标志", "N", colItem, 1)
            strPati = strPati & "," & GetJsonStrNode("pati_deptid", "病人科室ID", "N", colItem, 1)
            strPati = strPati & "," & GetJsonStrNode("pati_wardarea_id", "病人病区ID", "N", colItem, 1)
            strPati = strPati & "," & GetJsonStrNode("operator_name", "操作员姓名", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("operator_code", "操作员编号", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("outpati_tag", "门诊标志", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("rgst_id", "挂号ID", "N", colItem, 1)
            strPati = strPati & "," & GetJsonStrNode("emg_sign", "是否急诊", "N", colItem, 1)
        End If
        
        strItems = strItems & ",{" & GetJsonStrNode("fee_id", "费用ID", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("fee_no", "NO", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("serial_num", "序号", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("charge_tag", "划价", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("placer", "开单人", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("plcdept_id", "开单部门ID", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("sub_serial_num", "从属父号", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("fitem_id", "收费细目ID", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("item_type", "收费类别", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("unit", "计算单位", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("pharmacy_window", "发药窗口", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("packages_num", "付数", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("send_num", "数次", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("ext_mark", "附加标志", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("exe_deptid", "执行部门ID", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("price_ftrnum", "价格父号", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("income_item_id", "收入项目ID", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("receipt_name", "收据费目", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("price", "标准单价", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("fee_amrcvb", "应收金额", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("fee_ampaib", "实收金额", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("happen_time", "发生时间", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("create_time", "登记时间", "C", colItem)
        'strItems = strItems & "," & GetJsonStrNode("memo", "费用摘要", "C", colItem)
        strItems = strItems & ",""memo"":""" & zlStr.ToJsonStr(GetColVal(colItem, "费用摘要")) & """"
        strItems = strItems & "," & GetJsonStrNode("order_id", "医嘱序号", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("baby_num", "婴儿费", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("exe_properties", "执行性质", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("decoction_method", "煎法", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("morphology", "中药形态", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("bakstuff_batch", "批次", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("insurance", "保险项目否", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("insure_id", "保险大类ID", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("insure_code", "保险编码", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("fee_type", "费用类型", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("si_manp_money", "统筹金额", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("synchro", "更新同步标志", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("effective_time", "期效", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("receipt_issecret", "保密", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("takedept_id", "领药部门ID", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("group_id", "医疗小组ID", "N", colItem, 1)
        strItems = strItems & "}"
    Next
    
    strOut = "{""input"":{" & strPati & ",""item_list"":[" & Mid(strItems, 2) & "]}}"
    GetFeeJsonIn = strOut
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetFeeNodes(ByVal lng来源 As Long, ByVal lng类型 As Long, ByVal strPar As String, ByRef colOut As Collection)
'功能:解析费用结点
'参数:lng来源 费用来源,1-门诊费用记录,2-住院费用记录
'     lng类型 单据类型,1-收费,2-记帐
    Dim lngIdx As Long
    Dim varTmp As Variant
   
    
    On Error GoTo errH
    Set colOut = New Collection
    
    colOut.Add lng来源 & "", "费用来源"
    colOut.Add lng类型 & "", "单据类型"
    
    'ZL_门诊记帐记录_INSERT
    If lng来源 = 1 And lng类型 = 2 Then
        lngIdx = 0
        varTmp = Split(strPar, "<SPCHAR>")
        colOut.Add varTmp(lngIdx) & "", "FNO": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "序号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "病人ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "标识号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "姓名": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "性别": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "年龄": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "费别": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "加班标志": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "婴儿费": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "病人科室ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "开单部门ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "开单人": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "从属父号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收费细目ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收费类别": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "计算单位": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "付数": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "数次": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "附加标志": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "执行部门ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "价格父号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收入项目ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收据费目": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "标准单价": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "应收金额": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "实收金额": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "发生时间": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "登记时间": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "划价": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "F发药窗口": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "操作员编号": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "操作员姓名": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "F费用ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "记帐单ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "费用摘要": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "医嘱序号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "门诊标志": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "中药形态": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "煎法": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "主页ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "病人病区ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "批次": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "更新同步标志": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "期效": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "保密": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "挂号ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "是否急诊": lngIdx = lngIdx + 1 '--NUMBER
        
        If lngIdx <= UBound(varTmp) Then
            colOut.Add varTmp(lngIdx) & "", "NO": lngIdx = lngIdx + 1 '--VARCHAR2
            colOut.Add varTmp(lngIdx) & "", "费用ID": lngIdx = lngIdx + 1 '--NUMBER
            colOut.Add varTmp(lngIdx) & "", "发药窗口": lngIdx = lngIdx + 1 '--VARCHAR2
        End If
    End If
    
    'ZL_门诊划价记录_INSERT
    If lng来源 = 1 And lng类型 = 1 Then
        lngIdx = 0
        varTmp = Split(strPar, "<SPCHAR>")
        colOut.Add varTmp(lngIdx) & "", "FNO": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "序号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "病人ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "主页ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "标识号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "付款方式": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "姓名": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "性别": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "年龄": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "费别": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "加班标志": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "病人科室ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "开单部门ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "开单人": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "从属父号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收费细目ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收费类别": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "计算单位": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "F发药窗口": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "付数": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "数次": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "附加标志": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "执行部门ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "价格父号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收入项目ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收据费目": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "标准单价": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "应收金额": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "实收金额": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "发生时间": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "登记时间": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "操作员姓名": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "F费用ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "费用摘要": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "医嘱序号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "煎法": lngIdx = lngIdx + 1 '--VARCHAR2
        'colOut.Add varTmp(lngIdx) & "", "病人来源": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "保险编码": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "费用类型": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "保险项目否": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "保险大类ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "中药形态": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "执行人": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "病人病区ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "批次": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "更新同步标志": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "期效": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "保密": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "挂号ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "是否急诊": lngIdx = lngIdx + 1 '--NUMBER
        
        If lngIdx <= UBound(varTmp) Then
            colOut.Add varTmp(lngIdx) & "", "NO": lngIdx = lngIdx + 1 '--VARCHAR2
            colOut.Add varTmp(lngIdx) & "", "费用ID": lngIdx = lngIdx + 1 '--NUMBER
            colOut.Add varTmp(lngIdx) & "", "发药窗口": lngIdx = lngIdx + 1 '--VARCHAR2
        End If
    End If
    
    'ZL_住院记帐记录_Insert
    If lng来源 = 2 And lng类型 = 2 Then
        lngIdx = 0
        varTmp = Split(strPar, "<SPCHAR>")
        colOut.Add varTmp(lngIdx) & "", "FNO": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "序号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "病人ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "主页ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "标识号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "姓名": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "性别": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "年龄": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "床号": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "费别": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "病人病区ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "病人科室ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "加班标志": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "婴儿费": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "开单部门ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "开单人": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "从属父号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收费细目ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收费类别": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "计算单位": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "保险项目否": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "保险大类ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "保险编码": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "付数": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "数次": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "附加标志": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "执行部门ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "价格父号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收入项目ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "收据费目": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "标准单价": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "应收金额": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "实收金额": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "统筹金额": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "发生时间": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "登记时间": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "划价": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "操作员编号": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "操作员姓名": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "F费用ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "多病人单": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "记帐单ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "费用摘要": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "是否急诊": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "医嘱序号": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "简单记帐": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "费用类型": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "医技补临床费用": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "中药形态": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "医疗小组ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "煎法": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "执行性质": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "批次": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "领药部门ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "更新同步标志": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "期效": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "保密": lngIdx = lngIdx + 1 '--NUMBER
        
        If lngIdx <= UBound(varTmp) Then
            colOut.Add varTmp(lngIdx) & "", "NO": lngIdx = lngIdx + 1 '--VARCHAR2
            colOut.Add varTmp(lngIdx) & "", "费用ID": lngIdx = lngIdx + 1 '--NUMBER
            colOut.Add varTmp(lngIdx) & "", "发药窗口": lngIdx = lngIdx + 1 '--VARCHAR2
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetJsonStrNode(ByVal strNode As String, ByVal strKey As String, ByVal strType As String, ByVal colNode As Collection, Optional ByVal lngForce As Long) As String
'功能:生成JSON串的结点
'参数:strNode 结点名,
'     strKey  集合colNode中的关键字
'     strType 结点类型,"N"数字,"C" 字符
'     colNode 目标集合
'     lngForce 1-对于是数字类型的,如果是0转换为NULL
    Dim strOut As String
    Dim strVal As String
    Dim lngExist As Long
    
    strOut = """" & strNode & """:"
    
    strVal = GetColVal(colNode, strKey, strType, lngExist)
    
    If lngExist = -1 Or UCase(strVal) = "NULL" Then
        If strType = "N" Then
            strVal = "null"
        Else
            strVal = """"""
        End If
    ElseIf strType = "C" Then
        strVal = """" & strVal & """"
    ElseIf strType = "N" Then
        If strVal = "" Then
            strVal = "null"
        ElseIf lngForce = 1 And strVal = "0" Then
            strVal = "null"
        Else
            strVal = GetJsonNum(strVal)
        End If
    End If
    
    strOut = strOut & strVal
    
    GetJsonStrNode = strOut
End Function

Public Function ExsesvrGetNextID(strTable As String, Optional strFild As String, Optional ByVal strTitle As String, Optional ByVal lng模块 As Long, Optional ByVal lng数量 As Long, Optional ByRef varIDs As Variant) As Double
'功能:获取费用序列得到ID值
    Dim strJsonIn As String
    Dim strTmp As String
    
    If lng数量 <= 1 Then
        lng数量 = 0
    End If
    
    strJsonIn = "{""input"":{""table_name"":""" & strTable & """,""col_name"":""" & strFild & """,""quantity"":" & lng数量 & "}}"
    Call CallService("Zl_Exsesvr_Getnextid", strJsonIn, , strTitle, lng模块, True)
    
    If lng数量 > 1 Then
        strTmp = gobjService.GetJsonNodeValue("output.next_id")
    Else
        ExsesvrGetNextID = gobjService.GetJsonNodeValue("output.next_id")
        strTmp = ExsesvrGetNextID
    End If
    varIDs = Split(strTmp, ",")
        
End Function

Public Function GetFinishFeeJsonIn(ByVal colData As Collection) As String
'功能:自动发料的卫材费用更新为执行完成状态
 
    Dim i As Long
    Dim colItem As Collection
    Dim strItems As String
    Dim strOut As String
 
    Dim lngFee_origin As Long
    
    On Error GoTo errH
    
    If colData.Count = 0 Then
        Exit Function
    End If
    
    For i = 1 To colData.Count
        Set colItem = colData(i)
        If GetColVal(colItem, "自动发料") = "1" Then
            If Val(GetColVal(colItem, "费用ID")) > 0 Then
                strItems = strItems & "," & GetColVal(colItem, "费用ID")
            End If
        End If
    Next
    If strItems <> "" Then
        If GetColVal(colItem, "单据类型") = "2" And GetColVal(colItem, "病人来源") = "2" Then
            lngFee_origin = 2
        Else
            lngFee_origin = 1
        End If
    
    
        strOut = "{""input"":{""fee_ids"":""" & Mid(strItems, 2) & """,""oper_type"":1"
        strOut = strOut & "," & GetJsonStrNode("exe_people", "操作员姓名", "C", colItem)
        strOut = strOut & "," & GetJsonStrNode("exe_time", "登记时间", "C", colItem)
        strOut = strOut & ",""fee_origin"":" & lngFee_origin
        strOut = strOut & ",""exe_status"":1"
        strOut = strOut & "}}"
    End If
    GetFinishFeeJsonIn = strOut
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get皮试结果信息(ByVal lng病人ID As String, ByVal lng主页ID As Long, ByVal str挂号单 As String, Optional ByRef blnHaveData As Boolean) As ADODB.Recordset
'功能：获取病人皮试结果信息
    Dim strSQL As String
    Dim rs皮试 As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "select a.皮试结果, a.药名id, a.药品id from" & vbNewLine & _
        "(Select a.皮试结果, c.项目id As 药名id,0 as 药品id,a.开始执行时间" & vbNewLine & _
        "From 病人医嘱记录 A, 诊疗用法用量 C" & vbNewLine & _
        "Where a.诊疗项目id = c.用法id And Nvl(c.性质, 0) = 0 And Nvl(a.医嘱状态, 0) = 8 and a.皮试结果 is not null" & vbNewLine & _
        IIF(str挂号单 = "", " and a.病人id=[1] and a.主页id=[2]", " and a.挂号单=[3]") & vbNewLine & _
        "union all" & vbNewLine & _
        "Select a.皮试结果, b.药名id, b.药品id,a.开始执行时间" & vbNewLine & _
        "From 病人医嘱记录 A, 药品规格 B, 药品用法用量 C" & vbNewLine & _
        "Where a.诊疗项目id = c.用法id And b.药品id = c.药品id And Nvl(c.性质, 0) = 0 And Nvl(a.医嘱状态, 0) = 8 And  a.皮试结果<>'免试'" & vbNewLine & _
        IIF(str挂号单 = "", " and a.病人id=[1] and a.主页id=[2]", " and a.挂号单=[3]") & vbNewLine & _
        "union all" & vbNewLine & _
        "Select a.皮试结果, c.项目id As 药名id,0 as 药品id,a.开始执行时间" & vbNewLine & _
        "From 病人医嘱记录 A, 诊疗用法用量 C" & vbNewLine & _
        "Where a.诊疗项目id = c.用法id And Nvl(c.性质, 0) = 0 And a.皮试结果='免试'" & vbNewLine & _
        IIF(str挂号单 = "", " and a.病人id=[1] and a.主页id=[2]", " and a.挂号单=[3]") & vbNewLine & _
        " ) a" & vbNewLine & _
        "group by a.皮试结果, a.药名id, a.药品id,a.开始执行时间" & vbNewLine & _
        "order by a.开始执行时间 desc"
    Set rs皮试 = zlDatabase.OpenSQLRecord(strSQL, "mdlCISService", lng病人ID, lng主页ID, str挂号单)
    
    If Not rs皮试.EOF Then
        blnHaveData = True
    Else
        blnHaveData = False
    End If
    
    Set Get皮试结果信息 = rs皮试
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CISGetNextId(strTable As String, Optional strFild As String) As Long
    '------------------------------------------------------------------------------------
    '功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值(临床)
    '参数：
    '   strTable：表名称;strFild字段名，序列名称不一定是ID，例如记录ID
    '返回：
    '------------------------------------------------------------------------------------
    '不能用错误错处理,原因是序列失效和没有序列时,应该返回错误,不然返回零,就有问题!

    Dim strJsonIn As String
    Dim strJsonOut As String

    strJsonIn = "{""input"":{""table_name"":""" & Trim(strTable) & _
        """,""col_name"":""" & Trim(strFild) & """}}"
        
    If CallService("Zl_Cissvr_Getnextid", strJsonIn, strJsonOut, "zlCISkernel.mdlCISService.CISGetNextId", p住院医嘱下达, True) Then
         CISGetNextId = Val(gobjService.GetJsonNodeValue("output.next_id"))
    End If
End Function

Public Function GetAdviceInfo(intType As Integer, ByVal str医嘱IDs As String) As Collection
'功能：获取病人医嘱基本信息扩展功能
'参数：intType 查询类型：0:查询基本信息；1:查询基本信息+扩展信息

    Dim strJsonIn As String
    Dim strJsonOut As String
 
    On Error GoTo errH

    strJsonIn = "{""input"":{""query_type"":" & intType & ",""advice_ids"":""" & str医嘱IDs & """}}"
        
    If CallService("Zl_Cissvr_Getadviceinfo", strJsonIn, strJsonOut, "zlCISkernel.mdlCISService.GetAdviceInfo", p住院医嘱下达, True) Then
        Set GetAdviceInfo = gobjService.GetJsonListValue("output.advice_list")
    End If
     
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetRs危急值记录(ByVal intType As Integer, ByVal lngIdin As Long, Optional ByVal dt开始时间 As Date, Optional ByVal dt结束时间 As Date, _
                Optional ByVal int病人类型 As Integer, Optional ByVal lng报告科室id As Long, Optional ByVal lng确认科室id As Long, Optional ByVal int确认状态 As Integer) As ADODB.Recordset
'功能：通过危急值id获取病人危急值信息,并返回记录集
'intType 1-通过id查询,4-过滤危急值列表

    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim colList As Collection
    Dim rsReturn As Recordset
    Dim i As Long
    On Error GoTo errH

    If intType = 1 Then
        strJsonIn = "{""input"":{""use_type"":1" & _
            ",""cvalue_ids"":""" & lngIdin & """}}"
    ElseIf intType = 4 Then
        strJsonIn = "{""input"":{" & _
                        """use_type"":" & "4" & "," & _
                        """cvalue_time_begin"":""" & Format(dt开始时间, "yyyy-MM-dd HH:mm") & """," & _
                        """cvalue_time_end"":""" & Format(dt结束时间, "yyyy-MM-dd HH:mm") & """," & _
                        """pati_type"":" & int病人类型 & "," & _
                        """rpt_deptid"":" & lng报告科室id & "," & _
                        """cvalue_deptid"":" & lng确认科室id & "," & _
                        """cvalue_rec_status"":" & int确认状态 & _
                    "}}"
    End If

    Set rsReturn = New ADODB.Recordset
    rsReturn.Fields.Append "ID", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "医嘱ID", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "姓名", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "性别", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "年龄", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "报告时间", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "状态", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "是否危急值", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "危急值描述", adVarChar, 4000, adFldIsNullable
    
    
    rsReturn.Fields.Append "数据来源", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "病人id", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "主页id", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "挂号单", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "婴儿", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "标本id", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "报告科室id", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "报告人", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "处理情况", adVarChar, 4000, adFldIsNullable
    rsReturn.Fields.Append "确认时间", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "确认人", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "确认科室id", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "确认科室", adVarChar, 500, adFldIsNullable
    
    rsReturn.CursorLocation = adUseClient
    rsReturn.LockType = adLockOptimistic
    rsReturn.CursorType = adOpenStatic
    rsReturn.Open

    If CallService("Zl_Cissvr_Getcriticalinfo", strJsonIn, strJsonOut, "GetRs危急值记录", p住院医嘱下达, True) Then
        Set colList = gobjService.GetJsonListValue("output.cvalue_list")
        
        If Not colList Is Nothing Then
            If colList.Count > 0 Then

                For i = 1 To colList.Count
                    rsReturn.AddNew
                    rsReturn!ID = colList(i)("_cvalue_id")
                    rsReturn!医嘱ID = colList(i)("_advice_id")
                    rsReturn!姓名 = colList(i)("_pat_name")
                    rsReturn!性别 = colList(i)("_pat_sex")
                    rsReturn!年龄 = colList(i)("_pat_age")
                    rsReturn!报告时间 = colList(i)("_cvalue_rec_create_time")
                    rsReturn!状态 = colList(i)("_cvalue_rec_status")
                    rsReturn!是否危急值 = colList(i)("_cvitem_result")
                    rsReturn!危急值描述 = colList(i)("_cvalue_rec_desc")
                    rsReturn!数据来源 = colList(i)("_cvitem_source")
                    rsReturn!病人ID = colList(i)("_pati_id")
                    rsReturn!主页ID = colList(i)("_pati_pageid")
                    rsReturn!挂号单 = colList(i)("_rgst_no")
                    rsReturn!婴儿 = colList(i)("_baby_num")
                    rsReturn!标本id = colList(i)("_lspcm_id")
                    rsReturn!报告科室id = colList(i)("_rpt_deptid")
                    rsReturn!报告人 = colList(i)("_rec_rptor")
                    rsReturn!处理情况 = colList(i)("_proc_note")
                    rsReturn!确认时间 = colList(i)("_cvalue_cnfmtime")
                    rsReturn!确认人 = colList(i)("_cvalue_cnfmer")
                    rsReturn!确认科室id = colList(i)("_cnfm_deptid")
                    rsReturn!确认科室 = colList(i)("_cvalue_dept")
                    rsReturn.Update
                Next
                rsReturn.Filter = 0
            End If
        End If
        
    End If
    Set GetRs危急值记录 = rsReturn
     
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiBaseInfo(ByVal intType As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal lng医嘱ID As Long, Optional ByVal strNO As String) As Collection
'功能：获取病人基本信息扩展功能
'参数：intType 1-通过病人id和主页id查询  2-读取医嘱记录中的病人信息,3-通过挂号单号查询

    Dim strJsonIn As String
    Dim strJsonOut As String
    On Error GoTo errH

    strJsonIn = "{""input"":{""query_type"":" & intType & _
        ",""pati_id"":" & lng病人ID & _
        ",""page_id"":" & lng主页ID & _
        ",""advice_id"":" & lng医嘱ID & ",""reg_no"":""" & strNO & """}}"
        
    If CallService("Zl_Cissvr_Getpatibaseinfo", strJsonIn, strJsonOut, "zlCISkernel.mdlCISService.GetPatiBaseInfo", p住院医嘱下达, True) Then
        Set GetPatiBaseInfo = gobjService.GetJsonListValue("output.page_list")
    End If
     
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitSvr() As Boolean
'功能：初始化服务接口部件
    If gobjService Is Nothing Then
        On Error Resume Next
        Set gobjService = CreateObject("zlServiceCall.clsServiceCall")
        If Not gobjService.InitService(gcnOracle, gstrDBUser, glngSys) Then
            Set gobjService = Nothing
        End If
        err.Clear: On Error GoTo 0
    End If
    If gobjService Is Nothing Then
        MsgBox "zlServiceCall.clsServiceCall创建失败!", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not gobjService Is Nothing Then InitSvr = True
End Function

Public Function CallService(ByVal strServiceName As String, _
    ByVal strJson_In As String, Optional ByRef strJson_out As String, Optional ByVal strTitle As String, _
    Optional lngModule As Long, Optional blnShowErrMsg As Boolean = True, Optional ByVal strAskDate As String, Optional varExpend As String, _
    Optional lngSys As Long, Optional blnReadServiceErr As Boolean) As Boolean
'功能：调用服务
'相关说明见 zlServiceCall.clsServiceCall.CallService 接口
 
    If InitSvr() Then
        If Not gobjService.CallService(strServiceName, strJson_In, strJson_out, strTitle, lngModule, blnShowErrMsg, strAskDate, varExpend, lngSys, blnReadServiceErr) Then Exit Function
        If Not blnShowErrMsg Then
            If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
                CallService = False: Exit Function
            End If
        End If
        CallService = True
    End If
End Function

Public Function GetSvrOutInfo(ByVal strJson As String, Optional ByVal blnShowErrMsg As Boolean = True) As Boolean
'功能：针对本域【临床】检查方法调用的出参解析方式
'参数：strJson前一个检查过程执行后的的出参Json格式
'      blnShowErrMsg 是否内部弹出提示信息
'返回：true/false  code=1时为true,code=0时为false
    If InitSvr() Then
        GetSvrOutInfo = gobjService.SetJsonString(strJson)
        
        If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
            GetSvrOutInfo = False
            If blnShowErrMsg Then
                MsgBox gobjService.GetJsonNodeValue("output.message") & "", vbInformation, gstrSysName
            End If
        End If
    End If
End Function

Public Function GetToDateStr(ByVal strIn As String) As String
'功能：从oracle日期转换的串中解析日期
    Dim strTmp As String
    strTmp = Split(strIn, ",")(0)
    strTmp = Split(strTmp, "'")(1)
    GetToDateStr = strTmp
End Function
  
Public Function Exe医嘱取消执行完成(colProSQL As Collection, ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal lng单独执行 As Long, ByVal lng科室ID As Long, ByVal strTitle As String, ByVal lng模块 As Long) As Boolean
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim strFCOut As String '费用检查执行后的出参
    Dim lng费用性质 As String
    Dim blnTran As Boolean
    Dim strStuff As String
    Dim strDrug As String
    Dim i As Long
    Dim strFeeState As String
    Dim clFst As Collection '费用执行状态的更新的列表
    
    On Error GoTo errH
    
    strJsonOut = zlDatabase.CallProcedure("Zl_病人医嘱执行_Cancel_Check", strTitle & "_" & lng模块, _
                lng医嘱ID, lng发送号, lng单独执行, Empty)
    If Not GetSvrOutInfo(strJsonOut) Then
        Exit Function
    End If
    
    lng费用性质 = Val("" & gobjService.GetJsonNodeValue("output.fee_origin"))
    
    strJsonIn = "{""input"":{""is_finish"":2,""fee_origin"":" & lng费用性质 & ",""fee_nos"":""" & gobjService.GetJsonNodeValue("output.fee_nos") _
        & """,""order_status"":" & gobjService.GetJsonNodeValue("output.order_status") & ",""order_ids"":""" & gobjService.GetJsonNodeValue("output.order_ids") & """,""exe_deptid"":" & lng科室ID & "}}"
    
    If Not CallService("Zl_Exsesvr_GetOrderFeeExeInfo", strJsonIn, strJsonOut, strTitle, lng模块, True) Then
        Exit Function
    End If
    
    strStuff = gobjService.GetJsonNodeValue("output.stuffdtl_ids") & ""
    strDrug = gobjService.GetJsonNodeValue("output.rcpdtl_ids") & ""
    strFCOut = strJsonOut
    Set clFst = gobjService.GetJsonListValue("output.cancel_list")
    If strStuff <> "" Then
        '退料服务
        strFeeState = Get取消执行费用状态(strStuff, lng费用性质)
        strJsonIn = "{""input"":{""audit_operator"":""" & UserInfo.姓名 & """,""stuffdtl_ids"":""" & Replace(strStuff & ":", ",", ":,") & """}}"
        Call Exe自动退药退料(strJsonIn, strFeeState, "", "", strTitle, lng模块)
    End If
    
    If strDrug <> "" Then
        strFeeState = Get取消执行费用状态(strDrug, lng费用性质)
        '退药服务 数量都拼接 空进去
        strJsonIn = Get自动退药入参(strDrug)
        Call Exe自动退药退料("", "", strJsonIn, strFeeState, strTitle, lng模块)
    End If

    strJsonIn = ""
    strFeeState = ""
    If Not clFst Is Nothing Then
        If clFst.Count > 0 Then
            For i = 1 To clFst.Count
                strFeeState = strFeeState & ",{" & _
                    """fee_id"":" & clFst(i)("_fee_id") & _
                   ",""exe_nums"":0" & _
                   "}"
            Next
            strJsonIn = "{""input"":{""fee_origin"":" & lng费用性质 & ",""item_list"":[" & Mid(strFeeState, 2) & "]}}"
        End If
    End If
    
    gcnOracle.BeginTrans: blnTran = True
    For i = 1 To colProSQL.Count
        Call zlDatabase.ExecuteProcedure(colProSQL(i) & "", strTitle)
    Next
    If strJsonIn <> "" Then
        Call CallService("Zl_Exsesvr_UpdateExeInfo", strJsonIn, , strTitle, lng模块, False, , , , True)
    End If
    gcnOracle.CommitTrans: blnTran = False
    
    '最后再更新一次费用执行状态
    '因为前面有可能会失败
    
    Call Upd费用执行状态(lng医嘱ID, lng发送号)
    
    Exe医嘱取消执行完成 = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get自动退药入参(ByVal strIDs As String) As String
'功能：自动退药时的入参
    Dim i As Long, varTmp As Variant, strItem As String
    
    On Error GoTo errH
'    strJsonIn = "{""input"":{""audit_operator"":""" & UserInfo.姓名 & """,""rcpdtl_ids"":""" & Replace(strDrug & ":", ",", ":,") & """}}"
    varTmp = Split(strIDs, ",")
    For i = 0 To UBound(varTmp)
        strItem = strItem & ",{""rcpdtl_id"":" & varTmp(i)
        strItem = strItem & "}"
    Next
    Get自动退药入参 = "{""input"":{""audit_operator"":""" & UserInfo.姓名 & """,""rcpdtl_list"":[" & Mid(strItem, 2) & "]}}"
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get取消执行费用状态(ByVal strIDs As String, ByVal lng_fee_origin As Long) As String
'功能：根据费用明细获取取消执行费用服务入参
    Dim i As Long, strItem As String
    Dim varTmp As Variant, strTmp As String
    
    On Error GoTo errH
    
    varTmp = Split(strIDs, ",")
    For i = 0 To UBound(varTmp)
        strItem = strItem & ",{""fee_id"":" & varTmp(i)
        strItem = strItem & ",""exe_nums"":0"
        strItem = strItem & "}"
    Next
    
    strTmp = "{""input"":{""fee_origin"":" & lng_fee_origin & ",""item_list"":[" & Mid(strItem, 2) & "]}}"
    
    Get取消执行费用状态 = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Exe医嘱执行完成(colProSQL As Collection, ByVal lng医嘱ID As String, ByVal lng发送号 As Long, ByVal lng单独执行 As Long, ByVal lng科室ID As Long, _
   ByVal strTime As String, ByVal strTitle As String, ByVal lng模块 As Long) As Boolean
'功能:医嘱执行完成
'参数:colProSQL- 临床域可执行过程集合,可能有多个
'     lng医嘱ID- 执行完成的医嘱ID
'     lng发送号- 发送号
'     lng单独执行- 是否单独执行,0-不单独执行,1-单独执行
'     lng科室ID- 执行科室ID 关联费用记录的执行科室ID,0-不区分科室,1-指定科室
'     strTime- 执行时间,字符串日期格式:
'     strTitle- 标题
'     lng模块- 调用的模块号
'返回:true 成功,false 失败
    Dim i As Long
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim blnTran As Boolean
    Dim strSvrName As String
    Dim colPar As New Collection
    Dim clTag As Collection
     
    On Error GoTo errH
 
    strJsonOut = zlDatabase.CallProcedure("Zl_病人医嘱执行_Finish_Check", strTitle & "_" & lng模块, _
                lng医嘱ID, lng发送号, lng单独执行, Empty)
    If Not GetSvrOutInfo(strJsonOut) Then
        Exit Function
    End If
    
    If Not F医嘱执行完成(colProSQL, strJsonOut, strTime, lng科室ID, strTitle, lng模块, colPar) Then
        Exit Function
    End If
    
    
'C "药品发药"
'C "卫材发料"
'C "费用执行"
 
'Col "药品清异常"
'C "药品收费"

'Col "卫材清异常"
'C "卫材收费"

'Col "医嘱加异常"
'C "费用收费"


'Col "医嘱执行完成"

    strJsonIn = GetColVal(colPar, "费用收费")
    Set clTag = GetColObj(colPar, "医嘱加异常")
    strSvrName = "Zl_Exsesvr_Billverify"
    If strJsonIn <> "" Then
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To clTag.Count
            Call zlDatabase.ExecuteProcedure(clTag(i) & "", strTitle)
        Next
        Call CallService(strSvrName, strJsonIn, , strTitle, lng模块, False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    
    Call Exe药品收费确认(colPar, strTitle, lng模块)
    
    Call Exe卫材收费确认(colPar, strTitle, lng模块)
     
    strJsonIn = GetColVal(colPar, "药品发药")
    strSvrName = "Zl_Drugsvr_Autosenddrug"
    If strJsonIn <> "" Then
        Call CallService(strSvrName, strJsonIn, , strTitle, lng模块, False, , , , True)
    End If
    
    strJsonIn = GetColVal(colPar, "卫材发料")
    strSvrName = "Zl_Stuffsvr_Autosendstuff"
    If strJsonIn <> "" Then
        Call CallService(strSvrName, strJsonIn, , strTitle, lng模块, False, , , , True)
    End If
 
    strJsonIn = GetColVal(colPar, "费用执行")
    Set clTag = GetColObj(colPar, "医嘱执行完成")
    strSvrName = "Zl_Exsesvr_Updateexeinfo"
    If strJsonIn <> "" Or clTag.Count > 0 Then
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To clTag.Count
            Call zlDatabase.ExecuteProcedure(clTag(i) & "", strTitle)
        Next
        If strJsonIn <> "" Then Call CallService(strSvrName, strJsonIn, , strTitle, lng模块, False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    Exe医嘱执行完成 = True
    
    Call Upd费用执行状态(lng医嘱ID, lng发送号)
    
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Exe医嘱执行登记删除(colData As Collection, ByVal strTitle As String, ByVal lng模块 As Long) As Boolean
'功能:删除医嘱执行登记记录
    Dim strJsonOut As String
    Dim strItems As String
    Dim strOrderDel As String
    Dim strFeeSta As String
    Dim blnTran As Boolean
    Dim colProSQL As Collection
    Dim strTmp As String
    Dim strPJson As String '公用结点
    Dim colSQLlist As Collection
    Dim i As Long
    
    On Error GoTo errH
            
    '如果是自动取消执行完成需先调用
    If Val(GetColVal(colData, "自动取消") & "") = 1 Then
        strJsonOut = zlDatabase.CallProcedure("Zl_病人医嘱执行_Delete_check", strTitle & "_" & lng模块, _
                Val(GetColVal(colData, "医嘱ID")), Val(GetColVal(colData, "发送号")), _
                GetColVal(colData, "执行时间") & "", Val(GetColVal(colData, "单独执行")), 1, 1, Empty)
        If Not GetSvrOutInfo(strJsonOut) Then
            Exit Function
        End If
     
        If Val(gobjService.GetJsonNodeValue("output.auto_cancel") & "") = 1 Then
            Set colProSQL = New Collection
            strTmp = "Zl_病人医嘱执行_Cancel_S(" & colData("医嘱ID") & "," & colData("发送号") & ",'" & UserInfo.姓名 & "'," & colData("单独执行") & ",null)"
            colProSQL.Add strTmp
            If Not Exe医嘱取消执行完成(colProSQL, colData("医嘱ID"), colData("发送号"), colData("单独执行"), colData("执行部门ID"), strTitle, lng模块) Then
                Exit Function
            End If
        End If
    End If
    
    '删除执行登记
    strJsonOut = zlDatabase.CallProcedure("Zl_病人医嘱执行_Delete_check", strTitle & "_" & lng模块, _
                Val(GetColVal(colData, "医嘱ID")), Val(GetColVal(colData, "发送号")), _
                GetColVal(colData, "执行时间") & "", Val(GetColVal(colData, "单独执行")), 0, 0, Empty)
    If Not GetSvrOutInfo(strJsonOut) Then
        Exit Function
    End If
    
    strItems = strPJson
    strItems = strItems & ",""order_status"":" & Val(gobjService.GetJsonNodeValue("output.order_status") & "") '医嘱执行状态
    strItems = strItems & ",""upd_price_sta"":" & Val(gobjService.GetJsonNodeValue("output.upd_price_sta") & "")
    strItems = strItems & ",""price_req_time"":""" & gobjService.GetJsonNodeValue("output.price_req_time") & """"
    strItems = strItems & ",""lis_upd"":" & Val(gobjService.GetJsonNodeValue("output.lis_upd") & "")
    '医嘱删除执行登记的入参JSON拼装完成
    strOrderDel = "{""input"":{" & Mid(strItems, 2) & "}}"
    
    strOrderDel = "Zl_病人医嘱执行_Delete_s(" & Val(GetColVal(colData, "医嘱ID")) & "," & Val(GetColVal(colData, "发送号")) & _
                ",to_date('" & GetColVal(colData, "执行时间") & "','yyyy-mm-dd hh24:mi:ss')" & _
                "," & Val(GetColVal(colData, "单独执行")) & _
                "," & Val(gobjService.GetJsonNodeValue("output.upd_price_sta") & "") & _
                ",to_date('" & gobjService.GetJsonNodeValue("output.price_req_time") & "','yyyy-mm-dd hh24:mi:ss')" & _
                "," & Val(gobjService.GetJsonNodeValue("output.lis_upd") & "") & _
                "," & Val(gobjService.GetJsonNodeValue("output.order_status") & "") & _
                ")"
    
    strFeeSta = strFeeSta & ",""oper_type"":1"
    strFeeSta = strFeeSta & ",""fee_origin"":" & Val(gobjService.GetJsonNodeValue("output.fee_origin") & "") '费用来源(默认=2：1-门诊费用，2-住院费用)
    strFeeSta = strFeeSta & ",""exe_status"":" & Val(gobjService.GetJsonNodeValue("output.fee_status") & "")
    strFeeSta = strFeeSta & "," & GetJsonStrNode("exe_deptid", "执行部门ID", "N", colData)
    strFeeSta = strFeeSta & ",""order_ids"":""" & gobjService.GetJsonNodeValue("output.order_ids") & """"
    strFeeSta = strFeeSta & ",""fee_nos"":""" & gobjService.GetJsonNodeValue("output.fee_nos") & """"
    strFeeSta = "{""input"":{" & Mid(strFeeSta, 2) & "}}" '更新费用执行状态的入参JSON拼装完成
    
    Set colSQLlist = GetColObj(colData, "sqllist")
    
    
    gcnOracle.BeginTrans: blnTran = True
 
        Call zlDatabase.ExecuteProcedure(strOrderDel, "Exe医嘱执行登记删除")
        If Not CallService("Zl_Exsesvr_Updateexeinfo", strFeeSta, strJsonOut, strTitle, lng模块, False) Then
            gcnOracle.RollbackTrans: blnTran = False
            MsgBox "调用[Zl_Exsesvr_Updateexeinfo]失败：" & gobjService.GetJsonNodeValue("output.message"), vbInformation, gstrSysName
            Exit Function
        End If
        
        If Not colSQLlist Is Nothing Then
            For i = 1 To colSQLlist.Count
                If colSQLlist(i) <> "" Then
                    Call zlDatabase.ExecuteProcedure(colSQLlist(i), "Exe医嘱执行登记删除")
                End If
            Next
        End If
        
        
    gcnOracle.CommitTrans: blnTran = False
    
    Exe医嘱执行登记删除 = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Exe医嘱执行登记(colData As Collection, ByVal strTitle As String, ByVal lng模块 As Long) As Boolean
'功能:医嘱执行登记
    Dim strJsonOut As String
    Dim lng自动完成  As Long
    Dim strOrderIns As String
    Dim strFeeSta As String
    Dim blnTran As Boolean
    Dim colProSQL As Collection
    Dim strTmp As String
    Dim colSQLAfter As Collection, i As Long
  
    On Error GoTo errH
    
    strJsonOut = zlDatabase.CallProcedure("Zl_病人医嘱执行_Insert_check", strTitle & "_" & lng模块, _
                Val(GetColVal(colData, "医嘱ID")), Val(GetColVal(colData, "发送号")), CDate(GetColVal(colData, "要求时间")), _
                Val(GetColVal(colData, "本次数次")), _
                CDate(GetColVal(colData, "执行时间")), Val(GetColVal(colData, "单独执行")), Val(GetColVal(colData, "自动完成")), Val(GetColVal(colData, "执行结果")), _
                Val(GetColVal(colData, "配液检查")), Val(GetColVal(colData, "检验项目记帐")), Val(GetColVal(colData, "调用场合")), _
                Empty)
                 
    If Not GetSvrOutInfo(strJsonOut) Then
        Exit Function
    End If
    
    lng自动完成 = Val(gobjService.GetJsonNodeValue("output.auto_finish") & "")
    
    
    
'医嘱id_In       In Number,
'发送号_In       In Number,
'要求时间_In     In Date,
'本次数次_In     In Number,
'执行摘要_In     In Varchar2,
'执行人_In       In Varchar2,
'执行时间_In     In Date,
'执行结果_In     In Number,
'未执行原因_In   In Varchar2,
'操作员姓名_In   In Varchar2,
'执行部门id_In   In Number,
'输液通道_In     In Varchar2,
'记录来源_In     In Number,
'执行方式_In     In Number,
'病人id_In       In Number, --pati_id N 1 病人id
'主页id_In       In Number, --pati_pageid N 1 主页id
'挂号单_In       In Varchar2, --reg_no C 1 挂号单号
'执行状态_In     In Number, --exe_status N 1 执行状态
'检验采集更新_In In Number, --lis_upd N 1 更新检验采集
'组id_In         In Number, --order_main_id N 1 组id,主医嘱id
'更新计价状态_In In Number, --upd_price_sta N 1 计价状态更新
'计价要求时间_In In Date, --price_req_time C 1 计价要求时间
'医嘱ids_In      In Varchar2 --order_ids C 1 要更新状态的医嘱id串

    strOrderIns = "Zl_病人医嘱执行_Insert_S(" & Val(GetColVal(colData, "医嘱ID")) & "," & Val(GetColVal(colData, "发送号")) & ",to_date('" & GetColVal(colData, "要求时间") & "','yyyy-mm-dd hh24:mi:ss')," & Val(GetColVal(colData, "本次数次"))
    strOrderIns = strOrderIns & ",'" & GetColVal(colData, "执行摘要") & "','" & GetColVal(colData, "执行人") & "',to_date('" & GetColVal(colData, "执行时间") & "','yyyy-mm-dd hh24:mi:ss')," & GetColVal(colData, "执行结果")
    strOrderIns = strOrderIns & ",'" & GetColVal(colData, "未执行原因") & "','" & GetColVal(colData, "操作员姓名") & "'," & Val(GetColVal(colData, "执行部门ID"))
    strOrderIns = strOrderIns & ",'" & IIF("NULL" = UCase(GetColVal(colData, "输液通道")), "", GetColVal(colData, "输液通道")) & "'"
    strOrderIns = strOrderIns & "," & Val(GetColVal(colData, "记录来源")) & "," & Val(GetColVal(colData, "执行方式"))
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.pati_id") & "")
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.pati_pageid") & "")
    strOrderIns = strOrderIns & ",'" & gobjService.GetJsonNodeValue("output.reg_no") & "'"
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.order_status") & "")
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.lis_upd") & "")
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.order_main_id") & "")
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.upd_price_sta") & "")
    strOrderIns = strOrderIns & ",to_date('" & gobjService.GetJsonNodeValue("output.price_req_time") & "','yyyy-mm-dd hh24:mi:ss')"
    strOrderIns = strOrderIns & ",'" & gobjService.GetJsonNodeValue("output.order_ids") & "'"
    strOrderIns = strOrderIns & ")"
    
    strFeeSta = strFeeSta & ",""oper_type"":1,""exe_type"":1"
    strFeeSta = strFeeSta & ",""fee_origin"":" & Val(gobjService.GetJsonNodeValue("output.fee_origin") & "")
    strFeeSta = strFeeSta & ",""exe_status"":" & Val(gobjService.GetJsonNodeValue("output.fee_status") & "")
    strFeeSta = strFeeSta & "," & GetJsonStrNode("exe_deptid", "执行部门ID", "N", colData)
    strFeeSta = strFeeSta & "," & GetJsonStrNode("exe_people", "执行人", "C", colData)
    strFeeSta = strFeeSta & ",""exe_time"":""" & gobjService.GetJsonNodeValue("output.fee_exe_time") & """"
    strFeeSta = strFeeSta & ",""order_ids"":""" & gobjService.GetJsonNodeValue("output.order_ids") & """"
    strFeeSta = strFeeSta & ",""fee_nos"":""" & gobjService.GetJsonNodeValue("output.fee_nos") & """"
    strFeeSta = "{""input"":{" & Mid(strFeeSta, 2) & "}}" '费用执行状态更新入参封装完成
    
    
    Set colSQLAfter = GetColObj(colData, "sqllistafter")
    
    gcnOracle.BeginTrans: blnTran = True
        Call zlDatabase.ExecuteProcedure(strOrderIns, "Exe医嘱执行登记")
        If Not CallService("Zl_Exsesvr_Updateexeinfo", strFeeSta, strJsonOut, strTitle, lng模块, False) Then
            gcnOracle.RollbackTrans: blnTran = False
            MsgBox "调用[Zl_Exsesvr_Updateexeinfo]失败：" & gobjService.GetJsonNodeValue("output.message"), vbInformation, gstrSysName
            Exit Function
        End If
        
        
        
        If Not colSQLAfter Is Nothing Then
            For i = 1 To colSQLAfter.Count
                If colSQLAfter(i) <> "" Then
                    Call zlDatabase.ExecuteProcedure(colSQLAfter(i), "Exe医嘱执行登记")
                End If
            Next
        End If
        
    gcnOracle.CommitTrans: blnTran = False
    
    If lng自动完成 = 1 Then
        Set colProSQL = New Collection
        strTmp = "Zl_病人医嘱执行_Finish_S(" & colData("医嘱ID") & "," & colData("发送号") & ",to_date('" & colData("执行时间") & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.姓名 & "'," & colData("单独执行") & ")"
        colProSQL.Add strTmp
        Call Exe医嘱执行完成(colProSQL, colData("医嘱ID"), colData("发送号"), colData("单独执行"), colData("执行部门ID"), colData("执行时间"), strTitle, lng模块)
    End If
    
    Exe医嘱执行登记 = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Exe医嘱执行登记修改(colData As Collection, ByVal strTitle As String, ByVal lng模块 As Long, Optional colBefore As Collection, Optional colAfter As Collection) As Boolean
'功能：修改医嘱执行登记记录,原过程:Zl_病人医嘱执行_Update
'参数:colData 入参数"原执行时间,医嘱ID,发送号,要求时间,本次数次,执行摘要,执行人,执行时间,执行结果,未执行原因,单独执行,操作员编号,操作员姓名,执行部门ID"
'     strTitle- 标题
'     lng模块- 调用的模块号
'     colBefore  临床域可执行过程,可放到修改执行登记过程前执行的过程
'     colAfter   临床域可执行过程,可放到修改执行登记过程后执行的过程
'说明:过程内部会开启事务
'    zlAdviceRegRecUpd = AdviceRegRecUpd(colData, strTitle, lng模块, colBefore, colAfter)
    Dim i As Long
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim blnTran As Boolean
    Dim strFeeSta As String
    Dim strCisOutPar As String
    
    On Error GoTo errH
    
    strJsonIn = "{""input"":{""old_exe_time"":""" & GetColVal(colData, "原执行时间") & """" & _
            ",""order_id"":" & GetColVal(colData, "医嘱ID") & _
            ",""send_no"":" & GetColVal(colData, "发送号") & _
            ",""require_time"":""" & GetColVal(colData, "要求时间") & """" & _
            ",""reg_num"":" & GetJsonNum(GetColVal(colData, "本次数次")) & _
            ",""exe_memo"":""" & GetColVal(colData, "执行摘要") & """" & _
            ",""exe_people"":""" & GetColVal(colData, "执行人") & """" & _
            ",""exe_time"":""" & GetColVal(colData, "执行时间") & """" & _
            ",""exe_result"":" & GetColVal(colData, "执行结果") & _
            ",""no_exe_rea"":""" & GetColVal(colData, "未执行原因") & """" & _
            ",""exe_alone"":" & GetColVal(colData, "单独执行") & _
            ",""operator_name"":""" & GetColVal(colData, "操作员姓名") & """" & _
            "}}"
    
    
    
    gcnOracle.BeginTrans: blnTran = True
        If Not colBefore Is Nothing Then
            If colBefore.Count > 0 Then
                For i = 1 To colBefore.Count
                    Call zlDatabase.ExecuteProcedure(colBefore(i) & "", strTitle)
                Next
            End If
        End If
    
        strCisOutPar = zlDatabase.CallProcedure("Zl_病人医嘱执行_Update_S", strTitle & "_" & lng模块, CDate(GetColVal(colData, "原执行时间", , "0")), _
            Val(GetColVal(colData, "医嘱ID")), Val(GetColVal(colData, "发送号")), _
            CDate(GetColVal(colData, "要求时间", , "0")), Val(GetColVal(colData, "本次数次")), GetColVal(colData, "执行摘要"), GetColVal(colData, "执行人"), _
            CDate(GetColVal(colData, "执行时间", , "0")), Val(GetColVal(colData, "执行结果")), GetColVal(colData, "未执行原因"), Val(GetColVal(colData, "单独执行")), _
            UserInfo.编号, UserInfo.姓名, Val(GetColVal(colData, "执行部门ID", "N", "0")), _
            Empty)
        
        
        If Not colAfter Is Nothing Then
            If colAfter.Count > 0 Then
                For i = 1 To colAfter.Count
                    Call zlDatabase.ExecuteProcedure(colAfter(i) & "", strTitle)
                Next
            End If
        End If
        If strCisOutPar <> "" Then
            Call gobjService.SetJsonString(strCisOutPar)
            strFeeSta = strFeeSta & ",""oper_type"":1,""exe_type"":1"
            strFeeSta = strFeeSta & ",""fee_origin"":" & Val(gobjService.GetJsonNodeValue("output.fee_origin") & "")
            strFeeSta = strFeeSta & ",""exe_status"":" & Val(gobjService.GetJsonNodeValue("output.fee_status") & "")
            strFeeSta = strFeeSta & ",""exe_deptid"":" & GetColVal(colData, "执行部门ID", "N", "0")
            strFeeSta = strFeeSta & ",""exe_people"":""" & gobjService.GetJsonNodeValue("output.fee_exe_peo") & """"
            strFeeSta = strFeeSta & ",""exe_time"":""" & gobjService.GetJsonNodeValue("output.fee_exe_time") & """"
            strFeeSta = strFeeSta & ",""order_ids"":""" & gobjService.GetJsonNodeValue("output.order_ids") & """"
            strFeeSta = strFeeSta & ",""fee_nos"":""" & gobjService.GetJsonNodeValue("output.fee_nos") & """"
            
            strFeeSta = "{""input"":{" & Mid(strFeeSta, 2) & "}}" '费用执行状态更新入参封装完成
            Call CallService("Zl_Exsesvr_Updateexeinfo", strFeeSta, strJsonOut, strTitle, lng模块, False, , , , True)
        End If
    gcnOracle.CommitTrans: blnTran = False
    
    Exe医嘱执行登记修改 = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PatiSvrGetpatiinfo(intType As Integer, lng病人ID As Long, Optional lngModel As Long, Optional int包含家属 As Integer) As Collection
'  ---------------------------------------------------------------------------
'  --功能:获取病人信息（目前只用于通过病人id返回全部信息）
'  --入参：Json_In:格式
'  --    input
'  --      pati_id           N 1 病人id  病人ID<>0时，查询列表中的条件无效
'  --      query_type        N 1 查询类型:如：0-基本;1-基本+联系人;3-所有
'  --      query_card        N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
'  --      query_family      N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
'  --      query_drug        N 1 是否包含过敏药物:1-包含，0-不包含
'  --      query_immune      N 1 是否包含免疫修:1-包含;0-不包含
'  --      query_cons_list[] C 1 查询条件:可以选择一定条件进行查询（是And关系),只有一行
'  --        pati_ids        C   病人IDs:多个用逗号
'  --        pati_name       C   姓名:可以代%分号表表按姓名匹配
'  --        outpno  N   门诊号
'  --        pati_idcard     C   身份证号
'  --        contacts_idcard C   联系人身份证号
'  --        cardtype_id     N   医疗卡类别ID
'  --        card_no         C   卡号
'  --        qrcode          C   二维码
'  --        iccard_no       C   Ic卡号
'  --        insurance_num   C   医保号
'  --        qrspt_statu     C   查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
'  --        phone_number    C   手机号
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""query_type"":" & intType & ",""pati_id"":" & lng病人ID & IIF(int包含家属 = 0, "", ",""query_family"":1") & "}}"
        
    If CallService("Zl_Patisvr_Getpatiinfo", strJsonIn, strJsonOut, "PatiSvrGetpatiinfo", lngModel, False, , , , True) Then
        Set PatiSvrGetpatiinfo = gobjService.GetJsonListValue("output.pati_list")
    End If
End Function

Public Function PatiSvrGetVisitPatis(str病人IDs As String, Optional str就诊卡号 As String, Optional lngModel As Long) As Collection
'  ---------------------------------------------------------------------------
'  --功能:获取病人信息
'  --入参：Json_In:格式
'  --    input
'  --      pati_ids          C   病人IDs:多个用逗号
'  --      vcard_no          C   就诊卡号
'  -------------------------------------------
    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""pati_ids"":""" & str病人IDs & """,""vcard_no"":""" & str就诊卡号 & """}}"
    If CallService("Zl_Patisvr_Getvisitpatis", strJsonIn, strJsonOut, "PatiSvrGetVisitPatis", lngModel, False, , , , True) Then
        Set PatiSvrGetVisitPatis = gobjService.GetJsonListValue("output.pati_list", "pati_id")
    End If
End Function

Public Function ExseSvrGetPatisurplusinfo(str病人IDs As String, Optional lngModel As Long) As Collection
'  ---------------------------------------------------------------------------
'  --功能:获取病人费用余额和预交余额
'  --入参：Json_In:格式
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    
    strJsonIn = "{""input"":{""pati_ids"":""" & str病人IDs & """}}"
    If CallService("Zl_Exsesvr_Getpatisurplusinfo", strJsonIn, strJsonOut, "ExseSvrGetPatisurplusinfo", lngModel, False, , , , True) Then
        Set ExseSvrGetPatisurplusinfo = gobjService.GetJsonListValue("output.surplus_list")
    End If
End Function



Public Function ExseSvrGetremainmoney(lng病人ID As Long, lng主页ID As Long, str剩余款 As String, Optional str担保额 As String, Optional str预结费用 As String, Optional lngModel As Long) As Boolean
'  ---------------------------------------------------------------------------
'  --功能:获取病人剩余款和担保额
'  --入参：Json_In:格式
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""pati_id"":" & lng病人ID & ",""pati_pageid"":" & lng主页ID & ",""insure_account_balance"":0}}"
    If CallService("Zl_Exsesvr_Getremainmoney", strJsonIn, strJsonOut, "ExseSvrGetremainmoney", lngModel, False, , , , True) Then
        str剩余款 = gobjService.GetJsonNodeValue("output.remain_money") & ""
        str担保额 = gobjService.GetJsonNodeValue("output.guarantee_money") & ""
        str预结费用 = gobjService.GetJsonNodeValue("output.expected_money") & ""
    End If
    ExseSvrGetremainmoney = True
End Function

Public Function PatiSvrGetPatiExtendInfo(lng病人ID As Long, lng就诊ID As Long, str信息值 As String, Optional lngModel As Long) As Collection
'  ---------------------------------------------------------------------------
'  --功能:获取病人信息从表
'  --入参：Json_In:格式
'  -------------------------------------------
    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""pati_id"":" & lng病人ID & ",""visit_id"":" & lng就诊ID & ",""info_names"":""" & str信息值 & """}}"
    If CallService("Zl_Patisvr_Getpatiextendinfo", strJsonIn, strJsonOut, "PatiSvrGetPatiExtendInfo", lngModel, False, , , , True) Then
        Set PatiSvrGetPatiExtendInfo = gobjService.GetJsonListValue("output.slave_list")
    End If
End Function

Public Function ExseSvrGetinsureinfo(int险类 As Integer, lng收费细目ID As Long, str保险大类 As String, bln已设置 As Boolean, Optional lngModel As Long) As Boolean
'  ---------------------------------------------------------------------------
'  --功能:获取医保大类相关信息
'  --入参：Json_In:格式
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""insurance_type"":" & int险类 & ",""fee_item_id"":" & lng收费细目ID & "}}"
    If CallService("Zl_Exsesvr_GetInsureIteminfo", strJsonIn, strJsonOut, "ExseSvrGetinsureinfo", lngModel, False, , , , True) Then
        str保险大类 = gobjService.GetJsonNodeValue("output.insure_name")
        bln已设置 = Val(gobjService.GetJsonNodeValue("output.isexist")) > 0
    End If
    ExseSvrGetinsureinfo = True
End Function


Public Function ExseSvrGetBillDetailInfo(str费用ids As String, str费用Nos As String, Optional lngModel As Long) As Collection
'  ---------------------------------------------------------------------------
'  --功能：获取和药品发药业务相关的费用信息，主要用于界面显示
'  --入参：json格式
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""fee_ids"":""" & str费用ids & """,""bill_nos"":""" & str费用Nos & """}}"
    If CallService("Zl_Exsesvr_GetBillDetailInfo", strJsonIn, strJsonOut, "ExseSvrGetBillDetailInfo", lngModel, False, , , , True) Then
        Set ExseSvrGetBillDetailInfo = gobjService.GetJsonListValue("output.fee_list")
    End If
End Function
 
Public Function Update病人复诊标志(lng挂号ID As Long, int复诊标志 As Integer) As Boolean
'  --功能：门诊病人复诊标志

    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim strSQL As String
    Dim blnTrans As Boolean
    
    On Error GoTo errH
    strJsonIn = "{""input"":{""reg_id"":" & lng挂号ID & _
        ",""revst_sign"":" & int复诊标志 & "}}"
        
    strSQL = "Zl_病人就诊记录_复诊(" & lng挂号ID & "," & int复诊标志 & ")"
    
   gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, "标记为复诊")
    Call CallService("Zl_Exsesvr_Updateoutprevstsign", strJsonIn, strJsonOut, "Update病人复诊标志", p门诊医嘱下达, False, , , , True)
    gcnOracle.CommitTrans: blnTrans = False
    Update病人复诊标志 = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Check护理等级变更(ByVal lng病人ID As Long, ByVal lng主页ID As Long, dt登记时间 As Date, lngModel As Long) As String
'功能：护理等级变更前检查

    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim strOut As String
    
    
    strJsonIn = "{""input"":{""pati_id"":" & lng病人ID & ",""page_id"":" & lng主页ID & ",""create_time"":""" & Format(dt登记时间, "yyyy-MM-dd HH:mm:ss") & """}}"
    If CallService("Zl_Exsesvr_Chkpatichangenurse", strJsonIn, strJsonOut, "Check护理等级变更", lngModel, False) = False Then
        If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
            strOut = gobjService.GetJsonNodeValue("output.message")
        End If
    End If
    
    Check护理等级变更 = strOut
End Function


Public Function Check医嘱停止(ByVal lng医嘱ID As Long, dt终止时间 As Date, ByVal int内部调用 As Integer, ByVal int医师资格 As Integer, ByVal int停嘱审核 As Integer, ByVal lngModel As Long) As Boolean
'功能：医嘱停止前检查
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rs护理等级 As ADODB.Recordset
    Dim strOut As String
    
    On Error GoTo errH
    
    If int内部调用 = 0 And (int医师资格 > 0 Or int停嘱审核 = 0) Then
        strSQL = "Select a.医嘱状态, a.医嘱内容, a.病人id, a.主页id, Nvl(a.婴儿, 0) as 婴儿, Nvl(a.诊疗类别, '*') As 诊疗类别, b.操作类型" & vbNewLine & _
                "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
                "Where a.诊疗项目id = b.Id(+) And a.Id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Check医嘱停止", lng医嘱ID)
        
        If Not rsTmp.EOF Then
            If rsTmp!诊疗类别 & "" = "H" And rsTmp!操作类型 & "" = "1" And Val(rsTmp!婴儿 & "") = 0 Then
                '获取病人护理等级id
                strSQL = "Select c.收费细目id" & vbNewLine & _
                        "From 病人医嘱记录 A, 病人医嘱计价 C, 收费项目目录 D" & vbNewLine & _
                        "Where a.Id = c.医嘱id And c.收费细目id = d.Id And d.类别 = 'H' And Nvl(d.项目特性, 0) <> 0 And a.Id = Id_In And Rownum = 1 And" & vbNewLine & _
                        "      Exists (Select 1 From 病案主页 Where 病人id =[1] And 主页id = [2] And 护理等级id = c.收费细目id)"
                Set rs护理等级 = zlDatabase.OpenSQLRecord(strSQL, "Check医嘱停止", Val(rsTmp!病人ID & ""), Val(rsTmp!主页ID & ""))
                
                If Not rs护理等级 Is Nothing Then
                    If Not rs护理等级.EOF Then
                        If Val(rs护理等级!收费细目id & "") <> 0 Then
                             strOut = Check护理等级变更(Val(rsTmp!病人ID & ""), Val(rsTmp!主页ID & ""), dt终止时间, lngModel)
                             If strOut <> "" Then
                                MsgBox strOut, vbInformation, gstrSysName
                                Exit Function
                             End If
                        End If
                    End If
                End If
                
            End If
        End If
    End If
    Check医嘱停止 = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Check医嘱校对(ByVal lngModel As Long, ByVal lng医嘱ID As Long, int状态 As Integer, ByVal dt校对时间 As Date, Optional ByVal str校对说明 As String, Optional ByVal int自动校对 As Integer _
                    , Optional ByVal str操作员编号 As String, Optional ByVal str操作员姓名 As String, Optional str预约入院SQL As String) As Boolean
'功能：医嘱校对前检查
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rs护理等级 As ADODB.Recordset
    Dim rsTmp1 As ADODB.Recordset
    Dim strOut As String
    Dim colPi As Collection
    Dim colPati As Collection
    
    On Error GoTo errH
    
    str预约入院SQL = ""
    If int状态 = 3 Then
        strSQL = "Select a.医嘱期效, a.医嘱状态, a.开嘱时间, a.开嘱医生, a.开始执行时间, a.病人id, a.主页id, a.婴儿, a.医嘱内容, a.诊疗类别, a.诊疗项目id, a.前提id," & vbNewLine & _
                    "       Nvl(b.操作类型, '0') as 操作类型, Nvl(a.执行标记, 0), a.执行科室id, a.标本部位, a.开嘱科室id, Nvl(a.紧急标志, 0) As 紧急标志, a.病人科室id" & vbNewLine & _
                    "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
                    "Where a.诊疗项目id = b.Id(+) And a.Id = [1]"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Check医嘱停止", lng医嘱ID)
        
        If Not rsTmp.EOF Then
            If rsTmp!诊疗类别 & "" = "H" And rsTmp!操作类型 & "" = "1" And Val(rsTmp!医嘱期效 & "") = 0 Then
                '获取病人护理等级id
                strSQL = "Select b.id as 收费细目ID" & vbNewLine & _
                        "From 诊疗收费关系 A, 收费项目目录 B" & vbNewLine & _
                        "Where a.收费项目id = b.Id And b.类别 = 'H' And Nvl(b.项目特性, 0) <> 0 And a.诊疗项目id = [3] And Rownum = 1 And" & vbNewLine & _
                        "      Not Exists" & vbNewLine & _
                        " (Select 1 From 病案主页 Where 病人id = [1] And 主页id = [2] And 护理等级id = a.收费项目id)"

                Set rs护理等级 = zlDatabase.OpenSQLRecord(strSQL, "Check医嘱停止", Val(rsTmp!病人ID & ""), Val(rsTmp!主页ID & ""), Val(rsTmp!诊疗项目ID & ""))
                
                If Not rs护理等级 Is Nothing Then
                    If Not rs护理等级.EOF Then
                        If Val(rs护理等级!收费细目id & "") <> 0 Then
                             strOut = Check护理等级变更(Val(rsTmp!病人ID & ""), Val(rsTmp!主页ID & ""), CDate(rsTmp!开始执行时间 & ""), lngModel)
                             If strOut <> "" Then
                                MsgBox strOut, vbInformation, gstrSysName
                                Exit Function
                             End If
                        End If
                    End If
                End If
            ElseIf rsTmp!诊疗类别 & "" = "Z" And rsTmp!操作类型 & "" = "2" Then
                '对留观病人下达入院通知;
                '预约登记的条件：1.当前无预约,2.当前是门诊留观病人（在院时也允许，因为需要先预约,入院接收时检查了必须出院后才能接收）
                strSQL = "Select Count(*) as Count From 病案主页 Where 病人id = [1] And Nvl(主页id, 0) = 0"
                Set rsTmp1 = zlDatabase.OpenSQLRecord(strSQL, "Check医嘱停止", Val(rsTmp!病人ID & ""))
                If Val(rsTmp!Count & "") = 0 Then
                    strSQL = "Select Count(*) Into v_Count From 病案主页 Where 病人id = [1] And 主页id = [2] And 病人性质 <> 1"
                    Set rsTmp1 = zlDatabase.OpenSQLRecord(strSQL, "Check医嘱停止", Val(rsTmp!病人ID & ""), Val(rsTmp!主页ID & ""))
                    If Val(rsTmp!Count & "") = 0 Then
                        'Zl_入院病案主页_Insert
                        If Not CallService("Zl_Patisvr_Lockcheck", "{""input"":{""pati_id"":" & Val(rsTmp!病人ID & "") & "}}", , "Check医嘱停止", lngModel, True) Then
                            Exit Function
                        End If
                        
                        '获取住院预约入院的SQL
                        Set colPi = New Collection
                        Set colPati = PatiSvrGetpatiinfo(3, Val(rsTmp!病人ID & ""), lngModel)
                        colPi.Add 0, "_del_page"
                        colPi.Add 0, "_del_in"
                        colPi.Add str操作员姓名, "_operator_name"
                        colPi.Add str操作员编号, "_operator_code"
                        colPi.Add IIF(Val(rsTmp!紧急标志 & "") = 1, "急诊", ""), "_in_type"
                        colPi.Add Format(rsTmp!开始执行时间 & "", "yyyy-MM-dd HH:mm:ss"), "_create_time"
                        colPi.Add rsTmp!开嘱医生 & "", "_out_doctor"
                        colPi.Add Format(rsTmp!开始执行时间 & "", "yyyy-MM-dd HH:mm:ss"), "_in_date"
                        colPi.Add 0, "_rgst_id"
                        colPi.Add 0, "_inpatient_num"
                        colPi.Add 0, "_keep_num"
                        colPi.Add 1, "_status"
                        colPi.Add 0, "_pati_nature"
                        colPi.Add 0, "_again_in"
                        colPi.Add Val(rsTmp!开嘱科室ID & ""), "_pati_deptid"
                        
                        Call GetSQL住院预约登记(colPati, colPi, str预约入院SQL)
                    End If
                End If
            End If
        End If
    End If
    Check医嘱校对 = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSQL住院预约登记(colPati As Collection, colPi As Collection, strAPage As String) As Boolean
'功能:住院发送医生成预约入院登记
'参数:colPati 病人基本信息
'     colPi 入院病案信息
'出参:strAPage 病人案新增[Zl_病案主页_预约入院登记]入参Json格式
    Dim colP As Collection
    On Error GoTo errH
    
    strAPage = ""
                    
    Set colP = colPati(1)

    strAPage = strAPage & "(" & Val(GetColVal(colP, "_pati_id"))
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_del_page"))
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_del_in"))
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_rgst_id"))
    strAPage = strAPage & ",""" & GetColVal(colPi, "_inpatient_num") & """"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_name") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_sex") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_age") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_emp_phno") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_emp_postcode") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_fee_category") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_emp_addr") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_country_name") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pat_hous_addr") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pat_hous_postcode") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_marital_cstatus") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pat_home_addr") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pat_home_postcode") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pat_home_phno") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_contacts_addr") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_contacts_phno") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_contacts_relation") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_contacts_idcard") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_contacts_name") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_area") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_education") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_mdlpay_mode_name") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_ocpt_name") & "'"
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_insurance_type"))
    strAPage = strAPage & ",To_Date(" & GetColVal(colPi, "_in_date") & ", 'yyyy-mm-dd hh24:mi:ss')"
    strAPage = strAPage & ",'" & GetColVal(colPi, "_out_doctor") & "'"
    strAPage = strAPage & ",To_Date(" & GetColVal(colPi, "_create_time") & ", 'yyyy-mm-dd hh24:mi:ss')"
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_pati_nature"))
    strAPage = strAPage & ",'" & GetColVal(colP, "_in_type") & "'"
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_again_in"))
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_status"))
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_keep_num"))
    strAPage = strAPage & ",'" & GetColVal(colPi, "_operator_code") & "'"
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_pati_deptid"))
    strAPage = strAPage & ",'" & GetColVal(colPi, "_operator_name") & "')"
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get医保大类(ByVal int险类 As Integer, ByVal lng收费细目ID As Long, Optional ByRef colBat As Collection, Optional ByVal lngType As Long) As String
'功能：获取指定行的费用类型
'   int险类+lng收费细目ID 两都不为0时获取当前项目的险类名称
'   colBat 即是入参也是出参，入参传入需要批量获取的 收费细目ID逗号拼串[colBat(1)]，传出 收费细目ID,大类名称，二维数据
'   lngType 批量类型，0-批量获取医保大类名称,返回集合，1-批量获取是否设置了支付项目，返回字符串

    Dim strJsonIn As String, str大类 As String
    
    On Error GoTo errH
    
    If lng收费细目ID <> 0 And int险类 <> 0 Then
        strJsonIn = "{""input"":{""insurance_type"":" & int险类 & ",""fee_item_id"":" & lng收费细目ID & "}}"
        Call CallService("Zl_ExseSvr_GetInsureItemInfo", strJsonIn, , , , False, , , , True)
        str大类 = gobjService.GetJsonNodeValue("output.insure_name") & ""
        Get医保大类 = IIF(str大类 <> "", "医保大类:" & str大类, "")
    End If
    
    If Not colBat Is Nothing Then
        If colBat.Count > 0 Then
            strJsonIn = "{""input"":{""insurance_type"":" & int险类 & ",""fee_item_ids"":""" & colBat(1) & """,""query_type"":" & lngType & "}}"
            Call CallService("Zl_ExseSvr_GetInsureItemInfo", strJsonIn, , , , False, , , , True)
            
            If lngType = 0 Then
                Set colBat = gobjService.GetJsonListValue("output.item_list", "fee_item_id")
            Else
                Get医保大类 = gobjService.GetJsonNodeValue("output.fee_item_ids") & ""
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillChareApp(colList As Collection, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str销帐原因 As String, ByVal strTitle As String, ByVal lng模块 As Long) As Boolean
'功能:产生费用销帐申请，住院医嘱回退检查医嘱的时候自动产生药品用费用的销帐申请
'参数:colList 销帐申请列表,colList集合来源:Zl_Exsesvr_CheckOrderRoll中的charge_list[],医嘱回退发送的时候
'     lng病人ID
'     lng主页ID
'     str销帐原因
'--ZL_ExseSvr_billCharge,费用产生销帐申请
'--  input
'--    request_operator                C 1 申请人
'--    request_time                    C 1 申请时间
'--    request_type                    N 1 申请类别
'--    del_tag                         N 1 删除标志
'--    reason                          C 1 销帐原因
'--    item_list[]用于生成销帐申请的列表
'--        fee_id                      N 1 费用ID
'--        request_dept_id             N 1 销帐申请科室ID
'--        fee_item_id                 N 1 收费细目ID
'--        quantity                    N 1 数次
'--        audit_dept_id               N 1 审核部门ID
    Dim lng出院科室ID As Long, lng审核标志 As Long, lng状态 As Long
    Dim lng审核部门ID As Long, i As Long
    Dim strIt As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strJsonIn As String
    Dim colData As Collection
    Dim colRow As Collection
    Dim datCur As Date
    Dim strTime As String
    
    On Error GoTo errH
    
    strIt = ""
    strSQL = "select a.出院科室id,a.审核标志,a.状态 from 病案主页 a where a.病人id=[1] and a.主页id=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, strTitle, lng病人ID, lng主页ID)
    lng出院科室ID = Val("" & rsTmp!出院科室ID)
    lng审核标志 = Val("" & rsTmp!审核标志)
    lng状态 = Val("" & rsTmp!状态)
    For i = 1 To colList.Count '审核部门id/删除标志不传  del_tag / audit_dept_id
        strIt = strIt & ",{""fee_id"":" & colList(i)("_fee_id") & _
            ",""request_dept_id"":" & colList(i)("_request_dept_id") & _
            ",""item_id"":" & colList(i)("_fee_item_id") & ",""request_type"":1" & _
            ",""request_num"":" & GetJsonNum(colList(i)("_quantity") & "") & ",""sended_num"":0" & _
            ",""out_depti_id"":" & lng出院科室ID & ",""audit_sign"":" & lng审核标志 & ",""inp_state"":" & lng状态 & _
            "}"
    Next
    strJsonIn = "{""input"":{""item_list"":[" & Mid(strIt, 2) & "]}}"
    Call CallService("Zl_ExseSvr_CheckBillChargeOff", strJsonIn, , strTitle, lng模块, False, , , , True)
    Set colData = gobjService.GetJsonListValue("output.item_list", "fee_id")
    
    strIt = ""
    For i = 1 To colList.Count '审核部门id/删除标志不传  del_tag / audit_dept_id
        strIt = strIt & ",{""fee_id"":" & colList(i)("_fee_id") & _
            ",""request_dept_id"":" & colList(i)("_request_dept_id") & _
            ",""fee_item_id"":" & colList(i)("_fee_item_id") & _
            ",""quantity"":" & GetJsonNum(colList(i)("_quantity") & "")
        
        
        Set colRow = GetColObj(colData, "_" & colList(i)("_fee_id"))
        lng审核部门ID = GetColVal(colRow, "_audit_dept_id", "N")
        If lng审核部门ID <> 0 Then
            strIt = strIt & ",""audit_dept_id"":" & lng审核部门ID
        End If
        
        strIt = strIt & "}"
    Next
     
    datCur = zlDatabase.Currentdate
    strTime = Format(datCur, "yyyy-MM-dd HH:mm:ss")
    strJsonIn = "{""input"":{""request_operator"":""" & UserInfo.姓名 & """" & _
            ",""request_time"":""" & strTime & """" & _
            ",""request_type"":1" & _
            ",""reason"":""" & str销帐原因 & """" & _
            ",""item_list"":[" & Mid(strIt, 2) & "]" & _
            "}}"
    Call CallService("Zl_ExseSvr_BillCharge", strJsonIn, , strTitle, lng模块, False, , , , True)
    
    BillChareApp = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function BillChareAppPivas(ByVal strPivas As String, ByVal strExse As String, ByVal lng门诊记帐 As Long, ByVal lng配液ID As Long, ByVal lng申请科室ID As Long, ByVal blnAutoAduit As Boolean, ByVal str销帐原因 As String, ByVal strTitle As String, ByVal lng模块 As Long) As Boolean
'功能:护士站输液销帐申请
'参数:colList 销帐申请列表,colList集合来源:Zl_Exsesvr_CheckOrderRoll中的charge_list[],医嘱回退发送的时候
'       strPivas 输液内容列表：
'                strJsonIn = "{""input"":{""query_type"":0,""pivas_id"":" & Val(.RowData(.Row)) & "}}"
'                Call CallService("Zl_PivasSvr_GetPivasContent", strJsonIn, strPivasOutPar, Me.Caption, , False, , , , True)
'       strExse  费用服务出参：
'                strJsonIn = "{""input"":{""oper_type"":1,""fee_source"":" & IIF(lng门诊记帐 = 1, 1, 2) & ",""fee_ids"":""" & str费用ids & """}}"
'                Call CallService("Zl_ExseSvr_CheckBillChargeOff", strJsonIn, strExseOutPar, Me.Caption, , False, , , , True)
'       blnAutoAduit 是否自动审核,true 可以自动审核，false不能自动审核，需要获取已发数量
'       str销帐原因 销帐申请原因

'     lng病人ID
'     lng主页ID
'     str销帐原因
'--ZL_ExseSvr_billCharge,费用产生销帐申请
'--  input
'--    request_operator                C 1 申请人
'--    request_time                    C 1 申请时间
'--    request_type                    N 1 申请类别
'--    del_tag                         N 1 删除标志
'--    reason                          C 1 销帐原因
'--    item_list[]用于生成销帐申请的列表
'--        fee_id                      N 1 费用ID
'--        request_dept_id             N 1 销帐申请科室ID
'--        fee_item_id                 N 1 收费细目ID
'--        quantity                    N 1 数次
'--        audit_dept_id               N 1 审核部门ID
    
    Dim colPivas As Collection
    Dim colFee As Collection
    Dim strFees As String
    Dim lng病人ID As Long, lng主页ID As Long
    Dim strJsonIn As String
    Dim colDrug As Collection
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng出院科室ID As Long, lng审核标志 As Long, lng状态 As Long
    Dim lng审核部门ID As Long, i As Long
    Dim colTmp As New Collection
    Dim colRow As Collection
    Dim str已发数 As String
    Dim colEO As Collection
    Dim colRowOther As Collection
    Dim strIt As String
    Dim strTime As String
    Dim strParExseCharge  As String
    Dim strParPivasCharge  As String
    Dim blnTran As Boolean
    Dim str序号 As String
    Dim strDelDrug As String
    Dim strDelDrugPivas As String
    Dim lng自动审核 As Long ' 1-静配销帐自动审核，2-普通销帐申请模式
    Dim strPatiList As String
    Dim lng医嘱ID As Long, lng发送号 As Long, strSQLAdd As String, strSQLCls As String
    On Error GoTo errH
    
    '先将参数接收过来
    Call InitSvr
    Call gobjService.SetJsonString(strPivas)
    Set colPivas = gobjService.GetJsonListValue("output.item_list", "rcpdtl_id")
    strFees = gobjService.GetJsonNodeValue("output.fee_ids") & ""
    Call gobjService.SetJsonString(strExse)
    Set colFee = gobjService.GetJsonListValue("output.charge_list", "fee_id")
    lng病人ID = colFee(1)("_pati_id")
    lng主页ID = colFee(1)("_pati_pageid")
    
    strSQL = "select a.出院科室id,a.审核标志,a.状态 from 病案主页 a where a.病人id=[1] and a.主页id=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, strTitle, lng病人ID, lng主页ID)
    lng出院科室ID = Val("" & rsTmp!出院科室ID)
    lng审核标志 = Val("" & rsTmp!审核标志)
    lng状态 = Val("" & rsTmp!状态)
    '病人信息列表
    strPatiList = "{""pati_id"":" & lng病人ID & ",""pati_dept_id"":" & lng出院科室ID & ",""fee_audit_status"":" & lng审核标志 & ",""si_inp_status"":" & lng状态 & "}"
    If Not blnAutoAduit Then
        '不能自动审核，需要获取已发数量
        strJsonIn = "{""input"":{""rcpdtl_ids"":""" & strFees & """}}"
        Call CallService("Zl_Drugsvr_Getexecutednum", strJsonIn, , strTitle, lng模块, False, , , , True)
        Set colDrug = gobjService.GetJsonListValue("output.data", "rcpdtl_id")
    End If
    
    For i = 1 To colPivas.Count
        Set colRow = New Collection
 
        colRow.Add lng配液ID & "", "_pivas_id"
        colRow.Add colPivas(i)("_rcp_no") & "", "_fee_no"
        colRow.Add colPivas(i)("_rcpdtl_id") & "", "_fee_id"
        colRow.Add lng申请科室ID & "", "_request_dept_id"
        colRow.Add colPivas(i)("_drug_id") & "", "_item_id"
        colRow.Add IIF(blnAutoAduit, "0", "1"), "_request_type"
        colRow.Add colPivas(i)("_send_num") & "", "_request_num"
        
        If lng医嘱ID = 0 Then
            lng医嘱ID = colPivas(i)("_order_id")
            lng发送号 = colPivas(i)("_send_no")
        End If
        
        str序号 = ""
        If Not blnAutoAduit Then
            Set colRowOther = GetColObj(colDrug, "_" & colPivas(i)("_rcpdtl_id"))
            str已发数 = GetColVal(colRowOther, "_sended_num", "N")
        Else
            str已发数 = 0
            
            Set colRowOther = GetColObj(colFee, "_" & colPivas(i)("_rcpdtl_id"))
            str序号 = GetColVal(colRowOther, "_serial_num", "N")
            str序号 = str序号 & ":" & colPivas(i)("_send_num") & ":0" '可以自动审核,执行状态当成是0
        End If
        
        colRow.Add str已发数, "_sended_num"
        
        colRow.Add str序号, "_serial_num" '费用序号

        colTmp.Add colRow
    Next
    strIt = ""
    For i = 1 To colTmp.Count '销帐列表
        strIt = strIt & ",{"
        strIt = strIt & """fee_id"":" & colTmp(i)("_fee_id")
        strIt = strIt & ",""request_dept_id"":" & colTmp(i)("_request_dept_id")
        strIt = strIt & ",""item_id"":" & colTmp(i)("_item_id")
        strIt = strIt & ",""request_type"":" & colTmp(i)("_request_type")
        strIt = strIt & ",""request_num"":" & GetJsonNum("" & colTmp(i)("_request_num"))
        strIt = strIt & ",""sended_num"":" & GetJsonNum("" & colTmp(i)("_sended_num"))

        strIt = strIt & "}"
    Next
    
    strJsonIn = "{""input"":{""oper_type"":0,""item_list"":[" & Mid(strIt, 2) & "],""pati_list"":[" & strPatiList & "]}}"
    Call CallService("Zl_ExseSvr_CheckBillChargeOff", strJsonIn, , strTitle, lng模块, False, , , , True)
    Set colEO = gobjService.GetJsonListValue("output.item_list", "fee_id")
    
    strIt = ""
    For i = 1 To colTmp.Count
        strIt = strIt & ",{"
        strIt = strIt & """fee_id"":" & colTmp(i)("_fee_id")
        strIt = strIt & ",""request_dept_id"":" & colTmp(i)("_request_dept_id")
        strIt = strIt & ",""fee_item_id"":" & colTmp(i)("_item_id")
        strIt = strIt & ",""request_type"":" & colTmp(i)("_request_type")
        strIt = strIt & ",""quantity"":" & GetJsonNum("" & colTmp(i)("_request_num"))
        Set colRowOther = GetColObj(colEO, "_" & colTmp(i)("_fee_id"))
        lng审核部门ID = GetColVal(colRowOther, "_audit_dept_id", "N")
        If lng审核部门ID <> 0 Then
            strIt = strIt & ",""audit_dept_id"":" & lng审核部门ID
        End If
        
        If blnAutoAduit Then
            strIt = strIt & ",""auto_aduit"":1"
            strIt = strIt & ",""outpati_account"":" & lng门诊记帐
            strIt = strIt & ",""fee_no"":""" & colTmp(i)("_fee_no") & """"
            strIt = strIt & ",""serial_num"":""" & colTmp(i)("_serial_num") & """"
        End If
        strIt = strIt & "}"
        
        '如果是是自动审核还需要删除药品处方
        If blnAutoAduit Then
            strDelDrug = strDelDrug & ",{""rcpdtl_id"":" & colTmp(i)("_fee_id") & _
            ",""chargeoffs_num"":" & GetJsonNum("" & colTmp(i)("_request_num")) & _
            ",""dispensing_ids"":""" & colTmp(i)("_pivas_id") & """" & _
            "}"
        End If
    Next
    
    strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    strJsonIn = "{""input"":{""request_operator"":""" & UserInfo.姓名 & """,""request_code"":""" & UserInfo.编号 & """,""request_time"":""" & strTime & """,""request_type"":" & IIF(blnAutoAduit, "0", "1") & _
        ",""del_tag"":2,""reason"":""" & str销帐原因 & """,""item_list"":[" & Mid(strIt, 2) & "]}}"
    strParExseCharge = strJsonIn 'Zl_Exsesvr_BillChargeOff销帐申请入参拼接完成
    
    strJsonIn = "{""input"":{""pivas_ids"":""" & lng配液ID & """,""operator_status"":9,""operator_name"":""" & UserInfo.姓名 & """,""operator_time"":""" & strTime & """" & _
    ",""auto_aduit"":" & IIF(blnAutoAduit, 1, 0) & ",""operator_notes"":""" & str销帐原因 & """}}"
    strParPivasCharge = strJsonIn '静配销帐申请 Zl_Pivassvr_Statusupdate
    
    If strDelDrug <> "" Then
        strJsonIn = """pivas_list"":[{""pivas_ids"":""" & lng配液ID & """,""operator_name"":""" & UserInfo.姓名 & """,""operator_time"":""" & strTime & """,""reason"":""" & str销帐原因 & """}]"
        strParPivasCharge = ""
        strDelDrugPivas = "{""input"":{""item_list"":[" & Mid(strDelDrug, 2) & "]," & strJsonIn & "}}"
    End If
    lng自动审核 = 2
    If strDelDrugPivas <> "" Then lng自动审核 = 1
    
    strSQLAdd = "Zl_病人医嘱发送_更新同步标志(9,'" & lng发送号 & "','" & lng医嘱ID & "','" & UserInfo.姓名 & "','" & OS.ComputerName & "'," & lng配液ID & ")"
    strSQLCls = "Zl_病人医嘱发送_更新同步标志(10,'" & lng发送号 & "','" & lng医嘱ID & "','" & UserInfo.姓名 & "','" & OS.ComputerName & "'," & lng配液ID & ")"
    
    '如果不是自动审核就可以执行了,流程,先静配销帐,再费用销帐,实际上开启事务没用,先保留便于测试
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(strSQLCls, strTitle)
    Call zlDatabase.ExecuteProcedure(strSQLAdd, strTitle)
    If strDelDrugPivas <> "" Then
        Call CallService("Zl_Drugsvr_Delrecipebill", strDelDrugPivas, , strTitle, lng模块, False, , , , True)
    End If
    If strParPivasCharge <> "" Then
        Call CallService("Zl_Pivassvr_Statusupdate", strParPivasCharge, , strTitle, lng模块, False, , , , True)
    End If
    gcnOracle.CommitTrans: blnTran = False
    
    
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(strSQLCls, strTitle)
    Call CallService("Zl_Exsesvr_BillChargeOff", strParExseCharge, , strTitle, lng模块, False, , , , True)
    gcnOracle.CommitTrans: blnTran = False
    
    BillChareAppPivas = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExseSvrAdviceisexist(str医嘱IDs As String, Optional lngModel As Long) As Collection
'  ---------------------------------------------------------------------------
'  --功能:根据医嘱ID查询是否在费用表存在记录
'  --入参：Json_In:格式
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    
    strJsonIn = "{""input"":{""advice_ids"":""" & str医嘱IDs & """}}"
    If CallService("Zl_Exsesvr_Adviceisexist", strJsonIn, strJsonOut, "ExseSvrAdviceisexist", lngModel, False, , , , True) Then
        Set ExseSvrAdviceisexist = gobjService.GetJsonListValue("output.advice_list")
    End If
End Function


Public Function ExseSvrGetnextno(int序号 As Integer, Optional lng科室ID As Long, Optional lngModel As Long, Optional ByVal lng数量 As Long, Optional ByRef varNos As Variant) As String
'  ---------------------------------------------------------------------------
'  --功能:根据特定规则产生新的号码
'  --入参：Json_In:格式
'  -------------------------------------------
'lng数量 一次批量获取多个单据号,当传入值>1时返回一个数据varNos
    Dim strTmp As String
    Dim strJsonIn As String
    Dim strJsonOut As String
    
    If lng数量 <= 1 Then
        lng数量 = 0
    End If
    
    strJsonIn = "{""input"":{""item_num"":" & int序号 & ",""dept_id"":" & ZVal(lng科室ID) & ",""quantity"":" & lng数量 & "}}"
    If CallService("Zl_Exsesvr_Getnextno", strJsonIn, strJsonOut, "ExseSvrGetnextno", lngModel, False, , , , True) Then
        If lng数量 > 1 Then
            strTmp = gobjService.GetJsonNodeValue("output.next_no")
        Else
            strTmp = gobjService.GetJsonNodeValue("output.next_no")
            ExseSvrGetnextno = strTmp
        End If
        varNos = Split(strTmp, ",")
    End If
End Function

Public Function ExseSvrGetSpecCalcFeeItem(Optional lngModel As Long) As Collection
'  -------------------------------------------
'  --功能：获取打折费别明细列表
'  --入参：Json_In:格式
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String

    strJsonIn = ""
    If CallService("zl_ExseSvr_GetSpecCalcFeeItem", strJsonIn, strJsonOut, "ExseSvrGetSpecCalcFeeItem", lngModel, False, , , , True) Then
        Set ExseSvrGetSpecCalcFeeItem = gobjService.GetJsonListValue("output.feecategory_list")
    End If
End Function

Public Function ItemHaveCash(ByVal int病人来源 As Integer, ByVal bln单独执行 As Boolean, ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, _
    ByVal lng发送号 As Long, ByVal str类别 As String, ByVal str单据号 As String, ByVal int记录性质 As Integer, ByVal int门诊记帐 As Integer, ByVal int方式 As Integer, _
    Optional ByVal blnMove As Boolean, Optional ByRef blnIsAbnormal As Boolean) As Boolean
'功能：判断当前的执行医嘱是否已收费或记帐划价单是否已审核
'参数：int病人来源=1-门诊,2-住院
'      str类别=诊疗类别，用于从一组医嘱中区分分开执行的内容
'      int方式=0-检查是否存在未收费记录
'              1-检查是否存在已收费记录
'      int门诊记帐=1=住院发送到门诊记帐
'      返回：str医嘱IDs=该医嘱及相关的医嘱ID,NOs=医嘱发送的单据号和补的附费中的单据号
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strNoFilter As String
    Dim strOrders As String
    Dim lng主ID As Long
    Dim lngFO As Long
    Dim i As Long
    
    On Error GoTo errH
    
    If int病人来源 = 2 And int记录性质 = 2 And int门诊记帐 = 0 Then
        lngFO = 2
    Else
        lngFO = 1
    End If
    strOrders = lng医嘱ID
    strNoFilter = str单据号
    
    If Not bln单独执行 Then
        lng主ID = IIF(lng相关ID <> 0, lng相关ID, lng医嘱ID)
        strSQL = "Select a.医嘱id, a.No" & vbNewLine & _
            "From 病人医嘱记录 C, 病人医嘱附费 A" & vbNewLine & _
            "Where a.医嘱id In (Select ID From 病人医嘱记录 Where (ID =[1] Or 相关id =[1]) And 诊疗类别 =[2]) And a.发送号 =[3] And" & vbNewLine & _
            "      a.医嘱id = c.Id  and c.诊疗类别=[2] " & vbNewLine & _
            "group by a.医嘱id,a.no"
        If blnMove Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISService", lng主ID, str类别, lng发送号)
        
        For i = 1 To rsTmp.RecordCount
            If InStr("," & strOrders & ",", "," & rsTmp!医嘱ID & ",") = 0 Then
                strOrders = IIF(strOrders = "", "", strOrders & ",") & rsTmp!医嘱ID
            End If
            If InStr("," & strNoFilter & ",", "," & rsTmp!NO & ",") = 0 Then
                strNoFilter = IIF(strNoFilter = "", "", strNoFilter & ",") & rsTmp!NO
            End If
            rsTmp.MoveNext
        Next
    End If
     
    strSQL = "{""input"":{""fee_origin"":" & lngFO & ",""order_ids"":""" & strOrders & """,""fee_nos"":""" & strNoFilter & """,""oper_type"":" & int方式 & "}}"
    Call CallService("Zl_ExseSvr_GetOrderChargedInfo", strSQL, , "mdlCISService", , False, , , , True)
    
    '是否有异常费用
    If gobjService.GetJsonNodeValue("output.blance_sign") & "" = "1" Then
        blnIsAbnormal = True
    Else
        blnIsAbnormal = False
    End If
    
    If gobjService.GetJsonNodeValue("output.isexist") & "" = "1" Then
        ItemHaveCash = True
    End If
  
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetListNodeTxt(ByVal strJsonIn As String, ByVal strListName As String, Optional ByVal blnItem As Boolean) As String
'功能：从指定的服务出参中Json串中获取某个Json数组结点值，带结点名，逗号打头，
'     特殊说明：ZLHIS服务过程出参的标准结果,要求所取的数组为第一层 如:output.item_list
'参数: blnItem false-带结点名,true 只要元素串
'返回:json串中的某个片段,如果数组无元素结,则取元素时返回空串
'出参效果如下:
'        ,"charge_list":[{"fee_id":172780,"fee_item_id":10831},{"fee_id":172781,"fee_item_id":12391}]
'        ,{"fee_id":172780,"fee_item_id":10831},{"fee_id":172781,"fee_item_id":12391}

    Dim strTmp As String
    Dim objJson As New zl9ComLib.clsJson
    Dim strJsonPart As String
    
    On Error GoTo errH
    
    objJson.JSON = strJsonIn
    strJsonPart = objJson.PathItem("output." & strListName).ItemJSON
    
    If blnItem Then
        strJsonPart = Mid(strJsonPart, 2, Len(strJsonPart) - 2)
        If strJsonPart <> "" Then
            strTmp = "," & strJsonPart
        End If
    Else
        strTmp = ",""" & strListName & """:" & strJsonPart
    End If
    
    GetListNodeTxt = strTmp
    Exit Function
errH:
    err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub Exe自动退药退料(ByVal strReStuff As String, ByVal strClsS As String, ByVal strReDrug As String, ByVal strClsD As String, ByVal strTitle As String, ByVal lng模块 As Long)
'功能：自动退药退料
'说明：可能是重试操作，已经退了，所以要忽略错误

    On Error GoTo errH
    
    If strReStuff <> "" Then  '自动退材
        Call CallService("Zl_Stuffsvr_Autoreturnstuff", strReStuff, , strTitle, lng模块, False, , , , True)
        Call CallService("Zl_Exsesvr_Updateexeinfo", strClsS, , strTitle, lng模块, False, , , , True)
    End If
    
    If strReDrug <> "" Then  '自动退药
        Call CallService("Zl_Drugsvr_Autoreturndrug", strReDrug, , strTitle, lng模块, False, , , , True)
        Call CallService("Zl_Exsesvr_Updateexeinfo", strClsD, , strTitle, lng模块, False, , , , True)
    End If
    
    Exit Sub
errH:
    err.Clear
End Sub

Public Function Get医嘱作废执行过程SQL(ByVal strCisOutPar As String, ByVal strTime As String, ByRef strSQL As String) As Boolean
    Dim lng主医嘱ID As Long
    Dim colPage As Collection, col_cis_data_list As Collection
    Dim colCISData As Collection
    Dim lng病人ID As Long
    Dim lng删在院 As Long
    Dim lng删主页 As Long
    
    
    On Error GoTo errH
    
    strSQL = ""
    
'    CREATE OR REPLACE PROCEDURE ZL_病人医嘱记录_作废_S
'(
'  病人id_In      In Number,
'  主医嘱id_In    In Number,
'  删在院_In      In Number,
'  删主页_In      In Number,
'  操作员姓名_In  In Varchar2,
'  操作员编号_In  In Varchar2,
'  操作时间_In    In Varchar2,
'  护理等级医嘱id In Number,
'  回退变动_In    In Number,
'  护理等级停_In  In Number,
'  屏蔽打印_In    In Number,
'  婴儿序号_In    In Number,
'  主页id_In      In Number

    Call GetSvrOutInfo(strCisOutPar)
    Set colPage = gobjService.GetJsonListValue("output.page_list")
    Set col_cis_data_list = gobjService.GetJsonListValue("output.cis_data_list")
    Set colCISData = col_cis_data_list(1)
    If Not colPage Is Nothing Then
        If colPage.Count > 0 Then
            lng删在院 = Val(colPage(1)("_del_in"))
            lng删主页 = 1
        End If
    End If
    lng主医嘱ID = Val(GetColVal(colCISData, "_main_order_id"))
    lng病人ID = Val(gobjService.GetJsonNodeValue("output.order_info_list[0].pati_id") & "")
     
    strSQL = "ZL_病人医嘱记录_作废_S("
    strSQL = strSQL & ZVal(Val(GetColVal(colCISData, "_pati_id")))  '  病人id_In      In Number,
    strSQL = strSQL & "," & lng主医嘱ID '  主医嘱id_In    In Number,
    strSQL = strSQL & "," & ZVal(lng删在院) '  删在院_In      In Number,
    strSQL = strSQL & "," & ZVal(lng删主页) '  删主页_In      In Number,
    strSQL = strSQL & ",'" & UserInfo.姓名 & "'" '  操作员姓名_In  In Varchar2,
    strSQL = strSQL & ",'" & UserInfo.编号 & "'" '  操作员编号_In  In Varchar2,
    strSQL = strSQL & ",'" & strTime & "'" '  操作时间_In    In Varchar2,
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_nurse_order_id"))) '  护理等级医嘱id In Number, v_Jtmp := v_Jtmp || ',"nurse_order_id":1';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_undo_nurse")))   '  回退变动_In    In Number,v_Jtmp := v_Jtmp || ',"undo_nurse":1';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_nurse_stop")))  '  护理等级停_In  In Number, v_Jtmp := v_Jtmp || ',"nurse_stop":1';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_no_print")))   '  屏蔽打印_In    In Number, v_Jtmp := v_Jtmp || ',"no_print":1';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_baby_num")))   '  婴儿序号_In    In Number,v_Jtmp := v_Jtmp || ',"baby_num":' || Nvl(r_Advice.婴儿, 0);
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_pati_pageid")))  '  主页id_In      In Number ',"pati_pageid":' || Nvl(r_Advice.主页id, 0);
    strSQL = strSQL & ",'" & GetColVal(colCISData, "_del_allergy") & "'"     '删过敏_In       In Number := Null,',"del_allergy":' || Nvl(n_删过敏, 0);
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_oper_after_msg"))) '术后消息_In     In Number := Null, ',"oper_after_msg":' || Nvl(n_术后消息, 0);
    strSQL = strSQL & ",'" & GetColVal(colCISData, "_area_msg") & "'"   '病区执行消息_In In Varchar2 := Null, --医嘱id拼串逗号分割 ',"area_msg":"' || v_病区执行ids || '"';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_pacs_msg")))    '检查消息_In
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_lis_msg")))     '检验消息_In
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_blood_msg")))      '输血消息_In     In Number := Null, ',"blood_msg":' || Nvl(n_输血消息, 0);
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_oper_msg"))) '手术消息_In     In Number := Null,',"oper_msg":' || Nvl(n_手术消息, 0);
    strSQL = strSQL & ",'" & GetColVal(colCISData, "_rgst_no") & "'" '挂号单_In       In Varchar2 := Null,',"rgst_no":"' || r_Advice.挂号单 || '"';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_send_no")))     '发送号_In       In Number := Null,',"send_no":"' || v_发送号 || '"';
    strSQL = strSQL & ",'" & GetColVal(colCISData, "_bill_no") & "'"   'No_In           In Varchar2 := Null',"bill_no":"' || v_No || '"';
    strSQL = strSQL & ")"
    
    Get医嘱作废执行过程SQL = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function F_病人医嘱记录_作废(ByVal Id_In As String, ByVal int医生站模块号 As Integer) As Boolean
'功能：门诊病人医嘱作废门诊医嘱前相关检查
'作废门诊医嘱
'
'
'参数:
'  Id_In         In 病人医嘱记录.Id%Type,
'  操作员编号_In In 人员表.编号%Type := Null,
'  操作员姓名_In In 人员表.姓名%Type := Null,
'  护理医嘱id_In In 病人医嘱记录.Id%Type := Null,
'  作废时间_In   In 病人医嘱状态.操作时间%Type := Null
    Dim blnRis As Boolean, blnPacs As Boolean, blnDo As Boolean, col_fee_no_list As Collection
    Dim strCisOutPar As String, bln是预约入院医嘱 As Boolean, strJsonIn As String
    Dim col_order_info_list As Collection, lng挂号科室ID As Long, colSQL As New Collection
    Dim bln一并给药 As Boolean, bln已签名 As Boolean, strSQL As String
    Dim str医嘱内容 As String, strAdvice输血  As String, strRisIds  As String, strLISIDs  As String
    Dim strSign As String, str发送号串 As String, strTime As String
    Dim str_after_order_ids As String '先作废后退药的医嘱行ids
    Dim str_auto_item_ids As String '本科室自动完成的医嘱项目ids,格式：医嘱ID:收细目ID,医嘱ID:收细目ID,医嘱ID:收细目ID,,,,
    Dim lng医嘱ID As Long, str_fee_nos As String, lng_outpati_account As Long
    Dim strIDs As String, lng_bill_prop As Long, str_rcp_nos As String, str_stuff_nos As String, colPar As New Collection
    Dim lng病人ID  As Long, str挂号单 As String, str_all_order_ids As String
    Dim blnMoved As Boolean, colPage As Collection
    Dim strSource As String, str已执行费用IDs As String
    Dim intRule  As Integer, str过程作废电子签名 As String
    Dim lng证书ID As Long, strTimeStamp As String, strTimeStampCode As String, lng挂号ID As Long, lng签名id As Long
    
    On Error GoTo errH
    
    Call HaveRIS
    blnRis = (Not gobjRis Is Nothing) And gbln启用影像信息系统预约
    If blnRis = False Then
        Call CreateObjectPacs
        blnPacs = (Not gobjPACS Is Nothing) And gbln启用PACS系统预约
    End If
    strCisOutPar = zlDatabase.CallProcedure("Zl_病人医嘱记录_作废_Check", "医嘱作废", Val(Id_In), Empty)
    If Not GetSvrOutInfo(strCisOutPar) Then
        Exit Function
    End If
    Set col_order_info_list = gobjService.GetJsonListValue("output.order_info_list")
    Set col_fee_no_list = gobjService.GetJsonListValue("output.fee_no_list")
    bln是预约入院医嘱 = 1 = Val(GetColVal(col_order_info_list(1), "_is_order_appin"))
    bln已签名 = 1 = Val(GetColVal(col_order_info_list(1), "_is_sign"))
    bln一并给药 = 1 = Val(GetColVal(col_order_info_list(1), "_is_merge"))
    str医嘱内容 = GetColVal(col_order_info_list(1), "_advice_note") & ""
    lng医嘱ID = Val(GetColVal(col_order_info_list(1), "_main_order_id"))
    str_all_order_ids = GetColVal(col_order_info_list(1), "_all_order_ids") & ""
    strAdvice输血 = GetColVal(col_order_info_list(1), "_blood_order_ids") & ""
    strRisIds = GetColVal(col_order_info_list(1), "_ris_order_ids") & ""
    strLISIDs = GetColVal(col_order_info_list(1), "_lis_order_ids") & ""
    lng病人ID = Val(GetColVal(col_order_info_list(1), "_pati_id"))
    lng挂号ID = Val(GetColVal(col_order_info_list(1), "_rgst_id"))
    str挂号单 = GetColVal(col_order_info_list(1), "_rgst_no") & ""
    lng挂号科室ID = Val(GetColVal(col_order_info_list(1), "_rgst_deptid"))
    str_after_order_ids = GetColVal(col_order_info_list(1), "_after_order_ids") & ""
    str_auto_item_ids = GetColVal(col_order_info_list(1), "_auto_item_ids") & ""
    str发送号串 = GetColVal(col_order_info_list(1), "_send_nos") & ""
    Set colPage = gobjService.GetJsonListValue("output.page_list")
    lng_bill_prop = Val(GetColVal(col_fee_no_list(1), "_bill_prop"))
    str_rcp_nos = GetColVal(col_fee_no_list(1), "_rcp_nos") & ""
    str_stuff_nos = GetColVal(col_fee_no_list(1), "_stuff_nos") & ""
    str_fee_nos = GetColVal(col_fee_no_list(1), "_fee_nos") & ""
    lng_outpati_account = Val(GetColVal(col_fee_no_list(1), "_outpati_account"))
    
    '电子签名检查和提示
    If bln已签名 Then
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox "作废已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能作废。", vbInformation, gstrSysName
            Else
                MsgBox "作废已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能作废。", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        If gobjESign.CertificateStoped(UserInfo.姓名) = False Then strSign = vbCrLf & vbCrLf & "提示：该医嘱已经签名，作废时你需要再次签名。"
    End If
    
    If bln一并给药 Then
        If MsgBox("该组一并给药的医嘱将会一起作废，确实要作废吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If MsgBox("确实要作废医嘱""" & str医嘱内容 & """吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    '作废时进行电子签名
    If bln已签名 Then
        If gobjESign.CertificateStoped(UserInfo.姓名) = False Then
            '获取签名医嘱源文
            strIDs = lng医嘱ID
            intRule = ReadAdviceSignSource(4, lng病人ID, str挂号单, strIDs, 0, blnMoved, strSource)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "不能读取需要作废的已签名医嘱源文内容。", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode, , lng病人ID, 0, lng挂号ID)
            If strSign <> "" Then
                If strTimeStamp <> "" Then
                    strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    strTimeStamp = "NULL"
                End If
                lng签名id = zlDatabase.GetNextID("医嘱签名记录")
                str过程作废电子签名 = "zl_医嘱签名记录_Insert(" & lng签名id & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & strIDs & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
            Else
                MsgBox "作废医嘱电子签名失败。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '调用作废前外挂接口
    blnDo = Check外挂门诊医嘱作废前(lng病人ID, lng挂号ID, lng医嘱ID)
    If Not blnDo Then Exit Function
    
    strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    blnDo = Check已发药品卫材(str_all_order_ids, str_auto_item_ids, str_after_order_ids, str发送号串, strTime, lng_bill_prop, str_rcp_nos, str_stuff_nos, str已执行费用IDs, colPar)
    If Not blnDo Then Exit Function
    
    If str_fee_nos <> "" Then
        strJsonIn = "{""input"":{""bill_prop"":" & lng_bill_prop & _
        ",""outpati_account"":" & lng_outpati_account & ",""order_ids"":""" & str_all_order_ids & """" & _
         ",""fee_nos"":""" & str_fee_nos & """,""exe_fee_ids"":""" & str已执行费用IDs & """,""after_order_ids"":""" & str_after_order_ids & """}}"
        'colPar内容  "标异常回退处方","回退处方","标异常回退卫材","回退卫材","回退费用"
        If Not Check检查医嘱门诊作废费用(strJsonIn, strTime, colPar) Then
            Exit Function
        End If
    End If
    
    '门诊如果作废是 留观或者预约入院医嘱相关判断
    If Not colPage Is Nothing Then
        If colPage.Count > 0 Then
            blnDo = Get门作废入院医嘱(colPage, colPar)
            If Not blnDo Then Exit Function
        End If
    End If
    
    '本域的两个过程，作废医嘱，签名
    blnDo = Get医嘱作废执行过程SQL(strCisOutPar, strTime, strSQL)
    If Not blnDo Then Exit Function
    colSQL.Add strSQL
    If str过程作废电子签名 <> "" Then
        colSQL.Add str过程作废电子签名
    End If
    
    blnDo = Exe医嘱门诊作废(colPar, colSQL)
    If Not blnDo Then Exit Function
    '输血 LIS RIS  PACS 回退医嘱相关处理
    Call Get门诊医嘱作废其它接口(int医生站模块号, lng医嘱ID, strAdvice输血, strRisIds, strLISIDs, blnRis, blnPacs)
    
    '调用作废后外挂接口
    Call Check外挂门诊医嘱作废后(lng病人ID, lng挂号ID, lng医嘱ID)
    
    '调用预约中心服务
    If bln是预约入院医嘱 Then
        If Svr预约入院按科启用(lng挂号科室ID) Then
            Call Svr预约入院取消服务(lng挂号ID)
        End If
    End If
    
    F_病人医嘱记录_作废 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function F_病人医嘱记录_作废_Bat(ByVal Ids_In As String, ByVal int医生站模块号 As Integer) As Boolean
                           
'功能：门诊医嘱批量作废病人医嘱作废前相关检查
'作废门诊医嘱
'
'
'参数:
'  Id_In         In 病人医嘱记录.Id%Type,
'  操作员编号_In In 人员表.编号%Type := Null,
'  操作员姓名_In In 人员表.姓名%Type := Null,
'  护理医嘱id_In In 病人医嘱记录.Id%Type := Null,
'  作废时间_In   In 病人医嘱状态.操作时间%Type := Null
    Dim blnRis As Boolean, blnPacs As Boolean, blnDo As Boolean, col_fee_no_list As Collection
    Dim strCisOutPar As String, bln是预约入院医嘱 As Boolean, strJsonIn As String
    Dim col_order_info_list As Collection, lng挂号科室ID As Long, colSQL As New Collection
    Dim bln已签名 As Boolean, strSQL As String, str已执行费用IDs As String
    Dim strAdvice输血    As String, strRisIds  As String, strLISIDs  As String
    Dim strSign As String, str发送号串 As String, strTime As String, str_all_order_ids As String
    Dim str_after_order_ids As String '先作废后退药的医嘱行ids
    Dim str_auto_item_ids As String '本科室自动完成的医嘱项目ids,格式：医嘱ID:收细目ID,医嘱ID:收细目ID,医嘱ID:收细目ID,,,,
    Dim lng医嘱ID As Long, str_fee_nos As String, lng_outpati_account As Long
    Dim strIDs As String, lng_bill_prop As Long, str_rcp_nos As String, str_stuff_nos As String, colPar As New Collection
    Dim lng病人ID  As Long, str挂号单 As String, colCisOutPar As New Collection
    Dim blnMoved As Boolean, colPage As Collection, n As Long
    Dim strSource As String, Id_In As String, varBat医嘱ID As Variant
    Dim intRule  As Integer, str过程作废电子签名 As String
    Dim lng证书ID As Long, strTimeStamp As String, strTimeStampCode As String, lng挂号ID As Long, lng签名id As Long
    
    On Error GoTo errH
    varBat医嘱ID = Split(Ids_In, ",")
    
    Call HaveRIS
    blnRis = (Not gobjRis Is Nothing) And gbln启用影像信息系统预约
    If blnRis = False Then
        Call CreateObjectPacs
        blnPacs = (Not gobjPACS Is Nothing) And gbln启用PACS系统预约
    End If
    
    For n = 0 To UBound(varBat医嘱ID)
        Id_In = varBat医嘱ID(n)
        
        strCisOutPar = zlDatabase.CallProcedure("Zl_病人医嘱记录_作废_Check", "医嘱作废", Val(Id_In), Empty)
        If Not GetSvrOutInfo(strCisOutPar) Then
            Exit Function
        End If
        
        colCisOutPar.Add strCisOutPar
        
        Set col_order_info_list = gobjService.GetJsonListValue("output.order_info_list")
        Set col_fee_no_list = gobjService.GetJsonListValue("output.fee_no_list")
        
        If Not bln是预约入院医嘱 Then
            bln是预约入院医嘱 = 1 = Val(GetColVal(col_order_info_list(1), "_is_order_appin"))
        End If
        
        If Not bln已签名 Then
            bln已签名 = 1 = Val(GetColVal(col_order_info_list(1), "_is_sign"))
        End If
        
        strAdvice输血 = MergeStr(strAdvice输血, GetColVal(col_order_info_list(1), "_blood_order_ids") & "")
        strRisIds = MergeStr(strRisIds, GetColVal(col_order_info_list(1), "_ris_order_ids") & "")
        strLISIDs = MergeStr(strLISIDs, GetColVal(col_order_info_list(1), "_lis_order_ids") & "")
        str_all_order_ids = MergeStr(str_all_order_ids, GetColVal(col_order_info_list(1), "_all_order_ids") & "")
        
        If lng病人ID = 0 Then
            '一次发送相同信息
            lng病人ID = Val(GetColVal(col_order_info_list(1), "_pati_id"))
            lng挂号ID = Val(GetColVal(col_order_info_list(1), "_rgst_id"))
            str挂号单 = GetColVal(col_order_info_list(1), "_rgst_no") & ""
            lng挂号科室ID = Val(GetColVal(col_order_info_list(1), "_rgst_deptid"))
            str发送号串 = GetColVal(col_order_info_list(1), "_send_nos") & ""
            lng_outpati_account = Val(GetColVal(col_fee_no_list(1), "_outpati_account"))
            lng_bill_prop = Val(GetColVal(col_fee_no_list(1), "_bill_prop"))
        End If
        
        str_after_order_ids = MergeStr(str_after_order_ids, GetColVal(col_order_info_list(1), "_after_order_ids") & "")
        str_auto_item_ids = MergeStr(str_auto_item_ids, GetColVal(col_order_info_list(1), "_auto_item_ids") & "")
        
        '入院性质的医嘱一批发送中原则上只会有一条
        If colPage Is Nothing Then
            Set colPage = gobjService.GetJsonListValue("output.page_list")
        End If
        
        str_rcp_nos = MergeStr(str_rcp_nos, GetColVal(col_fee_no_list(1), "_rcp_nos") & "")
        str_stuff_nos = MergeStr(str_stuff_nos, GetColVal(col_fee_no_list(1), "_stuff_nos") & "")
        str_fee_nos = MergeStr(str_fee_nos, GetColVal(col_fee_no_list(1), "_fee_nos") & "")
        
    Next
    
    '电子签名检查和提示
    If bln已签名 Then
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox "作废已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能作废。", vbInformation, gstrSysName
            Else
                MsgBox "作废已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能作废。", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        If gobjESign.CertificateStoped(UserInfo.姓名) = False Then strSign = vbCrLf & vbCrLf & "提示：该医嘱已经签名，作废时你需要再次签名。"
    End If
    
    '作废时进行电子签名
    If bln已签名 Then
        If gobjESign.CertificateStoped(UserInfo.姓名) = False Then
            '获取签名医嘱源文
            strIDs = Ids_In
            intRule = ReadAdviceSignSource(4, lng病人ID, str挂号单, strIDs, 0, blnMoved, strSource)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "不能读取需要作废的已签名医嘱源文内容。", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode, , lng病人ID, 0, lng挂号ID)
            If strSign <> "" Then
                If strTimeStamp <> "" Then
                    strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    strTimeStamp = "NULL"
                End If
                lng签名id = zlDatabase.GetNextID("医嘱签名记录")
                str过程作废电子签名 = "zl_医嘱签名记录_Insert(" & lng签名id & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & strIDs & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
            Else
                MsgBox "作废医嘱电子签名失败。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    For n = 0 To UBound(varBat医嘱ID)
        lng医嘱ID = varBat医嘱ID(n)
        '调用作废前外挂接口
        blnDo = Check外挂门诊医嘱作废前(lng病人ID, lng挂号ID, lng医嘱ID)
        If Not blnDo Then Exit Function
    Next
    
    strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    blnDo = Check已发药品卫材(str_all_order_ids, str_auto_item_ids, str_after_order_ids, str发送号串, strTime, lng_bill_prop, str_rcp_nos, str_stuff_nos, str已执行费用IDs, colPar)
    If Not blnDo Then Exit Function
    
    If str_fee_nos <> "" Then
        strJsonIn = "{""input"":{""bill_prop"":" & lng_bill_prop & _
        ",""outpati_account"":" & lng_outpati_account & ",""order_ids"":""" & str_all_order_ids & """" & _
         ",""fee_nos"":""" & str_fee_nos & """,""exe_fee_ids"":""" & str已执行费用IDs & """,""after_order_ids"":""" & str_after_order_ids & """}}"
        'colPar内容  "标异常回退处方","回退处方","标异常回退卫材","回退卫材","回退费用"
        If Not Check检查医嘱门诊作废费用(strJsonIn, strTime, colPar) Then
            Exit Function
        End If
    End If
    
    '门诊如果作废是 留观或者预约入院医嘱相关判断
    If Not colPage Is Nothing Then
        If colPage.Count > 0 Then
            blnDo = Get门作废入院医嘱(colPage, colPar)
            If Not blnDo Then Exit Function
        End If
    End If
    
    
    For n = 1 To colCisOutPar.Count
        strCisOutPar = colCisOutPar(n)
        '本域的两个过程，作废医嘱，签名
        blnDo = Get医嘱作废执行过程SQL(strCisOutPar, strTime, strSQL)
        If Not blnDo Then Exit Function
        colSQL.Add strSQL
    
    Next
    
    If str过程作废电子签名 <> "" Then
        colSQL.Add str过程作废电子签名
    End If
    
    blnDo = Exe医嘱门诊作废(colPar, colSQL)
    If Not blnDo Then Exit Function
    
    '输血 LIS RIS  PACS 回退医嘱相关处理
    Call Get门诊医嘱作废其它接口(int医生站模块号, lng医嘱ID, strAdvice输血, strRisIds, strLISIDs, blnRis, blnPacs)
    
    For n = 0 To UBound(varBat医嘱ID)
        lng医嘱ID = varBat医嘱ID(n)
        '调用作废后外挂接口
        Call Check外挂门诊医嘱作废后(lng病人ID, lng挂号ID, lng医嘱ID)
    Next
    
    '调用预约中心服务
    If bln是预约入院医嘱 Then
        If Svr预约入院按科启用(lng挂号科室ID) Then
            Call Svr预约入院取消服务(lng挂号ID)
        End If
    End If
    
    F_病人医嘱记录_作废_Bat = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Function Check外挂门诊医嘱作废前(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, ByVal lng医嘱ID As Long) As Boolean
    Dim strErr As String
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    Call CreatePlugInOK(p门诊医嘱下达, 0)
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        strErr = ""
        blnDo = gobjPlugIn.AdviceRevokedBefore(glngSys, p门诊医嘱下达, lng病人ID, lng挂号ID, lng医嘱ID, 0, strErr)
        Call zlPlugInErrH(err, "AdviceRevokedBefore")
        If 0 = err.Number Then '接口没有出错的情况下再判断接口的返回值
            If Not blnDo Then
                MsgBox strErr, vbInformation, gstrSysName
                Exit Function
            End If
        End If
        blnDo = False
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
    End If
    Check外挂门诊医嘱作废前 = True
    Exit Function
errH:
    Check外挂门诊医嘱作废前 = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check外挂门诊医嘱作废后(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, ByVal lng医嘱ID As Long) As Boolean
    
    On Error GoTo errH
    
    Call CreatePlugInOK(p门诊医嘱下达, 0)
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.AdviceRevoked(glngSys, p门诊医嘱下达, lng病人ID, lng挂号ID, lng医嘱ID, 0)
        Call zlPlugInErrH(err, "AdviceRevoked")
    End If
    Check外挂门诊医嘱作废后 = True
    Exit Function
errH:
    Check外挂门诊医嘱作废后 = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Svr预约入院取消服务(ByVal lng挂号ID As Long)
'功能：调用预约中心服务取消住院申请(如果当前科室没有启用,预约中心也找不到该数据,也不会报错)
    Dim strJsIn As String
    Dim strJsOut As String
    Dim strErr As String
    Dim blnRet As Boolean
    
    strJsIn = "{""input_in"":{""rgst_id"": """ & lng挂号ID & """}}"
    zlWriteLog "预约中心服务调试日志", "mdlCISKernel", "住院申请取消", LOGLEVEL_Trace, "输入:" & strJsIn
    blnRet = sys.NewSystemSvr("预约中心", "住院申请取消", strJsIn, strJsOut, strErr)
    zlWriteLog "预约中心服务调试日志", "mdlCISKernel", "住院申请取消", LOGLEVEL_Trace, "输出:" & "strJsOut=" & strJsOut & ";函数值=" & blnRet & ";strErr=" & strErr
    If strErr <> "" Then
        MsgBox "预约入院取消服务:" & strErr, vbInformation, gstrSysName
    End If
End Sub

Private Function Get门诊医嘱作废其它接口(ByVal int医生站模块号 As Long, ByVal lng医嘱ID As Long, _
    ByVal strAdvice输血 As String, ByVal strRisIds As String, ByVal strLISIDs As String, _
    ByVal blnRis As Boolean, ByVal blnPacs As Boolean) As Boolean
'功能：门诊医嘱作废其它接口调用
    Dim strErr As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If strAdvice输血 <> "" Then
        If gobjPublicBlood.AdviceOperation(int医生站模块号, lng医嘱ID, 4, False, strErr) = False Then
            MsgBox "血库系统接口调用失败：" & strErr, vbInformation, gstrSysName
        End If
    End If
    
    If strLISIDs <> "" Then
        Call InitObjLis(int医生站模块号)
        '调用LIS作废申请单
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(CStr(lng医嘱ID), strErr) = False Then
                MsgBox "删除检验申请失败：" & strErr, vbInformation, gstrSysName
            End If
        End If
    End If
    
    'RIS回退，回退失败则退出
    If strRisIds <> "" Then '检查、手术、治疗
        If HaveRIS(True) Then
            On Error Resume Next
            If gobjRis.HISRollAdvice(lng医嘱ID) <> 1 Then
                MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISRollAdvice)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
            End If
            err.Clear: On Error GoTo 0
            On Error GoTo errH
        End If
        
        '删除预约信息处理
        If blnRis Then
            Set rsTmp = GetDataRIS预约(lng医嘱ID & "")
            If Not rsTmp.EOF Then
                On Error Resume Next
                If 0 = gobjRis.HISSchedulingEx(Val(rsTmp!ID & ""), Val(rsTmp!预约id & "")) Then
                    MsgBox "当前启用了影像信息系统接口，本次操作删除或修改了已经预约医嘱，但由于影像信息系统接口(HISSchedulingEx)取消息预约未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
                End If
                err.Clear: On Error GoTo errH
            End If
        ElseIf blnPacs Then
            Set rsTmp = GetDataPACS预约(lng医嘱ID & "")
            If Not rsTmp.EOF Then
                On Error Resume Next
                If False = gobjPACS.CancelSchedule(Val(rsTmp!ID & "")) Then
                    MsgBox "当前启用了PACS信息系统接口，本次操作删除或修改了已经预约医嘱，但由于PACS信息系统接口(CancelSchedule)取消息预约未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
                End If
                err.Clear: On Error GoTo errH
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check已发药品卫材(ByVal str_all_order_ids As String, ByVal strAutoitems As String, ByVal strOrderIds先废 As String, ByVal strSendNos As String, ByVal strTime As String, _
    ByVal lng_billtype As Long, ByVal str_rcp_nos As String, ByVal str_stuff_nos As String, ByRef str已执行费用IDs As String, colPar As Collection) As Boolean
'功能：检查已经发药品卫材，门诊病人作废医嘱
'参数：strAutoitems 自完成的项目，需要自动退，格式：医嘱ID:收费细目ID,,,,
'      strOrderIds先废 先做废后退药的医嘱行ids，逗号分割
    Dim i As Long
    Dim strJsonIn As String
    Dim colDrugExe As Collection
    Dim colStuffExe As Collection
    Dim lng医嘱ID As Long
    Dim lng收费细目ID As Long
    Dim strDelDrug As String
    Dim strDelStuff As String
    Dim strRetDrug As String
    Dim strRetStuff As String
    Dim strDelDrugJson As String
    Dim strDelStuffJson As String
    
    Dim str异常标记 As String
    Dim strSQL As String
    Dim strMsg As String
    Dim blnDo先废 As Boolean
    Dim str已执行药品卫材明细IDs As String '当是先作废后退药的时候，药品卫材已执行这部份费用不能退要留着，格式：费用id逗号拼串
    Dim bln先作废后退药模式 As Boolean
    Dim str已收费费用IDs As String
    
    On Error GoTo errH
     
    
    Check已发药品卫材 = True
    bln先作废后退药模式 = strOrderIds先废 <> ""
    
    If str_rcp_nos <> "" Then
        strJsonIn = "{""input"":{""billtype"":" & lng_billtype & _
            ",""rcp_nos"":""" & str_rcp_nos & """,""order_ids"":""" & str_all_order_ids & """}}"
        Call CallService("zl_DrugSvr_GetExecutedNum", strJsonIn, , , , False, , , , True)
        Set colDrugExe = gobjService.GetJsonListValue("output.data")
    End If
    
    If str_stuff_nos <> "" Then
        strJsonIn = "{""input"":{""billtype"":" & lng_billtype & ",""stuff_nos"":""" & str_stuff_nos & """,""order_ids"":""" & str_all_order_ids & """}}"
        Call CallService("zl_StuffSvr_GetExecutedNum", strJsonIn, , , , False, , , , True)
        Set colStuffExe = gobjService.GetJsonListValue("output.item_list")
    End If
    
    If bln先作废后退药模式 Then
        If Not Get已收费费用明细(lng_billtype, colDrugExe, colStuffExe, str已收费费用IDs) Then
            Check已发药品卫材 = False
            Exit Function
        End If
    End If
    
    If Not colDrugExe Is Nothing Then
        If colDrugExe.Count > 0 Then
            For i = 1 To colDrugExe.Count
                blnDo先废 = False
                lng医嘱ID = Val(colDrugExe(i)("_order_id") & "")
                lng收费细目ID = colDrugExe(i)("_drug_id")
                
                If Val(colDrugExe(i)("_sended_num") & "") <> 0 Then
                    '先作废后退判断处理
                    blnDo先废 = InStr("," & strOrderIds先废 & ",", "," & lng医嘱ID & ",") > 0
                    If blnDo先废 Then
                        '如果先作废后退药且已经执行此时这部份费用不能进行处理
                        str已执行药品卫材明细IDs = str已执行药品卫材明细IDs & "," & colDrugExe(i)("_rcpdtl_id")
                    Else
                        If InStr("," & strAutoitems & ",", "," & lng医嘱ID & ":" & lng收费细目ID & ",") = 0 Then
    
                            strMsg = "医嘱发送的费用单据""" & colDrugExe(i)("_rcp_no") & """中的内容已经被部分或完全执行，不能作废。"
                            Screen.MousePointer = 0
                            MsgBox strMsg, vbInformation, gstrSysName
                            Screen.MousePointer = 11
                            Check已发药品卫材 = False
                            Exit Function
                        End If
    
                        '自动退药+删除单据
                        'Zl_Drugsvr_Delrecipebill
                        strRetDrug = strRetDrug & "," & colDrugExe(i)("_rcpdtl_id") & ":" & Val(colDrugExe(i)("_sended_num") & "")
                    End If
                End If
                
                If Not blnDo先废 Then
                    If str已收费费用IDs = "" Or InStr("," & str已收费费用IDs & ",", "," & colDrugExe(i)("_rcpdtl_id") & ",") = 0 Then
                        If InStr("," & str异常标记 & ",", "," & lng医嘱ID & ",") = 0 Then
                            str异常标记 = str异常标记 & "," & lng医嘱ID
                        End If
                        
                        strDelDrug = strDelDrug & ",{""rcpdtl_id"":" & colDrugExe(i)("_rcpdtl_id")
                        strDelDrug = strDelDrug & ",""chargeoffs_num"":" & Val(colDrugExe(i)("_sended_num") & "")
                        strDelDrug = strDelDrug & "}"
                    End If
                End If
            Next
        End If
    End If
    
    If str异常标记 <> "" Then
        strSQL = "Zl_病人医嘱发送_更新同步标志(3,'" & strSendNos & "','" & Mid(str异常标记, 2) & "','" & UserInfo.姓名 & "','" & OS.ComputerName & "')"
        colPar.Add strSQL, "标异常回退处方"
    End If
    str异常标记 = ""
    If Not colStuffExe Is Nothing Then
        If colStuffExe.Count > 0 Then
            For i = 1 To colStuffExe.Count
                blnDo先废 = False
                lng医嘱ID = Val(colStuffExe(i)("_order_id") & "")
                lng收费细目ID = Val(colStuffExe(i)("_stuff_id") & "")
                
                If Val(colStuffExe(i)("_sended_num") & "") <> 0 Then
                    
                    blnDo先废 = InStr("," & strOrderIds先废 & ",", "," & lng医嘱ID & ",") > 0
                    
                    If InStr("," & strAutoitems & ",", "," & lng医嘱ID & ":" & lng收费细目ID & ",") = 0 Then
                    
                        strMsg = "医嘱发送的费用单据""" & colStuffExe(i)("_stuff_no") & """中的内容已经被部分或完全执行，不能作废。"
                        Screen.MousePointer = 0
                        MsgBox strMsg, vbInformation, gstrSysName
                        Screen.MousePointer = 11
                        
                        Check已发药品卫材 = False
                        Exit Function
                    End If
                    '自动退料+删除单据
                    'zl_stuffsvr_autoreturnstuff
                    'Zl_Stuffsvr_Delbill
                    strRetStuff = strRetStuff & "," & colStuffExe(i)("_stuffdtl_id") & ":" & Val(colStuffExe(i)("_sended_num") & "")
                    '如果自动退了料那就可以删卫材单据，对于卫材来说，先作废后退药是没有任何影响的
                    blnDo先废 = False
                End If
                
                '说明:如果是先作废后退药参数模式,如果是已经执行发料的卫材应该只进行退料不删除单据
                If Not blnDo先废 Then
                    If str已收费费用IDs = "" Or InStr("," & str已收费费用IDs & ",", "," & colStuffExe(i)("_stuffdtl_id") & ",") = 0 Then
                    
                        If InStr("," & str异常标记 & ",", "," & lng医嘱ID & ",") = 0 Then
                            str异常标记 = str异常标记 & "," & lng医嘱ID
                        End If
                        
                        strDelStuff = strDelStuff & ",{""stuffdtl_id"":" & colStuffExe(i)("_stuffdtl_id")
                        strDelStuff = strDelStuff & ",""return_num"":" & Val(colStuffExe(i)("_sended_num") & "")
                        strDelStuff = strDelStuff & "}"
                    End If
                End If
                
            Next
        End If
    End If
    
    If str异常标记 <> "" Then
        strSQL = "Zl_病人医嘱发送_更新同步标志(4,'" & strSendNos & "','" & Mid(str异常标记, 2) & "','" & UserInfo.姓名 & "','" & OS.ComputerName & "')"
        colPar.Add strSQL, "标异常回退卫材"
    End If
    
    '回退处方,Zl_病人医嘱发送_更新同步标志
    If strDelDrug <> "" Then
        strDelDrugJson = """item_list"":[" & Mid(strDelDrug, 2) & "]"
        If strRetDrug <> "" Then
            strDelDrugJson = strDelDrugJson & ",""return_list"":[{""audit_operator"":""" & zlStr.ToJsonStr(UserInfo.姓名) & """" & _
                ",""operator_time"":""" & strTime & """" & _
                ",""rcpdtl_ids"":""" & Mid(strRetDrug, 2) & """}]"
        End If
        strDelDrugJson = "{""input"":{" & strDelDrugJson & "}}"
        colPar.Add strDelDrugJson, "回退处方"
    End If
    
    '回退卫材,Zl_病人医嘱发送_更新同步标志,回退卫材的时候有可能有退料没得删院卫材单据
    If strDelStuff <> "" Or strRetStuff <> "" Then
        strDelStuffJson = ""
        If strDelStuff <> "" Then strDelStuffJson = ",""item_list"":[" & Mid(strDelStuff, 2) & "]"
        
        If strRetStuff <> "" Then
            strDelStuffJson = strDelStuffJson & ",""return_list"":[{""audit_operator"":""" & zlStr.ToJsonStr(UserInfo.姓名) & """" & _
            ",""operator_time"":""" & strTime & """" & _
            ",""stuffdtl_ids"":""" & Mid(strRetStuff, 2) & """}]"
        End If
        strDelStuffJson = "{""input"":{" & Mid(strDelStuffJson, 2) & "}}"
        colPar.Add strDelStuffJson, "回退卫材"
    End If
    str已执行费用IDs = Mid(str已执行药品卫材明细IDs, 2)
    Exit Function
errH:
    Check已发药品卫材 = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check检查医嘱门诊作废费用(ByVal strParIn As String, ByVal strTime As String, colPar As Collection) As Boolean
'功能:门诊医嘱作废费用域相关检查判断
    Dim strJsonIn As String
    Dim strPar As String
    Dim strExseOut As String
    Dim lng_exist_balance  As Long
    Dim lng_exist_verify  As Long
    Dim blnHaveMsg As Boolean
    
    On Error GoTo errH
    
 
    Call CallService("Zl_Exsesvr_CheckOrderRevoke", strParIn, strExseOut, , , False, , , , True)
    lng_exist_balance = Val(gobjService.GetJsonNodeValue("output.exist_balance") & "")
    lng_exist_verify = Val(gobjService.GetJsonNodeValue("output.exist_verify") & "")
 
    '检查作废医嘱对应的费用结帐情况
    If lng_exist_balance = 1 Then
        If gbytBillOpt = 1 Then
            If MsgBox("要作废医嘱的对应费用中存在已结帐的费用，确实要作废吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
            blnHaveMsg = True
        ElseIf gbytBillOpt = 2 Then
            MsgBox "要作废医嘱的对应费用中存在已结帐的费用，不能作废。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '已审核记帐费用检查
    If InStr(GetInsidePrivs(p门诊医嘱下达), "作废已审核记帐医嘱") = 0 Then
        If lng_exist_verify = 1 Then
            MsgBox "要作废医嘱的对应记帐划价费用已经审核，不能作废。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '因为前面界面有交互提示，防止并发再重新调用一次
    If blnHaveMsg Then
        Call CallService("Zl_Exsesvr_CheckOrderRevoke", strParIn, strExseOut, , , False, , , , True)
    End If
    
    strPar = GetListNodeTxt(strExseOut, "del_list")
    If strPar <> "" Then
        strJsonIn = "{""input"":{""operator_name"":""" & UserInfo.姓名 & """,""operator_code"":""" & UserInfo.编号 & """,""operator_time"":""" & strTime & """" & strPar & "}}"
        colPar.Add strJsonIn, "回退费用"
    End If
    
    Check检查医嘱门诊作废费用 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Function Exe医嘱门诊作废(colPar As Collection, colSQL As Collection) As Boolean
'功能:回退医嘱发送执行过程
'参数:colPar   "标异常回退处方","回退处方","标异常回退卫材","回退卫材","回退费用",colSQL"回退医嘱"
    Dim blnTran As Boolean
    Dim strSQL As String
    Dim strSvrName As String
    Dim strSvrPar As String
    Dim i As Long
    Dim lngTestMode As Long '测试代码用于调试,0-正常,1-调试模式
    
    On Error GoTo errH
    
    lngTestMode = 1
'    If lngTestMode = 1 Then gcnOracle.BeginTrans: blnTran = True
    
    '药品，因为有静配，可以不用标记异常，因为有可能没得收发记录了，还要需进一步处理静配数据
    strSQL = GetColVal(colPar, "标异常回退处方")
    strSvrPar = GetColVal(colPar, "回退处方")
    strSvrName = "Zl_DrugSvr_DelRecipeBill"
    
    If strSvrPar <> "" Then
        gcnOracle.BeginTrans: blnTran = True
        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "门诊医嘱作废")
        Call CallService(strSvrName, strSvrPar, , "门诊医嘱作废", , False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    
    '卫材 先作废后退药模式是  回退卫材的时候有可能有退料,没得删除卫材单据
    strSQL = GetColVal(colPar, "标异常回退卫材")
    strSvrPar = GetColVal(colPar, "回退卫材")
    strSvrName = "Zl_StuffSvr_DelBill"
    If strSvrPar <> "" Then
        gcnOracle.BeginTrans: blnTran = True
        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "门诊医嘱作废")
        Call CallService(strSvrName, strSvrPar, , "门诊医嘱作废", , False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    
    
    '医嘱
    'strSQL = GetColVal(colSQL, "回退医嘱")
    strSvrPar = GetColVal(colPar, "回退费用")
    strSvrName = "Zl_ExseSvr_DelBill"
    If colSQL.Count > 0 Then
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To colSQL.Count
            strSQL = colSQL(i)
'            Debug.Print strSQL
            Call zlDatabase.ExecuteProcedure(strSQL, "门诊医嘱作废")
        Next
        If strSvrPar <> "" Then
            Call CallService(strSvrName, strSvrPar, , "门诊医嘱作废", , False, , , , True)
        End If
         gcnOracle.CommitTrans: blnTran = False
    End If
    
    
    strSvrPar = GetColVal(colPar, "病人域信息更新")
    strSvrName = "Zl_Patisvr_UpdateInpatiState"
    If strSvrPar <> "" Then
        Call CallService(strSvrName, strSvrPar, , "门诊医嘱作废", , True)
    End If
    
    'If lngTestMode = 1 Then gcnOracle.RollbackTrans: blnTran = False
    'gcnOracle.CommitTrans: blnTran = False
    Exe医嘱门诊作废 = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get门作废入院医嘱(colPage As Collection, colPar As Collection) As Boolean
'功能：门诊作废特殊医嘱
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim colPati As Collection
    Dim colPd As Collection
    Dim colP As Collection
    Dim lng病人ID As Long, lngD在院 As Long, lngD主页 As Long, strUpPait As String, strTitle As String, lng模块 As Long
    
    On Error GoTo errH
    
    If Not colPage Is Nothing Then
        If colPage.Count > 0 Then
            lng病人ID = colPage(1)("_pati_id")
            If Not CallService("Zl_Patisvr_Lockcheck", "{""input"":{""pati_id"":" & lng病人ID & "}}", , strTitle, lng模块, True) Then
                Exit Function
            End If
            lngD在院 = colPage(1)("_del_in")
            lngD主页 = 1
            strJsonIn = "{""input"":{""pati_id"":" & lng病人ID & ",""query_type"":3}}"
            Call CallService("Zl_Patisvr_Getpatiinfo", strJsonIn, strJsonOut, strTitle, lng模块, True)
            Set colPati = gobjService.GetJsonListValue("output.pati_list")
            Set colP = colPati(1)
            Set colPd = colPage(1)
            strUpPait = strUpPait & "," & GetJsonStrNode("pati_id", "_pati_id", "N", colP, 1)
            strUpPait = strUpPait & "," & GetJsonStrNode("pati_pageid", "_pati_pageid", "N", colPd, 1)
            strUpPait = strUpPait & "," & GetJsonStrNode("inpatient_num", "_inpatient_num", "C", colPd)
            strUpPait = strUpPait & "," & GetJsonStrNode("in_time", "_in_time", "C", colPd)
            strUpPait = strUpPait & "," & GetJsonStrNode("adtd_time", "_adtd_time", "C", colPd)
            strUpPait = strUpPait & ",""pati_deptid"":null"
            strUpPait = strUpPait & ",""wardarea_id"":null"
            strUpPait = strUpPait & ",""pati_bed"":"""""
            strUpPait = strUpPait & ",""inp_status"":null"
            If Val("" & colPd("_pati_pageid")) = 0 Then
                strUpPait = strUpPait & ",""inp_times"":null"
            Else
                strUpPait = strUpPait & "," & GetJsonStrNode("inp_times", "_inp_times", "N", colP, 1)
            End If
            strUpPait = "{""input"":{""pati_list"":[{" & Mid(strUpPait, 2) & "}]}}"
            
        End If
    End If
    
    If strUpPait <> "" Then
        colPar.Add strUpPait, "病人域信息更新"
    End If
    
    Get门作废入院医嘱 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function F医嘱执行完成(colProSQL As Collection, ByVal strCisPar As String, ByVal strTime As String, ByVal lng科室ID As Long, ByVal strTitle As String, ByVal lng模块 As Long, ByRef colPar As Collection) As Boolean
'功能：医嘱执行完成相关检查
'参数：strCisPar 医嘱执行完成相关检查，strTime 操作时间，lng科室ID 执行科室ID
'      colProSQL 可执行过程
    Dim strJsonIn As String, i As Long
    Dim strExseOutPar As String, lng发送号 As Long
    Dim col_fee_list As Collection, strJsonF As String, strSQL As String
    Dim lng_is_affirm As Long, lng_fee_origin As Long, lng_is_verify As Long
    Dim lng病区ID As Long '用于判断本科自动发料发药
    Dim str自动发料IDs As String
    Dim str自动发药IDs As String
    Dim str费用ids As String
    Dim col药品收费确认 As New Collection
    Dim col卫材收费确认 As New Collection
    Dim col费用收费确认 As New Collection
    Dim col_order_tag_list As Collection, colPati As Collection, cl标记加 As New Collection, cl标记减 As New Collection
    Dim strPatiInfo As String, blnDo As Boolean, lng病人ID As Long, str标记 As String
    Dim strItem药品卫材审核 As String
    Dim str费用ids药品 As String
    Dim str费用ids卫材 As String
    
    On Error GoTo errH
    
    Call GetSvrOutInfo(strCisPar)
    Set col_order_tag_list = gobjService.GetJsonListValue("output.order_tag_list")
    lng发送号 = Val(gobjService.GetJsonNodeValue("output.send_no") & "")
    lng病区ID = Val(gobjService.GetJsonNodeValue("output.wardarea_id") & "")
    lng病人ID = Val(gobjService.GetJsonNodeValue("output.pati_id") & "")
    strJsonIn = "{""input"":{""is_finish"":1"
    strJsonIn = strJsonIn & ",""fee_nos"":""" & gobjService.GetJsonNodeValue("output.fee_nos") & """"
    strJsonIn = strJsonIn & ",""fee_order_ids"":""" & gobjService.GetJsonNodeValue("output.fee_order_ids") & """"
    strJsonIn = strJsonIn & ",""exe_deptid"":" & lng科室ID
    strJsonIn = strJsonIn & "}}"
    
    If Not CallService("Zl_ExseSvr_GetOrderFeeExeInfo", strJsonIn, strExseOutPar, strTitle, lng模块, True) Then
        Exit Function
    End If
    
    lng_is_affirm = Val(gobjService.GetJsonNodeValue("output.is_affirm") & "")
    
    '分两种情况，一种是要审核费用，一种是不需要审核费用
    Set col_fee_list = gobjService.GetJsonListValue("output.fee_list")
    If Not col_fee_list Is Nothing Then
        If col_fee_list.Count > 0 Then
            For i = 1 To col_fee_list.Count
                lng_is_verify = Val(col_fee_list(i)("_is_verify") & "")
                If InStr(",5,6,7,", "," & col_fee_list(i)("_fee_type") & ",") > 0 Then
                    blnDo = F医嘱执行异常判断(col_order_tag_list, col_fee_list(i))
                    If lng科室ID <> 0 And Not blnDo Then
                        If lng病区ID = Val(col_fee_list(i)("_exe_dept_id") & "") Then
                            If Val(col_fee_list(i)("_rec_state") & "") = 1 Or lng_is_verify = 1 Then
                                If Val(col_fee_list(i)("_fee_origin") & "") = 2 And Val(col_fee_list(i)("_bill_prop") & "") = 2 Then
                                    str自动发药IDs = str自动发药IDs & "," & col_fee_list(i)("_fee_id")
                                    str费用ids药品 = str费用ids药品 & "," & col_fee_list(i)("_fee_id")
                                End If
                            End If
                        End If
                        If lng_is_verify = 1 Then
                            col药品收费确认.Add col_fee_list(i)
                            col费用收费确认.Add col_fee_list(i)
                        End If
                    End If
                ElseIf Val(col_fee_list(i)("_stuff_used") & "") = 1 Then
                    blnDo = F医嘱执行异常判断(col_order_tag_list, col_fee_list(i))
                    If lng科室ID <> 0 And Not blnDo Then
                        '收费单和记帐单才可以发料
                        If Val(col_fee_list(i)("_rec_state") & "") = 1 Or lng_is_verify = 1 Then
                            If Val(col_fee_list(i)("_fee_origin") & "") = 1 And (Val(col_fee_list(i)("_bill_prop") & "") = 1 Or Val(col_fee_list(i)("_bill_prop") & "") = 11) Or Val(col_fee_list(i)("_bill_prop") & "") = 2 Then
                                str自动发料IDs = str自动发料IDs & "," & col_fee_list(i)("_fee_id")
                                str费用ids卫材 = str费用ids卫材 & "," & col_fee_list(i)("_fee_id")
                            End If
                        End If
                        
                        If lng_is_verify = 1 Then
                            col卫材收费确认.Add col_fee_list(i)
                            col费用收费确认.Add col_fee_list(i)
                        End If
                    End If
                Else
                    str费用ids = str费用ids & "," & col_fee_list(i)("_fee_id")
                    If lng_is_verify = 1 Then
                        col费用收费确认.Add col_fee_list(i)
                    End If
                End If
                
                lng_fee_origin = Val(col_fee_list(i)("_fee_origin") & "")
            Next
            
            
            
            If str费用ids药品 <> "" Then
                strJsonIn = "{""input"":{""fee_ids"":""" & Mid(str费用ids药品, 2) & """,""oper_type"":1,""exe_people"":""" & UserInfo.姓名 & """,""exe_time"":""" & strTime & """,""fee_origin"":" & lng_fee_origin & ",""exe_status"":1}}"
                colPar.Add strJsonIn, "费用执行药品"
            End If
            
            If str费用ids卫材 <> "" Then
                strJsonIn = "{""input"":{""fee_ids"":""" & Mid(str费用ids卫材, 2) & """,""oper_type"":1,""exe_people"":""" & UserInfo.姓名 & """,""exe_time"":""" & strTime & """,""fee_origin"":" & lng_fee_origin & ",""exe_status"":1}}"
                colPar.Add strJsonIn, "费用执行卫材"
            End If
            
            
            If str费用ids <> "" Then
                strJsonIn = "{""input"":{""fee_ids"":""" & Mid(str费用ids, 2) & """,""oper_type"":1,""exe_people"":""" & UserInfo.姓名 & """,""exe_time"":""" & strTime & """,""fee_origin"":" & lng_fee_origin & ",""exe_status"":1}}"
                colPar.Add strJsonIn, "费用执行"
            End If
           
        End If
    End If
     
    '如果有收费确认先执行收费确认，主要是指记帐划价的执行完成审核
    If col药品收费确认.Count > 0 Or col卫材收费确认.Count > 0 Then
        '--获取病人基本信息
        strJsonIn = "{""input"":{""pati_id"":" & lng病人ID & ",""query_type"":1}}"
        Call CallService("Zl_Patisvr_Getpatiinfo", strJsonIn, , strTitle, lng模块, False, , , , True)
        Set colPati = gobjService.GetJsonListValue("output.pati_list")
        Set colPati = colPati(1)
        
        strPatiInfo = """pati_id"":" & lng病人ID & _
            "," & GetJsonStrNode("pati_name", "_pati_name", "C", colPati) & _
            "," & GetJsonStrNode("pati_sex", "_pati_sex", "C", colPati) & _
            "," & GetJsonStrNode("pati_age", "_pati_age", "C", colPati) & _
            "," & GetJsonStrNode("pati_outpno", "_outpatient_num", "C", colPati) & _
            ",""auditor"":""" & UserInfo.姓名 & """,""auditor_code"":""" & UserInfo.编号 & """,""audit_time"":""" & strTime & """"
        
        If col药品收费确认.Count > 0 Then
            strItem药品卫材审核 = ""
            
            For i = 1 To col药品收费确认.Count
                str标记 = "Zl_病人医嘱发送_更新同步标志(5,'" & lng发送号 & "','" & col药品收费确认(i)("_order_id") & "','" & UserInfo.姓名 & "','" & OS.ComputerName & "')"
                cl标记加.Add str标记
                
                str标记 = "Zl_病人医嘱发送_更新同步标志(7,'" & lng发送号 & "','" & col药品收费确认(i)("_order_id") & "')"
                cl标记减.Add str标记
                
                strItem药品卫材审核 = strItem药品卫材审核 & ",{" & GetJsonStrNode("billtype", "_bill_prop", "N", col药品收费确认(i))
                strItem药品卫材审核 = strItem药品卫材审核 & "," & GetJsonStrNode("rcp_no", "_fee_no", "C", col药品收费确认(i))
                strItem药品卫材审核 = strItem药品卫材审核 & "," & GetJsonStrNode("rcpdtl_ids", "_fee_id", "C", col药品收费确认(i))
                
                If i = col药品收费确认.Count Then
                    If str自动发药IDs <> "" Then
                        strItem药品卫材审核 = strItem药品卫材审核 & ",""drug_auto_send"":1"
                        strItem药品卫材审核 = strItem药品卫材审核 & ",""auto_send_ids"":""" & Mid(str自动发药IDs, 2) & """"
                        str自动发药IDs = ""
                    End If
                End If
                
                strItem药品卫材审核 = strItem药品卫材审核 & "}"
                
            Next
            
            strJsonIn = "{""input"":{" & strPatiInfo & ",""item_list"":[" & Mid(strItem药品卫材审核, 2) & "]}}"
              
            colPar.Add cl标记减, "药品清异常"
            colPar.Add strJsonIn, "药品收费"
            Set cl标记减 = New Collection
            
        End If
        
        If col卫材收费确认.Count > 0 Then
            strItem药品卫材审核 = ""
            For i = 1 To col卫材收费确认.Count
                str标记 = "Zl_病人医嘱发送_更新同步标志(6,'" & lng发送号 & "','" & col卫材收费确认(i)("_order_id") & "','" & UserInfo.姓名 & "','" & OS.ComputerName & "')"
                cl标记加.Add str标记
                
                str标记 = "Zl_病人医嘱发送_更新同步标志(8,'" & lng发送号 & "','" & col卫材收费确认(i)("_order_id") & "')"
                cl标记减.Add str标记
                
                strItem药品卫材审核 = strItem药品卫材审核 & ",{" & GetJsonStrNode("billtype", "_bill_prop", "N", col卫材收费确认(i))
                strItem药品卫材审核 = strItem药品卫材审核 & "," & GetJsonStrNode("stuff_no", "_fee_no", "C", col卫材收费确认(i))
                strItem药品卫材审核 = strItem药品卫材审核 & "," & GetJsonStrNode("stuffdtl_ids", "_fee_id", "C", col卫材收费确认(i))
                
                If i = col卫材收费确认.Count Then
                    If str自动发料IDs <> "" Then
                        strItem药品卫材审核 = strItem药品卫材审核 & ",""stuff_auto_send"":1"
                        strItem药品卫材审核 = strItem药品卫材审核 & ",""auto_send_ids"":""" & Mid(str自动发料IDs, 2) & """"
                        str自动发料IDs = ""
                    End If
                End If
                
                strItem药品卫材审核 = strItem药品卫材审核 & "}"
                
            Next
            
            strJsonIn = "{""input"":{" & strPatiInfo & ",""item_list"":[" & Mid(strItem药品卫材审核, 2) & "]}}"
            colPar.Add cl标记减, "卫材清异常"
            colPar.Add strJsonIn, "卫材收费"
        End If
    End If
    
    '费用审核
    If col费用收费确认.Count > 0 Then
        strJsonF = ""
        For i = 1 To col费用收费确认.Count
            strJsonF = strJsonF & ",{ " & GetJsonStrNode("fee_source", "_fee_origin", "N", col费用收费确认(i))
            strJsonF = strJsonF & "," & GetJsonStrNode("fee_no", "_fee_no", "C", col费用收费确认(i))
            strJsonF = strJsonF & "," & GetJsonStrNode("serial_nums", "_serial_num", "C", col费用收费确认(i))
            strJsonF = strJsonF & ",""pati_id"":" & lng病人ID
            strJsonF = strJsonF & "}"
        Next
        strJsonIn = "{""input"":{""operator_name"":""" & UserInfo.姓名 & """,""operator_code"":""" & UserInfo.编号 & """,""operator_time"":""" & strTime & """,""item_list"":[" & Mid(strJsonF, 2) & "]}}"
        colPar.Add cl标记加, "医嘱加异常"
        colPar.Add strJsonIn, "费用收费"
    End If
    
    If str自动发药IDs <> "" Then
        strJsonIn = "{""input"":{""rcpdtl_ids"":""" & Mid(str自动发药IDs, 2) & """,""send_type"":1,""operator_name"":""" & UserInfo.姓名 & """,""operator_code"":""" & UserInfo.编号 & """}}"
        colPar.Add strJsonIn, "药品发药"
    End If
    
    If str自动发料IDs <> "" Then
        strJsonIn = "{""input"":{""stuffdtl_ids"":""" & Mid(str自动发料IDs, 2) & """,""send_type"":1,""operator_name"":""" & UserInfo.姓名 & """,""operator_code"":""" & UserInfo.编号 & """}}"
        colPar.Add strJsonIn, "卫材发料"
    End If
    
    '改造 医嘱执行完成过程
    Set colPati = New Collection
    For i = 1 To colProSQL.Count
        If InStr("ZLHIS_" & UCase(colProSQL(i) & ""), UCase("ZLHIS_Zl_病人医嘱执行_Finish_S(")) > 0 Then
            strSQL = ""
            If Not Get医嘱执行完成SQL(strCisPar, strTime, strSQL) Then
                Exit Function
            End If
'            colProSQL(i) = strSQL
        Else
            strSQL = colProSQL(i)
        End If
        colPati.Add strSQL
    Next
    
    colPar.Add colPati, "医嘱执行完成"
    
    
    F医嘱执行完成 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function F医嘱执行异常判断(col_order_tag As Collection, col_item As Collection) As Boolean
'功能:医嘱执行时判断是否是异常的药品卫材
'返回：true 是否异常，false 正常
    Dim i As Long
    
    On Error GoTo errH
     
    If Not col_order_tag Is Nothing Then
        For i = 1 To col_order_tag.Count
            If col_order_tag(i)("_drug_tag") <> 0 And col_order_tag(i)("_order_id") = col_item("_order_id") And col_item("_stuff_used") = 0 Or _
                col_order_tag(i)("_stuff_tag") <> 0 And col_order_tag(i)("_order_id") = col_item("_order_id") And col_item("_stuff_used") = 1 Then
                  
                F医嘱执行异常判断 = True
                Exit Function
            End If
        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get医嘱执行完成SQL(ByVal strCisPar As String, ByVal strTime As String, ByRef strSQL As String) As Boolean
    On Error GoTo errH
    Call GetSvrOutInfo(strCisPar)
'    Set col_order_tag_list = gobjService.GetJsonListValue("output.order_tag_list")
'    lng发送号 = Val(gobjService.GetJsonNodeValue("output.send_no") & "")
'    lng病区ID = Val(gobjService.GetJsonNodeValue("output.wardarea_id") & "")
    
'Create Or Replace Procedure Zl_病人医嘱执行_Finish_s
'(
'完成医嘱ids_In     Varchar2,
'采集完成医嘱ids_In Varchar2,
'医嘱id_In          Number,
'发送号_In          Number,
'完成时间_In        Date,
'操作员姓名_In      Varchar2,
'病人id_In          Number := Null,
'主页id_In          Number := Null,
'挂号单_In          Varchar2 := Null,
'手术开始时间_In    Varchar2 := Null,
'手术执行科室id_In Number:=Null

    strSQL = "Zl_病人医嘱执行_Finish_S("
    strSQL = strSQL & "'" & gobjService.GetJsonNodeValue("output.finish_order_ids") & "'" '完成医嘱ids_In     Varchar2,
    strSQL = strSQL & ",'" & gobjService.GetJsonNodeValue("output.lis_order_ids") & "'" '采集完成医嘱ids_In Varchar2,
    strSQL = strSQL & "," & gobjService.GetJsonNodeValue("output.order_id") '医嘱id_In          Number,
    strSQL = strSQL & "," & gobjService.GetJsonNodeValue("output.send_no") '发送号_In          Number,
    strSQL = strSQL & ",to_date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')" '完成时间_In        Date,
    strSQL = strSQL & ",'" & UserInfo.姓名 & "'"  '操作员姓名_In      Varchar2,
    strSQL = strSQL & "," & gobjService.GetJsonNodeValue("output.pati_id") '病人id_In          Number := Null,
    strSQL = strSQL & "," & ZVal(Val("" & gobjService.GetJsonNodeValue("output.pati_pageid"))) '主页id_In          Number := Null,
    strSQL = strSQL & ",'" & gobjService.GetJsonNodeValue("output.rgst_no") & "'" '挂号单_In          Varchar2 := Null,
    strSQL = strSQL & ",'" & gobjService.GetJsonNodeValue("output.operate_time") & "'" '手术开始时间_In    Varchar2 := Null,
    strSQL = strSQL & "," & ZVal(Val("" & gobjService.GetJsonNodeValue("output.operate_deptid"))) '手术执行科室id_In Number:=Null
    strSQL = strSQL & ")"
    
    Get医嘱执行完成SQL = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Check住院医嘱作废(ByVal strPars As String, ByRef colOut As Collection, ByVal strTitle As String, ByVal lng模块 As Long) As Boolean
'功能：住院医嘱作废检查,同时返回医嘱作废的入参
'入参：strPars 要作废的医嘱信息，医嘱ID:护理等级医嘱ID,医嘱ID:护理等级医嘱ID,....
'出参：colOut对应的入参集合
    Dim i As Long
    Dim varTmp As Variant
    Dim strJsonIn As String
    Dim strTime As String
    Dim strCHK作废 As String
    
    On Error GoTo errH
    Set colOut = New Collection
    varTmp = Split(strPars, ",")
    strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    For i = 0 To UBound(varTmp)
        strCHK作废 = zlDatabase.CallProcedure("Zl_病人医嘱记录_作废_Check", strTitle & "_" & lng模块, Val("" & varTmp(i)), Empty)
        If Not GetSvrOutInfo(strCHK作废) Then
            Exit Function
        End If
        If Not Get医嘱作废执行过程SQL(strCHK作废, strTime, strJsonIn) Then
            Exit Function
        End If
        colOut.Add strJsonIn
        strJsonIn = ""
    Next
    
    Check住院医嘱作废 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get已收费费用明细(ByVal lng_billtype As Long, colDrugExe As Collection, colStuffExe As Collection, ByRef str已收费费用IDs As String) As Boolean
'功能：先作废后退药模式时判断是否已收费
    Dim i As Long
    Dim strJsonIn As String
    Dim strFeeids As String
    Dim strFid As String
    On Error GoTo errH
    
    If Not colDrugExe Is Nothing Then
        If colDrugExe.Count > 0 Then
            For i = 1 To colDrugExe.Count
                strFid = colDrugExe(i)("_rcpdtl_id") & ""
                If InStr("," & strFeeids & ",", "," & strFid & ",") = 0 Then
                    strFeeids = strFeeids & "," & strFid
                End If
            Next
        End If
    End If
    
    If Not colStuffExe Is Nothing Then
        If colStuffExe.Count > 0 Then
            For i = 1 To colStuffExe.Count
                strFid = colStuffExe(i)("_stuffdtl_id") & ""
                If InStr("," & strFeeids & ",", "," & strFid & ",") = 0 Then
                    strFeeids = strFeeids & "," & strFid
                End If
            Next
        End If
    End If
    
    If strFeeids <> "" Then
        strJsonIn = "{""input"":{""fee_origin"":" & lng_billtype & ",""fee_ids"":""" & Mid(strFeeids, 2) & """,""query_type"":7}}"
        Call CallService("Zl_Exsesvr_Getorderfeeinfo", strJsonIn, , , , False, , , , True)
        str已收费费用IDs = gobjService.GetJsonNodeValue("output.fee_ids") & ""
    End If
    
    Get已收费费用明细 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function Exe药品收费确认(colPar As Collection, ByVal strTitle As String, ByVal lng模块 As Long) As Boolean
'功能：医嘱执行完成，记帐划价单自动审核费用时，对药品费用进行收费确认
'说明：
    Dim strJsonIn As String, clTag As Collection, i As Long, strSvrName As String, blnTran As Boolean
    On Error GoTo errH
    strJsonIn = GetColVal(colPar, "药品收费")
    Set clTag = GetColObj(colPar, "药品清异常")
    strSvrName = "Zl_DrugSvr_RecipeAffirm"
    If strJsonIn <> "" Then
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To clTag.Count
            Call zlDatabase.ExecuteProcedure(clTag(i) & "", strTitle)
        Next
        Call CallService(strSvrName, strJsonIn, , strTitle, lng模块, False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    strJsonIn = GetColVal(colPar, "费用执行药品")
    If strJsonIn <> "" Then
'        Call CallService("Zl_ExseSvr_BillInforUpdate", strJsonIn, , strTitle, lng模块, False, , , , True)
    End If
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Exe卫材收费确认(colPar As Collection, ByVal strTitle As String, ByVal lng模块 As Long) As Boolean
'功能：医嘱执行完成，记帐划价单自动审核费用时，对卫材费用进行收费确认
'说明：
    Dim strJsonIn As String, clTag As Collection, i As Long, strSvrName As String, blnTran As Boolean
    On Error GoTo errH
    strJsonIn = GetColVal(colPar, "卫材收费")
    Set clTag = GetColObj(colPar, "卫材清异常")
    strSvrName = "Zl_StuffSvr_BillAffirm"
    If strJsonIn <> "" Then
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To clTag.Count
            Call zlDatabase.ExecuteProcedure(clTag(i) & "", strTitle)
        Next
        Call CallService(strSvrName, strJsonIn, , strTitle, lng模块, False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    strJsonIn = GetColVal(colPar, "费用执行卫材")
    If strJsonIn <> "" Then
'        Call CallService("Zl_ExseSvr_BillInforUpdate", strJsonIn, , strTitle, lng模块, False, , , , True)
    End If
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Upd费用执行状态(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long) As Boolean
'功能：费用执行状态统一更新,纯调服务的方式
    Dim rsExc As ADODB.Recordset '医嘱执行计价
    Dim rsSend As ADODB.Recordset
    Dim lng门诊记帐 As Long, lng记录性质 As Long, lng主页ID As Long
    Dim colStu As Collection, colRcp As Collection, strJsonIn As String
    Dim strHead As String, strTime As String, strItem As String, strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select a.医嘱id,a.相关id,a.诊疗类别,a.NO,a.收费细目id,a.门诊记帐,a.记录性质,Sum(a.已执行数) As 已执行数 from (" & vbNewLine & _
        "Select b.医嘱id,c.相关id,c.诊疗类别,b.No,a.收费细目id,Decode(a.执行状态,1,1,Decode(a.执行状态,1,1,0))*a.数量 As 已执行数,b.门诊记帐,b.记录性质" & vbNewLine & _
        "From 医嘱执行计价 A,病人医嘱发送 B,病人医嘱记录 C" & vbNewLine & _
        "Where b.医嘱id=c.id and a.数量<>0 and a.医嘱id =[1] And a.发送号 =[2] And a.医嘱id = b.医嘱id And a.发送号 = b.发送号) a" & vbNewLine & _
        "Group By a.医嘱id, a.NO, a.收费细目id,a.门诊记帐,a.记录性质,a.相关id,a.诊疗类别"

    Set rsExc = zlDatabase.OpenSQLRecord(strSQL, "Upd费用执行状态", lng医嘱ID, lng发送号)
    If Not rsExc.EOF Then
        lng门诊记帐 = Val(rsExc!门诊记帐 & "")
        lng记录性质 = Val(rsExc!记录性质 & "")
'        lng主页ID = Val(rsExc!主页ID & "")
'        If Val(rsExc!相关ID & "") <> 0 And InStr(",5,6,7,", "," & rsExc!诊疗类别 & ",") > 0 Then
            strSQL = "select a.医嘱id,a.NO,a.门诊记帐,a.记录性质 from 病人医嘱发送 a, 病人医嘱记录 b where a.发送号+0=[2] and a.医嘱id=b.id and b.相关id=[1]"
            Set rsSend = zlDatabase.OpenSQLRecord(strSQL, "Upd费用执行状态", lng医嘱ID, lng发送号)
'        End If
    End If
    
    Call Get执行状态更单据(rsExc, rsSend, colRcp, colStu)
    
    '住院没有收费帐单这一说
    If lng门诊记帐 = 1 Then
        strHead = strHead & ",""fee_origin"":1"
    ElseIf lng记录性质 = 2 Then
        strHead = strHead & ",""fee_origin"":2"
    Else
        strHead = strHead & ",""fee_origin"":1"
    End If
    strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    strHead = strHead & ",""operator_name"":""" & UserInfo.姓名 & """"
    strHead = strHead & ",""operator_code"":""" & UserInfo.编号 & """"
    strHead = strHead & ",""operator_time"":""" & strTime & """"
    
    strItem = Get状态明细(rsExc, colRcp, colStu, strTime)
    
    If strItem <> "" Then
        strJsonIn = "{""input"":{" & Mid(strHead, 2) & ",""item_list"":" & strItem & "}}"
        Call In更新费用执行状态(strJsonIn)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function In更新费用执行状态(ByVal strJsonIn As String)
    On Error GoTo errH
    Call CallService("Zl_Exsesvr_UpdateExeInfo", strJsonIn, , "In更新费用执行状态", , False, , , , True)
    Exit Function
errH:
    err.Clear
End Function

Private Function Get状态明细(rsA As ADODB.Recordset, colStu As Collection, colRcp As Collection, ByVal strTime As String) As String
'功能：获取状态明细情况
    Dim i As Long
    Dim strAll As String
    Dim strItem As String, strList As String
    
    On Error GoTo errH
    
    If Not colStu Is Nothing Then
        If colStu.Count > 0 Then
        For i = 1 To colStu.Count
            strItem = "{""fee_no"":""" & colStu(i)("_stuff_no") & """"
            strItem = strItem & ",""advice_id"":" & colStu(i)("_order_id")
            strItem = strItem & ",""fee_item_id"":" & colStu(i)("_stuff_id"): strAll = strAll & "," & strItem '缓存一次，于用后面去重
            strItem = strItem & ",""exe_nums"":" & GetJsonNum("" & colStu(i)("_sended_num"))
            strItem = strItem & ",""exe_people"":""" & IIF(Val("" & colStu(i)("_sended_num")) = 0, "", UserInfo.姓名) & """"
            strItem = strItem & ",""exe_time"":""" & IIF(Val("" & colStu(i)("_sended_num")) = 0, "", strTime) & """"
            strItem = strItem & "}"
            strList = strList & "," & strItem
        Next
        End If
    End If
    
    
    If Not colRcp Is Nothing Then
        If colRcp.Count > 0 Then
        For i = 1 To colRcp.Count
            strItem = "{""fee_no"":""" & colRcp(i)("_rcp_no") & """"
            strItem = strItem & ",""advice_id"":" & colRcp(i)("_order_id")
            strItem = strItem & ",""fee_item_id"":" & colRcp(i)("_drug_id"): strAll = strAll & "," & strItem '缓存一次，于用后面去重
            strItem = strItem & ",""exe_nums"":" & GetJsonNum("" & colRcp(i)("_sended_num"))
            strItem = strItem & ",""exe_people"":""" & IIF(Val("" & colRcp(i)("_sended_num")) = 0, "", UserInfo.姓名) & """"
            strItem = strItem & ",""exe_time"":""" & IIF(Val("" & colRcp(i)("_sended_num")) = 0, "", strTime) & """"
            strItem = strItem & "}"
            strList = strList & "," & strItem
        Next
        End If
    End If
    
    If Not rsA.EOF Then
        For i = 1 To rsA.RecordCount
        
            strItem = "{""fee_no"":""" & rsA!NO & """"
            strItem = strItem & ",""advice_id"":" & rsA!医嘱ID
            strItem = strItem & ",""fee_item_id"":" & rsA!收费细目id
            
            If InStr("," & strAll & ",", "," & strItem & ",") = 0 Then
                strItem = strItem & ",""exe_nums"":" & GetJsonNum("" & rsA!已执行数)
                
                strItem = strItem & ",""exe_people"":""" & IIF(Val("" & rsA!已执行数) = 0, "", UserInfo.姓名) & """"
                strItem = strItem & ",""exe_time"":""" & IIF(Val("" & rsA!已执行数) = 0, "", strTime) & """"
                
                strItem = strItem & "}"
                strList = strList & "," & strItem
            End If
            
            rsA.MoveNext
        Next
    End If
    If strList <> "" Then
    Get状态明细 = "[" & Mid(strList, 2) & "]"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get执行状态更单据(rsA As ADODB.Recordset, rsB As ADODB.Recordset, ByRef colStu As Collection, ByRef colRcp As Collection) As Boolean
'功能：获取更新执行状态时对应的药品卫材单据中对应的执行数
    Dim i As Long
    Dim strNos As String, strOrderIds As String
    
    
    On Error GoTo errH
    
    If rsA.EOF Then Exit Function
    
    For i = 1 To rsA.RecordCount
        
        If InStr("," & strNos & ",", "," & rsA!NO & ",") = 0 Then strNos = strNos & "," & rsA!NO
        
        If InStr("," & strOrderIds & ",", "," & rsA!医嘱ID & ",") = 0 Then strOrderIds = strOrderIds & "," & rsA!医嘱ID
    
        
        rsA.MoveNext
    Next
    rsA.MoveFirst
    
    If Not rsB Is Nothing Then
        If Not rsB.EOF Then
            For i = 1 To rsB.RecordCount
                
                If InStr("," & strNos & ",", "," & rsB!NO & ",") = 0 Then strNos = strNos & "," & rsB!NO
        
                If InStr("," & strOrderIds & ",", "," & rsB!医嘱ID & ",") = 0 Then strOrderIds = strNos & "," & rsB!医嘱ID
                
                rsB.MoveNext
            Next
            rsB.MoveFirst
        End If
    End If
    
    
    Call Get卫材执行数(Mid(strNos, 2), Mid(strOrderIds, 2), colStu)
    
    Call Get药品执行数(Mid(strNos, 2), Mid(strOrderIds, 2), colRcp)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get药品执行数(ByVal strN As String, ByVal strOd As String, ByRef colOut As Collection) As Boolean
'功能：根据单据号+医嘱ID串获取药品已执行数据
'参数：strN 单据号逗号拼串，strOd 医嘱ID逗号拼串
    Dim strJson As String
    
    On Error GoTo errH
    
    If strN = "" Then Exit Function
    
    strJson = "{""input"":{""billtype"":3,""rcp_nos"":""" & strN & """,""order_ids"":""" & strOd & """}}"
    Call CallService("Zl_Drugsvr_Getexecutednum", strJson, , "更新费用执行状态", , False, , , , True)
    Set colOut = gobjService.GetJsonListValue("output.data")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get卫材执行数(ByVal strN As String, ByVal strOd As String, ByRef colOut As Collection) As Boolean
'功能：根据单据号+医嘱ID串获取卫材已执行数据
'参数：strN 单据号逗号拼串，strOd 医嘱ID逗号拼串
    Dim strJson As String
    
    On Error GoTo errH
    
    If strN = "" Then Exit Function
    
    strJson = "{""input"":{""billtype"":3,""stuff_nos"":""" & strN & """,""order_ids"":""" & strOd & """}}"
    Call CallService("Zl_Stuffsvr_Getexecutednum", strJson, , "更新费用执行状态", , False, , , , True)
    Set colOut = gobjService.GetJsonListValue("output.item_list")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get发药窗口(ByVal bln记帐 As Boolean, ByVal lng病人ID As Long, ByVal str药房明细 As String, ByRef colOut As Collection)
'功能：获取发药窗口
'参数：bln记帐 是否是记帐单 true 是记帐单，false-收费单
    Dim i As Long
    Dim varTmp As Variant
    Dim strItem As String
    Dim strJson As String
    Dim lng_billtype As Long
    
    On Error GoTo errH
    lng_billtype = IIF(bln记帐, 2, 1)
    varTmp = Split(str药房明细, ",")
    For i = 0 To UBound(varTmp)
        strItem = strItem & ",{""billtype"":" & lng_billtype
        strItem = strItem & ",""pharmacy_id"":" & varTmp(i)
        strItem = strItem & ",""pati_id"":" & lng病人ID
        strItem = strItem & ",""valid_days"":null"
        strItem = strItem & ",""defaultwindow"":null"
        strItem = strItem & "}"
    Next
    strJson = "{""input"":{""item_list"":[" & Mid(strItem, 2) & "]}}"
    
    Call CallService("Zl_Drugsvr_Getsendwindows", strJson, , "Get发药窗口", , False, , , , True)
    Set colOut = gobjService.GetJsonListValue("output.item_list")


    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
