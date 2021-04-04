Attribute VB_Name = "mdlSplitService"
Option Explicit

Public Function zlSplitService_CheckErrData(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strType As String, ByVal strInputById As String, ByVal strInputByNo As String, _
    ByRef colOutlist As Collection, ByRef colOutExpenseList As Collection, ByRef strErrMsg_Out As String) As Boolean
    '发药/退药调用服务：检查费用异常状态，矫正收费、审核状态
    'strInputByid：1.按费用id进行检查，费用id,费用id...
    'strInputByNO: 2.按处方NO进行检查，单据性质:no1,no2|...
    '出参：1.如果update_drug_status=2 then 需要更新处方的已收费/审核状态，再调用费用服务更新对方的记费同步状态
    '      2.如果rcp_no_new<>"" then 需要更新处方的NO，费用ID等
    '      3.如果fee_status=2 then 不能进行发药，退药等操作
    Dim arrInput As Variant, i As Integer
    Dim strJson As String, StrJson_In As String, strJson_List As String
    
    On Error GoTo ErrHandler
    strErrMsg_Out = ""
    'Zl_Exsesvr_Checkerrordata
    '  --功能：根据费用NO或费用ID检查收费记账异常信息
    '  --入参：Json_In:格式
    '  --  input
    '  --      fee_type              C   1 费用类别，'4'-卫材，'5,6,7'-药品
    '  --      rcpdtl_ids            C   1 处方明细ids,多个用逗号分隔
    '  --      bill_list[]                  数组，费用NO信息
    '  --         billtype               N   1 单据类型:1-收费处方;2-记帐处方
    '  --         rcp_nos                C   1 处方Nos,多个用逗号分隔
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --     code                   N   1 应答吗：0-失败；1-成功
    '  --     message                C   1 应答消息：失败时返回具体的错误信息
    '  --     billid_list[]                   按费用ID传入时返回id列表
    '  --        rcpdtl_id           N   1 处方明细id
    '  --        fee_status          N   1 费用状态： 0-划价,1-记帐
    '  --        cancel_status       N   1 作废状态:0-正常状态,1-作废同步标志异常
    '  --        update_status       N   1 记费同步状态:0-正常状态,1-未更新药品/卫材记帐状态
    '  --     billno_list[]                 按NO传入时返回NO列表
    '  --         billtype               N   1 单据类型:1-收费处方;2-记帐处方
    '  --         rcp_no                 C   1 处方no
    '  --         fee_status             N   1 费用状态：针对收费时,0-未收费,1-已收费,2-异常收费;针对记帐时,0-划价,1-记帐
    '  --         cancel_status          N   1 作废状态:0-正常状态,1-作废同步标志异常
    '  --         update_drug_status     N   1 记费同步状态:0-正常状态,2-未更新药品/卫材收费状态
    '  --     expense_list[]               仅药品才有
    '  --         billtype               N   1 (原始)单据类型:1-收费处方;2-记帐处方
    '  --         rcp_no                 C   1 (原始)处方no
    '  --         rcpdtl_id              N   1 (原始)处方明细id
    '  --         rcp_no_new             C   1 新生成的处方NO
    '  --         rcpdtl_id_new          N   1 新生成处方明细id
    '  --         pati_pageid            N   1  主页ID
    If strInputById = "" And strInputByNo = "" Then zlSplitService_CheckErrData = True: Exit Function
    
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_type", strType, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("rcpdtl_ids", strInputById, 0)
    If strInputByNo <> "" Then
        arrInput = Split(strInputByNo, "|")
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("billtype", Split(arrInput(i), ":")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("rcp_nos", Split(arrInput(i), ":")(1), 0)
            strJson_List = IIf(strJson_List = "", "", strJson_List & ",") & "{" & strJson & "}"
        Next
        strJson_List = ",""bill_list"":[" & strJson_List & "]"
    End If
    
    '汇总
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"

    If objServiceCall.CallService("Zl_Exsesvr_CheckErrorData", StrJson_In, , "", lngMode, False, , , , True) = False Then Exit Function
    If strInputById <> "" Then
        Set colOutlist = objServiceCall.GetJsonListValue("output.billid_list", "rcpdtl_id")
    Else
       Set colOutlist = objServiceCall.GetJsonListValue("output.billno_list", "billtype,rcp_no")
    End If
    Set colOutExpenseList = objServiceCall.GetJsonListValue("output.expense_list")
    zlSplitService_CheckErrData = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

Public Function zlSplitService_CisUpdateSyncState(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal intType As Integer, ByVal strInput As String, ByRef strErrMsg_Out As String) As Boolean
    '更新医嘱同步标记录更新
    'intType：1-静配，2-药品 ，3-卫材
    'strInput：医嘱id,发送号|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '服务出参list返回集合
    Dim varList As Variant   '集合元素
    
    On Error GoTo ErrHandle
    strErrMsg_Out = ""
    '服务：Zl_CisSvr_UpdateSyncState
    '  ---------------------------------------------------------------------------
    '  --功能：同步标记录更新
    '  --入参：Json_In:格式
    '  --  input
    '  --      order_list[]
    '  --          order_id          N 1 医嘱id
    '  --          send_no           N 1 发送号
    '  --          sign_type         N 1 设置标记录的类型，
    '  --                                  说明：1-清除静配标记录
    '  --                                        2-清除 生成药品同步标记
    '  --                                        3-清除 生成卫材同步标记
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code          N 1 应答吗：0-失败；1-成功
    '  --    message       C 1 应答消息：失败时返回具体的错误信息
    '  ---------------------------------------------------------------------------
    
    If strInput = "" Then zlSplitService_CisUpdateSyncState = True: Exit Function
    
    If strInput <> "" Then
        arrInput = Split(strInput, "|")
        
        For i = 0 To UBound(arrInput)
            
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("order_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("send_no", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("sign_type", intType, 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
    End If
    StrJson_In = IIf(StrJson_In = "", "", StrJson_In & ",") & """order_list"":[" & strJson_List & "]"
    
    '汇总
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_CisSvr_UpdateSyncState"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, , "", lngMode, False, , , , True) = False Then Exit Function
    
    zlSplitService_CisUpdateSyncState = True
    Exit Function
ErrHandle:
    strErrMsg_Out = err.Description
End Function


Public Function zlSplitService_CheckAdviceaffirm(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOutput As String, ByRef strErrMsg_Out As String) As Boolean
    '返回指定病人ID，主页ID，挂号ID的医嘱发送信息
    'strInput：病人id,主页ID,挂号ID,挂号单号|...
    'strOutPut：医嘱id,发送号|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_List As String
    Dim strService As String
    Dim colOutlist As Collection, colOrderList As Collection
    Dim arrOrder As Variant
    
    On Error GoTo ErrHandle
    strErrMsg_Out = ""
    'Zl_CISSvr_GetAffirmErrorData
    ' ---------------------------------------------------------------------------
    '  --功能：临床医嘱执行完成时自动审核费用，异常后未对药品和卫材进行收费确认，针对此类异常数据获取服务
    '  --入参：Json_In:格式
    '  --  input
    '  --      pati_list[]病人关键信息，用于获取医嘱
    '  --           pati_id                    N 1 病人id
    '  --           pati_pageid                N 1 主页id，住院病人传入，门诊传0
    '  --           rgst_id                    N 1 挂号id，门诊病人传入，住院病人传空
    '  --           rgst_no                    C 1 挂号单号
    '  --出参: Json_Out,格式如下
    '  --   output:
    '  --    code                  N 1 应答吗：0-失败；1-成功
    '  --    message               C 1 应答消息：失败时返回具体的错误信息
    '  --     pati_bill_list[]
    '  --         pati_id                      N 1 病人id
    '  --         pati_pageid                  N 0 主页id，住院病人传入，门诊传0
    '  --         rgst_id                      N 0 挂号id，门诊病人传入，住院病人传空
    '  --         rgst_no                      C 0 挂号单号
    '  --         order_ids                    C 1 有异常的所有医嘱id拼串
    '  --         fee_nos                      C 1 有异常的所有单据号拼串
    '  --         order_list[]医嘱发送信息列表
    '  --             send_no                  N 1 发送号
    '  --             advice_id                N 1 医嘱id
    '  --             fee_no                   C 1 单据号
    '  --             bill_prop                N 1 记录性质
    '  --             outpati_account          N 1 是否门诊记帐 0-不是门诊记帐，1-是门诊记帐
    '  --             pati_source              N 1 病人来源 1-门诊医嘱，2-住院医嘱
   
    
    If strInput <> "" Then
        arrInput = Split(strInput, "|")
        
        For i = 0 To UBound(arrInput)
            StrJson_In = ""
            StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(arrInput(i), ",")(0), 1)
            StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageid", Split(arrInput(i), ",")(1), 1)
            StrJson_In = StrJson_In & "," & GetJsonNodeString("rgst_id", Split(arrInput(i), ",")(2), 1)
            StrJson_In = StrJson_In & "," & GetJsonNodeString("rgst_no", Split(arrInput(i), ",")(3), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & StrJson_In & "}"
            Else
                strJson_List = strJson_List & "," & "{" & StrJson_In & "}"
            End If
        Next
    End If
    StrJson_In = IIf(StrJson_In = "", "", StrJson_In & ",") & """pati_list"":[" & strJson_List & "]"
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_CISSvr_GetAffirmErrorData"

    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False, , , , True) = False Then Exit Function
    
    '数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.pati_bill_list")
    If colOutlist Is Nothing Then Exit Function
    
    '只需要返回内容：医嘱id,发送号|...
    For i = 0 To colOutlist.Count - 1
        '循环取子节点数组
        Set colOrderList = objServiceCall.GetJsonListValue("output.pati_bill_list[" & i & "].order_list")
        
        For Each arrOrder In colOrderList
            strOutput = IIf(strOutput = "", "", strOutput & "|") & arrOrder("_advice_id") & "," & arrOrder("_send_no")
        Next
    Next

    zlSplitService_CheckAdviceaffirm = True
    Exit Function
ErrHandle:
    strErrMsg_Out = err.Description
End Function

Public Function GetJsonNodeString(ByVal strNodeName As String, ByVal strValue As String, Optional ByVal intType As Integer, _
    Optional ByVal blnZeroToNull As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取Json接点串
    '入参:strNodeName-接点名
    '     strValue-值
    '     intType-类型:0-字符;1-数字
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 18:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    
    strJson = Chr(34) & strNodeName & Chr(34)
    
    If intType = 0 Then
        strJson = strJson & ":" & Chr(34) & zlStr.ToJsonStr(strValue) & Chr(34)
    Else
        If strValue = "" Or (blnZeroToNull And Val(strValue) = 0) Then
            strJson = strJson & ":null"
        Else
            strJson = strJson & ":" & IIf(Mid(strValue, 1, 1) = ".", "0", "") & strValue
        End If
    End If
        
    GetJsonNodeString = strJson
End Function

Public Function zlSplitService_CallAccountVerify(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strMain As String, ByVal strInput As String) As Boolean
    '发料后服务：审核划价记账单
    '入参：
    '   strMain：操作时间,操作员姓名,操作员编号
    '   strInput：费用来源|单据号|病人ID|序号s||费用来源|单据号|病人ID|序号s||...
    Dim arrInput As Variant, arrItem As Variant, arrMain As Variant
    Dim i As Integer, StrJson_In As String
    Dim strJson_bill As String, strJson_item As String, strJson_List As String
    
    'Zl_Exsesvr_Billverify
    '  --功能：费用单据审核
    '  --入参：Json_In:格式
    '  --  input
    '  --    operator_time         C 1 操作时间:yyyy-mm-dd hh24:mi:ss
    '  --    operator_name         C 1 操作员姓名
    '  --    operator_code         C 1 操作员编号
    '  --    item_list
    '  --        fee_source        N 1 费用来源:1-门诊;2-住院
    '  --        fee_no            C 1 费用单据号
    '  --        serial_nums       C 0 序号串，不传表示整张单据
    '  --        pharmacy_window   C 0 发药窗口，费用来源为门诊时传入，格式：库房ID1:发药窗口1,库房ID2:发药窗口2,....
    '  --        pati_id           N 0 病人id，费用来源为住院且按病人审核时传入(主要针对记帐表)
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code          N 1 应答吗：0-失败；1-成功
    '  --    message       C 1 应答消息：失败时返回具体的错误信息
    If strInput = "" Then Exit Function
    
    '费用来源|单据号|病人ID|序号s||费用来源|单据号|病人ID|序号s||...
    arrInput = Split(strInput, "||")
    strJson_List = ""
    
    For i = 0 To UBound(arrInput)
        arrItem = Split(arrInput(i), "|")
        
        strJson_item = ""
        strJson_item = strJson_item & "" & GetJsonNodeString("fee_source", arrItem(0), 1)
        strJson_item = strJson_item & "," & GetJsonNodeString("fee_no", arrItem(1), 0)
        strJson_item = strJson_item & "," & GetJsonNodeString("pati_id", arrItem(2), 1)
        strJson_item = strJson_item & "," & GetJsonNodeString("serial_nums", arrItem(3), 0)
        
        strJson_List = IIf(strJson_List = "", "", strJson_List & ",") & "{" & strJson_item & "}"
    Next
    
    arrMain = Split(strMain, ",")
    strJson_bill = ""
    strJson_bill = strJson_bill & "" & GetJsonNodeString("operator_time", arrMain(0), 0)
    strJson_bill = strJson_bill & "," & GetJsonNodeString("operator_name", arrMain(1), 0)
    strJson_bill = strJson_bill & "," & GetJsonNodeString("operator_code", arrMain(2), 0)
    
    StrJson_In = "{""input"":{" & strJson_bill & ",""item_list"":[" & strJson_List & "]" & "}}"
        
    '调用服务
    If objServiceCall.CallService("Zl_Exsesvr_Billverify", StrJson_In, , "", lngMode, False, , , , True) = False Then
        MsgBox "调用“Zl_Exsesvr_Billverify”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    zlSplitService_CallAccountVerify = True
End Function

Public Function zlSplitService_CallCheckDrugById_bak(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, colOutlist As Collection) As Boolean
    '发药/退药调用服务：检查费用异常状态，矫正收费、审核状态
    '按费用id进行检查
    'strInput： 费用id,费用id...
    '出参：1.如果update_stuff_status=2 then 需要更新处方的已收费/审核状态，再调用费用服务更新对方的记费同步状态
    '      2.如果fee_status=2 then 不能进行发药，退药等操作
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '服务出参list返回集合
    Dim varList As Variant   '集合元素
    
    On Error GoTo ErrHandle
    
    '服务：Zl_Exsesvr_In_Checkstuff
    
'  ---------------------------------------------------------------------------
'  --功能：根据费用ID检查病人住院异常费用信息
'  --入参：Json_In:格式
'  --  input
'  --      stuff_ids             C   1 发料明细ids,多个用逗号分隔
'  --出参: Json_Out,格式如下
'  --  output
'  --     code                   N   1 应答吗：0-失败；1-成功
'  --     message                C   1 应答消息：失败时返回具体的错误信息
'  --     bill_list[]
'  --        stuff_id            N   1 发料处方明细id
'  --        fee_status          N   1 费用状态： 0-划价,1-记帐
'  --        cancel_status       N   1 作废状态:0-正常状态,1-作废同步标志异常
'  --        update_stuff_status N   1 记费同步状态:0-正常状态,1-未更新药品/卫材记帐状态
'  ------------------------------------------------------------------------------------------------------------
    
    If strInput = "" Then zlSplitService_CallCheckDrugById_bak = True: Exit Function
    
    strJson_List = ""
    If strInput <> "" Then
        strJson = ""
        strJson = strJson & "" & GetJsonNodeString("stuff_ids", strInput, 0)
    End If
    
    '汇总
    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "Zl_Exsesvr_In_Checkstuff"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_Exsesvr_In_Checkstuff”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallCheckDrugById_bak = False
        Exit Function
    End If
    
    Set colOutlist = objServiceCall.GetJsonListValue("output.bill_list", "stuff_id")
    
    zlSplitService_CallCheckDrugById_bak = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_CallCheckDrugByNo_bak(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, colOutlist As Collection, colOutExpenseList As Collection) As Boolean
    '发药/退药调用服务：检查费用异常状态，矫正收费、审核状态
    '按处方NO进行检查
    'strInput：单据性质:no1,no2|...
    '出参：1.如果update_stuff_status=2 then 需要更新处方的已收费/审核状态，再调用费用服务更新对方的记费同步状态
    '      2.如果rcp_no_new<>"" then 需要更新处方的NO，费用ID等
    '      3.如果fee_status=2 then 不能进行发药，退药等操作
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '服务出参list返回集合
    Dim varList As Variant   '集合元素
    
    On Error GoTo ErrHandle
    
    '服务：Zl_Exsesvr_Out_Checkdrug
    
'  --功能：根据单据类型和NO号检查病人门诊异常费用信息
'  --入参：Json_In:格式
'  --  input
'  --    bill_list[]
'  --         billtype               N   1 单据类型:1-收费处方;2-记帐处方
'  --         rcp_nos                C   1 处方Nos,多个用逗号分隔
'  --出参: Json_Out,格式如下
'  --  output
'  --     code                       N   1 应答吗：0-失败；1-成功
'  --     message                    C   1 应答消息：失败时返回具体的错误信息
'  --     bill_list[]
'  --         billtype               N   1 单据类型:1-收费处方;2-记帐处方
'  --         rcp_no                 C   1 处方no
'  --         fee_status             N   1 费用状态：针对收费时,0-未收费,1-已收费,2-异常收费;针对记帐时,0-划价,1-记帐
'  --         cancel_status          N   1 作废状态:0-正常状态,1-作废同步标志异常
'  --         update_stuff_status     N   1 记费同步状态:0-正常状态,2-未更新药品/卫材收费状态
'  --     expense_list[]
'  --         rcp_no                 C   1 (原始)处方no
'  --         rcpdtl_id              N   1 (原始)处方id
'  --         rcp_no_new             C   1 新生成的处方NO
'  --         rcpdtl_id_new          N   1 新生成处方id
    
    strJson_List = ""
    If strInput <> "" Then
        arrInput = Split(strInput, "|")
        
        For i = 0 To UBound(arrInput)
            
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("billtype", Split(arrInput(i), ":")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("rcp_nos", Split(arrInput(i), ":")(1), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = strJson_List
    
    '汇总
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Out_Checkstuff"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_Exsesvr_Out_Checkstuff”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallCheckDrugByNo_bak = False
        Exit Function
    End If
    
    Set colOutlist = objServiceCall.GetJsonListValue("output.bill_list", "billtype,stuff_no")
    Set colOutExpenseList = objServiceCall.GetJsonListValue("output.expense_list")
    
    zlSplitService_CallCheckDrugByNo_bak = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function zlSplitService_UpdateExseInfo(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strBase As String, ByVal strItemList As String, ByVal strDeptchangeList As String, ByVal strDelNoList As String) As Boolean
    '发药/退药调用服务：更新费用执行部门、执行人、发药窗口及执行状态等信息
    'strBase：基本信息，格式 费用来源,操作员编码,操作员姓名,操作时间
    'strItemList：更新费用ID列表，格式 费用ID,已执行数量(发药数量),执行人,执行时间,发药窗口|...
    '             暂时这几个值都从药品数据中都查询出来传入（或按功能从界面中组织，比如改变发药窗口等）
    'strDeptchangeList：更新执行部门 格式 费用id,原执行部门id,现执行部门id|...
    'strDelNoList：退药自动销账 格式 NO;序号(序号1:数量:执行状态1,序号2:数量2:执行状态2,...);操作状态|...
    Dim arrInput As Variant, strPart As String
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim str执行人 As String, str执行时间 As String
    
    On Error GoTo ErrHandle
    
    'Zl_Exsesvr_Updateexeinfo
'    ---------------------------------------------------------------------------
'  --功能：更新费用执行部门、执行人、发药窗口及执行状态等信息
'  --入参：Json_In:格式
'  --input
'  --  fee_origin            N  1  费用来源(默认=2：1-门诊费用，2-住院费用)
'  --  operator_code         C     操作员编码
'  --  operator_name         C     操作员姓名
'  --  operator_time         C     操作时间
'  --  item_list                   按列表更新执行相关信息，传入列表时同时需要传入fee_origin
'  --    fee_id              C  1  费用id
'  --    exe_nums            N  1  已执行数量:为0表示，未执行
'  --    exe_people          C     执行人:部分执行或完全执行时，需要传入，不传入时，以operator_name为准
'  --    exe_time            D     执行时间:yyyy-mm-dd hh24:mi:ss,:部分执行或完全执行时，需要传入，不传入时，以"create_time"为准
'  --    pharmacy_window     C     发药窗口:药品及卫材有效,无此接点，不会更新发药窗口
'  --  deptchange_List       C  1  执行科室变更信息列表
'  --    fee_id              C  1  费用id
'  --    exe_old_deptid      N     原执行科室ID
'  --    exe_deptid          N  1  执行部门id
'  --  delrcp_list           C     [数组]取消输液时，需要同步销帐
'  --    rcp_no              C  1  处方no
'  --    serial_nums         C  1  格式: 序号1:数量:执行状态1,序号2:数量2:执行状态2,...
'  --    operator_status     N     操作状态：0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
'  --出参: Json_Out,格式如下
'  --output
'  --   code                 C  1  应答码：0-失败；1-成功
'  --   message              C  1  应答消息：失败时返回具体的错误信息
'  ---------------------------------------------------------------------------

    
    '1.基本信息
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_origin", Split(strBase, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_code", Split(strBase, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_name", Split(strBase, ",")(2), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_time", Split(strBase, ",")(3), 0)
        
    '2.更新费用id对应的费用信息
    strJson_List = ""
    If strItemList <> "" Then
        arrInput = Split(strItemList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("fee_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_nums", Split(arrInput(i), ",")(1), 1)
            
            '执行人如果为空则不传入该节点
            If Split(arrInput(i), ",")(2) <> "" Then
                If Split(arrInput(i), ",")(2) = "null" Then
                    '特殊的，如果传入的是null，则传入空串，用于取消配药时清空执行人
                    strJson = strJson & "," & GetJsonNodeString("exe_people", "", 0)
                Else
                    strJson = strJson & "," & GetJsonNodeString("exe_people", Split(arrInput(i), ",")(2), 0)
                End If
            End If
            
            '执行时间如果为空则不传入该节点
            If Split(arrInput(i), ",")(3) <> "" Then
                strJson = strJson & "," & GetJsonNodeString("exe_time", Split(arrInput(i), ",")(3), 0)
            End If
            
            '发药窗口如果为空则不传该节点
            If Split(arrInput(i), ",")(4) <> "" Then
                strJson = strJson & "," & GetJsonNodeString("pharmacy_window", Split(arrInput(i), ",")(4), 0)
            End If
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""item_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '3.单独更新费用id对应的执行部门id
    strJson_List = ""
    If strDeptchangeList <> "" Then
        arrInput = Split(strDeptchangeList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("fee_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_old_deptid", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_deptid", Split(arrInput(i), ",")(2), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""deptchange_List"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '4.单独费用销账
    strJson_List = ""
    If strDelNoList <> "" Then
        arrInput = Split(strDelNoList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), ";")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_nums", Split(arrInput(i), ";")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_status", Split(arrInput(i), ";")(2), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""delrcp_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '汇总
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Updateexeinfo"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_Exsesvr_Updateexeinfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_UpdateExseInfo = False
        Exit Function
    End If

    zlSplitService_UpdateExseInfo = True

    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_WriteOff(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strBase As String, ByVal strAccAudit As String) As Boolean
    '退药同时销帐时调用的其他服务：合并销帐申请审核，重审销帐审核，门诊/住院记录删除，更新费用记录状态，销帐检查等功能
    '合并为一个服务可以有效控制发药/退药事务
    'strBase：基本信息，格式 费用来源,操作员编码,操作员姓名,操作时间
    'strAccAudit：销账列表，格式 费用ID,申请时间,操作类型,申请类别,销账数量,已发数量|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim arrList As Variant, arrPart As Variant
    Dim strPart1 As String, strPart2 As String, strPart3 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String
    Dim n As Integer
    Dim strJson_Part As String, strCheckJson_In As String
    
    On Error GoTo ErrHandle
    
'    'Zl_Exsesvr_Drugwriteoff
'      ---------------------------------------------------------------------------
'  --功能：药品及卫材费用销帐(包含审核通过、重审拒绝，取消拒绝)
'  --入参：Json_In:格式
'  --input
'  --  fee_origin            N  1  费用来源（1-门诊，2-住院）
'  --  operator_code         C  1  操作员编码
'  --  operator_name         C  1  操作员姓名
'  --  operator_time         C     操作时间:yyyy-mm-dd hh24:mi:ss
'  --  rcpdtl_list                 [数组]每个处方明细信息
'  --    rcpdtl_id           N  1  处方明细id(费用id)
'  --    request_time        D  1  申请时间
'  --    oper_type           N  1  操作类型:0-审核通过;1-审核不通过 2-审核拒绝 3-取消拒绝;
'  --    request_type        N  1  申请类别（默认传1）
'  --    quantity            N  1  销帐数量
'  --    sended_num          N  1  已发数量
'
'  --出参: Json_Out,格式如下
'  --output
'  --   code                          C 1 应答码：0-失败；1-成功
'  --   message                       C 1 应答消息：失败时返回具体的错误信息
'  ---------------------------------------------------------------------------
    
    '1.基本信息
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_origin", Split(strBase, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_code", Split(strBase, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_name", Split(strBase, ",")(2), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_time", Split(strBase, ",")(3), 0)
    
    '2.销账信息
    strJson_List = ""
    If strAccAudit <> "" Then
        arrInput = Split(strAccAudit, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcpdtl_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("request_time", Split(arrInput(i), ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("oper_type", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("request_type", Split(arrInput(i), ",")(3), 1)
            strJson = strJson & "," & GetJsonNodeString("quantity", Split(arrInput(i), ",")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("sended_num", Split(arrInput(i), ",")(5), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""rcpdtl_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
     
    '汇总
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Drugwriteoff"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_Exsesvr_Drugwriteoff”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_WriteOff = False
        Exit Function
    End If
    
    zlSplitService_WriteOff = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_WriteOffCheck(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strBase As String, ByVal strAccAudit As String, ByVal strPatiList As String, ByRef strCheckMsg As String) As Boolean
    '退药同时销帐时调用的其他服务：销帐检查
    '合并为一个服务可以有效控制发药/退药事务
    'strBase：基本信息，格式 禁止部分销帐,费用来源
    'strAccAudit：销账列表，格式  费用ID,申请时间,操作类型,申请类别,销账数量,已发数量|...
    'strPatiList：病人列表，格式  病人id,费用审核标志,住院状态,病案编目日期|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim arrList As Variant, arrPart As Variant
    Dim strPart1 As String, strPart2 As String, strPart3 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String
    Dim n As Integer
    Dim strJson_Part As String, strCheckJson_In As String
    Dim colOutlist As New Collection
    
    On Error GoTo ErrHandle
''    'Zl_Exsesvr_Drugwriteoff_Check
'  ---------------------------------------------------------------------------
'  --功能：药品及卫材费用销帐审核检查
'  --入参：Json_In:格式
'  --input      药品费用销帐前检查
'  --  part_ban_writeoffs    N  1  禁止部分销帐:0-允许;1-不允许部分销帐(含整张单据的部分或某笔的部份)
'  --  fee_origin            N  1  费用来源:1-门诊，2-住院
'  --  rcpdtl_list[]               本次销帐列表
'  --    oper_type           N  1  操作类型:0-审核通过 1-审核不通过 2-审核拒绝 3-取消拒绝;
'  --    rcpdtl_id           N  1  处方明细ID(费用ID)
'  --    request_time        D     申请时间
'  --    request_type        N     申请类别：缺省为1
'  --    quantity            N  1  销帐数量：为零或null时,按费用ID申请数量直接销帐
'  --    sended_num          N  1  已发数量
'  --  pati_list[]                 病人信息
'  --    pati_id             N     病人ID,为NULL或0时，表示整张单据
'  --    fee_audit_status    N     费用审核标志:0或空-未审核;1-已审核或开始审核;2-完成审核,结合结帐权限
'  --    si_inp_status       N     住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
'  --    catalog_date        C     病案编目日期：yyyy-mm-dd hh24:mi:ss
'  --出参: Json_Out,格式如下
'  --output
'  --   code                          C 1 应答码：0-失败；1-成功
'  --   message                       C 1 应答消息：失败时返回具体的错误信息
'  --  tip_list[]  C  1  提示列表:主要是可能存在多个提示询问方式，所以用列表,禁止时，返回一条信息
'  --    tip_mode  C  1  控制方式:1-提示询问;2-禁止
'  --    tip_message  C  1  提示信息
'  ---------------------------------------------------------------------------
    
    '1.基本信息
    StrJson_In = StrJson_In & "" & GetJsonNodeString("part_ban_writeoffs", Split(strBase, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("fee_origin", Split(strBase, ",")(1), 1)
    
    '2.销账信息
    strJson_List = ""
    If strAccAudit <> "" Then
        arrInput = Split(strAccAudit, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""            '销账列表：费用ID,申请时间,操作类型,申请类别,销账数量,已发数量|...
            strJson = strJson & "" & GetJsonNodeString("rcpdtl_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("request_time", Split(arrInput(i), ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("oper_type", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("request_type", Split(arrInput(i), ",")(3), 1)
            strJson = strJson & "," & GetJsonNodeString("quantity", Split(arrInput(i), ",")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("sended_num", Split(arrInput(i), ",")(5), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""rcpdtl_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '3.病人信息
    strJson_List = ""
    If strPatiList <> "" Then
        arrInput = Split(strPatiList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""            '病人id,费用审核标志,住院状态,病案编目日期|...
            strJson = strJson & "" & GetJsonNodeString("pati_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_audit_status", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("si_inp_status", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("catalog_date", Split(arrInput(i), ",")(3), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""pati_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
     
    '汇总
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Drugwriteoff_Check"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_Exsesvr_Drugwriteoff_Check”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        '服务调用失败时
        strCheckMsg = objServiceCall.GetJsonNodeValue("output.message")
        zlSplitService_WriteOffCheck = False
        Exit Function
    Else
        '服务调用成功，返回业务检查结果
        Set colOutlist = objServiceCall.GetJsonListValue("output.tip_list")
        If colOutlist.Count > 0 Then
            '控制方式,提示信息
            strCheckMsg = colOutlist(1)(1) & "," & colOutlist(1)(2)
        End If
    End If
    
    zlSplitService_WriteOffCheck = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function zlSplitService_CallUpdateSynchrosign(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal intType As Integer, strInputById As String, ByRef strErrMsg_Out As String) As Boolean
    '更新费用记费同步标记
    '按处方NO进行更新
    'intType：0-更新记费同步标志，1-更新转费同步标志
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '服务出参list返回集合
    Dim varList As Variant   '集合元素
    
    On Error GoTo ErrHandle
    strErrMsg_Out = ""
    If strInputById = "" Then zlSplitService_CallUpdateSynchrosign = True: Exit Function
    '服务：Zl_Exsesvr_Sync_Update
    '  --功能：更新收费同步标志
    '  --入参：Json_In:格式
    '  --  input
    '  --    sign_type           N 1 标志类型：0-记费同步标志,1-转费同步标志
    '  --    detail_ids  C  1  处方明细id串(费用id串),支持多个id，用“,”分隔
    '  --    bill_list[]
    '  --      billtype               N   1 单据类型:1-收费处方;2-记帐处方
    '  --      rcp_no                 C   1 处方No
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --     code                   N   1 应答吗：0-失败；1-成功
    '  --     message                C   1 应答消息：失败时返回具体的错误信息
    
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("sign_type", intType, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("detail_ids", strInputById, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Sync_Update"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False, , , , True) = False Then Exit Function
    
    zlSplitService_CallUpdateSynchrosign = True
    Exit Function
ErrHandle:
    strErrMsg_Out = err.Description
End Function

Public Function zlSplitService_GetAdvice(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strAdvice_ids As String, _
    ByRef colAdvice As Collection) As Boolean
    '取医嘱信息
    'strRcpdtl_ids：医嘱id,医嘱id...
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("advice_ids", strAdvice_ids, 0)
'    strJson_In = strJson_In & "," & GetJsonNodeString("advice_ids", strPatis, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_CisSvr_GetAdviceInfo"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“zl_CisSvr_GetAdviceInfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '返回数组
    Set colAdvice = objServiceCall.GetJsonListValue("output.advice_list")
    
    If colAdvice Is Nothing Then Exit Function
    
    zlSplitService_GetAdvice = True
End Function







Public Function zlSplitService_GetCloseAccount(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal colInput As Collection, _
    ByRef colCloseAccount As Collection) As Boolean
    '取销帐记录
    'colInput：查询条件组合，Json中input各节点作为元素的KEY值，集合某元素为空表示该节点值为空
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    
    On Error GoTo ErrHandle
    
'    input           根据条件查询费用销帐信息
'        audit_dept_id   N       审核部门ID(药房)
'        request_begin_time  D       申请开始时间
'        request_end_time    D       申请结束时间
'        audit_begin_time    D       审核开始时间
'        audit_end_time  D       审核结束时间

'        cancel_status    N   1   状态
'        request_dept_id N       申请部门ID
'        request_operator    C       申请人
'        pati_id  N       病人ID
'        cancel_condition    C       销账条件

'        cancel_check    N       核查（选择参数【销账申请需要核查】时传入，0-未核查 1-已核查）
'        rcpdtl_id   C       处方明细id,[数组]：[1,2,3]
'        request_dept_ids   C     申请部门id串，用于批量查询
'        item_ids           C     收费细目id串,用于批量查询
'    output
'        code    C   1   应答码：0-失败；1-成功
'        message C   1   "应答消息：
'        fee_cancel_list         [数组]满足条件的每个费用销帐记录
'        rcpdtl_id N       处方明细id(费用id)
'        request_type  N       申请类别
'        item_id   N       收费细目id
'        request_dept_id   N       申请部门id
'        request_dept  C       申请部门
'        audit_dept_id N       审核部门id
'        quantity  N       数量
'        request_operator  C       申请人
'        request_time  D       申请时间
'        auditor   C       审核人
'        audit_time    D       审核时间
'        cancel_status N       状态
'        cancel_reason C       销帐原因
'        checker   C       核查人
'        price_retail  N       零售价
'        advice_id N       医嘱id
'        pati_id   N       病人ID
'        pati_name C       病人姓名
'        inpatient_num C       住院号

    '入参
    StrJson_In = ""
    If ExistsColObject(colInput, "audit_dept_id") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("audit_dept_id", colInput("audit_dept_id"), 1)
    If ExistsColObject(colInput, "request_begin_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_begin_time", colInput("request_begin_time"), 0)
    If ExistsColObject(colInput, "request_end_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_end_time", colInput("request_end_time"), 0)
    If ExistsColObject(colInput, "audit_begin_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("audit_begin_time", colInput("audit_begin_time"), 0)
    If ExistsColObject(colInput, "audit_end_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("audit_end_time", colInput("audit_end_time"), 0)
    
    If ExistsColObject(colInput, "cancel_status") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_status", colInput("cancel_status"), 1)
    If ExistsColObject(colInput, "request_dept_id") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_dept_id", colInput("request_dept_id"), 1)
    If ExistsColObject(colInput, "request_operator") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_operator", colInput("request_operator"), 0)
    If ExistsColObject(colInput, "pati_id") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_id", colInput("pati_id"), 1)
    If ExistsColObject(colInput, "cancel_condition") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_condition", colInput("cancel_condition"), 0)
    
    If ExistsColObject(colInput, "cancel_check") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_check", colInput("cancel_check"), 1)
    If ExistsColObject(colInput, "rcpdtl_id") Then StrJson_In = StrJson_In & "," & """rcpdtl_id"":" & "[" & colInput("rcpdtl_id") & "]"
    If ExistsColObject(colInput, "request_dept_ids") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_dept_ids", colInput("request_dept_ids"), 0)
    If ExistsColObject(colInput, "item_ids") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("item_ids", colInput("item_ids"), 0)
    
    If StrJson_In = "" Then Exit Function
    
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_ExseSvr_GetChargeOffInfo"
    
    '测试
'    StrJson_In = "{""input"":{""audit_dept_id"":305,""cancel_status"":0,""rcpdtl_id"":[4528923,4528923]}}"
'    StrJson_In = "{""input"":{""audit_dept_id"":305,""cancel_status"":0,""rcpdtl_id"":[1]}}"
'    strService = "Zl_ExseSvr_GetChargeOffInfo"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_ExseSvr_GetChargeOffInfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    Set colCloseAccount = objServiceCall.GetJsonListValue("output.fee_cancel_list")
    
    If colCloseAccount Is Nothing Then Exit Function
    
    zlSplitService_GetCloseAccount = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_GetFee(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strRcpdtl_ids As String, _
    ByRef colPati As Collection, Optional ByVal strKeyNodes As String) As Boolean
    '取费用信息
    'strRcpdtl_ids：费用id,费用id...
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    
    'Zl_Exsesvr_GetBillDetailInfo
'  -------------------------------------------------------------------------------------------------
'  --功能：获取和药品发药业务相关的费用信息，主要用于界面显示
'  --入参：json格式
'  --Input
'  --   fee_ids    C     费用id，支持多个id，格式： 费用id,费用id,…
'  --   bill_nos   C     费用no,记录性质，格式: no,记录性质|,...
'  --出参：json格式
'  --Json_Out
'  --fee_list      [数组]每个费用ID信息
'  --  bill_prop           N    记录性质:1-收费单;2-记帐单;3-自动记帐单;4-挂号单;5-就诊卡;6-预交单
'  --  bill_no             C    单据号
'  --  fee_id              N    处方明细id(费用id)
'  --  fee_num             N    序号
'  --  iden_id             N    标识号
'  --  pati_bed            C    床号
'  --  fee_ampaid          N    实收金额
'  --  packages_num        N    付数
'  --  quantity            N    数次
'  --  placer              C    开单人
'  --  operator_code       C    操作员编号
'  --  operator_name       C    操作员姓名
'  --  create_time         D    登记时间
'  --  happen_time         D    发生时间
'  --  rcp_type            N    处方类别(按整个NO来说，1-西药，2-中药，3-混合)
'  --  fee_type            C    费别
'  --  rec_status          N    记录状态
'  --  register_id         N    挂号id
'  --  register_no         C    挂号NO
'  --  register_time       D    挂号登记时间
'  --  income_item_id      N    收入项目id
'  --  fee_origin          N    费用来源(1-门诊费用，2-住院费用)
'  --  bill_deptid         N    开单部门id
'  --  order_id            N    医嘱ID
'  --  fee_item_id         N    收费细目id
'  --  fee_status         N    费用状态
'  -------------------------------------------------------------------------------------------------
  
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_ids", strRcpdtl_ids, 0)
'    strJson_In = strJson_In & "," & GetJsonNodeString("fee_ids", strPatis, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_GetBillDetailInfo"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Exsesvr_GetBillDetailInfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '返回数组
    Set colPati = objServiceCall.GetJsonListValue("output.fee_list", strKeyNodes)
    
    If colPati Is Nothing Then Exit Function
    
    zlSplitService_GetFee = True
End Function

Public Function zlSplitService_GetNOByInvoice(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
     ByRef colOutlist As Collection) As Boolean
    '通过票据号取费用NO
    'strInput：票据号
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim strOutNos As String
    Dim n As Integer
    
    'zl_ExseSvr_GetNoByInvoice
'  -------------------------------------------------------------------------------------------------
'  --功能：按票据号发药或退药中通过录入发票号获取对应的药品处方NO
'  --入参：json格式
'  --Input
'  --   invc_no  C  1  票据号
'  --出参：json格式
'  --Json_Out
'  --  code  C  1  应答码：0-失败；1-成功
'  --  message  C  1  "应答消息： 成功时返回处方No，[数组] 失败时返回具体的错误信息"
'  --  rcp_nos  C  1 处理方单据号：多个用逗号分隔
'  -------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandle
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("invc_no", strInput, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
        
    strService = "zl_ExseSvr_GetNoByInvoice"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“zl_ExseSvr_GetNoByInvoice”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    strOutNos = objServiceCall.GetJsonNodeValue("output.rcp_nos")
    
    If strOutNos <> "" Then
        For n = 0 To UBound(Split(strOutNos, ","))
            colOutlist.Add strOutNos
        Next
    End If

    zlSplitService_GetNOByInvoice = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetExseSpec(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
     ByRef strOutput As String) As Boolean
    '获取产生过费用的规格（材料）
    'strInput：卫材id
    'strOutPut：卫材id，如果没有则返回0
    Dim StrJson_In As String
    
    On Error GoTo ErrHandle
    'Zl_Exsesvr_Getexsespec
    '  --功能：检查该规格是否产生过费用记录
    '  --input   根据材料id检查是否产生过费用记录
    '  --  item_id       N   1   收费细目id
    '  --output
    '  --  code          C   1   应答码：0-失败；1-成功
    '  --  message       C   1   应答消息：
    '  --  item_id       N   1   收费细目id
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("item_id", Val(strInput), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Getexsespec", StrJson_In, "", "", lngMode) = False Then
        MsgBox "调用“Zl_Exsesvr_Getexsespec”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    strOutput = Nvl(objServiceCall.GetJsonNodeValue("output.item_id"))
    zlSplitService_GetExseSpec = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_StuffAdjustPriceType(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    Optional ByVal strKey As String) As Boolean
    '卫材价格属性调整时产生的调价盈亏和库存变化数据处理
    'strInput:材料ID，原价格类型，新价格类型|...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    Dim arrPart As Variant
    Dim strJson_Part As String, strJson_List As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_AdjustPriceType
'  ---------------------------------------------------------------------------
'  --input      卫材价格属性调整时产生的调价盈亏和库存变化数据处理
'  --    item_list[]         材料列表
'  --       stuff_id      N    药品id
'  --       price_type_old    N    原价格类型：0-定价；1-时价
'  --       price_type_new    N    新价格类型：0-定价；1-时价
'  --output
'  --  code  C  1  应答码：0-失败；1-成功
'  --  message  C  1  应答消息：
'  ---------------------------------------------------------------------------
  
    If strInput = "" Then Exit Function
    
    '入参
    strJson_List = ""
    arrPart = Split(strInput, "|")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("stuff_id", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("price_type_old", Split(arrPart(i), ",")(1), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("price_type_new", Split(arrPart(i), ",")(2), 1)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    StrJson_In = """item_list"":[" & strJson_List & "]"
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_AdjustPriceType"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_StuffSvr_AdjustPriceType”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
   
    zlSplitService_StuffAdjustPriceType = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_StuffExistRec(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal lng材料ID As Long, _
    ByRef intExist As Integer, Optional ByVal strKey As String) As Boolean
    '获取指定的卫材是否存在收发记录
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_CheckStuffExistRec
'  ---------------------------------------------------------------------------
'  --input      判断卫材是否存在收发记录
'  --  stuff_id      N    材料id
'  --output
'  --  code  C  1  应答码：0-失败；1-成功
'  --  message  C  1  应答消息：
'  --  isexist  N 1 是否存在: 1-存在;0-不存在
'  ---------------------------------------------------------------------------
  
    If lng材料ID = 0 Then Exit Function
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("stuff_id", lng材料ID, 1)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_CheckStuffExistRec"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_StuffSvr_CheckStuffExistRec”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '返回值
    intExist = Nvl(objServiceCall.GetJsonNodeValue("output.isexist"), 0)
    
    zlSplitService_StuffExistRec = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_HightCostStuffExistRec(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal lng材料ID As Long, _
    ByRef intExist As Integer, Optional ByVal strKey As String) As Boolean
    '判断高值卫材是否存在使用记录
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_CheckHCostExistRec
'  ---------------------------------------------------------------------------
'  --input      判断高值卫材是否存在使用记录
'  --  stuff_id      N    材料id
'  --output
'  --  code  C  1  应答码：0-失败；1-成功
'  --  message  C  1  应答消息：
'  --  isexist  N 1 是否存在: 1-存在;0-不存在
'  ---------------------------------------------------------------------------
  
    If lng材料ID = 0 Then Exit Function
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("stuff_id", lng材料ID, 1)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_CheckHCostExistRec"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_StuffSvr_CheckHCostExistRec”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '返回值
    intExist = Nvl(objServiceCall.GetJsonNodeValue("output.isexist"), 0)
    
    zlSplitService_HightCostStuffExistRec = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_StuffExistStock(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef intExist As Integer, Optional ByVal strKey As String) As Boolean
    '判断卫材是否存在库数据
    'strInput:材料/诊疗ID，按品种/规格（0-按规格，1-按品种）
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_CheckExistStock
'  ---------------------------------------------------------------------------
'  --input      判断卫材是否存在库数据
'  --  stuff_id      N  1  卫材id
'  --  is_item      N  1  是否按品种查询：0-按规格查询，1-按品种查询
'  --output
'  --  code  C  1  应答码：0-失败；1-成功
'  --  message  C  1  应答消息：
'  --  isexist  N 1 是否存在: 1-存在;0-不存在
'  ---------------------------------------------------------------------------
  
    If strInput = "" Then Exit Function
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("stuff_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_item", Split(strInput, ",")(1), 1)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_CheckExistStock"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_StuffSvr_CheckExistStock”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '返回值
    intExist = Nvl(objServiceCall.GetJsonNodeValue("output.isexist"), 0)
    
    zlSplitService_StuffExistStock = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_StuffExecutePrice(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal lng材料ID As Long, _
Optional ByVal strKey As String) As Boolean
    '检查卫材售价，成本价是否存在已生效但未执行的价格，如果存在则执行调价
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_ExecutePrice
'  ---------------------------------------------------------------------------
'  --input      检查卫材售价，成本价是否存在已生效但未执行的价格，如果存在则执行调价
'  --  stuff_id      N    材料id
'  --output
'  --  code  C  1  应答码：0-失败；1-成功
'  --  message  C  1  应答消息：
'  ---------------------------------------------------------------------------
  
    If lng材料ID = 0 Then Exit Function
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("stuff_id", lng材料ID, 1)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_ExecutePrice"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_StuffSvr_ExecutePrice”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    
    zlSplitService_StuffExecutePrice = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_StuffGetCostPriceAdjust(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '获取卫材成本价调价记录
    'strInput:品种ID，单位（0-散装单位，1-包装单位）
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_GetCostPriceAdjust
'  ---------------------------------------------------------------------------
'  --功能：获取卫材成本价调价记录
'  --input
'  --  stuff_id      N   1 材料id
'  --  show_unit    N   1   显示单位:0-散装单位;1-库房单位
'  --output
'  --  code  C  1  应答码：0-失败；1-成功
'  --  message  C  1  应答消息：
'  --  price_list[]  卫材成本价调价记录
'  --     stuff_id   N 1  材料ID
'  --     stuff_name   C 1  材料信息
'  --     stock_name   C 1  库房
'  --     batch_number   C 1  批号
'  --     effective_time   C 1  效期
'  --     place_name   C 1  产地
'  --     unit_name   C 1  单位
'  --     cost_old   N 1  原成本价
'  --     cost_new    N 1  现成本价
'  --     adjust_time   C 1  调价时间
'  --     adjust_reson   C 1  调价说明
'  --     adjust_no   C 1  调价单据号
'  --     drug_revoke_time  C 1 撤档时间
'  --     node_no      C    0  站点编码
'  --     is_stock    N   1 是否有库存数据  0-否，1-是
'  ---------------------------------------------------------------------------
  
    If strInput = "" Then Exit Function
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("stuff_id", Val(Split(strInput, ",")(0)), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("show_unit", Val(Split(strInput, ",")(1)), 1)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_GetCostPriceAdjust"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_StuffSvr_GetCostPriceAdjust”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '返回值
    Set colOutlist = objServiceCall.GetJsonListValue("output.price_list")
    
    zlSplitService_StuffGetCostPriceAdjust = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetStockShow(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '获取指定库房的库存数据，用于显示
    'strInput:库房id，库房id...
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_GetStockShow
'  ---------------------------------------------------------------------------
'  --功能：获取指定库房的库存数据，用于显示
'  --入参：Json_In:格式
'  --  input
'  --    warehouse_ids        C   1   库房ID串
'  --出参: Json_Out,格式如下
'  --  output
'  --    code                N 1 应答吗：0-失败；1-成功
'  --    message             C 1 应答消息：失败时返回具体的错误信息
'  --    item_list
'  --      stuff_id              N   1   材料ID
'  --      warehouse_id          N   1   库房ID
'  --      stock                N   1   可用数量
'  --      real_stock          N  1 实际库存
'  --      avg_price           N  1 平均售价
'  --      avg_cost            N  1 平均成本价
'  ---------------------------------------------------------------------------
  
    If strInput = "" Then Exit Function
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("warehouse_ids", strInput, 0)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_GetStockShow"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_StuffSvr_GetStockShow”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '返回值
    Set colOutlist = objServiceCall.GetJsonListValue("output.item_list", strKey)
    
    zlSplitService_GetStockShow = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function





Public Function zlSplitService_GetRequestCancel(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection) As Boolean
    '取销帐申请记录
    'strInput：费用id,费用id...
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    
'  --查询是否存在销帐申请记录
'  --入参      json
'  --  input      查询是否存在销帐申请记录
'  --    rcpdtl_id          C     单据明细id,[数组]：[1,2,3]
'  --    request_type       N    申请类别
'  --    cancel_status       N  1 状态
'  --出参      json
'  -- output
'  --   code     C  1   应答码：0-失败；1-成功
'  --   message  C  1   应答消息：
'  --   fee_cancel_list      [数组]满足条件的每个费用销帐记录
'  --     rcpdtl_id          N    处方明细id(费用id)

    '入参
    StrJson_In = ""
'    StrJson_In = StrJson_In & GetJsonNodeString("rcpdtl_id", "[" & strInput & "]", 0)
'    If Not IsNull(colInput("rcpdtl_id")) Then StrJson_In = StrJson_In & "," & """rcpdtl_id"":" & "[" & colInput("rcpdtl_id") & "]"
    
    StrJson_In = StrJson_In & """rcpdtl_id"":" & "[" & strInput & "]"
    StrJson_In = StrJson_In & "," & GetJsonNodeString("request_type", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_status", 0, 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Getrequestcancel"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Exsesvr_Getrequestcancel”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.fee_cancel_list")
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetRequestCancel = True
End Function

Public Function zlSplitService_CallAccountDel_Check(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef strMsg As String, Optional ByRef int执行状态 As Integer = 1) As Boolean
    '门诊/住院记录销账检查
    'strInput:服务需要的入参格式：no,已结禁止销帐(1),医保禁止部分销帐(0),操作状态(1),费用来源|序号,销帐数量;序号,销帐数量...|费用id,已发数量;费用id,已发数量...
    Dim arrPart As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String, strJson_Part As String
    Dim strService As String
    Dim colJsonOutList As Collection    '服务出参list返回集合
    Dim varList As Variant   '集合元素
    Dim strPart1 As String, strPart2 As String, strPart3 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String
    
'    on error goto errHandle
    
    '服务：Zl_ExseSvr_DelBill_Check 门诊/住院记录通用服务
    
'  ---------------------------------------------------------------------------
'  --功能：针对指定单据指定行行进行销帐
'  --入参：Json_In:格式
'  --input
'  --        fee_no                  C   1   费用单据号
'  --        fee_bill_type           N   1   单据性质:2-门诊记帐单,3-自动记帐单
'  --        balance_ban_writeoffs                   N   1   已结禁止销帐:如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
'  --        part_ban_writeoffs                  N   1   禁止部分销帐:1-不允许；0-允许
'  --        fee_origin        N 1            费用来源（1-门诊记帐，2-住院记帐）
'  --        item_list[]                         本次销帐列表
'  --            serial_num              N   1   序号
'  --            quantity                N   1   销帐数量(为零时，按序号直接销帐)
'  --        excute_list[]                           药品及卫材所对应已执行列表
'  --            fee_id              N   1   费用ID
'  --            sended_num              N   1   已发数量
'  --出参: Json_Out,格式如下
'  --    output
'  --        code                    N   1   应答吗：0-失败；1-成功
'  --        message                 C   1   应答消息：失败时返回具体的错误信息
'  --        item_list[]                         单据数据列表
'  --            serial_num              N   1   序号
'  --            quantity                N   1   销帐数量
'  --            execute_tag             N   1   执行状态：0-未执行;1-已执行;2-部分执行
'
'  ---------------------------------------------------------------------------

    If strInput = "" Then zlSplitService_CallAccountDel_Check = True: Exit Function
    
    strJson_List = ""
    
    arrPart = Split(strInput, "|")
    
    strPart1 = arrPart(0)
    strPart2 = arrPart(1)
    strPart3 = arrPart(2)
        
    strJson_In_Part1 = ""
    strJson_In_Part1 = strJson_In_Part1 & "" & GetJsonNodeString("fee_no", Split(strPart1, ",")(0), 0)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("fee_bill_type", 2, 1)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("balance_ban_writeoffs", Split(strPart1, ",")(1), 1)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("part_ban_writeoffs", Split(strPart1, ",")(2), 1)
'    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("oper_type", Split(strPart1, ",")(3), 1)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("fee_origin", Split(strPart1, ",")(4), 1)

    strJson_List = ""
    arrPart = Split(strPart2, ";")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("serial_num", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("quantity", Split(arrPart(i), ",")(1), 1)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    strJson_In_Part2 = ",""item_list"":[" & strJson_List & "]"
    
    strJson_List = ""
    arrPart = Split(strPart3, ";")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("fee_id", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("sended_num", Split(arrPart(i), ",")(1), 1)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    strJson_In_Part3 = ",""excute_list"":[" & strJson_List & "]"
    
    '汇总
    StrJson_In = "{""input"":{" & strJson_In_Part1 & strJson_In_Part2 & strJson_In_Part3 & "}}"
    
    strService = "Zl_ExseSvr_DelBill_Check"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_ExseSvr_DelBill_Check”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    int执行状态 = objServiceCall.GetJsonNodeValue("output.item_list[0].execute_tag")
    
    If strJson_Out = "0" Then
        strMsg = strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
        zlSplitService_CallAccountDel_Check = False
        Exit Function
    End If

    zlSplitService_CallAccountDel_Check = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function zlSplitService_GetPati(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '取病人信息
    'strInput：项目名称;项目内容
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String, strJson_List As String
    Dim strService As String
    Dim arrItem As Variant
    Dim i As Integer
        
'  ---------------------------------------------------------------------------
'Zl_Patisvr_Getpatiinfo
'  --功能:获取病人信息
'  --入参：Json_In:格式
'  --    input
'  --      pati_id           N 1 病人id  病人ID<>0时，查询列表中的条件无效
'  --      query_type        N 1 查询类型:如：0-基本;1-基本+联系人;3-所有
'  --      query_card        N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
'  --      query_family      N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
'  --      query_drug        N 1 是否包含过敏药物:1-包含，0-不包含
'  --      query_immune      N 1 是否包含免疫修:1-包含;0-不包含
'  --      query_cons_list   C 1 查询条件:可以选择一定条件进行查询（是And关系),只有一行
'  --        pati_ids        C   病人IDs:多个用逗号
'  --        pati_name       C   姓名:可以代%分号表表按姓名匹配
'  --        outpatient_num  C   门诊号
'  --        pati_idcard     C   身份证号
'  --        contacts_idcard C   联系人身份证号
'  --        cardtype_id     N   医疗卡类别ID
'  --        medc_card_name  N   医疗卡名称
'  --        card_no         C   卡号
'  --        qrcode          C   二维码
'  --        iccard_no       C   Ic卡号
'  --        visit_card      C   就诊卡号
'  --        insurance_num   C   医保号
'  --        qrspt_statu     C   查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
'  --        phone_number    C   手机号
'  --        pati_bed        C   床号
'  --出参      json
'  --output
'  --    code                N   1   应答码：0-失败；1-成功
'  --    message             C   1   应答消息： 失败时返回具体的错误信息
'  --    pati_list[]                 病人信息列表
'  --    pati_id             N   1   病人id
'  --    pati_pageid         N   1   主页id：病人信息.主页ID
'  --    pati_name           C   1   姓名
'  --    pati_sex            C   1   性别
'  --    pati_age            C   1   年龄
'  --    pati_birthdate      C   1   出生日期：yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   1   费别
'  --    outpatient_num      C   1   门诊号
'  --    inpatient_num       C   1   住院号
'  --    mdlpay_mode_name    C   1   医疗付款方式名称
'  --    mdlpay_mode_code    C   1   医疗付款方式编码
'  --    pati_nation         C   1   民族
'  --    insurance_num       C   1   医保号
'  --    pati_idcard         C   1   身份证号
'  --    vcard_no            C   1   就诊卡号
'  --    iccard_no           C   1   Ic卡号
'  --    health_num          C   1   健康号
'  --    inp_times           N   1   住院次数
'  --    pati_education      C   1   学历
'  --    ocpt_name           C   1   职业
'  --    pati_identity       C   1   身份
'  --    ntvplc_name         C   1   籍贯
'  --    country_name        C   1   国籍
'  --    pati_marital_cstatus    C   1   婚姻状况
'  --    pat_home_addr           C   1   家庭地址
'  --    pat_home_phno           C   1   家庭电话
'  --    pat_home_postcode   C   1   家庭地址邮编
'  --    pati_area           C   1   区域
'  --    pati_birthplace     C   1   出生地点
'  --    pat_hous_addr       C   1   户口地址
'  --    pat_hous_postcode   C   1   户口地址邮编
'  --    emp_name            C   1   工作单位名称
'  --    emp_phno            C   1   单位电话
'  --    emp_postcode        C   1   单位邮编
'  --    emp_bank_name       C   1   单位开户行
'  --    emp_bank_accnum     C   1   单位帐号
'  --    emp_addr             C   1   单位地址
'  --    ctt_unit_id         N   1   合同单位ID
'  --    phone_number        C   1   手机号
'  --    pati_bed            C   1   当前床号
'  --    pati_type           C   1   病人类型(普通，医保，留观)
'  --    insurance_type      C   1   险类
'  --    insurance_name      C   1   险类名称
'  --    pati_wardarea_id    N   1   当前病区id
'  --    pati_wardarea_name  C   1   当前病区名称
'  --    pati_dept_id        N   1   当前科室id
'  --    pati_dept_name      C   1   当前科室名称
'  --    adta_time           C   1   入院时间:yyyy-mm-dd hh24:mi:ss
'  --    adtd_time           C   1   出院时间:yyyy-mm-dd hh24:mi:ss
'  --    contacts_name       C   1   联系人姓名
'  --    contacts_relation   C   1   联系人关系
'  --    contacts_idcard     C   1   联系人身份证号
'  --    contacts_addr       C   1   联系人地址
'  --    contacts_phno       C   1   联系人电话
'  --    pat_grdn_name       C   1   监护人
'  --    cert_no_other       C   1   其他证件
'  --    is_inhspt            C   1   是否在院:1-在院 ;0-不在院
'  --    pati_show_color      N   1   病人显示颜色
'  --    visit_room           C   1   就诊诊室
'  --    visit_statu          N   1   就诊状态
'  --    visit_time           C   1   就诊时间:yyyy-mm-dd hh24:mi:ss
'  --    create_time          C   1   登记时间:yyyy-mm-dd hh24:mi:ss
'  --    pati_email           C   1   email
'  --    pati_qq              C   1   qq
'  --    card_captcha         C   1  卡验证码
'  --    family_list[]        C   1   家属成员:病人家属() query_family=1返回
'  --        family_id        N   1   家属id  query_family=1
'  --        family_relation  C   1   关系
'  --    drug_list[]          C   1   过敏药物列表    query_drug=1时返回
'  --        pat_algc_cadn_id N   1   过敏药品ID
'  --        pat_algc_cadn    C   1   过敏药物名称
'  --        allergy_info     C   1   过每药物反应
'  --    immune_list[]        C   1   病人免疫列表    query_immune=1时返回
'  --        vaccinate_time   C   1   接种时间:yyyy-mm-dd hh24:mi:ss
'  --        vaccinate_name   C   1   接种名称
'  --    card_list[]          C   1   病人医疗卡信息列表(如果条件中传入了卡类别ID的，则返回该卡类别的卡信息)  query_card=1时返回
'  --        cardtype_id      N   1   医疗卡类别ID
'  --        card_no          C   1   卡号
'  --        card_pwd         C   1   密码
'  ---------------------------------------------------------------------------
    
    On Error GoTo ErrHandle
    
'  --      pati_id           N 1 病人id  病人ID<>0时，查询列表中的条件无效
'  --      query_type        N 1 查询类型:如：0-基本;1-基本+联系人;3-所有
'  --      query_card        N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
'  --      query_family      N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
'  --      query_drug        N 1 是否包含过敏药物:1-包含，0-不包含
'  --      query_immune      N 1 是否包含免疫修:1-包含;0-不包含
'  --      query_cons_list   C 1 查询条件:可以选择一定条件进行查询（是And关系),只有一行
'  --        pati_ids        C   病人IDs:多个用逗号
    
    '入参
    StrJson_In = ""
    
    If Split(strInput, ";")(0) = "病人id" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ";")(1), 1)
    Else
        StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", 0, 1)
    End If
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", 3, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_card", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_family", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_drug", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_immune", 0, 1)
    
    Select Case Split(strInput, ";")(0)
        Case "病人ids"
            strJson_List = GetJsonNodeString("pati_ids", Split(strInput, ";")(1), 0)
        Case "医保号"
            strJson_List = GetJsonNodeString("insurance_num", Split(strInput, ";")(1), 0)
    End Select
    
    If strJson_List <> "" Then
        strJson_List = strJson_List & "," & GetJsonNodeString("qrspt_statu", 2, 1)
    End If
    
    strJson_List = ",""query_cons_list"":{" & strJson_List & "}"
    
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"
    strService = "Zl_Patisvr_Getpatiinfo"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Patisvr_Getpatiinfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.pati_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetPati = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallPatiIsOut(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strPaitId As Long, ByVal strPageId As Long, ByRef intOutSign As Integer) As Boolean
    '发药/退药调用服务：查询病人是否已出院
    '病人信息：病人id，主页id
    'intOutSign：0-未出院，1-出院
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '服务出参list返回集合
    Dim varList As Variant   '集合元素
    
    On Error GoTo ErrHandle
    
    '服务：zl_CisSvr_PatiIsOut
'    input           查询病人是否已经出院
'       pati_id N   1   病人id
'       pati_pageid  N   1   主页id
'
'    output
'      code    C   1   应答码：0-失败；1-成功
'      message C   1   "应答消息：
'      pati_outsign    N       出院标记：0-未出院，1-出院

    
    If strPaitId = 0 Then Exit Function

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", strPaitId, 1)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", strPageId, 1)

    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "zl_CisSvr_PatiIsOut"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“zl_CisSvr_PatiIsOut”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallPatiIsOut = False
        Exit Function
    Else
        intOutSign = Val(objServiceCall.GetJsonNodeValue("output.pati_outsign"))
    End If

    zlSplitService_CallPatiIsOut = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiByRange(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '按范围查找病人信息
    'strInput： 查询条件，目前仅按病区ID查找
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    
    'Zl_Cissvr_Getpatpageinfbyrange
'  ---------------------------------------------------------------------------
'  --功能:获取病案信息
'  --入参：Json_In:格式
'  --  input
'  --    query_type          N 1 查询类型:0-基本;1-基本扩展
'  --    wararea_ids         C   病区ids:多个用逗号
'  --    dept_ids            C   科室IDs:个用逗号
'  --    pati_ids            C   病人ids:多个用逗号分离
'  --    pati_pageIds        C   主页IDs:病人id:主页id,…
'  --    adta_start_time     C   入院开始时间:yyyy-mm-dd hh24:mi:ss
'  --    adta_end_time       C   入院结束时间:yyyy-mm-dd hh24:mi:ss
'  --    adtd_start_time     C   出院开始时间:yyyy-mm-dd hh24:mi:ss
'  --    adtd_end_time       C   出院结束时间:yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   费别
'  --    inp_status          N   住院状态:0-在院病人;1-出院病人;2-在院或出院
'  --    pati_natures        C   病人性质：多个用逗号分0-普通住院病人,1-门诊留观病人,2-住院留观病人，NULL-表示不区分
'  --    pati_name           C   姓名:可以代%分号表表按姓名匹配
'  --    nodeno              C   站点编号
'  --    change_dept_pati    N   是否查询转科病人
'  --出参      json
'  --output
'  -- code                   N 1 应答码：0-失败；1-成功
'  -- message                C 1 应答消息： 失败时返回具体的错误信息
'  --   page_list[]          数据组  √  √
'  --    pati_id             N    病人id  √  √
'  --    pati_pageid         N    主页id  √  √
'  --    pati_name           C    姓名  √  √
'  --    pati_sex            C    性别  √  √
'  --    pati_age            C    年龄  √  √
'  --    inpatient_num       C    住院号  √  √
'  --    pati_bed            C    出院病床  √  √
'  --    insurance_type      N    险类  √  √
'  --    fee_category        C    费别  √  √
'  --    pati_type           C    病人类型(普通,医保,留观)  √  √
'  --    adta_time           C    入院时间:yyyy-mm-dd hh24:mi:ss  √  √
'  --    adtd_time           C    出院时间:yyyy-mm-dd hh24:mi:ss  √  √
'  --    si_inp_status       N    住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)  √  √
'  --    pati_nature         N    病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
'  --    pati_wardarea_id    N    当前病区id
'  --    pati_wardarea_name  C    当前病区名称
'  --    pati_dept_id        N    当前科室id
'  --    pati_dept_name      C    当前科室名称
'  --    mdlpay_mode_name    C    医疗付款方式名称
'  --    mdlpay_mode_code    C    医疗付款方式编码
'  --    pat_rsdpscn         C    住院医师
'  --    pati_desc           C    病人备注
'  --    catalog_date        C    编目日期:yyyy-mm-dd hh24:mi:ss
'  --    create_pati         C    登记人
'  --    in_objective     C    住院目的
'  ---------------------------------------------------------------------------
    
    On Error GoTo ErrHandle
    
'  --    query_type          N 1 查询类型:0-基本;1-基本扩展
'  --    wararea_ids         C   病区ids:多个用逗号
'  --    dept_ids            C   科室IDs:个用逗号
'  --    pati_ids            C   病人ids:多个用逗号分离
'  --    pati_pageIds        C   主页IDs:病人id:主页id,…
'  --    adta_start_time     C   入院开始时间:yyyy-mm-dd hh24:mi:ss
'  --    adta_end_time       C   入院结束时间:yyyy-mm-dd hh24:mi:ss
'  --    adtd_start_time     C   出院开始时间:yyyy-mm-dd hh24:mi:ss
'  --    adtd_end_time       C   出院结束时间:yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   费别
'  --    inp_status          N   住院状态:0-在院病人;1-出院病人;2-在院或出院
'  --    pati_natures        C   病人性质：多个用逗号分0-普通住院病人,1-门诊留观病人,2-住院留观病人，NULL-表示不区分
'  --    pati_name           C   姓名:可以代%分号表表按姓名匹配
'  --    nodeno              C   站点编号
'  --    change_dept_pati    N   是否查询转科病人
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("wararea_ids", Val(strInput), 0)

    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatpageinfbyrange"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Cissvr_Getpatpageinfbyrange”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.page_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetPatiByRange = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiId(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOutput As String, Optional ByVal strKey As String) As Boolean
    '取病人信息，根据住院号，或病区+床号
    'strInput：住院号 或 病区id|床号，根据是否有“|”来区分
    'strOutPut：返回信息，病人ID
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
'  ---------------------------------------------------------------------------
'Zl_Cissvr_Getpatiid
'  --功能:获取病人信息
'  --入参：Json_In:格式
'  --input
'  --   wardarea_id          N 1 当前病区id
'  --   pati_bed             C 1 当前床号
'  --   inpatient_num        C 1 住院号
'  --output
'  --    code                N 1 应答码：0-失败；1-成功
'  --    message             C 1 应答消息： 失败时返回具体的错误信息
'  --    pati_id             N 1 病人ID:未找到时也成功，返回0
'  --    pati_pageid         N   主页ID
'  ---------------------------------------------------------------------------
    
    On Error GoTo ErrHandle
    '入参
    StrJson_In = ""
    
    If InStr(strInput, "|") > 0 Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("wardarea_id", Split(strInput, "|")(0), 1)
        StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_bed", Split(strInput, "|")(1), 0)
    Else
        StrJson_In = StrJson_In & "" & GetJsonNodeString("inpatient_num", strInput, 0)
    End If
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatiid"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Patisvr_Getpatiinfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回数据
    strOutput = objServiceCall.GetJsonNodeValue("output.pati_id")
    
    zlSplitService_GetPatiId = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_GetPatiName(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal colInput As Collection, _
    ByRef colPati As Collection, Optional ByVal strKey As String, Optional ByVal bytQueryType As Integer = 3) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:用于按一定条件单独查询病人信息
    '入参:
    '   colInput 查询条件组合，Json中input各节点作为元素的KEY值，集合某元素为空表示该节点值为空
    '   bytQueryType 病人信息查询类型:如：0-基本;1-基本+联系人;3-所有
    '出参:
    '返回:
    '说明:目前支持的查询条件，一般都是按其中一种查询：病人ID，门诊号，姓名，就诊卡号，医保号，床号
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    
    On Error GoTo ErrHandle
    
    'Zl_Patisvr_Getpatiinfo
'  ---------------------------------------------------------------------------
'  --功能:获取病人信息
'  --入参：Json_In:格式
'  --    input
'  --      pati_id           N 1 病人id  病人ID<>0时，查询列表中的条件无效
'  --      query_type        N 1 查询类型:如：0-基本;1-基本+联系人;3-所有
'  --      query_card        N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
'  --      query_family      N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
'  --      query_drug        N 1 是否包含过敏药物:1-包含，0-不包含
'  --      query_immune      N 1 是否包含免疫修:1-包含;0-不包含
'  --      query_cons_list   C 1 查询条件:可以选择一定条件进行查询（是And关系),只有一行
'  --        pati_ids        C   病人IDs:多个用逗号
'  --        pati_name       C   姓名:可以代%分号表表按姓名匹配
'  --        outpatient_num  C   门诊号
'  --        pati_idcard     C   身份证号
'  --        contacts_idcard C   联系人身份证号
'  --        cardtype_id     N   医疗卡类别ID
'  --        medc_card_name  N   医疗卡名称
'  --        card_no         C   卡号
'  --        qrcode          C   二维码
'  --        iccard_no       C   Ic卡号
'  --        visit_card      C   就诊卡号
'  --        insurance_num   C   医保号
'  --        qrspt_statu     C   查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
'  --        phone_number    C   手机号
'  --        pati_bed        C   床号
'  --        dept_id         N   当前科室ID
'  --出参      json
'  --output
'  --    code                N   1   应答码：0-失败；1-成功
'  --    message             C   1   应答消息： 失败时返回具体的错误信息
'  --    pati_list[]                 病人信息列表
'  --    pati_id             N   1   病人id
'  --    pati_pageid         N   1   主页id：病人信息.主页ID
'  --    pati_name           C   1   姓名
'  --    pati_sex            C   1   性别
'  --    pati_age            C   1   年龄
'  --    pati_birthdate      C   1   出生日期：yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   1   费别
'  --    outpatient_num      C   1   门诊号
'  --    inpatient_num       C   1   住院号
'  --    mdlpay_mode_name    C   1   医疗付款方式名称
'  --    mdlpay_mode_code    C   1   医疗付款方式编码
'  --    pati_nation         C   1   民族
'  --    insurance_num       C   1   医保号
'  --    pati_idcard         C   1   身份证号
'  --    vcard_no            C   1   就诊卡号
'  --    iccard_no           C   1   Ic卡号
'  --    health_num          C   1   健康号
'  --    inp_times           N   1   住院次数
'  --    pati_education      C   1   学历
'  --    ocpt_name           C   1   职业
'  --    pati_identity       C   1   身份
'  --    ntvplc_name         C   1   籍贯
'  --    country_name        C   1   国籍
'  --    pati_marital_cstatus    C   1   婚姻状况
'  --    pat_home_addr           C   1   家庭地址
'  --    pat_home_phno           C   1   家庭电话
'  --    pat_home_postcode   C   1   家庭地址邮编
'  --    pati_area           C   1   区域
'  --    pati_birthplace     C   1   出生地点
'  --    pat_hous_addr       C   1   户口地址
'  --    pat_hous_postcode   C   1   户口地址邮编
'  --    emp_name            C   1   工作单位名称
'  --    emp_phno            C   1   单位电话
'  --    emp_postcode        C   1   单位邮编
'  --    emp_bank_name       C   1   单位开户行
'  --    emp_bank_accnum     C   1   单位帐号
'  --    emp_addr             C   1   单位地址
'  --    ctt_unit_id         N   1   合同单位ID
'  --    phone_number        C   1   手机号
'  --    pati_bed            C   1   当前床号
'  --    pati_type           C   1   病人类型(普通，医保，留观)
'  --    insurance_type      C   1   险类
'  --    insurance_name      C   1   险类名称
'  --    pati_wardarea_id    N   1   当前病区id
'  --    pati_wardarea_name  C   1   当前病区名称
'  --    pati_dept_id        N   1   当前科室id
'  --    pati_dept_name      C   1   当前科室名称
'  --    adta_time           C   1   入院时间:yyyy-mm-dd hh24:mi:ss
'  --    adtd_time           C   1   出院时间:yyyy-mm-dd hh24:mi:ss
'  --    contacts_name       C   1   联系人姓名
'  --    contacts_relation   C   1   联系人关系
'  --    contacts_idcard     C   1   联系人身份证号
'  --    contacts_addr       C   1   联系人地址
'  --    contacts_phno       C   1   联系人电话
'  --    pat_grdn_name       C   1   监护人
'  --    cert_no_other       C   1   其他证件
'  --    is_inhspt            C   1   是否在院:1-在院 ;0-不在院
'  --    pati_show_color      N   1   病人显示颜色
'  --    visit_room           C   1   就诊诊室
'  --    visit_statu          N   1   就诊状态
'  --    visit_time           C   1   就诊时间:yyyy-mm-dd hh24:mi:ss
'  --    create_time          C   1   登记时间:yyyy-mm-dd hh24:mi:ss
'  --    pati_email           C   1   email
'  --    pati_qq              C   1   qq
'  --    card_captcha         C   1  卡验证码
'  --    family_list[]        C   1   家属成员:病人家属() query_family=1返回
'  --        family_id        N   1   家属id  query_family=1
'  --        family_relation  C   1   关系
'  --    drug_list[]          C   1   过敏药物列表    query_drug=1时返回
'  --        pat_algc_cadn_id N   1   过敏药品ID
'  --        pat_algc_cadn    C   1   过敏药物名称
'  --        allergy_info     C   1   过每药物反应
'  --    immune_list[]        C   1   病人免疫列表    query_immune=1时返回
'  --        vaccinate_time   C   1   接种时间:yyyy-mm-dd hh24:mi:ss
'  --        vaccinate_name   C   1   接种名称
'  --    card_list[]          C   1   病人医疗卡信息列表(如果条件中传入了卡类别ID的，则返回该卡类别的卡信息)  query_card=1时返回
'  --        cardtype_id      N   1   医疗卡类别ID
'  --        card_no          C   1   卡号
'  --        card_pwd         C   1   密码
'  ---------------------------------------------------------------------------
'  --      pati_id           N 1 病人id  病人ID<>0时，查询列表中的条件无效
'  --      query_type        N 1 查询类型:如：0-基本;1-基本+联系人;3-所有
'  --      query_card        N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
'  --      query_family      N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
'  --      query_drug        N 1 是否包含过敏药物:1-包含，0-不包含
'  --      query_immune      N 1 是否包含免疫修:1-包含;0-不包含
'  --      query_cons_list   C 1 查询条件:可以选择一定条件进行查询（是And关系),只有一行

    '入参
    StrJson_In = ""
    If Not IsNull(colInput("pati_id")) Then
        StrJson_In = GetJsonNodeString("pati_id", colInput("pati_id"), 1)
    Else
        StrJson_In = GetJsonNodeString("pati_id", 0, 1)
    End If
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", bytQueryType, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_card", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_family", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_drug", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_immune", 0, 1)
    
    If IsNull(colInput("pati_id")) Then
        '可能按其中一种方式查询：门诊号，姓名，就诊卡号，医保号，床号
        strJson_List = ""
        If Not IsNull(colInput("outpatient_num")) Then strJson_List = strJson_List & "," & GetJsonNodeString("outpatient_num", colInput("outpatient_num"), 0)
        If Not IsNull(colInput("pati_name")) Then strJson_List = strJson_List & "," & GetJsonNodeString("pati_name", colInput("pati_name"), 0)
        If Not IsNull(colInput("pati_vcard_no")) Then strJson_List = strJson_List & "," & GetJsonNodeString("visit_card", colInput("pati_vcard_no"), 0)
        'If Not IsNull(colInput("insurance_num")) Then StrJson_In = StrJson_In & "," & GetJsonNodeString("insurance_num", colInput("insurance_num"), 0)
        If Not IsNull(colInput("pati_bed")) Then strJson_List = strJson_List & "," & GetJsonNodeString("pati_bed", colInput("pati_bed"), 0)
        If strJson_List <> "" Then
            strJson_List = GetJsonNodeString("qrspt_statu", 2, 1) & strJson_List '2-门诊及在院
            strJson_List = ",""query_cons_list"":{" & strJson_List & "}"
        ElseIf ExistsColObject(colInput, "pati_deptid") Then '按科室查询
            strJson_List = GetJsonNodeString("qrspt_statu", 1, 1)   '1-在院
            strJson_List = strJson_List & "," & GetJsonNodeString("dept_id", colInput("pati_deptid"), 1)
            strJson_List = ",""query_cons_list"":{" & strJson_List & "}"
        End If
    End If
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"
    
    strService = "Zl_Patisvr_Getpatiinfo"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Patisvr_Getpatiinfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    Set colPati = objServiceCall.GetJsonListValue("output.pati_list", strKey)
    
    If colPati Is Nothing Then Exit Function
    
    zlSplitService_GetPatiName = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function zlSplitService_GetPatiPage(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String, Optional ByRef colOutListBaby As Collection) As Boolean
    '取病案主页信息
    'strInput：病人id:主页id,...
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '集合元素
    Dim colTmp As New Collection, colbaby As New Collection
    Dim i As Integer, n As Integer
    
'---------------------------------------------------------------------------
'Zl_Cissvr_Getpatipageinfo
'  --功能:获取病案主页相关信息
'  --入参：Json_In:格式
'  --    input
'  --      query_type          C 1 查询类型:0-基本信息;1-基本信息的展;2-仅取主页
'  --      pati_pageids        C 1 病人信息,格式两种:一种是:病人id:主页ID,…;一种：病人id,…
'  --      is_babyinfo         N 1 是否包含婴儿信息:1-包含;0-不包含
'  --      is_transdeptinfo    N 1 是否包含转科信息:1-包含;0-不包含
'  --      is_lastpage         N 1 是否取最后一次住院
'  --出参: Json_Out,格式如下
'  --  output
'  --    code                  N 1 应答码：0-失败；1-成功
'  --    message               C 1 应答消息：失败时返回具体的错误信息
'  --    pati_count            N 1 查询的病人信息条数
'  --    page_list[]             1 数据组
'  --      pati_id             N 1 病人id
'  --      pati_pageid         N 1 主页id
'  --      pati_name           C 1 姓名
'  --      pati_sex            C 1 性别
'  --      pati_age            C 1 年龄
'  --      inpatient_num       C 1 住院号
'  --      fee_category        C 1 费别
'  --      mdlpay_mode_name    C 1 医疗付款方式名称
'  --      mdlpay_mode_code    C 1 医疗付款方式编码
'  --      pati_bed            C 1 当前床号
'  --      pati_type           C 1 病人类型(普通，医保，留观)
'  --      pati_show_color     N    病人类型颜色
'  --      pati_education      C 1 学历
'  --      ocpt_name           C 1 职业
'  --      country_name        C 1 国籍
'  --      pati_marital_cstatus  C 1 婚姻状况
'  --      pati_nature         N 1 病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
'  --      audit_sign          N 1 审核标志:病案主页.审核标志
'  --      si_inp_status       N 1 住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)
'  --      pati_wardarea_id    N 1 当前病区id
'  --      pati_wardarea_name  C 1 当前病区名称
'  --      pati_dept_id        N 1 当前科室id
'  --      pati_dept_name      C 1 当前科室名称
'  --      adta_time           C 1 入院时间:yyyy-mm-dd hh24:mi:ss
'  --      adtd_time           C 1 出院时间:yyyy-mm-dd hh24:mi:ss
'  --      insurance_type      N 1 险类
'  --      rgst_id             N 1 挂号id
'  --      catalog_date        C 1 编目日期:yyyy-mm-dd hh24:mi:ss
'  --      in_objective        C 1 住院目的
'  --      reg_name            C 1 登记人
'  --      reg_date            C 1 住院登记时间
'  --      pat_rsdpscn         C 1 住院医师
'  --      pati_desc           C 1 病人备注
'  --      baby_list[]           1 婴儿信息，[数组]
'  --        pati_id           N 1 病人id
'  --        pati_pageid       N 1 主页id
'  --        baby_num          N 1 婴儿序号
'  --        baby_name         C 1 婴儿姓名
'  --        baby_sex          C 1 婴儿性别
'  --        baby_date         C 1 出生时间
'  --      trans_list[]        C   转科列表信息
'  --        start_reason      C 1 开始原因
'  --        start_time        C 1 开始时间:yyyy-mm-dd hh24:mi:ss
'  --        dept_name         C 1 科室名称
'  ---------------------------------------------------------------------------
    
    On Error GoTo ErrHandle
    
'  --    input
'  --      query_type          C 1 查询类型:0-基本信息;1-基本信息的展;2-仅取主页
'  --      pati_pageids        C 1 病人信息,格式两种:一种是:病人id:主页ID,…;一种：病人id,…
'  --      is_babyinfo         N 1 是否包含婴儿信息:1-包含;0-不包含
'  --      is_transdeptinfo    N 1 是否包含转科信息:1-包含;0-不包含
'  --      is_lastpage         N 1 是否取最后一次住院
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageids", strInput, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_babyinfo", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_transdeptinfo", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_lastpage", 0, 1)
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatipageinfo"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Cissvr_Getpatipageinfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.page_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    Set colOutListBaby = New Collection
    
    '婴儿数据的病人ID,主页id相同，婴儿序号不同，要么不用key，要么用 病人id+主页id+婴儿序号 作为key
    For n = 1 To colOutlist.Count
        Set colbaby = objServiceCall.GetJsonListValue("output.page_list[" & n - 1 & "].baby_list")
        If colbaby.Count > 0 Then
            For i = 1 To colbaby.Count
                Set colTmp = Nothing
                
                colTmp.Add colbaby(i)("_pati_id"), "_pati_id"
                colTmp.Add colbaby(i)("_pati_pageid"), "_pati_pageid"
                colTmp.Add colbaby(i)("_baby_num"), "_baby_num"
                colTmp.Add colbaby(i)("_baby_name"), "_baby_name"
                colTmp.Add colbaby(i)("_baby_sex"), "_baby_sex"
                colTmp.Add colbaby(i)("_baby_date"), "_baby_date"
                
                colOutListBaby.Add colTmp, "_" & colbaby(i)("_pati_id") & "_" & colbaby(i)("_pati_pageid") & "_" & colbaby(i)("_baby_num")
            Next
        End If
    Next
    
    zlSplitService_GetPatiPage = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_CallIsCloseAcc(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef inState As Integer) As Boolean
    '发药/退药调用服务：查询是否已结帐
    'strInput：费用来源|No
    'inState：返回结帐状态
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '服务出参list返回集合
    Dim varList As Variant   '集合元素
    
    On Error GoTo ErrHandle
    'Zl_Exsesvr_Getfeebalancestate
'  ---------------------------------------------------------------------------
'  --功能:根据单据号信息，获取单据对应的结帐状态
'  --入参：Json_In:格式
'  --input
'  --    query_mode  N 1 查询方式:0-门诊记帐;1-住院记帐
'  --    bill_nos  C 1 单据号
'  --出参: Json_Out,格式如下
'  -- output
'  --    code  C 1 应答码：0-失败；1-成功
'  --    message C 1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
'  --    state N 1 状态:-1-不存在记帐单据;0-未结帐;1-部分结帐;2-全部结帐
'
'  ---------------------------------------------------------------------------
    
    If strInput = "" Then Exit Function

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_mode", Split(strInput, "|")(0), 1)
    strJson = strJson & "," & GetJsonNodeString("bill_nos", Split(strInput, "|")(1), 0)

    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "Zl_Exsesvr_Getfeebalancestate"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_Exsesvr_Getfeebalancestate”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallIsCloseAcc = False
        Exit Function
    Else
        inState = Val(objServiceCall.GetJsonNodeValue("output.state"))
    End If

    zlSplitService_CallIsCloseAcc = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetExseBillByTime(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal colQueryCons As Collection, ByRef strBill_Out As String, ByRef strErrMsg_Out As String) As Boolean
    '按时间范围获取费用单据
    '入参：
    '   colQueryCons = 查询条件，成员(Key)：费用来源,开始时间,结束时间,执行部门IDS,不含执行部门IDS
    '                           其中，费用来源：0-不区分;1-门诊;2-住院
    '出参:
    '   strBill_Out = 单据信息，格式：单据1:NO,单据1:NO,...；其中，单据：24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料
    Dim StrJson_In As String
    Dim colOutlist As Collection, colTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    strBill_Out = "": strErrMsg_Out = ""
    'Zl_Exsesvr_Getbillbytime
    '  --功能：按时间范围获取费用单据
    '  --入参：json格式
    '  --  input
    '  --    query_type          N 0 查询方式:0-获取药品医嘱费用单据，1-获取卫材医嘱费用单据
    '  --    fee_source          N 1 费用来源:0-不区分;1-门诊;2-住院
    '  --    start_time          C 1 开始时间，格式：yyyy-mm-dd hh24:mi:ss
    '  --    end_time            C 1 结束时间，格式：yyyy-mm-dd hh24:mi:ss
    '  --    exe_deptids         C 0 执行部门ID，多个用逗英文号分隔
    '  --    excp_exe_deptids    C 0 不包含的执行部门ID，多个用逗英文号分隔
    '  --出参：json格式
    '  --  output
    '  --    code                C 1 应答码：0-失败；1-成功
    '  --    message             C 1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '  --    bill_nos            C 1 单据信息:格式：单据类型1:NO1,单据类型2:NO2,...
    '  --                            其中，单据类型: 1-收费处方;2-记帐单处方;3-记帐表处方
     
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("fee_source", colQueryCons("费用来源"), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("start_time", Format(colQueryCons("开始时间"), "yyyy-MM-dd HH:mm:ss"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("end_time", Format(colQueryCons("结束时间"), "yyyy-MM-dd HH:mm:ss"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("exe_deptids", colQueryCons("执行部门IDS"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("excp_exe_deptids", colQueryCons("不含执行部门IDS"), 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Getbillbytime", StrJson_In, "", "", lngMode, False, , , , True) = False Then Exit Function
    
    strBill_Out = objServiceCall.GetJsonNodeValue("output.bill_nos")
    
    If strBill_Out <> "" Then
        '单据类型转换：24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料
        strBill_Out = "," & strBill_Out
        strBill_Out = Replace(strBill_Out, ",1:", ",24:")
        strBill_Out = Replace(strBill_Out, ",2:", ",25:")
        strBill_Out = Replace(strBill_Out, ",3:", ",26:")
        strBill_Out = zlStr.TrimEx(strBill_Out, ",")
    End If
    
    zlSplitService_GetExseBillByTime = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

