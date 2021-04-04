Attribute VB_Name = "mdlService"
Option Explicit

Public Enum JSON_TYPE
    Json_Text = 0 '字符
    Json_num = 1 '数值
End Enum

Public Function OpenJson(ByVal strJson As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:设置一个Json串
'入参:strJson-Json串
'出参:
'返回:设置成功,返回true,否则返回False
'编制:余伟节
'日期:2019-08-11 19:36:34
'---------------------------------------------------------------------------------------------------------------------------------------------
    If InitSvr() Then
        OpenJson = gobjService.SetJsonString(strJson)
    End If
End Function

Public Function GetSvrOutInfo(ByVal strJson As String, Optional ByVal blnShowErrMsg As Boolean = True, Optional ByVal strErrMsg As String) As Boolean
'功能：针对本域【临床】检查方法调用的出参解析方式
'参数：strJson前一个检查过程执行后的的出参Json格式
'      blnShowErrMsg 是否内部弹出提示信息
'返回：true/false  code=1时为true,code=0时为false
    If InitSvr() Then
        Call gobjService.SetJsonString(strJson)
        If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
            If blnShowErrMsg Then
                MsgBox gobjService.GetJsonNodeValue("output.message") & "", vbInformation, gstrSysName
            Else
                strErrMsg = gobjService.GetJsonNodeValue("output.message") & ""
            End If
        Else
            GetSvrOutInfo = True
        End If
    End If
End Function

Public Function GetNode(ByVal strKey As String, ByVal varValue As Variant, Optional ByVal blnFirst As Boolean, _
Optional ByVal bytType As Byte) As String
'功能:获取单个JSON元素键值对
'   strKey-   Json key值
'   strValue- Json value值 数字\字符串
'   blnFirst  = T 第一个节点;F-非第一节点
'   bytType   =1 拼接字符"{}";=2 拼接字符"[]"
    Dim strTemp As String
    
    Select Case TypeName(varValue)
    
    Case "String"
        If Left(varValue, 1) = Chr(91) Or Left(varValue, 1) = Chr(123) Then
            strTemp = varValue
        Else
            strTemp = Chr(34) & zlStr.ToJsonStr(varValue) & Chr(34)
        End If
    Case "Empty"
        strTemp = "null"
    Case Else
        strTemp = varValue
    End Select
    strTemp = IIf(Not blnFirst, ",", "") & Chr(34) & LCase(strKey) & Chr(34) & ":" & strTemp
    If bytType = 1 Then strTemp = Chr(123) & strTemp & Chr(125)
    If bytType = 2 Then strTemp = Chr(91) & strTemp & Chr(93)
    GetNode = strTemp
End Function

Public Function GetJsonListValue(ByVal strListPathNode As String, Optional ByVal strKeyNodes As String, Optional ByVal varNullValue As Variant) As Collection
'功能：获取Json中的数组数据或子结点数据到集合中
'参数：
'  strList=Json数组结点或父结点名及路径，如：output，output.pati_list，output.pati_list[0].baby_list
'  strKeys=数组中作为关键字的结点名，可以多个用","号分隔，如"pati_id,pati_pageid"。注意关键字结点的数据不允许存在重复
'  varNullValue=当数组中的结点值为为null时，返回的转换值
    If InitSvr() Then
        Set GetJsonListValue = gobjService.GetJsonListValue(strListPathNode, strKeyNodes, varNullValue)
    End If
End Function

Public Function GetJsonNodeValue(ByVal strPathNode As String, Optional ByVal varNullValue As Variant) As Variant
'功能：获取Json指定结点的值
'参数：
'  strElement=结点及路径，如：output.message，output.pati_list[0].phone_number,output.num_list
'  varNullValue=当结点值为为null时，返回的转换值
    If InitSvr() Then
        GetJsonNodeValue = gobjService.GetJsonNodeValue(strPathNode, varNullValue)
    End If
End Function

Public Function InitSvr() As Boolean
'功能：初始化服务接口部件
    If gobjService Is Nothing Then
        On Error Resume Next
        Set gobjService = CreateObject("zlServiceCall.clsServiceCall")
        If Not gobjService.InitService(gcnOracle, gstrDBUser, glngSys) Then
            Set gobjService = Nothing
        End If
        Err.Clear: On Error GoTo 0
    End If
    If gobjService Is Nothing Then
        MsgBox "zlServiceCall.clsServiceCall创建失败!", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not gobjService Is Nothing Then InitSvr = True
End Function

Public Function CallService(ByVal strServiceName As String, _
    ByVal strJson_In As String, Optional ByRef strJson_out As String, Optional ByVal strTitle As String, _
    Optional lngModule As Long, Optional blnShowErrMsg As Boolean = True, Optional ByVal strAskDate As String, _
    Optional varExpend As String, Optional lngSys As Long, Optional blnReadServiceErr As Boolean) As Boolean
'功能：调用服务
'相关说明见 zlServiceCall.clsServiceCall.CallService 接口
    If InitSvr() Then
        If lngModule = 0 Then lngModule = glngModul
        If Not gobjService.CallService(strServiceName, strJson_In, strJson_out, strTitle, lngModule, blnShowErrMsg, strAskDate, varExpend, lngSys, blnReadServiceErr) Then Exit Function
        If Not blnShowErrMsg Then
            If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
                varExpend = gobjService.GetJsonNodeValue("output.message")
                CallService = False: Exit Function
            End If
        End If
        CallService = True
    End If
End Function

Public Function GetJsonSurety(ByVal bytFunc As Byte, ByVal lngPatiID As Long, ByVal lngPageID As Long, Optional ByVal strGuarantor As String, _
   Optional ByVal dblAmount As Double, Optional ByVal bytType As Byte, Optional ByVal strReason As String, Optional ByVal strDueTime As String, _
   Optional ByVal strCreateTime As String) As String
'功能:获取担保JSON串
'参数:
'      bytFunc          功能ID 1-新增;2-更新;3-删除
'      lngPatiId        病人id
'      lngPageId        主页ID
'      strGuarantor     担保人
'      dblAmount        担保额
'      bytType          担保性质
'      strReason        担保原因
'      strDueTime       到期时间   格式化 "yyyy-MM-dd HH:mm:ss"
'      strCreateTime    登记时间   更新或删除时传入
    Dim strIn As String
    
    strIn = GetNode("func_id", bytFunc, True)
    strIn = strIn & GetNode("pati_id", lngPatiID)
    strIn = strIn & GetNode("pati_pageid", lngPageID)
    If bytFunc <> 3 Then
        strIn = strIn & GetNode("guarantor", strGuarantor)
        strIn = strIn & GetNode("garnt_amount", dblAmount)
        strIn = strIn & GetNode("garnt_prop", bytType)
        strIn = strIn & GetNode("garnt_reason", strReason)
        If strDueTime <> "" Then
            strIn = strIn & GetNode("due_time", strDueTime)
        End If
    End If

    strIn = strIn & GetNode("operator_code", UserInfo.编号)
    strIn = strIn & GetNode("operator_name", UserInfo.姓名)

    If bytFunc <> 1 Then
        strIn = strIn & GetNode("create_time", strCreateTime)
    End If
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    GetJsonSurety = strIn
End Function


'---------------------------------------------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------------------------------------------
Public Function ZL_PatiSvr_GetPatiId(ByRef colPati As Collection, ByVal strFindName As String, ByVal strFindValue As String, Optional ByVal lngPatiID As Long) As Boolean
'功能:获取病人ID
'参数:
'       _pati_id 病人ID
'       _pati_pageid 主页ID
    Dim strIn As String
    Dim strOut As String
    
    strIn = GetNode("find_name", strFindName, True)
    strIn = strIn & GetNode("find_text", strFindValue)
    If lngPatiID > 0 Then strIn = strIn & GetNode("pati_id", lngPatiID)
    strIn = GetNode("other_cons_find", "{" & strIn & "}", True, 1)
    strIn = GetNode("input", strIn, True, 1)
    If Not CallService("Zl_Patisvr_Getpatiid", strIn, strOut, "获取病人ID", P病人入院管理) Then Exit Function
    Set colPati = GetJsonListValue("output.pati_list", "pati_id")
    If colPati.Count = 0 Then
        colPati.Add 0, "_pati_id"
    Else
        Set colPati = colPati(1)
    End If
    ZL_PatiSvr_GetPatiId = True
End Function

Public Function Zl_PatiSvr_GetNextId(ByVal strTableName As String, ByVal strColName As String) As Long
'  -----------------------------------------
'  --功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值
'  --入参：Json_In:格式
'  --input
'  --  table_name    C  1 表名
'  --  col_name      C  1 字段名  序列名称不一定是ID，例如记录ID
'  -- 出参:
'  --  output
'  --  next_id      N   1  序列
'  -------------------------------------------
    Dim strIn As String
    Dim strOut As String
    
    strIn = GetNode("table_name", strTableName, True)
    strIn = strIn & GetNode("col_name", strColName)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    
    If Not CallService("Zl_PatiSvr_GetNextId", strIn, strOut, "读取指定表名对应的序列", P病人入院管理) Then Exit Function
    Zl_PatiSvr_GetNextId = GetJsonNodeValue("output.next_id")
End Function

Public Function zl_PatiSvr_GetNextNo(ByVal lngItemNum As Long, Optional ByVal lngDeptId As Long) As String
'  ---------------------------------------------------------------------------
'  --功能：功能：根据特定规则产生新的号码
'  lngItemNum:1  病人身份ID ;2  病人住院号;3  病人门诊号
'  --入参：Json_In:格式
'  --  input
'  --    item_num            N   1   项目序号
'  --    dept_id             N   0   科室ID
'  --出参: Json_Out,格式如下
'  --  output
'  --    code                N   1   应答码：0-失败；1-成功
'  --    message             C   1   应答消息：失败时返回具体的错误信息
'  --    next_no             C   1   下一个号码
'  ---------------------------------------------------------------------------
    Dim strIn As String
    Dim strOut As String
    
    strIn = GetNode("item_num", lngItemNum, True)
    strIn = strIn & GetNode("dept_id", lngDeptId)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    
    If Not CallService("zl_PatiSvr_GetNextNo", strIn, strOut, "根据特定规则产生新的号码", P病人入院管理) Then Exit Function
    zl_PatiSvr_GetNextNo = GetJsonNodeValue("output.next_no")
End Function

Public Function zl_PatiSvr_GetPatiInfo(ByVal lngPatiID As Long, Optional ByVal bytQueryType As Byte, Optional ByVal bytCard As Byte, _
    Optional ByVal bytFamily As Byte, Optional ByVal bytDrug As Byte, Optional ByVal bytImmune As Byte, _
    Optional ByVal strPatiIDs As String, Optional ByVal strPatiName As String, Optional strOutNum As String, _
    Optional strIdCard As String, Optional strContactId As String, Optional dblCardTypeId As Double, _
    Optional strMedcCardName As String, Optional strCardNO As String, Optional strQRcode As String, _
    Optional strICCardNo As String, Optional strVisitCard As String, Optional strInsuranceNum As String, _
    Optional intStatu As Integer = -1, Optional strPhoneNumber As String, Optional strBed As String) As Boolean
'功能:获取病人信息
'参数:
'     lngPatiID           N 1 病人id  病人ID<>0时，查询列表中的条件无效
'     bytQueryType        N 1 查询类型:如：0-基本;1-基本+联系人;2-所有
'     bytCard           N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
'     bytFamily         N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
'     bytDrug           N 1 是否包含过敏药物:1-包含，0-不包含
'     bytImmune         N 1 是否包含免疫修:1-包含;0-不包含
'       strPatiIds          C   病人IDs:多个用逗号
'       strPatiName         C   姓名:可以代%分号表表按姓名匹配
'       dblOutNum           N   门诊号
'       strIdCard           C   身份证号
'       strContactId        C   联系人身份证号
'       dblCardTypeId       N   医疗卡类别ID
'       strMedcCardName     C   医疗卡名称
'       strCardNo           C   卡号
'       strQRcode           C   二维码
'       strICCardNo         C   Ic卡号
'       strVisitCard        C   就诊卡号
'       strInsuranceNum     C   医保号
'       intStatu            C   查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
'       strPhoneNumber      C   手机号
'       strBed              C   当前床号
    Dim strIn As String
    Dim strOut As String
    Dim strTemp As String

    On Error GoTo ErrH

    strIn = GetNode("pati_id", lngPatiID, True)
    If bytQueryType > 0 Then strIn = strIn & GetNode("query_type", bytQueryType)
    If bytCard > 0 Then strIn = strIn & GetNode("query_card", bytCard)
    If bytFamily > 0 Then strIn = strIn & GetNode("query_family", bytFamily)
    If bytDrug > 0 Then strIn = strIn & GetNode("query_drug", bytDrug)
    If bytImmune > 0 Then strIn = strIn & GetNode("query_immune", bytImmune)
    If bytFamily > 0 Then strIn = strIn & GetNode("query_family", bytFamily)
    If lngPatiID = 0 Then
        If strPatiIDs <> "" Then strTemp = strTemp & GetNode("pati_ids", strPatiIDs)
        If strPatiName <> "" Then strTemp = strTemp & GetNode("pati_name", strPatiName)
        If strOutNum <> "" Then strTemp = strTemp & GetNode("outpatient_num", strOutNum)
        If strIdCard <> "" Then strTemp = strTemp & GetNode("pati_idcard", strIdCard)
    
        If strContactId <> "" Then strTemp = strTemp & GetNode("contacts_idcard", strContactId)
        If dblCardTypeId > 0 Then strTemp = strTemp & GetNode("cardtype_id", dblCardTypeId)
        If strMedcCardName <> "" Then strTemp = strTemp & GetNode("medc_card_name", strMedcCardName)
        If strCardNO <> "" Then strTemp = strTemp & GetNode("card_no", strCardNO)
        If strQRcode <> "" Then strTemp = strTemp & GetNode("qrcode", strQRcode)
        If strICCardNo <> "" Then strTemp = strTemp & GetNode("iccard_no", strICCardNo)
    
        If strVisitCard <> "" Then strTemp = strTemp & GetNode("visit_card", strVisitCard)
        If strInsuranceNum <> "" Then strTemp = strTemp & GetNode("insurance_num", strInsuranceNum)
        If intStatu >= 0 Then strTemp = strTemp & GetNode("qrspt_statu", intStatu)
        If strPhoneNumber <> "" Then strTemp = strTemp & GetNode("phone_number", strPhoneNumber)
        If strBed <> "" Then strTemp = strTemp & GetNode("pati_bed", strBed)
    
        If strTemp <> "" Then
            strIn = strIn & GetNode("query_cons_list", "{" & Mid(strTemp, 2) & "}")
        End If
    End If
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    zl_PatiSvr_GetPatiInfo = CallService("Zl_Patisvr_Getpatiinfo", strIn, strOut, "获取病人基本信息", P病人入院管理, False, , , , True)

    Exit Function

ErrH:
    MsgBox "在zl9InPatient.mdlService.zl_PatiSvr_GetPatiInfo的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description, vbExclamation, gstrSysName

End Function

Public Function Zl_Patisvr_Getpatiaddrssinfo(ByVal lngPatiID As Long, ByVal lngPageID As Long, Optional ByVal bytType As Byte) As Boolean
'  --功能：根据病人ID,获取病人的地址信息
'  --入参：Json_In:格式
'  --  input
'  --    pati_id              N   1  病人ID
'  --    pati_pageid          N      主页id
'  --    addr_type            N      地址类别:1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址；为0时表示查询所有类型的地址信息
'  --出参: Json_Out,格式如下
'  --  output
'  --    code                 N   1   应答吗：0-失败；1-成功
'  --    message              C   1   应答消息：失败时返回具体的错误信息
'  --    addr_list[]          C       地址列表信息
'  --      pat_addr_type      C   1   地址类别
'  --      pat_addr_state     C   1   地址_省
'  --      pat_addr_city      C   1   地址_市
'  --      pat_addr_county    C   1   地址_县
'  --      pat_addr_township  C   1   地址_乡
'  --      pat_addr_other     C   1   地址_其他
'  --      pat_region_code    C   1   区划代码
    Dim strIn As String
    Dim strOut As String
        
    strIn = GetNode("pati_id", lngPatiID, True)
    strIn = strIn & GetNode("pati_pageid", lngPageID)
    strIn = strIn & GetNode("addr_type", bytType)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    Zl_Patisvr_Getpatiaddrssinfo = CallService("Zl_Patisvr_Getpatiaddrssinfo", strIn, strOut, "获取结构化地址", P病人入院管理)
End Function
 

  
'----------------------------------------------------------------------------------------------------------------------------
'------费用相关服务
'----------------------------------------------------------------------------------------------------------------------------
 
Public Function Zl_Exsesvr_CheckPatiChangeUndo(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal strUndoType As String, _
    ByVal strBeginTime As String, ByVal lngFeeItemID As Long) As Boolean
'功能:撤销病人变动记录前检查
'参数：
'   strUndoType-撤销方式
'   strBeginTime-开始时间
'   lngFeeItemID-费用项目ID
    Dim strIn As String
    Dim strOut As String
    
    strIn = GetNode("pati_id", lngPatiID, True)
    strIn = strIn & GetNode("pati_pageid", lngPageID)
    strIn = strIn & GetNode("undo_type", strUndoType)
    strIn = strIn & GetNode("create_time", strBeginTime)
    strIn = strIn & GetNode("fee_item_id", lngFeeItemID)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    Zl_Exsesvr_CheckPatiChangeUndo = CallService("Zl_Exsesvr_Checkpatichangeundo", strIn, strOut, "撤销病人变动检查", P病人入院管理)
End Function

Public Function zl_ExseSvr_GetInsuranceDisease(ByVal lng病种ID As Long) As Boolean
'功能:获取保险病种名
    Dim strIn As String
    Dim strOut As String
    
    strIn = GetNode("si_dz_id", lng病种ID, True)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    zl_ExseSvr_GetInsuranceDisease = CallService("Zl_Exsesvr_Getinsurancedisease", strIn, strOut, "获取保险病种", P病人入院管理)
End Function

Public Function Zl_Exsesvr_Getpatisurplusinfo(ByVal varPatiId As Variant, ByVal bytFunc As Byte, _
    ByRef colFee As Collection, Optional ByVal lngPageID As Long, Optional ByVal lngModel As Long) As Boolean
'参数:
'    bytFunc= 1-获取病人费用余额;2-获取病人某一次住院的未结费用;3-批量获取病人费用余额
'出参：
'    colFee

    Dim strIn As String
    Dim strOut As String
    Dim colList As Collection
    '  output
    '    code              N 1 应答码：0-失败；1-成功
    '    message           C 1 应答消息：失败时返回具体的错误信息
    '    infee_surplus     N 1 住院费用余额
    '    surplus_list[]    C 1 余额列表
    '      pati_Id         N   病人ID
    '      outdpst_surplus N 1 门诊预交余额
    '      indpst_surplus  N 1 住院预交余额
    '      outfee_surplus  N 1 门诊费用余额
    '      infee_surplus   N 1 住院费用余额
    ' 获取病人是否结清
    On Error GoTo ErrH
    
    Set colFee = New Collection
    If bytFunc = 2 Then
        strIn = GetNode("pati_id", CLng(varPatiId), True)
        strIn = strIn & GetNode("pati_pageid", lngPageID)
    Else
        strIn = GetNode("pati_ids", CStr(varPatiId), True)
    End If
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    If Not CallService("Zl_Exsesvr_Getpatisurplusinfo", strIn, strOut, "获取病人费用", lngModel) Then Exit Function
    If bytFunc = 1 Then
        Set colList = GetJsonListValue("output.surplus_list[0]")
        If Not colList Is Nothing Then
            If colList.Count > 0 Then
                colFee.Add colList("_outfee_surplus"), "门诊费用余额"
                colFee.Add colList("_infee_surplus"), "住院费用余额"
                colFee.Add colList("_indpst_surplus"), "住院预交余额"
            End If
        End If
        If colFee.Count = 0 Then
            colFee.Add 0, "门诊费用余额"
            colFee.Add 0, "住院费用余额"
            colFee.Add 0, "住院预交余额"
        End If
    ElseIf bytFunc = 2 Then
        colFee.Add Val(GetJsonNodeValue("output.infee_surplus") & ""), "住院费用余额"
    Else
        Set colFee = GetJsonListValue("output.surplus_list", "pati_id")
    End If
    Zl_Exsesvr_Getpatisurplusinfo = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Zl_Exsesvr_GetReceiveInvoice(ByVal bytKind As Long, ByRef colList As Collection, Optional ByVal bytFunc As Byte, _
    Optional ByVal strReceiver As String, Optional ByVal bytUseMode As Byte, Optional ByVal strUseType As String, _
    Optional ByVal strRecvIds As String) As Boolean
'参数:
'   bytFunc= 0-获取票据领用信息 1-获取获取指定票种的共用票据批次
'   strReceiver-领用人
'   strRecvIds-领用ids:票据领用id,多个用逗号
'  ---------------------------------------------------------------------------
'  --功能:获取票据领用信息
'  --入参：Json_In:格式
'  --    input
'  --      oper_fun  N 1 0-获取票据领用信息 1-获取获取指定票种的共用票据批次
'  --      recv_ids C 1 领用ids:票据领用id,多个用逗号
'  --      inv_type  N 1 票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
'  --      use_mode  N 1 使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
'  --      use_type C 1 票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
'  --      recvtr  C 1 领用人
'  --      min_nums  N 1 发票最少数量
'  --      nodeno  C  1  站点
'  --出参: Json_Out,格式如下
'  --    output
'  --    code  C 1 应答码：0-失败；1-成功
'  --    message C 1 "应答消息： 成功时返回成功信息,失败时返回具体的错误信息"
'  --    item_list C
'  --      recv_id N 1 领用ID
'  --      use_mode  N 1 使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
'  --      use_type C 1 票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
'  --      prefix_text C 1 前缀文本
'  --      start_no  C 1 开始号码
'  --      end_no  C 1 终止号码
'  --      inv_no_cur  C 1 当前号码
'  --      surplus_num C 1 剩余数量
'  --      create_time C 1 登记时间:yyyy-mm-dd hh24:mi:ss
'  --      use_time  C 1 使用时间:yyyy-mm-dd hh24:mi:ss
'  --      recvtr  C 1 领用人
'  --      use_typecode      C 1 使用类别编码
'  --      use_typeid        N 1 使用类别id
'
'  ---------------------------------------------------------------------------
    Dim strIn As String
    Dim strOut As String

    On Error GoTo ErrH

    strIn = GetNode("oper_fun", bytFunc, True)
    strIn = strIn & GetNode("inv_type", bytKind)
    If bytFunc = 0 Then
        If strRecvIds <> "" Then strIn = strIn & GetNode("recv_ids", strRecvIds)
        strIn = strIn & GetNode("recvtr", strReceiver)
        strIn = strIn & GetNode("use_mode", bytUseMode)
        strIn = strIn & GetNode("use_type", strUseType)
    End If
    strIn = strIn & GetNode("nodeno", gstrNodeNo)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    If Not CallService("Zl_Exsesvr_GetReceiveInvoice", strIn, strOut, "获取票据领用信息") Then Exit Function
    Set colList = GetJsonListValue("output.item_list")
    Zl_Exsesvr_GetReceiveInvoice = True
    Exit Function
ErrH:
   If ErrCenter() = 1 Then Resume
   Call SaveErrLog
End Function

Public Function zl_PatiSvr_GetPatiInfoEx(ByVal lng病人ID As Long, _
    ByVal colOtherFindCons As Collection, ByRef colPatiDatas_Out As Collection, _
    Optional ByVal int查询类型 As Integer = 0, _
    Optional ByVal bln包含家属 As Boolean, _
    Optional ByVal bln包含过敏药物 As Boolean, _
    Optional ByVal bln包含免疫信息 As Boolean, _
    Optional ByVal bln包含卡信息 As Boolean, Optional ByVal blnNotShowErrMsg As Boolean = True, _
    Optional ByVal bln是否密文显示 As Boolean, _
    Optional ByRef strErrMsg As String, Optional ByVal strKeyNodes As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人详细信息服务接口
    '入参:colOtherFindCons-其他查找条件(array(查询名称,查询值)
    '             查询名称:病人IDS,姓名,性别,出生日期等,见query_cons_list[]列表中的描述部分
    '      int查询类型-0-基本;1-基本+联系人;2-所有
    '      strKeyNodes= 指定返回集合索引方式
    '出参:colPatiDatas_Out-返回病人信息集
    '返回:成功返回true,否则返回False
    '编制:YWJ
    '日期:2019-10-31 18:02:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim colData As Variant
    Dim intReturn As Integer, strJsonTemp As String
    
    
    On Error GoTo errHandle
    
    If Not colOtherFindCons Is Nothing Then
        For i = 1 To colOtherFindCons.Count
            Select Case UCase(colOtherFindCons(i)(0))
            Case "病人IDS"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_ids", colOtherFindCons(i)(1), Json_Text)
            Case "姓名"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_name", colOtherFindCons(i)(1), Json_Text)
            Case "门诊号"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("outpatient_num", Val(colOtherFindCons(i)(1)), Json_Text, True)
            Case "住院号"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("inpatient_num", Val(colOtherFindCons(i)(1)), Json_Text, True)
            Case "身份证号", "二代身份证", "身份证"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_idcard", colOtherFindCons(i)(1), Json_Text)
            Case "联系人身份证"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("contacts_idcard", colOtherFindCons(i)(1), Json_Text)
            Case "医保号", "医保证号"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_num", colOtherFindCons(i)(1), Json_Text)
            Case "医疗卡类别ID", "卡类别ID"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(colOtherFindCons(i)(1)), Json_num, True)
            Case "卡号"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("card_no", colOtherFindCons(i)(1), Json_Text)
            Case "二维码"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("qrcode", colOtherFindCons(i)(1), Json_Text)
            Case "IC卡号", "IC", "IC卡"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("iccard_no", colOtherFindCons(i)(1), Json_Text)
            Case "查询住院状态", "住院状态"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("qrspt_statu", Val(colOtherFindCons(i)(1)), Json_num, True)
            Case "手机号", "手机"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("phone_number", colOtherFindCons(i)(1), Json_Text)
            Case "就诊卡号", "就诊卡"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("vcard_no", colOtherFindCons(i)(1), Json_Text)
            Case "查找天数"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("search_days", colOtherFindCons(i)(1), Json_num, True)
            Case Else
                strErrMsg = "目前暂不不支持按类别为【" & UCase(colOtherFindCons(i)(0)) & "】来查找病人！"
                If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
            End Select
        Next
        If strJsonTemp <> "" Then strJsonTemp = Mid(strJsonTemp, 2)
    End If
    If lng病人ID = 0 And strJsonTemp = "" Then
        strErrMsg = "无有效的查询条件，请检查！"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '其他条件
    'zl_PatiSvr_GetPatiInfo
    'input
    '    pati_id N   1   病人id：病人ID<>0时，查询列表中的条件无效
    '    query_type  N   1   查询类型:如：0-基本;1-基本+联系人;2-所有
    '    query_card  N   1   是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
    '    query_family    N   1   是否包含家属:1-包含家属信息，0-不包含家属信息
    '    query_drug  N   1   是否包含过敏药物:1-包含，0-不包含
    '    query_immune    N   1   是否包含免疫信息:1-包含;0-不包含
    '    query_cons_list[]   C   1   查询条件:可以选择一定条件进行查询（是And关系),只有一行
    '        pati_ids    C       病人IDs:多个用逗号
    '        pati_name   C       姓名:可以代%分号表表按姓名匹配
    '        outpatient_num  N       门诊号
    '        pati_idcard C       身份证号
    '        contacts_idcard C       联系人身份证号
    '        cardtype_id N       医疗卡类别ID
    '        card_no C       卡号
    '        qrcode  C       二维码
    '        iccard_no   C       Ic卡号
    '        insurance_num   C       医保号
    '        qrspt_statu C       查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
    '        phone_number    C       手机号
       
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_type", int查询类型, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_family", IIf(bln包含家属, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_drug", IIf(bln包含过敏药物, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_immune", IIf(bln包含免疫信息, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_card", IIf(bln包含卡信息, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("passshowcard", IIf(bln是否密文显示, 1, 0), Json_num, True)
    strJson = strJson & "," & GetNodeString("query_cons_list") & ":{" & strJsonTemp & "}"
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetPatiInfo"
    If CallService(strServiceName, strJson, , "", , False) = False Then Exit Function
    
    
    '出参            json    基本    基本+联系人 所有
    'output
    '    code    N   1   应答码：0-失败；1-成功  √  √  √
    '    message C   1   应答消息： 失败时返回具体的错误信息 √  √  √
    '    pati_list[]         病人信息列表    √  √  √
    '    pati_id N   1   病人id  √  √  √
    '    pati_pageid N   1   主页id：病人信息.主页ID √  √  √
    '    pati_name   C   1   姓名    √  √  √
    '    pati_sex    C   1   性别    √  √  √
    '    pati_age    C   1   年龄    √  √  √
    '    pati_birthdate  C   1   出生日期：yyyy-mm-dd hh24:mi:ss √  √  √
    '    fee_category    C   1   费别    √  √  √
    '    outpatient_num  C   1   门诊号  √  √  √
    '    inpatient_num   C   1   住院号
    '    mdlpay_mode_name    C   1   医疗付款方式名称    √  √  √
    '    mdlpay_mode_code    C   1   医疗付款方式编码    √  √  √
    '    pati_nation C   1   民族    √  √
    '    insurance_num   C   1   医保号  √  √  √
    '    pati_idcard C   1   身份证号    √  √  √
    '    vcard_no    C   1   就诊卡号            √
    '    iccard_no   C   1   Ic卡号          √
    '    health_num  C   1   健康号          √
    '    pati_education  C   1   学历            √
    '    ocpt_name   C   1   职业            √
    '    pati_identity   C   1   身份            √
    '    ntvplc_name C   1   籍贯            √
    '    country_name    C   1   国籍            √
    '    pati_marital_cstatus    C   1   婚姻状况            √
    '    pat_home_addr   C   1   家庭地址    √  √  √
    '    pat_home_phno   C   1   家庭电话    √  √  √
    '    pat_home_postcode   C   1   家庭地址邮编            √
    '    pati_area   C   1   区域            √
    '    pati_birthplace C   1   出生地点    √  √  √
    '    pat_hous_addr   C   1   户口地址            √
    '    pat_hous_postcode   C   1   户口地址邮编            √
    '    emp_name    C   1   工作单位名称            √
    '    emp_phno    C   1   单位电话            √
    '    emp_postcode    C   1   单位邮编            √
    '    emp_bank_name   C   1   单位开户行          √
    '    ctt_unit_id N   1   合同单位ID          √
    '    phone_number    C   1   手机号  √  √  √
    '    pati_bed    C   1   当前床号    √  √  √
    '    pati_type   C   1   病人类型(普通，医保，留观)          √
    '    balance_mode    N   1   结算模式(0-收费，1-记账)            √
    '    insurance_type  C   1   险类    √  √  √
    '    pati_wardarea_id    N   1   当前病区id          √
    '    pati_wardarea_name  C   1   当前病区名称            √
    '    pati_dept_id    N   1   当前科室id          √
    '    pati_dept_name  C   1   当前科室名称            √
    '    adta_time   C   1   入院时间:yyyy-mm-dd hh24:mi:ss          √
    '    adtd_time   C   1   出院时间:yyyy-mm-dd hh24:mi:ss          √
    '    contacts_name   C   1   联系人姓名      √  √
    '    contacts_relation   C   1   联系人关系      √  √
    '    contacts_idcard C   1   联系人身份证号      √  √
    '    contacts_addr   C   1   联系人地址      √  √
    '    contacts_phno   C   1   联系人电话      √  √
    '    pat_grdn_name   C   1   监护人          √
    '    cert_no_other   C   1   其他证件            √
    '    is_inhspt   C   1   是否在院:1-在院 ;0-不在院   √  √  √
    '    pati_show_color N   1   病人显示颜色            √
    '    visit_room  C   1   就诊诊室            √
    '    visit_statu N   1   就诊状态            √
    '    visit_time  C   1   就诊时间:yyyy-mm-dd hh24:mi:ss          √
    '    create_time C   1   登记时间:yyyy-mm-dd hh24:mi:ss          √
    '    family_list[]   C   1   家属成员:病人家属() query_family=1返回
    '        family_id   N   1   家属id  query_family=1
    '        family_relation C   1   关系
    '    drug_list[] C   1   抗菌药物列表    query_drug=1时返回
    '        pat_algc_cadn_id    N   1   过敏药品ID
    '        pat_algc_cadn   C   1   过敏药物名称
    '        allergy_info    C   1   过每药物反应
    '    immune_list[]   C   1   病人免疫列表    query_immune=1时返回
    '        vaccinate_time  C   1   接种时间:yyyy-mm-dd hh24:mi:ss
    '        vaccinate_name  C   1   接种名称
    '    card_list[] C   1   病人医疗卡信息列表(如果条件中传入了卡类别ID的，则返回该卡类别的卡信息)  query_card=1时返回
    '        cardtype_id N   1   医疗卡类别ID
    '        card_no C   1   卡号
    '        card_pwd    C   1   密码
    intReturn = gobjService.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg = gobjService.GetJsonNodeValue("output.message")
        If strErrMsg = "" Then
            strErrMsg = "未找到符合条件的病人信息，请检查！"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set colData = gobjService.GetJsonListValue("output.pati_list", strKeyNodes)
    
    If colData Is Nothing Then
        strErrMsg = "未找到符合条件的病人信息，请检查！"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If colData.Count = 0 Then
        strErrMsg = "未找到符合条件的病人信息，请检查！"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set colPatiDatas_Out = colData
    zl_PatiSvr_GetPatiInfoEx = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Patisvr_GetPatiExtendInfo(ByVal lngPatiID As Long, colPatiEx As Collection, Optional ByVal strInfoName As String, Optional ByVal lngVisitID As Long) As Boolean
'  ---------------------------------------------------------------------------
'  --功能：获取病人信息从表
'  --入参：Json_In:格式
'  --  input
'  --    pati_id             N 1 病人id
'  --    info_names          C 1 信息名：多个用逗号
'  --    visit_id            N  就诊id
'  --出参: Json_Out,格式如下
'  --  output
'  --    code                N 1   应答码：0-失败；1-成功
'  --    message             C 1   应答消息：失败时返回具体的错误信息
'  --    slave_list[]        C     病人信息从表列表
'  --     info_name          C 1   信息名
'  --     info_value         N 1   信息值
'  --     visit_id           N 1   就诊id
'  ---------------------------------------------------------------------------
    Dim strIn As String
    Dim strOut As String
    On Error GoTo ErrH

    strIn = GetNode("pati_id", lngPatiID, True)
    strIn = strIn & GetNode("info_names", strInfoName)
    strIn = strIn & GetNode("visit_id", lngVisitID)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    If Not CallService("Zl_Patisvr_GetPatiExtendInfo", strIn, strOut, "获取病人从表信息") Then Exit Function
    Set colPatiEx = GetJsonListValue("output.slave_list", "info_name")
    Zl_Patisvr_GetPatiExtendInfo = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Zl_Patisvr_GetpatAllergicDrugs(ByVal lngPatiID As Long, colDrug As Collection) As Boolean
'---------------------------------------------------------------------------
'--功能:获取病人信息的过敏药物信息
'--入参：Json_In:格式
'--  input
'--    pati_id             N   1 病人id
'--出参: Json_Out,格式如下
'--  output
'--    code                N   1   应答码：0-失败；1-成功
'--    message             C   1   应答消息：失败时返回具体的错误信息
'--    drug_list[]         C       过敏药物列表
'--      medicinal_id      N   1   过敏药品ID
'--      medicinal_name    C   1   过敏药物名称
'--      allergy_info      C   1   过每药物反应
'---------------------------------------------------------------------------
    Dim strIn As String
    Dim strOut As String
    On Error GoTo ErrH

    strIn = GetNode("pati_id", lngPatiID, True)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    If Not CallService("Zl_Patisvr_GetpatAllergicDrugs", strIn, strOut, "获取病人过敏药物") Then Exit Function
    Set colDrug = GetJsonListValue("drug_list")
    Zl_Patisvr_GetpatAllergicDrugs = True
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Zl_Patisvr_GetPatImmuneInfo(ByVal lngPatiID As Long, colDrug As Collection) As Boolean
'  ---------------------------------------------------------------------------
'  --功能:获取病人信息的免疫信息
'  --入参：Json_In:格式
'  --  input
'  --    pati_id           N   1 病人id
'  --出参: Json_Out,格式如下
'  --  output
'  --    code              N   1   应答码：0-失败；1-成功
'  --    message           C   1   应答消息：失败时返回具体的错误信息
'  --    immune_list[]     C       病人免疫列表
'  --      vaccinate_time    C   1   接种时间:yyyy-mm-dd hh24:mi:ss
'  --      vaccinate_name    C   1   接种名称
'  ---------------------------------------------------------------------------
    Dim strIn As String
    Dim strOut As String
    On Error GoTo ErrH

    strIn = GetNode("pati_id", lngPatiID, True)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    If Not CallService("Zl_Patisvr_GetPatImmuneInfo", strIn, strOut, "获取病人免疫信息") Then Exit Function
    Set colDrug = GetJsonListValue("immune_list")
    Zl_Patisvr_GetPatImmuneInfo = True
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetColVal(ByVal colData As Collection, ByVal strKey As String, Optional ByVal strType As String, Optional ByVal strDef As String, Optional ByRef lngExist As Long) As String
'功能:通集合关键字获取集合的值,基本数据类型,数字或字符
'入参：strType  N/n  表示数字类型，c表示字符串
'      strDef  缺省值，当出错时以这个值为缺省值返回
'出参:lngExist 集合中是否存在这个结点值,0-存在,-1不存在
    Dim strValue As String
    
    On Error GoTo ErrH
    
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
ErrH:
    Err.Clear
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

Public Function GetJsonNodeString(ByVal strNodeName As String, ByVal strValue As String, _
    Optional ByVal intType As JSON_TYPE, Optional ByVal blnZeroToNull As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取Json接点串
    '入参:strNodeName-接点名
    '     strValue-值
    '     intType-类型:0-字符;1-数字
    '     blnZeroToEmpty-是否将数值0转换为Null，仅类型为数字时有效
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 18:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    strJson = Chr(34) & strNodeName & Chr(34)
    If intType = Json_Text Then
        strJson = strJson & ":" & Chr(34) & strValue & Chr(34)
    Else
        If strValue = "" Or (blnZeroToNull And Val(strValue) = 0) Then
            strJson = strJson & ":null"
        Else
            strJson = strJson & ":" & IIf(Mid(strValue, 1, 1) = ".", "0", "") & strValue
        End If
    End If
    GetJsonNodeString = strJson
End Function

Public Function GetNodeString(ByVal strNodeName As String) As String
    GetNodeString = Chr(34) & strNodeName & Chr(34)
End Function

Public Function GetMapPait(ByVal strName As String) As String
    Dim strValue As String
    '--    pati_id             N   1   病人id
    '--    pati_pageid         N   1   主页id：病人信息.主页ID
    '--    pati_name           C   1   姓名
    '--    pati_sex            C   1   性别
    '--    pati_age            C   1   年龄
    '--    pati_birthdate      C   1   出生日期：yyyy-mm-dd hh24:mi:ss
    '--    fee_category        C   1   费别
    '--    outpatient_num      C   1   门诊号
    '--    inpatient_num       C   1   住院号
    '--    mdlpay_mode_name    C   1   医疗付款方式名称
    '--    mdlpay_mode_code    C   1   医疗付款方式编码
    '--    pati_nation         C   1   民族
    '--    insurance_num       C   1   医保号
    '--    pati_idcard         C   1   身份证号
    '--    vcard_no            C   1   就诊卡号
    '--    iccard_no           C   1   Ic卡号
    '--    health_num          C   1   健康号
    '--    inp_times           N   1   住院次数
    '--    pati_education      C   1   学历
    '--    ocpt_name           C   1   职业
    '--    pati_identity       C   1   身份
    '--    ntvplc_name         C   1   籍贯
    '--    country_name        C   1   国籍
    '--    pati_marital_cstatus    C   1   婚姻状况
    '--    pat_home_addr           C   1   家庭地址
    '--    pat_home_phno           C   1   家庭电话
    '--    pat_home_postcode   C   1   家庭地址邮编
    '--    pati_area           C   1   区域
    '--    pati_birthplace     C   1   出生地点
    '--    pat_hous_addr       C   1   户口地址
    '--    pat_hous_postcode   C   1   户口地址邮编
    '--    emp_name            C   1   工作单位名称
    '--    emp_phno            C   1   单位电话
    '--    emp_postcode        C   1   单位邮编
    '--    emp_bank_name       C   1   单位开户行
    '--    emp_bank_accnum     C   1   单位帐号
    '--    emp_addr             C   1   单位地址
    '--    ctt_unit_id         N   1   合同单位ID
    '--    phone_number        C   1   手机号
    '--    pati_bed            C   1   当前床号
    '--    pati_type           C   1   病人类型(普通，医保，留观)
    '--    insurance_type      C   1   险类
    '--    insurance_name      C   1   险类名称
    '--    pati_wardarea_id    N   1   当前病区id
    '--    pati_wardarea_name  C   1   当前病区名称
    '--    pati_dept_id        N   1   当前科室id
    '--    pati_dept_name      C   1   当前科室名称
    '--    adta_time           C   1   入院时间:yyyy-mm-dd hh24:mi:ss
    '--    adtd_time           C   1   出院时间:yyyy-mm-dd hh24:mi:ss
    '--    contacts_name       C   1   联系人姓名
    '--    contacts_relation   C   1   联系人关系
    '--    contacts_idcard     C   1   联系人身份证号
    '--    contacts_addr       C   1   联系人地址
    '--    contacts_phno       C   1   联系人电话
    '--    pat_grdn_name       C   1   监护人
    '--    cert_no_other       C   1   其他证件
    '--    is_inhspt            C   1   是否在院:1-在院 ;0-不在院
    '--    pati_show_color      N   1   病人显示颜色
    '--    visit_room           C   1   就诊诊室
    '--    visit_statu          N   1   就诊状态
    '--    visit_time           C   1   就诊时间:yyyy-mm-dd hh24:mi:ss
    '--    create_time          C   1   登记时间:yyyy-mm-dd hh24:mi:ss
    '--    pati_email           C   1   email
    '--    pati_qq              C   1   qq
    '--    card_captcha         C   1  卡验证码
    '--    insurance_pwd        C       医保密码
    Select Case UCase(strName)
    
    Case UCase("病人ID")
        strValue = "_pati_id"
    Case UCase("主页ID")
        strValue = "_pati_pageid"
    Case "姓名"
        strValue = "_pati_name"
    Case "性别"
        strValue = "_pati_sex"
    Case "年龄"
        strValue = "_pati_age"
    Case "出生日期"
        strValue = "_pati_birthdate"
    Case "费别"
        strValue = "_fee_category"
    Case "门诊号"
        strValue = "_outpatient_num"
    Case "住院号"
        strValue = "_inpatient_num"
    Case "医疗付款方式名称"
        strValue = "_mdlpay_mode_name"
    Case "医疗付款方式编码"
        strValue = "_mdlpay_mode_code"
    Case "民族"
        strValue = "_pati_nation"
    Case "医保号"
        strValue = "_insurance_num"
    Case "身份证号"
        strValue = "_pati_idcard"
    Case "就诊卡号", "就诊卡"
        strValue = "_vcard_no"
    Case UCase("IC卡号")
        strValue = "_iccard_no"
    Case "健康号"
        strValue = "_health_num"
    Case "住院次数"
        strValue = "_inp_times"
    Case "学历"
        strValue = "_pati_education"
    Case "职业"
        strValue = "_ocpt_name"
    Case "身份"
        strValue = "_pati_identity"
    Case "籍贯"
        strValue = "_ntvplc_name"
    Case "国籍"
        strValue = "_country_name"
    Case "婚姻状况"
        strValue = "_pati_marital_cstatus"
    Case "家庭地址"
        strValue = "_pat_home_addr"
    Case "家庭电话"
        strValue = "_pat_home_phno"
    Case "家庭地址邮编"
        strValue = "_pat_home_postcode"
    Case "区域"
        strValue = "_pati_area"
    Case "出生地点"
        strValue = "_pati_birthplace"
    Case "户口地址"
        strValue = "_pat_hous_addr"
    Case "户口地址邮编"
        strValue = "_pat_hous_postcode"
    Case "工作单位名称"
        strValue = "_emp_name"
    Case "单位电话"
        strValue = "_emp_phno"
    Case "单位邮编"
        strValue = "_emp_postcode"
    Case "单位开户行"
        strValue = "_emp_bank_name"
    Case "单位帐号"
        strValue = "_emp_bank_accnum"
    Case "合同单位ID"
        strValue = "_ctt_unit_id"
    Case "手机号"
        strValue = "_phone_number"
    Case "当前床号"
        strValue = "_pati_bed"
    Case "病人类型"
        strValue = "_pati_type"
    Case "险类"
        strValue = "_insurance_type"
    Case "险类名称"
        strValue = "_insurance_name"
    Case "当前病区id"
        strValue = "_pati_wardarea_id"
    Case "当前病区名称"
        strValue = "_pati_wardarea_name"
    Case "当前科室id"
        strValue = "_pati_dept_id"
    Case "当前科室名称"
        strValue = "_pati_dept_name"
    Case "入院时间"
        strValue = "_adta_time"
    Case "出院时间"
        strValue = "_adtd_time"
    End Select
    GetMapPait = strValue
End Function

Public Function GetJsonInPatiState(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngWardID As Long, ByVal lngDeptId As Long, _
    ByVal strPatiBed As String, ByVal strInTime As String, ByVal bytInpStatus As Byte, strPatiNum As String) As String
'功能:
'更新病人就诊状态
'--input
'--    pati_list[]              数组
'--      pati_id              N 1   病人id
'--      pati_pageid          N 1   主页id
'--      outpatient_num       C 1   门诊号
'--      inpatient_num        C 1   住院号
'--      in_time              C 1   入院时间
'--      adtd_time            C 1   出院时间
'--      pati_deptid          N 1   当前科室id
'--      wardarea_id          N 1   当前病区id
'--      pati_bed             C 1   当前床号
'--      inp_status           N 1   是否在院，0/1
'--      inp_times            N 1   住院次数
'--      inp_times_increment  N 1   =1时-住院次数自增;=-1住院次数自减
'避免入院登记时病人信息同步失败,入住时再次更新。门诊号、住院次数 暂不考虑更新
    Dim strJsonIn As String
    
    strJsonIn = GetJsonNodeString("pati_id", lngPatiID, Json_num)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("pati_pageid", lngPageID, Json_num)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("wardarea_id", lngWardID, Json_num)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("pati_deptid", lngDeptId, Json_num)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("pati_bed", strPatiBed, Json_Text)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("in_time", strInTime, Json_Text)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("inp_status", bytInpStatus, Json_num)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("inpatient_num", strPatiNum, Json_Text)
    strJsonIn = GetNode("pati_list", "[{" & strJsonIn & "}]", True, 1)
    GetJsonInPatiState = GetNode("input", strJsonIn, True, 1)
End Function


