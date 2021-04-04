VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpenceSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'*********************************************************************************************************************************************
'费用相关服务
' 接口
'       1.zl_ExseSvr_Actualmoney-获取实收金额
'       2.zl_ExseSvr_GetNextNo-获取一下张单据号
'       3.zl_ExseSvr_GetNextID:获取下一个序列值
'       3.zl_Exsesvr_UpdCardFeeBlncInfo-修改卡费结算信息
'       4.zl_Exsesvr_DelCardfeeInfo-删除或销帐卡费费用
'       5.zl_ExseSvr_GetCardFeeInfoByNo-根据单据号获取卡费信息
'       6.zl_ExseSvr_AddDepositInfo-增加预交信息
'       7.zl_Exsesvr_Upddepositblncinfo-修正预产结算信息
'       8.zl_Exsesvr_DelDepositErrorRec-删除预交信息
'       9.zlExseSvr_GetChargeOffApply-根据病人id,获取费用销账申请数据
'      10.zlExseSvr_GetNeedAuditItems-获取指定病人的费用审批项目
'      11.zlExsesvr_ExistsPriceBill-根据病人id或医嘱ID,判断该病人或医嘱是否存在划价单
'      12.zlExseSvr_Getremainmoney-获取病人的相关费用信息等信息
'      13.zlExsesvr_GetPatiTotalMoney-获取病人的应收及实收金额
'      14.ZlExseSvr_GetOrderFeeState-获取对应医嘱的收费状态
'      15 zlExseSvr_GetFeeBalanceState-根据记帐单号，获取结帐状态
'      16.zlExseSvr_GetFeeChargeState-根据费用单据号，获取单据号所对应的收费状态
'      17.zlExseSvr_GetFeeInfoByBlncId-根据结帐ID，返回该结帐ID的费用明细数据
'         zlExseSvr_GetBalanceInfoById-根据结帐ID，返回该结帐ID的结算明细数据
'
'编制:刘兴洪
'日期:2019-11-25 11:11:39
'*********************************************************************************************************************************************
Private mlngErrNum As Long, mstrSource As String, mstrErrMsg As String
Private mobjServiceCall As Object
Private mobjService As clsService
Private mlngModule As Long
Public Function zlInitCommon(ByVal lngSys As Long, ByVal lngModule As Long, ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String, _
    Optional objServiceCall As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '     objServiceCall-费用服务对象
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    mlngModule = lngModule
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    Call InitVar
    
    If Not objServiceCall Is Nothing Then
        Set mobjServiceCall = objServiceCall
    Else
        If GetServiceCall(mobjServiceCall) = False Then Set mobjServiceCall = Nothing: Exit Function
    End If
    
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetServiceCall(ByRef objServiceCall_Out As Object, Optional blnShowErrMsg As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共服务对象
    '出参:objServiceCall_Out-返回公共服务对象
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
     
    If Not mobjServiceCall Is Nothing Then Set objServiceCall_Out = mobjServiceCall: GetServiceCall = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
    If Err <> 0 Then
        mstrErrMsg = "部件【zlServiceCall】丢失，请与系统管理员联系，恢复该部件！"
        If blnShowErrMsg Then
            MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
            Err = 0: On Error GoTo 0
        Else
            Err.Raise Err.Number, Err.Source, mstrErrMsg: Exit Function
        End If
        Exit Function
    End If
    
    On Error GoTo errHandle
    If mobjServiceCall.InitService(gcnOracle, gstrDBUser, glngSys, glngModul) = False Then
        Set mobjServiceCall = Nothing: Exit Function
    End If
    
    Set objServiceCall_Out = mobjServiceCall
    GetServiceCall = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Private Function GetService(ByRef objService_Out As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共服务对象
    '出参:objService_Out-返回公共服务对象
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    If Not mobjService Is Nothing Then Set objService_Out = mobjService: GetService = True: Exit Function
    
    Set mobjService = New clsService
    If mobjService.zlInitCommon(glngSys, mlngModule, gcnOracle, gstrDBUser) = False Then Set mobjService = Nothing: Exit Function
    Set objService_Out = mobjService:
    GetService = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zl_ExseSvr_GetNextNo(ByVal int序号 As Integer, ByRef strNO_Out As String, Optional ByVal lng科室id As Long, _
    Optional blnShowErrMsg As Boolean = True, Optional ByRef strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取费用单据号
    '入参:
    '出参:strErrMsg_Out-错误信息
    '     strNo_Out-返回的一下张单据号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
  
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的单据号!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'input
    '    item_num    N   1   项目序号
    '    dept_id N       科室ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("item_num", int序号, Json_num)
    strJson = strJson & "," & GetJsonNodeString("recv_id", lng科室id, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetNextNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    output
    '        code    N   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：失败时返回具体的错误信息
    '        next_no C       下一个号码

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取单据号，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strNO_Out = objServiceCall.GetJsonNodeValue("output.next_no")
    zl_ExseSvr_GetNextNo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_ExseSvr_GetNextID(ByVal strTableName As String, ByRef lngID_Out As Long, _
    Optional blnShowErrMsg As Boolean = True, Optional ByRef strErrmsg_Out As String, Optional ByVal str字段名 As String = "ID") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取下一个序列值
    '入参:
    '出参:strErrMsg_Out-错误信息
    '     lngID_Out-返回的一下序列值
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
  
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的序列!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
  
    'input
    '  table_name    C  1 表名
    '  col_name      C  1 字段名  序列名称不一定是ID，例如记录ID
    ' 出参:
    '  output
    '  next_id      N   1  序列
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("table_name", strTableName, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("col_name", str字段名, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetNextID"
    If objServiceCall.CallService(strServiceName, strJson, , "", mlngModule, False) = False Then Exit Function
    '    output
    '        code    N   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：失败时返回具体的错误信息
    '        next_no C       下一个号码

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取" & strTableName & "的序列值，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lngID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.next_id")))
    zl_ExseSvr_GetNextID = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Zl住院记帐记录_Verify_Check(ByVal strNOs As String, _
    ByVal cllPatient As Collection, ByVal str序号s As String, _
    Optional ByRef cllData_out As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:住院记帐审核合法性检查
    '入参:
    '   strNOs-单据号:允许传入多个，多个时，序号无效
    '   cllPatient-病人信息集，成员:Collect(病人ID,费用审核标志,住院状态)
    '   str序号s-多个用逗号，可以不传入
    '   lngModule -调用模块号
    '出参:
    '   cllData_out-返回的数据集，成员:Array(_rcp_no,_stuff_rcpdtl_ids,_drug_rcpdtl_ids,_autosendstuff_rcpdtl_ids)
    '返回:记帐审核合法，返回true,否则返回False
     '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, strJsonList As String
    Dim cllTemp As Collection, i As Long
     
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_住院记帐记录_Verify_Check
    '  --功能：审核一张住院记帐划价单的合法性检查
    '  --入参：Json_In:格式
    '  --  input
    '  --    fee_nos                       C   1   单据号，允许传入多张,多张时，用逗 \号分隔
    '  --    serials_num                   C   1   序号,多个用逗号分离,为空为所有,当fee_nos传入多张单据时，该项无效
    '  --    pati_list[]                   病人信息，仅审核这些病人的费用
    '  --      pati_id                     N   1   病人ID
    '  --      fee_audit_status            N   1   费用审核标志:0或空-未审核;1-已审核或开始审核(结合参数:病人审核方式来控制);2-完成审核,结合结帐权限[禁止未审核病人结帐]进行管理控制
    '  --      si_inp_status               N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                          N   1   应答吗：0-失败；1-成功
    '  --    message                       C   1   应答消息：失败时返回具体的错误信息
    '  --    item_list[]
    '  --      rcp_no                      C   1   处方单号
    '  --      stuff_rcpdtl_ids            C   1   卫材处方明细IDs:卫生材料所涉及的费用ids
    '  --      drug_rcpdtl_ids             C   1   药品处方明细IDs:药品所涉及的费用ids
    '  --      autosendstuff_rcpdtl_ids    C   1   发料处方明细IDs:自动发放卫生材料所涉及的费用IDs
    strJsonList = ""
    For i = 1 To cllPatient.Count
        Set cllTemp = cllPatient(i)
        strJson = ""
        strJson = strJson & "" & GetJsonNodeString("pati_id", cllTemp("病人ID"), Json_num)
        strJson = strJson & "," & GetJsonNodeString("fee_audit_status", cllTemp("费用审核标志"), Json_num)
        strJson = strJson & "," & GetJsonNodeString("si_inp_status", cllTemp("住院状态"), Json_num)
        strJsonList = strJsonList & ",{" & strJson & "}"
    Next
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_nos", strNOs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("serials_num", str序号s, Json_Text)
    strJson = strJson & ",""pati_list"":[" & Mid(strJsonList, 2) & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_住院记帐记录_Verify_Check", strJson, , "", mlngModule) = False Then Exit Function
    Set cllData_out = objServiceCall.GetJsonListValue("output.item_list")
    
    Zl住院记帐记录_Verify_Check = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlReadInPatientDelBillData(ByVal bytBillType As Byte, ByVal strNo As String, _
    ByVal bln住院单位 As Boolean, ByVal bln已结禁止销帐 As Boolean, ByVal bln禁止部分销帐 As Boolean, _
    ByVal str登记时间 As String, ByVal str收费类别s As String, ByVal str排除收费类别s As String, _
    ByRef rsBillData_out As ADODB.Recordset, ByRef rsIncome_out As ADODB.Recordset, _
    Optional strFrmCaption As String, Optional ByVal lngModule As Long, Optional ByVal str病人IDs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取需要销帐的单据数据
    '入参:
    '   bytBillType-1-收费单;2-记帐单;3-自动记帐单;
    '   strNo-单据号
    '   str登记时间-单据时间
    '   bln禁止部分销帐-
    '   str收费类别s-多个用逗号分离,如"5,6,7"
    '   str病人IDs-多病人单所销帐的病人IDs,如"1,2,3"
    '出参:
    '   rsBillData_out-单据数据
    '   rsIncome_out-收入汇总集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-16 20:04:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJsonTemp As String, strJson As String, i As Long, strJsonList As String
    Dim rsRecipe As ADODB.Recordset, cllData_out As Collection, cllTemp As Collection
    Dim objServiceCall As Object
    Dim blnExistDrug As Boolean, blnExistStuff As Boolean
    Dim cllAdviceData As Collection, rsAdviceExecuted As ADODB.Recordset '已执行数据：医嘱序号,NO,收费细目ID,已执行数
    Dim objService As clsService
    
    On Error GoTo errHandle
    If lngModule <> 0 Then mlngModule = lngModule
    
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If GetService(objService) = False Then Exit Function
      
    '获取准退数
    If CheckExistDrugOrStuff(strNo, bytBillType, False, blnExistDrug, blnExistStuff) = False Then Exit Function
    If zlGetDrugStuff_ExecutedNum(strNo, 2, rsRecipe, blnExistDrug, blnExistStuff) = False Then Exit Function
    
    If CheckExistAdviceFeeByNo(strNo, bytBillType, False, cllAdviceData) = False Then Exit Function
    If objService.ZlCisSvr_GetAdviceExcutNums(bytBillType, cllAdviceData, rsAdviceExecuted, mlngModule) = False Then Exit Function

    'Zl_住院记帐记录_Del_Get
    '  --功能：获取住院记帐销帐单据
    '  --入参：Json_In,格式如下
    '  --  input
    '  --      fee_no                    C   1   费用单据号
    '  --      balance_ban_writeoffs     N   1   已结禁止销帐:如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
    '  --      create_time               C       原单据的登记时间
    '  --      part_ban_writeoffs        N   1   禁止部分销帐:1-不允许；0-允许
    '  --      show_unit                 N   1   显示单位:0-售价单位;1-住院单位
    '  --      update_exestatu           N   1   是否更新执行状态：在获取时，发现执行状态不正确时，是否立即更新
    '  --      billtype                  N   1   单据类型：1-收费单(不支持);2-记帐单,3-自动记帐单
    '  --      oper_type                 N   1   操作方式:0-按单据信息返回;1-仅只返回销帐信息(将不返回收入列表及单据列表中的非销帐信息项
    '  --      pati_ids                  C   1   多病人单有效，表示只销指定病人Id的费用,多个用逗号分离:1,2,3,...
    '  --      fee_type_codes            C   0   收费类别编码:多个用逗号分离:"4,5,6,7…"
    '  --      no_fee_type_codes         C   0   排除的收费类别编码:多个用逗号分离:"4,5,6,7…
    '  --      executed_list[]           单据已执行列表(药品、卫材费用),即使已执行数为0也要传入
    '  --        fee_id                  N   1   费用ID
    '  --        sended_num              N   1   已发数量
    '  --      advice_excute_list[]      单据已执行列表(医嘱费用),即使已执行数为0也要传入
    '  --        advice_id               N   1   医嘱ID
    '  --        fee_item_id             N   1   收费细目ID
    '  --        execute_num             N   1   已执行数
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("balance_ban_writeoffs", IIf(bln已结禁止销帐, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("create_time", str登记时间, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("part_ban_writeoffs", IIf(bln禁止部分销帐, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("show_unit", IIf(bln住院单位, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("update_exestatu", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("billtype", bytBillType, Json_num)
    strJson = strJson & "," & GetJsonNodeString("oper_type", 0, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_ids", str病人IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("fee_type_codes", str收费类别s, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("no_fee_type_codes", str排除收费类别s, Json_Text)
    
    strJsonList = ""
    With rsRecipe
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strJsonTemp = ""
            strJsonTemp = strJsonTemp & "" & GetJsonNodeString("fee_id", Val(Nvl(!处方明细id)), Json_num)
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("sended_num", Val(Nvl(!已发数量)), Json_num)
            strJsonList = strJsonList & ",{" & strJsonTemp & "}"
            .MoveNext
        Loop
    End With
    If strJsonList <> "" Then
        strJson = strJson & "," & GetNodeString("executed_list") & ":[" & Mid(strJsonList, 2) & "]"
    End If
    
    strJsonList = ""
    With rsAdviceExecuted
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strJsonTemp = ""
            strJsonTemp = strJsonTemp & "" & GetJsonNodeString("advice_id", Val(Nvl(!医嘱序号)), Json_num)
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_item_id", Val(Nvl(!收费细目ID)), Json_num)
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("execute_num", Val(Nvl(!已执行数)), Json_num)
            strJsonList = strJsonList & ",{" & strJsonTemp & "}"
            .MoveNext
        Loop
    End With
    If strJsonList <> "" Then
        strJson = strJson & "," & GetNodeString("advice_excute_list") & ":[" & Mid(strJsonList, 2) & "]"
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_住院记帐记录_Del_Get", strJson, "", strFrmCaption, mlngModule, True) = False Then Exit Function
    Set cllData_out = objServiceCall.GetJsonListValue("output.item_list")
    If cllData_out Is Nothing Then Exit Function
    
    If zlBuildingDelBillingRecStru(rsBillData_out, rsIncome_out) = False Then Exit Function
    
    For i = 1 To cllData_out.Count
        Set cllTemp = cllData_out(i)
        With rsBillData_out
            .AddNew
            !病人ID = cllTemp("_pati_id")
            !主页ID = cllTemp("_pati_pageid")
            !姓名 = cllTemp("_pati_name")
            !性别 = cllTemp("_pati_sex")
            !年龄 = cllTemp("_pati_age")
            !床号 = cllTemp("_pati_bed")
            !费别 = cllTemp("_fee_category")
            !婴儿费 = cllTemp("_baby_num")
            !费用id = cllTemp("_fee_id")
            !序号 = cllTemp("_serial_num")
            !编码 = cllTemp("_chargetype_code")
            !类别 = cllTemp("_chargetype_name")
            !名称 = cllTemp("_item_name")
            !商品名 = cllTemp("_drug_brandname")
            !规格 = cllTemp("_item_spec")
            !费用类型 = cllTemp("_fee_type")
            !付数 = cllTemp("_packages_num")
            !数次 = cllTemp("_quantity")
            !计算单位 = cllTemp("_item_unit")
            !单价 = cllTemp("_price")
            !应收金额 = cllTemp("_fee_amrcvb")
            !实收金额 = cllTemp("_fee_ampaib")
            !执行部门 = cllTemp("_exe_deptname")
            !附加标志 = cllTemp("_rad_tag")
            !医嘱序号 = cllTemp("_adviceid")
            !跟踪在用 = cllTemp("_stuff_used")
            .Update
        End With
    Next
    If rsBillData_out.RecordCount > 0 Then rsBillData_out.MoveFirst
    
    Set cllData_out = objServiceCall.GetJsonListValue("output.income_list")
    If Not cllData_out Is Nothing Then
        For i = 1 To cllData_out.Count
            Set cllTemp = cllData_out(i)
            With rsIncome_out
                .AddNew
                !名称 = cllTemp("_income_name")
                !应收金额 = cllTemp("_fee_amrcvb")
                !实收金额 = cllTemp("_fee_ampaib")
                .Update
            End With
        Next
    End If
    If rsIncome_out.RecordCount > 0 Then rsIncome_out.MoveFirst
    zlReadInPatientDelBillData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlBuildingDelBillingRecStru(ByRef rsBilling As ADODB.Recordset, _
    ByRef rsIncom As ADODB.Recordset, _
    Optional ByVal blnOutpatient As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建本地记帐的记录集结构
    '出参:rsBilling-返回记帐的本地记录集
    '     rsIncom-返回收入项目数据集
    '返回:返回本地记录集成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-15 12:02:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set rsBilling = New ADODB.Recordset
    rsBilling.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "主页id", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "姓名", adVarChar, 50, adFldIsNullable
    rsBilling.Fields.Append "性别", adVarChar, 20, adFldIsNullable
    rsBilling.Fields.Append "年龄", adVarChar, 20, adFldIsNullable
    rsBilling.Fields.Append "床号", adVarChar, 20, adFldIsNullable
    rsBilling.Fields.Append "费别", adVarChar, 50, adFldIsNullable
    rsBilling.Fields.Append "婴儿费", adBigInt, , adFldIsNullable

    rsBilling.Fields.Append "序号", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "编码", adVarChar, 20, adFldIsNullable
    rsBilling.Fields.Append "类别", adVarChar, 30, adFldIsNullable
    rsBilling.Fields.Append "名称", adVarChar, 200, adFldIsNullable
    rsBilling.Fields.Append "计算单位", adVarChar, 20, adFldIsNullable
    rsBilling.Fields.Append "商品名", adVarChar, 200, adFldIsNullable
    rsBilling.Fields.Append "规格", adVarChar, 100, adFldIsNullable
    rsBilling.Fields.Append "费用类型", adVarChar, 100, adFldIsNullable
    rsBilling.Fields.Append "付数", adDouble, , adFldIsNullable
    rsBilling.Fields.Append "数次", adDouble, , adFldIsNullable
    rsBilling.Fields.Append "单价", adDouble, , adFldIsNullable
    rsBilling.Fields.Append "应收金额", adDouble, , adFldIsNullable
    rsBilling.Fields.Append "实收金额", adDouble, , adFldIsNullable
    rsBilling.Fields.Append "执行部门", adVarChar, 100, adFldIsNullable
    rsBilling.Fields.Append "附加标志", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "医嘱序号", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "跟踪在用", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "费用ID", adBigInt, , adFldIsNullable
    If blnOutpatient Then
        rsBilling.Fields.Append "NO", adVarChar, 8, adFldIsNullable
        rsBilling.Fields.Append "从属父号", adBigInt, , adFldIsNullable
        rsBilling.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
        rsBilling.Fields.Append "项目编码", adVarChar, 20, adFldIsNullable
        rsBilling.Fields.Append "医嘱内容", adLongVarChar, 1000, adFldIsNullable
        rsBilling.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
        rsBilling.Fields.Append "执行部门ID", adBigInt, , adFldIsNullable
        rsBilling.Fields.Append "操作员姓名", adVarChar, 50, adFldIsNullable
        rsBilling.Fields.Append "登记时间", adDate, , adFldIsNullable
        rsBilling.Fields.Append "挂号单", adVarChar, 8, adFldIsNullable
        rsBilling.Fields.Append "门诊标志", adInteger, , adFldIsNullable
        rsBilling.Fields.Append "商品条码", adVarChar, 100, adFldIsNullable
        rsBilling.Fields.Append "内部条码", adVarChar, 100, adFldIsNullable
        rsBilling.Fields.Append "执行状态", adInteger, , adFldIsNullable
        rsBilling.Fields.Append "门诊包装", adDouble, , adFldIsNullable
        rsBilling.Fields.Append "门诊单位", adVarChar, 10, adFldIsNullable
        rsBilling.Fields.Append "住院包装", adDouble, , adFldIsNullable
        rsBilling.Fields.Append "住院单位", adVarChar, 10, adFldIsNullable
        
        rsBilling.Fields.Append "结帐ID", adBigInt, , adFldIsNullable
        rsBilling.Fields.Append "记录标志", adInteger, , adFldIsNullable
        rsBilling.Fields.Append "原始数量", adDouble, , adFldIsNullable
        rsBilling.Fields.Append "准退数量", adDouble, , adFldIsNullable
        rsBilling.Fields.Append "剩余数量", adDouble, , adFldIsNullable
    End If
    rsBilling.CursorLocation = adUseClient
    rsBilling.LockType = adLockOptimistic
    rsBilling.CursorType = adOpenStatic
    rsBilling.Open
    
    Set rsIncom = New ADODB.Recordset
    rsIncom.Fields.Append "名称", adVarChar, 200, adFldIsNullable
    rsIncom.Fields.Append "应收金额", adDouble, , adFldIsNullable
    rsIncom.Fields.Append "实收金额", adDouble, , adFldIsNullable
    rsIncom.CursorLocation = adUseClient
    rsIncom.LockType = adLockOptimistic
    rsIncom.CursorType = adOpenStatic
    rsIncom.Open

    zlBuildingDelBillingRecStru = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlReadOutPatientDelBillData(cllInPara As Collection, ByVal strFrmCaption As String, _
    Optional ByRef rsBillData_out As ADODB.Recordset, Optional ByRef rsIncome_out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取门诊费用单据剩余/准退费用信息
    '入参:
    '   cllInPara 参数列表，集合，成员(Key)：
    '    [单据类型,单据号,显示单位(N),药品名称显示方式(N),已结禁止销帐(N),操作方式(N),医嘱ID(N),
    '     更新执行状态(N),取收入汇总(N),仅取剩余数量(N),收入汇总方式(N),取医嘱内容(N)]
    '       单据类型:1-收费单;2-记帐单;3-自动记帐单
    '       单据号:多个英文逗号分隔；格式：A001,A002,...
    '       显示单位:0-售价单位;1-门诊单位;2-住院单位 仅操作方式=0时有效
    '       药品名称显示方式:0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
    '       已结禁止销帐:已结账数是否禁止销账,记帐单有效:1-是;0-否
    '       操作方式:0-按单据详细信息返回;1-仅只返回简要信息：费用ID,序号,付数,数次,执行状态
    '       更新执行状态：在获取时，发现执行状态不正确时，是否立即更新:1-是;0-否
    '       取收入汇总：是否获取收入项目统计信息:1-是;0-否
    '       仅取有剩余数量：仅取有剩余数量的数据（不减去已执行数量）:1-是;0-否
    '       收入汇总方式：0-按收据费目,1-按收入项目,2-按单据
    '出参:
    '   rsBillData_out-单据数据，字段：
    '       序号,编码,类别,名称,计算单位,商品名,规格,费用类型,付数,数次,单价,
    '       应收金额,实收金额,执行部门,附加标志,医嘱序号,跟踪在用,门诊包装,门诊单位,住院包装,住院单位,
    '       NO,费用ID,从属父号,费别,收费细目ID,项目编码,医嘱内容,开单科室,执行部门ID,
    '       操作员姓名,登记时间,挂号单,门诊标志,商品条码,内部条码,
    '       结帐ID,记录标志,原始数量,准退数量,剩余数量
    '   rsIncome_out-收入汇总集，字段：
    '       名称,应收金,实收金额
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJsonTemp As String, strJson As String, i As Long, strJsonList As String
    Dim rsRecipe As ADODB.Recordset, cllData_out As Collection, cllTemp As Collection
    Dim cllItems As Collection, cllIncome As Collection
    Dim objServiceCall As Object
    Dim blnExistDrug As Boolean, blnExistStuff As Boolean
    Dim varNos As Variant, p As Integer
    Dim byt显示单位 As Byte '0-售价单位;1-门诊单位;2-住院单位
    Dim cllAdviceDatas As Collection, cllAdviceData As Collection
    Dim rsAdviceExecuted As ADODB.Recordset   '已执行数据：医嘱序号,NO,收费细目ID,已执行数
    Dim strAdviceExecutedList As String
    Dim str医嘱IDs As String, rsAdviceData As ADODB.Recordset
    Dim cllFilter As Collection, objService As clsService
    
    On Error GoTo ErrHandler
    '获取费用服务接口
    If cllInPara Is Nothing Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If GetService(objService) = False Then Exit Function
    
    byt显示单位 = Val(zlGetNodeValueFromCollect(cllInPara, "显示单位", "N"))
    
    If Val(zlGetNodeValueFromCollect(cllInPara, "仅取剩余数量", "N")) = 0 Then
        '获取准退数
        If CheckExistDrugOrStuff(cllInPara("单据号"), cllInPara("单据类型"), True, blnExistDrug, blnExistStuff) = False Then Exit Function
        If zlGetDrugStuff_ExecutedNum(cllInPara("单据号"), cllInPara("单据类型"), rsRecipe, blnExistDrug, blnExistStuff) = False Then Exit Function
        
        If CheckExistAdviceFeeByNo(cllInPara("单据号"), cllInPara("单据类型"), True, cllAdviceDatas) = False Then Exit Function
        If objService.ZlCisSvr_GetAdviceExcutNums(cllInPara("单据类型"), cllAdviceDatas, rsAdviceExecuted, mlngModule) = False Then Exit Function
    End If
    
    'Zl_门诊费用记录_Del_Get
    '  --功能：获取准退数，同时更新药品/卫材记录的执行状态
    '  --入参：Json_In,格式如下
    '  --  input
    '  --    fee_bill_type           N   1   单据性质:1-门诊收费单,2-门诊记帐单,3-自动记帐单
    '  --    oper_type               N   1   操作方式:0-按单据详细信息返回;1-仅只返回简要信息：费用ID,序号,付数,数次,执行状态
    '  --    balance_ban_writeoffs   N   1   已结禁止销帐:如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
    '  --    update_exestatu         N   1   是否更新执行状态：在获取时，发现执行状态不正确时，是否立即更新
    '  --    drug_name_display_mode  N   1   药品名称显示方式:0-显示通用名，1-显示商品名，2-同时显示通用名和商品名；仅oper_type=0有效
    '  --    return_incomelist       N   1   是否返回收入项目统计信息:1-是;0-否；仅oper_type=0有效
    '  --    return_remain           N   1   返回有剩余数量的数据（不减去已执行数量）:1-是;0-否；仅oper_type=0有效
    '  --    income_group_type       N   1   收入汇总方式：0-按收据费目,1-按收入项目,2-按单据
    '  --    bill_list[]
    '  --      fee_no                C   1   费用单据号
    '  --      order_id              N   1   医嘱ID
    '  --      executed_list[]           药品及卫材所对应的已发送数量
    '  --        fee_id              N   1   费用ID
    '  --        sended_num          N   1   已发数量
    '  --        barcode_goods       C   0   商品条码
    '  --        barcode_inside      C   0   内部条码
    '  --      advice_excute_list[]    单据已执行列表(医嘱费用),即使已执行数为0也要传入
    '  --        advice_id           N   1   医嘱ID
    '  --        fee_item_id         N   1   收费细目ID
    '  --        execute_num         N   1   已执行数
    varNos = Split(cllInPara("单据号"), ",")
    strJson = ""
    For p = 0 To UBound(varNos)
        strJsonList = ""
        If Not rsRecipe Is Nothing Then
            rsRecipe.Filter = "处方单号='" & varNos(p) & "'"
            Do While Not rsRecipe.EOF
                strJsonTemp = ""
                strJsonTemp = strJsonTemp & "" & GetJsonNodeString("fee_id", Nvl(rsRecipe!处方明细id), Json_num)
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("sended_num", Nvl(rsRecipe!已发数量), Json_num)
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("barcode_goods", Nvl(rsRecipe!商品条码), Json_Text)
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("barcode_inside", Nvl(rsRecipe!内部条码), Json_Text)
                strJsonList = strJsonList & ",{" & strJsonTemp & "}"
                rsRecipe.MoveNext
            Loop
        End If
        
        strAdviceExecutedList = ""
        If Not rsAdviceExecuted Is Nothing Then
            rsAdviceExecuted.Filter = "No='" & varNos(p) & "'"
            Do While Not rsAdviceExecuted.EOF
                strJsonTemp = ""
                strJsonTemp = strJsonTemp & "" & GetJsonNodeString("advice_id", Nvl(rsAdviceExecuted!医嘱序号), Json_num)
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_item_id", Nvl(rsAdviceExecuted!收费细目ID), Json_num)
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("execute_num", Nvl(rsAdviceExecuted!已执行数), Json_num)
                strAdviceExecutedList = strAdviceExecutedList & ",{" & strJsonTemp & "}"
                rsAdviceExecuted.MoveNext
            Loop
        End If
        
        strJsonTemp = ""
        strJsonTemp = strJsonTemp & "" & GetJsonNodeString("fee_no", varNos(p), Json_Text)
        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("order_id", zlGetNodeValueFromCollect(cllInPara, "医嘱ID", "N"), Json_num)
        strJsonTemp = strJsonTemp & ",""executed_list"":[" & Mid(strJsonList, 2) & "]"
        strJsonTemp = strJsonTemp & ",""advice_excute_list"":[" & Mid(strAdviceExecutedList, 2) & "]"
        strJson = strJson & ",{" & strJsonTemp & "}"
    Next
    strJson = Mid(strJson, 2)
    
    strJsonTemp = ""
    strJsonTemp = strJsonTemp & "" & GetJsonNodeString("fee_bill_type", cllInPara("单据类型"), Json_num)
    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("oper_type", zlGetNodeValueFromCollect(cllInPara, "操作方式", "N"), Json_num)
    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("balance_ban_writeoffs", zlGetNodeValueFromCollect(cllInPara, "已结禁止销帐", "N"), Json_num)
    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("update_exestatu", zlGetNodeValueFromCollect(cllInPara, "更新执行状态", "N"), Json_num)

    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("drug_name_display_mode", zlGetNodeValueFromCollect(cllInPara, "药品名称显示方式", "N"), Json_num)
    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("return_incomelist", zlGetNodeValueFromCollect(cllInPara, "取收入汇总", "N"), Json_num)
    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("return_remain", zlGetNodeValueFromCollect(cllInPara, "仅取剩余数量", "N"), Json_num)
    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("income_group_type", zlGetNodeValueFromCollect(cllInPara, "收入汇总方式", "N"), Json_num)
    strJsonTemp = strJsonTemp & ",""bill_list"":[" & strJson & "]"
    strJson = "{""input"":{" & strJsonTemp & "}}"
    
    If objServiceCall.CallService("Zl_门诊费用记录_Del_Get", strJson, "", strFrmCaption, mlngModule, True) = False Then Exit Function
    Set cllData_out = objServiceCall.GetJsonListValue("output.bill_list")
    If cllData_out Is Nothing Then Exit Function
    
    If Val(zlGetNodeValueFromCollect(cllInPara, "取医嘱内容", "N")) = 1 Then
        str医嘱IDs = ""
        If CheckExistAdviceFeeByNo(cllInPara("单据号"), cllInPara("单据类型"), True, cllAdviceDatas, True, True) = False Then Exit Function
        For i = 1 To cllAdviceDatas.Count
            Set cllAdviceData = cllAdviceDatas(i)
            str医嘱IDs = str医嘱IDs & "," & cllAdviceData("医嘱ID")
        Next
        If str医嘱IDs <> "" Then
            Set cllFilter = New Collection
            cllFilter.Add Array("医嘱IDS", Mid(str医嘱IDs, 2))
            If objService.ZlCissvr_GetAdviceInfo(cllFilter, rsAdviceData) = False Then Exit Function
        End If
    End If
    
    If zlBuildingDelBillingRecStru(rsBillData_out, rsIncome_out, True) = False Then Exit Function
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                    N   1   应答码：0-失败；1-成功
    '  --    message                 C   1   应答消息：失败时返回具体的错误信息
    '  ------oper_type=0返回以下节点---------------------------------------------
    '  --    bill_list[]
    '  --      fee_no                C   1   费用单据号
    '  --      item_list[]
    '  --        fee_id              N   1   费用ID（原始记录）
    '  --        serial_num          N   1   序号
    '  --        parent_serial_num   N   0   从属父号
    '  --        fee_category        C   0   费别
    '  --        chargetype_code     C   1   收费类别编码
    '  --        chargetype_name     C   1   收费类别名称
    '  --        item_id             N   1   收费项目ID
    '  --        item_code           N   1   收费项目编码
    '  --        item_name           C   1   收费项目名称
    '  --        drug_brandname      C   0   商品名
    '  --        item_spec           C   1   规格
    '  --        fee_type            C   0   费用类型
    '  --        stuff_used          N   0   是否跟踪在用：1-跟踪在用;0-未跟踪在用
    '  --        outp_packg_cvsfctr  N   0   门诊包装系数
    '  --        outp_unit           C   0   门诊单位
    '  --        inp_packg_cvsfctr   N   0   住院包装系数
    '  --        inp_unit            C   0   住院单位
    '  --        packages_num        N   1   付数
    '  --        quantity            N   1   数次(only_allowreturn=1,准退数量;0,剩余数次)
    '  --        item_unit           C   1   单位
    '  --        price               N   1   单价
    '  --        fee_amrcvb          N   1   应收金额
    '  --        fee_ampaib          N   1   实收金额
    '  --        rad_tag             N   0   附加标志
    '  --        adviceid            N   0   医嘱ID
    '  --        plcdept             C   0   开单科室名称
    '  --        exe_deptid          N   1   执行部门ID
    '  --        exe_deptname        C   1   执行部门名称
    '  --        operator_name       C   1   操作员姓名
    '  --        create_time         C   1   登记时间
    '  --        rgst_no             C   0   挂号单
    '  --        outpati_flag        N   1   门诊标志:1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
    '  --        barcode_goods       C   0   商品条码
    '  --        barcode_inside      C   0   内部条码
    '  --        balance_id          N   0   结帐ID
    '  --        fee_status          N   1   记录标志
    '  --        old_quantity        N   1   原始数量
    '  --        return_quantity     N   1   准退数量
    '  --        remain_quantity     N   1   剩余数量
    '  --      income_list[]           收入统计列表
    '  --        income_name       C   1   收入项目名称
    '  --        fee_amrcvb        N   1   应收金额
    '  --        fee_ampaib        N   1   实收金额
    '  ------oper_type=0返回以下节点---------------------------------------------
    '  --    bill_list[]
    '  --      fee_no                C   1   费用单据号
    '  --      item_list[]
    '  --        fee_id              N   1   费用ID（原始记录）
    '  --        serial_num          N   1   序号
    '  --        packages_num        N   1   付数
    '  --        quantity            N   1   数次(only_allowreturn=1,准退数量;0,剩余数次)
    '  --        execute_tag         N   1   执行状态：0-未执行;1-已执行;2-部分执行
    
    For p = 1 To cllData_out.Count
        Set cllItems = objServiceCall.GetJsonListValue("output.bill_list[" & p - 1 & "].item_list")
        If Not cllItems Is Nothing Then
            For i = 1 To cllItems.Count
                Set cllTemp = cllItems(i)
                With rsBillData_out
                    .AddNew
                    If Val(zlGetNodeValueFromCollect(cllInPara, "操作方式", "N")) = 0 Then
                        !序号 = cllTemp("_serial_num")
                        !编码 = cllTemp("_chargetype_code")
                        !类别 = cllTemp("_chargetype_name")
                        !名称 = cllTemp("_item_name")
                        !商品名 = cllTemp("_drug_brandname")
                        !规格 = cllTemp("_item_spec")
                        !费用类型 = cllTemp("_fee_type")
                        !付数 = cllTemp("_packages_num")
                        !数次 = cllTemp("_quantity")
                        !计算单位 = cllTemp("_item_unit")
                        !单价 = cllTemp("_price")
                        !应收金额 = cllTemp("_fee_amrcvb")
                        !实收金额 = cllTemp("_fee_ampaib")
                        !执行部门 = cllTemp("_exe_deptname")
                        !附加标志 = cllTemp("_rad_tag")
                        !医嘱序号 = cllTemp("_adviceid")
                        !跟踪在用 = cllTemp("_stuff_used")
                        
                        !NO = cllData_out(p)("_fee_no")
                        !费用id = cllTemp("_fee_id")
                        !从属父号 = cllTemp("_parent_serial_num")
                        !费别 = cllTemp("_fee_category")
                        !收费细目ID = cllTemp("_item_id")
                        !项目编码 = cllTemp("_item_code")
                        !开单科室 = cllTemp("_plcdept")
                        !执行部门ID = cllTemp("_exe_deptid")
                        !操作员姓名 = cllTemp("_operator_name")
                        !登记时间 = cllTemp("_create_time")
                        !挂号单 = cllTemp("_rgst_no")
                        !门诊标志 = cllTemp("_outpati_flag")
                        !商品条码 = cllTemp("_barcode_goods")
                        !内部条码 = cllTemp("_barcode_inside")
                        
                        !门诊包装 = cllTemp("_outp_packg_cvsfctr")
                        !门诊单位 = cllTemp("_outp_unit")
                        !住院包装 = cllTemp("_inp_packg_cvsfctr")
                        !住院单位 = cllTemp("_inp_unit")
                        
                        !结帐ID = cllTemp("_balance_id")
                        !记录标志 = cllTemp("_fee_status")
                        !原始数量 = cllTemp("_old_quantity")
                        !准退数量 = cllTemp("_return_quantity")
                        !剩余数量 = cllTemp("_remain_quantity")
                        
                        '按显示单位转换
                        If byt显示单位 = 1 Then '门诊单位
                            !数次 = !数次 / Nvl(!门诊包装, 1)
                            !计算单位 = Nvl(!门诊单位, !计算单位)
                            !单价 = !单价 * Nvl(!门诊包装, 1)
                        ElseIf byt显示单位 = 2 Then '住院单位
                            !数次 = !数次 / Nvl(!住院包装, 1)
                            !计算单位 = Nvl(!住院单位, !计算单位)
                            !单价 = !单价 * Nvl(!住院包装, 1)
                        End If
                        
                        '医嘱id,相关id,医嘱内容,挂号单号,诊疗类别
                        If Not rsAdviceData Is Nothing And Val(Nvl(!医嘱序号)) <> 0 Then
                            rsAdviceData.Filter = "医嘱id=" & Val(Nvl(!医嘱序号))
                            If Not rsAdviceData.EOF Then !医嘱内容 = Nvl(rsAdviceData!医嘱内容)
                        End If
                    Else
                        !NO = cllData_out(p)("_fee_no")
                        !费用id = cllTemp("_fee_id")
                        !序号 = cllTemp("_serial_num")
                        !付数 = cllTemp("_packages_num")
                        !数次 = cllTemp("_quantity")
                        !执行状态 = cllTemp("_execute_tag")
                    End If
                    .Update
                End With
            Next
        End If
        
        Set cllIncome = objServiceCall.GetJsonListValue("output.bill_list[" & p - 1 & "].income_list")
        If Not cllIncome Is Nothing Then
            For i = 1 To cllIncome.Count
                Set cllTemp = cllIncome(i)
                With rsIncome_out
                    .AddNew
                    !名称 = cllTemp("_income_name")
                    !应收金额 = cllTemp("_fee_amrcvb")
                    !实收金额 = cllTemp("_fee_ampaib")
                    .Update
                End With
            Next
        End If
    Next
    If rsBillData_out.RecordCount > 0 Then rsBillData_out.MoveFirst
    If rsIncome_out.RecordCount > 0 Then rsIncome_out.MoveFirst
    zlReadOutPatientDelBillData = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function GetPatiJsonByFeeID(ByVal str费用IDs As String, ByRef strPatiJson As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用ID获取病人信息JSON串
    '入参:
    '   str费用IDs=费用ID,多个用英文逗号分隔
    '出参:
    '   strPatiJson=病人信息JSON对象字符串，格式如下:
    '       {"pati_id":201,"pati_name":"王二小","pati_dept_id":1201,"pati_name":"王小五","fee_audit_status":0,"si_inp_status":0},...
    '返回:获取成功返回True，失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSubTable As String, varPara() As Variant, rsFee As ADODB.Recordset
    Dim rsPati As ADODB.Recordset, strJson As String
    Dim strSql As String, str病人IDs As String
    Dim objService As clsService
    
    On Error GoTo ErrHandler
    strPatiJson = ""
    If str费用IDs = "" Then GetPatiJsonByFeeID = True: Exit Function
    If GetService(objService) = False Then Exit Function
    
    If zlGetVarBoundSQL(0, str费用IDs, strSubTable, varPara) = False Then Exit Function
    
    strSql = "Select a.病人ID,a.主页ID From 住院费用记录 A,(" & strSubTable & ") B Where a.ID=b.Column_value" & _
            " Union All" & _
            " Select a.病人ID,a.主页ID From 门诊费用记录 A,(" & strSubTable & ") B Where a.ID=b.Column_value"
    
    strSql = "Select f_List2str(Cast(Collect(Distinct 病人id || ':' || Nvl(主页id,0)) As t_Strlist)) As 病人IDS" & _
            " From (" & strSql & ")"
    Set rsFee = gobjDatabase.OpenSQLRecordByArray(strSql, "批量查询费用", varPara)
    str病人IDs = Nvl(rsFee!病人IDs)
    
    If str病人IDs = "" Then Exit Function

    If objService.ZlCissvr_GetPatiPageInfo(1, str病人IDs, rsPati) = False Then Exit Function
    strPatiJson = ""
    Do While Not rsPati.EOF
        '  --    pati_list[]                 病人信息
        '  --      pati_id                   N   1   病人ID
        '  --      pati_name                 C   1   病人姓名
        '  --      pati_dept_id              N   1   当前科室id
        '  --      fee_audit_status          N   1   费用审核标志:0或空-未审核;1-已审核或开始审核;2-完成审核
        '  --      si_inp_status             N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
        strJson = "{"
        strJson = strJson & "" & GetJsonNodeString("pati_id", Val(Nvl(rsPati!病人ID)), Json_num)
        strJson = strJson & "," & GetJsonNodeString("pati_name", Nvl(rsPati!姓名), Json_Text)
        strJson = strJson & "," & GetJsonNodeString("pati_dept_id", Val(Nvl(rsPati!当前科室id)), Json_num)
        strJson = strJson & "," & GetJsonNodeString("fee_audit_status", Val(Nvl(rsPati!审核标志)), Json_num)
        strJson = strJson & "," & GetJsonNodeString("si_inp_status", Val(Nvl(rsPati!住院状态)), Json_num)
        strJson = strJson & "}"
        strPatiJson = strPatiJson & "," & strJson
        rsPati.MoveNext
    Loop
    If strPatiJson <> "" Then strPatiJson = Mid(strPatiJson, 2)
    
    GetPatiJsonByFeeID = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zl病人费用销帐_Insert(ByVal cllApplyDatas As Collection, ByRef str申请费用ids_Out As String, _
    Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人费用销帐申请
    '入参:
    '   cllApplyDatas-销账申请数据集(RowData=collect:(费用ID,单据号,收费细目ID,申请科室ID,审核科室ID,申请数量,申请人,申请时间,申请类别,收费类别,销帐原因))
    '出参:
    '   str申请费用ids_Out-本次申请所涉及的费用ID(主要是后继审核要用)
    '返回:申请返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl已发数量 As Double, lng审核科室ID As Long
    Dim objServiceCall As Object, cllTemp As Collection
    Dim strJson As String, strTemp As String, strJsonList As String
    Dim strDrugNos As String, strStufNos As String, rsTemp As ADODB.Recordset
    Dim cllData_out As Collection, i As Long, strSql As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str费用IDs As String, strPatiJson As String
    Dim m As Long
    
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    str申请费用ids_Out = ""
    If GetServiceCall(objServiceCall) = False Or cllApplyDatas Is Nothing Then Exit Function
    
    If cllApplyDatas.Count = 0 Then Exit Function
    
    '第一步：先进行检查
    strDrugNos = "": strStufNos = ""
    For i = 1 To cllApplyDatas.Count
        Set cllTemp = cllApplyDatas(i)
        If InStr(strDrugNos & ",", "," & cllTemp("单据号")) = 0 And InStr(",5,6,7,", "," & cllTemp("收费类别") & ",") > 0 Then
            strDrugNos = strDrugNos & "," & cllTemp("单据号")
        End If
        If InStr(strStufNos & ",", "," & cllTemp("单据号")) = 0 And InStr(",4,", "," & cllTemp("收费类别") & ",") > 0 Then
            strStufNos = strStufNos & "," & cllTemp("单据号")
        End If
    Next
    
    Set rsTemp = Nothing
    '药品
    If strDrugNos <> "" Then
        strDrugNos = Mid(strDrugNos, 2)
        If zlGetDrugStuff_ExecutedNum(strDrugNos, 2, rsTemp, True, False, True) = False Then Exit Function
    End If
    '卫材
    If strStufNos <> "" Then
        strStufNos = Mid(strStufNos, 2)
        If zlGetDrugStuff_ExecutedNum(strStufNos, 2, rsTemp, False, True, True) = False Then Exit Function
    End If
    
    'Zl_病人费用销帐_Insert_Check
    '  --功能：针对指定单据指定行行进行销帐申请检查
    '  --入参：Json_In:格式
    '  --  input
    '  --    item_list[]                 本次销帐列表
    '  --      fee_id                    N   1   费用ID
    '  --      request_dept_id           N   1   销帐申请科室ID
    '  --      audit_dept_id             N   1   审核科室ID
    '  --      request_type              N   1   申请类别:对药品和卫材有效:0-未发药(料);1-已发药(料);其他为0
    '  --      request_num               N   1   申请数量
    '  --      sended_num                N   1   已发数量
    '  --    pati_list[]                 病人信息
    '  --      pati_id                   N   1   病人ID
    '  --      pati_name                 C   1   病人姓名
    '  --      pati_dept_id              N   1   当前科室id
    '  --      fee_audit_status          N   1   费用审核标志:0或空-未审核;1-已审核或开始审核;2-完成审核
    '  --      si_inp_status             N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                    N   1   应答吗：0-失败；1-成功
    '  --    message                 C   1   应答消息：失败时返回具体的错误信息
    '  --    item_list[]                      单据数据列表
    '  --      fee_id                 N 1 费用ID
    '  --      request_dept_id        N 1 销帐申请科室ID
    '  --      audit_dept_id          N 1 销帐审核科室ID
    '  --      request_type           N 1 申请类别:对药品和卫材有效:0-未发药(料);1-已发药(料);其他为0
    strJsonList = ""
    For i = 1 To cllApplyDatas.Count
        Set cllTemp = cllApplyDatas(i)
        dbl已发数量 = 0
        If Not rsTemp Is Nothing Then
            rsTemp.Filter = "处方明细ID=" & cllTemp("费用ID")
            If Not rsTemp.EOF Then dbl已发数量 = Val(Nvl(rsTemp!已发数量))
        End If
        
        strTemp = "{"
        strTemp = strTemp & "" & GetJsonNodeString("fee_id", cllTemp("费用ID"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("request_dept_id", cllTemp("申请科室ID"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("audit_dept_id", zlGetNodeValueFromCollect(cllTemp, "审核科室ID", "N"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("request_type", cllTemp("申请类别"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("request_num", cllTemp("申请数量"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("sended_num", dbl已发数量, Json_num)
        strTemp = strTemp & "}"
        strJsonList = strJsonList & "," & strTemp
        
        str费用IDs = str费用IDs & "," & cllTemp("费用ID")
    Next
    If strJsonList <> "" Then strJsonList = Mid(strJsonList, 2)
    If strJsonList = "" Then Exit Function
    
    If GetPatiJsonByFeeID(Mid(str费用IDs, 2), strPatiJson) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & """item_list"":[" & strJsonList & "]"
    strJson = strJson & ",""pati_list"":[" & strPatiJson & "]"
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_病人费用销帐_Insert_Check", strJson, "", strFrmCaption, mlngModule, True) = False Then Exit Function
    
    '获取返回值
    Set cllData_out = objServiceCall.GetJsonListValue("output.item_list")
    If cllData_out Is Nothing Then Exit Function
    Set cllPro = New Collection
    For i = 1 To cllApplyDatas.Count
        '费用ID,单据号,收费细目ID,申请科室ID,审核科室ID,申请数量, 配药ids,申请人,申请时间,申请类别,收费类别,销帐原因
        Set cllTemp = cllApplyDatas(i)
        If IsEmpty(zlGetNodeValueFromCollect(cllTemp, "审核科室ID", "N")) Then
            lng审核科室ID = 0
        Else
            lng审核科室ID = zlGetNodeValueFromCollect(cllTemp, "审核科室ID", "N")
        End If
        If lng审核科室ID = 0 Then
            For m = 1 To cllData_out.Count
                If Val(Nvl(cllData_out(m)("_fee_id"))) = Val(Nvl(cllTemp("费用ID"))) And Val(Nvl(cllData_out(m)("_request_type"))) = Val(Nvl(cllTemp("申请类别"))) Then
                    lng审核科室ID = Val(Nvl(cllData_out(m)("_audit_dept_id"))): Exit For
                End If
            Next
        End If
        strSql = "Zl_病人费用销帐_Insert_S("
        '    Id_In         In 病人费用销帐.费用id%Type,
        strSql = strSql & "" & cllTemp("费用ID") & ","
        '    收费细目id_In In 病人费用销帐.收费细目id%Type,
        strSql = strSql & "" & cllTemp("收费细目ID") & ","
        '    申请部门id_In In 病人费用销帐.申请部门id%Type,
        strSql = strSql & "" & cllTemp("申请科室ID") & ","
        '    数量_In       In 病人费用销帐.数量%Type,
        strSql = strSql & "" & cllTemp("申请数量") & ","
        '    申请人_In     In 病人费用销帐.申请人%Type,
        strSql = strSql & "'" & cllTemp("申请人") & "',"
        '    申请时间_In   In 病人费用销帐.申请时间%Type,
        strSql = strSql & "to_date('" & cllTemp("申请时间") & "','yyyy-mm-dd HH24:mi:ss'),"
        '    申请类别_In   In 病人费用销帐.申请类别%Type,
        strSql = strSql & "" & cllTemp("申请类别") & ","
        '    销帐原因_In   In 病人费用销帐.销帐原因%Type := Null,
        strSql = strSql & "'" & cllTemp("销帐原因") & "',"
        '    审核部门id_In In 病人费用销帐.审核部门id%Type := Null,
        strSql = strSql & "" & ZVal(lng审核科室ID) & ","
        '    删除标志_In   In Integer := 0:取消以前申请的重新生成(按批次申请时，不能取消，因为费用id相同，每个批次可分别申请)
        strSql = strSql & "" & IIf(InStr(1, "," & str申请费用ids_Out & ",", "," & cllTemp("费用ID") & ",") > 0, 0, 1) & " )"
        zlAddArray cllPro, strSql
        
        If InStr(str申请费用ids_Out & ",", "," & cllTemp("费用ID") & ",") = 0 Then
            str申请费用ids_Out = str申请费用ids_Out & "," & cllTemp("费用ID")
        End If
    Next
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, strFrmCaption
    blnTrans = False
    
    zl病人费用销帐_Insert = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function Zl_病人费用销帐_Audit_Check(ByVal cllAuditDatas As Collection, _
    Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long, _
    Optional ByRef rsSendDatas As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人费用销帐申请审核合法性检查
    '入参:
    '   cllAuditDatas-销账审核数据集(RowData=collect:(费用ID,单据号,收费类别,卫材是否自动退料,申请时间,申请类别))
    '          成员值说明:1.卫生材料是否自动退料:1-自动退料;0-不自动退料
    '                     2.申请类别:0-未发药(料);1-已发药(料);其他为0
    '   rsSendDatas-not Nothing 代表已经在外部获取出已发药或料的数据，内部直接会用，不用再调用服务
    '出参:
    '   rsSendDatas-药品及卫生材料的已执行数量的数据集
    '返回:合法返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl已发数量 As Double
    Dim objServiceCall As Object, cllTemp As Collection
    Dim strJson As String, strTemp As String, strJsonList As String
    Dim strDrugNos As String, strStufNos As String
    Dim i As Long, str费用IDs As String, strPatiJson As String
    
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Or cllAuditDatas Is Nothing Then Exit Function
    
    If cllAuditDatas.Count = 0 Then Exit Function
    
    If rsSendDatas Is Nothing Then
        strDrugNos = "": strStufNos = ""
        For i = 1 To cllAuditDatas.Count
            Set cllTemp = cllAuditDatas(i)
            If InStr(strDrugNos & ",", "," & cllTemp("单据号")) = 0 And InStr(",5,6,7,", "," & cllTemp("收费类别") & ",") > 0 Then
                strDrugNos = strDrugNos & "," & cllTemp("单据号")
            End If
            If InStr(strStufNos & ",", "," & cllTemp("单据号")) = 0 And InStr(",4,", "," & cllTemp("收费类别") & ",") > 0 Then
                strStufNos = strStufNos & "," & cllTemp("单据号")
            End If
        Next
        
        Set rsSendDatas = Nothing
        '药品
        If strDrugNos <> "" Then
            strDrugNos = Mid(strDrugNos, 2)
            If zlGetDrugStuff_ExecutedNum(strDrugNos, 2, rsSendDatas, True, False, True) = False Then Exit Function
        End If
        '卫材
        If strStufNos <> "" Then
            strStufNos = Mid(strStufNos, 2)
            If zlGetDrugStuff_ExecutedNum(strStufNos, 2, rsSendDatas, False, True, True) = False Then Exit Function
        End If
    End If
    
    '第一步：先进行检查
    'Zl_病人费用销帐_Audit_Check
    '  --功能：针对指定单据指定行行进行销帐申请检查
    '  --入参：Json_In:格式
    '  --  input
    '  --    no_consistence              销帐数据不一致时不检查并发
    '  --    item_list[]                 本次销帐列表
    '  --      fee_id                    N 1 费用ID
    '  --      stuff_auto_return         N   卫生材料是否自动退料:1-自动退料;0-不自动退料
    '  --      request_time              C 1 申请时间:yyyy-mm-dd hh24:mi:ss
    '  --      request_type              N 1 申请类别:对药品和卫材有效:0-未发药(料);1-已发药(料);其他为0
    '  --      sended_num                N 1 已发数量
    '  --    pati_list[]                 病人信息
    '  --      pati_id                   N   1   病人ID
    '  --      pati_name                 C   1   病人姓名
    '  --      fee_audit_status          N   1   费用审核标志:0或空-未审核;1-已审核或开始审核;2-完成审核
    '  --      si_inp_status             N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
    '  --出参: Json_Out,格式如下
    '  --    output
    '  --        code                    N   1   应答吗：0-失败；1-成功
    '  --        message                 C   1   应答消息：失败时返回具体的错误信息
    strJsonList = ""
    For i = 1 To cllAuditDatas.Count
        Set cllTemp = cllAuditDatas(i)
        dbl已发数量 = 0
        If Not rsSendDatas Is Nothing Then
              rsSendDatas.Filter = "处方明细ID=" & cllTemp("费用ID")
              If Not rsSendDatas.EOF Then dbl已发数量 = Val(Nvl(rsSendDatas!已发数量))
        End If
        
        strTemp = "{"
        strTemp = strTemp & "" & GetJsonNodeString("fee_id", cllTemp("费用ID"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("stuff_auto_return", cllTemp("卫材是否自动退料"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("request_time", zlGetNodeValueFromCollect(cllTemp, "申请时间", "C"), Json_Text)
        strTemp = strTemp & "," & GetJsonNodeString("request_type", cllTemp("申请类别"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("sended_num", dbl已发数量, Json_num)
        strTemp = strTemp & "}"
        strJsonList = strJsonList & "," & strTemp
        
        str费用IDs = str费用IDs & "," & cllTemp("费用ID")
    Next
    If strJsonList <> "" Then strJsonList = Mid(strJsonList, 2)
    If strJsonList = "" Then Exit Function
    
    If GetPatiJsonByFeeID(Mid(str费用IDs, 2), strPatiJson) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & """item_list"":[" & strJsonList & "]"
    strJson = strJson & ",""pati_list"":[" & strPatiJson & "]"
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_病人费用销帐_Audit_Check", strJson, "", strFrmCaption, mlngModule, True) = False Then Exit Function
    
    Zl_病人费用销帐_Audit_Check = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function Zl_病人费用销帐_Cancel_Check(ByVal cllAuditDatas As Collection, _
    Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long, _
    Optional ByRef rsSendDatas As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:针对“取消拒发”或“重审拒发”的功能进行合法性检查
    '入参:
    '   cllAuditDatas-销账审核数据集(RowData=collect:(操作状态,费用ID,单据号(N),收费类别,卫材是否自动退料(N),申请时间,申请类别(N),操作员姓名))
    '          成员值说明:0.操作状态:0-审核拒绝的申请;1-取消拒绝的申请（操作状态_In=1时无效))
    '                     1.卫生材料是否自动退料:1-自动退料;0-不自动退料（操作状态_In=1时无效))
    '                     2.申请类别:0-未发药(料);1-已发药(料);其他为0（操作状态_In=1时无效))
    '                     3.N代表可选
    '   rsSendDatas-not Nothing 代表已经在外部获取出已发药或料的数据，内部直接会用，不用再调用服务
    '出参:
    '   rsSendDatas-药品及卫生材料的已执行数量的数据集
    '返回:合法返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl已发数量 As Double
    Dim objServiceCall As Object, cllTemp As Collection
    Dim strJson As String, strTemp As String, strJsonList As String
    Dim strDrugFeeIDs As String, strStuffgFeeIDs As String
    Dim i As Long, str费用IDs As String, strPatiJson As String
    Dim bln重审拒发 As Boolean
    
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Or cllAuditDatas Is Nothing Then Exit Function
    
    If cllAuditDatas.Count = 0 Then Exit Function
    
    bln重审拒发 = False
    strDrugFeeIDs = "": strStuffgFeeIDs = ""
    For i = 1 To cllAuditDatas.Count
        Set cllTemp = cllAuditDatas(i)
        If InStr(strDrugFeeIDs & ",", "," & cllTemp("费用ID")) = 0 And InStr(",5,6,7,", "," & cllTemp("收费类别") & ",") > 0 Then
            strDrugFeeIDs = strDrugFeeIDs & "," & cllTemp("费用ID")
        End If
        If InStr(strStuffgFeeIDs & ",", "," & cllTemp("费用ID")) = 0 And InStr(",4,", "," & cllTemp("收费类别") & ",") > 0 Then
            strStuffgFeeIDs = strStuffgFeeIDs & "," & cllTemp("费用ID")
        End If
        If Not bln重审拒发 Then bln重审拒发 = Val(Nvl(cllTemp("操作状态"))) = 0
    Next
        
    If rsSendDatas Is Nothing Then
        Set rsSendDatas = Nothing
        If bln重审拒发 Then
            '药品
             If strDrugFeeIDs <> "" Then
                 strDrugFeeIDs = Mid(strDrugFeeIDs, 2)
                 If zlGetDrugStuff_ExecutedNum("", 2, rsSendDatas, True, False, True, strDrugFeeIDs, mlngModule) = False Then Exit Function
             End If
             '卫材
             If strStuffgFeeIDs <> "" Then
                 strStuffgFeeIDs = Mid(strStuffgFeeIDs, 2)
                 If zlGetDrugStuff_ExecutedNum("", 2, rsSendDatas, False, True, True, strStuffgFeeIDs, mlngModule) = False Then Exit Function
             End If
        End If
    End If
    
    '第一步：先进行检查
    'Zl_病人费用销帐_Cancel_Check
    '  --功能：针对指定单据指定行行进行销帐申请检查
    '  --入参：Json_In-格式如下
    '  --  input
    '  --    item_list[]                 本次销帐列表
    '  --      fee_id                    N   1   费用ID
    '  --      oper_type                 N   1   操作状态_In:0-审核拒绝的申请;1-取消拒绝的申请
    '  --      request_type              N   申请类别（操作状态_In=1时无效)::对药品和卫材有效:0-未发药(料);1-已发药(料);其他为0
    '  --      request_time              C   1   申请时间:yyyy-mm-dd hh24:mi:ss
    '  --      stuff_auto_return         N   卫生材料是否自动退料（操作状态_In=1时无效):1-自动退料;0-不自动退料
    '  --      operator_name             C   1   操作员姓名
    '  --      sended_num                N   已发数量（操作状态_In=1时无效)
    '  --    pati_list[]                 病人信息
    '  --      pati_id                   N   1   病人ID
    '  --      pati_name                 C   1   病人姓名
    '  --      fee_audit_status          N   1   费用审核标志:0或空-未审核;1-已审核或开始审核;2-完成审核
    '  --      si_inp_status             N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
    '  --出参 Json_Out-格式如下
    '  --    output
    '  --        code                    N   1   应答吗：0-失败；1-成功
    '  --        message                 C   1   应答消息：失败时返回具体的错误信息
    strJsonList = ""
    For i = 1 To cllAuditDatas.Count
        Set cllTemp = cllAuditDatas(i)
        dbl已发数量 = 0
        If bln重审拒发 Then
            If Not rsSendDatas Is Nothing Then
                  rsSendDatas.Filter = "处方明细ID=" & cllTemp("费用ID")
                  If Not rsSendDatas.EOF Then dbl已发数量 = Val(Nvl(rsSendDatas!已发数量))
            End If
        End If
        
        strTemp = "{"
        strTemp = strTemp & "" & GetJsonNodeString("fee_id", cllTemp("费用ID"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("oper_type", cllTemp("操作状态"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("request_type", zlGetNodeValueFromCollect(cllTemp, "申请类别", "N"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("request_time", zlGetNodeValueFromCollect(cllTemp, "申请时间", "C"), Json_Text)
        strTemp = strTemp & "," & GetJsonNodeString("stuff_auto_return", zlGetNodeValueFromCollect(cllTemp, "卫材是否自动退料", "N"), Json_num)
        strTemp = strTemp & "," & GetJsonNodeString("operator_name", zlGetNodeValueFromCollect(cllTemp, "操作员姓名", "C"), Json_Text)
        strTemp = strTemp & "," & GetJsonNodeString("sended_num", dbl已发数量, Json_num)
        strTemp = strTemp & "}"
        strJsonList = strJsonList & "," & strTemp
        
        str费用IDs = str费用IDs & "," & cllTemp("费用ID")
    Next
    If strJsonList <> "" Then strJsonList = Mid(strJsonList, 2)
    If strJsonList = "" Then Exit Function
    
    If GetPatiJsonByFeeID(Mid(str费用IDs, 2), strPatiJson) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & """item_list"":[" & strJsonList & "]"
    strJson = strJson & ",""pati_list"":[" & strPatiJson & "]"
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_病人费用销帐_Cancel_Check", strJson, "", strFrmCaption, mlngModule, True) = False Then Exit Function
    
    Zl_病人费用销帐_Cancel_Check = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function GetPatiJsonByFeeBill(ByVal int单据性质 As Integer, _
    ByVal strNo As String, ByVal str序号 As String, ByRef strPatiJson As String, _
    Optional ByVal int费用来源 As Integer = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用单据获取病人信息JSON串
    '入参:
    '   int单据性质-1-收费;2-记帐;3-自动记帐
    '   strNO-单据号
    '   str序号-序号
    '   int费用来源-1-门诊;2-住院
    '出参:
    '   strPatiJson=病人信息JSON对象字符串，格式如下:
    '       {"pati_id":201,"pati_name":"王小二","fee_audit_status":0,"si_inp_status":0},{"pati_id":202,"pati_name":"王小五","fee_audit_status":0,"si_inp_status":0},...
    '返回:获取成功返回True，失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSubTable As String, varPara() As Variant, rsTemp As ADODB.Recordset
    Dim rsPati As ADODB.Recordset, strJson As String
    Dim strSql As String, str病人IDs As String
    Dim objService As clsService
    
    On Error GoTo ErrHandler
    strPatiJson = ""
    If GetService(objService) = False Then Exit Function
    
    strSql = "Select f_List2str(Cast(Collect(Distinct 病人id || ':' || Nvl(主页id,0)) As t_Strlist)) As 病人IDS" & _
            " From " & IIf(int费用来源 = 2, "住院费用记录", "门诊费用记录") & "" & _
            " Where 记录性质=[1] And NO=[2] And 记录状态 In(0,1)" & _
            "       And (Instr(','||[3]||',',','||Nvl(价格父号,序号)||',')>0 Or [3] Is Null)"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取病人ID", int单据性质, strNo, str序号)
    str病人IDs = Nvl(rsTemp!病人IDs)
    
    If str病人IDs = "" Then Exit Function
    
    If objService.ZlCissvr_GetPatiPageInfo(0, str病人IDs, rsPati) = False Then Exit Function
    strPatiJson = ""
    Do While Not rsPati.EOF
        '  --    pati_list[]                 病人信息
        '  --      pati_id                   N   1   病人ID
        '  --      pati_name                 C   1   病人姓名
        '  --      fee_audit_status          N   1   费用审核标志:0或空-未审核;1-已审核或开始审核;2-完成审核
        '  --      si_inp_status             N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
        strJson = "{"
        strJson = strJson & "" & GetJsonNodeString("pati_id", Val(Nvl(rsPati!病人ID)), Json_num)
        strJson = strJson & "," & GetJsonNodeString("pati_name", Nvl(rsPati!姓名), Json_Text)
        strJson = strJson & "," & GetJsonNodeString("fee_audit_status", Val(Nvl(rsPati!审核标志)), Json_num)
        strJson = strJson & "," & GetJsonNodeString("si_inp_status", Val(Nvl(rsPati!住院状态)), Json_num)
        strJson = strJson & "}"
        strPatiJson = strPatiJson & "," & strJson
        rsPati.MoveNext
    Loop
    If strPatiJson <> "" Then strPatiJson = Mid(strPatiJson, 2)
    
    GetPatiJsonByFeeBill = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlFeeExecute_Check(ByVal strNo As String, ByVal str序号 As String, _
    ByVal int费用来源 As Integer, ByVal int单据性质 As Integer, _
    Optional ByRef strSendStuffFeeIDs_Out As String, _
    Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long, _
    Optional ByVal byt自动发料控制 As Byte) As Boolean
    '---------------------------------------------------------------------------------------
    '功能: 执行登记合法性检查
    '入参:
    '   strNO-单据号
    '   str序号-序号
    '   int费用来源-1-门诊;2-住院
    '   int单据性质-1-收费;2-记帐;3-自动记帐
    '   byt自动发料控制-是否执行卫材自动发料，1-是，2-否，0-参数号为63的系统参数控制
    '出参:
    '   strSendStuffFeeIDs_Out-本次执行所涉及自动发放卫材所对应的费用Ids,多个用逗号分离
    '返回:
    '---------------------------------------------------------------------------------------
    Dim strJson As String, str费用IDs As String, strPatiJson As String
    Dim objServiceCall As Object
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    strSendStuffFeeIDs_Out = ""
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strPatiJson = ""
    If int费用来源 = 2 Then
        If GetPatiJsonByFeeBill(int单据性质, strNo, str序号, strPatiJson, int费用来源) = False Then Exit Function
    End If
    
    'Zl_病人费用记录_Execute_Check
    '(
    '  Json_In  Clob,
    '  Json_Out Out Clob
    ') As
    '  ----------------------------------------------------------------------------------------------
    '  --功能：检查费用执行的合法性
    '  --入参：Json_In:格式
    '  --  input
    '  --    fee_no                      C   1   单据号
    '  --    billtype                    N   1   单据类别:1-收费单;2-记帐单;3-自动记账单
    '  --    fee_source                  N   1   费用来源:1-门诊费用;2-住院费用(缺省)
    '  --    serial_nums                 C   0   序号，多个用逗号，如:1,2,3…
    '  --    auto_send_stuff             N   0   是否执行卫材自动发料，1-是，2-否，0-参数号为63的系统参数控制
    '  --    pati_list[]                 病人信息,仅费用来源为2有效
    '  --      pati_id                   N   1   病人ID
    '  --      pati_name                 C   1   病人姓名
    '  --      fee_audit_status          N   1   费用审核标志:0或空-未审核;1-已审核或开始审核;2-完成审核
    '  --      si_inp_status             N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --      code                      N   1   应答码：0-失败；1-成功
    '  --      message                   C   1   应答消息：失败时返回具体的错误信息
    '  --      auto_send_stuffdtl_ids    C   1   发料处方明细IDs:自动发放卫生材料所涉及的费用IDs
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("billtype", int单据性质, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_source", int费用来源, Json_num)
    strJson = strJson & "," & GetJsonNodeString("serial_nums", str序号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("auto_send_stuff", byt自动发料控制, Json_num)
    strJson = strJson & ",""pati_list"":[" & strPatiJson & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_病人费用记录_Execute_Check", strJson, "", strFrmCaption, mlngModule, True) = False Then Exit Function
    strSendStuffFeeIDs_Out = Nvl(objServiceCall.GetJsonNodeValue("output.auto_send_stuffdtl_ids"))  '所涉及的需要自动发料的卫生材料的费用IDs
    zlFeeExecute_Check = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlFeeUnExecute_Check(ByVal strNo As String, ByVal str序号 As String, _
    ByVal int费用来源 As Integer, ByVal int单据性质 As Integer, ByRef strStuffFeeIDs_out As String, _
    Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 取消执行登记前检查
      ' 入参:strNO-单据号
    '      str序号-序号
    '      int费用来源-1-门诊;2-住院
    '      int单据性质-1-收费;2-记帐;3-自动记帐
    ' 出参 :strStuffFeeIDs_out-所涉及的卫生材料费用IDs
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/8/20 21:45
    '---------------------------------------------------------------------------------------
    Dim varReturn As Variant
     
    On Error GoTo errH
    '    Zl_病人费用记录_Unexecute_Chk
    '  单据号_In    住院费用记录.No%Type,
    '  单据性质_In  住院费用记录.记录性质%Type,1-收费;2-记帐;3-自动记帐
    '  费用来源_In  Integer,1-门诊;2-住院
    '  序号_In      Varchar2,费用行的行号，格式:1,2,3...收入项目子行以父行为序号为准,为空表示单据所有行
    '  应签码_In    Out Integer,0-失败；1-成功
    '  应答信息_Out Out Varchar2  应答消息-失败时返回具体的错误信息;1-返回的卫生材料所涉及的费用ids,用逗号分离
    
    varReturn = gobjDatabase.CallProcedure("Zl_病人费用记录_Unexecute_Chk", strFrmCaption, strNo, int单据性质, int费用来源, str序号, Empty, Empty)
    If varReturn(0) <> 1 Then
        If varReturn(1) <> "" Then
            MsgBox varReturn(1), vbOKOnly + vbInformation, gstrSysName
        Else
            MsgBox "取消执行登记检查失败，请检查或稍后重试!", vbOKOnly + vbInformation, gstrSysName
        End If
        Exit Function
    End If
    strStuffFeeIDs_out = varReturn(1)
    zlFeeUnExecute_Check = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPriceChargeCheck(ByVal strNo As String, ByVal lng病人ID As Long, _
    ByVal strFrmCaption As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:划价收费 检查数据合法性
    '入参:
    '   strNo:退费单据号
    '出参:
    '返回:检查通过返回True,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varOut As Variant
    
    On Error GoTo ErrHandler
    'Zl_病人划价收费_Insert_Check(
    '  No_In         In 门诊费用记录.No%Type,
    '  病人id_In     In 门诊费用记录.病人id%Type,
    '  操作员姓名_In In 门诊费用记录.操作员姓名%Type,
    '  应答码_Out    Out Number,0-失败；1-成功；2-失败，用于服务窗并发判断
    '  应答消息_Out  Out Varchar2)
    varOut = zlDatabase.CallProcedure("Zl_病人划价收费_Insert_Check", strFrmCaption, strNo, lng病人ID, UserInfo.姓名, Empty, Empty)
    If varOut(0) <> 1 Then
        MsgBox varOut(1), vbInformation, gstrSysName
        Exit Function
    End If
    
    zlPriceChargeCheck = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlChargeTurnCheck(ByVal strNOs As String, ByVal cllPati As Collection, _
    ByVal strFrmCaption As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊费用转住院 检查数据合法性
    '入参:
    '   strNo:单据信息，格式：记录性质,单据号|记录性质,单据号|... 其中，记录性质：1-门诊收费，2-门诊记帐
    '   cllPati:病人信息：成员:病人ID,主页ID,姓名,审核标志,住院状态
    '出参:
    '   strErrMsg_Out 错误信息
    '返回:检查通过返回True,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    Dim objServiceCall As Object
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_门诊费用转住院_Insert_Check
    '  --功能：门诊费用记录转为住院费用记录前检查数据合法性
    '  --入参：Json_In:格式
    '  --  input
    '  --      fee_nos           C   1   单据信息，格式：记录性质,单据号|记录性质,单据号|... 其中，记录性质：1-门诊收费，2-门诊记帐
    '  --      pati_id           N   1   病人ID
    '  --      pati_pageid       N   1   主页ID
    '  --      pati_name         C   1   病人姓名
    '  --      fee_audit_status  N   1   费用审核标志:0或空-未审核;1-已审核或开始审核(;2-完成审核,结合结帐权限[禁止未审核病人结帐]进行管理控制
    '  --      si_inp_status     N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --      code              N   1   应答码：0-失败；1-成功
    '  --      message           C   1   应答消息：失败时返回具体的错误信息
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_nos", strNOs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_id", cllPati("病人ID"), Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", cllPati("主页ID"), Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_name", cllPati("姓名"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("fee_audit_status", cllPati("审核标志"), Json_num)
    strJson = strJson & "," & GetJsonNodeString("si_inp_status", cllPati("住院状态"), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_门诊费用转住院_Insert_Check", strJson, , strFrmCaption, mlngModule, True) = False Then Exit Function
    
    zlChargeTurnCheck = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mobjServiceCall = Nothing
    Set mobjService = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub

Public Function zlDrugOutRecipeAffirm(ByVal strNOs As String, ByVal byt单据性质 As Byte, _
    ByVal byt单据类型 As Byte, Optional ByVal blnAutoSend As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费成功后药品费用确认，同时自动发药
    '入参:
    '   byt单据性质 1-门诊单据,2-住院单据
    '   byt单据类型 1-收费单;2-记帐单
    '   blnAutoSend 是否自动发药
    '出参:
    '返回:执行成功，返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsData As ADODB.Recordset
    Dim clldata As Collection, lng病人ID As Long
    Dim str操作员编号 As String, str操作员姓名 As String, str操作时间 As String
    Dim blnTrans As Boolean
    Dim strPreNo As String, str费用IDs As String
    Dim str发药窗口 As String '发药窗口1:药房ID1| …|发药窗口n:药房Idn
    Dim strAll费用IDs As String, strTemp As String
    Dim strTable As String, strAutoSendIDs As String
    Dim cllAutoSend As Collection, cllExtend As Collection
    Dim objService As clsService
    Dim str禁止发药费用IDs As String
    
    On Error GoTo ErrHandler
    If GetService(objService) = False Then Exit Function
    
    strTable = IIf(byt单据性质 = 1, "门诊费用记录", "住院费用记录")
    
    strSql = "Select /*+cardinality(b,10)*/a.Id,a.No,a.发药窗口,a.记录状态,a.执行部门ID,a.操作员编号,a.操作员姓名,a.登记时间," & _
            "       a.病人ID,a.姓名,a.性别,a.年龄,a.标识号,a.门诊标志" & _
            " From " & strTable & " A,Table(f_Str2list([2])) B" & _
            " Where a.No=b.Column_Value And a.记录性质=[1] And a.收费类别 In('5','6','7') And Nvl(a.记费同步标志,0)=2" & _
            " Order By a.No,a.序号"
    Set rsData = gobjDatabase.OpenSQLRecord(strSql, "获取费用数据", byt单据类型, strNOs)
    If rsData.EOF Then Exit Function
    
    If blnAutoSend Then '进入输液配药中心的禁止自动发药
        If mobjService.zlPivasSvr_Getinfusion_Record(strNOs, str禁止发药费用IDs) = False Then Exit Function
    End If
    
    With rsData
        lng病人ID = Val(Nvl(!病人ID))
        str操作员编号 = Nvl(!操作员编号)
        str操作员姓名 = Nvl(!操作员姓名)
        str操作时间 = Format(Nvl(!登记时间), "yyyy-MM-dd HH:mm:ss")
        
        Set cllExtend = New Collection
        cllExtend.Add Nvl(!姓名), "姓名"
        cllExtend.Add Nvl(!性别), "性别"
        cllExtend.Add Nvl(!年龄), "年龄"
        cllExtend.Add IIf(Val(Nvl(!门诊标志)) = 2, "", Nvl(!标识号)), "门诊号"
        
        Set cllAutoSend = New Collection
        Set clldata = New Collection: strAll费用IDs = ""
        str费用IDs = "": str发药窗口 = "": strAutoSendIDs = ""
        Do While Not .EOF
            If strPreNo <> Nvl(!NO) Then
                If strPreNo <> "" Then
                    'Array(单据类型,单据号,费用IDs,发药窗口,是否自动发放,自动发放明细IDs)
                    clldata.Add Array(byt单据类型, strPreNo, Mid(str费用IDs, 2), Mid(str发药窗口, 2), _
                        IIf(strAutoSendIDs = "", 0, 1), Mid(strAutoSendIDs, 2))
                        
                    strAll费用IDs = strAll费用IDs & "," & Mid(str费用IDs, 2)
                    If strAutoSendIDs <> "" Then cllAutoSend.Add Array(Mid(strAutoSendIDs, 2), 1)
                End If
                
                strPreNo = Nvl(!NO)
                str费用IDs = "": str发药窗口 = "": strAutoSendIDs = ""
            End If
            
            str费用IDs = str费用IDs & "," & Nvl(!ID)
            
            '发药窗口
            strTemp = Nvl(!发药窗口) & ":" & Nvl(!执行部门ID)
            If Nvl(!发药窗口) <> "" And InStr("|" & str发药窗口 & "|", "|" & strTemp & "|") = 0 Then
                str发药窗口 = str发药窗口 & "|" & strTemp
            End If
            
            '已收费或已记帐审核后自动发药
            If Val(Nvl(!记录状态)) = 1 And blnAutoSend And InStr("," & str禁止发药费用IDs & ",", "," & Nvl(!ID) & ",") = 0 Then
                strAutoSendIDs = strAutoSendIDs & "," & Nvl(!ID)
            End If
            
            rsData.MoveNext
        Loop
    End With
    If strPreNo <> "" Then
        'Array(单据类型,单据号,费用IDs,发药窗口,是否自动发放,自动发放明细IDs)
        clldata.Add Array(byt单据类型, strPreNo, Mid(str费用IDs, 2), Mid(str发药窗口, 2), _
            IIf(strAutoSendIDs = "", 0, 1), Mid(strAutoSendIDs, 2))
        
        strAll费用IDs = strAll费用IDs & "," & Mid(str费用IDs, 2)
        If strAutoSendIDs <> "" Then cllAutoSend.Add Array(Mid(strAutoSendIDs, 2), 1)
    End If
    strAll费用IDs = Mid(strAll费用IDs, 2)
    
    zlDrugOutRecipeAffirm = ExecuteDrugBillAffirm(byt单据性质, clldata, strAll费用IDs, cllAutoSend, _
        gobjDatabase.Currentdate, lng病人ID, cllExtend)
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function ExecuteDrugBillAffirm(ByVal byt单据性质 As Byte, _
    ByVal cllDrugData As Collection, ByVal strDrugFeeIDs As String, _
    ByVal cllSendDrug As Collection, ByVal dt操作时间 As Date, _
    Optional ByVal lng病人ID As Long, Optional ByVal cllExtend As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费、记帐划价审核成功后卫材费用确认，同时自动发料
    '入参:
    '   byt单据性质 1-门诊单据,2-住院单据
    '   cllDrugData 卫材数据，Array(单据类型,单据号,费用IDs,是否自动发放,自动发放明细IDs)
    '   strDrugFeeIDs 费用ID
    '   cllSendDrug 自动发药数据,Array(费用ids,执行状态,已发料
    '   lng病人ID-病人ID，针对收费时有效
    '   cllExtend-扩展信息集，主要针对门诊传入了病人ID时，成员:(姓名,性别,年龄,门诊号)
    '出参:
    '返回:执行成功，返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService
    
    On Error GoTo ErrHandler
    '获取费用服务接口
    If GetService(objService) = False Then Exit Function
    If cllDrugData Is Nothing Then Exit Function
    If cllDrugData.Count = 0 Then Exit Function
    
    gcnOracle.BeginTrans
        If zlExecuteUpdateSyncSymbol(strDrugFeeIDs, 0, byt单据性质, 2) = False Then gcnOracle.RollbackTrans: Exit Function
        If zlUpdateExcuteStatu(cllSendDrug, byt单据性质, , , _
            UserInfo.姓名, Format(dt操作时间, "yyyy-mm-dd HH:MM:SS")) = False Then gcnOracle.RollbackTrans: Exit Function
        If objService.zlDrugSvr_RecipeAffirm(cllDrugData, lng病人ID, _
            Format(dt操作时间, "yyyy-mm-dd HH:MM:SS"), UserInfo.编号, UserInfo.姓名, _
            mlngModule, , cllExtend) = False Then gcnOracle.RollbackTrans: Exit Function
    gcnOracle.CommitTrans
    ExecuteDrugBillAffirm = True
    Exit Function
ErrHandler:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffOutBillAffirm(ByVal strNOs As String, ByVal byt单据性质 As Byte, _
    ByVal byt单据类型 As Byte, Optional ByVal blnAutoSend As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费、记帐划价审核成功后卫材费用确认，同时自动发料
    '入参:
    '   byt单据性质 1-门诊单据,2-住院单据
    '   byt单据类型 1-收费单;2-记帐单
    '   blnAutoSend 是否自动发料
    '出参:
    '返回:执行成功，返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsData As ADODB.Recordset
    Dim clldata As Collection, lng病人ID As Long
    Dim str操作员编号 As String, str操作员姓名 As String, str操作时间 As String
    Dim strPreNo As String, str费用IDs As String
    Dim strAll费用IDs As String
    Dim strTable As String, strAutoSendIDs As String
    Dim cllAutoSend As Collection, cllExtend As Collection
    Dim objService As clsService
    
    On Error GoTo ErrHandler
    '获取费用服务接口
    If GetService(objService) = False Then Exit Function
    
    strTable = IIf(byt单据性质 = 1, "门诊费用记录", "住院费用记录")
    
    strSql = "Select /*+cardinality(b,10)*/a.Id,a.No,a.记录状态,a.操作员编号,a.操作员姓名,a.登记时间," & _
            "       a.病人ID,a.姓名,a.性别,a.年龄,a.标识号,a.门诊标志" & _
            " From " & strTable & " A,Table(f_Str2list([2])) B" & _
            " Where a.No=b.Column_Value And a.记录性质=[1] And a.收费类别='4' And Nvl(a.记费同步标志,0)=2" & _
            " Order By a.No,a.序号"
    Set rsData = gobjDatabase.OpenSQLRecord(strSql, "获取费用数据", byt单据类型, strNOs)
    If rsData.EOF Then Exit Function
     
    With rsData
        lng病人ID = Val(Nvl(!病人ID))
        str操作员编号 = Nvl(!操作员编号)
        str操作员姓名 = Nvl(!操作员姓名)
        str操作时间 = Format(Nvl(!登记时间), "yyyy-MM-dd HH:mm:ss")
        
        Set cllExtend = New Collection
        cllExtend.Add Nvl(!姓名), "姓名"
        cllExtend.Add Nvl(!性别), "性别"
        cllExtend.Add Nvl(!年龄), "年龄"
        cllExtend.Add IIf(Val(Nvl(!门诊标志)) = 2, "", Nvl(!标识号)), "门诊号"
        
        Set cllAutoSend = New Collection
        Set clldata = New Collection: strAll费用IDs = ""
        str费用IDs = "": strAutoSendIDs = ""
        Do While Not .EOF
            If strPreNo <> Nvl(!NO) Then
                If strPreNo <> "" Then
                    'Array(单据类型,单据号,费用IDs,是否自动发放,自动发放明细IDs)
                    clldata.Add Array(byt单据类型, strPreNo, Mid(str费用IDs, 2), _
                        IIf(blnAutoSend, 1, 0), Mid(strAutoSendIDs, 2))
                        
                    strAll费用IDs = strAll费用IDs & "," & Mid(str费用IDs, 2)
                    If strAutoSendIDs <> "" Then cllAutoSend.Add Array(Mid(strAutoSendIDs, 2), 1)
                End If
                
                strPreNo = Nvl(!NO)
                str费用IDs = "": strAutoSendIDs = ""
            End If
            
            str费用IDs = str费用IDs & "," & Nvl(!ID)
            
            '已收费或已记帐审核后自动发料
            If Val(Nvl(!记录状态)) = 1 And blnAutoSend Then strAutoSendIDs = strAutoSendIDs & "," & Nvl(!ID)
            
            rsData.MoveNext
        Loop
    End With
    If strPreNo <> "" Then
        'Array(单据类型,单据号,费用IDs,是否自动发放,自动发放明细IDs)
        clldata.Add Array(byt单据类型, strPreNo, Mid(str费用IDs, 2), _
            IIf(blnAutoSend, 1, 0), Mid(strAutoSendIDs, 2))
            
        strAll费用IDs = strAll费用IDs & "," & Mid(str费用IDs, 2)
        If strAutoSendIDs <> "" Then cllAutoSend.Add Array(Mid(strAutoSendIDs, 2), 1)
    End If
    strAll费用IDs = Mid(strAll费用IDs, 2)
    
     zlStuffOutBillAffirm = ExecuteStuffBillAffirm(byt单据性质, clldata, strAll费用IDs, cllAutoSend, _
        gobjDatabase.Currentdate, lng病人ID, cllExtend)
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function ExecuteStuffBillAffirm(ByVal byt单据性质 As Byte, _
    ByVal cllStuffData As Collection, ByVal strStuffFeeIDs As String, _
    ByVal cllSendStuff As Collection, ByVal dt操作时间 As Date, _
    Optional ByVal lng病人ID As Long, Optional ByVal cllExtend As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费、记帐划价审核成功后卫材费用确认，同时自动发料
    '入参:
    '   byt单据性质 1-门诊单据,2-住院单据
    '   cllStuffData 卫材数据，Array(单据类型,单据号,费用IDs,是否自动发放,自动发放明细IDs)
    '   strStuffFeeIDs 费用ID
    '   cllSendStuff 自动发药数据,Array(费用ids,执行状态,已发料
    '   lng病人ID-病人ID，针对收费时有效
    '   cllExtend-扩展信息集，主要针对门诊传入了病人ID时，成员:(姓名,性别,年龄,门诊号)
    '出参:
    '返回:执行成功，返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService
    
    On Error GoTo ErrHandler
    '获取费用服务接口
    If GetService(objService) = False Then Exit Function
    If cllStuffData Is Nothing Then Exit Function
    If cllStuffData.Count = 0 Then Exit Function
    
    gcnOracle.BeginTrans
        If zlExecuteUpdateSyncSymbol(strStuffFeeIDs, 0, byt单据性质, 2) = False Then gcnOracle.RollbackTrans: Exit Function
        If zlUpdateExcuteStatu(cllSendStuff, byt单据性质, , , _
            UserInfo.姓名, Format(dt操作时间, "yyyy-mm-dd HH:MM:SS")) = False Then gcnOracle.RollbackTrans: Exit Function
        If objService.zlStuffSvr_BillAffirm(cllStuffData, lng病人ID, _
            Format(dt操作时间, "yyyy-mm-dd HH:MM:SS"), UserInfo.编号, UserInfo.姓名, _
            mlngModule, , cllExtend) = False Then gcnOracle.RollbackTrans: Exit Function
    gcnOracle.CommitTrans
    ExecuteStuffBillAffirm = True
    Exit Function
ErrHandler:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function Get发药窗口_Batch(ByVal strNo As String, ByVal strSerialNums As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:给门诊记帐单分配发药窗口
    '入参:
    '   strNO 记帐单据号
    '   strSerialNums 费用序号，格式：序号1,序号2,...
    '出参:
    '返回:已分配的发药窗口，格式：库房ID:发药窗口,库房2:发药窗口,...
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsData As ADODB.Recordset, rsWindows As ADODB.Recordset
    Dim lng病人ID As Long, str药房IDs As String
    Dim strWhere As String, str发药窗口 As String
    
    On Error GoTo ErrHandler
    If strSerialNums <> "" Then strWhere = " And Instr(','||[2]||',',','||Nvl(a.价格父号,a.序号)|| ',') > 0"
    strSql = "Select Distinct a.病人id, a.执行部门id" & _
            " From 门诊费用记录 A" & _
            " Where a.记录性质 = 2 And a.No = [1] And a.收费类别 In ('5', '6', '7')" & _
                    strWhere
    Set rsData = gobjDatabase.OpenSQLRecord(strSql, "获取执行科室", strNo, strSerialNums)
    If rsData.EOF Then Exit Function
    
    str药房IDs = ""
    Do While Not rsData.EOF
        lng病人ID = Nvl(rsData!病人ID)
        str药房IDs = str药房IDs & "," & Nvl(rsData!执行部门ID)
        rsData.MoveNext
    Loop
    str药房IDs = Mid(str药房IDs, 2)
    
    
    Set rsWindows = Get发药窗口(2, lng病人ID, str药房IDs, mlngModule)
    If rsWindows Is Nothing Then Exit Function
    str发药窗口 = ""
    Do While Not rsWindows.EOF
        str发药窗口 = str发药窗口 & "," & Nvl(rsWindows!药房Id)
        str发药窗口 = str发药窗口 & ":" & Nvl(rsWindows!发药窗口)
        rsWindows.MoveNext
    Loop
    Get发药窗口_Batch = Mid(str发药窗口, 2)
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlGetAdviceBillingState(ByVal bytBillType As Byte, ByVal strNOs As String, _
    ByVal bln门诊 As Boolean, ByVal byt操作类型 As Byte, _
    ByRef cllAdviceData As Collection, Optional ByVal lngModule As Long, _
    Optional ByVal cllDeleteLists As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱计费状态
    '入参:
    '   bytBillType 单据类型: 1-收费单;2-记帐单;3-自动记帐单
    '   strNos -单据信息, 格式：NO1:序号1,序号2,...|NO1:序号1,序号2,...|...
    '   byt操作类型 操作类型：0-删除划价单;1-收费或记帐;2-退费或销帐
    '   cllDeleteLists 集合,退费数据
    '       |-cllDeleteList 集合,销账数据,操作类型为0或2时必须传=cllDeleteLists("_NO")
    '           |-元素集合：序号,销帐数量(为零时，按序号直接销帐)
    '出参:
    '   cllAdviceData 集合,含计费状态的医嘱数据
    '       |-元素集合：医嘱ID,费用单号,计费状态,删除附费
    '返回:获取成功返回True,获取失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNos As Variant, i As Long, j As Long, strSql As String
    Dim strSendDrugWindows As String, varData As Variant
    Dim objService As clsService
    Dim str医嘱IDs As String, rsAdviceData As ADODB.Recordset
    Dim cllAdviceList_out As Collection, cllTemp As Collection
    Dim clldata As Collection, cllAdviceList As Collection
    Dim cllDeleteList As Collection
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetService(objService) = False Then Exit Function
    
    Set cllAdviceData = New Collection
    varNos = Split(strNOs, "|")
    For i = 0 To UBound(varNos)
        varData = Split(varNos(i) & ":", ":")
        
        '计算计费状态
        str医嘱IDs = ""
        If CheckExistAdviceFeeByDetail(varNos(i), bytBillType, bln门诊, str医嘱IDs, True) = False Then Exit Function
        If str医嘱IDs <> "" Then
            If objService.ZlCissvr_GetAllGroupAdviceIDs(str医嘱IDs, rsAdviceData, mlngModule) = False Then Exit Function
            Set cllAdviceList = New Collection
            Do While Not rsAdviceData.EOF
                Set cllTemp = New Collection
                cllTemp.Add Nvl(rsAdviceData!医嘱id), "医嘱ID"
                cllTemp.Add Nvl(rsAdviceData!相关ID), "相关ID"
                cllTemp.Add Nvl(rsAdviceData!诊疗类别), "诊疗类别"
                cllAdviceList.Add cllTemp
                rsAdviceData.MoveNext
            Loop
            
            '销账数据
            Set cllDeleteList = Nothing
            If Not cllDeleteLists Is Nothing Then Set cllDeleteList = cllDeleteLists("_" & varData(0))
            
            If Zl_医嘱发送_计费状态_Get(bytBillType, varData(0), bln门诊, byt操作类型, _
                str医嘱IDs, cllDeleteList, cllAdviceList, cllAdviceList_out, mlngModule) = False Then Exit Function
            If Not cllAdviceList_out Is Nothing Then
                For j = 1 To cllAdviceList_out.Count
                    Set clldata = cllAdviceList_out(j)
                    Set cllTemp = New Collection
                    cllTemp.Add clldata("医嘱ID"), "医嘱ID"
                    cllTemp.Add varData(0), "费用单号"
                    cllTemp.Add clldata("计费状态"), "计费状态"
                    cllTemp.Add 0, "删除附费"
                    cllAdviceData.Add cllTemp
                Next
            End If
        End If
    Next
    ZlGetAdviceBillingState = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlVerfyBillingPriceBill(ByVal byt病人来源 As Byte, ByVal strNOs As String, _
   Optional ByVal strVerfyDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记帐划价单审核为记帐单
    '入参:byt病人来源:1-门诊;2-住院
    '    strNos -单据信息, 格式：NO1:序号1,序号2,...|NO1:序号1,序号2,...|...
    '    strVerfyDate-审核日期(为空时，以当前时间为准):格式：yyyy-mm-dd HH:MM:SS
    '出参:
    '返回:执行成功返回True,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNos As Variant, i As Long, j As Long, strSql As String
    Dim strSendDrugWindows As String, varData As Variant
    Dim blnTrans As Boolean, cllOutPro As Collection
    Dim objService As clsService
    Dim cllAdviceDatas As Collection
    
    On Error GoTo errH
    Set cllOutPro = New Collection
    If GetService(objService) = False Then Exit Function
    
    If strVerfyDate = "" Or Not IsDate(strVerfyDate) Then
        strVerfyDate = Format(gobjDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        strVerfyDate = Format(CDate(strVerfyDate), "yyyy-mm-dd HH:MM:SS")
    End If
    
    '计算医嘱计费状态
    If ZlGetAdviceBillingState(Val("2-记帐单"), strNOs, (byt病人来源 = 1), _
        Val("1-收费或记帐"), cllAdviceDatas, mlngModule) = False Then Exit Function
        
    varNos = Split(strNOs, "|")
    For i = 0 To UBound(varNos)
        varData = Split(varNos(i) & ":", ":")
        
        If byt病人来源 = 1 Then
            strSendDrugWindows = Get发药窗口_Batch(varData(0), varData(1))
            'Zl_门诊记帐记录_Verify_S
            strSql = "Zl_门诊记帐记录_Verify_S("
            '  No_In         门诊费用记录.No%Type,
            strSql = strSql & "'" & varData(0) & "',"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '  序号_In       Varchar2 := Null,--格式如"1,3,5,7,8",为空表示审核所有未审核的行
            strSql = strSql & "'" & varData(1) & "',"
            '  审核时间_In   门诊费用记录.登记时间%Type := Null
            strSql = strSql & "to_date('" & strVerfyDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '  发药窗口_In   Varchar2 := Null--库房ID:窗口,库房2:发药窗口,....
            strSql = strSql & "'" & strSendDrugWindows & "',"
            '  更新同步标志_In     Number := 1
            strSql = strSql & "" & 1 & ")"
            
        Else
            'Zl_住院记帐记录_Verify_S
            strSql = "Zl_住院记帐记录_Verify_S("
            '  No_In         住院费用记录.No%Type,
            strSql = strSql & "'" & varData(0) & "',"
            '  操作员编号_In 住院费用记录.操作员编号%Type,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '  序号_In       Varchar2 := Null,--格式如"1,3,5,7,8",为空表示审核所有未审核的行
            strSql = strSql & "'" & varData(1) & "',"
            '  病人id_In     住院费用记录.病人id%Type := Null,
            strSql = strSql & "NULL,"
            '  审核时间_In   住院费用记录.登记时间%Type := Null
            strSql = strSql & "to_date('" & strVerfyDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '  更新同步标志_In     Number := 1
            strSql = strSql & "" & 1 & ")"
        End If
        Call zlAddArray(cllOutPro, strSql)
    Next
    
    blnTrans = True
    zlExecuteProcedureArrAy cllOutPro, "执行划价审核", True, False
    
    '更改医嘱数据的计费状态
    If cllAdviceDatas.Count > 0 Then
        If objService.ZlCisSvr_UpdAdviceChargeTag(Val("2-记帐单"), cllAdviceDatas, Nothing, mlngModule) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    End If
    
    gcnOracle.CommitTrans
    zlVerfyBillingPriceBill = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDelPriceBillCheck(ByVal strNOs As String, Optional ByRef cllBillLists As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除划价单前调用 Zl_门诊划价记录_Delete_Check 检查数据合法性
    '入参:
    '   strNos:划价单据信息,格式：A001:序号1|序号2|...,A002:序号1|序号2|...,...
    '出参:
    '   cllBillLists:删除执行入参数据
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBillList As Collection, cllBillDetails As Collection, cllBillDetail As Collection
    Dim varNos As Variant, i As Integer, j As Integer
    Dim varNo As Variant, varNums As Variant, strNo As String, strNums As String
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    
    On Error GoTo ErrHandler
    Set cllBillLists = New Collection
    
    varNos = Split(strNOs, ",")
    For i = 0 To UBound(varNos)
        varNo = Split(varNos(i) & ":", ":")
        strNo = varNo(0): strNums = varNo(1)
        
        varNums = Split(strNums, "|")
        Set cllBillDetails = New Collection
        For j = 0 To UBound(varNums)
            Set cllBillDetail = New Collection
            cllBillDetail.Add varNums(j), "序号"
            cllBillDetails.Add cllBillDetail, "_" & varNums(j)
        Next
    
        Set cllBillList = New Collection
        cllBillList.Add strNo, "单据号"
        cllBillList.Add cllBillDetails, "_cllBillDetails"
        cllBillLists.Add cllBillList, "_" & strNo
    Next
    
    If zlExcute_DelRecipeBill_Check(cllBillLists, 1, mlngModule, True, True) = False Then Exit Function
    
    zlDelPriceBillCheck = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDelPriceBill(ByVal cllBillLists As Collection, Optional ByVal blnExcute As Boolean = True, Optional ByRef cllPro_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除划价单
    '入参:
    '   cllBillLists:删除执行入参数据
    '   blnExcute:立即执行作废过程
    '       ↓cllPro：如果不立即执行，则返回作废过程集合
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService
    Dim strJsonDelDrugRcp_out As String, strJsonDelStuffRcp_out As String
    Dim strDrugSendFeeIDs_Out  As String, strStuffSendFeeIDs_Out As String
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    
    On Error GoTo ErrHandler
    '获取费用服务接口
    If GetService(objService) = False Then Exit Function
    
    If blnExcute Then
        If zlExcute_DelRecipeBill(cllBillLists, "zlDelPriceBill", 1, True, mlngModule, True) = False Then Exit Function
    Else
        If cllPro_out Is Nothing Then Set cllPro_out = New Collection
        If zlGetDelBill_SaveData(cllBillLists, cllPro_out, _
            strJsonDelDrugRcp_out, strJsonDelStuffRcp_out, _
            strDrugSendFeeIDs_Out, strStuffSendFeeIDs_Out, 1, True, True) = False Then Exit Function
    End If
    
    zlDelPriceBill = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlAdjustFeeData(ByVal strNOs As String, Optional ByVal blnDel As Boolean, _
    Optional ByVal blnShowErrMsg As Boolean = True, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊费用转住院后调整关联费用信息
    '入参:
    '   strNos 住院费用记录的单据号,多个用英文的逗号分隔
    '   blnDel 是否销账前同步数据，此时必须同步成功才能进行下一步操作
    '   blnShowErrMsg 是否显示错误信息
    '出参:
    '   strErrMsg blnShowErrMsg=False时返回错误信息
    '返回:成功，返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim lng病人ID As Long, lng主页ID As Long, byt单据类型 As Byte
    Dim strSubTable As String, varPara() As Variant
    Dim i  As Long, var单据类型 As Variant
    
    On Error GoTo ErrHandler
    If zlGetVarBoundSQL(1, strNOs, strSubTable, varPara, 0) = False Then Exit Function
    strSql = _
            " Select a.病人ID, a.主页ID, b.记录性质 As 单据类型, a.收费类别," & _
            "              b.No As 原单据号, b.Id As 原费用id, a.No As 新单据号, a.Id As 新费用id" & _
            " From 住院费用记录 A, 门诊费用记录 B, 费用审核记录 C, (" & strSubTable & ") B, 材料特性 M" & _
            " Where a.Id = c.转出id And b.Id = c.费用id" & _
            "       And a.No = b.Column_Value And a.记录性质 = 2 And a.收费细目id = m.材料id(+)" & _
            "       And (a.收费类别 In ('5','6','7') Or a.收费类别 = '4' And Nvl(m.跟踪在用, 0) = 1)" & _
            "       And Nvl(a.转费同步标志,0)<>0 And Nvl(a.记费同步标志,0)<>1"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSql, "获取费用数据", varPara)
    If rsTemp.RecordCount = 0 Then zlAdjustFeeData = True: Exit Function
    
    var单据类型 = Array(1, 2) '单据类型：1-门诊收费单,2-门诊记帐单
    For i = 0 To UBound(var单据类型)
        rsTemp.Filter = "单据类型=" & var单据类型(i)
        If rsTemp.RecordCount > 0 Then
            lng病人ID = Val(Nvl(rsTemp!病人ID))
            lng主页ID = Val(Nvl(rsTemp!主页ID))
            byt单据类型 = Val(Nvl(rsTemp!单据类型))
            If zlAdjustFeeDataDrugStuff(lng病人ID, lng主页ID, byt单据类型, rsTemp, False, blnShowErrMsg, strErrMsg) = False And blnDel Then Exit Function
            If zlAdjustFeeDataDrugStuff(lng病人ID, lng主页ID, byt单据类型, rsTemp, True, blnShowErrMsg, strErrMsg) = False And blnDel Then Exit Function
        End If
    Next
    zlAdjustFeeData = True
    Exit Function
ErrHandler:
    If blnShowErrMsg = False Then strErrMsg = Err.Description: Exit Function
    
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function zlAdjustFeeDataDrugStuff(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal byt单据类型 As Byte, _
    ByVal rsData As ADODB.Recordset, ByVal bln卫材 As Boolean, _
    Optional ByVal blnShowErrMsg As Boolean = True, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊费用转住院后调整药品、卫材关联费用信息
    '入参:
    '   byt单据类型 1-门诊收费单,2-门诊记帐单
    '   rsData 费用数据，字段：收费类别,原单据号,原费用id,新单据号,新费用id
    '   blnShowErrMsg 是否显示错误信息
    '出参:
    '   strErrMsg blnShowErrMsg=False时返回错误信息
    '返回:成功，返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTrans As Boolean
    Dim cllUpdataDatas As Collection, cllUpdataData As Collection, strFeeIds As String
    Dim objService As clsService
    Dim strFliter As String
    
    '   cllUpdataDatas(collect)-数据集，格式如下
    '     |-cllUpdataData(_新费用ID)(collect)每行明细数据集，成员：(原单据号,原费用ID,新单据号,新费用ID)
    
    On Error GoTo ErrHandler
    
    '获取费用服务接口
    If GetService(objService) = False Then Exit Function
    If rsData Is Nothing Then Exit Function
    
    If bln卫材 Then
        strFliter = "单据类型=" & byt单据类型 & " And 收费类别='4'"
    Else
        strFliter = "(单据类型=" & byt单据类型 & " And 收费类别='5')"
        strFliter = strFliter & " Or (单据类型=" & byt单据类型 & " And 收费类别='6')"
        strFliter = strFliter & " Or (单据类型=" & byt单据类型 & " And 收费类别='7')"
    End If
    
    rsData.Filter = strFliter
    If rsData.EOF Then zlAdjustFeeDataDrugStuff = True: Exit Function
    
    strFeeIds = ""
    Set cllUpdataDatas = New Collection
    Do While Not rsData.EOF
        strFeeIds = strFeeIds & "," & Val(Nvl(rsData!新费用ID))
        
        Set cllUpdataData = New Collection
        cllUpdataData.Add Nvl(rsData!原单据号), "原单据号"
        cllUpdataData.Add Val(Nvl(rsData!原费用ID)), "原费用ID"
        cllUpdataData.Add Nvl(rsData!新单据号), "新单据号"
        cllUpdataData.Add Val(Nvl(rsData!新费用ID)), "新费用ID"
        cllUpdataDatas.Add cllUpdataData, "_" & Val(Nvl(rsData!新费用ID))
        
        rsData.MoveNext
    Loop
    If strFeeIds <> "" Then strFeeIds = Mid(strFeeIds, 2)
    
    gcnOracle.BeginTrans: blnTrans = True
        If zlExecuteUpdateSyncSymbol(strFeeIds, 2, 2, 1) = False Then gcnOracle.RollbackTrans: Exit Function
        If bln卫材 Then
            If objService.zlStuffSvr_AdjustFeeData(lng病人ID, lng主页ID, byt单据类型, cllUpdataDatas, _
                mlngModule, blnShowErrMsg, strErrMsg) = False Then gcnOracle.RollbackTrans: Exit Function
        Else
            If objService.zlDrugsvr_AdjustFeeData(lng病人ID, lng主页ID, byt单据类型, cllUpdataDatas, _
                mlngModule, blnShowErrMsg, strErrMsg) = False Then gcnOracle.RollbackTrans: Exit Function
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    zlAdjustFeeDataDrugStuff = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If blnShowErrMsg = False Then strErrMsg = Err.Description: Exit Function
    
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlUpdateExcuteStatu(ByVal clldata As Variant, _
    Optional ByVal int费用来源 As Integer, Optional ByVal blnAutoCalc As Boolean = False, Optional ByVal str执行情况 As String, _
    Optional str操作员姓名 As String, Optional str操作时间 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改病人费用记录的执行状态
    '入参:cllData-更新的数据集(array(费用ids,执行状态,已发料),自动计算时，需要传入已发料，费用ids只能为一个费用ID,执行状态缺省为0
    '     int费用来源-1-门诊;2-住院;0-不区分门诊或住院
    '     blnAutoCalc-是否自动计算执行状态
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJsonTemp As String, strJson As String
    Dim varData As Variant, i As Long
    
    On Error GoTo ErrHandler
    If clldata Is Nothing Then zlUpdateExcuteStatu = True: Exit Function
    If clldata.Count = 0 Then zlUpdateExcuteStatu = True: Exit Function
    
    'Zl_费用记录_执行状态_Update
    '  --功能：更新费用记录的执行状态
    '  --入参：Json_In:格式
    '  --  input
    '  --    fee_source          N   1-门诊费用;2-住院费用;0-不区分门诊或住院
    '  --    auto_calc           N   是否自动根据已执行数计算执行状态并进行修正:1-自动;0-按传入的为准
    '        operator_time                   C       操作时间:yyyy-mm-dd hh24:mi:ss
    '        operator_name                   C       操作员姓名
    '        execute_note                    C       执行情况说明
    '  --    item_list[]
    '  --        fee_ids         C   1   费用IDs，如果需要自动修正，则只能传入一个费用ID
    '  --        execute_statu   N   1   执行状态:0:未执行;1:完全执行;2:部份执(Auto_Calc=1时，无此接点)
    '            execute_num                         执行数(药品卫材的已发料数)(Auto_Calc=0时，无此接点)
    
    '可能存在只修正记录状态的情况，所以不能取缺省值，需要调用者传入
    'If str操作员姓名 = "" Then str操作员姓名 = UserInfo.姓名
    'If str操作时间 = "" Then str操作时间 = Format(gobjDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    strJson = ""
    For i = 1 To clldata.Count
        varData = clldata(i)
        strJsonTemp = ""
        strJsonTemp = strJsonTemp & "" & GetJsonNodeString("fee_ids", CStr(varData(0)), Json_Text)
        If blnAutoCalc Then
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("execute_num", Val(varData(2)), Json_num)
        Else
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("execute_statu", Val(varData(1)), Json_num)
        End If
        strJson = strJson & ",{" & strJsonTemp & "}"
    Next
    strJsonTemp = Mid(strJson, 2)
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_source", int费用来源, Json_num)
    strJson = strJson & "," & GetJsonNodeString("auto_calc", IIf(blnAutoCalc, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("operator_name", str操作员姓名, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_time", str操作时间, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("execute_note", str执行情况, Json_Text)
    strJson = strJson & ",""item_list"":[" & strJsonTemp & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    Call zlDatabase.CallProcedure("Zl_费用记录_执行状态_Update", "更新执行状态", strJson)
    zlUpdateExcuteStatu = True
    Exit Function
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function zlExecuteUpdateSyncSymbol(ByVal str费用IDs As String, ByVal byt标志类型 As Byte, _
    ByVal byt门诊标志 As Byte, ByVal byt原标志值 As Byte, Optional ByVal byt新标志值 As Byte = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新费用记录中的药品/卫材同步标志
    '入参:
    '   str费用IDs 需要更新的费用ID，多个用英文逗号分隔
    '   byt标志类型 标志类型：0-记费同步标志,1-作废同步标志,2-转费同步标志
    '   byt门诊标志 门诊标志：1-门诊，2-住院
    '出参:
    '返回:
    '  记费同步标志：NULL或0-正常，1-未生成处方单/发料单，2-未更新药品/卫材收费状态
    '  作废同步标志：NULL或0-正常，1-药品/卫材已作废但费用未作废(禁止发放/退回)
    '  转费同步标志：NULL或0-正常，1-门诊费用已转到住院，药品/卫材数据未同步
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    
    On Error GoTo errHandle
    'Zl_病人收费记录_更新同步标志
    '  --功能：更新费用记录中的药品/卫材同步标志
    '  --入参：Json_In:格式
    '  --  input
    '  --    fee_ids             C 1 费用IDs
    '  --    outpati_flag        N 1 门诊标志：1-门诊，2-住院
    '  --    sign_type           N 1 标志类型：0-记费同步标志,1-作废同步标志,2-转费同步标志
    '  --    old_sign            N 1 原标志
    '  --    new_sign            N 1 新标志
    '  --说明：
    '  --    记费同步标志：NULL或0-正常，1-未生成处方单/发料单，2-未更新药品/卫材收费状态
    '  --    作废同步标志：NULL或0-正常，1-药品/卫材已作废但费用未作废(禁止发放/退回)
    '  --    转费同步标志：NULL或0-正常，1-门诊费用已转到住院，药品/卫材数据未同步
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_ids", str费用IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("outpati_flag", byt门诊标志, Json_num)
    strJson = strJson & "," & GetJsonNodeString("sign_type", byt标志类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("old_sign", byt原标志值, Json_num)
    strJson = strJson & "," & GetJsonNodeString("new_sign", byt新标志值, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    Call zlDatabase.CallProcedure("Zl_病人收费记录_更新同步标志", "更新同步标志", strJson)
    zlExecuteUpdateSyncSymbol = True
    Exit Function
errHandle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function zlExcute_SaveRecipeBill_Check(cllRcpBillData As Collection, ByVal intBillType As Integer, _
    Optional ByRef cllDrawDeptids_Out As Collection, _
    Optional ByVal bln划价 As Boolean, Optional ByVal bln门诊 As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行处方保存前检查操作
    '入参:
    ' cllRcpBillData(结构)
    '   |-cllPati(Collect):病人信息，成员：
    '     (病人ID,主页ID,姓名,性别,年龄,费别,床号,病人科室ID,病区ID,险类,
    '      险类名称,病人来源,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '      患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '      身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):单据信息集=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):单据信息，成员：
    '       (单据号,是否多病人单,医疗小组id,医技补临床费用,简单记帐,记帐单id,记录性质,是否划价,是否急诊,加班标志,开单科室ID,
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,是否附费(N),[cllBillDetails(collect)])=cllBillists(_单据号)
    '       |-cllBillDetails(Collect):单据明细集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,从属父号,药名ID,收费细目ID,价格父号,收入项目ID,费别,婴儿序号,收费类别,计算单位,是否保险项目,保险大类ID,
    '            保险编码,收据费目,付数,数次,单价,应收金额,实收金额,统筹金额,附加标志,执行部门ID,是否自动发放,费用摘要(N),
    '            医嘱ID(N),就诊ID(N),费用类型,中药形态,煎法,执行性质,是否备货卫材(N),备货材料批次(N),是否跟踪在用,发药窗口(N),
    '            组内序号(N),诊疗项目ID(N),给药途径ID(N),给药途径名称(N),给药途径分类(N),给药频次ID(N)，给药频次名称（N),
    '            医嘱紧急标志(N),医嘱期效(N),计价特性(N),频次(N),单量（N),用法(N),皮试结果(N),超量说明(N),使用嘱托(N),
    '            发药方式(N),药品含量(N),门诊执行天数(N),煎法(N)
    '            【记帐表增加】:(病人ID,主页ID,姓名,性别,年龄,床号,病人科室ID,病区ID,险类,
    '              险类名称,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '              病人来源,患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '              身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N),
    '              医疗小组id)=cllBillDetails(_序号)
    '    以上元素中，带(N)的，表示可选节点
    '  intBillType-单据类型(1-收费单;2-记帐单;3-记帐表)
    '  bln划价-是否划价单
    '出参:
    '   cllDrawDeptIds_Out-领药部门ID集((病人ID,领药部门ID),_病人ID)
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-14 14:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJsonFeeCheck_out As String, strJsonDrugCheck_Out As String, strJsonStuffCheck_out As String
    Dim objService As clsService
                
    On Error GoTo errHandle
    If lngModule <> 0 Then mlngModule = lngModule
    If GetService(objService) = False Then Exit Function
    
    '先合法性数据检查
    '1.先组织相关数据
    If zlGetNewBill_CheckData(cllRcpBillData, strJsonFeeCheck_out, strJsonDrugCheck_Out, strJsonStuffCheck_out, _
        intBillType, bln门诊) = False Then Exit Function
    '2.先进行数据合法性检查
    '2.1 费用数据合法性较验
    If zlFeeNewRecipeBill_Check(strJsonFeeCheck_out, intBillType, bln划价, cllDrawDeptids_Out, bln门诊) = False Then Exit Function
    '2.2 药品生成处方检查
    If strJsonDrugCheck_Out <> "" Then
        If objService.zlDrugSvr_NewRecipeBill_Check(strJsonDrugCheck_Out) = False Then Exit Function
    End If
    '2.3 卫材生成处方检查
    If strJsonStuffCheck_out <> "" Then
        If objService.zlStuffSvr_NewBill_Check(strJsonStuffCheck_out) = False Then Exit Function
    End If
    
    zlExcute_SaveRecipeBill_Check = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlFeeNewRecipeBill_Check(ByVal strJsonAsk As String, ByVal intBillType As Integer, _
    Optional ByVal bln划价 As Boolean, _
    Optional ByRef cllData_out As Collection, Optional ByVal bln门诊 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:住院记帐审核合法性检查
    '入参:strJsonAsk-请求的Json串
    '     intBillType-1-收费处方;2- 记帐单处方;3- 记帐表处方
    '     bln划价-是否划价
    '出参:
    '    cllData_out -返回的数据集
    '             intBillType=2,3时:每个元素为一个集合(Collection),每个元素的集合为(Zl_住院记帐记录_Insert_Check)的出参的各元素
    '             item_list[]                         数据组
    '                 pati_id             N   1    病人ID
    '                 takedept_id             N   1   领用部门ID
    '返回:记帐合法，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    
    Err = 0: On Error GoTo errHandle
    
    If strJsonAsk = "" Then
        If bln门诊 Then zlFeeNewRecipeBill_Check = True: Exit Function
        MsgBox "未传入本次记帐的病人信息，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    Select Case intBillType
    Case 1 '收费单
    Case 2, 3 '记账单
        If Not bln门诊 Then
            If objServiceCall.CallService("Zl_住院记帐记录_Insert_Check", strJsonAsk, , "", mlngModule, True) = False Then Exit Function
            Set cllData_out = objServiceCall.GetJsonListValue("output.item_list", "pati_id,takedept_id")
        End If
    Case Else   '其他
    
    End Select
    zlFeeNewRecipeBill_Check = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExcute_SaveRecipeBill(ByVal cllRctBillData As Collection, _
    ByVal strFrmCaption As String, ByVal intBillType As Integer, Optional ByVal bln划价 As Boolean, _
    Optional ByVal cllDrawDeptIds As Collection, Optional bln门诊 As Boolean, _
    Optional cllPro As Collection, Optional blnSendMateria As Boolean, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行处方数据保存操作
    '入参:
    ' cllRcpBillData(结构)
    '   |-cllPati(Collect):病人信息，成员：
    '     (病人ID,主页ID,姓名,性别,年龄,费别,床号,病人科室ID,病区ID,险类,
    '      险类名称,病人来源,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '      患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '      身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):单据信息集=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):单据信息，成员：
    '       (单据号,是否多病人单,医疗小组id,医技补临床费用,简单记帐,记帐单id,记录性质,是否划价,是否急诊,加班标志,开单科室ID,
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,是否附费(N),[cllBillDetails(collect)])=cllBillists(_单据号)
    '       |-cllBillDetails(Collect):单据明细集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,从属父号,药名ID,收费细目ID,价格父号,收入项目ID,费别,婴儿序号,收费类别,计算单位,是否保险项目,保险大类ID,
    '            保险编码,收据费目,付数,数次,单价,应收金额,实收金额,统筹金额,附加标志,执行部门ID,是否自动发放,费用摘要(N),
    '            医嘱ID(N),就诊ID(N),费用类型,中药形态,煎法,执行性质,是否备货卫材(N),备货材料批次(N),是否跟踪在用,发药窗口(N),
    '            组内序号(N),诊疗项目ID(N),给药途径ID(N),给药途径名称(N),给药途径分类(N),给药频次ID(N)，给药频次名称（N),
    '            医嘱紧急标志(N),医嘱期效(N),计价特性(N),频次(N),单量（N),用法(N),皮试结果(N),超量说明(N),使用嘱托(N),
    '            发药方式(N),药品含量(N),门诊执行天数(N),煎法(N)
    '            【记帐表增加】:(病人ID,主页ID,姓名,性别,年龄,床号,病人科室ID,病区ID,险类,
    '              险类名称,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '              病人来源,患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '              身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N),
    '              医疗小组id)=cllBillDetails(_序号)
    '    以上元素中，带(N)的，表示可选节点
    '  intBillType-单据类型(1-收费单;2-记帐单;3-记帐表)
    '  cllDrawDeptIds-  领药部门ID集((病人ID,领药部门ID),_病人ID)
    '  cllPro 需要一起执行的SQL语句
    '出参:blnSendMateria:记帐后自动发药
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-13 22:52:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJosnNewDrugRcp_out As String, strJosnNewStuffRcp_out As String
    Dim i As Long, strInfo As String
    Dim strDrugSendFeeIDs_Out  As String, strStuffSendFeeIDs_Out As String
    Dim strAutoSendDrugSendFeeIDs  As String, strAutoSendStuffSendFeeIDs As String
    Dim blnTrans As Boolean
    Dim cllBillLists As Collection, cllBillList As Collection
    Dim cllPatient As Collection, strNo As String, strInsure As String
    Dim cllBillDetails As Collection, cllBillDetail As Collection
    
    On Error GoTo errHandle
    If lngModule <> 0 Then mlngModule = lngModule
    '1.先收集费用保存数据过程及生成处方数据信息
    If zlGetNewBill_SaveData(cllRctBillData, cllPro, strJosnNewDrugRcp_out, strJosnNewStuffRcp_out, _
        strDrugSendFeeIDs_Out, strStuffSendFeeIDs_Out, intBillType, bln划价, cllDrawDeptIds, bln门诊, _
        strAutoSendDrugSendFeeIDs, strAutoSendStuffSendFeeIDs) = False Then Exit Function
    If cllPro.Count = 0 Then Exit Function
    
    Set cllBillLists = cllRctBillData("_cllBillLists")
    Set cllPatient = cllRctBillData("_patiinfor")
    Set cllBillList = cllBillLists(1)
    strNo = cllBillList("单据号")
    
    '收集险类数据
    strInsure = ""
    If intBillType <> 1 And Not bln划价 And Not bln门诊 Then
        If intBillType = 3 Then '多病人单
            Set cllBillDetails = cllBillList("_cllBillDetails")
            For i = 1 To cllBillDetails.Count
                Set cllBillDetail = cllBillDetails(i)
                If Val(cllBillDetail("险类")) <> 0 And InStr("," & strInsure & ",", "," & Val(cllBillDetail("险类")) & ",") = 0 Then
                    strInsure = strInsure & "," & Val(cllBillDetail("险类"))
                End If
            Next
            If strInsure <> "" Then strInsure = Mid(strInsure, 2)
        ElseIf Val(cllPatient("险类")) <> 0 Then
            strInsure = Val(cllPatient("险类"))
        End If
    End If
    
    '2.执行相关费用数据保存
    gcnOracle.BeginTrans: blnTrans = True
        zlExecuteProcedureArrAy cllPro, strFrmCaption, True, True
        
        '医保记帐完成前上传
        If strInsure <> "" Then
            If InsureTranChargeDetail(strNo, strInsure, strInfo, False) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                If strInfo <> "" Then MsgBox strInfo, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '3.存在药品，需要生成处方
    If strJosnNewDrugRcp_out <> "" Then
        Call zlExcute_SaveDrugNewRecipeBill(strDrugSendFeeIDs_Out, strJosnNewDrugRcp_out, strAutoSendDrugSendFeeIDs, bln门诊, blnSendMateria)
    End If
    '4.存在卫生材料，需要生成处方
    If strJosnNewStuffRcp_out <> "" Then
        Call zlExcute_SaveStuffNewRecipeBill(strStuffSendFeeIDs_Out, strJosnNewStuffRcp_out, strAutoSendStuffSendFeeIDs, bln门诊)
    End If
        
    '5.医保记帐完成后上传
    If strInsure <> "" Then
        If InsureTranChargeDetail(strNo, strInsure, strInfo, True) = False Then
            If strInfo <> "" Then
                MsgBox strInfo, vbInformation, gstrSysName
            Else
                MsgBox "单据""" & strNo & """的数据向医保传送失败,该单据已保存！", vbInformation, gstrSysName
            End If
        End If
    End If
    zlExcute_SaveRecipeBill = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function InsureTranChargeDetail(ByVal strNo As String, ByVal strInsures As String, ByRef strErrMsg As String, _
    ByVal bln记帐已完成 As Boolean, Optional ByVal bln执行作废上传 As Boolean, _
    Optional ByVal bln记帐完成后一并上传 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保上传费用明细
    '入参:
    '   strNo-单据号
    '   strInsures-险类，多个用逗号分隔（记帐表）
    '   strErrMsg-返回错误信息
    '   bln记帐已完成-是否记帐完后在调用
    '   bln执行作废上传-执行作废上传为True,否则为执行记帐上传
    '   bln记帐完成后一并上传-是否记帐已完成后一并上传，未在记帐已完成前(事务中)上传
    '出参:
    '返回:上传成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varInsure As Variant, intInsure As Integer, i As Integer
    Dim bln记帐上传 As Boolean, bln记帐完成后上传 As Boolean
    Dim strInfo As String
    
    strErrMsg = ""
    If strInsures = "" Then InsureTranChargeDetail = True: Exit Function
     
    varInsure = Split(strInsures, ",")
    '只要第一个成功则返回True（主要针对记帐表）
    For i = 0 To UBound(varInsure)
        intInsure = Val(varInsure(i))
        If Not bln执行作废上传 Then
            bln记帐上传 = gclsInsure.GetCapability(support记帐上传, , intInsure)
        Else
            bln记帐上传 = gclsInsure.GetCapability(support记帐作废上传, , intInsure)
        End If
        bln记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , intInsure)
        
        If bln记帐上传 _
            And ((Not bln记帐已完成 And Not bln记帐完成后上传) _
                 Or (bln记帐已完成 And (bln记帐完成后上传 Or bln记帐完成后一并上传))) Then
            strInfo = ""
            If gclsInsure.TranChargeDetail(2, strNo, 2, IIf(bln执行作废上传, 2, 1), strInfo, , intInsure) = False Then
                strErrMsg = strInfo
                If i = 0 Then Exit Function
            End If
        End If
    Next
    InsureTranChargeDetail = True
End Function

Public Function zlExcute_SaveDrugNewRecipeBill(ByVal strSendDrugFeeIDs As String, _
    ByVal strJosnNewDrugRcp As String, ByVal strAutoSendDrugSendFeeIDs As String, _
    ByVal bln门诊 As Boolean, Optional ByRef blnSendMateria As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行药品生成处方
    '入参:
    '   strSendDrugFeeIds-所涉及的发药费用IDs
    '   strJosnNewDrugRcp-生成药品处方的Json串
    '   strAutoSendDrugSendFeeIDs-自动发药涉及的费用IDs
    '   lngModule-模块号
    '出参:blnSendMateria-记账后自动发药
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-15 14:09:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String, blnTrans As Boolean
    Dim clldata  As Collection
    Dim objService As clsService
    On Error GoTo errHandle
    
    If GetService(objService) = False Then Exit Function
    Set clldata = New Collection
    If strAutoSendDrugSendFeeIDs <> "" Then
        clldata.Add Array(strAutoSendDrugSendFeeIDs, 1)
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
        If zlExecuteUpdateSyncSymbol(strSendDrugFeeIDs, 0, IIf(bln门诊, 1, 2), 1) = False Then
            blnSendMateria = False
            gcnOracle.RollbackTrans: Exit Function
        End If
        If zlUpdateExcuteStatu(clldata, IIf(bln门诊, 1, 2), , , _
            UserInfo.姓名, Format(gobjDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")) = False Then
            blnSendMateria = False
            gcnOracle.RollbackTrans: Exit Function
        End If
        If objService.zlDrugSvr_NewRecipeBill(strJosnNewDrugRcp, strErrMsg) = False Then
            blnSendMateria = False
            gcnOracle.RollbackTrans
            If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, gstrSysName
            Exit Function '未提交的,回退后直接退出
        End If
     gcnOracle.CommitTrans: blnTrans = False
    
    zlExcute_SaveDrugNewRecipeBill = True
    Exit Function
errHandle:
    blnSendMateria = False
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExcute_SaveStuffNewRecipeBill(ByVal strSendStuffFeeIDs As String, _
    ByVal strJosnNewStuffRcp As String, ByVal strAutoSendStuffSendFeeIDs As String, _
    ByVal bln门诊 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行卫生材料生成处方
    '入参:
    '   strSendDrugFeeIds-所涉及的发药费用IDs
    '   strJosnNewStuffRcp-生成卫材的处方的Json串
    '   strAutoSendStuffSendFeeIDs-自动发料涉及的费用IDs
    '   lngModule-模块号
    '出参:blnSendMateria：记账后自动发药
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-15 14:09:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String, blnTrans As Boolean
    Dim clldata  As Collection
    Dim objService As clsService
    On Error GoTo errHandle
    
    If GetService(objService) = False Then Exit Function
    
    On Error GoTo errHandle
    Set clldata = New Collection
    If strAutoSendStuffSendFeeIDs <> "" Then
        clldata.Add Array(strAutoSendStuffSendFeeIDs, 1)
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
        If zlExecuteUpdateSyncSymbol(strSendStuffFeeIDs, 0, IIf(bln门诊, 1, 2), 1) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        If zlUpdateExcuteStatu(clldata, IIf(bln门诊, 1, 2), , , _
            UserInfo.姓名, Format(gobjDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")) = False Then
                gcnOracle.RollbackTrans: Exit Function
            End If
        If objService.zlStuffSvr_NewBill(strJosnNewStuffRcp, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, gstrSysName
             Exit Function
        End If
     gcnOracle.CommitTrans: blnTrans = False
    
    zlExcute_SaveStuffNewRecipeBill = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlGetRecipe_ID(ByVal strNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取处方ID
    '入参:strNo-单据号
    '返回:返回处方ID
    '编制:刘兴洪
    '日期:2019-08-12 18:57:00
    '说明:暂用NO转数字(字母用Asci代替)+序号替代
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKey As String
    strKey = Mid(strNo, 1, 1)
    If IsNumeric(strNo) Then zlGetRecipe_ID = strNo: Exit Function
    zlGetRecipe_ID = Asc(strKey) & Mid(strNo, 2)
End Function


Private Function zlGetRecipe_BillType(ByVal lng开单科室ID As Long, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str年龄 As String, ByVal str药品IDs As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取处方类型
    '入参:
    '   str药品IDs 多个英文逗号分隔
    '出参:
    '返回:处方类型:0-普通,1-儿科,2-急诊,3-精二,4-精一,5-麻醉
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If str药品IDs = "" Then Exit Function
    'Zl_Recipebill_Get
    '  开单科室id_In Number,
    '  病人id_In     Number,
    '  主页id_In     Number,
    '  病人年龄_In   Varchar2,
    '  药品ids_In Varchar2
    strSql = "Select Zl_Recipebill_Get([1],[2],[3],[4],[5]) As 处方类型 From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取处方", lng开单科室ID, lng病人ID, lng主页ID, str年龄, str药品IDs)
    If rsTemp.EOF Then Exit Function
    
    zlGetRecipe_BillType = Val((rsTemp!处方类型))
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetNewBill_CheckData(ByVal cllRctBillData As Collection, _
    ByRef strJsonFeeCheck_out As String, ByRef strJsonDrugCheck_Out As String, _
    ByRef strJsonStuffCheck_out As String, Optional ByVal intBillType As Integer, Optional bln门诊 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收集记帐单保存数据相关信息
    '入参:
    '  cllRcpBillData(结构)
    '   |-cllPati(Collect):病人信息，成员：
    '     (病人ID,主页ID,姓名,性别,年龄,费别,床号,病人科室ID,病区ID,险类,
    '      险类名称,病人来源,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '      患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '      身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):单据信息集=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):单据信息，成员：
    '       (单据号,是否多病人单,医疗小组id,医技补临床费用,简单记帐,记帐单id,记录性质,是否划价,是否急诊,加班标志,开单科室ID,
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,是否附费(N),[cllBillDetails(collect)])=cllBillists(_单据号)
    '       |-cllBillDetails(Collect):单据明细集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,从属父号,药名ID,收费细目ID,价格父号,收入项目ID,费别,婴儿序号,收费类别,计算单位,是否保险项目,保险大类ID,
    '            保险编码,收据费目,付数,数次,单价,应收金额,实收金额,统筹金额,附加标志,执行部门ID,是否自动发放,费用摘要(N),
    '            医嘱ID(N),就诊ID(N),费用类型,中药形态,煎法,执行性质,是否备货卫材(N),备货材料批次(N),是否跟踪在用,发药窗口(N),
    '            组内序号(N),诊疗项目ID(N),给药途径ID(N),给药途径名称(N),给药途径分类(N),给药频次ID(N)，给药频次名称（N),
    '            医嘱紧急标志(N),医嘱期效(N),计价特性(N),频次(N),单量（N),用法(N),皮试结果(N),超量说明(N),使用嘱托(N),
    '            发药方式(N),药品含量(N),门诊执行天数(N),煎法(N)
    '            【记帐表增加】:(病人ID,主页ID,姓名,性别,年龄,床号,病人科室ID,病区ID,险类,
    '              险类名称,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '              病人来源,患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '              身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N),
    '              医疗小组id)=cllBillDetails(_序号)
    '    以上元素中，带(N)的，表示可选节点
    '   intBillType-单据类型(1-收费单;2-记帐单;3-记帐表)
    '出参:cllPro_Out-保存单据的SQL
    '     stJsonrFeeCheck_out-费用事务的Json串
    '     strJsonDrugCheck_out-药品服务检查Json串
    '     strJsonStuffCheck_out-卫材服务检查
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-13 23:05:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBillLists As Collection, cllBillList As Collection, cllBillDetails As Collection, cllBillDetail As Collection
    Dim i As Long, j As Long, m As Long, blnFind As Boolean
    Dim cllFeeChecks As Collection, cllRowData As Collection, cllPati As Collection
    Dim rsTotal As ADODB.Recordset
    Dim lng批次 As Long
    Dim strPatiPages As String
    
    On Error GoTo errHandle
   
    '初始化数据结构
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "库房ID", adDouble, , adFldIsNullable
    rsTotal.Fields.Append "收费细目ID", adDouble, , adFldIsNullable
    
    rsTotal.Fields.Append "数量", adDouble, , adFldIsNullable
    rsTotal.Fields.Append "单价", adDouble, , adFldIsNullable
    rsTotal.Fields.Append "是否药品", adInteger, , adFldIsNullable
    rsTotal.Fields.Append "是否备货卫材", adInteger, , adFldIsNullable
    rsTotal.Fields.Append "批次", adDouble, , adFldIsNullable
    
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open

    On Error GoTo errHandle
    Set cllFeeChecks = New Collection
    
    Set cllPati = cllRctBillData("_patiinfor")
    Set cllBillLists = cllRctBillData("_cllBillLists")
    If cllBillLists.Count = 0 Then Exit Function
    
    '1.构建病人信息集
    If intBillType = 2 And bln门诊 = False Then '住院记帐单
        'Zl_住院记帐记录_Insert_Check
        '  --功能：记帐插入时，先进行相关数据的合法性检查
        '  --入参：Json_In:格式
        '  --input
        '  --    item_list[]            列表数组
        '  --        pati_id             N   病人ID
        '  --        pati_pageid         N   主页Id
        '  --        pati_deptid         N   病人科室 id
        '  --        pati_wardarea_id    n   病人病区iD
        '  --        plcdept_id          N   开单科室ID
        '  --        takedept_id         N   领用部门ID
        '  --        pati_name           C   病人姓名
        '  --        fee_audit_status    N   费用审核标志
        '  --        si_inp_status       N   住院状态
        Set cllRowData = New Collection
        cllRowData.Add Array(cllPati("病人ID"), "pati_id", "N"), "_pati_id"
        cllRowData.Add Array(cllPati("主页ID"), "pati_pageid", "N"), "_pati_pageid"
        cllRowData.Add Array(cllPati("病人科室ID"), "pati_deptid", "N"), "_pati_deptid"
        cllRowData.Add Array(cllPati("病区ID"), "pati_wardarea_id", "N"), "_pati_wardarea_id"
        Set cllBillList = cllBillLists(1)
        cllRowData.Add Array(cllBillList("开单科室ID"), "plcdept_id", "N"), "_plcdept_id"
        cllRowData.Add Array(cllBillList("领药部门ID"), "takedept_id", "N"), "_takedept_id"
        cllRowData.Add Array(cllPati("姓名"), "pati_name", "C"), "_pati_name"
        cllRowData.Add Array(cllPati("费用审核标志"), "fee_audit_status", "N"), "_fee_audit_status"
        cllRowData.Add Array(cllPati("住院状态"), "si_inp_status", "N"), "_si_inp_status"
        cllFeeChecks.Add cllRowData
    End If
     
    strPatiPages = ""
    For i = 1 To cllBillLists.Count
        Set cllBillList = cllBillLists(i)
        Set cllBillDetails = cllBillList("_cllBillDetails")
        
        For j = 1 To cllBillDetails.Count
            Set cllBillDetail = cllBillDetails(j)
            
            If intBillType = 3 Then '住院记帐表
                'Zl_住院记帐记录_Insert_Check
                '  --功能：记帐插入时，先进行相关数据的合法性检查
                '  --入参：Json_In:格式
                '  --input
                '  --    item_list[]            列表数组
                '  --        pati_id             N   病人ID
                '  --        pati_pageid         N   主页Id
                '  --        pati_deptid         N   病人科室 id
                '  --        pati_wardarea_id    n   病人病区iD
                '  --        plcdept_id          N   开单科室ID
                '  --        takedept_id         N   领用部门ID
                '  --        pati_name           C   病人姓名
                '  --        fee_audit_status    N   费用审核标志
                '  --        si_inp_status       N   住院状态
                blnFind = False
                For m = 1 To cllFeeChecks.Count
                     If cllFeeChecks(m)("_pati_id")(0) = cllBillDetail("病人ID") Then
                        blnFind = True: Exit For
                     End If
                Next
                If Not blnFind Then
                    Set cllRowData = New Collection
                    cllRowData.Add Array(cllBillDetail("病人ID"), "pati_id", "N"), "_pati_id"
                    cllRowData.Add Array(cllBillDetail("主页ID"), "pati_pageid", "N"), "_pati_pageid"
                    cllRowData.Add Array(cllBillDetail("病人科室ID"), "pati_deptid", "N"), "_pati_deptid"
                    cllRowData.Add Array(cllBillDetail("病区ID"), "pati_wardarea_id", "N"), "_pati_wardarea_id"
                    cllRowData.Add Array(cllBillList("开单科室ID"), "plcdept_id", "N"), "_plcdept_id"
                    cllRowData.Add Array(cllBillList("领药部门ID"), "takedept_id", "N"), "_takedept_id"
                    cllRowData.Add Array(cllBillDetail("姓名"), "pati_name", "C"), "_pati_name"
                    cllRowData.Add Array(cllBillDetail("费用审核标志"), "fee_audit_status", "N"), "_fee_audit_status"
                    cllRowData.Add Array(cllBillDetail("住院状态"), "si_inp_status", "N"), "_si_inp_status"
                    
                    cllFeeChecks.Add cllRowData
                    
                    strPatiPages = strPatiPages & "|" & cllBillDetail("病人ID") & "," & cllBillDetail("主页ID")
                End If
            End If
             
            If InStr(",5,6,7,", cllBillDetail("收费类别") & ",") > 0 Then
                rsTotal.Filter = "库房ID=" & cllBillDetail("执行部门ID") & " and 收费细目ID=" & cllBillDetail("收费细目ID") & " And 是否药品 = 1"
                If rsTotal.EOF Then rsTotal.AddNew
                
                rsTotal!库房ID = cllBillDetail("执行部门ID")
                rsTotal!收费细目ID = cllBillDetail("收费细目ID")
                rsTotal!数量 = Val(Nvl(rsTotal!数量)) + IIf(cllBillDetail("付数") = 0, 1, cllBillDetail("付数")) * cllBillDetail("数次")
                rsTotal!单价 = Val(cllBillDetail("单价"))
                rsTotal!是否药品 = 1
                rsTotal.Update
            ElseIf cllBillDetail("收费类别") = 4 And cllBillDetail("是否跟踪在用") = 1 Then
                lng批次 = zlGetNodeValueFromCollect(cllBillDetail, "备货材料批次", "N")
                If lng批次 < 0 Then lng批次 = 0
                
                rsTotal.Filter = "库房ID=" & cllBillDetail("执行部门ID") & " and 收费细目ID=" & cllBillDetail("收费细目ID") & " And 批次=" & lng批次 & " And 是否药品 = 0"
                If rsTotal.EOF Then rsTotal.AddNew
                
                rsTotal!库房ID = cllBillDetail("执行部门ID")
                rsTotal!收费细目ID = cllBillDetail("收费细目ID")
                rsTotal!数量 = Val(Nvl(rsTotal!数量)) + IIf(cllBillDetail("付数") = 0, 1, cllBillDetail("付数")) * cllBillDetail("数次")
                rsTotal!单价 = Val(cllBillDetail("单价"))
                rsTotal!是否药品 = 0
                rsTotal!是否备货卫材 = zlGetNodeValueFromCollect(cllBillDetail, "是否备货卫材", "N")
                rsTotal!批次 = lng批次
                rsTotal.Update
            End If
        Next
    Next
    
    rsTotal.Filter = "是否药品=1"
    strJsonDrugCheck_Out = DrugNewRecipeBill_Check_GetJsonString(rsTotal)
    rsTotal.Filter = "是否药品=0"
    strJsonStuffCheck_out = StuffNewBill_Check_GetJsonString(rsTotal)
    strJsonFeeCheck_out = NewRecipeBill_Check_GetJsonString(cllFeeChecks)
    zlGetNewBill_CheckData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetSaveSQL_InRecipeBill(ByVal intBillType As Integer, ByVal cllPati As Collection, ByVal cllBillList As Collection, ByVal cllBillDetail As Collection, _
    ByRef lng费用ID_out As Long, ByRef strSQL_Out As String, Optional ByVal blnMarkSyncSymbol As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保存住院记帐单对应的SQL
    '入参:cllPati-病人信息集
    '     cllBillList-单据信息集
    '     cllBillDetail-单据明细集
    '     blnMarkSyncSymbol-是否更新同步标志
    '出参:
    '     lng费用ID_out-费用Id
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-15 11:29:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    On Error GoTo errHandle

    lng费用ID_out = gobjDatabase.GetNextId("病人费用记录")
    
    'zl_住院记帐记录_INSERT_S
    strSQL_Out = "zl_住院记帐记录_INSERT_S("
    '    No_In             住院费用记录.No%Type,
    strSQL_Out = strSQL_Out & "'" & cllBillList("单据号") & "',"
    '    序号_In           住院费用记录.序号%Type,
    strSQL_Out = strSQL_Out & "" & cllBillDetail("序号") & ","
    If intBillType = 3 Then
        '    病人id_In         住院费用记录.病人id%Type,
        strSQL_Out = strSQL_Out & "" & cllBillDetail("病人ID") & ","
        '    主页id_In         住院费用记录.主页id%Type,
        strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("主页ID")) & ","
        '    标识号_In         住院费用记录.标识号%Type,
        strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("标识号")) & ","
        '    姓名_In           住院费用记录.姓名%Type,
        strSQL_Out = strSQL_Out & "'" & cllBillDetail("姓名") & "',"
        '    性别_In           住院费用记录.性别%Type,
        strSQL_Out = strSQL_Out & "'" & cllBillDetail("性别") & "',"
        '    年龄_In           住院费用记录.年龄%Type,
        strSQL_Out = strSQL_Out & "'" & cllBillDetail("年龄") & "',"
        '    床号_In           住院费用记录.床号%Type,
        strSQL_Out = strSQL_Out & "'" & cllBillDetail("床号") & "',"
        '    费别_In           住院费用记录.费别%Type,
        strSQL_Out = strSQL_Out & "'" & cllBillDetail("费别") & "',"
        '    病区id_In         住院费用记录.病人病区id%Type,
        strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("病区ID")) & ","
        '    科室id_In         住院费用记录.病人科室id%Type,
        strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("病人科室ID")) & ","
    Else
        '    病人id_In         住院费用记录.病人id%Type,
        strSQL_Out = strSQL_Out & "" & cllPati("病人ID") & ","
        '    主页id_In         住院费用记录.主页id%Type,
        strSQL_Out = strSQL_Out & "" & ZVal(cllPati("主页ID")) & ","
        '    标识号_In         住院费用记录.标识号%Type,
        strSQL_Out = strSQL_Out & "" & ZVal(cllPati("标识号")) & ","
        '    姓名_In           住院费用记录.姓名%Type,
        strSQL_Out = strSQL_Out & "'" & cllPati("姓名") & "',"
        '    性别_In           住院费用记录.性别%Type,
        strSQL_Out = strSQL_Out & "'" & cllPati("性别") & "',"
        '    年龄_In           住院费用记录.年龄%Type,
        strSQL_Out = strSQL_Out & "'" & cllPati("年龄") & "',"
        '    床号_In           住院费用记录.床号%Type,
        strSQL_Out = strSQL_Out & "'" & cllPati("床号") & "',"
        '    费别_In           住院费用记录.费别%Type,
        strSQL_Out = strSQL_Out & "'" & cllBillDetail("费别") & "',"
        '    病区id_In         住院费用记录.病人病区id%Type,
        strSQL_Out = strSQL_Out & "" & ZVal(cllPati("病区ID")) & ","
        '    科室id_In         住院费用记录.病人科室id%Type,
        strSQL_Out = strSQL_Out & "" & ZVal(cllPati("病人科室ID")) & ","
    
    End If
    '    加班标志_In       住院费用记录.加班标志%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillList("加班标志")) & ","
    '    婴儿费_In         住院费用记录.婴儿费%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("婴儿序号")) & ","
    '    开单部门id_In     住院费用记录.开单部门id%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillList("开单科室ID")) & ","
    '    开单人_In         住院费用记录.开单人%Type,
    strSQL_Out = strSQL_Out & "'" & cllBillList("开单人") & "',"
    '    从属父号_In       住院费用记录.从属父号%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("从属父号")) & ","
    '    收费细目id_In     住院费用记录.收费细目id%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("收费细目ID")) & ","
    '    收费类别_In       住院费用记录.收费类别%Type,
    strSQL_Out = strSQL_Out & "'" & cllBillDetail("收费类别") & "',"
    '    计算单位_In       住院费用记录.计算单位%Type,
    strSQL_Out = strSQL_Out & "'" & cllBillDetail("计算单位") & "',"
    '    保险项目否_In     住院费用记录.保险项目否%Type,
    strSQL_Out = strSQL_Out & "" & cllBillDetail("是否保险项目") & ","
    '    保险大类id_In     住院费用记录.保险大类id%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("保险大类ID")) & ","
    '    保险编码_In       住院费用记录.保险编码%Type,
    strSQL_Out = strSQL_Out & "'" & cllBillDetail("保险编码") & "',"
    '    付数_In           住院费用记录.付数%Type,
    strSQL_Out = strSQL_Out & "" & IIf(cllBillDetail("付数") = 0, 1, cllBillDetail("付数")) & ","
    '    数次_In           住院费用记录.数次%Type,
    strSQL_Out = strSQL_Out & "" & cllBillDetail("数次") & ","
    '    附加标志_In       住院费用记录.附加标志%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("附加标志")) & ","
    '    执行部门id_In     住院费用记录.执行部门id%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("执行部门ID")) & ","
    '    价格父号_In       住院费用记录.价格父号%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("价格父号")) & ","
    '    收入项目id_In     住院费用记录.收入项目id%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("收入项目ID")) & ","
    '    收据费目_In       住院费用记录.收据费目%Type,
    strSQL_Out = strSQL_Out & "'" & cllBillDetail("收据费目") & "',"
    '    标准单价_In       住院费用记录.标准单价%Type,
    strSQL_Out = strSQL_Out & "" & cllBillDetail("单价") & ","
    '    应收金额_In       住院费用记录.应收金额%Type,
    strSQL_Out = strSQL_Out & "" & cllBillDetail("应收金额") & ","
    '    实收金额_In       住院费用记录.实收金额%Type,
    strSQL_Out = strSQL_Out & "" & cllBillDetail("实收金额") & ","
    '    统筹金额_In       住院费用记录.统筹金额%Type,
    strSQL_Out = strSQL_Out & "" & cllBillDetail("统筹金额") & ","
    '    发生时间_In       住院费用记录.发生时间%Type,
    strSQL_Out = strSQL_Out & "to_date('" & cllBillList("发生时间") & "','yyyy-mm-dd hh24:mi:ss'),"
    '    登记时间_In       住院费用记录.登记时间%Type,
    strSQL_Out = strSQL_Out & "to_date('" & cllBillList("登记时间") & "','yyyy-mm-dd hh24:mi:ss'),"
    '    划价_In           Number,
    strSQL_Out = strSQL_Out & "" & cllBillList("是否划价") & ","
    '    操作员编号_In     住院费用记录.操作员编号%Type,
    strSQL_Out = strSQL_Out & "'" & cllBillList("操作员编号") & "',"
    '    操作员姓名_In     住院费用记录.操作员姓名%Type,
    strSQL_Out = strSQL_Out & "'" & cllBillList("操作员姓名") & "',"
    '    费用id_In         In Out 住院费用记录.Id%Type,
    strSQL_Out = strSQL_Out & "" & ZVal(lng费用ID_out) & ","
    '    多病人单_In       Number := 0,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillList("是否多病人单")) & ","
    '    记帐单id_In       住院费用记录.记帐单id%Type := Null,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillList("记帐单id")) & ","
    '    费用摘要_In       住院费用记录.摘要%Type := Null,
    strSQL_Out = strSQL_Out & "'" & cllBillDetail("费用摘要") & "',"
    '    是否急诊_In       住院费用记录.是否急诊%Type := 0,
    strSQL_Out = strSQL_Out & "" & cllBillList("是否急诊") & ","
    '    医嘱序号_In       住院费用记录.医嘱序号%Type := Null,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("医嘱ID")) & ","
    '    简单记帐_In       Number := 0,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillList("简单记帐")) & ","
    '    费用类型_In       住院费用记录.费用类型%Type := Null,
    strSQL_Out = strSQL_Out & "'" & cllBillDetail("费用类型") & "',"
    '    医技补临床费用_In Number := 0,
    strSQL_Out = strSQL_Out & "" & ZVal(cllBillList("医技补临床费用")) & ","
    '    中药形态_In       住院费用记录.结论%Type := Null,
    strSQL_Out = strSQL_Out & "'" & cllBillDetail("中药形态") & "',"
    '    医疗小组id_In     住院费用记录.医疗小组id%Type := -1,
    If intBillType = 3 Then
        strSQL_Out = strSQL_Out & "" & ZVal(cllBillDetail("医疗小组id")) & ","
    Else
        strSQL_Out = strSQL_Out & "" & ZVal(cllBillList("医疗小组id")) & ","
    End If
    '    煎法_In           住院费用记录.结论%Type := Null,
    strSQL_Out = strSQL_Out & "'" & cllBillDetail("煎法") & "',"
    '    执行性质_In       住院费用记录.结论%Type := 0
    strSQL_Out = strSQL_Out & "'" & cllBillDetail("执行性质") & "',"
    '    批次_In         住院费用记录.批次%Type := Null,
    strSQL_Out = strSQL_Out & "" & ZVal(zlGetNodeValueFromCollect(cllBillDetail, "备货材料批次", "N")) & ","
    '   领药部门ID_In         住院费用记录.领药部门ID%Type := Null,
    If InStr(",5,6,7,", "," & cllBillDetail("收费类别") & ",") > 0 Then
        strSQL_Out = strSQL_Out & ZVal(zlGetNodeValueFromCollect(cllBillList, "领药部门ID", "N")) & ","
    Else
        strSQL_Out = strSQL_Out & "Null" & ","
    End If
    '  更新同步标志_In Number := 1
    strSQL_Out = strSQL_Out & "" & IIf(blnMarkSyncSymbol, 1, 0) & ","
    '  期效_In         门诊费用记录.医嘱期效%Type := 0,
    strSQL_Out = strSQL_Out & "NULL,"
    '  是否保密_In     门诊费用记录.是否保密%Type := 0,
    strSQL_Out = strSQL_Out & "NULL,"
    '  是否附费_In     门诊费用记录.是否附费%Type := 0
    strSQL_Out = strSQL_Out & "" & ZVal(zlGetNodeValueFromCollect(cllBillList, "是否附费", "N")) & ")"
    
    zlGetSaveSQL_InRecipeBill = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlGetNewBill_SaveData(ByVal cllRctBillData As Collection, ByRef cllPro_out As Collection, _
    ByRef strJsonNewDrugRcp_out As String, ByRef strJsonNewStuffRcp_out As String, _
    ByRef strDrugSendFeeIDs_Out As String, ByRef strStuffSendFessIDs_Out As String, _
    Optional ByVal intBillType As Integer, _
    Optional ByVal bln划价 As Boolean, Optional ByVal cllDrawDeptIds As Collection, Optional bln门诊 As Boolean, _
    Optional ByRef strAutoSendDrugSendFeeIDs_Out As String, Optional ByRef strAutoSendStuffSendFeeIDs_Out As String, _
    Optional ByVal blnMarkSyncSymbol As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收集记帐或收费单保存数据相关信息
    '入参:
    '  cllRcpBillData(结构)
    '   |-cllPati(Collect):病人信息，成员：
    '     (病人ID,主页ID,姓名,性别,年龄,费别,床号,病人科室ID,病区ID,险类,
    '      险类名称,病人来源,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '      患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '      身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):单据信息集=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):单据信息，成员：
    '       (单据号,是否多病人单,医疗小组id,医技补临床费用,简单记帐,记帐单id,记录性质,是否划价,是否急诊,加班标志,开单科室ID,
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,是否附费(N),[cllBillDetails(collect)])=cllBillists(_单据号)
    '       |-cllBillDetails(Collect):单据明细集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,从属父号,药名ID,收费细目ID,价格父号,收入项目ID,费别,婴儿序号,收费类别,计算单位,是否保险项目,保险大类ID,
    '            保险编码,收据费目,付数,数次,单价,应收金额,实收金额,统筹金额,附加标志,执行部门ID,是否自动发放,费用摘要(N),
    '            医嘱ID(N),就诊ID(N),费用类型,中药形态,煎法,执行性质,是否备货卫材(N),备货材料批次(N),是否跟踪在用,发药窗口(N),
    '            组内序号(N),诊疗项目ID(N),给药途径ID(N),给药途径名称(N),给药途径分类(N),给药频次ID(N)，给药频次名称（N),
    '            医嘱紧急标志(N),医嘱期效(N),计价特性(N),频次(N),单量（N),用法(N),皮试结果(N),超量说明(N),使用嘱托(N),
    '            发药方式(N),药品含量(N),门诊执行天数(N),煎法(N)
    '            【记帐表增加】:(病人ID,主页ID,姓名,性别,年龄,床号,病人科室ID,病区ID,险类,
    '              险类名称,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '              病人来源,患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '              身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N),
    '              医疗小组id)=cllBillDetails(_序号)
    '    以上元素中，带(N)的，表示可选节点
    '   intBillType-单据类型(1-收费单;2-记帐单;3-记帐表)
    '   cllDrawDeptIds-领药部门Ids((病人ID,领药部门ID)(_病人ID)
    '   blnMarkSyncSymbol-是否标记同步状态
    '出参:cllPro_Out-保存单据的SQL
    '     strJosnNewDrugRcp_out-存在药品时，返回生成药品处方的Json串
    '     strJosnNewStuffRcp_out-存在卫材时，返回生成卫材的Json串
    '     strDrugSendFeeIDs_Out-药品需要发药所涉及的费用ID(修改同步标志用)
    '     strStuffSendFessIDs_Out-卫材需要发料所涉及的费用ID(修改同步标志用)
    '     strAutoSendDrugSendFeeIDs_Out-自动发药所涉及的费用ID(修改费用执行状态用)
    '     strAutoSendStuffSendFeeIDs_Out-自动发料所涉及的费用ID(修改费用执行状态用)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-13 23:05:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, q As Long, lng费用ID As Long
    Dim lng病人ID As Long
    Dim cllPati As Collection
    Dim strSql As String, lng领药部门ID As Long
    Dim str药品IDs As String, byt处方类型 As Byte, k As Integer
    Dim byt取药特性 As Byte, byt计价特性 As Byte
     
    Dim cllDrugBillLists As Collection, cllStuffBillLists As Collection '药品单据集及卫材单据集
    'Billists(collect)(以下每个元素都是Array(元素值,节点名,类型),_节点名
     '   |-BillList
     '       |-单据头信息:recipe_id,rcp_no等
     '       |-BillDetails
     '           |-BillDetail(collect) 单据明细信息集
     '               |-明细数据:rcpdtl_id,serial_num等
    Dim cllDrugBillList As Collection, cllStuffBillList As Collection
    Dim cllDrugDetails As Collection, cllStuffDetails As Collection
    Dim cllDrugDetail As Collection, cllStuffDetail As Collection
    Dim cllBillLists As Collection, cllBillList As Collection, cllBillDetails As Collection, cllBillDetail As Collection
    Dim cllSvrPatiInfor As Collection
   
    On Error GoTo errHandle
    strDrugSendFeeIDs_Out = "": strStuffSendFessIDs_Out = ""
    strJsonNewStuffRcp_out = "": strDrugSendFeeIDs_Out = "": strJsonNewDrugRcp_out = ""
    strAutoSendDrugSendFeeIDs_Out = "": strAutoSendStuffSendFeeIDs_Out = ""
    
    Set cllPati = cllRctBillData("_patiinfor")
    '1.构建病人信息集
    Set cllSvrPatiInfor = New Collection
    cllSvrPatiInfor.Add Array(intBillType, "billtype", "N"), "_billtype"
    If intBillType <> 3 Then '非记帐表
        cllSvrPatiInfor.Add Array(cllPati("病人来源"), "pati_source", "N"), "_pati_source"
        cllSvrPatiInfor.Add Array(cllPati("病人ID"), "pati_id", "N"), "_pati_id"
        cllSvrPatiInfor.Add Array(cllPati("主页ID"), "pati_pageid", "N"), "pati_pageid"
        cllSvrPatiInfor.Add Array(cllPati("姓名"), "pati_name", "C"), "_pati_name"
        cllSvrPatiInfor.Add Array(cllPati("性别"), "pati_sex", "C"), "_pati_sex"
        cllSvrPatiInfor.Add Array(zlGetPatiSex_Code(cllPati("性别")), "pati_sex_code", "C"), "_pati_sex_code"
        cllSvrPatiInfor.Add Array(cllPati("年龄"), "pati_age", "C"), "_pati_age"
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "身份", "C"), "pati_identity", "C"), "_pati_identity"
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "出生日期", "C"), "pati_birthdate", "C"), "_pati_birthdate"
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "身份证号", "C"), "pati_idcard", "C"), "_pati_idcard"
        cllSvrPatiInfor.Add Array(cllPati("病人科室ID"), "pati_deptid", "N"), "_pati_deptid"
        cllSvrPatiInfor.Add Array(cllPati("病区ID"), "pati_wardarea_id", "N"), "_pati_wardarea_id"
        cllSvrPatiInfor.Add Array(cllPati("险类"), "si_type_id", "N"), "_si_type_id"
        lng病人ID = cllPati("病人ID")
        If Val(cllPati("险类")) <> 0 Then
            If cllPati("险类名称") = "" Then
                cllSvrPatiInfor.Add Array(GetInsureName(Val(cllPati("险类"))), "si_type_name", "C"), "_si_type_name"
            Else
                cllSvrPatiInfor.Add Array(cllPati("险类名称"), "si_type_name", "C"), "_si_type_name"
            End If
        Else
            cllSvrPatiInfor.Add Array("", "si_type_name", "C"), "_si_type_name"
        End If
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "挂号单ID", "N"), "rgst_id", "N"), "_rgst_id"
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "代办人姓名", "C"), "recipe_proxy_name", "C"), "_recipe_proxy_name"
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "代办人身份证号", "C"), "recipe_proxy_idno", "C"), "_recipe_proxy_idno"
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "患者体重", "C"), "recipe_pat_bodywt", "C"), "_recipe_pat_bodywt"
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "患者体重单位", "C"), "recipe_pat_bodywt_unit", "C"), "_recipe_pat_bodywt_unit"
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "诊断记录ID", "C"), "outp_diag_rec_id", "C"), "_outp_diag_rec_id"
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "诊断ID", "C"), "diag_id", "C"), "_diag_id"
        cllSvrPatiInfor.Add Array(zlGetNodeValueFromCollect(cllPati, "诊断名称", "C"), "diag_name", "C"), "_diag_name"
    End If
    
    Set cllBillLists = cllRctBillData("_cllBillLists")
    Set cllDrugBillLists = New Collection: Set cllStuffBillLists = New Collection
    If cllPro_out Is Nothing Then Set cllPro_out = New Collection
    For i = 1 To cllBillLists.Count
        Set cllBillList = cllBillLists(i)
        Set cllBillDetails = cllBillList("_cllBillDetails")
        Set cllDrugDetails = New Collection
        Set cllStuffDetails = New Collection
                    
        For j = 1 To cllBillDetails.Count
            Set cllBillDetail = cllBillDetails(j)
            
            Select Case intBillType
            Case 1 '收费
                If zlGetSaveSQL_OutRecipeBill(cllPati, cllBillList, cllBillDetail, lng费用ID, strSql, blnMarkSyncSymbol) = False Then Exit Function
                zlAddArray cllPro_out, strSql
            Case 2, 3 '记帐
                If Not bln门诊 Then
                    If zlGetSaveSQL_InRecipeBill(intBillType, cllPati, cllBillList, cllBillDetail, lng费用ID, strSql, blnMarkSyncSymbol) = False Then Exit Function
                    zlAddArray cllPro_out, strSql
                Else    '门诊记帐
                    If zlGetSaveSQL_OutRecipeBill(cllPati, cllBillList, cllBillDetail, lng费用ID, strSql, blnMarkSyncSymbol) = False Then Exit Function
                    zlAddArray cllPro_out, strSql
                End If
            Case Else
                Exit Function
            End Select
            
            If InStr(",5,6,7,", "," & cllBillDetail("收费类别") & ",") > 0 Then
                Set cllDrugDetail = New Collection
                If intBillType = 3 Then '记帐表
                    cllDrugDetail.Add Array(cllBillDetail("病人ID"), "pati_id", "N"), "_pati_id" 'N1病人ID记帐表（ 单据=10）时，有此节点
                    cllDrugDetail.Add Array(cllBillDetail("主页ID"), "pati_pageid", "N"), "_pati_pageid" 'N主页ID
                    cllDrugDetail.Add Array(cllBillDetail("姓名"), "pati_name", "C"), "_pati_name" 'C1病人姓名
                    cllDrugDetail.Add Array(cllBillDetail("性别"), "pati_sex", "C"), "_pati_sex" 'C1性别
                    cllDrugDetail.Add Array(zlGetPatiSex_Code(cllBillDetail("性别")), "pati_sex_code", "C"), "_pati_sex_code"     'C1性别编号（新门诊)
                    cllDrugDetail.Add Array(cllBillDetail("年龄"), "pati_age", "C"), "_pati_age" 'C1年龄
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "身份", "C"), "pati_identity", "C"), "_pati_identity" 'C身份
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "出生日期", "C"), "pati_birthdate", "C"), "_pati_birthdate" 'C出生日期:yyyy-mm-dd hh:mi:ss
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "身份证号", "C"), "pati_idcard", "C"), "_pati_idcard" 'C身份证号
                    cllDrugDetail.Add Array(cllBillDetail("病区ID"), "pati_wardarea_id", "N"), "_pati_wardarea_id" 'N病人病区ID
                    cllDrugDetail.Add Array(cllBillDetail("病人科室ID"), "pati_deptid", "N"), "_pati_deptid" 'N1病人科室ID
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "挂号单ID", "N"), "rgst_id", "N"), "_rgst_id"
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "代办人姓名", "C"), "recipe_proxy_name", "C"), "_recipe_proxy_name"
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "代办人身份证号", "C"), "recipe_proxy_idno", "C"), "_recipe_proxy_idno"
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "患者体重", "C"), "recipe_pat_bodywt", "C"), "_recipe_pat_bodywt"
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "患者体重单位", "C"), "recipe_pat_bodywt_unit", "C"), "_recipe_pat_bodywt_unit"
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "诊断记录ID", "C"), "outp_diag_rec_id", "C"), "_outp_diag_rec_id"
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "诊断ID", "C"), "diag_id", "C"), "_diag_id"
                    cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "诊断名称", "C"), "diag_name", "C"), "_diag_name"
                End If
                
                cllDrugDetail.Add Array(lng费用ID, "rcpdtl_id", "N"), "_rcpdtl_id" 'N1处方明细id,目前传入的是费用id
                cllDrugDetail.Add Array(cllBillDetail("序号"), "serial_num", "N"), "_serial_num" 'N1序号:(变更(包括存储)：序号和组号，1、2、3、3、3、4…)
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "组内序号", "N"), "group_sno", "N"), "_group_sno"   'N组内序号 (包括存储)：1、2、3
                cllDrugDetail.Add Array(cllBillDetail("执行部门ID"), "pharmacy_id", "N"), "_pharmacy_id" 'N1药房ID
                cllDrugDetail.Add Array(zlGetDeptName(cllBillDetail("执行部门ID")), "pharmacy_name", "C"), "_pharmacy_name" 'C1药房名称(新门诊)
                
                lng领药部门ID = 0
                If IsEmpty(cllBillList("领药部门ID")) = False Then lng领药部门ID = cllBillList("领药部门ID")
                If lng领药部门ID = 0 And Not cllDrawDeptIds Is Nothing Then
                    For q = 1 To cllDrawDeptIds.Count
                        If cllDrawDeptIds(q)("_pati_id") = lng病人ID Then
                            lng领药部门ID = zlGetNodeValueFromCollect(cllDrawDeptIds(q), "_takedept_id", "N")
                        End If
                    Next
                End If
                cllDrugDetail.Add Array(IIf(lng领药部门ID = 0, Empty, lng领药部门ID), "takedept_id", "N"), "_takedept_id" 'N1领药部门ID:针对住院才传入
                
                cllDrugDetail.Add Array(cllBillDetail("药名ID"), "cadn_id", "N"), "_cadn_id"  'N1药品通用名称id(药名ID)(新门诊)
                cllDrugDetail.Add Array(cllBillDetail("收费细目ID"), "drug_id", "N"), "_drug_id" 'N1药品ID
                cllDrugDetail.Add Array(cllBillDetail("婴儿序号"), "baby_num", "N"), "_baby_num" 'N婴儿序号
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "医嘱ID", "N"), "advice_id", "N"), "_advice_id"    'N0医嘱ID可选参数，医嘱记录产生
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "给药途径ID", "N"), "drug_method_id", "N"), "_drug_method_id"  'N1给药途径id(新门诊):诊疗项目ID
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "给药途径名称", "C"), "drug_method_name", "C"), "_drug_method_name"  'C1给药途径名称
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "给药途径分类", "C"), "drug_method_class_code", "C"), "_drug_method_class_code"  'C1给药途径分类(新门诊):执行分类编号
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "给药频次ID", "C"), "drug_freq_id", "C"), "_drug_freq_id"  'C1给药频次id(新门诊):诊疗频率编码
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "给药频次名称", "C"), "drug_freq_name", "C"), "_drug_freq_name"  'C1给药频次名称d(新门诊):
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "医嘱紧急标志", "N"), "emergency_tag", "N"), "_emergency_tag" 'N医嘱记录中的紧急标志(0-普通;1-紧急;2-补录(对门诊无效))
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "医嘱期效", "N"), "effectivetime", "N"), "_effectivetime" 'N0医嘱期效
                
                byt计价特性 = 0: byt取药特性 = 0
                If CollectionExitsValue(cllBillDetail, "执行性质") Then '0-正常,3-离院带药,4-自取药
                    byt取药特性 = Decode(Val(cllBillDetail("执行性质")), 3, 1, 4, 2, 0)
                ElseIf CollectionExitsValue(cllBillDetail, "计价特性") Then '计价特性：0-2-计价特性,3-离院带药,4-自取药   计价特性(0-正常计价；1-不计价；2-手工计价))
                    If Val(cllBillDetail("计价特性")) = 3 Or Val(cllBillDetail("计价特性")) = 4 Then
                        byt取药特性 = Decode(Val(cllBillDetail("计价特性")), 3, 1, 4, 2, 0)
                    Else
                        byt计价特性 = Val(cllBillDetail("计价特性"))
                    End If
                End If
                cllDrugDetail.Add Array(byt计价特性, "fee_mode", "N"), "_fee_mode" 'N 0 计价特性:0-正常计价；1-不计价；2-手工计价
                cllDrugDetail.Add Array(byt取药特性, "use_mode", "N"), "_use_mode" 'N 0 取药特性:0-正常方式，1-离院带药，2-自取药
                
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "频次", "N"), "frequency", "N"), "_frequency" 'N0
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "单量", "N"), "single", "N"), "_single" 'N0单量
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "用法", "C"), "usage", "C"), "_usage" 'C0用法
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "皮试结果", "N"), "rcpdtl_st_result", "N"), "_rcpdtl_st_result" 'N"皮试结果(新门诊)1-阴性，2-阳性，3-免试，4-连续用药处方生成时已确定或已有皮试结果。ZLHIS目前支持不全"
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "超量说明", "C"), "rcpdtL_excs_desc", "C"), "_rcpdtL_excs_desc" 'C超量说明(新门诊)
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "使用嘱托", "C"), "rcpdtL_drask", "C"), "_rcpdtL_drask"    'C使用嘱托(新门诊)
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "发药方式", "N"), "disps_mode_code", "N"), "_disps_mode_code" 'N1"发药方式(新门诊)1-正常发放；2-科室贮药；3-自备药；4-代购药ZLHIS目前支持不全(2,4)"
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "药品含量", "N"), "drug_content", "N"), "_drug_content" 'N药品含量（剂量系数）(新门诊)：
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "门诊执行天数", "N"), "rcpdtl_outp_drugdays", "N"), "_rcpdtl_outp_drugdays" 'N本院门诊执行天数(新门诊)：ZLHIS是给药执行次数，要转换为天数传
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "煎法", "C"), "decoction_method", "C"), "_decoction_method" 'C0煎法
                cllDrugDetail.Add Array(cllBillDetail("付数"), "packages_num", "N"), "_packages_num" 'N1发药付数
                cllDrugDetail.Add Array(cllBillDetail("数次"), "send_num", "N"), "_send_num" 'N1发药数量
                cllDrugDetail.Add Array(cllBillDetail("计算单位"), "send_unit", "C"), "_send_unit" 'C1发药单位：zlhis零售单位
                cllDrugDetail.Add Array(cllBillDetail("单价"), "price", "N"), "_price" 'N售价
                cllDrugDetail.Add Array(cllBillDetail("应收金额"), "money", "N"), "_money"  'N零售金额(新门诊)
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "发药窗口", "C"), "pharmacy_window", "C"), "_pharmacy_window" 'C0发药窗口
                cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "费用摘要", "C"), "memo", "C"), "_memo"    'C0摘要
                cllDrugDetail.Add Array(IIf(bln门诊, 1, 2), "fee_source", "N"), "_fee_source" 'N1费用来源
                cllDrugDetail.Add Array(cllBillDetail("是否自动发放"), "drug_auto_send", "N"), "_drug_auto_send" '          N   0   是否自动发放药品:0-不自动发药,1-自动发药
                cllDrugDetails.Add cllDrugDetail, "_" & cllBillDetail("序号")
                
                strDrugSendFeeIDs_Out = strDrugSendFeeIDs_Out & "," & lng费用ID
                If Val(cllBillDetail("是否自动发放")) = 1 Then
                    strAutoSendDrugSendFeeIDs_Out = strAutoSendDrugSendFeeIDs_Out & "," & lng费用ID
                End If
            ElseIf cllBillDetail("收费类别") = "4" And cllBillDetail("是否跟踪在用") = 1 Then
                Set cllStuffDetail = New Collection
                If intBillType = 3 Then '记帐表
                    cllStuffDetail.Add Array(cllBillDetail("病人ID"), "pati_id", "N"), "_pati_id" 'N1病人ID记帐表（ 单据=10）时，有此节点
                    cllStuffDetail.Add Array(cllBillDetail("主页ID"), "pati_pageid", "N"), "_pati_pageid" 'N主页ID
                    cllStuffDetail.Add Array(cllBillDetail("姓名"), "pati_name", "C"), "_pati_name" 'C1病人姓名
                    cllStuffDetail.Add Array(cllBillDetail("性别"), "pati_sex", "C"), "_pati_sex" 'C1性别
                    cllStuffDetail.Add Array(zlGetPatiSex_Code(cllBillDetail("性别")), "pati_sex_code", "C"), "_pati_sex_code"     'C1性别编号（新门诊)
                    cllStuffDetail.Add Array(cllBillDetail("年龄"), "pati_age", "C"), "_pati_age" 'C1年龄
                    cllStuffDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "身份", "C"), "pati_identity", "C"), "_pati_identity" 'C身份
                    cllStuffDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "出生日期", "C"), "pati_birthdate", "C"), "_pati_birthdate" 'C出生日期:yyyy-mm-dd hh:mi:ss
                    cllStuffDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "身份证号", "C"), "pati_idcard", "C"), "_pati_idcard" 'C身份证号
                    cllStuffDetail.Add Array(cllBillDetail("病区ID"), "pati_wardarea_id", "N"), "_pati_wardarea_id" 'N病人病区ID
                    cllStuffDetail.Add Array(cllBillDetail("病人科室ID"), "pati_deptid", "N"), "_pati_deptid" 'N1病人科室ID
                End If
    
                cllStuffDetail.Add Array(cllBillDetail("序号"), "serial_num", "N"), "_serial_num" 'N1序号
                cllStuffDetail.Add Array(cllBillDetail("执行部门ID"), "warehouse_id", "N"), "_warehouse_id"  'N1库房ID
                cllStuffDetail.Add Array(lng费用ID, "stuffdtl_id", "N"), "_stuffdtl_id" 'N1处方明细id,目前传入的是门诊(住院)费用记录.ID
                cllStuffDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "是否备货卫材", "N"), "is_bakstuff", "N"), "_is_bakstuff"   'N1是否备货卫材:有高值卫材才需要传入，非0表示是高值卫材模式(如扫码时使用)
                
                Dim lng批次 As Long
                lng批次 = Val(zlGetNodeValueFromCollect(cllBillDetail, "备货材料批次", "N"))
                cllStuffDetail.Add Array(IIf(lng批次 > 0, lng批次, Empty), "bakstuff_batch", "N"), "_bakstuff_batch" '备货材料批次
                
                cllStuffDetail.Add Array(cllBillDetail("收费细目ID"), "stuff_id", "N"), "_stuff_id"  'N1卫材ID
                cllStuffDetail.Add Array(cllBillDetail("婴儿序号"), "baby_num", "N"), "_baby_num" 'N婴儿序号
                cllStuffDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "医嘱ID", "N"), "advice_id", "N"), "_advice_id"  'N0医嘱ID可选参数，医嘱记录产生
                cllStuffDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "医嘱紧急标志", "N"), "emergency_tag", "N"), "_emergency_tag"  'N医嘱记录中的紧急标志(0-普通;1-紧急;2-补录(对门诊无效))
                cllStuffDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "医嘱期效", "N"), "effectivetime", "N"), "_effectivetime" 'C0医嘱期效
                cllStuffDetail.Add Array(cllBillDetail("付数"), "packages_num", "N"), "_packages_num"   'N1付数
                cllStuffDetail.Add Array(cllBillDetail("数次"), "outbound_num", "N"), "_outbound_num"   'N1出库数量
                cllStuffDetail.Add Array(cllBillDetail("单价"), "price", "N"), "_price"  'C标准单价(售价)
                cllStuffDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "发药窗口", "C"), "warehouse_window", "C"), "_warehouse_window"  'C0发料窗口
                cllStuffDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "费用摘要", "C"), "memo", "C"), "_memo" 'C0摘要
                cllStuffDetail.Add Array(cllBillDetail("是否自动发放"), "stuff_auto_send", "N"), "_stuff_auto_send"  '          N   0   是否自动发料:0-不自动发料,1-自动发料
                cllStuffDetail.Add Array(IIf(bln门诊, 1, 2), "fee_source", "N"), "_fee_source" 'N0费用来源 '1-门诊;2-住院;4-体检
                cllStuffDetails.Add cllStuffDetail, "_" & cllBillDetail("序号")
                
                strStuffSendFessIDs_Out = strStuffSendFessIDs_Out & "," & lng费用ID
                If Val(cllBillDetail("是否自动发放")) = 1 Then
                    strAutoSendStuffSendFeeIDs_Out = strAutoSendStuffSendFeeIDs_Out & "," & lng费用ID
                End If
            End If
        Next
    
        '单据
        If cllStuffDetails.Count <> 0 Then '存在卫生材料
            Set cllStuffBillList = New Collection
            cllStuffBillList.Add Array(zlGetRecipe_ID(cllBillList("单据号")), "bill_id", "N"), "_bill_id"  '   N   1   处方id(新门诊):ZLHIS无，暂用NO转数字(字母用Asci代替)+序号替代
            cllStuffBillList.Add Array(cllBillList("单据号"), "stuff_no", "C"), "_stuff_no" '   C   1   NO
            cllStuffBillList.Add Array(IIf(cllBillList("是否划价") = 1, 0, 1), "charge_tag", "N"), "_charge_tag" '   N   1   收费标志:0-未收费或记帐划价;1-已收费或记帐
            cllStuffBillList.Add Array(cllBillList("划价人"), "fee_acnter", "C"), "_fee_acnter" '   C       划价人
            cllStuffBillList.Add Array(cllBillList("开单科室ID"), "plcdept_id", "C"), "plcdept_id"  '   C       开单科室id（新门诊):应该要
            cllStuffBillList.Add Array(zlGetDeptName(cllBillList("开单科室ID")), "plcdept", "C"), "plcdept" '   C       开单科室名称（新门诊):应该要
            cllStuffBillList.Add Array(zlGetPersonID(cllBillList("开单人")), "placer_id", "C"), "placer_id" '   C       开单医师id（新门诊):应该要
            cllStuffBillList.Add Array(cllBillList("开单人"), "placer", "C"), "placer" '   C       开单医师（新门诊)   增加:应该要
            If intBillType <> 3 Then
                cllStuffBillList.Add Array(cllPati("医疗付款方式编码"), "apply_fee_category_code", "C"), "_apply_fee_category_code" '   C       申请单费别编码(医疗付款方式编码)(新门诊)    增加:应该要；
                cllStuffBillList.Add Array(cllPati("医疗付款方式名称"), "apply_fee_category_name", "C"), "_apply_fee_category_name" '   C       申请单费别名称（医疗付款方式名称）(新门诊)  增加:应该要；
            Else
                '记帐表暂不传
            End If
            cllStuffBillList.Add Array(cllBillList("操作员姓名"), "operator_name", "C"), "_operator_name" '   C   1   操作员姓名
            cllStuffBillList.Add Array(cllBillList("操作员编号"), "operator_code", "C"), "_operator_code" '   C   1   操作员编号
            cllStuffBillList.Add Array(cllBillList("登记时间"), "create_time", "C"), "_create_time" '   C   1   登记时间:yyyy-mm-dd hh:mi:ss
            cllStuffBillList.Add Array(cllStuffDetails, "item_list", "L"), "_item_list"
              
            '增加单据信息集
            cllStuffBillLists.Add Array(cllStuffBillList, "bill_list", "L"), "_" & cllBillList("单据号")
        End If
        
        If cllDrugDetails.Count <> 0 Then '存在药品
            str药品IDs = ""
            For k = 1 To cllDrugDetails.Count
                str药品IDs = str药品IDs & "," & cllDrugDetails(k)("_drug_id")(0)
            Next
            If intBillType <> 3 Then '记帐表作为“0-普通处方”
                byt处方类型 = zlGetRecipe_BillType(cllBillList("开单科室ID"), cllPati("病人ID"), cllPati("主页ID"), cllPati("年龄"), Mid(str药品IDs, 2))
            End If
            
            Set cllDrugBillList = New Collection
            cllDrugBillList.Add Array(zlGetRecipe_ID(cllBillList("单据号")), "recipe_id", "N"), "_recipe_id" '   N   1   处方id(新门诊):ZLHIS无，暂用NO转数字(字母用Asci代替)+序号替代
            cllDrugBillList.Add Array(cllBillList("单据号"), "rcp_no", "C"), "_rcp_no" '   C   1   NO
            cllDrugBillList.Add Array(byt处方类型, "recipe_type", "N"), "_recipe_type" '   C   0   处方类型:0和空-普通,1-儿科,2-急诊,3-精二,4-精一,5-麻醉
            cllDrugBillList.Add Array(IIf(cllBillList("是否划价") = 1, 0, 1), "charge_tag", "N"), "_charge_tag" '   N   1   收费标志:0-未收费或记帐划价;1-已收费或记帐
            cllDrugBillList.Add Array(IIf(cllBillList("是否划价") = 1, UserInfo.姓名, ""), "fee_acnter", "C"), "_fee_acnter" '   C       划价人
            cllDrugBillList.Add Array(cllBillList("开单科室ID"), "recipe_plcdept_id", "C"), "_recipe_plcdept_id" '   C       开单科室id（新门诊)
            cllDrugBillList.Add Array(zlGetDeptName(cllBillList("开单科室ID")), "recipe_plcdept", "C"), "_recipe_plcdept" '   C       开单科室名称（新门诊)
            cllDrugBillList.Add Array(zlGetPersonID(cllBillList("开单人")), "recipe_placer_id", "C"), "_recipe_placer_id" '   C       开单医师id（新门诊)
            cllDrugBillList.Add Array(cllBillList("开单人"), "recipe_placer", "C"), "_recipe_placer" '   C       开单医师（新门诊)   增加
            If intBillType <> 3 Then
                cllDrugBillList.Add Array(cllPati("医疗付款方式编码"), "apply_fee_category_code", "C"), "_apply_fee_category_code" '   C       申请单费别编码(医疗付款方式编码)(新门诊)    增加；
                cllDrugBillList.Add Array(cllPati("医疗付款方式名称"), "apply_fee_category_name", "C"), "_apply_fee_category_name" '   C       申请单费别名称（医疗付款方式名称）(新门诊)  增加；
            Else
                '记帐表暂不传
            End If
            cllDrugBillList.Add Array(cllBillList("操作员姓名"), "operator_name", "C"), "_operator_name" '   C   1   操作员姓名
            cllDrugBillList.Add Array(cllBillList("操作员编号"), "operator_code", "C"), "_operator_code" '   C   1   操作员编号
            cllDrugBillList.Add Array(cllBillList("登记时间"), "create_time", "C"), "_create_time" '   C   1   登记时间:yyyy-mm-dd hh:mi:ss
            cllDrugBillList.Add Array(cllDrugDetails, "item_list", "L"), "_item_list"
            
            '增加单据信息集
            cllDrugBillLists.Add Array(cllDrugBillList, "bill_list", "L"), "_" & cllBillList("单据号")
        End If
    Next
    
    If strDrugSendFeeIDs_Out <> "" Then strDrugSendFeeIDs_Out = Mid(strDrugSendFeeIDs_Out, 2)
    If strStuffSendFessIDs_Out <> "" Then strStuffSendFessIDs_Out = Mid(strStuffSendFessIDs_Out, 2)
    
    If strAutoSendDrugSendFeeIDs_Out <> "" Then strAutoSendDrugSendFeeIDs_Out = Mid(strAutoSendDrugSendFeeIDs_Out, 2)
    If strAutoSendStuffSendFeeIDs_Out <> "" Then strAutoSendStuffSendFeeIDs_Out = Mid(strAutoSendStuffSendFeeIDs_Out, 2)
    
    If Not cllDrugBillLists Is Nothing Then
        If cllDrugBillLists.Count <> 0 Then
            If GetRecipBillListsToJsonString(cllSvrPatiInfor, cllDrugBillLists, strJsonNewDrugRcp_out) = False Then Exit Function
        End If
    End If
    
    If Not cllStuffBillLists Is Nothing Then
        If cllStuffBillLists.Count <> 0 Then
            If GetRecipBillListsToJsonString(cllSvrPatiInfor, cllStuffBillLists, strJsonNewStuffRcp_out) = False Then Exit Function
        End If
    End If
    zlGetNewBill_SaveData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function zlGetSaveSQL_OutRecipeBill(ByVal cllPati As Collection, _
    ByVal cllBillList As Collection, ByVal cllBillDetail As Collection, _
    ByRef lng费用ID_out As Long, ByRef strSQL_Out As String, _
    Optional ByVal blnMarkSyncSymbol As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取保存门诊划价单、记帐单对应的SQL
    '入参:
    '   cllPati-病人信息集
    '   cllBillList-单据信息集
    '   cllBillDetail-单据明细集
    '   blnMarkSyncSymbol-是否更新同步标志
    '出参:
    '     lng费用ID_out-费用Id
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '  cllRcpBillData(结构)
    '   |-cllPati(Collect):病人信息，成员：
    '     (病人ID,主页ID,姓名,性别,年龄,费别,床号,病人科室ID,病区ID,险类,
    '      险类名称,病人来源,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '      患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '      身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):单据信息集=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):单据信息，成员：
    '       (单据号,是否多病人单,医疗小组id,医技补临床费用,简单记帐,记帐单id,记录性质,是否划价,是否急诊,加班标志,开单科室ID,
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,是否附费(N),[cllBillDetails(collect)])=cllBillists(_单据号)
    '       |-cllBillDetails(Collect):单据明细集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,从属父号,药名ID,收费细目ID,价格父号,收入项目ID,费别,婴儿序号,收费类别,计算单位,是否保险项目,保险大类ID,
    '            保险编码,收据费目,付数,数次,单价,应收金额,实收金额,统筹金额,附加标志,执行部门ID,是否自动发放,费用摘要(N),
    '            医嘱ID(N),就诊ID(N),费用类型,中药形态,煎法,执行性质,是否备货卫材(N),备货材料批次(N),是否跟踪在用,发药窗口(N),
    '            组内序号(N),诊疗项目ID(N),给药途径ID(N),给药途径名称(N),给药途径分类(N),给药频次ID(N)，给药频次名称（N),
    '            医嘱紧急标志(N),医嘱期效(N),计价特性(N),频次(N),单量（N),用法(N),皮试结果(N),超量说明(N),使用嘱托(N),
    '            发药方式(N),药品含量(N),门诊执行天数(N),煎法(N)
    '            【记帐表增加】:(病人ID,主页ID,姓名,性别,年龄,床号,病人科室ID,病区ID,险类,
    '              险类名称,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '              病人来源,患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '              身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N),
    '              医疗小组id)=cllBillDetails(_序号)
    '    以上元素中，带(N)的，表示可选节点
    Dim byt记录性质 As Byte, strSql As String
    
    On Error GoTo ErrHandler
    byt记录性质 = cllBillList("记录性质") '1-收费单;2-记帐单
    
    lng费用ID_out = gobjDatabase.GetNextId("病人费用记录")
    
    If byt记录性质 = 1 Then
        'ZL_门诊划价记录_INSERT_S
        strSql = "ZL_门诊划价记录_INSERT_S("
        '  No_In         门诊费用记录.No%Type,
        strSql = strSql & "'" & cllBillList("单据号") & "',"
        '  序号_In       门诊费用记录.序号%Type,
        strSql = strSql & "" & cllBillDetail("序号") & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSql = strSql & "" & ZVal(cllPati("病人ID")) & ","
        '  主页id_In     住院费用记录.主页id%Type,
        strSql = strSql & "" & ZVal(cllPati("主页ID")) & ","
        '  标识号_In     门诊费用记录.标识号%Type,
        strSql = strSql & "" & ZVal(cllPati("标识号")) & ","
        
        '  付款方式_In   门诊费用记录.付款方式%Type,
        strSql = strSql & "'" & cllPati("医疗付款方式编码") & "',"
        '  姓名_In       门诊费用记录.姓名%Type,
        strSql = strSql & "'" & cllPati("姓名") & "',"
        '  性别_In       门诊费用记录.性别%Type,
        strSql = strSql & "'" & cllPati("性别") & "',"
        '  年龄_In       门诊费用记录.年龄%Type,
        strSql = strSql & "'" & cllPati("年龄") & "',"
        '  费别_In       门诊费用记录.费别%Type,
        strSql = strSql & "'" & cllBillDetail("费别") & "',"
        
        '  加班标志_In   门诊费用记录.加班标志%Type,
        strSql = strSql & "" & ZVal(cllBillList("加班标志")) & ","
        '  病人科室id_In 门诊费用记录.病人科室id%Type,
        strSql = strSql & "" & ZVal(cllBillList("病人科室ID")) & ","
        '  开单部门id_In 门诊费用记录.开单部门id%Type,
        strSql = strSql & "" & ZVal(cllBillList("开单科室ID")) & ","
        '  开单人_In     门诊费用记录.开单人%Type,
        strSql = strSql & "'" & cllBillList("开单人") & "',"
        '  从属父号_In   门诊费用记录.从属父号%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("从属父号")) & ","
        
        '  收费细目id_In 门诊费用记录.收费细目id%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("收费细目ID")) & ","
        '  收费类别_In   门诊费用记录.收费类别%Type,
        strSql = strSql & "'" & cllBillDetail("收费类别") & "',"
        '  计算单位_In   门诊费用记录.计算单位%Type,
        strSql = strSql & "'" & cllBillDetail("计算单位") & "',"
        '  发药窗口_In   门诊费用记录.发药窗口%Type,
        strSql = strSql & "'" & zlGetNodeValueFromCollect(cllBillDetail, "发药窗口", "C") & "',"
        '  付数_In       门诊费用记录.付数%Type,
        strSql = strSql & "" & Val(cllBillDetail("付数")) & ","
        
        '  数次_In       门诊费用记录.数次%Type,
        strSql = strSql & "" & Val(cllBillDetail("数次")) & ","
        '  附加标志_In   门诊费用记录.附加标志%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("附加标志")) & ","
        '  执行部门id_In 门诊费用记录.执行部门id%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("执行部门ID")) & ","
        '  价格父号_In   门诊费用记录.价格父号%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("价格父号")) & ","
        '  收入项目id_In 门诊费用记录.收入项目id%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("收入项目ID")) & ","
        
        '  收据费目_In   门诊费用记录.收据费目%Type,
        strSql = strSql & "'" & cllBillDetail("收据费目") & "',"
        '  标准单价_In   门诊费用记录.标准单价%Type,
        strSql = strSql & "" & Val(cllBillDetail("单价")) & ","
        '  应收金额_In   门诊费用记录.应收金额%Type,
        strSql = strSql & "" & Val(cllBillDetail("应收金额")) & ","
        '  实收金额_In   门诊费用记录.实收金额%Type,
        strSql = strSql & "" & Val(cllBillDetail("实收金额")) & ","
        '  发生时间_In   门诊费用记录.发生时间%Type,
        strSql = strSql & " TO_Date('" & Format(cllBillList("发生时间"), "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        
        '  登记时间_In   门诊费用记录.登记时间%Type,
        strSql = strSql & " TO_Date('" & Format(cllBillList("登记时间"), "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
        strSql = strSql & "'" & cllBillList("操作员姓名") & "',"
        '  费用id_In     门诊费用记录.Id%Type := Null,
        strSql = strSql & "" & lng费用ID_out & ","
        '  费用摘要_In   门诊费用记录.摘要%Type := Null,
        strSql = strSql & "'" & cllBillDetail("费用摘要") & "',"
        '  医嘱序号_In   门诊费用记录.医嘱序号%Type := Null,
        strSql = strSql & "" & ZVal(cllBillDetail("医嘱ID")) & ","
        
        '  煎法_In       药品收发记录.用法%Type := Null,
        strSql = strSql & "'" & cllBillDetail("煎法") & "',"
        '  病人来源_In   Number := 1,
        strSql = strSql & "" & Val(cllPati("病人来源")) & ","
        '  保险编码_In   门诊费用记录.保险编码%Type := Null,
        strSql = strSql & "'" & cllBillDetail("保险编码") & "',"
        '  费用类型_In   门诊费用记录.费用类型%Type := Null,
        strSql = strSql & "'" & cllBillDetail("费用类型") & "',"
        '  保险项目否_In 门诊费用记录.保险项目否%Type := Null,
        strSql = strSql & "" & Val(cllBillDetail("是否保险项目")) & ","
        
        '  保险大类id_In 门诊费用记录.保险大类id%Type := Null,
        strSql = strSql & "" & ZVal(cllBillDetail("保险大类ID")) & ","
        '  中药形态_In   门诊费用记录.结论%Type := Null,
        If cllBillDetail("收费类别") = "7" Then
            strSql = strSql & "'" & cllBillDetail("中药形态") & "',"
        Else
            strSql = strSql & "NULL" & ","
        End If
        '  执行人_In     门诊费用记录.执行人%Type := Null,
        strSql = strSql & "NULL,"
        '  病人病区id_In 门诊费用记录.病人病区id%Type := Null
        strSql = strSql & "" & ZVal(cllPati("病区ID")) & ","
        '  批次_In         门诊费用记录.批次%Type := Null,
        strSql = strSql & "" & ZVal(zlGetNodeValueFromCollect(cllBillDetail, "备货材料批次", "N")) & ","
        
        '  更新同步标志_In Number := 1,
        strSql = strSql & "" & IIf(blnMarkSyncSymbol, 1, 0) & ","
        '  就诊id_In       病人挂号记录.Id%Type := Null,
        strSql = strSql & "" & ZVal(zlGetNodeValueFromCollect(cllBillDetail, "就诊ID", "N")) & ","
        '  是否急诊_In     病人挂号记录.急诊%Type := Null,
        strSql = strSql & "" & ZVal(zlGetNodeValueFromCollect(cllBillList, "是否急诊", "N")) & ","
        '  期效_In         门诊费用记录.医嘱期效%Type := 0,
        strSql = strSql & "NULL,"
        '  是否保密_In     门诊费用记录.是否保密%Type := 0,
        strSql = strSql & "NULL,"
        '  是否附费_In     门诊费用记录.是否附费%Type := 0
        strSql = strSql & "" & ZVal(zlGetNodeValueFromCollect(cllBillList, "是否附费", "N")) & ")"
    Else
        'Zl_门诊记帐记录_Insert_S
        strSql = "Zl_门诊记帐记录_Insert_S("
        '  No_In         门诊费用记录.No%Type,
        strSql = strSql & "'" & cllBillList("单据号") & "',"
        '  序号_In       门诊费用记录.序号%Type,
        strSql = strSql & "" & cllBillDetail("序号") & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSql = strSql & "" & ZVal(cllPati("病人ID")) & ","
        '  标识号_In     门诊费用记录.标识号%Type,
        strSql = strSql & "" & ZVal(cllPati("标识号")) & ","
        '  姓名_In       门诊费用记录.姓名%Type,
        strSql = strSql & "'" & cllPati("姓名") & "',"
        
        '  性别_In       门诊费用记录.性别%Type,
        strSql = strSql & "'" & cllPati("性别") & "',"
        '  年龄_In       门诊费用记录.年龄%Type,
        strSql = strSql & "'" & cllPati("年龄") & "',"
        '  费别_In       门诊费用记录.费别%Type,
        strSql = strSql & "'" & cllBillDetail("费别") & "',"
        '  加班标志_In   门诊费用记录.加班标志%Type,
        strSql = strSql & "" & ZVal(cllBillList("加班标志")) & ","
        '   婴儿费_In     门诊费用记录.婴儿费%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("婴儿序号")) & ","
        
        '  病人科室id_In 门诊费用记录.病人科室id%Type,
        strSql = strSql & "" & ZVal(cllBillList("病人科室ID")) & ","
        '  开单部门id_In 门诊费用记录.开单部门id%Type,
        strSql = strSql & "" & ZVal(cllBillList("开单科室ID")) & ","
        '  开单人_In     门诊费用记录.开单人%Type,
        strSql = strSql & "'" & cllBillList("开单人") & "',"
        '  从属父号_In   门诊费用记录.从属父号%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("从属父号")) & ","
        '  收费细目id_In 门诊费用记录.收费细目id%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("收费细目ID")) & ","
        
        '  收费类别_In   门诊费用记录.收费类别%Type,
        strSql = strSql & "'" & cllBillDetail("收费类别") & "',"
        '  计算单位_In   门诊费用记录.计算单位%Type,
        strSql = strSql & "'" & cllBillDetail("计算单位") & "',"
        '  付数_In       门诊费用记录.付数%Type,
        strSql = strSql & "" & Val(cllBillDetail("付数")) & ","
        '  数次_In       门诊费用记录.数次%Type,
        strSql = strSql & "" & Val(cllBillDetail("数次")) & ","
        '  附加标志_In   门诊费用记录.附加标志%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("附加标志")) & ","
        
        '  执行部门id_In 门诊费用记录.执行部门id%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("执行部门ID")) & ","
        '  价格父号_In   门诊费用记录.价格父号%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("价格父号")) & ","
        '  收入项目id_In 门诊费用记录.收入项目id%Type,
        strSql = strSql & "" & ZVal(cllBillDetail("收入项目ID")) & ","
        '  收据费目_In   门诊费用记录.收据费目%Type,
        strSql = strSql & "'" & cllBillDetail("收据费目") & "',"
        '  标准单价_In   门诊费用记录.标准单价%Type,
        strSql = strSql & "" & Val(cllBillDetail("单价")) & ","
        
        '  应收金额_In   门诊费用记录.应收金额%Type,
        strSql = strSql & "" & Val(cllBillDetail("应收金额")) & ","
        '  实收金额_In   门诊费用记录.实收金额%Type,
        strSql = strSql & "" & Val(cllBillDetail("实收金额")) & ","
        '  发生时间_In   门诊费用记录.发生时间%Type,
        strSql = strSql & " TO_Date('" & Format(cllBillList("发生时间"), "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  登记时间_In   门诊费用记录.登记时间%Type,
        strSql = strSql & " TO_Date('" & Format(cllBillList("登记时间"), "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  划价_In       Number,
        strSql = strSql & "" & IIf(Val(cllBillList("是否划价")), 1, 0) & " ,"
        
        '  发药窗口_In   门诊费用记录.发药窗口%Type,
        strSql = strSql & "'" & zlGetNodeValueFromCollect(cllBillDetail, "发药窗口", "C") & "',"
        '  操作员编号_In 门诊费用记录.操作员编号%Type,
        strSql = strSql & "'" & cllBillList("操作员编号") & "',"
        '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
        strSql = strSql & "'" & cllBillList("操作员姓名") & "',"
        '  费用id_In     门诊费用记录.Id%Type := Null,
        strSql = strSql & "" & lng费用ID_out & ","
        '  记帐单id_In   门诊费用记录.记帐单id%Type := Null,
        strSql = strSql & "" & ZVal(cllBillList("记帐单ID")) & ","
        
        '  费用摘要_In   门诊费用记录.摘要%Type := Null,
        strSql = strSql & "'" & cllBillDetail("费用摘要") & "',"
        '  医嘱序号_In   门诊费用记录.医嘱序号%Type := Null,
        strSql = strSql & "" & ZVal(cllBillDetail("医嘱ID")) & ","
        '  门诊标志_In   门诊费用记录.门诊标志%Type := 1,
        strSql = strSql & "" & cllPati("病人来源") & ","
        '  中药形态_In   门诊费用记录.结论%Type := Null,
        If cllBillDetail("收费类别") = "7" Then
            strSql = strSql & "'" & cllBillDetail("中药形态") & "',"
        Else
            strSql = strSql & "NULL" & ","
        End If
        '  煎法_In       药品收发记录.用法%Type := Null,
        strSql = strSql & "'" & cllBillDetail("煎法") & "',"
        
        '  主页id_In     门诊费用记录.主页id%Type := Null,
        strSql = strSql & "" & ZVal(cllPati("主页ID")) & ","
        '  病人病区id_In 门诊费用记录.病人病区id%Type := Null
        strSql = strSql & "" & ZVal(cllPati("病区ID")) & ","
        '  批次_In         门诊费用记录.批次%Type := Null,
        strSql = strSql & "" & ZVal(zlGetNodeValueFromCollect(cllBillDetail, "备货材料批次", "N")) & ","
        '  更新同步标志_In Number := 1,
        strSql = strSql & "" & IIf(blnMarkSyncSymbol, 1, 0) & ","
        '  就诊id_In       病人挂号记录.Id%Type := Null,
        strSql = strSql & "" & ZVal(zlGetNodeValueFromCollect(cllBillDetail, "就诊ID", "N")) & ","
        
        '  是否急诊_In     病人挂号记录.急诊%Type := Null
        strSql = strSql & "" & ZVal(zlGetNodeValueFromCollect(cllBillList, "是否急诊", "N")) & ","
        '  期效_In         门诊费用记录.医嘱期效%Type := 0,
        strSql = strSql & "NULL,"
        '  是否保密_In     门诊费用记录.是否保密%Type := 0,
        strSql = strSql & "NULL,"
        '  是否附费_In     门诊费用记录.是否附费%Type := 0
        strSql = strSql & "" & ZVal(zlGetNodeValueFromCollect(cllBillList, "是否附费", "N")) & ")"
    End If
    strSQL_Out = strSql
    
    zlGetSaveSQL_OutRecipeBill = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function zlExcute_DelRecipeBill_Check(cllBillLists As Collection, ByVal intBillType As Integer, _
    Optional ByVal lngModule As Long, Optional ByVal bln划价 As Boolean, _
    Optional ByVal bln门诊 As Boolean, Optional ByVal cllPatients As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行处方作废前检查操作
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    '   intBillType-单据类型(1-收费单;2-记帐单;3-自动记帐单)
    '   cllPatients(Collect):-病人信息集，仅住院记帐传入
    '     |-cllPatient(Collect):-每个病人信息，成员：
    '       (病人ID,主页ID,险类,审核标志,住院状态,编目日期)=cllPatients(_病人ID)
    '出参:
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-14 14:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnIsExist  As Boolean, p As Long, strNo As String
    Dim cllFeeCheck_out As Collection, cllBillLists_Out As Collection
    Dim cllBillList As Collection, i As Long
    Dim rsRecipeExecuted As ADODB.Recordset, cllAdviceDatas As Collection, cllAdviceDatas_NO As Collection
    Dim cllBillDetails As Collection, cllBillDetail As Collection
    Dim cllDeleteData As Collection, cllTemp As Collection
    Dim str医嘱ID_No As String, varData As Variant, rsAdviceData As ADODB.Recordset
    Dim cllAdviceList As Collection, cllAdviceList_out As Collection
    Dim objService As clsService
    
    On Error GoTo errHandle
    If lngModule <> 0 Then mlngModule = lngModule
    If GetService(objService) = False Then Exit Function
    
    '先合法性数据检查
    '1.先组织相关数据
    If zlGetDelBill_CheckData(cllBillLists, cllFeeCheck_out, intBillType, bln划价, bln门诊, _
        rsRecipeExecuted, , , cllPatients) = False Then Exit Function
    
    '2. 费用数据合法性较验
    If zlFeeDelRecipeBill_Check(cllFeeCheck_out, intBillType, bln划价, cllBillLists_Out, bln门诊) = False Then Exit Function
    '更新销账数量和执行状态
    For p = 1 To cllBillLists.Count
        Set cllBillList = cllBillLists(p)
        Call cllBillList.Remove("_cllBillDetails")
        cllBillList.Add cllBillLists_Out("_" & cllBillList("单据号"))("_cllBillDetails"), "_cllBillDetails"
    Next
    
    '门诊退费不用处理医嘱的数据
    If Not (intBillType = 1 And Not bln划价) Then
        '3.检查正在执行医嘱
        Set cllAdviceDatas = New Collection
        For p = 1 To cllBillLists.Count
            Set cllBillList = cllBillLists(p)
             
            strNo = cllBillList("单据号")
            Set cllBillDetails = cllBillList("_cllBillDetails")
            For i = 1 To cllBillDetails.Count
                Set cllBillDetail = cllBillDetails(i)
                strNo = strNo & IIf(i = 1, ":", ",") & cllBillDetail("序号")
            Next
            
            If CheckExistAdviceFeeByDetail(strNo, intBillType, bln门诊, str医嘱ID_No, True) = False Then Exit Function
            If str医嘱ID_No <> "" Then
                varData = Split(str医嘱ID_No, ",")
                For i = 0 To UBound(varData)
                    Set cllAdviceDatas_NO = New Collection
                    cllAdviceDatas_NO.Add varData(i), "医嘱ID"
                    cllAdviceDatas_NO.Add cllBillList("单据号"), "费用单号"
                Next
                cllAdviceDatas.Add cllAdviceDatas_NO, "_" & cllBillList("单据号")
            End If
            cllBillList.Add str医嘱ID_No, "医嘱ID" '临时使用
        Next
        
        If objService.ZlCissvr_AdviceExecuting(intBillType, cllAdviceDatas, blnIsExist, mlngModule) = False Then Exit Function
        If blnIsExist Then
            MsgBox "单据中存在对应的医嘱正在执行的情况，不能删除！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '4.医嘱计费状态计算
        For p = 1 To cllBillLists.Count
            Set cllBillList = cllBillLists(p)
            
            Set cllDeleteData = New Collection
            Set cllBillDetails = cllBillList("_cllBillDetails")
            For i = 1 To cllBillDetails.Count
                Set cllBillDetail = cllBillDetails(i)
                
                Set cllTemp = New Collection
                cllTemp.Add cllBillDetail("序号"), "序号"
                cllTemp.Add cllBillDetail("销帐数量"), "销帐数量"
                cllDeleteData.Add cllTemp
            Next
            
            Set cllAdviceList_out = New Collection
            str医嘱ID_No = cllBillList("医嘱ID")
            If str医嘱ID_No = "" Then
                '可能是医嘱附费单据
                If CheckExistAdviceFeeByDetail(strNo, intBillType, bln门诊, str医嘱ID_No, True, True) = False Then Exit Function
                
            Else
                If objService.ZlCissvr_GetAllGroupAdviceIDs(str医嘱ID_No, rsAdviceData, mlngModule) = False Then Exit Function
                Set cllAdviceList = New Collection
                Do While Not rsAdviceData.EOF
                    Set cllTemp = New Collection
                    cllTemp.Add Nvl(rsAdviceData!医嘱id), "医嘱ID"
                    cllTemp.Add Nvl(rsAdviceData!相关ID), "相关ID"
                    cllTemp.Add Nvl(rsAdviceData!诊疗类别), "诊疗类别"
                    cllAdviceList.Add cllTemp
                    
                    rsAdviceData.MoveNext
                Loop
            End If
            
            If str医嘱ID_No <> "" Then
                If Zl_医嘱发送_计费状态_Get(intBillType, cllBillList("单据号"), bln门诊, IIf(bln划价, 0, 2), _
                    str医嘱ID_No, cllDeleteData, cllAdviceList, cllAdviceList_out, mlngModule) = False Then Exit Function
            End If
            cllBillList.Add cllAdviceList_out, "_cllAdviceUpdateDatas"
        Next
    End If
    
    zlExcute_DelRecipeBill_Check = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetDelBill_CheckData(cllBillLists As Collection, ByRef cllFeeCheck_out As Collection, _
    Optional ByVal intBillType As Integer, Optional bln划价 As Boolean, Optional bln门诊 As Boolean, _
    Optional rsRecipeExecuted_Out As ADODB.Recordset, _
    Optional ByVal blnNotShowErr As Boolean, Optional ByRef strErrmsg_Out As String, _
    Optional ByVal cllPatients As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行处方作废前检查操作
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    '   intBillType-单据类型(1-收费单;2-记帐单;3-自动记帐单)
    '   cllPatients(Collect):-病人信息集，仅住院记帐传入
    '     |-cllPatient(Collect):-每个病人信息，成员：
    '       (病人ID,主页ID,险类,审核标志,住院状态,编目日期)=cllPatients(_病人ID)
    '出参:
    '   cllFeeCheck_Out-Array(no,费用检查json),key(_NO)
    '   rsRecipeExecuted_Out-返回当前的调用服务后药品\卫材的已发药或发料的数据，以便外部不用再调用该服务
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-14 14:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strNOs As String, strNo As String, strJsonFeeCheck As String, p As Integer
    Dim rsRecipeExecuted As ADODB.Recordset
    Dim clldata As Collection, cllDatas As Collection
    Dim cllBakCheck As Collection, cllBakChecks As Collection, cllBillList As Collection
    Dim cllBillDetails As Collection, cllBillDetail As Collection
    Dim blnExistDrug As Boolean, blnExistStuff As Boolean
    Dim cllAdviceData As Collection, rsAdviceExecuted As ADODB.Recordset '已执行数据：医嘱序号,NO,收费细目ID,已执行数
    Dim cllPatient As Collection, objService As clsService
    
    On Error GoTo errHandle
    If GetService(objService) = False Then Exit Function
    Set cllFeeCheck_out = New Collection
    Set cllBakChecks = New Collection
    
    '1.构建单据信息集
    '  --      item_list[]             本次销帐列表
    '  --          serial_num          N   1   序号
    '  --          quantity            N   1   销帐数量(为零时，按序号直接销帐)
    For p = 1 To cllBillLists.Count
        Set cllBillList = cllBillLists(p)
        
        strNOs = strNOs & "," & cllBillList("单据号")
        
        '  --      fee_no                  C   1   费用单据号
        '  --      fee_bill_type           N   1   单据性质:2-门诊记帐单,3-自动记帐单
        '  --      balance_ban_writeoffs   N   1   已结禁止销帐:如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
        '  --      part_ban_writeoffs      N   1   禁止部分销帐:1-不允许；0-允许
        '  --      oper_type               N   1   操作状态：0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
        Set cllBakCheck = New Collection
        cllBakCheck.Add Array(cllBillList("单据号"), "fee_no", "C"), "_fee_no"
        cllBakCheck.Add Array(intBillType, "fee_bill_type", "N"), "_fee_bill_type"
        If intBillType <> 1 Then
            cllBakCheck.Add Array(cllBillList("已结禁止销帐"), "balance_ban_writeoffs", "N"), "_balance_ban_writeoffs"
            cllBakCheck.Add Array(cllBillList("禁止部分销帐"), "part_ban_writeoffs", "N"), "_part_ban_writeoffs"
            cllBakCheck.Add Array(zlGetNodeValueFromCollect(cllBillList, "操作状态", "N"), "oper_type", "N"), "_oper_type"
        End If
        
        Set cllBillDetails = cllBillList("_cllBillDetails")
        
        Set cllDatas = New Collection
        For i = 1 To cllBillDetails.Count
            
            Set cllBillDetail = cllBillDetails(i)
            
            Set clldata = New Collection
            Select Case intBillType
            Case 1 '收费
                If bln划价 Then
                    clldata.Add Array(cllBillDetail("序号"), "serial_num", "N"), "_serial_num" 'N1序号
                Else
                    clldata.Add Array(cllBillDetail("序号"), "serial_num", "N"), "_serial_num" 'N1序号
                    clldata.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "销帐数量", "N"), "charge_offs_num", "N"), "_charge_offs_num" 'N1销帐数量
                End If
                    
            Case 2, 3 '记帐
                clldata.Add Array(cllBillDetail("序号"), "serial_num", "N"), "_serial_num" 'N1序号
                clldata.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "销帐数量", "N"), "quantity", "N"), "_quantity" 'N1销帐数量
                
            Case Else
                Exit Function
            End Select
        
            cllDatas.Add clldata, "_" & cllBillDetail("序号")
        Next
        cllBakCheck.Add Array(cllDatas, "item_list", "Collection"), "_item_list"
        cllBakChecks.Add cllBakCheck, cllBillList("单据号")
    Next
    If strNOs = "" Then Exit Function
    
    strNOs = Mid(strNOs, 2)
    '2.构建已执行数量信息集
    '  --      excute_list[]           单据已执行列表(药品、卫材费用),即使已执行数为0也要传入
    '  --          fee_id              N   1   费用ID
    '  --          sended_num          N   1   已发数量
    '  --      advice_excute_list[]    单据已执行列表(医嘱费用),即使已执行数为0也要传入
    '  --          advice_id           N   1   医嘱ID
    '  --          fee_item_id         N   1   收费细目ID
    '  --          execute_num         N   1   已执行数
    If CheckExistDrugOrStuff(strNOs, intBillType, bln门诊, blnExistDrug, blnExistStuff) = False Then Exit Function
    If zlGetDrugStuff_ExecutedNum(strNOs, intBillType, rsRecipeExecuted, blnExistDrug, blnExistStuff) = False Then Exit Function
    Set rsRecipeExecuted_Out = rsRecipeExecuted
    
    If CheckExistAdviceFeeByNo(strNOs, intBillType, bln门诊, cllAdviceData) = False Then Exit Function
    If objService.ZlCisSvr_GetAdviceExcutNums(intBillType, cllAdviceData, rsAdviceExecuted, mlngModule) = False Then Exit Function
    
    For i = 1 To cllBakChecks.Count
        Set cllBakCheck = cllBakChecks(i)
        strNo = cllBakCheck("_fee_no")(0)
        
        '药品\卫材费用
        If Not rsRecipeExecuted Is Nothing Then
            Set cllDatas = New Collection
            rsRecipeExecuted.Filter = "处方单号 = '" & strNo & "'"
            Do While Not rsRecipeExecuted.EOF
                Set clldata = New Collection
                clldata.Add Array(Nvl(rsRecipeExecuted!处方明细id), "fee_id", "N"), "_fee_id"
                clldata.Add Array(Nvl(rsRecipeExecuted!已发数量), "sended_num", "N"), "_sended_num"
                cllDatas.Add clldata, "_" & Nvl(rsRecipeExecuted!处方明细id)
                rsRecipeExecuted.MoveNext
            Loop
            If cllDatas.Count > 0 Then
                cllBakCheck.Add Array(cllDatas, "excute_list", "Collection"), "_excute_list"
            End If
        End If
        
        '医嘱费用
        If Not rsAdviceExecuted Is Nothing Then
            Set cllDatas = New Collection
            rsAdviceExecuted.Filter = "No = '" & strNo & "'"
            Do While Not rsAdviceExecuted.EOF
                Set clldata = New Collection
                clldata.Add Array(Nvl(rsAdviceExecuted!医嘱序号), "advice_id", "N"), "_advice_id"  '
                clldata.Add Array(Nvl(rsAdviceExecuted!收费细目ID), "fee_item_id", "N"), "_fee_item_id"
                clldata.Add Array(Nvl(rsAdviceExecuted!已执行数), "execute_num", "N"), "_execute_num"
                cllDatas.Add clldata, "_" & Nvl(rsAdviceExecuted!收费细目ID)
                rsAdviceExecuted.MoveNext
            Loop
            If cllDatas.Count > 0 Then
                cllBakCheck.Add Array(cllDatas, "advice_excute_list", "Collection"), "_advice_excute_list"
            End If
        End If
    Next
    
    '3.住院记帐，需要传病人信息集
    '  --      pati_list[]             病人信息，仅审核这些病人的费用
    '  --          pati_id             N   1   病人ID
    '  --          fee_audit_status    N   1   费用审核标志:0或空-未审核;1-已审核或开始审核(结合参数:病人审核方式来控制);2-完成审核,结合结帐权限[禁止未审核病人结帐]进行管理控制
    '  --          si_inp_status       N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
    '  --          catalog_date        C   0   病案编目日期：yyyy-mm-dd hh24:mi:ss
    If (intBillType = 2 Or intBillType = 3) And Not bln门诊 Then
        For p = 1 To cllBakChecks.Count
            Set cllBakCheck = cllBakChecks(p)
            
            Set cllDatas = New Collection
            For i = 1 To cllPatients.Count
                Set cllPatient = cllPatients(i)
                Set clldata = New Collection
                clldata.Add Array(cllPatient("病人ID"), "pati_id", "N"), "_pati_id"
                clldata.Add Array(cllPatient("审核标志"), "fee_audit_status", "N"), "_fee_audit_status"
                clldata.Add Array(cllPatient("住院状态"), "si_inp_status", "N"), "_si_inp_status"
                clldata.Add Array(cllPatient("编目日期"), "catalog_date", "C"), "_catalog_date"
                cllDatas.Add clldata, "_" & cllPatient("病人ID")
            Next
            If cllDatas.Count > 0 Then
                cllBakCheck.Add Array(cllDatas, "pati_list", "Collection"), "_pati_list"
            End If
        Next
    End If
    
    '组织作废单据检查条件的Json入参串
    For i = 1 To cllBakChecks.Count
        Set cllBakCheck = cllBakChecks(i)
        strNo = cllBakCheck("_fee_no")(0)
        strJsonFeeCheck = DelRecipeBill_Check_GetJsonString(cllBakCheck)
        cllFeeCheck_out.Add Array(strNo, strJsonFeeCheck), "_" & strNo
    Next
    
    zlGetDelBill_CheckData = True
    Exit Function
errHandle:
    If blnNotShowErr Then strErrmsg_Out = Err.Description: Exit Function
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlExcute_DelRecipeBill(ByVal cllBillLists As Collection, _
    ByVal strFrmCaption As String, ByVal intBillType As Integer, _
    Optional ByVal bln划价 As Boolean, Optional ByVal lngModule As Long, _
    Optional bln门诊 As Boolean, Optional ByVal cllPro As Collection, _
    Optional ByVal cllPatients As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行处方数据作废操作
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    '   intBillType-单据类型(1-收费单;2-记帐单;3-自动记帐单)
    '   cllPro 需要一起执行的SQL语句
    '   cllPatients(Collect):-病人信息集，仅住院记帐传入
    '     |-cllPatient(Collect):-每个病人信息，成员：
    '       (病人ID,主页ID,险类,审核标志,住院状态,编目日期)=cllPatients(_病人ID)
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-13 22:52:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJosnDelDrugRcp_Out As String, strJosnDelStuffRcp_Out As String
    Dim i As Long, strInfo As String
    Dim p As Long, cllTemp As Collection
    Dim strDrugSendFeeIDs_Out  As String, strStuffSendFeeIDs_Out As String, strNo As String
    Dim blnTrans As Boolean, strInsure As String, intInsure As Integer
    Dim cllAdviceUpdateDatas As Collection, cllAdviceUpdateData As Collection
    Dim cllAdviceDatas As Collection, str医嘱IDs As String
    Dim objService As clsService
    
    On Error GoTo errHandle
    If GetService(objService) = False Then Exit Function
    If lngModule <> 0 Then mlngModule = lngModule
    If cllPro Is Nothing Then Set cllPro = New Collection
    
    '1.先收集费用保存数据过程及生成处方数据信息
    If zlGetDelBill_SaveData(cllBillLists, cllPro, _
        strJosnDelDrugRcp_Out, strJosnDelStuffRcp_Out, _
        strDrugSendFeeIDs_Out, strStuffSendFeeIDs_Out, intBillType, bln划价, bln门诊) = False Then Exit Function
    If cllPro.Count = 0 Then Exit Function
    
    '收集险类
    strInsure = ""
    If intBillType <> 1 And Not bln划价 And Not bln门诊 Then
        For i = 1 To cllPatients.Count
            intInsure = Val(cllPatients(i)("险类"))
            If intInsure <> 0 And InStr("," & strInsure & ",", "," & intInsure & ",") = 0 Then
                strInsure = strInsure & "," & intInsure
            End If
        Next
        If strInsure <> "" Then strInsure = Mid(strInsure, 2)
    End If
    
    str医嘱IDs = ""
    Set cllAdviceDatas = New Collection
    For p = 1 To cllBillLists.Count
        Set cllAdviceUpdateDatas = cllBillLists(p)("_cllAdviceUpdateDatas")
        If Not cllAdviceUpdateDatas Is Nothing Then
            For i = 1 To cllAdviceUpdateDatas.Count
                Set cllAdviceUpdateData = cllAdviceUpdateDatas(i)
                Set cllTemp = New Collection
                cllTemp.Add cllAdviceUpdateData("医嘱ID"), "医嘱ID"
                cllTemp.Add cllBillLists(p)("单据号"), "费用单号"
                cllTemp.Add cllAdviceUpdateData("计费状态"), "计费状态"
                cllTemp.Add cllAdviceUpdateData("删除附费"), "删除附费"
                
                If InStr("," & str医嘱IDs & ",", "," & cllAdviceUpdateData("医嘱ID") & ",") = 0 Then
                    str医嘱IDs = str医嘱IDs & "," & cllAdviceUpdateData("医嘱ID")
                End If
                
                cllAdviceDatas.Add cllTemp
            Next
        End If
    Next
    If str医嘱IDs <> "" Then str医嘱IDs = Mid(str医嘱IDs, 2)
    
    '2.存在药品，需要作废处方
    If strJosnDelDrugRcp_Out <> "" Then
        If zlExcute_SaveDrugDelRecipeBill(strDrugSendFeeIDs_Out, strJosnDelDrugRcp_Out, bln门诊) = False Then Exit Function
    End If
    '3.存在卫生材料，需要作废发料单
    If strJosnDelStuffRcp_Out <> "" Then
        If zlExcute_SaveStuffDelBill(strStuffSendFeeIDs_Out, strJosnDelStuffRcp_Out, bln门诊) = False Then Exit Function
    End If
    
    '4.执行相关费用数据保存
    gcnOracle.BeginTrans: blnTrans = True
        zlExecuteProcedureArrAy cllPro, strFrmCaption, True, True
        
        If str医嘱IDs <> "" Then
            '更改医嘱数据的计费状态及删除附费数据
            If cllAdviceDatas.Count > 0 Then
                If objService.ZlCisSvr_UpdAdviceChargeTag(intBillType, cllAdviceDatas, Nothing, mlngModule) = False Then
                    gcnOracle.RollbackTrans: Exit Function
                End If
            End If
        ElseIf strInsure <> "" Then
            '医保记帐审核完成前上传
            If InsureTranChargeDetail(strNo, strInsure, strInfo, False, True) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                If strInfo <> "" Then MsgBox strInfo, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '医保记帐作废完成后上传(有医嘱费用时，不分记帐完成前与后)
    If strInsure <> "" Then
        If InsureTranChargeDetail(strNo, strInsure, strInfo, True, True, str医嘱IDs <> "") = False Then
            If strInfo <> "" Then
                MsgBox strInfo, vbInformation, gstrSysName
            Else
                MsgBox "单据""" & strNo & """的数据向医保传送失败，该单据已作废！", vbInformation, gstrSysName
            End If
        End If
    End If
    
    zlExcute_DelRecipeBill = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function zlGetDelBill_SaveData(ByVal cllBillLists As Collection, ByRef cllPro_out As Collection, ByRef strJsonDelDrugRcp_out As String, _
    ByRef strJsonDelStuffRcp_out As String, ByRef strDrugSendFeeIDs_Out As String, ByRef strStuffSendFeeIDs_Out As String, _
    Optional ByVal intBillType As Integer, Optional ByVal bln划价 As Boolean, Optional bln门诊 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行处方数据作废操作
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    '   intBillType-单据类型(1-收费处方;2-记帐单处方;3-自动记帐单)
    '   cllPro 需要一起执行的SQL语句
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim strSql As String, str序号 As String, strNo As String
    Dim rsTmp As ADODB.Recordset
    Dim cllBillList As Collection
    Dim cllBillDetails As Collection, cllBillDetail As Collection
    Dim cllDrugDetails As Collection, cllStuffDetails As Collection
    Dim cllDrugDetail As Collection, cllStuffDetail As Collection
    
    On Error GoTo errHandle
    strJsonDelDrugRcp_out = "": strJsonDelStuffRcp_out = ""
    Set cllDrugDetails = New Collection
    Set cllStuffDetails = New Collection
    
    For i = 1 To cllBillLists.Count
        Set cllBillList = cllBillLists(i)
        Set cllBillDetails = cllBillList("_cllBillDetails")
        strNo = cllBillList("单据号")
        str序号 = ""
        
        For j = 1 To cllBillDetails.Count
            Set cllBillDetail = cllBillDetails(j)
            str序号 = str序号 & "," & cllBillDetail("序号")
        Next
        If str序号 <> "" Then str序号 = Mid(str序号, 2)
        
        If bln门诊 Then
            strSql = "Select Max(Decode(a.记录状态,2,0,Decode(a.记录性质,[2],a.Id,0))) As 费用ID," & _
                    "        a.序号, a.收费类别, Sum(Nvl(a.付数,1)*a.数次) As 剩余数次, w.跟踪在用," & _
                    "        Max(Decode(a.记录状态,2,0,Decode(a.记录性质,[2],a.记费同步标志,0))) As 记费同步标志," & _
                    "        Max(Decode(a.记录状态,2,0,Decode(a.记录性质,[2],a.作废同步标志,0))) As 作废同步标志" & _
                    " From 门诊费用记录 A, 材料特性 W" & _
                    " Where a.No = [1] And Mod(a.记录性质,10) = [2] And a.收费细目ID=w.材料ID(+) And a.价格父号 Is Null" & _
                            IIf(str序号 = "", "", " And Instr(','||[3]||',', ','||Nvl(a.价格父号,a.序号)||',') > 0") & _
                    " Group By a.序号, a.收费类别, w.跟踪在用" & _
                    " Having Nvl(Sum(Nvl(a.付数,1)*a.数次),0)<>0"
        Else
            strSql = "Select Max(Decode(a.记录状态,2,0,a.Id)) As 费用ID," & _
                    "        a.序号, a.收费类别, Sum(Nvl(a.付数,1)*a.数次) As 剩余数次, w.跟踪在用," & _
                    "        Max(Decode(a.记录状态,2,0,a.记费同步标志)) As 记费同步标志," & _
                    "        Max(Decode(a.记录状态,2,0,a.作废同步标志)) As 作废同步标志" & _
                    " From 住院费用记录 a, 材料特性 w" & vbNewLine & _
                    " Where a.No = [1] And a.记录性质 = [2] And a.收费细目ID=w.材料ID(+) And a.价格父号 Is Null" & _
                            IIf(str序号 = "", "", " And Instr(','||[3]||',', ','||Nvl(a.价格父号,a.序号)||',') > 0") & _
                    " Group By a.序号, a.收费类别, w.跟踪在用" & _
                    " Having Nvl(Sum(Nvl(a.付数,1)*a.数次),0)<>0"
        End If
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取退费明细", strNo, intBillType, str序号)
        If rsTmp.EOF Then
            MsgBox "未找到单据号为【" & strNo & "】的单据信息！", vbInformation, gstrSysName
            Exit Function
        End If
                
        '1.构建单据信息集
        Select Case intBillType
        Case 1 '划价、收费
            If zlGetDelSQL_OutRecipeBill(cllBillList, intBillType, bln划价, cllPro_out) = False Then Exit Function
        Case 2, 3 '记帐
            If Not bln门诊 Then
                If zlGetDelSQL_InRecipeBill(cllBillList, intBillType, cllPro_out) = False Then Exit Function
            Else '门诊记帐
                If zlGetDelSQL_OutRecipeBill(cllBillList, intBillType, False, cllPro_out) = False Then Exit Function
            End If
        Case Else
            Exit Function
        End Select

        With rsTmp
            Do While Not .EOF
                Set cllBillDetail = cllBillDetails("_" & Nvl(!序号))
                
                '未生成处方的和已作废的不再调用药品/卫材作废服务
                If Val(Nvl(!记费同步标志)) <> 1 And Val(Nvl(!作废同步标志)) = 0 Then
                    If InStr(",5,6,7,", "," & Nvl(!收费类别) & ",") > 0 Then
                        Set cllDrugDetail = New Collection
                        cllDrugDetail.Add Array(Nvl(!费用id), "rcpdtl_id", "N"), "_rcpdtl_id" 'N1处方明细id,目前传入的是费用id
                        cllDrugDetail.Add Array(cllBillDetail("销帐数量"), "chargeoffs_num", "N"), "_chargeoffs_num" 'N1销账数量
                        cllDrugDetail.Add Array(zlGetNodeValueFromCollect(cllBillDetail, "配药IDs", "C"), "dispensing_ids", "C"), "_dispensing_ids"    'C1配药记录IDs
                        cllDrugDetails.Add cllDrugDetail, "_" & Nvl(!费用id)
                        
                        strDrugSendFeeIDs_Out = strDrugSendFeeIDs_Out & "," & Nvl(!费用id)
                    ElseIf Nvl(!收费类别) = "4" And Val(Nvl(!跟踪在用)) = 1 Then
                        Set cllStuffDetail = New Collection
                        cllStuffDetail.Add Array(Nvl(!费用id), "stuffdtl_id", "N"), "_stuffdtl_id" 'N1处方明细id,目前传入的是费用id
                        cllStuffDetail.Add Array(cllBillDetail("销帐数量"), "return_num", "N"), "_return_num" 'N1退料数量
                        cllStuffDetails.Add cllStuffDetail, "_" & Nvl(!费用id)
                        
                        strStuffSendFeeIDs_Out = strStuffSendFeeIDs_Out & "," & Nvl(!费用id)
                    End If
                End If
                rsTmp.MoveNext
            Loop
        End With
    Next
    
    If cllDrugDetails.Count <> 0 Then
        If GetRecipeBillDetailsToJson(cllDrugDetails, strJsonDelDrugRcp_out) = False Then Exit Function
        strJsonDelDrugRcp_out = "{""input"":{" & strJsonDelDrugRcp_out & "}}"
        
        strDrugSendFeeIDs_Out = Mid(strDrugSendFeeIDs_Out, 2)
    End If
    
    If cllStuffDetails.Count <> 0 Then
        If GetRecipeBillDetailsToJson(cllStuffDetails, strJsonDelStuffRcp_out) = False Then Exit Function
        strJsonDelStuffRcp_out = "{""input"":{" & strJsonDelStuffRcp_out & "}}"
        
        strStuffSendFeeIDs_Out = Mid(strStuffSendFeeIDs_Out, 2)
    End If
    
    zlGetDelBill_SaveData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetDelSQL_InRecipeBill(ByVal cllBillList As Collection, _
    ByVal intBillType As Integer, ByRef cllPro_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取作废门诊划价单、记帐单对应的SQL
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    '   intBillType-单据类型(1-收费处方;2-记帐单处方;3-记帐表;4-自动记帐单)
    '出参:
    '
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, str序号 As String
    Dim cllBillDetails As Collection, cllBillDetail As Collection, cll序号 As Collection
    Dim i As Long
    
    On Error GoTo ErrHandler
    Set cllBillDetails = cllBillList("_cllBillDetails")
    
    Set cll序号 = New Collection
    For i = 1 To cllBillDetails.Count
        Set cllBillDetail = cllBillDetails(i)
        str序号 = str序号 & "," & cllBillDetail("序号") & ":" & cllBillDetail("销帐数量") & ":" & cllBillDetail("执行状态")
        
        If Len(str序号) > 3950 Then
            cll序号.Add Mid(str序号, 2)
            str序号 = ""
        End If
    Next
    If str序号 <> "" Then cll序号.Add Mid(str序号, 2)
    
    For i = 1 To cll序号.Count
        'Zl_住院记帐记录_Delete_S
        strSql = "Zl_住院记帐记录_Delete_S("
        '  No_In         住院费用记录.No%Type,
        strSql = strSql & "'" & cllBillList("单据号") & "',"
        '  序号_In       Varchar2,
        strSql = strSql & "'" & cll序号(i) & "',"
        '  操作员编号_In 门诊费用记录.操作员编号%Type,
        strSql = strSql & "'" & cllBillList("操作员编号") & "',"
        '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
        strSql = strSql & "'" & cllBillList("操作员姓名") & "',"
        '  记录性质_In   住院费用记录.记录性质%Type := 2,
        strSql = strSql & "" & intBillType & ","
        '  操作状态_In   Number := 0,
        strSql = strSql & "" & ZVal(zlGetNodeValueFromCollect(cllBillList, "操作状态", "N")) & ","
        '  登记时间_In   门诊费用记录.登记时间%Type,
        strSql = strSql & " TO_Date('" & Format(cllBillList("登记时间"), "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))"
    
        zlAddArray cllPro_out, strSql
    Next
    zlGetDelSQL_InRecipeBill = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function zlGetDelSQL_OutRecipeBill(ByVal cllBillList As Collection, ByVal intBillType As Integer, _
                ByVal bln划价 As Boolean, ByRef cllPro_out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取作废门诊划价单、记帐单对应的SQL
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    '    intBillType-单据类型(1-收费处方;2-记帐单处方;3-记帐表;4-自动记帐单)
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBillDetails As Collection, cllBillDetail As Collection, cll序号 As Collection
    Dim strSql As String, str序号 As String
    Dim i As Long
    
    On Error GoTo ErrHandler
    Set cllBillDetails = cllBillList("_cllBillDetails")
    
    Set cll序号 = New Collection
    If intBillType = 1 Then
        If bln划价 Then
            For i = 1 To cllBillDetails.Count
                Set cllBillDetail = cllBillDetails(i)
                str序号 = str序号 & "," & cllBillDetail("序号")
                
                If Len(str序号) > 3950 Then
                    cll序号.Add Mid(str序号, 2)
                    str序号 = ""
                End If
            Next
            If str序号 <> "" Then cll序号.Add Mid(str序号, 2)
            
            For i = 1 To cll序号.Count
                'zl_门诊划价记录_DELETE_S
                '  --功能：删除一张门诊划价单据
                '  --参数：
                '  --  序号_In：格式: 序号1,序号2,...
                strSql = "zl_门诊划价记录_DELETE_S("
                '  No_In         门诊费用记录.No%Type,
                strSql = strSql & "'" & cllBillList("单据号") & "',"
                '  序号_In       Varchar2,
                strSql = strSql & "'" & cll序号(i) & "')"
                
                zlAddArray cllPro_out, strSql
            Next
        Else '退费
            For i = 1 To cllBillDetails.Count
                Set cllBillDetail = cllBillDetails(i)
                str序号 = str序号 & "," & cllBillDetail("序号") & ":" & cllBillDetail("销帐数量") & ":" & cllBillDetail("执行状态")
                
                If Len(str序号) > 3950 Then
                    cll序号.Add Mid(str序号, 2)
                    str序号 = ""
                End If
            Next
            If str序号 <> "" Then cll序号.Add Mid(str序号, 2)
            
            For i = 1 To cll序号.Count
                'Zl_门诊收费记录_销帐_S
                '  --功能：作废一张门诊收费单据
                '  --参数：
                '  --  序号_In：格式: 序号1:数量:执行状态,序号2:数量2:执行状态,...
                '  --  结算作废_IN：是否结算作废调用
                strSql = "Zl_门诊收费记录_销帐_S("
                '  No_In         门诊费用记录.No%Type,
                strSql = strSql & "'" & cllBillList("单据号") & "',"
                '  操作员编号_In 门诊费用记录.操作员编号%Type,
                strSql = strSql & "'" & cllBillList("操作员编号") & "',"
                '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
                strSql = strSql & "'" & cllBillList("操作员姓名") & "',"
                '  序号_In       Varchar2,
                strSql = strSql & "'" & cll序号(i) & "',"
                '  退费时间_In   门诊费用记录.登记时间%Type := Null,
                strSql = strSql & "To_Date('" & Format(cllBillList("登记时间"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')" & ","
                '  退费摘要_In   门诊费用记录.摘要%Type := Null,
                strSql = strSql & "'" & cllBillList("摘要") & "',"
                '  结帐id_In     病人预交记录.结帐id%Type := Null,
                strSql = strSql & "" & cllBillList("结帐ID") & ","
                '  结算作废_In Number:=0
                strSql = strSql & "" & cllBillList("结算作废") & ")"
                
                zlAddArray cllPro_out, strSql
            Next
        End If
            
    Else
        For i = 1 To cllBillDetails.Count
            Set cllBillDetail = cllBillDetails(i)
            str序号 = str序号 & "," & cllBillDetail("序号") & ":" & cllBillDetail("销帐数量") & ":" & cllBillDetail("执行状态")
            
            If Len(str序号) > 3950 Then
                cll序号.Add Mid(str序号, 2)
                str序号 = ""
            End If
        Next
        If str序号 <> "" Then cll序号.Add Mid(str序号, 2)
            
        For i = 1 To cll序号.Count
            'zl_门诊记帐记录_DELETE_S
            '  --功能：冲销一张门诊记帐单据中指定序号行
            '  --参数：
            '  --  序号_In：格式: 序号1:数量:执行状态1,序号2:数量2:执行状态2,...
            strSql = "zl_门诊记帐记录_DELETE_S("
            '  No_In         门诊费用记录.No%Type,
            strSql = strSql & "'" & cllBillList("单据号") & "',"
            '  序号_In       Varchar2,
            strSql = strSql & "'" & cll序号(i) & "',"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSql = strSql & "'" & cllBillList("操作员编号") & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            strSql = strSql & "'" & cllBillList("操作员姓名") & "',"
            '  登记时间_In   门诊费用记录.登记时间%Type,
            strSql = strSql & " TO_Date('" & Format(cllBillList("登记时间"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
            '  记录性质_In   住院费用记录.记录性质%Type := 2
            strSql = strSql & "" & intBillType & ")"
            zlAddArray cllPro_out, strSql
        Next
    End If
    
    zlGetDelSQL_OutRecipeBill = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlExcute_SaveDrugDelRecipeBill(ByVal strSendDrugFeeIDs As String, _
    ByVal strJosnDelDrugRcp As String, ByVal bln门诊 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废药品处方
    '入参:strSendDrugFeeIds-所涉及的发药费用IDs
    '     strJosnNewDrugRcp-作废处方的Json串
    '   lngModule-模块号
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-15 14:09:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String, blnTrans As Boolean
    Dim objService As clsService
    
    On Error GoTo errHandle
    If GetService(objService) = False Then Exit Function
    
    gcnOracle.BeginTrans: blnTrans = True
        If zlExecuteUpdateSyncSymbol(strSendDrugFeeIDs, 1, IIf(bln门诊, 1, 2), 0, 1) = False Then gcnOracle.RollbackTrans: Exit Function
        If objService.zlDrugSvr_DelRecipeBill(strJosnDelDrugRcp, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, gstrSysName
            Exit Function '未提交的,回退后直接退出
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    zlExcute_SaveDrugDelRecipeBill = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function zlExcute_SaveStuffDelBill(ByVal strSendStuffFeeIDs As String, _
    ByVal strJosnDelStuffRcp As String, ByVal bln门诊 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废卫生材料处方
    '入参:strSendDrugFeeIds-所涉及的发药费用IDs
    '     strJosnDelStuffRcp-作废处方的Json串
    '   lngModule-模块号
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-15 14:09:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String, blnTrans As Boolean
    Dim objService As clsService
    
    On Error GoTo errHandle
    If GetService(objService) = False Then Exit Function
    
    gcnOracle.BeginTrans: blnTrans = True
        If zlExecuteUpdateSyncSymbol(strSendStuffFeeIDs, 1, IIf(bln门诊, 1, 2), 0, 1) = False Then gcnOracle.RollbackTrans: Exit Function
        If objService.zlStuffSvr_DelBill(strJosnDelStuffRcp, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, gstrSysName
             Exit Function
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    zlExcute_SaveStuffDelBill = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlFeeDelRecipeBill_Check(ByVal cllFeeCheck As Collection, ByVal intBillType As Integer, Optional ByVal bln划价 As Boolean, _
    Optional ByRef cllBillLists_Out As Collection, Optional ByVal bln门诊 As Boolean, _
    Optional ByVal blnNotShowErr As Boolean, Optional ByRef strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废门诊划价、收费、记账、住院记帐审核合法性检查
    '入参:
    '     intBillType-1-收费单;2-记帐单;3-自动记帐单
    '     bln划价-是否划价
    '出参:
    '   cllBillLists_Out(结构):检查返回的销帐数据
    '    |-cllBillList(Collect):单据基础信息=cllBillLists(_No)
    '        |-成员(cllBillDetails,单据号)
    '        |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '            |-cllBillDetail(Collect):单据信息集=cllBillDetails(_序号)
    '                |-成员(序号,销帐数量,执行状态)
    '   执行状态：0-未执行;1-已执行;2-部分执行
    '返回:合法，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim objServiceCall As Object
    Dim cllData_out As Collection
    Dim cllBillList As Collection, cllTemp As Collection
    Dim cllBillDetails As Collection, cllBillDetail As Collection
    Dim varValue As Variant
    
    On Error GoTo errHandle
    If cllFeeCheck.Count = 0 Then
        strErrmsg_Out = "未传入本次销帐的单据信息，请检查!"
        If Not blnNotShowErr Then MsgBox strErrmsg_Out, vbInformation, gstrSysName
        Exit Function
    End If
    
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    Set cllBillLists_Out = New Collection
    
    For j = 1 To cllFeeCheck.Count
        Select Case intBillType
        Case 1 '收费单
            If bln划价 Then
                If objServiceCall.CallService("Zl_门诊划价记录_Delete_Check", cllFeeCheck(j)(1), , "", mlngModule, Not blnNotShowErr) = False Then Exit Function
            Else
                If objServiceCall.CallService("Zl_门诊收费记录_销帐_Check", cllFeeCheck(j)(1), , "", mlngModule, Not blnNotShowErr) = False Then Exit Function
            End If
        Case 2, 3 '记账单
            If Not bln门诊 Then
                If objServiceCall.CallService("Zl_住院记帐记录_Delete_Check", cllFeeCheck(j)(1), , "", mlngModule, Not blnNotShowErr) = False Then Exit Function
            Else
                If objServiceCall.CallService("Zl_门诊记帐记录_Delete_Check", cllFeeCheck(j)(1), , "", mlngModule, Not blnNotShowErr) = False Then Exit Function
            End If
        Case Else   '其他
            
        End Select
    
        '  --出参: Json_Out,格式如下
        '  --  output
        '  --      code                    N   1   应答码：0-失败；1-成功
        '  --      message                 C   1   应答消息：失败时返回具体的错误信息
        '  --      item_list[]             单据数据列表
        '  --          serial_num          N   1   序号
        '  --          quantity            N   1   销帐数量
        '  --          execute_tag         N   1   执行状态：0-未执行;1-已执行;2-部分执行
        
        varValue = objServiceCall.GetJsonNodeValue("output.code")
        If varValue <> 1 Then
            strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
            If Not blnNotShowErr Then MsgBox strErrmsg_Out, vbInformation, gstrSysName
            Exit Function
        End If
        
        Set cllData_out = objServiceCall.GetJsonListValue("output.item_list")
               
        Set cllBillDetails = New Collection
        If Not cllData_out Is Nothing Then
            For i = 1 To cllData_out.Count
                Set cllTemp = cllData_out(i)
                Set cllBillDetail = New Collection
                cllBillDetail.Add cllTemp("_serial_num"), "序号"
                cllBillDetail.Add cllTemp("_quantity"), "销帐数量"
                cllBillDetail.Add cllTemp("_execute_tag"), "执行状态"
                
                cllBillDetails.Add cllBillDetail, "_" & cllTemp("_serial_num")
            Next
        End If
        
        Set cllBillList = New Collection
        cllBillList.Add cllFeeCheck(j)(0), "单据号"
        cllBillList.Add cllBillDetails, "_cllBillDetails"
        
        cllBillLists_Out.Add cllBillList, "_" & cllFeeCheck(j)(0)
    Next
    
    zlFeeDelRecipeBill_Check = True
    Exit Function
errHandle:
    If blnNotShowErr Then strErrmsg_Out = Err.Description: Exit Function
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRefuseSendList(ByVal lng病人ID As Long, ByVal str主页IDs As String, ByRef str费用IDs_out As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的拒发药清单
    '入参:lng病人ID-病人ID
    '     str主页Ids-多个用逗号
    '出参:str费用IDs_out-返回拒发药的卫生材料及药品有的费用ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService
    Dim str费用IDs As String
        
    On Error GoTo errHandle
    If lngModule <> 0 Then mlngModule = lngModule
    If GetService(objService) = False Then Exit Function
    '先读取药品
    If objService.zlDrugSvr_GetRefuseSendList(lng病人ID, str主页IDs, str费用IDs, mlngModule) = False Then Exit Function
    str费用IDs_out = str费用IDs
    
    '再取卫生材料
    str费用IDs = ""
    If objService.zlStuffSvr_GetRefuseSendList(lng病人ID, str主页IDs, str费用IDs, mlngModule) = False Then Exit Function
    If str费用IDs_out <> "" Then
        str费用IDs_out = str费用IDs_out & "," & str费用IDs
    Else
        str费用IDs_out = str费用IDs
    End If
    
    zlGetRefuseSendList = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter()() = 1 Then
        Resume
    End If
End Function

Public Function zlGetWaitExcuteDrugAndStuff(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef strNotSendInfor_out As String, _
    Optional ByVal int检查离院带药 As Integer = 0, Optional int婴儿序号 As Integer = -1, Optional int门诊标志 As Integer = 1, _
    Optional ByVal strNOs As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人在药房是否还有未发药的药品或卫材
    '入参:lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     lng检查离院带药-离院带药
    '     int门诊标志-1-门诊;2-住院
    '     lngModule -调用模块号
    '出参:
    '    strNotSendInfor_out-返回未发药或未发料信息
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService
    Dim strDrugInfor As String, strStuffInfor As String
        
    On Error GoTo errHandle
    If lngModule <> 0 Then mlngModule = lngModule
    If GetService(objService) = False Then Exit Function
    '1.先检查药品
    If objService.zlDrugSvr_PatiCheckExecute(lng病人ID, lng主页ID, strDrugInfor, int检查离院带药, int婴儿序号, int门诊标志, strNOs, mlngModule) = False Then Exit Function

    '2.再检查卫生材料
    If objService.zlStuffsvr_PatiCheckExecute(lng病人ID, lng主页ID, strStuffInfor, int检查离院带药, int婴儿序号, int门诊标志, strNOs, mlngModule) = False Then Exit Function
 
    If strDrugInfor <> "" And strStuffInfor <> "" Then
        strNotSendInfor_out = strDrugInfor & vbCrLf & vbCrLf & strStuffInfor
    ElseIf strDrugInfor <> "" Then
        strNotSendInfor_out = strDrugInfor
    Else
        strNotSendInfor_out = strStuffInfor
    End If
    zlGetWaitExcuteDrugAndStuff = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetDrugStuff_ExecutedNum(ByVal strNOs As String, ByVal bytBillType As Byte, _
    ByRef rsSended As ADODB.Recordset, ByVal blnExistDrug As Boolean, ByVal blnExistStuff As Boolean, _
    Optional blnAppend As Boolean, Optional ByVal str费用IDs As String, _
    Optional ByVal blnNotShowErr As Boolean, Optional ByRef strErrmsg_Out As String, _
    Optional ByVal lngModule As Long) As Boolean
    '获取单据药品卫材费用已执行数量
    '入参:
    '   strNos:单据号，多个用英文逗号分隔
    '   bytBillType:1-收费单,2-记帐单,3-自动记帐单
    '   blnExistDrug:是否含药品
    '   blnExistStuff:是否含卫材
    '   blnAppend-是否追加记录
    '出参:
    '   rsSended:处方卫材数据：处方单号,药品ID,处方明细ID,已发数量,商品条码,内部条码
    '说明:
    Dim objService As clsService
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetService(objService) = False Then Exit Function
    
    If rsSended Is Nothing Or blnAppend = False Then
        Set rsSended = New ADODB.Recordset
        With rsSended.Fields
            .Append "处方单号", adVarChar, 50, adFldIsNullable
            .Append "处方明细ID", adBigInt, 18, adFldIsNullable
            .Append "药品ID", adBigInt, 18, adFldIsNullable
            .Append "已发数量", adDouble, , adFldIsNullable
            .Append "商品条码", adVarChar, 50, adFldIsNullable
            .Append "内部条码", adVarChar, 50, adFldIsNullable
        End With
        rsSended.CursorLocation = adUseClient
        rsSended.LockType = adLockOptimistic
        rsSended.CursorType = adOpenStatic
        rsSended.Open
    End If
    If bytBillType = 3 Then zlGetDrugStuff_ExecutedNum = True: Exit Function
    
    '药品
    If blnExistDrug Then
        If objService.zlDrugSvr_GetExecutedNum(strNOs, bytBillType, rsData, True, str费用IDs, _
            mlngModule, blnNotShowErr, strErrmsg_Out) = False Then Exit Function
        If Not rsData Is Nothing Then
            Do While Not rsData.EOF
                rsSended.AddNew
                    rsSended!处方单号 = Nvl(rsData!处方单号)
                    rsSended!处方明细id = Nvl(rsData!处方明细id)
                    rsSended!药品ID = Nvl(rsData!药品ID)
                    rsSended!已发数量 = Nvl(rsData!已发数量)
                rsSended.Update
                rsData.MoveNext
            Loop
        End If
    End If

    '卫材
    If blnExistStuff Then
        If objService.zlStuffSvr_GetExecutedNum(strNOs, bytBillType, rsData, True, str费用IDs, _
            mlngModule, blnNotShowErr, strErrmsg_Out) = False Then Exit Function
        If Not rsData Is Nothing Then
            Do While Not rsData.EOF
                rsSended.AddNew
                    rsSended!处方单号 = Nvl(rsData!处方单号)
                    rsSended!处方明细id = Nvl(rsData!处方明细id)
                    rsSended!药品ID = Nvl(rsData!药品ID)
                    rsSended!已发数量 = Nvl(rsData!已发数量)
                    rsSended!商品条码 = Nvl(rsData!商品条码)
                    rsSended!内部条码 = Nvl(rsData!内部条码)
                rsSended.Update
                rsData.MoveNext
            Loop
        End If
    End If
    If rsSended.RecordCount > 0 Then rsSended.MoveFirst

    zlGetDrugStuff_ExecutedNum = True
    Exit Function
ErrHandler:
    If blnNotShowErr Then strErrmsg_Out = Err.Description: Exit Function
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function NewRecipeBill_Check_GetJsonString(ByVal cllCheckDatas As Collection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取生成药品处方检查条件的Json入参串
    '入参:cllCheckDatas-检查条件集,每个元素为Array(元素值，节点值,类型),_节点值)
    '返回:Json格式串
    '编制:刘兴洪
    '日期:2019-08-14 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long, j As Long
    Dim strJsonTemp As String, strJson As String
    If cllCheckDatas Is Nothing Then Exit Function
    
    For i = 1 To cllCheckDatas.Count
        Set cllTemp = cllCheckDatas(i)
        strJsonTemp = ""
        For j = 1 To cllTemp.Count
            
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString(cllTemp(j)(1), cllTemp(j)(0), IIf(cllTemp(j)(2) = "N", Json_num, Json_Text))
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJson = strJson & ",{" & strJsonTemp & "}"
        End If
    Next
    If strJson <> "" Then
        strJson = Mid(strJson, 2)
        strJson = GetNodeString("item_list") & ":[" & strJson & "]"
        strJson = "{" & GetNodeString("input") & ":{" & strJson & "}" & "}"
    End If
    NewRecipeBill_Check_GetJsonString = strJson
End Function

Private Function GetRecipBillListsToJsonString(ByVal cllPatiInfor As Collection, ByVal cllRcpBillLists As Collection, ByRef strJson_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取处方明细给Json字符串
    '入参:cllPatiInfor-病人信息集 （billtype,pati_source等)
    '     cllRcpBillLists-处方明细集,以下每个元素中都包含一个数组Array(元素值,元素节点,类型(N,C)) ,key(_元素节点)
    '       |-BillList
    '           |-BillDetails
    '               |-BillDetail(collect) 单据明细信息集
    '                   |-明细数据:rcpdtl_id,serial_num等
    '出参:strJson_out-返回Json串
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-12 21:55:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBillList As Collection, cllBillDetails As Collection
    Dim i As Long, j As Long, strBillJson As String
    Dim strJson As String, strJsonTemp As String, varData As Variant

    Err = 0: On Error GoTo errHandle:
    strJson_out = ""
    For i = 1 To cllRcpBillLists.Count
        Set cllBillList = cllRcpBillLists(i)(0)
        strBillJson = ""
        For j = 1 To cllBillList.Count
            varData = cllBillList(j)
            If varData(1) <> "item_list" Then
                If IsEmpty(varData(0)) Then
                    strBillJson = strBillJson & "," & GetJsonNodeString(varData(1), "", IIf(varData(2) = "C", 0, 1))
                Else
                    strBillJson = strBillJson & "," & GetJsonNodeString(varData(1), varData(0), IIf(varData(2) = "C", 0, 1))
                End If
            Else
                Set cllBillDetails = varData(0)
                If GetRecipeBillDetailsToJson(cllBillDetails, strJsonTemp) = False Then Exit Function
                strBillJson = strBillJson & "," & strJsonTemp
            End If
        Next
        strJson = strJson & ",{" & Mid(strBillJson, 2) & "}"
    Next
    strJson = GetNodeString("bill_list") & ":[" & Mid(strJson, 2) & "]"
    
    '加载病人信息
    strJsonTemp = ""
    For i = 1 To cllPatiInfor.Count
        varData = cllPatiInfor(i)
        If IsEmpty(varData(0)) Then
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString(varData(1), "", IIf(varData(2) = "C", 0, 1))
        Else
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString(varData(1), varData(0), IIf(varData(2) = "C", 0, 1))
        End If
    Next
    If strJsonTemp <> "" Then strJsonTemp = Mid(strJsonTemp, 2)
    
    strJson_out = "{" & GetNodeString("input") & ":{" & strJsonTemp & "," & strJson & "}}"
    
    GetRecipBillListsToJsonString = True
        
'    'deubJson
'    Dim objFile As New FileSystemObject
'    Dim objText As TextStream
'
'    Call objFile.CreateTextFile(App.Path & "\temp.sql", True)
'
'    Set objText = objFile.OpenTextFile(App.Path & "\temp.sql", ForAppending)
'    objText.WriteLine strJson_out
'    objText.Close
    
   Exit Function
errHandle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetRecipeBillDetailsToJson(ByVal cllBillDetails As Collection, ByRef strJsonOut As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据明细集数据
    '入参:cllBillDetails(collect)-单据明细集
    '       |-BillDetail(collect) 单据明细信息集
    '           |-明细数据:rcpdtl_id,serial_num等
    '出参:strJson_out-返回Json串
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-12 22:10:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim cllBillDetail As Collection, varData As Variant
    Dim strJson As String, strJsonTemp As String
    
    strJson = ""
    Err = 0: On Error GoTo errHandle:
    For i = 1 To cllBillDetails.Count
        Set cllBillDetail = cllBillDetails(i)
        strJsonTemp = ""
        For j = 1 To cllBillDetail.Count
            varData = cllBillDetail(j)
            If IsEmpty(varData(0)) Then
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString(varData(1), "", IIf(varData(2) = "C", 0, 1))
            Else
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString(varData(1), varData(0), IIf(varData(2) = "C", 0, 1))
            End If
        Next
        strJson = strJson & ",{" & Mid(strJsonTemp, 2) & "}"
    Next
    strJsonOut = GetNodeString("item_list") & ":[" & Mid(strJson, 2) & "]"
    
    GetRecipeBillDetailsToJson = True
    Exit Function
errHandle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function DrugNewRecipeBill_Check_GetJsonString(ByVal rsTotal As ADODB.Recordset) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取生成药品处方检查条件的Json入参串
    '入参:
    '   rsTotal-当前的汇总记录集(收费细目ID,库房ID,数量,单价)
    '出参:
    '返回:返回Json串
    '编制:刘兴洪
    '日期:2019-08-14 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    
    If rsTotal Is Nothing Then Exit Function
    If rsTotal.State <> 1 Then Exit Function
    strJson = ""
    With rsTotal
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strJson = strJson & ",{" & GetJsonNodeString("drug_id", Val(Nvl(!收费细目ID)), 1)
            strJson = strJson & "," & GetJsonNodeString("pharmacy_id", Val(Nvl(!库房ID)), 1)
            strJson = strJson & "," & GetJsonNodeString("send_num", Val(Nvl(!数量)), 1)
            strJson = strJson & "," & GetJsonNodeString("price", Val(Nvl(!单价)), 1)
            strJson = strJson & "}"
            .MoveNext
        Loop
    End With
    If strJson <> "" Then
        strJson = Mid(strJson, 2)
        strJson = GetNodeString("fee_list") & ":[" & strJson & "]"
        strJson = "{" & GetNodeString("input") & ":{" & strJson & "}" & "}"
    End If
    DrugNewRecipeBill_Check_GetJsonString = strJson
End Function

Private Function StuffNewBill_Check_GetJsonString(ByVal rsTotal As ADODB.Recordset) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取生成药品处方检查条件的Json入参串
    '入参:
    '   rsTotal-当前的汇总记录集(收费细目ID,库房ID,数量,是否备货卫材,批次,单价)
    '出参:
    '返回:返回Json串
    '编制:刘兴洪
    '日期:2019-08-14 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    If rsTotal Is Nothing Then Exit Function
    If rsTotal.State <> 1 Then Exit Function
    strJson = ""
    With rsTotal
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strJson = strJson & ",{" & GetJsonNodeString("stuff_id", Val(Nvl(!收费细目ID)), 1)
            strJson = strJson & "," & GetJsonNodeString("warehouse_id", Val(Nvl(!库房ID)), 1)
            strJson = strJson & "," & GetJsonNodeString("send_num", Val(Nvl(!数量)), 1)
            strJson = strJson & "," & GetJsonNodeString("is_bakstuff", Val(Nvl(!是否备货卫材)), 1)
            strJson = strJson & "," & GetJsonNodeString("bakstuff_batch", IIf(Val(Nvl(!批次)) > 0, Val(Nvl(!批次)), Empty), 1)
            strJson = strJson & "," & GetJsonNodeString("price", Val(Nvl(!单价)), 1)
            strJson = strJson & "}"
            .MoveNext
        Loop
    End With
    If strJson <> "" Then
        strJson = Mid(strJson, 2)
        strJson = GetNodeString("fee_list") & ":[" & strJson & "]"
        strJson = "{" & GetNodeString("input") & ":{" & strJson & "}" & "}"
    End If
    StuffNewBill_Check_GetJsonString = strJson
End Function

Private Function DelRecipeBill_Check_GetJsonString(ByVal cllBakCheck As Collection, Optional blnBuildHead As Boolean = True) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取作废单据检查条件的Json入参串
    '入参:cllBakCheck-检查条件集,每个元素为Array(元素值，节点值,类型),_节点值)
    '     blnBuildHead-是否构健头
    '返回:Json格式串
    '编制:刘兴洪
    '日期:2019-08-14 11:20:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllItems As Collection, cllItem As Collection, i As Long, j As Long, m As Long
    Dim varTemp As Variant, strTemp, strJsonTemp As String, strJson As String
    If cllBakCheck Is Nothing Then Exit Function
    
    For i = 1 To cllBakCheck.Count
        strJsonTemp = "": strTemp = ""
        varTemp = cllBakCheck(i)
        If varTemp(2) = "Collection" Then
            Set cllItems = varTemp(0)
            strJsonTemp = "": strTemp = ""
            For m = 1 To cllItems.Count
                strTemp = ""
                Set cllItem = cllItems(m)
                For j = 1 To cllItem.Count
                    strTemp = strTemp & "," & GetJsonNodeString(cllItem(j)(1), cllItem(j)(0), IIf(cllItem(j)(2) = "N", Json_num, Json_Text))
                Next
                If strTemp <> "" Then
                    strJsonTemp = strJsonTemp & ",{" & Mid(strTemp, 2) & "}"
                End If
            Next
            If strJsonTemp <> "" Then
                strJsonTemp = GetNodeString(varTemp(1)) & ":[" & Mid(strJsonTemp, 2) & "]"
                strJson = strJson & "," & strJsonTemp
            End If
        Else
            strTemp = GetJsonNodeString(varTemp(1), varTemp(0), IIf(varTemp(2) = "N", Json_num, Json_Text))
            strJson = strJson & "," & strTemp
        End If
    Next
    If strJson <> "" Then
        strJson = Mid(strJson, 2)
        If blnBuildHead Then
            strJson = "{" & GetNodeString("input") & ":{" & strJson & "}" & "}"
        End If
    End If
    
    DelRecipeBill_Check_GetJsonString = strJson
End Function

Public Function Zl_医嘱发送_计费状态_Get(ByVal bytBillType As Byte, ByVal strNo As String, _
    ByVal bln门诊 As Boolean, ByVal byt操作类型 As Byte, ByVal str医嘱IDs As String, _
    ByVal cllDeleteList As Collection, ByVal cllAdviceList As Collection, _
    ByRef cllAdviceList_out As Collection, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:计算医嘱数据的计费状态
    '入参:
    '   bytBillType 单据类型: 1-收费单;2-记帐单;3-自动记帐单
    '   byt操作类型 操作类型：0-删除划价单;1-收费或记帐;2-退费或销帐
    '   cllDeleteList 集合,销账数据,操作类型为0或2时必须传
    '       |元素集合：序号,销帐数量(为零时，按序号直接销帐)
    '   cllAdviceList 集合,医嘱数据
    '       |元素集合：医嘱ID,相关ID,诊疗类别
    '出参:
    '   cllAdviceList 集合,含计费状态的医嘱数据
    '       |元素集合：医嘱ID,计费状态,删除附费
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strTemp As String, strJsonList As String
    Dim objServiceCall As Object
    Dim cllDatas As Collection, clldata As Collection, cllTemp As Collection
    Dim i As Integer
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_医嘱发送_计费状态_Get
    '  --功能：计算医嘱数据的计费状态
    '  --入参：Json_In:格式
    '  --input
    '  --  fee_source                N   1   费用来源：0-门诊;1-住院
    '  --  billtype                  N   1   单据类型: 1-收费单;2-记帐单
    '  --  fee_no                    N   1   费用单据号
    '  --  advice_ids                C   1   医嘱IDs,格式：医嘱ID1,医嘱ID2,...
    '  --  operationtype             N   1   操作类型：0-删除划价单;1-收费或记帐;2-退费或销帐
    '  --  delete_list[]             销账列表,操作类型为0或2时必须传
    '  --    serial_num              N   1   序号
    '  --    quantity                N   1   销帐数量(为零时，按序号直接销帐)
    '  --  advice_list[]
    '  --    advice_id               N   1   医嘱id
    '  --    advice_related_id       N   1   相关id
    '  --    clinic_Type             N   1   诊疗类别
    '  --出参: Json_Out,格式如下
    '  --output
    '  --  code                      N   1   应答吗：0-失败；1-成功
    '  --  message                   C   1   应答消息：失败时返回具体的错误信息
    '  --  item_list
    '  --    advice_id               N   1   医嘱序号
    '  --    update_charge_statu     N   1   更新的计费状态:-1-无须计费(通常无执行和院外执行的都无须计费);0-未计费;1-已计费(或记帐)，2-部分收费/退费(记帐/销帐)，3-全部收费(仅门诊收费有此项)，4-全部退费(销帐)
    '  --    advice_del_annex        N   0   是否删除对应的医嘱附费数据:1-删除;0-不删除
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_source", IIf(bln门诊, 0, 1), Json_num)
    strJson = strJson & "," & GetJsonNodeString("billtype", bytBillType, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("advice_ids", str医嘱IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operationtype", byt操作类型, Json_num)
    '销账数据
    If Not cllDeleteList Is Nothing Then
        If cllDeleteList.Count <> 0 Then
            strJsonList = ""
            For i = 1 To cllDeleteList.Count
                Set clldata = cllDeleteList(i)
                strTemp = ""
                strTemp = strTemp & "" & GetJsonNodeString("serial_num", clldata("序号"), Json_num)
                strTemp = strTemp & "," & GetJsonNodeString("quantity", clldata("销帐数量"), Json_num)
                strJsonList = strJsonList & ",{" & strTemp & "}"
            Next
            strJson = strJson & ",""delete_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    '医嘱数据
    strJsonList = ""
    If Not cllAdviceList Is Nothing Then
        For i = 1 To cllAdviceList.Count
            Set clldata = cllAdviceList(i)
            strTemp = ""
            strTemp = strTemp & "" & GetJsonNodeString("advice_id", clldata("医嘱ID"), Json_num)
            strTemp = strTemp & "," & GetJsonNodeString("advice_related_id", clldata("相关ID"), Json_num)
            strTemp = strTemp & "," & GetJsonNodeString("clinic_Type", clldata("诊疗类别"), Json_Text)
            strJsonList = strJsonList & ",{" & strTemp & "}"
        Next
    End If
    strJson = strJson & ",""advice_list"":[" & Mid(strJsonList, 2) & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_医嘱发送_计费状态_Get", strJson, , "", mlngModule) = False Then Exit Function
    Set cllDatas = objServiceCall.GetJsonListValue("output.item_list")
    
    Set cllAdviceList_out = New Collection
    For i = 1 To cllDatas.Count
        Set clldata = cllDatas(i)
        Set cllTemp = New Collection
        cllTemp.Add clldata("_advice_id"), "医嘱ID"
        cllTemp.Add clldata("_update_charge_statu"), "计费状态"
        cllTemp.Add clldata("_advice_del_annex"), "删除附费"
        cllAdviceList_out.Add cllTemp
    Next
    
    Zl_医嘱发送_计费状态_Get = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function zlExseSvr_GetNewBlackLists(ByVal lng病人ID As Long, ByVal strLastTime As String, ByVal strRegNos As String, _
                ByRef cllBlackDatas As Collection, Optional ByVal strFrmCaption As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取未记录的失约记录
    ' 入参 : strLastTime-最后一次记录的时间
    '        strRegNos-已经记录过的预约单据
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/29 15:15
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    On Error GoTo errHand
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Exsesvr_Getsavebadrecords
    '  --功能:根据病人ID,获取需要保存的黑名单数据
    '  --入参：Json_In:格式
    '  --    input
    '  --      pati_id N 1 病人id
    '  --      last_rgsappt_time C 1 进入不良记录的最后一次时间：yyyy-mm-dd hh24:mi:ss
    '  --      operator_name C 1 操作员姓名
    '  --      badRecord_regnos  C 1 进入黑名单的挂号单号,多个用逗号分离
    '  --
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code              N   1   应答码：0-失败；1-成功
    '  --    message           C   1   应答消息：失败时返回具体的错误信息
    '  --    badrec_list C   保存的不良记录列表
    '  --      pati_id N 1 病人id
    '  --      behavior_category C 1 行为类别:如预约挂号
    '  --      happen_time C 1 发生时间:yyyy-mm-dd hh24:mi:ss
    '  --      add_time  C 1 加入时间:yyyy-mm-dd hh24:mi:ss
    '  --      add_note  C 1 加入原因：如预约超期
    '  --      add_memo  C 1 加入说明
    '  --      additional_info C 1 附加信息
    '  --      creator C 1 登记人
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("last_rgsappt_time", strLastTime, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("blackLst_regnos", strRegNos, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_ExseSvr_GetNewBlackLists", strJson, , strFrmCaption, mlngModule) = False Then Exit Function
    Set cllBlackDatas = objServiceCall.GetJsonListValue("output.badrec_list")
    
    zlExseSvr_GetNewBlackLists = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlRegistCheck(ByVal byt排班模式 As Byte, ByVal byt操作类型 As Byte, ByVal int结算模式 As Integer, ByVal lng病人ID As Long, ByVal str费别 As String, _
                ByVal strNo As String, ByVal str号别 As String, ByVal lng执行部门ID As Long, ByVal str发生时间 As String, _
                ByVal lngSN As Long, ByVal bln预约挂号 As Boolean, Optional ByVal bln急诊 As Boolean, Optional ByVal bln记账 As Boolean, _
                Optional ByRef lng记录ID As Long, Optional ByVal lng预约顺序号 As Long, Optional ByVal lng挂号项目ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:挂号检查数据合法性
    '入参:
    '   strNo:单据号
    '出参:
    '返回:检查通过返回True,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnContOrder As Boolean
    Dim int原结算模式 As Integer, int有效天数 As Integer
    Dim strJson As String, strJsonOut As String
    Dim objServiceCall As Object, objService As clsService
    Dim arrRtn As Variant
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If GetService(objService) = False Then Exit Function
    
    If byt排班模式 = 1 Then
        arrRtn = gobjDatabase.CallProcedure("Zl_病人挂号记录_出诊_Check_S", "zlRegistCheck", strNo, lng病人ID, str费别, _
            UserInfo.姓名, CDate(str发生时间), str号别, IIf(bln急诊, 1, 0), IIf(lngSN = 0, "", lngSN), IIf(bln预约挂号, 1, 0), byt操作类型, _
            int结算模式, IIf(bln记账, 1, 0), lng执行部门ID, lng挂号项目ID, lng记录ID, ZValStr(lng预约顺序号), Empty)
        int原结算模式 = Val(Nvl(arrRtn))
        
    Else
        arrRtn = gobjDatabase.CallProcedure("Zl_病人挂号记录_Check_S", "zlRegistCheck", strNo, lng病人ID, str费别, _
            UserInfo.姓名, CDate(str发生时间), str号别, IIf(bln急诊, 1, 0), IIf(lngSN = 0, "", lngSN), IIf(bln预约挂号, 1, 0), byt操作类型, _
            int结算模式, IIf(bln记账, 1, 0), Empty, Empty)
        int原结算模式 = Val(Nvl(arrRtn(0)))
        lng记录ID = Val(Nvl(arrRtn(1)))
    End If
    
    If int原结算模式 = 1 And int结算模式 = 0 And lng病人ID <> 0 And Not bln预约挂号 Then
        int有效天数 = Val(Left(gobjDatabase.GetPara(21, glngSys, , "01") & "1", 1))
        If objService.ZlCISSvr_ChkOutPatiExistOrder(lng病人ID, int有效天数, blnContOrder) = False Then Exit Function
        If blnContOrder Then
            MsgBox "当前病人的就诊模式为先诊疗后结算,不允许调整该病人的就诊模式!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    zlRegistCheck = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlRegistReceiveCheck(ByVal str挂号单号 As String, ByVal lng病人ID As Long, ByVal str发生时间 As String, _
                ByVal int结算模式 As Integer, ByVal bln记帐费用 As Boolean, _
                ByVal bln急诊 As Boolean, Optional ByVal bln三方调用 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:挂号检查数据合法性
    '入参:
    '   strNo:单据号
    '出参:
    '返回:检查通过返回True,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnContOrder As Boolean
    Dim int原结算模式 As Integer, int有效天数 As Integer
    Dim strJson As String, strJsonOut As String
    Dim objServiceCall As Object, objService As clsService
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If GetService(objService) = False Then Exit Function
    
    'Zl_预约挂号接收_Check_S
    Dim arrRtn As Variant
    arrRtn = gobjDatabase.CallProcedure("Zl_预约挂号接收_Check_S", "zlRegistReceiveCheck", str挂号单号, lng病人ID, CDate(str发生时间), UserInfo.姓名, int结算模式, IIf(bln记帐费用, 1, 0), IIf(bln急诊, 1, 0), IIf(bln三方调用, 1, 0), Empty)
    int原结算模式 = Val(Nvl(arrRtn))
    
    If int原结算模式 = 1 And int结算模式 = 0 And lng病人ID <> 0 Then
        int有效天数 = Val(Left(gobjDatabase.GetPara(21, glngSys, , "01") & "1", 1))
        If objService.ZlCISSvr_ChkOutPatiExistOrder(lng病人ID, int有效天数, blnContOrder) = False Then Exit Function
        If blnContOrder Then
            MsgBox "当前病人的就诊模式为先诊疗后结算,不允许调整该病人的就诊模式!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    zlRegistReceiveCheck = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function zl_ExseSvr_BulidRegistPrice(ByVal str挂号单 As String, ByVal lng挂号科室ID As Long, _
                ByVal str医生姓名 As String, ByRef str划价单_Out As String, Optional ByVal cllVisitInfo As Collection, _
                Optional ByVal lng病人ID As Long, Optional ByVal str姓名 As String, Optional ByVal str性别 As String, Optional ByVal str年龄 As String, _
                Optional ByVal str出生日期 As String, Optional ByVal str身份证号 As String, _
                Optional blnNotShowErr As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据挂号单构建收费划价单
    '入参:strNO-挂号单号
    '     lng科室ID-传入科室ID
    '     str医生姓名-传入医生姓名(传入医生和科室ID时）将会将划价单的执行人，执行科室及执行状态及执行时间填上)
    '     cllVisitInfo(Collect):成员:Array(节点名,节点值)
    '                               >节点名:接诊科室,急诊标志,回诊标志,执行时间
    '出参:
    '     str划价单_Out-返回当前生成的划价单
    '返回:执行成功或获取保存的SQL数据成功,返回true,检查失败返回Fale
    '编制:刘兴洪
    '日期:2018-01-12 10:53:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strData As String, strServiceName   As String
    Dim i As Long
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, Not blnNotShowErr) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法构建收费划价单!"
        If Not blnNotShowErr Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_BulidRegistPrice
    '    input           根据挂号单构键收费划价单
    '    rgst_no        C   1   挂号单号
    '    rgst_dept_id   N   1   挂号科室ID
    '    rgst_dr        C   1   医生姓名
    '    operator_name  C   1   操作员姓名
    '    site_no        C       站点

    strJson = strJson & "" & GetJsonNodeString("rgst_no", str挂号单, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("rgst_dept_id", lng挂号科室ID, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("rgst_dr", str医生姓名, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_name", UserInfo.姓名, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("site_no", gstrNodeNo, Json_Text)
    
    strJson = strJson & "," & GetJsonNodeString("pait_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_name", str姓名, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_sex", str性别, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_age", str年龄, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("birth_date", str出生日期, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_idcard", str身份证号, Json_Text)
    
    strData = ""
    If Not cllVisitInfo Is Nothing Then
        For i = 1 To cllVisitInfo.Count
            If IsArray(cllVisitInfo(i)) Then
                Select Case cllVisitInfo(i)(0)
                    Case "接诊科室ID"
                        strData = strData & "," & GetJsonNodeString("exe_deptid", cllVisitInfo(i)(1), Json_num, True)
                    Case "接诊诊室"
                        strData = strData & "," & GetJsonNodeString("outp_room_name", cllVisitInfo(i)(1), Json_Text)
                    Case "急诊标志"
                        strData = strData & "," & GetJsonNodeString("emg_sign", cllVisitInfo(i)(1), Json_num)
                    Case "回诊标志"
                        strData = strData & "," & GetJsonNodeString("revisit_sign", cllVisitInfo(i)(1), Json_num)
                    Case "执行时间"
                        strData = strData & "," & GetJsonNodeString("exe_time", cllVisitInfo(i)(1), Json_Text)
                End Select
            End If
        Next
        If strData <> "" Then
            strJson = strJson & ",""rgst_visitinfo"":{" & Mid(strData, 2) & "}"
        End If
    End If
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_BulidRegistPrice"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, blnNotShowErr) = False Then Exit Function

    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取指定条件下的挂号单据，请检查！"
        End If
        Exit Function
    End If
    str划价单_Out = Nvl(objServiceCall.GetJsonNodeValue("output.fee_no"))
    zl_ExseSvr_BulidRegistPrice = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnNotShowErr Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_ExseSvr_CheckNoIsCharge(ByVal strNo As String, ByRef intChargeed_Out As Integer, _
                Optional blnCheckPriceNoIsChargeed As Boolean = True, _
                Optional ByVal lng科室id As Long, Optional ByVal str医生姓名 As String, _
                Optional blnNotShowErr As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定挂号单是否已经收费
    '入参:strNO-挂号单号
    '     lng科室ID-传入科室ID
    '     str医生姓名-传入医生姓名(传入医生和科室ID时）将会将划价单的执行人，执行科室及执行状态及执行时间填上)
    '     cllVisitInfo(Collect):成员:Array(节点名,节点值)
    '                               >节点名:接诊科室,急诊标志,回诊标志,执行时间
    '出参:
    '     intChargeed_Out-划价单状态
    '返回:执行成功或获取保存的SQL数据成功,返回true,检查失败返回Fale
    '编制:刘兴洪
    '日期:2018-01-12 10:53:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strData As String, strServiceName   As String
    Dim i As Long
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, Not blnNotShowErr) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取挂号单信息!"
        If Not blnNotShowErr Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_CheckNoIsCharge
    '    --input
    '    --  fee_no          C  1  单据号
    '    --  checkCharge     N       检查划价单是否收费
    '    --output
    '    --  code          N   1   应答码：0-失败；1-成功
    '    --  message       C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    --  fee_status    C   1   单据状态：-1-未找到对应的挂号单据,0-未收费;1-挂号单已收;2-免挂号模式下，还未产生划价记录; 3-挂号单对应的收费划价单已全收费(存在多张划价单时，必须全收的);4-挂号单对应的划价单存在部分收费)

    strJson = strJson & "" & GetJsonNodeString("fee_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("checkCharge", IIf(blnCheckPriceNoIsChargeed, 1, 0), Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_CheckNoIsCharge"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, Not blnNotShowErr) = False Then Exit Function

    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    intChargeed_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.fee_status")))
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取指定条件下的挂号单据，请检查！"
        End If
        Exit Function
    End If
    zl_ExseSvr_CheckNoIsCharge = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnNotShowErr Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_ExseSvr_DelRegistPrice(ByVal strNo As String, Optional ByVal lng科室id As Long, _
                Optional ByVal str医生姓名 As String, Optional ByVal cllVisitInfo As Collection, _
                Optional blnNotShowErr As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除挂号单对应的划价单
    '入参:strNO-挂号单号
    '出参:
    '返回:执行成功或获取保存的SQL数据成功,返回true,检查失败返回Fale
    '编制:刘兴洪
    '日期:2018-01-12 10:53:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strData As String, strServiceName   As String
    Dim i As Long
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, Not blnNotShowErr) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取挂号单信息!"
        If Not blnNotShowErr Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_CheckNoIsCharge
    '    --input
    '    --  fee_no          C  1  单据号
    '    --  checkCharge     N       检查划价单是否收费
    '    --output
    '    --  code          N   1   应答码：0-失败；1-成功
    '    --  message       C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    --  fee_status    C   1   单据状态：-1-未找到对应的挂号单据,0-未收费;1-挂号单已收;2-免挂号模式下，还未产生划价记录; 3-挂号单对应的收费划价单已全收费(存在多张划价单时，必须全收的);4-挂号单对应的划价单存在部分收费)

    strJson = strJson & "" & GetJsonNodeString("fee_no", strNo, Json_Text)
    If lng科室id <> 0 Then
        strJson = strJson & "," & GetJsonNodeString("rgst_dept_id", lng科室id, Json_num)
    End If
    If str医生姓名 <> "" Then
        strJson = strJson & "," & GetJsonNodeString("rgst_dr", str医生姓名, Json_Text)
    End If
    
    strData = ""
    If Not cllVisitInfo Is Nothing Then
        For i = 1 To cllVisitInfo.Count
            If IsArray(cllVisitInfo(i)) Then
                Select Case cllVisitInfo(i)(0)
                    Case "执行部门ID"
                        strData = strData & "," & GetJsonNodeString("exe_deptid", cllVisitInfo(i)(1), Json_num)
                    Case "执行人"
                        strData = strData & "," & GetJsonNodeString("exetr", cllVisitInfo(i)(1), Json_Text)
                    Case "转诊标志"
                        strData = strData & "," & GetJsonNodeString("referral_sign", cllVisitInfo(i)(1), Json_num)
                    Case "转诊科室ID"
                        strData = strData & "," & GetJsonNodeString("referral_dept_id", cllVisitInfo(i)(1), Json_num)
                    Case "转诊医生"
                        strData = strData & "," & GetJsonNodeString("referral_dr", cllVisitInfo(i)(1), Json_Text)
                End Select
            End If
        Next
        If strData <> "" Then
            strJson = strJson & ",""rgst_visitinfo"":{" & Mid(strData, 2) & "}"
        End If
    End If
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_DelRegistPrice"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, Not blnNotShowErr) = False Then Exit Function

    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取指定条件下的挂号单据，请检查！"
        End If
        Exit Function
    End If
    zl_ExseSvr_DelRegistPrice = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnNotShowErr Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_ExseSvr_CheckDelRegistPrice(ByVal strNo As String, ByRef intChargeed_Out As Integer, _
                Optional ByVal lng科室id As Long, Optional ByVal str医生姓名 As String, _
                Optional blnNotShowErr As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定挂号划价单是否允许删除
    '入参:strNO-挂号单号
    '     lng科室ID-传入科室ID
    '     str医生姓名-传入医生姓名
    '出参:
    '     str划价单_Out-返回当前的划价单
    '     intChargeed_Out-划价单状态
    '返回:执行成功或获取保存的SQL数据成功,返回true,检查失败返回Fale
    '编制:刘兴洪
    '日期:2018-01-12 10:53:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strData As String, strServiceName   As String
    Dim i As Long
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, Not blnNotShowErr) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取挂号单信息!"
        If Not blnNotShowErr Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_CheckNoIsCharge
    '    --input
    '    --  fee_no          C  1  单据号
    '    --  checkCharge     N       检查划价单是否收费
    '    --output
    '    --  code          N   1   应答码：0-失败；1-成功
    '    --  message       C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    --  fee_status    C   1   单据状态：-1-未找到对应的挂号单据,0-未收费;1-挂号单已收;2-免挂号模式下，还未产生划价记录; 3-挂号单对应的收费划价单已全收费(存在多张划价单时，必须全收的);4-挂号单对应的划价单存在部分收费)

    strJson = strJson & "" & GetJsonNodeString("fee_no", strNo, Json_Text)
    If lng科室id <> 0 Then
        strJson = strJson & "," & GetJsonNodeString("rgst_dept_id", lng科室id, Json_num)
    End If
    If str医生姓名 <> "" Then
        strJson = strJson & "," & GetJsonNodeString("rgst_dr", str医生姓名, Json_Text)
    End If
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_CheckDelRegistPrice"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, Not blnNotShowErr) = False Then Exit Function

    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取指定条件下的挂号单据，请检查！"
        End If
        Exit Function
    End If
    intChargeed_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.fee_status")))
    zl_ExseSvr_CheckDelRegistPrice = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnNotShowErr Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlExcute_BindOutPno(ByVal lng病人ID As Long, ByVal str门诊号 As String, ByVal str登记时间 As String, _
                Optional ByVal str医保号 As String, Optional ByVal intInsure As Integer, Optional ByVal str病案类型 As String = "一般", _
                Optional ByVal strFrmCaption As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 更新病人挂号及费用记录的门诊号
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/21 15:24
    '---------------------------------------------------------------------------------------
    Dim blnTrans As Boolean
    Dim strErrMsg As String, strSql As String
    Dim intReturn As Integer
    Dim objServiceCall As Object, objService As clsService
    
    On Error GoTo errHand
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If GetService(objService) = False Then Exit Function
    
    'Zl_门诊费用数据_Update
    '  --功能：调整门诊费用数据的信息
    strSql = "Zl_门诊费用数据_Update_S(" & lng病人ID & "," & str门诊号 & ")"
    
    gcnOracle.BeginTrans: blnTrans = True
        Call gobjDatabase.ExecuteProcedure(strSql, "zlExcute_BindOutPno")
        
        If objService.zlPatiSvr_BindOutPno(lng病人ID, str门诊号, str医保号, intInsure, True, strErrMsg) = False Then
            If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "绑定门诊号失败" & IIf(strErrMsg <> "", "，原因：" & strErrMsg, "！"), vbInformation, gstrSysName
            Exit Function
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '临床域：更新门诊病案，失败了也不影响业务
    Call objService.zlCisSvr_UpdateOutMedRec(lng病人ID, str门诊号, str登记时间, str病案类型)
    
    zlExcute_BindOutPno = True
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlRegistAutoIntoBlacklist(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 挂号时检查违约并加入黑名单
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/29 14:45
    '---------------------------------------------------------------------------------------
    Dim cllBlackDatas As Collection
    Dim objService As clsService
    Dim strLastTime As String, strRegNos As String
    
    On Error GoTo errHand
    
    '获取费用服务接口
    If GetService(objService) = False Then Exit Function
    
    Call objService.ZlPatiSvr_GetblackRegNos(lng病人ID, strLastTime, strRegNos)
    '获取新的失约记录
    If zlExseSvr_GetNewBlackLists(lng病人ID, strLastTime, strRegNos, cllBlackDatas, "获取失约记录") = False Then Exit Function
    If cllBlackDatas Is Nothing Then zlRegistAutoIntoBlacklist = True: Exit Function
    If cllBlackDatas.Count = 0 Then zlRegistAutoIntoBlacklist = True: Exit Function
    
    '更新记录
    If objService.zlPatiSvr_Savebadrecord(cllBlackDatas) = False Then Exit Function
    
    zlRegistAutoIntoBlacklist = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExcute_InBillVerfiy(ByVal strNo As String, ByVal str序号 As String, ByVal dt审核时间 As Date, _
    ByVal lngModule As Long, Optional ByVal strInsure As String, Optional ByRef blnAutoSendDrug As Boolean, _
    Optional ByVal bln是否多病人单 As Boolean, Optional ByVal lng病人ID As Long, _
    Optional ByRef blnStuffSync As Boolean, Optional ByRef blnDrugSync As Boolean, _
    Optional ByRef blnAutoSendStuff As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行住院记帐单审核操作
    '入参:
    '   strNo-执行的单据号
    '   str序号-审核的序号
    '   blnAutoSendDrug-是否已自动发放药品
    '   strInsure-记帐表时，多个医保(险类1,险类2,...)
    '   bln是否多病人单-是否多病人单(记帐表)
    '   lng病人ID-只审核指定病人,用于按病人审核记帐表
    '   blnAutoSendStuff-是否自动发药卫材，True=根据 Zl_住院记帐记录_Verify_Check 返回控制自动发料，false=不自动发料
    '出参:
    '   blnAutoSendDrug-True:已经自动发放药品;false-未自动发放成功药品
    '   blnStuffSync-卫材数据是否已同步
    '   blnDrugSync-药品数据是否已同步
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim clldata As Collection, cllTemp As Collection
    Dim cllStuffData As Collection, cllDrugData As Collection 'Array(单据类型,单据号,费用IDs,发药窗口,是否自动发放,自动发放明细IDs)
    Dim cllSendDrug As Collection, cllSendStuff As Collection '自动发药及发料,Array(费用ids,执行状态,已发料)
    Dim strInfo As String, strSql As String
    Dim i As Long, j As Long, blnTrans As Boolean
    Dim strDrugFeeIDs As String, strStuffFeeIDs As String
    Dim str医嘱IDs As String
    Dim objService As clsService
    Dim rsAdviceData As ADODB.Recordset, cllAdviceList As Collection, cllAdviceList_out As Collection
    Dim cllAdviceDatas As Collection, str禁止发药费用IDs As String, strAutoSendDrug As String
    Dim varIDs As Variant, m As Long
    
    On Error GoTo ErrHandler
    blnStuffSync = False: blnDrugSync = False
    If GetService(objService) = False Then Exit Function
    If strNo = "" Then
        MsgBox "没有住院划价单据,请先输入！", vbInformation, gstrSysName
        Exit Function
    End If
    Set cllAdviceDatas = New Collection
    '1.执行前需要先进行审核检查
    If Charge_Verfiy_IsValied(IIf(bln是否多病人单, 3, 2), strNo, str序号, IIf(bln是否多病人单, 0, lng病人ID), clldata, lngModule) = False Then Exit Function
    
    If blnAutoSendDrug Then '进入输液配药中心的禁止自动发药
        If mobjService.zlPivasSvr_Getinfusion_Record(strNo, str禁止发药费用IDs) = False Then Exit Function
    End If
    
    '先收集药品及卫生材料相关数据
    Set cllStuffData = New Collection: Set cllDrugData = New Collection 'Array(单据类型,单据号,费用IDs,发药窗口,是否自动发放,自动发放明细IDs)
    Set cllSendDrug = New Collection: Set cllSendStuff = New Collection 'Array(费用ids,执行状态,已发料)
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
           'Collect(_rcp_no,_stuff_rcpdtl_ids,_drug_rcpdtl_ids,_autosendstuff_rcpdtl_ids)
           Set cllTemp = clldata(i)
           If cllTemp("_drug_rcpdtl_ids") <> "" Then
                strAutoSendDrug = ""
                If blnAutoSendDrug Then
                    '排除进入输液配置中心的药品
                    If str禁止发药费用IDs = "" Then
                        strAutoSendDrug = cllTemp("_drug_rcpdtl_ids")
                    Else
                        varIDs = Split(cllTemp("_drug_rcpdtl_ids"), ",")
                        For m = 0 To UBound(varIDs)
                            If InStr("," & str禁止发药费用IDs & ",", "," & varIDs(m) & ",") = 0 Then
                                strAutoSendDrug = IIf(strAutoSendDrug = "", "", strAutoSendDrug & ",") & varIDs(m)
                            End If
                        Next
                    End If
                End If
                
                'Array(单据类型,单据号,费用IDs,发药窗口,是否自动发放,自动发放明细IDs)
                cllDrugData.Add Array(2, cllTemp("_rcp_no"), cllTemp("_drug_rcpdtl_ids"), "", IIf(strAutoSendDrug = "", 0, 1), strAutoSendDrug)
                If strAutoSendDrug <> "" Then cllSendDrug.Add Array(strAutoSendDrug, 1, 0)
                
                strDrugFeeIDs = strDrugFeeIDs & "," & cllTemp("_drug_rcpdtl_ids")
           End If
           
           If cllTemp("_stuff_rcpdtl_ids") <> "" Then
               If cllTemp("_autosendstuff_rcpdtl_ids") <> "" And blnAutoSendStuff Then
                    'Array(单据类型,单据号,费用IDs,发药窗口,是否自动发放,自动发放明细IDs)
                    cllStuffData.Add Array(2, cllTemp("_rcp_no"), cllTemp("_stuff_rcpdtl_ids"), 1, cllTemp("_autosendstuff_rcpdtl_ids"))
                    cllSendStuff.Add Array(cllTemp("_autosendstuff_rcpdtl_ids"), 1, 0)
               Else
                    'Array(单据类型,单据号,费用IDs,发药窗口,是否自动发放,自动发放明细IDs)
                    cllStuffData.Add Array(2, cllTemp("_rcp_no"), cllTemp("_stuff_rcpdtl_ids"), 0, "")
               End If
               strStuffFeeIDs = strStuffFeeIDs & "," & cllTemp("_stuff_rcpdtl_ids")
           End If
        Next
    End If
    If strDrugFeeIDs <> "" Then strDrugFeeIDs = Mid(strDrugFeeIDs, 2)
    If strStuffFeeIDs <> "" Then strStuffFeeIDs = Mid(strStuffFeeIDs, 2)
     
    '计算医嘱计费状态
    If CheckExistAdviceFeeByDetail(strNo, Val("2-记帐单"), False, str医嘱IDs, True) = False Then Exit Function
    If str医嘱IDs <> "" Then
        If objService.ZlCissvr_GetAllGroupAdviceIDs(str医嘱IDs, rsAdviceData, mlngModule) = False Then Exit Function
        Set cllAdviceList = New Collection
        Do While Not rsAdviceData.EOF
            Set cllTemp = New Collection
            cllTemp.Add Nvl(rsAdviceData!医嘱id), "医嘱ID"
            cllTemp.Add Nvl(rsAdviceData!相关ID), "相关ID"
            cllTemp.Add Nvl(rsAdviceData!诊疗类别), "诊疗类别"
            cllAdviceList.Add cllTemp
            rsAdviceData.MoveNext
        Loop
        If Zl_医嘱发送_计费状态_Get(Val("2-记帐单"), strNo, False, Val("1-收费或记帐"), _
            str医嘱IDs, Nothing, cllAdviceList, cllAdviceList_out, mlngModule) = False Then Exit Function
        If Not cllAdviceList_out Is Nothing Then
            For j = 1 To cllAdviceList_out.Count
                Set clldata = cllAdviceList_out(j)
                Set cllTemp = New Collection
                cllTemp.Add clldata("医嘱ID"), "医嘱ID"
                cllTemp.Add strNo, "费用单号"
                cllTemp.Add clldata("计费状态"), "计费状态"
                cllTemp.Add 0, "删除附费"
                cllAdviceDatas.Add cllTemp
            Next
        End If
    End If
    
    '2.执行费用数据保存
    'Zl_住院记帐记录_Verify_S
    strSql = "Zl_住院记帐记录_Verify_S("
    '  No_In           住院费用记录.No%Type,
    strSql = strSql & "'" & strNo & "',"
    '  操作员编号_In   住院费用记录.操作员编号%Type,
    strSql = strSql & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In   住院费用记录.操作员姓名%Type,
    strSql = strSql & "'" & UserInfo.姓名 & "',"
    '  序号_In         Varchar2 := Null,--格式如"1,3,5,7,8",为空表示审核所有未审核的行
    strSql = strSql & "'" & str序号 & "',"
    '  病人id_In       住院费用记录.病人id%Type := Null,--只审核指定病人,用于按病人审核记帐表
    strSql = strSql & "" & IIf(lng病人ID = 0, "NULL", lng病人ID) & ","
    '  审核时间_In     住院费用记录.登记时间%Type := Null,
    strSql = strSql & "To_Date('" & Format(dt审核时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    '  更新同步标志_In Number:=1
    
    gcnOracle.BeginTrans: blnTrans = True
        Call gobjDatabase.ExecuteProcedure(strSql, "住院记帐审核")
        
        If str医嘱IDs <> "" Then
            '更改医嘱数据的计费状态
            If cllAdviceDatas.Count > 0 Then
                If objService.ZlCisSvr_UpdAdviceChargeTag(Val("2-记帐单"), cllAdviceDatas, Nothing, mlngModule) = False Then
                    gcnOracle.RollbackTrans: Exit Function
                End If
            End If
        ElseIf strInsure <> "" Then
            '医保记帐审核完成前上传
            If InsureTranChargeDetail(strNo, strInsure, strInfo, False) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                If strInfo <> "" Then MsgBox strInfo, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '3.执行药品处方确认
    blnDrugSync = ExecuteDrugBillAffirm(2, cllDrugData, strDrugFeeIDs, cllSendDrug, dt审核时间)
    blnAutoSendDrug = blnDrugSync And Not cllSendDrug Is Nothing
    blnAutoSendDrug = blnAutoSendDrug And cllSendDrug.Count > 0
    
    '4.执行卫材处方确认
    blnStuffSync = ExecuteStuffBillAffirm(2, cllStuffData, strStuffFeeIDs, cllSendStuff, dt审核时间)
     
    '5.医保记帐审核完成后上传(有医嘱费用时，不分记帐完成前与后)
    If strInsure <> "" Then
        If InsureTranChargeDetail(strNo, strInsure, strInfo, True, False, str医嘱IDs <> "") = False Then
            If strInfo <> "" Then
                MsgBox strInfo, vbInformation, gstrSysName
            Else
                MsgBox "单据""" & strNo & """的数据向医保传送失败，该单据已审核！", vbInformation, gstrSysName
            End If
        End If
    End If
    
    zlExcute_InBillVerfiy = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function Charge_Verfiy_IsValied(ByVal intBillType As Integer, _
    ByVal strNOs As String, ByVal str序号s As String, ByVal lng病人ID As Long, _
    Optional ByRef cllData_out As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:住院记帐审核合法性检查
    '入参:
    '   intBillType-单据类型(1-收费单;2-记帐单;3-记帐表)
    '   strNOs-单据号:允许传入多个，多个时，序号无效
    '   str序号s-多个用逗号，可以不传入
    '   lng病人ID-只审核指定病人,用于按病人审核记帐表
    '   lngModule -调用模块号
    '出参:
    '   cllData_out-返回的数据集，成员:Array(_rcp_no,_stuff_rcpdtl_ids,_drug_rcpdtl_ids,_autosendstuff_rcpdtl_ids)
    '返回:记帐审核合法，返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim str病人信息 As String, rsPatiPageInfo As ADODB.Recordset
    Dim str病人IDs As String, varData As Variant, i As Long
    Dim cllTemp As Collection, cllPatient As Collection
    
    On Error GoTo ErrHandler
    If GetService(objService) = False Then Exit Function
    
    strSql = "Select /*+cardinality(j,10)*/Distinct 病人id, 主页id" & _
            " From 住院费用记录 A, Table(f_Str2list([1])) J" & _
            " Where a.记录性质 = 2 And a.记录状态 = 0 And a.No = j.Column_Value"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取病人ID", strNOs)
    If rsTemp.EOF Then
        MsgBox "未找到病人信息，不能完成审核！", vbInformation, gstrSysName
        Exit Function
    End If
    str病人信息 = "": str病人IDs = ""
    Do While Not rsTemp.EOF
        If lng病人ID = 0 Or lng病人ID = Nvl(rsTemp!病人ID) Then
            str病人信息 = str病人信息 & "," & Nvl(rsTemp!病人ID) & ":" & Nvl(rsTemp!主页ID)
            str病人IDs = str病人IDs & "," & Nvl(rsTemp!病人ID)
        End If
        rsTemp.MoveNext
    Loop
    If str病人信息 <> "" Then str病人信息 = Mid(str病人信息, 2)
    If str病人IDs <> "" Then str病人IDs = Mid(str病人IDs, 2)
    
    '获取病人主页信息
    If objService.ZlCissvr_GetPatiPageInfo(Val("0-基本信息"), str病人信息, rsPatiPageInfo, , , lngModule) = False Then Exit Function
    
    Set cllPatient = New Collection
    varData = Split(str病人IDs, ",")
    For i = 0 To UBound(varData)
        rsPatiPageInfo.Filter = "病人ID=" & varData(i)
        If rsPatiPageInfo.EOF Then
            MsgBox "获取[" & varData(i) & "]的病人信息失败！", vbInformation, gstrSysName
            Exit Function
        End If
        'Collect(病人ID,费用审核标志,住院状态)
        Set cllTemp = New Collection
        cllTemp.Add varData(i), "病人ID"
        cllTemp.Add Nvl(rsPatiPageInfo!审核标志), "费用审核标志"
        cllTemp.Add Nvl(rsPatiPageInfo!住院状态), "住院状态"
        cllPatient.Add cllTemp
    Next
    
    If Zl住院记帐记录_Verify_Check(strNOs, cllPatient, str序号s, cllData_out, lngModule) = False Then Exit Function
    Charge_Verfiy_IsValied = True
    Exit Function
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zl病人挂号记录_批量换号(ByVal cllRgst As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:住院记帐审核合法性检查
    '入参:
    '   intBillType-单据类型(1-收费单;2-记帐单;3-记帐表)
    '   strNOs-单据号:允许传入多个，多个时，序号无效
    '   str序号s-多个用逗号，可以不传入
    '   lng病人ID-只审核指定病人,用于按病人审核记帐表
    '   lngModule -调用模块号
    '出参:
    '   cllData_out-返回的数据集，成员:Array(_rcp_no,_stuff_rcpdtl_ids,_drug_rcpdtl_ids,_autosendstuff_rcpdtl_ids)
    '返回:记帐审核合法，返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService, objServiceCall As Object, cllPatiVisit As Collection
    Dim blnTrans As Boolean
    Dim intReturn As Integer
    Dim strJson As String, strJsonOut As String, strErrMsg As String
    Dim str病人IDs As String, varData As Variant, i As Long
    
    
    On Error GoTo ErrHandler
    If cllRgst Is Nothing Then Exit Function
    If cllRgst.Count = 0 Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If GetService(objService) = False Then Exit Function
    If lngModule <> 0 Then mlngModule = lngModule
    
    strJson = ""
    For i = 1 To cllRgst.Count
        Select Case cllRgst(i)(0)
        Case "挂号单号"
            strJson = strJson & "," & GetJsonNodeString("rgst_nos", cllRgst(i)(1), Json_Text)
        Case "new号别"
            strJson = strJson & "," & GetJsonNodeString("rgst_newcode", cllRgst(i)(1), Json_Text)
        Case "new医生姓名"
            strJson = strJson & "," & GetJsonNodeString("rgst_newdr", cllRgst(i)(1), Json_Text)
        Case "new医生ID"
            strJson = strJson & "," & GetJsonNodeString("rgst_newdr_id", cllRgst(i)(1), Json_num)
        Case "new科室ID"
            strJson = strJson & "," & GetJsonNodeString("rgst_newdept_id", cllRgst(i)(1), Json_num)
        Case "new出诊记录ID"
            strJson = strJson & "," & GetJsonNodeString("rgst_newrec_id", cllRgst(i)(1), Json_num, True)
       Case "old号别"
            strJson = strJson & "," & GetJsonNodeString("rgst_oldcode", cllRgst(i)(1), Json_Text)
        Case "old医生姓名"
            strJson = strJson & "," & GetJsonNodeString("rgst_olddr", cllRgst(i)(1), Json_Text)
        Case "old医生ID"
            strJson = strJson & "," & GetJsonNodeString("rgst_olddr_id", cllRgst(i)(1), Json_num)
        Case "old科室ID"
            strJson = strJson & "," & GetJsonNodeString("rgst_olddept_id", cllRgst(i)(1), Json_num)
        Case "old出诊记录ID"
            strJson = strJson & "," & GetJsonNodeString("rgst_oldrec_id", cllRgst(i)(1), Json_num, True)
        Case "操作员姓名"
            strJson = strJson & "," & GetJsonNodeString("operator_name", cllRgst(i)(1), Json_Text)
        End Select
    Next
    strJson = "{""input"":{" & Mid(strJson, 2) & "}}"
    
    gcnOracle.BeginTrans: blnTrans = True
        Call objServiceCall.CallService("Zl_病人挂号记录_批量换号", strJson, strJsonOut, , lngModule, True, strErrMsg)
        intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
        If intReturn <> 1 Then
            If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
            strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
            MsgBox "批量换号失败" & IIf(strErrMsg <> "", "，原因：" & strErrMsg, "！"), vbInformation, gstrSysName
            Exit Function
        End If
        Set cllPatiVisit = objServiceCall.GetJsonListValue("output.rgst_list")
    gcnOracle.CommitTrans: blnTrans = False
    
    If Not cllPatiVisit Is Nothing Then
        If cllPatiVisit.Count > 0 Then
            Call mobjService.zlPatisvr_BatupdoutPativisit(cllPatiVisit, True)
        End If
    End If
    zl病人挂号记录_批量换号 = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function



Public Function zl_ExseSvr_Actualmoney(ByVal str费别 As String, ByVal lng收费细目id As Long, _
    ByVal lng收入项目id As Long, ByVal dbl应收金额 As Double, ByRef dbl实收金额_Out As Double, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional ByVal dbl数量 As Double, Optional ByVal dbl成本价 As Long, Optional ByVal lng医嘱ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费别计算打折金额
    '入参:str费别- 费别名称
    '     lng收费细目id-细目id
    '     lng收入项目id
    '     blnShowErrMsg-是否显示错误信息
    '     dbl数量\dbl成本价\lng医嘱id:药品及卫材才传入
    '出参:dbl实收金额_Out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的打折信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_Actualmoney
    '    input           获以实收金额
    '        fee_category    C   1   费别
    '        fee_item_id N   1   收费细目id
    '        income_item_id  N   1   收入项目id
    '        fee_amrcvb  N   1   应收金额
    '        quantity    N   1   数量
    '        price_cost  N   1   成本价
    '        order_id    N   1   医嘱id
    strJson = strJson & "" & GetJsonNodeString("fee_category", str费别, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("fee_item_id", lng收费细目id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("income_item_id", lng收入项目id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_amrcvb", dbl应收金额, Json_num)
    strJson = strJson & "," & GetJsonNodeString("quantity", dbl数量, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("price_cost", dbl成本价, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("order_id", lng医嘱ID, Json_num, True)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_Actualmoney"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    fee_category    C   1   费别
    '    net_receipts_fee    N   1   实收金额

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取指定条件下的医嘱数据，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    dbl实收金额_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.fee_ampaib")))
    zl_ExseSvr_Actualmoney = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetShareInvoiceGroupID(ByVal bytKind As Byte) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定票种的共用票据批次
    '编制:刘兴洪
    '日期:2011-04-29 10:24:48
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, cllBillInfor As Collection
    On Error GoTo errH
    If zl_ExseSvr_GetReceiveInvoice(bytKind, "", cllBillInfor, False, "", , , , , , 1, True, rsTemp) = False Then Exit Function
    Set GetShareInvoiceGroupID = rsTemp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "") As Long
'功能：获取张数够用并且指定票据在其可用范围内的领用ID
'参数：bytKind      =   票种
'      intNum       =   要打印的票据张数
'      lngLastUseID =   上次使用的领用ID
'      lngShareUseID=   本地参数指定的共用ID
'      strBill      =   当前票据号，用于检查领用批次的票据范围
'      strUseType-使用类别
'返回：
'      >0   =   成功，可用的领用ID
'      =0   =   失败
'      -1   =   没有自用(用完或不够，或未领用),未设置共用
'      -2   =   没有自用(用完或不够，或未领用),设置的共用已用完或不够
'      -3   =   指定票据号不在当前所有可用领用批次的有效票据号范围内
'      -4   =   指定批次的票据不够用
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    Dim cllBillInfo As Collection, cllTemp As Collection
 
    On Error GoTo errH
    '1.上次的领用批次是否可用并够用
    If lngLastUseID > 0 Then
    
        If zl_ExseSvr_GetReceiveInvoice(bytKind, lngLastUseID, cllBillInfo, False, strUseType, "", intNum, True, , glngModul) = False Then Exit Function
        '        item_list C
        '            recv_id N   1   领用ID
        '            use_mode    N   1   使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
        '            use_type    C   1   票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
        '            prefix_text C   1   前缀文本
        '            start_no    C   1   开始号码
        '            end_no  C   1   终止号码
        '            inv_no_cur  C   1   当前号码
        '            surplus_num C   1   剩余数量
        '            create_time C   1   登记时间:yyyy-mm-dd hh24:mi:ss
        '            use_time    C   1   使用时间:yyyy-mm-dd hh24:mi:ss
        '            recvtr  C   1   领用人
        If Not cllBillInfo Is Nothing Then
            If cllBillInfo.Count <> 0 Then
                 If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '可能没有当前票据号
                
                Set cllTemp = cllBillInfo(1)
                blnTmp = False
                strPre = Nvl(cllTemp("_prefix_text"))
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(Nvl(cllTemp("_start_no"))) And UCase(strBill) <= UCase(Nvl(cllTemp("_end_no"))) And Len(strBill) = Len(Nvl(cllTemp("_start_no")))) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
            ElseIf intNum > 1 Then  '不是确定领用批次调用时,当前票据号所在批次不够用
                GetInvoiceGroupID = -4: Exit Function
            End If
        ElseIf intNum > 1 Then  '不是确定领用批次调用时,当前票据号所在批次不够用
            GetInvoiceGroupID = -4: Exit Function
        End If
    End If
    
    '2.上次的领用批次不够用或不可用时,取自已领的并且自用的
    '  有多批，最近使用的优先,少的先用,先领先用
    If zl_ExseSvr_GetReceiveInvoice(bytKind, 0, cllBillInfo, True, strUseType, UserInfo.姓名, intNum, True, , glngModul) Then
        '        item_list C
        '            recv_id N   1   领用ID
        '            use_mode    N   1   使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
        '            use_type    C   1   票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
        '            prefix_text C   1   前缀文本
        '            start_no    C   1   开始号码
        '            end_no  C   1   终止号码
        '            inv_no_cur  C   1   当前号码
        '            surplus_num C   1   剩余数量
        '            create_time C   1   登记时间:yyyy-mm-dd hh24:mi:ss
        '            use_time    C   1   使用时间:yyyy-mm-dd hh24:mi:ss
        '            recvtr  C   1   领用人
        For i = 1 To cllBillInfo.Count
            Set cllTemp = cllBillInfo(i)
            
            If strBill = "" Then GetInvoiceGroupID = Val(Nvl(cllTemp("_recv_id"))): Exit Function '第一次使用时没有当前票据号
            blnTmp = False
            strPre = Nvl(cllTemp("_prefix_text"))
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(Nvl(cllTemp("_start_no"))) And UCase(strBill) <= UCase(Nvl(cllTemp("_end_no"))) And Len(strBill) = Len(UCase(Nvl(cllTemp("_start_no"))))) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = Val(Nvl(cllTemp("_recv_id"))): Exit Function
        Next
        lngReturn = IIf(cllBillInfo.Count > 0, -3, -1)
    Else
        lngReturn = -1
    End If

    '3.没有自用的,使用本地参数指定的共用批次
    If lngShareUseID > 0 Then
        If zl_ExseSvr_GetReceiveInvoice(bytKind, lngLastUseID, cllBillInfo, False, strUseType, "", intNum, True, , glngModul) = False Then
            GetInvoiceGroupID = lngReturn   '返回未找到的原因代码
            Exit Function
        End If
        
        If cllBillInfo.Count = 0 Then
            lngReturn = -2
            GetInvoiceGroupID = lngReturn
            Exit Function
        End If
        
        Set cllTemp = cllBillInfo(1)
        If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '第一次使用时没有当前票据号
        blnTmp = False
        strPre = Nvl(cllTemp("_prefix_text"))
        If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
            blnTmp = True
        ElseIf Not (UCase(strBill) >= UCase(Nvl(cllTemp("_start_no"))) And UCase(strBill) <= UCase(Nvl(cllTemp("_end_no"))) And Len(strBill) = Len(UCase(Nvl(cllTemp("_start_no"))))) Then
            blnTmp = True
        End If
        If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
        lngReturn = -3
    End If
    GetInvoiceGroupID = lngReturn   '返回未找到的原因代码
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function CheckUsedBill(bytKind As Byte, ByVal lng领用ID As Long, _
    Optional ByVal strBill As String, _
     Optional ByVal strUseType As String = "") As Long
    '功能：检查当前操作员是否有可用票据领用(自用或共用),并返回可用的领用ID
    '参数：bytKind=票种
    '      lng领用ID=第一次检查时为本地设置的共用领用ID,以后为上次使用的领用ID
    '      strBill=要检查范围的票据号
    '说明：
    '    1.在检查范围时,如果病人有多批自用票据,则只要在其中一批之中就行了
    '    2.在检查范围时,长度也在检查范围之内。
    '    3.当有多批自用时,缺省按少的先用,先领先用,"最近使用的优先"原则
    '返回：
    '      正常：票据领用ID>0
    '      0=失败
    '      -1:没有自用(用完或未领用)、也没有共用(未设置)
    '      -2:设置的共用已用完
    '      -3:指定票据号不在当前可用范围内(包含多批自用票据的情况)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSql As String, blnTmp As Boolean, lngReturn As Long
    Dim cllBillInfos As Collection, cllBillItem As Collection
    Dim cllSharess  As Collection, cllSharessItem As Collection
    Dim i As Long
    
    On Error GoTo errH
    
      '2.上次的领用批次不够用或不可用时,取自已领的并且自用的
    '  有多批，最近使用的优先,少的先用,先领先用
    If Not zl_ExseSvr_GetReceiveInvoice(bytKind, "", cllBillInfos, True, strUseType, UserInfo.姓名, 1, True, , glngModul) Then
        Set cllBillInfos = New Collection
    End If
    If cllBillInfos Is Nothing Then Set cllBillInfos = New Collection


    If lng领用ID = 0 Then
         '程序中第一次检查,且没有设置本地共用
         If cllBillInfos.Count = 0 Then CheckUsedBill = -1: Exit Function  '也没有自用票据
         '有自用票据 , 按优先原则返回
         Set cllBillItem = cllBillInfos(1)
         lngReturn = Val(Nvl(cllBillItem("_recv_id")))
    Else
        '上次使用的领用ID或第一次检查的共用ID,先判断性质
        If Not zl_ExseSvr_GetReceiveInvoice(bytKind, lng领用ID, cllSharess, True, strUseType, "", 0, True, , glngModul) Then
            Set cllSharess = New Collection
        End If
        
        If cllSharess.Count = 0 Then CheckUsedBill = -2: Exit Function
        Set cllSharessItem = cllSharess(1)
        
        
        If Val(Nvl(cllSharessItem("_use_mode"))) = 2 Then '共用,要先看有没有自用
            If cllBillInfos.Count <> 0 Then
                '有自用的，优先
                Set cllBillItem = cllBillInfos(1)
                lngReturn = Val(cllBillItem("_recv_id"))
            Else
                '没有自用取共用
                If Val(Nvl(cllSharessItem("_surplus_num"))) = 0 Then CheckUsedBill = -2: Exit Function '共用已经用完
                lngReturn = Val(cllSharessItem("_recv_id"))
                blnTmp = True
            End If
        Else
            '自用票据
            If Val(Nvl(cllSharessItem("_surplus_num"))) > 0 Then
                '有剩余
                lngReturn = Val(cllSharessItem("_recv_id"))
            Else
                '其它有剩余的自用
                If cllBillInfos.Count = 0 Then CheckUsedBill = -1: Exit Function      '其它自用也没有剩余
                Set cllBillItem = cllBillInfos(1)
                lngReturn = Val(cllBillItem("_recv_id"))
            End If
        End If
    End If
    
    '检查票号范围是否正确
    If strBill <> "" Then
        If blnTmp Then
            '在共用范围内范围判断
            If UCase(Left(strBill, Len(Nvl(cllSharessItem("_prefix_text"))))) <> UCase(Nvl(cllSharessItem("_prefix_text"))) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(Nvl(cllSharessItem("_start_no"))) And UCase(strBill) <= UCase(Nvl(cllSharessItem("_end_no"))) And Len(strBill) = Len(UCase(Nvl(cllSharessItem("_start_no"))))) Then
                lngReturn = -3
            End If
        Else
            '在可用自用范围内判断
            blnTmp = False
            Set cllBillItem = mdlPubJson.zlGetNodeObjectFromCollect(cllBillInfos, "_" & lngReturn)
           ' rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(Nvl(cllBillItem("_prefix_text"))))) <> UCase(Nvl(cllBillItem("_prefix_text"))) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(Nvl(cllBillItem("_start_no"))) And UCase(strBill) <= UCase(Nvl(cllBillItem("_end_no"))) And Len(strBill) = Len(UCase(Nvl(cllBillItem("_start_no"))))) Then
                blnTmp = True
            End If
            
            If blnTmp Then
                '该批不满足,则在其它自用中检查
                lngReturn = -3
                For i = 1 To cllBillInfos.Count
                    Set cllBillItem = cllBillInfos(i)
                    
                    If Nvl(cllBillItem("_recv_id")) <> lngReturn Then
                        blnTmp = False
                        
                        If UCase(Left(strBill, Len(Nvl(cllBillItem("_prefix_text"))))) <> UCase(Nvl(cllBillItem("_prefix_text"))) Then
                            blnTmp = True
                        ElseIf Not (UCase(strBill) >= UCase(Nvl(cllBillItem("_start_no"))) And UCase(strBill) <= UCase(Nvl(cllBillItem("_end_no"))) And Len(strBill) = Len(UCase(Nvl(cllBillItem("_start_no"))))) Then
                            blnTmp = True
                        End If
                        If Not blnTmp Then lngReturn = Val(cllBillItem("_recv_id")): Exit For
                    End If
                Next
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    CheckUsedBill = 0
End Function

Public Function GetNextBill(lng领用ID As Long) As String
'功能：根据领用批次ID,获取下一个实际票据号
'说明：1.当取不到范围内的有效票据时,返回空由用户输入
'      2.排开已报损的号码
    Dim strSql As String, strBill As String
    Dim cllBillInfos As Collection, cllTemp As Collection
    Dim str下一张发票 As String
    
    On Error GoTo errH
    
    If Not zl_ExseSvr_GetReceiveInvoice(0, lng领用ID, cllBillInfos, True, , UserInfo.姓名, 1, True, , glngModul) Then
        Set cllBillInfos = New Collection
    End If
    
    If cllBillInfos.Count = 0 Then Exit Function
    
    Set cllTemp = cllBillInfos(1)
    '取下一个号码
    If Nvl(cllTemp("_inv_no_cur")) = "" Then
        strBill = UCase(Nvl(cllTemp("_start_no")))
    Else
        strBill = UCase(gobjCommFun.IncStr(Nvl(cllTemp("_inv_no_cur"))))
    End If
    
    '检查使用明细是否使用该票据
    If Not zl_ExseSvr_GetNextInvoice(lng领用ID, strBill, str下一张发票) Then Exit Function
    If str下一张发票 = "" Then Exit Function
    strBill = str下一张发票
    GetNextBill = strBill
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function



Public Function zl_GetInvoicePreperty(ByVal lngModule As Long, _
    ByVal int票据 As Integer, Optional str使用类别 As String) As clsFactProperty
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票格式
    '入参:int票据:1- 收费收据, 2 - 预交收据, 3 - 结帐收据, 4 - 挂号收据, 5 - 就诊卡
    '返回:发票的相关数据
    '编制:刘兴洪
    '日期:2011-07-19 16:43:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFactPro As clsFactProperty, strFactType As String, varData As Variant, varTemp As Variant
    Dim strShareTypeUseID As String, lng共用票据 As Long, lng使用票据 As Long
    Dim strFactTypeFormat As String, strFacePrintMode As String
    Dim intPrintMode As Long, intPrintMode1 As Long, lng领用ID As Long
    Dim strSql As String, rsTemp As ADODB.Recordset, strValue As String
    Dim i As Long, lngFormat As Long, lngFormat1 As Long
    
    strFactType = Switch(int票据 = 1, "共用收费票据批次", int票据 = 2, "共用预交票据批次", int票据 = 3, "共用结帐票据批次", int票据 = 4, "共用挂号票据批次", int票据 = 5, "共用医疗卡批次", True, "")
    strFactTypeFormat = Switch(int票据 = 1, "收费发票格式", int票据 = 2, "预交发票格式", int票据 = 3, "结帐发票格式", int票据 = 4, "挂号发票格式", int票据 = 5, "医疗卡发票格式", True, "")
    strFacePrintMode = Switch(int票据 = 1, "收费发票打印方式", int票据 = 2, "预交发票打印方式", int票据 = 3, "病人结帐打印", int票据 = 4, "挂号发票打印方式", int票据 = 5, "医疗卡发票打印方式", True, "")
    
    Set objFactPro = New clsFactProperty
    If strFactType = "" Then Set zl_GetInvoicePreperty = objFactPro: Exit Function
    
    
    '票号严格控制
    If int票据 >= 1 And int票据 <= 4 Then
        strValue = gobjDatabase.GetPara("票号严格控制", glngSys, , "00000")
        objFactPro.严格控制 = Mid(strValue, int票据, 1) = "1"
    End If
    
    '用四位数表示, 每一位代表不同的业务类型:
    '第一位:         收费
    '第二位:         预交
    '第三位:         结帐
    '第四位:         挂号
    '每位用1或0表示,1表示严格控制;0-表示非严格控制

    
    '78751:李南春,2014/10/20,增加预交票据打印格式
    objFactPro.使用类别 = str使用类别
 
    strFactTypeFormat = Trim(gobjDatabase.GetPara(strFactTypeFormat, glngSys, lngModule, ""))
    '格式:使用类别1,格式1|使用类别2,格式2...
    varData = Split(strFactTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lngFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then lngFormat1 = lngFormat
        If Trim(varTemp(0)) = str使用类别 And lngFormat <> 0 Then
            objFactPro.打印格式 = lngFormat: Exit For
        End If
    Next
    If objFactPro.打印格式 = 0 And lngFormat1 <> 0 Then objFactPro.打印格式 = lngFormat
 
    '打印方式(0-不打印;1-自动打印;2-提示打印)
    '问题50656
'    If int票据 = 2 Then
'        '预交暂为自动打印
'        objFactPro.打印方式 = 1
'    Else
        '因为Getpara就缓存了的,所以不用先用变量进行记录
        If lngModule = 1131 Then
            strFacePrintMode = Trim(gobjDatabase.GetPara("预交款票据打印", glngSys, lngModule, ""))
            objFactPro.打印方式 = Val(strFacePrintMode)
        Else
            strFacePrintMode = Trim(gobjDatabase.GetPara(strFacePrintMode, glngSys, lngModule, ""))
            objFactPro.打印方式 = -1
            '格式:使用类别1,打印方式1|使用类别2,打印方式2...
            varData = Split(strFacePrintMode, "|")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",,", ",")
                intPrintMode = Val(varTemp(1))
                If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
                If Trim(varTemp(0)) = str使用类别 Then
                    objFactPro.打印方式 = intPrintMode: Exit For
                End If
            Next
            If objFactPro.打印方式 < 0 Then objFactPro.打印方式 = intPrintMode1
        End If
'    End If
    '共享批次
    
    '格式:领用ID1,使用类别1|....
    strShareTypeUseID = Trim(gobjDatabase.GetPara(strFactType, glngSys, lngModule, "0"))
    varData = Split(strShareTypeUseID, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lng领用ID = Val(varTemp(0))
        If int票据 = 2 Or int票据 = 5 Then
            If Val(varTemp(1)) = 0 Then lng共用票据 = lng领用ID    '共用的.
            If Val(varTemp(1)) = Val(str使用类别) And lng领用ID <> 0 Then
                lng使用票据 = lng领用ID
            End If
        Else
            If Trim(varTemp(1)) = "" Then lng共用票据 = lng领用ID    '共用的.
            If Trim(varTemp(1)) = str使用类别 And lng领用ID <> 0 Then
                lng使用票据 = lng领用ID
            End If
        End If
    Next
    
    On Error GoTo errHandle
    '优先顺序
    '1.先使用
    '2.使用类别不区分的
    '3.具体使用类别的
    Dim cllBillInfo As Collection
    If zl_ExseSvr_GetReceiveInvoice(0, lng共用票据 & "," & lng使用票据, cllBillInfo, False) = False Then
        Set zl_GetInvoicePreperty = objFactPro
        Exit Function
    End If
    If cllBillInfo.Count <> 0 Then
         objFactPro.共享批次ID = Val(Nvl(cllBillInfo(1)("_recv_id"))) ' '共用的领用ID
    End If
    Set zl_GetInvoicePreperty = objFactPro
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zl_GetInvoiceUserType(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票的使用类别
    '返回:发票的使用类别
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str使用类别 As String
    On Error GoTo errHandle
    If zl_ExseSvr_GetPatiInvoiceClass(lng病人ID, lng主页ID, intInsure, str使用类别) = False Then Exit Function
    zl_GetInvoiceUserType = str使用类别
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Public Function zlStartFactUseType(ByVal int票种 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否使用了使用类别的
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-10 16:11:47
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln是否启用 As Boolean
    
    On Error GoTo errHandle
    If zl_ExseSvr_InvoiceClassUsed(int票种, bln是否启用, True, , glngModul) = False Then Exit Function
    zlStartFactUseType = bln是否启用
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


 

Public Function zl_ExseSvr_GetReceiveInvoice(ByVal int票种 As Integer, ByVal str领用ids As String, ByRef cllBillInfo_Out As Collection, _
    Optional ByVal bln是否自用 As Boolean = True, Optional ByVal str使用类别 As String, Optional ByVal str领用人 As String, _
    Optional ByVal int最少发票数 As Integer = 0, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long, Optional ByVal bytOperFun As Byte = 0, Optional ByVal blnResolveToRecord As Boolean, _
    Optional ByRef rsInvoice_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:卡费按门诊医疗票据使用时进行门诊医疗票据的相关更新操作
    '入参:int票种-1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    '     bln是否自用-true-自用：该票据仅供领用者自己使用；false-共用：该票据由多个人员共同使用
    '     str领用ids-多个用逗号(比如:1,2...)
    '     str使用类别-票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
    '     blnShowErrMsg-是否显示错误信息
    '     bytOperFun-0-获取票据领用信息 1-获取获取指定票种的共用票据批次
    '     blnResolveToRecord-是否转换为记录集
    '出参:cllBillInfo_Out-返回票据领用信息集
    '     rsInvoice_Out-返回发票信息集(blnResolveToRecord=true时)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
  
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的票据信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    'zl_ExseSvr_GetReceiveInvoice
    '   input
    '        oper_fun    N   1   0-获取票据领用信息 1-获取获取指定票种的共用票据批次
    '        recv_ids N   1   领用ids:票据领用id
    '        inv_type    N   1   票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    '        use_mode    N   1   使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
    '        use_type    C   1   票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
    '        recvtr  C   1   领用人
    '        min_nums  N 1 发票最少数量
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_fun", bytOperFun, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_type", int票种, Json_num)
    strJson = strJson & "," & GetJsonNodeString("recv_ids", str领用ids, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("use_mode", IIf(bln是否自用, 1, 2), Json_num)
    strJson = strJson & "," & GetJsonNodeString("use_type", str使用类别, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("min_nums", int最少发票数, Json_num)
    strJson = strJson & "," & GetJsonNodeString("recvtr", str领用人, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetReceiveInvoice"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '        code    C   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：成功时返回成功信息失败时返回具体的错误信息 ""
    '        item_list C
    '            recv_id N   1   领用ID
    '            use_mode    N   1   使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
    '            use_type    C   1   票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
    '            prefix_text C   1   前缀文本
    '            start_no    C   1   开始号码
    '            end_no  C   1   终止号码
    '            inv_no_cur  C   1   当前号码
    '            surplus_num C   1   剩余数量
    '            create_time C   1   登记时间:yyyy-mm-dd hh24:mi:ss
    '            use_time    C   1   使用时间:yyyy-mm-dd hh24:mi:ss
    '            recvtr  C   1   领用人

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "未找到符合条件的票据领用信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set cllBillInfo_Out = objServiceCall.GetJsonListValue("output.item_list", "recv_id")
    If blnResolveToRecord Then Call zlGetReceiveInvoiceRecFromCollect(cllBillInfo_Out, rsInvoice_Out)
    zl_ExseSvr_GetReceiveInvoice = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
 



Public Function zl_ExseSvr_GetNextInvoice(ByVal lng领用ID As Integer, ByVal str发票号 As String, ByRef str下一张发票_Out As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:卡费按门诊医疗票据使用时进行门诊医疗票据的相关更新操作
    '入参:str发票号-下一张发票号
    '     blnShowErrMsg-是否显示错误信息
    '出参:str下一张发票_Out-返回有效的票据号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    str下一张发票_Out = ""
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的票据号!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetNextInvoice
    '    input
    '       recv_id N   1   领用id:票据领用id
    '       inv_no  C   1   发票号

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("recv_id", lng领用ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_no", str发票号, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetNextInvoice"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '       code    C   1   应答码：0-失败；1-成功
    '       message C   1   "应答消息：
    '       inv_no  C   1   下一个发票号
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "未找到符合条件的票据信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    str下一张发票_Out = objServiceCall.GetJsonNodeValue("output.inv_no")
    zl_ExseSvr_GetNextInvoice = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function






Public Function zl_ExseSvr_GetPatiInvoiceClass(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int险类 As Integer, ByRef str使用类别_Out As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据使用类别
    '入参:lng病人id
    '     lng主页id
    '    int险类
    '     blnShowErrMsg-是否显示错误信息
    '出参:str使用类别_Out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的票据使用类别!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetPatiInvoiceClass
    ' input
    '    pati_id N   1   病人id
    '    pati_pageid N   1   主页id
    '    insure_type N   1   险类

    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("insure_type", int险类, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetPatiInvoiceClass"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '       code    C   1   应答码：0-失败；1-成功
    '       message C   1   应答消息： 失败时返回具体的错误信息
    '       use_type    C   1   票据使用类别
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取票据的使用类别，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    str使用类别_Out = Trim(Nvl(objServiceCall.GetJsonNodeValue("output.use_type")))
    zl_ExseSvr_GetPatiInvoiceClass = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zl_ExseSvr_InvoiceClassUsed(ByVal int票种 As Integer, bln是否启用_Out As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定票种是否启用了使用类别
    '入参:int票种-1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    '     blnShowErrMsg-是否显示错误信息
    '出参:bln是否启用_Out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的票据使用类别!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_InvoiceClassUsed
    ' input
    '    inv_type    N   1   票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
 

    strJson = strJson & "" & GetJsonNodeString("inv_type", int票种, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_InvoiceClassUsed"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '       code    C   1   应答码：0-失败；1-成功
    '       message C   1   应答消息： 失败时返回具体的错误信息
    '       is_start    N   1   是否启用:1-启用了的，0-未启用

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取票据的使用类别，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln是否启用_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.is_start"))) = 1
    zl_ExseSvr_InvoiceClassUsed = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function




Public Function zlGetReceiveInvoiceRecFromCollect(ByVal cllInvoice As Collection, ByRef rsInvoice_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据服务返回的票据信息集合,按记录集方式返回信息
    '入参:cllCardFee-当前集合
    '
    '出参:rsCardFee_Out-返回的卡费用集合
    '     objBalanceItems_out-结算信息列表，主要是可能存在记帐，需要给objBalanceItems_out
    '     dblMoney_Out:实收金额
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection
    Dim i As Long
    
    On Error GoTo errHandle
      
    Set rsInvoice_Out = New ADODB.Recordset
    With rsInvoice_Out
        If .State = adStateOpen Then .Close
        
        .Fields.Append "ID", adBigInt, , adFldIsNullable
        .Fields.Append "使用类别编码", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "使用类别ID", adBigInt, , adFldIsNullable
        .Fields.Append "使用类别", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "领用人", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "登记时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "开始号码", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "终止号码", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "剩余数量", adDouble, , adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    If cllInvoice Is Nothing Then Exit Function
    If cllInvoice.Count = 0 Then Exit Function
    
    '    recv_id N   1   领用ID
    '    use_mode    N   1   使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
    '    use_type    C   1   票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
    '    prefix_text C   1   前缀文本
    '    start_no    C   1   开始号码
    '    end_no  C   1   终止号码
    '    inv_no_cur  C   1   当前号码
    '    surplus_num C   1   剩余数量
    '    create_time C   1   登记时间:yyyy-mm-dd hh24:mi:ss
    '    use_time    C   1   使用时间:yyyy-mm-dd hh24:mi:ss
    '    recvtr  C   1   领用人
    '    use_typecode    C   1   使用类别编码
    '    use_typeid  N   1   使用类别id
    For i = 1 To cllInvoice.Count
        Set cllTemp = cllInvoice(i)
        With rsInvoice_Out
            .AddNew
            !ID = Val(Nvl(cllTemp("_recv_id")))
            !使用类别编码 = Nvl(cllTemp("_use_typecode"))
            !使用类别ID = Val(Nvl(cllTemp("_use_typeid")))
            !使用类别 = Nvl(cllTemp("_use_type"))
            !领用人 = Nvl(cllTemp("_recvtr"))
            !登记时间 = Nvl(cllTemp("_create_time"))
            !开始号码 = Nvl(cllTemp("_start_no"))
            !终止号码 = Nvl(cllTemp("_end_no"))
            !剩余数量 = Val(Nvl(cllTemp("_surplus_num")))
            .Update
        End With
    Next
    zlGetReceiveInvoiceRecFromCollect = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function Zl_Exsesvr_UpdCardFeeBlncInfo(ByVal int操作状态 As Integer, ByVal clsSendCardInfo As Collection, ByVal cllUpdateDate As Collection, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改卡费结算数据
    '入参:int操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
    '   clsSendCardInfo-发卡信息(卡类别ID,变动类型,卡号,原卡号,IC卡号,密码,加密密码,终止使用时间,卡费,病历费,摘要,卡号重用,领用ID),格式:array(名称,值),"_名称"
    '   cllUpdateDate-修改的结算数据
    '         |--billinfo-单据信息,"_billinfo"
    '              |-预交单号,预交ID,收费单号,结帐ID,操作员编号,操作员姓名,收款时间,病人ID)
    '         |--balanceinfo-结算信息,"_balanceinfo"
    '                |--(结算方式,结算号码,卡类别id,结算卡序号,卡号,交易流水号,交易说明,摘要,合作单位)
    '                |--其他信息集,
    '                |-----其他信息:交易名称,交易内容
    '     blnShowErrMsg-是否显示错误信息
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, j As Long, m As Long, strServiceName  As String
    Dim clldata As Collection, cllTemp As Collection, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strSendCardJson As String, strBalanceJson As String, strOthersJson As String
 
    Dim strJsonTemp As String, cllOthers As Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法修改卡费结算信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    'Zl_Exsesvr_UpdCardFeeBlncInfo
    'input
    '   oper_fun    N   1   操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
    '    pati_id N   1   病人id
    '    fee_no  C   1   费用单号：本次要调整的费用单据
    '    balance_id  N       结帐ID
    '    operator_name   C   1   操作员姓名
    '    operator_code   C   1   操作员编号
    '    create_time C   1   操作时间:yyyy-mm-dd hh:mi:ss
    '    sendcard_info 发卡信息
    '        send_mode   N   1   发卡方式;0-发卡,1-补卡,2-换卡
    '        cardtype_id C   1   卡类别id
    '        cardno  C   1   卡号:本次发放或绑定或补卡的卡号
    '        recv_id N   1   领用id:票据领用ID(卡号)
    '        cardno_reusing  N   1   卡号重用:1-卡号允许重复使用用;0-不允许重复使用
    '        cardno_old  C   1   原卡卡号:换卡时，需要传入原卡号
    '    balance_info    C       结算信息
    '        deposit_no  C       预交单号
    '        deposit_id  N       预交ID
    '        pay_mode    C   1   结算方式
    '        blnc_no C   1   结算号码
    '        cardtype_id N   1   卡类别id
    '        consumer_no N   1   结算卡序号，即卡消费接口目录.编号
    '        cardno  C   1   卡号
    '        swapno  C   1   交易流水号
    '        swapmemo    C   1   交易说明
    '        memo    C   1   摘要
    '        cprtion_unit    C   1   合作单位
    '        other_list[]    C   1   其他交易信息
    '            swap_name   C   1   交易名称
    '            swap_note   C   1   交易内容

    Set clldata = cllUpdateDate("_billinfo")
    strJson = ""
    strSendCardJson = ""
    
    strJson = strJson & "," & GetJsonNodeString("oper_fun", int操作状态, Json_num)
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "预交单号"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("deposit_no", Trim(varTemp(1)), Json_Text)
        Case "预交ID"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("deposit_id", Val(varTemp(1)), Json_num, True)
        Case "病人ID"
            strJson = strJson & "," & GetJsonNodeString("pati_id", Val(varTemp(1)), Json_num, True)
        Case "收费单号"
            strJson = strJson & "," & GetJsonNodeString("fee_no", Trim(varTemp(1)), Json_Text)
        Case "结帐ID"
            strJson = strJson & "," & GetJsonNodeString("balance_id", Val(varTemp(1)), Json_num, True)
        Case "操作员编号"
            strJson = strJson & "," & GetJsonNodeString("operator_code", Trim(varTemp(1)), Json_Text)
        Case "操作员姓名"
            strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
        
        Case "收款时间"
            strJson = strJson & "," & GetJsonNodeString("create_time", Trim(varTemp(1)), Json_Text)
        End Select
    Next
    
    '        send_mode   N   1   发卡方式;0-发卡,1-补卡,2-换卡；3-退卡
    '        cardtype_id C   1   卡类别id
    '        cardno  C   1   卡号:本次发放或绑定或补卡的卡号
    '        recv_id N   1   领用id:票据领用ID(卡号)
    '        cardno_reusing  N   1   卡号重用:1-卡号允许重复使用用;0-不允许重复使用
    '        cardno_old  C   1   原卡卡号:换卡时，需要传入原卡号
    If Not clsSendCardInfo Is Nothing Then
        strSendCardJson = ""
        For i = 1 To clsSendCardInfo.Count
            varTemp = clsSendCardInfo(i)
            Select Case UCase(varTemp(0))
            Case "卡类别ID"
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num, True)
            Case "变动类型"
                '1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失);7-终止时间调整
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("send_mode", Decode(Val(varTemp(1)), 1, 0, 11, 0, 2, 2, 3, 1, 3), Json_num)
            Case "卡号", "医疗卡号"
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
            Case "原卡号"
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("cardno_old", Trim(varTemp(1)), Json_Text)
            Case "IC卡号"
            Case "密码"
            Case "加密密码"
            Case "终止使用时间"
            Case "卡费"
            Case "病历费"
            Case "摘要"
            Case "卡号重用"
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("cardno_reusing", Val(varTemp(1)), Json_num, True)
            Case "领用ID"
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("recv_id", Val(varTemp(1)), Json_num, True)
            End Select
        Next
        If strSendCardJson = "" Then Exit Function
        strSendCardJson = Mid(strSendCardJson, 2)
        strJson = strJson & "," & GetNodeString("sendcard_info") & ":{" & strSendCardJson & "}"
    End If
    '结算信息
    '        deposit_no  C       预交单号
    '        deposit_id  N       预交ID
    '        pay_mode    C   1   结算方式
    '        blnc_no C   1   结算号码
    '        cardtype_id N   1   卡类别id
    '        consumer_no N   1   结算卡序号，即卡消费接口目录.编号
    '        cardno  C   1   卡号
    '        swapno  C   1   交易流水号
    '        swapmemo    C   1   交易说明
    '        memo    C   1   摘要
    '        statu   N   1   0-完成结算;1-接口调用前,2-接口调用成功
    '        cprtion_unit    C   1   合作单位
    '        other_list[]    C   1   其他交易信息
    '            swap_name   C   1   交易名称
    '            swap_note   C   1   交易内容

    Set clldata = cllUpdateDate("_balanceinfo")
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "结算方式"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("pay_mode", Trim(varTemp(1)), Json_Text)
        Case "结算号码"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("blnc_no", Trim(varTemp(1)), Json_Text)
        Case "卡类别ID"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num)
        Case "结算卡序号"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("consumer_no", Val(varTemp(1)), Json_num)
        Case "卡号"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
        Case "交易流水号"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("swapno", Trim(varTemp(1)), Json_Text)
        Case "交易说明"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("swapmemo", Trim(varTemp(1)), Json_Text)
        Case "摘要"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("memo", Trim(varTemp(1)), Json_Text)
        Case "合作单位"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("cprtion_unit", Trim(varTemp(1)), Json_Text)
        Case "校对标志"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("statu", Val(varTemp(1)), Json_num)
        Case UCase("其他信息集") '其他扩展交易信息
            Set cllOthers = varTemp(1)
            strOthersJson = ""
            For j = 1 To cllOthers.Count
                 Set cllTemp = cllOthers(j)
                 strJsonTemp = ""
                 For m = 1 To cllTemp.Count
                    varTemp = cllTemp(m)
          
                    Select Case UCase(varTemp(0))
                    Case "交易名称"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swap_name", Trim(varTemp(1)), Json_Text)
                    Case "交易内容"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swap_note", Trim(varTemp(1)), Json_Text)
                    End Select
                 Next
                 If strJsonTemp <> "" Then
                    strJsonTemp = Mid(strJsonTemp, 2)
                    strOthersJson = strOthersJson & ",{" & strJsonTemp & "}"
                 End If
            Next
        End Select
    Next
    If strOthersJson <> "" Then
        strOthersJson = Mid(strOthersJson, 2)
       strBalanceJson = strBalanceJson & "," & GetNodeString("other_list") & ":[" & strOthersJson & "]"
    End If
     
    If strBalanceJson <> "" Then
        strBalanceJson = Mid(strBalanceJson, 2)
        strJson = strJson & "," & GetNodeString("balance_info") & ":{" & strBalanceJson & "}"
    End If
    If strJson = "" Then Exit Function
    strJson = Mid(strJson, 2)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    
    strServiceName = "Zl_Exsesvr_UpdCardFeeBlncInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    Zl_Exsesvr_UpdCardFeeBlncInfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function Zl_Exsesvr_AddCardFeeInfo(ByVal int操作状态 As Integer, cllCardData As Collection, ByRef lng结帐ID_Out As Long, lng预交ID_Out As Long, _
    Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加卡费及预交数据
    '入参:int操作状态-操作状态:0-正常的预交款或卡费缴款;1-保存为未生效的预交款或异常的卡费;2-保存为记帐单;3-保存为划价单
    '     cllData: 卡数据对象
    '          |--billinfo:(结算合计,操作员编号,操作员姓名,登记时间),Key="_billinfo"
    '          |--patinfo:(病人ID,主页ID,病人姓名,性别,年龄,门诊号,住院号,付款方式编号,费别,险类),Key="_patinfo"
    '          |--cardinfo:发卡信息(卡号,卡类别ID,发卡方式(0-发卡,1-补卡,2-换卡),卡号重用,领用id),key="_cardinfo"
    '          |--cardfeelists:key="_cardfeelists"
    '               |---cardfeelist:(卡费单据号,序号,价格父号,从属父号,收费类别,收费细目id,收入项目id,标准单价,收据费目,应收金额,实收金额,病人科室id,开单部门id,病人病区id,
    '                                 执行部门id,是否病历费,保险编码,保险项目否,统筹金额,摘要,发卡卡号,发卡卡类别ID,发卡方式(0-发卡,1-补卡,2-换卡)) ,Key="_" & 序号
    
    '          |--balanceinfo:(结算方式,结算号码,卡类别id,结算卡序号,消费卡ID,支付卡号,交易流水号,交易说明,合作单位) Key="_balanceinfo"
    '          |--depositinfo:(预交单据号,发票号,预交类别,主页id,缴款科室id,缴款金额,缴款单位,单位开户行,摘要,领用id),Key="_depositinfo",无预交时，不传入
    '          以上，格式为:,格式：array(名称,值)
    '          int操作状态=2-保存为记帐单;3-保存为划价单 的，则无"balanceinfo"和"depositinfo"节点
    '     blnShowErrMsg-是否显示错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllRow As Collection, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    Dim strJsonTemp  As String, strJsonFee As String, blnHaveDeposit As Boolean
    Dim j As Long
    
    If cllCardData Is Nothing Then Exit Function
    If cllCardData.Count = 0 Then Exit Function
    
    If blnShowErrMsg Then On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "连接费用域服务失败，无法增加卡费信息!"
        If Not blnShowErrMsg Then Err.Raise -1001, strErrMsg, strErrMsg: Exit Function
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    '1.先取单据结算信息
    '    oper_fun    N   1   操作状态:0-正常的预交款或卡费缴款;1-保存为未生效的预交款或异常的卡费;2-保存为记帐单;3-保存为划价单
    '    blnc_total  N   1   结算合计:预交+卡费
    '    operator_name   C   1   操作员姓名
    '    operator_code   C   1   操作员编号
    '    create_time C   1   登记时间或收款时间:yyyy-mm-dd hh:mi:ss
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_fun", int操作状态, Json_num)
    Set clldata = cllCardData("_billinfo")
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "结算合计"
            strJson = strJson & "," & GetJsonNodeString("blnc_total", Val(Trim(varTemp(1))), Json_num)
        Case "操作员姓名", "操作员"
            strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
        Case "操作员编号"
            strJson = strJson & "," & GetJsonNodeString("operator_code", Trim(varTemp(1)), Json_Text)
        Case "登记时间", "收款时间"
            strJson = strJson & "," & GetJsonNodeString("create_time", Trim(varTemp(1)), Json_Text)
        Case "结帐ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("balance_id", Val(varTemp(1)), Json_num, True)
        Case Else
        End Select
    Next
    
    '2.取病人信息
    Set clldata = cllCardData("_patinfo")
    
    '    pati_info   C       病人信息
    '        pati_id N   1   病人ID
    '        pati_pageid N   1   主页id
    '        pati_name   C   1   病人姓名
    '        pati_sex    C   1   性别
    '        pati_age    C   1   年龄
    '        outpatient_num  C   1   门诊号
    '        inpatient_num   C   1   住院号
    '        mdlpay_name    C   1   付款方式名称
    '        fee_category    C   1   费别
    '        insurance_type  N   1   险类
    strJsonTemp = ""
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "病人ID"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_id", Val(Trim(varTemp(1))), Json_num)
        Case "主页ID"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_pageid", Val(Trim(varTemp(1))), Json_num)
        Case "姓名", "病人姓名"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_name", Trim(varTemp(1)), Json_Text)
        Case "性别"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_sex", Trim(varTemp(1)), Json_Text)
        Case "门诊号"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("outpatient_num", Trim(varTemp(1)), Json_Text)
        Case "住院号"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("inpatient_num", Trim(varTemp(1)), Json_Text)
        Case "年龄"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_age", Trim(varTemp(1)), Json_Text)
        Case "费别"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_category", Trim(varTemp(1)), Json_Text)
        Case "险类"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_type", Val(varTemp(1)), Json_num, True)
        Case "付款方式名称"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("mdlpay_name", Trim(Trim(varTemp(1))), Json_Text)
        Case Else
        End Select
    Next
    If strJsonTemp = "" Then Exit Function
    
    strJsonTemp = Mid(strJsonTemp, 2)
    strJsonTemp = GetNodeString("pati_info") & ":{" & strJsonTemp & "}"
    strJson = strJson & "," & strJsonTemp
    
    '3.取发卡信息
    '    card_info   C       医疗卡信息
    '        cardno  C   1   发卡卡号
    '        cardtype_id N   1   发卡卡类别ID
    '        send_mode   N   1   发卡方式;0-发卡,1-补卡,2-换卡
    '        cardno_reusing  N   1   卡号重用:1-重用;0-不以许重用
    '        recv_id N   1   领用id:领用Id
    '
    '
    Set clldata = cllCardData("_cardinfo")
    strJsonTemp = ""
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "发卡卡号", "卡号"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
        Case "发卡卡类别ID", "卡类别ID"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(Trim(varTemp(1))), Json_num)
        Case "发卡方式"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("send_mode", Val(Trim(varTemp(1))), Json_num)
        Case "卡号重用", "是否卡号重用", "是否卡号重复使用"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno_reusing", Val(Trim(varTemp(1))), Json_num)
        Case "领用ID"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("recv_id", Val(Trim(varTemp(1))), Json_num, True)
        Case Else
        End Select
    Next
    If strJsonTemp = "" Then Exit Function
    
    
    strJsonTemp = Mid(strJsonTemp, 2)
    strJsonTemp = GetNodeString("card_info") & ":{" & strJsonTemp & "}"
    strJson = strJson & "," & strJsonTemp
    '2.取费用信息
    Set clldata = cllCardData("_cardfeelists")
    '    cardfee_list[]  C   1   卡费列表
    '      fee_no  C   1   卡费单据号
    '      serial_num  N   1   序号
    '      price_ftrnum    N   1   价格父号
    '      subde_ftrnum    N   1   从属父号
    '      receipt_type    C   1   收费类别
    '      fitem_id    N   1   收费细目id
    '      income_item_id  N   1   收入项目id
    '      price   N   1   标准单价
    '      receipt_fee C   1   收据费目
    '      fee_amrcvb  N   1   应收金额
    '      fee_ampaib  N   1   实收金额
    '      pati_deptid N   1   病人科室id
    '      bill_deptid N   1   开单部门id
    '      pati_wardarea_id    N   1   病人病区id
    '      exedept_id  N   1   执行部门id
    '      mrbkfee_sign N   1   是否病历费:1-是病历费;0-不是病历费
    '      insurance_code  C   1   保险编码
    '      insurance_type_id   N   1   保险大类id
    '      insurance_sign  N   1   保险项目否:1-是保险项目;0-不是保险项目
    '      si_manp_money   N   1   统筹金额
    '      memo    C   1   摘要
 
     strJsonFee = ""
    For i = 1 To clldata.Count
        Set cllRow = clldata(i)
        strJsonTemp = ""
        For j = 1 To cllRow.Count
            varTemp = cllRow(j)
            Select Case UCase(varTemp(0))
            Case "卡费单据号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_no", Trim(varTemp(1)), Json_Text)
            Case "序号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("serial_num", Val(Trim(varTemp(1))), Json_num)
            Case "价格父号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("price_ftrnum", Val(Trim(varTemp(1))), Json_num, True)
            Case "从属父号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("subde_ftrnum", Val(Trim(varTemp(1))), Json_num, True)
            Case "收费类别"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("receipt_type", Trim(varTemp(1)), Json_Text)
            Case "收费细目ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fitem_id", Val(Trim(varTemp(1))), Json_num)
            Case "收入项目ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("income_item_id", Val(Trim(varTemp(1))), Json_num)
            Case "标准单价"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("price", Val(Trim(varTemp(1))), Json_num)
            Case "收据费目"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("receipt_fee", Trim(varTemp(1)), Json_Text)
            Case "应收金额"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_amrcvb", Val(Trim(varTemp(1))), Json_num)
            Case "实收金额"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_ampaib", Val(Trim(varTemp(1))), Json_num)
            Case "病人科室ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_deptid", Val(Trim(varTemp(1))), Json_num)
            Case "开单部门ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("bill_deptid", Val(Trim(varTemp(1))), Json_num)
            Case "病人病区ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_wardarea_id", Val(Trim(varTemp(1))), Json_num)
            Case "执行部门ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("exedept_id", Val(Trim(varTemp(1))), Json_num)
            Case "是否病历费"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("mrbkfee_sign", Val(Trim(varTemp(1))), Json_num)
            Case "保险编码"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_code", Trim(varTemp(1)), Json_Text)
            Case "保险项目否"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_sign", Val(Trim(varTemp(1))), Json_num)
            Case "统筹金额"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("si_manp_money", Val(Trim(varTemp(1))), Json_num)
            Case "摘要"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("memo", Trim(varTemp(1)), Json_Text)
            End Select
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonFee = strJsonFee & ",{" & strJsonTemp & "}"
        End If
    Next
    If strJsonFee = "" Then Exit Function
    strJsonFee = Mid(strJsonFee, 2)
    strJsonFee = GetNodeString("cardfee_list") & ":[" & strJsonFee & "]"
    strJson = strJson & "," & strJsonFee
    
    
    If int操作状态 < 2 Then '2-保存为记帐单;3-保存为划价单
        '非记帐或划价时，才有效
        '4.取结算信息
        'balance_info    C       结算信息:目前只支持一种结算方式
        '    blnc_mode    C   1   结算方式
        '    blnc_no C   1   结算号码
        '    cardtype_id N   1   卡类别id
        '    consumer_no N   1   结算卡序号，即卡消费接口目录.编号
        '    consume_card_id N   1   消费卡ID
        '    cardno  C   1   支付卡号
        '    swapno  C   1   交易流水号
        '    swapmemo    C   1   交易说明
        '    cprtion_unit    C   1   合作单位
        
        Set clldata = cllCardData("_balanceinfo")
        strJsonTemp = ""
        For i = 1 To clldata.Count
            varTemp = clldata(i)
            Select Case UCase(varTemp(0))
            Case "结算方式"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("blnc_mode", Trim(varTemp(1)), Json_Text)
            Case "结算号码"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("blnc_no", Trim(varTemp(1)), Json_Text)
            Case "卡类别ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(Trim(varTemp(1))), Json_num, True)
            Case "结算卡序号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("consumer_no", Val(Trim(varTemp(1))), Json_num, True)
            Case "支付卡号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
            Case "交易流水号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swapno", Trim(varTemp(1)), Json_Text)
            Case "交易说明"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swapmemo", Trim(varTemp(1)), Json_Text)
            Case "合作单位"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cprtion_unit", Trim(varTemp(1)), Json_Text)
            Case "消费卡ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("consume_card_id", Val(varTemp(1)), Json_num, True)

            Case Else
            End Select
        Next
        If strJsonTemp = "" Then Exit Function
        strJsonTemp = Mid(strJsonTemp, 2)
        strJsonTemp = GetNodeString("balance_info") & ":{" & strJsonTemp & "}"
        strJson = strJson & "," & strJsonTemp
        
        '5.获取预交款
        
        Err = 0: On Error Resume Next
        Set clldata = cllCardData("_depositinfo")
        
        blnHaveDeposit = True
        If Err <> 0 Then
            Err = 0: On Error GoTo 0
            blnHaveDeposit = False
        End If
        
        On Error GoTo errHandle
        
        If blnHaveDeposit Then
            'deposit_info    C   1   预交款列表
            '    deposit_no  C   1   预交单据号
            '    fact_no C   1   发票号
            '    deposit_type    N       预交类别:1-门诊;2-住院
            '    pati_pageid N   1   主页id
            '    dept_id N   1   缴款科室id
            '    money   N   1   缴款金额
            '    emp_name    C   1   缴款单位
            '    emp_bank_name   C   1   单位开户行
            '    emp_bank_actno  C   1   开户行账号
            '    memo    C   1   摘要
            '    recv_id N   1   领用id:领用Id
            '
            strJsonTemp = ""
            For i = 1 To clldata.Count
                varTemp = clldata(i)
                Select Case UCase(varTemp(0))
                
                Case "预交单据号"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("deposit_no", Trim(varTemp(1)), Json_Text)
                Case "发票号"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fact_no", Trim(varTemp(1)), Json_Text)
                Case "预交类别"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("deposit_type", Val(Trim(varTemp(1))), Json_num, True)
                Case "主页ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_pageid", Val(Trim(varTemp(1))), Json_num, True)
                Case "缴款科室ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("dept_id", Val(Trim(varTemp(1))), Json_num, True)
                Case "缴款金额"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("money", Val(Trim(varTemp(1))), Json_num, True)
                Case "缴款单位"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("emp_bank_name", Trim(varTemp(1)), Json_Text)
                Case "单位开户行"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("emp_name", Trim(varTemp(1)), Json_Text)
                Case "摘要"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("memo", Trim(varTemp(1)), Json_Text)
                Case "领用ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("recv_id", Val(Trim(varTemp(1))), Json_num, True)
                Case "预交ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("deposit_id", Val(Trim(varTemp(1))), Json_num, True)
                End Select
            Next
            
            If strJsonTemp = "" Then Exit Function
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonTemp = "" & GetNodeString("deposit_info") & ":{" & strJsonTemp & "}"
            strJson = strJson & "," & strJsonTemp
        End If
    End If

    If strJson = "" Then Exit Function
     
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    
    strServiceName = "Zl_Exsesvr_Addcardfeeinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then
        Exit Function
    End If
    '    output
    '       code    C   1   应答码：0-失败；1-成功
    '       message C   1   应答消息：失败时返回具体的错误信息
    '       deposit_id  N   1   预交ID
    '       balance_id  N   1   结帐ID
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        mstrErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If mstrErrMsg = "" Then
            mstrErrMsg = "费用处理失败，请检查！"
        End If
        If blnShowErrMsg Then MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    

    lng预交ID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.deposit_id")))
    lng结帐ID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.balance_id")))
    Zl_Exsesvr_AddCardFeeInfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function zl_ExseSvr_AddDepositInfo(ByVal int操作状态 As Integer, cllDepositData As Collection, lng预交ID_Out As Long, _
    Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加预交数据服务
    '入参:int操作状态-操作状态:0-正常的预交款 ;1-保存为未生效的预交款
    '     cllDepositData: 预交数据对象
    '          |--billinfo:(操作员编号,操作员姓名,收款时间),Key="_billinfo"
    '          |--balanceinfo:(结算方式,结算号码,卡类别id,结算卡序号,支付卡号,交易流水号,交易说明,合作单位,险类,医保号,医保密码,消费卡ID) Key="_balanceinfo"
    '          |--depositinfo:(预交单据号,发票号,预交类别,病人ID,主页id,姓名,性别,年龄,门诊号,住院号,付款方式编号,付款方式名称,缴款科室id,缴款金额,缴款单位,单位开户行,摘要,领用id),Key="_depositinfo",无预交时，不传入
    '          以上，格式为:,格式：array(名称,值)
    '     blnShowErrMsg-是否显示错误信息
    '出参:lng预交ID_Out-预交ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllRow As Collection, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    Dim strJsonTemp  As String, strJsonFee As String, blnHaveDeposit As Boolean
    Dim j As Long
    
    If cllDepositData Is Nothing Then Exit Function
    If cllDepositData.Count = 0 Then Exit Function
    
    If blnShowErrMsg Then On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "连接费用域服务失败，无法增加预交信息!"
        If Not blnShowErrMsg Then Err.Raise -1001, strErrMsg, strErrMsg: Exit Function
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    '1.先取单据结算信息
    '    oper_fun    N   1   操作状态:0-正常的预交款或卡费缴款;1-保存为未生效的预交款或异常的卡费;2-保存为记帐单;3-保存为划价单
    '    operator_name   C   1   操作员姓名
    '    operator_code   C   1   操作员编号
    '    create_time C   1   登记时间或收款时间:yyyy-mm-dd hh:mi:ss
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_fun", int操作状态, Json_num)
    Set clldata = cllDepositData("_billinfo")
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "操作员姓名", "操作员"
            strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
        Case "操作员编号"
            strJson = strJson & "," & GetJsonNodeString("operator_code", Trim(varTemp(1)), Json_Text)
        Case "登记时间", "收款时间"
            strJson = strJson & "," & GetJsonNodeString("create_time", Trim(varTemp(1)), Json_Text)
        Case Else
        End Select
    Next
   
       Set clldata = cllDepositData("_balanceinfo")
       strJsonTemp = ""
       For i = 1 To clldata.Count
           varTemp = clldata(i)
           Select Case UCase(varTemp(0))
           Case "结算方式"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("blnc_mode", Trim(varTemp(1)), Json_Text)
           Case "结算号码"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("blnc_no", Trim(varTemp(1)), Json_Text)
           Case "卡类别ID"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(Trim(varTemp(1))), Json_num, True)
           Case "结算卡序号"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("consumer_no", Val(Trim(varTemp(1))), Json_num, True)
           Case "支付卡号"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
           Case "交易流水号"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swapno", Trim(varTemp(1)), Json_Text)
           Case "交易说明"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swapmemo", Trim(varTemp(1)), Json_Text)
           Case "合作单位"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cprtion_unit", Trim(varTemp(1)), Json_Text)
           Case "险类"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_type", Val(varTemp(1)), Json_num, True)
           Case "医保号"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_num", Trim(varTemp(1)), Json_Text)
           Case "医保密码", "密码"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_pwd", Trim(varTemp(1)), Json_Text)
           
           Case "消费卡ID"
               strJsonTemp = strJsonTemp & "," & GetJsonNodeString("consume_card_id", Val(varTemp(1)), Json_num, True)
           Case Else
           End Select
       Next
       If strJsonTemp = "" Then Exit Function
       strJsonTemp = Mid(strJsonTemp, 2)
       strJsonTemp = GetNodeString("balance_info") & ":{" & strJsonTemp & "}"
       strJson = strJson & "," & strJsonTemp
       
       '5.获取预交款
       
       Err = 0: On Error Resume Next
       Set clldata = cllDepositData("_depositinfo")
       
       blnHaveDeposit = True
       If Err <> 0 Then
           Err = 0: On Error GoTo 0
           blnHaveDeposit = False
       End If
       
       On Error GoTo errHandle
       
       If blnHaveDeposit Then
           'deposit_info    C   1   预交款列表
            '    pati_id N   1   病人ID
            '    pati_pageid N   1   主页id
            '    pati_name   C   1   病人姓名
            '    pati_sex    C   1   性别
            '    pati_age    C   1   年龄
            '    outpatient_num  C   1   门诊号
            '    inpatient_num   C   1   住院号
            '    mdlpay_mode_name    C   1   付款方式名称
            '    deposit_id  N   1   预交ID
            '    deposit_no  C   1   预交单据号
            '    invc_no C   1   发票号
            '    deposit_type    N       预交类别:1-门诊;2-住院
            '    dept_id N   1   缴款科室id
            '    money   N   1   缴款金额
            '    emp_name    C   1   缴款单位
            '    emp_bank_name   C   1   单位开户行
            '    emp_bank_actno  C   1   开户行账号
            '    memo    C   1   摘要
            '    recv_id N   1   领用id:领用Id

           strJsonTemp = ""
           For i = 1 To clldata.Count
               varTemp = clldata(i)
                Select Case UCase(varTemp(0))
                Case "预交ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("deposit_id", Val(Trim(varTemp(1))), Json_num, True)
                Case "预交单据号"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("deposit_no", Trim(varTemp(1)), Json_Text)
                Case "发票号"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("invc_no", Trim(varTemp(1)), Json_Text)
                Case "预交类别"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("deposit_type", Val(Trim(varTemp(1))), Json_num, True)
                Case "主页ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_pageid", Val(Trim(varTemp(1))), Json_num, True)
                Case "缴款科室ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("dept_id", Val(Trim(varTemp(1))), Json_num, True)
                Case "缴款金额"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("money", Val(Trim(varTemp(1))), Json_num, True)
                Case "缴款单位"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("emp_bank_name", Trim(varTemp(1)), Json_Text)
                Case "单位开户行"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("emp_name", Trim(varTemp(1)), Json_Text)
                Case "摘要"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("memo", Trim(varTemp(1)), Json_Text)
                Case "领用ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("recv_id", Val(Trim(varTemp(1))), Json_num, True)
                Case "病人ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_id", Val(Trim(varTemp(1))), Json_num, True)
                Case "姓名", "病人姓名"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_name", Trim(varTemp(1)), Json_Text)
                Case "性别"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_sex", Trim(varTemp(1)), Json_Text)
                Case "年龄"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_age", Trim(varTemp(1)), Json_Text)
                Case "门诊号"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("outpatient_num", Trim(Trim(varTemp(1))), Json_Text)
                Case "住院号"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("inpatient_num", Trim(Trim(varTemp(1))), Json_Text)
                Case "付款方式编号"
                Case "付款方式名称"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("mdlpay_name", Trim(varTemp(1)), Json_Text)
               End Select
           Next
           
           If strJsonTemp = "" Then Exit Function
           strJsonTemp = Mid(strJsonTemp, 2)
           strJsonTemp = "" & GetNodeString("deposit_info") & ":{" & strJsonTemp & "}"
           strJson = strJson & "," & strJsonTemp
       End If


    If strJson = "" Then Exit Function
     
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    
    strServiceName = "zl_ExseSvr_AddDepositInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then
        Exit Function
    End If
    '    output
    '       code    C   1   应答码：0-失败；1-成功
    '       message C   1   应答消息：失败时返回具体的错误信息
    '       deposit_id  N   1   预交ID
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        mstrErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If mstrErrMsg = "" Then
            mstrErrMsg = "费用处理失败，请检查！"
        End If
        If blnShowErrMsg Then MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng预交ID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.deposit_id")))
    zl_ExseSvr_AddDepositInfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_ExseSvr_GetDepositInfo(ByVal strNo As String, ByVal int记录状态 As Integer, ByRef cllDeposit_Out As Collection, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人预交数据信息
    '入参:strNo-预交单号
    '     int记录状态-1-原始充值的预交记录;2-退款的预交记录
    '     blnShowErrMsg-是否显示错误信息
    '出参:cll_Deposit_Out-返回预交票据数据集,key="_"+名称
    '       |-病人ID,主页id,预交ID,预交单据号,发票号,预交类别 ,缴款科室id,缴款金额,缴款单位 ,单位开户行,开户行账号,摘要,操作员姓名,操作员编号,收款时间 ,结算方式,结算号码,
    '       | 卡类别id,结算卡序号 ,消费卡ID,支付卡号,交易流水号,交易说明,合作单位,结算状态 ,关联交易ID,险类,医保号,医保密码
    '
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Collection, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    Set cllDeposit_Out = New Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的预交信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    'zl_ExseSvr_GetDepositInfo
    '    input
    '        deposit_no  C   1   单据号:预交单据号
    '        rec_state   N   1   记录状态:1-原始充值的预交记录;2-退款的预交记录
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("deposit_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("rec_state", int记录状态, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetDepositInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", mlngModule, False) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：失败时返回具体的错误信息
    '    deposit_info    C   1   预交款信息
    '        pati_id N   1   病人ID
    '        pati_pageid N   1   主页id
    '        deposit_id  N   1   预交ID
    '        deposit_no  C   1   预交单据号
    '        invc_no C   1   发票号
    '        deposit_type    N       预交类别:1-门诊;2-住院
    '        dept_id N   1   缴款科室id
    '        money   N   1   缴款金额
    '        emp_name    C   1   缴款单位
    '        emp_bank_name   C   1   单位开户行
    '        emp_bank_actno  C   1   开户行账号
    '        memo    C   1   摘要
    '        operator_name   C   1   操作员姓名
    '        operator_code   C   1   操作员编号
    '        create_time C   1   收款时间:yyyy-mm-dd hh:mi:ss
    '    balance_info    C       结算信息:目前只支持一种结算方式
    '        blnc_mode   C   1   结算方式
    '        blnc_no C   1   结算号码
    '        cardtype_id N   1   卡类别id
    '        consumer_no N   1   结算卡序号，即卡消费接口目录.编号
    '        consume_card_id N   1   消费卡ID
    '        cardno  C   1   支付卡号
    '        swapno  C   1   交易流水号
    '        swapmemo    C   1   交易说明
    '        cprtion_unit    C   1   合作单位
    '        blnc_state  N   1   结算状态:0或NULL正常缴款记录;1-未调用接口;2-接口调用完成
    '        relation_id C   1   关联交易ID
    '        insurance_type  N       险类
    '        insurance_num   C       医保号
    '        insurance_pwd   C       医保密码
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "未找到符合条件的预交数据，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '单据信息
    Set clldata = objServiceCall.GetJsonListValue("output.deposit_info")
    cllDeposit_Out.Add Val(Nvl(clldata("_pati_id"))), "_病人ID"
    cllDeposit_Out.Add Val(Nvl(clldata("_pati_pageid"))), "_主页ID"
    cllDeposit_Out.Add Val(Nvl(clldata("_deposit_id"))), "_预交ID"
    cllDeposit_Out.Add Trim(Nvl(clldata("_deposit_no"))), "_预交单据号"
    cllDeposit_Out.Add Trim(Nvl(clldata("_invc_no"))), "_发票号"
    cllDeposit_Out.Add Val(Nvl(clldata("_deposit_type"))), "_预交类别"
    cllDeposit_Out.Add Val(Nvl(clldata("_dept_id"))), "_缴款科室ID"
    cllDeposit_Out.Add Val(Nvl(clldata("_money"))), "_缴款金额"
    cllDeposit_Out.Add Trim(Nvl(clldata("_emp_name"))), "_缴款单位"
    cllDeposit_Out.Add Trim(Nvl(clldata("_emp_bank_name"))), "_单位开户行"
    cllDeposit_Out.Add Trim(Nvl(clldata("_emp_bank_actno"))), "_开户行账号"
    cllDeposit_Out.Add Trim(Nvl(clldata("_memo"))), "_摘要"
    cllDeposit_Out.Add Trim(Nvl(clldata("_operator_name"))), "_操作员姓名"
    cllDeposit_Out.Add Trim(Nvl(clldata("_operator_code"))), "_操作员编号"
    cllDeposit_Out.Add Trim(Nvl(clldata("_create_time"))), "_收款时间"
    
    Set clldata = objServiceCall.GetJsonListValue("output.balance_info")
    cllDeposit_Out.Add Trim(Nvl(clldata("_blnc_mode"))), "_结算方式"
    cllDeposit_Out.Add Trim(Nvl(clldata("_blnc_no"))), "_结算号码"
    cllDeposit_Out.Add Val(Nvl(clldata("_cardtype_id"))), "_卡类别ID"
    cllDeposit_Out.Add Val(Nvl(clldata("_consumer_no"))), "_结算卡序号"
    cllDeposit_Out.Add Val(Nvl(clldata("_consume_card_id"))), "_消费卡ID"
    cllDeposit_Out.Add Trim(Nvl(clldata("_cardno"))), "_支付卡号"
    cllDeposit_Out.Add Trim(Nvl(clldata("_swapno"))), "_交易流水号"
    cllDeposit_Out.Add Trim(Nvl(clldata("_swapmemo"))), "_交易说明"
    cllDeposit_Out.Add Trim(Nvl(clldata("_cprtion_unit"))), "_合作单位"
    cllDeposit_Out.Add Val(Nvl(clldata("_blnc_state"))), "_结算状态"
    cllDeposit_Out.Add Trim(Nvl(clldata("_relation_id"))), "_关联交易ID"
    cllDeposit_Out.Add Val(Nvl(clldata("_insurance_type"))), "_险类"
    cllDeposit_Out.Add Trim(Nvl(clldata("_insurance_num"))), "_医保号"
    cllDeposit_Out.Add Trim(Nvl(clldata("_insurance_pwd"))), "_医保密码"
    zl_ExseSvr_GetDepositInfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zl_ExseSvr_GetCardFeeInfoByNo(ByVal strNo As String, ByVal int查询类型 As Integer, _
    ByRef cllFeeData_out As Collection, ByRef cllPriceBill_out As Collection, ByRef cllBalance_Out As Collection, ByRef cllDeposit_Out As Collection, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional bln查询预交 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的卡费单据，获取费用及结算及预交信息
    '入参:strNO- 单据号
    '     int查询类型-0-读取正常单据:1-读取作废单据;2-剩余费用单据
    '     blnShowErrMsg-是否显示错误信息
    '出参:cllFeeData_Out-发卡费用单信息
    '     cllPriceBill_out-划价单信息
    '     cllBalance_out-发卡结算信息
    '     cllDeposit_Out-发卡同时缴预交信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
     Set cllFeeData_out = New Collection: Set cllBalance_Out = New Collection: Set cllDeposit_Out = New Collection
     Set cllPriceBill_out = New Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的卡费信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetCardFeeInfoByNo
    '    input
    '       fee_no  C   1   单据号:费用单据号
    '       query_type  N       查询类型：0-读取正常单据:1-读取作废单据;2-剩余费用单据
    '       query_deposit   N   1   是否包含预交:1-包含预交信息，0-不包含预交信息



    strJson = strJson & "" & GetJsonNodeString("fee_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("query_type", int查询类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("query_deposit", IIf(bln查询预交, 1, 0), Json_num)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetCardFeeInfoByNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   "应答消息：
    '    fee_list            [数组]每个费用ID信息
    '        fee_id  N   1   费用id
    '        fee_num N   1   序号
    '        pati_id N   1   病人id
    '        pati_name   C   1   姓名
    '        pati_sex    C   1   性别
    '        pati_age    C   1   年龄
    '        fee_category    C   1   费别
    '        item_id N   1   收费项目id
    '        income_item_id  N   1   收入项目id
    '        quantity    N   1   数次
    '        fee_amrcvb  N   1   应收金额
    '        fee_ampaid  N   1   实收金额
    '        placer  C   1   开单人
    '        operator_code   C   1   操作员编号
    '        operator_name   C   1   操作员姓名
    '        create_time D   1   登记时间
    '        happen_time D   1   发生时间
    '        rec_status  N   1   记录状态
    '        mrbkfee_sign N   1   是否病历费:1-是病历费;0-不是病历费
    '        invoice_no  N   1   发票号
    '        kpbooks_sign N   1   是否记帐:1-是记帐;0-现收
    '        fee_status   N   1   费用状态:1-异常状态;0-正常费用
    '        cardtype_id N   1   卡类别ID
    '        card_no C   1   卡号
    '    pricebill_info  C       卡费生成划价费用信息
    '        fee_no  C       划价单号
    '        fee_amrcvb  N   1   应收金额
    '        fee_ampaid  N   1   实收金额
    '        charged_statu   N   1   收费状态:0-未收费;1-已收费;2-已全退
    '    balance_list[]  C       结算信息列表
    '        blnc_mode   C   1   结算方式名称
    '        balance_id  N   1   结帐ID
    '        blnc_money  N   1   结帐金额
    '        pay_cardno  N   1   支付卡号
    '        pay_swapno  C   1   交易流水号
    '        pay_swapmemo    C   1   交易说明
    '        relation_id N   1   关联交易id
    '        cardtype_id N   1   卡类别id
    '        consume_card    N   1   是否消费卡:1-是;0-不是
    '        blnc_nature N   1   结算性质:1-现金结算方式,2-其他非医保结算 , 8-结算卡结算 ,9-误差费
    '        blnc_statu  N   1   结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
    '        consume_card_id N   1   消费卡id
    '        blnc_no C 1 结算号码
    '        blnc_memo C 1 摘要
    '        original_id N 1 原结帐ID:冲销时返回
    '        original_money N,1 原始金额,求剩余款数时返回

    '    deposit_info    C       预交信息
    '        deposit_id  N   1   预交id
    '        deposit_no  C   1   预交单据号
    '        deposit_money   N   1   预交金额
    '        blnc_mode   C   1   结算方式
    '        pay_cardno  N   1   支付卡号
    '        pay_swapno  C   1   交易流水号
    '        pay_swapmemo    C   1   交易说明
    '        relation_id N   1   关联交易id
    '        cardtype_id N   1   卡类别id
    '        consume_card    N   1   是否消费卡:1-是;0-不是
    '        blnc_nature N   1   结算性质
    '        blnc_statu  N   1   结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
    '        consume_card_id N   1   消费卡id
    '        blnc_no C   1   结算号码
    '        blnc_memo   C   1   摘要

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取单据号为【" & strNo & "】的单据信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '获取费用信息集
    Set cllFeeData_out = objServiceCall.GetJsonListValue("output.fee_list")
    Set cllBalance_Out = objServiceCall.GetJsonListValue("output.balance_list")
    
    Set cllDeposit_Out = objServiceCall.GetJsonListValue("output.deposit_info")
    '划价单
   Set cllPriceBill_out = objServiceCall.GetJsonListValue("output.pricebill_info")
    
    If cllFeeData_out Is Nothing Then Set cllFeeData_out = New Collection
    If cllBalance_Out Is Nothing Then Set cllBalance_Out = New Collection
    If cllDeposit_Out Is Nothing Then Set cllDeposit_Out = New Collection
    If cllPriceBill_out Is Nothing Then Set cllPriceBill_out = New Collection
    
    zl_ExseSvr_GetCardFeeInfoByNo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Exsesvr_Upddepositblncinfo(ByVal int操作状态 As Integer, ByVal cllUpdateDate As Collection, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改预交结算信息
    '入参:int操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
     '   cllUpdateDate-修改的结算数据
    '         |--billinfo-单据信息,"_billinfo"
    '              |-预交单号,预交ID,操作员编号,操作员姓名,收款时间,病人ID,发票号，领用ID)
    '         |--balanceinfo-结算信息,"_balanceinfo"
    '                |--(结算方式,结算号码,卡类别id,结算卡序号,卡号,交易流水号,交易说明,摘要,合作单位)
    '                |--其他信息集,
    '                |-----其他信息:交易名称,交易内容
    '     blnShowErrMsg-是否显示错误信息
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, j As Long, m As Long, strServiceName  As String
    Dim clldata As Collection, cllTemp As Collection, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strSendCardJson As String, strBalanceJson As String, strOthersJson As String
 
    Dim strJsonTemp As String, cllOthers As Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法修正预交结算信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    'Zl_Exsesvr_Upddepositblncinfo
    'input
    '    oper_state  N   1   操作状态:0-完成结算;1-接口调用前修正;2-接口调用后修正
    '    pati_id N   1   病人id
    '    deposit_no  C       预交单号
    '    deposit_id  N       预交ID
    '    operator_name   C   1   操作员姓名
    '    operator_code   C   1   操作员编号
    '    create_time C   1   操作时间:yyyy-mm-dd hh:mi:ss
    '    invc_no C   1   发票号
    '    recv_id N   1   领用id:领用Id
    '    balance_info    C       结算信息
    '        blnc_mode   C   1   结算方式
    '        blnc_no C   1   结算号码
    '        cardtype_id N   1   卡类别id
    '        consumer_no N   1   结算卡序号，即卡消费接口目录.编号
    '        cardno  C   1   卡号
    '        swapno  C   1   交易流水号
    '        swapmemo    C   1   交易说明
    '        memo    C   1   摘要
    '        cprtion_unit    C   1   合作单位
    '        other_list[]    C   1   其他交易信息
    '            swap_name   C   1   交易名称
    '            swap_note   C   1   交易内容

    Set clldata = cllUpdateDate("_billinfo")
    strJson = ""
    strSendCardJson = ""
    
    strJson = strJson & "," & GetJsonNodeString("oper_state", int操作状态, Json_num)
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "预交单号"
            strJson = strJson & "," & GetJsonNodeString("deposit_no", Trim(varTemp(1)), Json_Text)
        Case "预交ID"
            strJson = strJson & "," & GetJsonNodeString("deposit_id", Val(varTemp(1)), Json_num, True)
        Case "病人ID"
            strJson = strJson & "," & GetJsonNodeString("pati_id", Val(varTemp(1)), Json_num, True)
        Case "发票号"
            strJson = strJson & "," & GetJsonNodeString("invc_no", Trim(varTemp(1)), Json_Text)
        Case "领用ID"
            strJson = strJson & "," & GetJsonNodeString("recv_id", Val(varTemp(1)), Json_num, True)
        Case "操作员编号"
            strJson = strJson & "," & GetJsonNodeString("operator_code", Trim(varTemp(1)), Json_Text)
        Case "操作员姓名"
            strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
        Case "收款时间"
            strJson = strJson & "," & GetJsonNodeString("create_time", Trim(varTemp(1)), Json_Text)
        End Select
    Next
     
    '结算信息
    '        deposit_no  C       预交单号
    '        deposit_id  N       预交ID
    '        pay_mode    C   1   结算方式
    '        blnc_no C   1   结算号码
    '        cardtype_id N   1   卡类别id
    '        consumer_no N   1   结算卡序号，即卡消费接口目录.编号
    '        cardno  C   1   卡号
    '        swapno  C   1   交易流水号
    '        swapmemo    C   1   交易说明
    '        memo    C   1   摘要
    '        statu   N   1   0-完成结算;1-接口调用前,2-接口调用成功
    '        cprtion_unit    C   1   合作单位
    '        other_list[]    C   1   其他交易信息
    '            swap_name   C   1   交易名称
    '            swap_note   C   1   交易内容

    Set clldata = cllUpdateDate("_balanceinfo")
    strBalanceJson = ""
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "结算方式"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("blnc_mode", Trim(varTemp(1)), Json_Text)
        Case "结算号码"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("blnc_no", Trim(varTemp(1)), Json_Text)
        Case "卡类别ID"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num)
        Case "结算卡序号"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("consumer_no", Val(varTemp(1)), Json_num)
        Case "卡号"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
        Case "交易流水号"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("swapno", Trim(varTemp(1)), Json_Text)
        Case "交易说明"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("swapmemo", Trim(varTemp(1)), Json_Text)
        Case "摘要"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("memo", Trim(varTemp(1)), Json_Text)
        Case "合作单位"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("cprtion_unit", Trim(varTemp(1)), Json_Text)
        Case UCase("其他信息集") '其他扩展交易信息
            Set cllOthers = varTemp(1)
            strOthersJson = ""
            For j = 1 To cllOthers.Count
                 Set cllTemp = cllOthers(j)
                 strJsonTemp = ""
                 For m = 1 To cllTemp.Count
                    varTemp = cllTemp(m)
          
                    Select Case UCase(varTemp(0))
                    Case "交易名称"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swap_name", Trim(varTemp(1)), Json_Text)
                    Case "交易内容"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swap_note", Trim(varTemp(1)), Json_Text)
                    End Select
                 Next
                 If strJsonTemp <> "" Then
                    strJsonTemp = Mid(strJsonTemp, 2)
                    strOthersJson = strOthersJson & ",{" & strJsonTemp & "}"
                 End If
            Next
        End Select
    Next
    If strOthersJson <> "" Then
        strOthersJson = Mid(strOthersJson, 2)
       strBalanceJson = strBalanceJson & "," & GetNodeString("other_list") & ":[" & strOthersJson & "]"
    End If
     
    If strBalanceJson <> "" Then
        strBalanceJson = Mid(strBalanceJson, 2)
        strJson = strJson & "," & GetNodeString("balance_info") & ":{" & strBalanceJson & "}"
    End If
    If strJson = "" Then Exit Function
    strJson = Mid(strJson, 2)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    
    strServiceName = "Zl_Exsesvr_Upddepositblncinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", mlngModule, False) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    Zl_Exsesvr_Upddepositblncinfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Exsesvr_DelDepositErrorRec(ByVal int操作状态 As Integer, ByVal strDepositNo As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除预交信息
    '入参:int操作状态:0-删除异常充值单据，1-删除异常退款单据，2-删除异常余额退款单据
     '   strDepositNo-单据号
    '     blnShowErrMsg-是否显示错误信息
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, j As Long, m As Long, strServiceName  As String
    Dim clldata As Collection, cllTemp As Collection, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strSendCardJson As String, strBalanceJson As String, strOthersJson As String
    Dim strJsonTemp As String, cllOthers As Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法删除预交异常数据!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    'Zl_Exsesvr_Deldepositerrorrec
    'input
    '    deposit_no  C   1   预交单号
    '    oper_state  N   1   操作状态:0-删除异常充值单据，1-删除异常退款单据，2-删除异常余额退款单据
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_state", int操作状态, Json_num)
    strJson = strJson & "," & GetJsonNodeString("deposit_no", strDepositNo, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Deldepositerrorrec"
    If objServiceCall.CallService(strServiceName, strJson, , "", mlngModule, False) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    Zl_Exsesvr_DelDepositErrorRec = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Exsesvr_DelCardfeeInfo(ByVal int操作状态 As Integer, cllDelFeeData As Collection, ByRef lng结帐ID_Out As Long, lng预交ID_Out As Long, _
    Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退卡费、病历费及预交数据
    '入参:int操作状态-操作状态:0-正常的预交款或卡费的退款记录;1-保存为异常的退款记录;2-作废异常数据
    '     cllDelFeeData-退费数据
    '        |-(卡费单号,预交单号,是否退卡费,是否退病历费,操作员姓名,操作员编号,退费时间,结算信息) array(名称,值) ,"_名称)
    '        |-结算信息:(退款金额,结算方式,结算号码,卡类别id,结算卡序号,支付卡号,交易流水号,交易说明,合作单位,关联交易ID) Key="_结算信息"
    '     blnShowErrMsg-是否显示错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Collection, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    Dim strJsonTemp  As String
    Dim j As Long
    
    If cllDelFeeData Is Nothing Then Exit Function
    If cllDelFeeData.Count = 0 Then Exit Function
    
    If blnShowErrMsg Then On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "连接费用域服务失败，无法删除卡费信息!"
        If Not blnShowErrMsg Then Err.Raise -1001, strErrMsg, strErrMsg: Exit Function
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    '1.先取单据结算信息
    'oper_fun    N   1   操作状态:0-正常的预交款或卡费的退款记录;1-保存为异常的退款记录;2-作废异常数据
    'cardfee_no  C   1   卡费单号
    'deposit_no  C   1   预交单号
    'cardfee_sign    N   1   是否退卡费:1-是退卡费;0-不退卡费
    'mrbkfee_sign N   1   是否退病历费:1-退病历费;0-不退病历费
    'operator_name   C   1   操作员姓名
    'operator_code   C   1   操作员编号
    'del_time    C   1   退费时间:yyyy-mm-dd hh:mi:ss
    'balance_info    C       只存在一条数据
    '    moeny   N   1   退款金额
    '    blnc_mode    C   1   结算方式
    '    blnc_no C   1   结算号码
    '    memo    C   1   摘要
    '    cardtype_id N   1   卡类别id
    '    consumer_no N   1   结算卡序号：即卡消费接口目录.编号
    '    cardno  C   1   卡号
    '    swapno  C   1   交易流水号
    '    swapmemo    C   1   交易说明
    '    cprtion_unit    C   1   合作单位
    '    relation_id N   1   关联交易ID

    strJson = "": strJsonTemp = ""
    strJson = strJson & "" & GetJsonNodeString("oper_fun", int操作状态, Json_num)
    For i = 1 To cllDelFeeData.Count
        varTemp = cllDelFeeData(i)
        Select Case UCase(varTemp(0))
        Case "卡费单号"
            strJson = strJson & "," & GetJsonNodeString("cardfee_no", Trim(varTemp(1)), Json_Text)
        Case "预交单号"
            strJson = strJson & "," & GetJsonNodeString("deposit_no", Trim(varTemp(1)), Json_Text)
        Case "是否退卡费"
            strJson = strJson & "," & GetJsonNodeString("cardfee_sign", Val(varTemp(1)), Json_num)
        Case "是否退病历费"
            strJson = strJson & "," & GetJsonNodeString("mrbkfee_sign", Val(varTemp(1)), Json_num)
        Case "操作员姓名"
            strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
        Case "操作员编号"
            strJson = strJson & "," & GetJsonNodeString("operator_code", Trim(varTemp(1)), Json_Text)
        Case "退费时间"
            strJson = strJson & "," & GetJsonNodeString("del_time", Trim(varTemp(1)), Json_Text)
        Case "结算信息"
            Set cllTemp = varTemp(1)
            If Not cllTemp Is Nothing Then
                For j = 1 To cllTemp.Count
                     varTemp = cllTemp(j)
                    Select Case UCase(varTemp(0))
                    Case "退款金额"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("moeny", Val(varTemp(1)), Json_num)
                    Case "结算方式"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("blnc_mode", Trim(varTemp(1)), Json_Text)
                    Case "结算号码"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("blnc_no", Trim(varTemp(1)), Json_Text)
                    Case "卡类别ID"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num)
                    Case "结算卡序号"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("consumer_no", Val(varTemp(1)), Json_num)
                    Case "支付卡号"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
                    Case "交易流水号"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swapno", Trim(varTemp(1)), Json_Text)
                    Case "交易说明"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swapmemo", Trim(varTemp(1)), Json_Text)
                    Case "合作单位"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cprtion_unit", Trim(varTemp(1)), Json_Text)
                    Case "关联交易ID"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("relation_id", Val(varTemp(1)), Json_num)
                    End Select
                Next
            
            End If
        End Select
    Next
    If strJsonTemp <> "" Then
        strJsonTemp = Mid(strJsonTemp, 2)
        strJsonTemp = "," & GetNodeString("balance_info") & ":{" & strJsonTemp & "}"
    End If
    strJson = strJson & strJsonTemp
    If strJson = "" Then Exit Function
     
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    
    strServiceName = "Zl_Exsesvr_Delcardfeeinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then
        Exit Function
    End If
    '    output
    '       code    C   1   应答码：0-失败；1-成功
    '       message C   1   应答消息：失败时返回具体的错误信息
    '       deposit_id  N   1   预交ID
    '       balance_id  N   1   结帐ID
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        mstrErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If mstrErrMsg = "" Then
            mstrErrMsg = "费用处理失败，请检查！"
        End If
        If blnShowErrMsg Then MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng预交ID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.deposit_id")))
    lng结帐ID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.balance_id")))
    Zl_Exsesvr_DelCardfeeInfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlUpdateAdviceChargeTag(ByVal bytBillType As Byte, _
    ByVal cllAdviceDatas As Collection, Optional ByVal cllFeeDetails As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更改病人医嘱数据的计费状态及删除附费数据
    '入参:
    '   bytBillType 单据类型：1-收费单;2-记帐单;3-自动记帐单
    '   cllAdviceDatas(collect)-数据集，格式如下
    '     |-cllAdviceData(collect)每行明细数据集
    '        |-成员(医嘱ID,发送号(N),费用单号(N),计费状态,删除附费(N))
    '   cllFeeDetails(collect)费用明细集，仅门诊退费时传入，用于标记 医嘱执行计价.执行状态 为 2-已退费
    '     |-cllFeeDetail(collect)每行明细数据
    '        |-成员(医嘱ID,费用单号,收费明细ID)
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService
    
    On Error GoTo ErrHandler
    If cllAdviceDatas.Count = 0 Then ZlUpdateAdviceChargeTag = True: Exit Function
    
    If GetService(objService) = False Then Exit Function
    ZlUpdateAdviceChargeTag = objService.ZlCisSvr_UpdAdviceChargeTag(bytBillType, cllAdviceDatas, cllFeeDetails)
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function




Public Function zlExseSvr_GetChargeOffApply(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByRef blnExists_Out As Boolean, ByRef rsApplyData_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id,获取病人的销帐申请请数据
    '入参:lng病人ID-病人id
    '出参:blnExists_Out- 是否存在销帐申请
    '     rsApplyData_Out-销帐申请数据
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, clldata As Collection, cllTemp As Collection
    Dim strJson As String, i As Long
    Dim intMode As Integer, strMessage As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'input
    '    pati_id N   1   病人id
    '    pati_pageid N   1   主页id
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_ExseSvr_GetChargeOffApply", strJson, , "", mlngModule) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    exists  N   1   是否存在:1-存在;0-不存在
    '    apply_list[]    C       申请列表
    '        fee_no  C   1   费用单据号
    '        fitem_name  C   1   收费项目名称
    '        audit_dept_name C   1   审核部分名称


    intMode = Val(Nvl(objServiceCall.GetJsonNodeValue("output.code")))
    If intMode > 0 Then
        strMessage = Nvl(objServiceCall.GetJsonNodeValue("output.message"))
        Call MsgBox(strMessage, vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
        Exit Function
    End If
    blnExists_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.exists")) = 1)
    Set clldata = objServiceCall.GetJsonListValue("output.apply_list")
    
    Set rsApplyData_Out = New ADODB.Recordset
    With rsApplyData_Out.Fields
        .Append "NO", adLongVarChar, 50, adFldIsNullable
        .Append "项目名称", adLongVarChar, 200, adFldIsNullable
        .Append "审核科室", adLongVarChar, 100, adFldIsNullable
    End With
    rsApplyData_Out.CursorLocation = adUseClient
    rsApplyData_Out.LockType = adLockOptimistic
    rsApplyData_Out.CursorType = adOpenStatic
    rsApplyData_Out.Open
    zlExseSvr_GetChargeOffApply = True
            
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsApplyData_Out
            .AddNew
                !NO = Nvl(cllTemp("_fee_no"))
                !项目名称 = Nvl(cllTemp("_fitem_name"))
                !审核科室 = Nvl(cllTemp("_audit_dept_name"))
            .Update
        End With
    Next
    zlExseSvr_GetChargeOffApply = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlExseSvr_GetNeedAuditItems(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng项目id As Long, _
    ByRef rsAuditItems_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的费用审批项目
    '入参:lng病人ID-
    '     lng项目id=收费项目id
    '出参: rsAuditItems_Out-销帐申请数据
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, clldata As Collection, cllTemp As Collection
    Dim strJson As String, i As Long
    Dim intMode As Integer, strMessage As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'input
    '    pati_id N   1   病人id
    '    pati_pageid N   1   主页id
    '    fitem_id    N   1   收费项目id

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng项目id, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_ExseSvr_GetNeedAuditItems", strJson, , "", mlngModule) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    item_list
    '        fitem_id N   1   项目ID
    '        limit_quantity  N   1   使用限量
    '        used_quantity   N   1   已用数量
    '        avail_quantity  N   1   可用数量


    intMode = Val(Nvl(objServiceCall.GetJsonNodeValue("output.code")))
    If intMode > 0 Then
        strMessage = Nvl(objServiceCall.GetJsonNodeValue("output.message"))
        Call MsgBox(strMessage, vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
        Exit Function
    End If
     
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    Set rsAuditItems_Out = New ADODB.Recordset
    With rsAuditItems_Out.Fields
        .Append "项目ID", adBigInt, , adFldIsNullable
        .Append "使用限量", adDouble, , adFldIsNullable
        .Append "已用数量", adDouble, , adFldIsNullable
        .Append "可用数量", adDouble, , adFldIsNullable
    End With
    
    rsAuditItems_Out.CursorLocation = adUseClient
    rsAuditItems_Out.LockType = adLockOptimistic
    rsAuditItems_Out.CursorType = adOpenStatic
    rsAuditItems_Out.Open
    zlExseSvr_GetNeedAuditItems = True
            
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsAuditItems_Out
            .AddNew
                !项目ID = Nvl(cllTemp("_fitem_id"))
                !使用限量 = Nvl(cllTemp("_limit_quantity"))
                !已用数量 = Nvl(cllTemp("_used_quantity"))
                !可用数量 = Nvl(cllTemp("_avail_quantity"))
            .Update
        End With
    Next
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlExsesvr_ExistsPriceBill(ByVal intTYPE As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str医嘱IDs As String, ByRef blnExists_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id或医嘱ID,判断该病人或医嘱是否存在划价单
    '入参:lng病人ID-病人id
    '     intType-1-收费单,2-记帐单
    '出参:blnExists_Out- 是否存在划价单
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, clldata As Collection, cllTemp As Collection
    Dim strJson As String, i As Long
    Dim intMode As Integer, strMessage As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '     入参 json
    '    input
    '        pati_id N   1   病人id
    '        pati_pageid N   1   主页id
    '        advice_ids  C       医嘱id:多个用逗号
    '        billtype    N   1   单据类型:1-收费划价单;2-记帐划价单

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("billtype", intTYPE, Json_num)
    strJson = strJson & "," & GetJsonNodeString("advice_ids", str医嘱IDs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Existspricebill", strJson, , "", mlngModule) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    exists  N   1   是否存在:1-存在;0-不存在
  
    intMode = Val(Nvl(objServiceCall.GetJsonNodeValue("output.code")))
    If intMode > 0 Then
        strMessage = Nvl(objServiceCall.GetJsonNodeValue("output.message"))
        Call MsgBox(strMessage, vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
        Exit Function
    End If
    blnExists_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.exists")) = 1)
    zlExsesvr_ExistsPriceBill = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlExseSvr_GetRemainMoney(ByVal bytTYPE As Byte, ByVal lng病人ID As Long, ByRef objPatiFee As clsPatiFeeinfor, _
    Optional ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的相关费用信息
    '入参:bytType-0-门诊;1-住院
    '     lng病人ID-病人ID
     '出参:
    '返回:获取成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, strJson As String
    Dim objService As clsService
    
    On Error GoTo errHandle
    Set objPatiFee = New clsPatiFeeinfor
    
    If bytTYPE = 1 Then
        If lng主页ID = 0 Then '获取最后一次主页ID
            If GetService(objService) = False Then Exit Function
            If objService.ZlCissvr_GetPatiMaxPageID(lng病人ID, lng主页ID) = False Then Exit Function
        End If
    End If
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Exsesvr_Getremainmoney
    '  --获取病人费用余额
    '  --input      获取病人费用余额
    '  --  pati_id                 N  1  病人ID
    '  --  pati_pageid             N  1  主页ID
    '  --  insure_account_balance  N  1  医保账户余额
    '  --  query_type              N  0  查询方式；0-按单病人查询，1-批量查询病人余额，2-批量查询，担保额和适用病人信息
    '  --  pati_ids                C  0  批量查询病人关键信息拼串，两种格式：1-病人ID1:主页ID1,病人ID2:主页ID2,....；2-病人ID1,病人ID2,....
    '  --  fee_source              N  1  费用来源：0-不区分，1-门诊，2-住院；查询方式=1且仅按病人ID查询时有效
    '  --output
    '  --  code                    C  1  应答码：0-失败；1-成功
    '  --  message                 C  1  应答消息
    '  --  remain_money            N     剩余款
    '  --  guarantee_money         N     担保额
    '  --  expected_money          N     预结费用
    '  --  prepay_money            N  0  预交余额
    '  --  nobalance_money         N  0  未结费用金额
    '  --  item_list[]当传入批量病人信息时才返回，该列表可以不返回
    '  --       pati_id            N 1 病人id
    '  --       pati_pageid        N 1 主页id
    '  --       prepay_money       N 0 预交余额
    '  --       nobalance_money    N 0 未结费用金额
    '  --       remain_money       N 1 剩余款
    '  --       guarantee_money    N 1 担保额
    '  --       pati_type          C 1 适用病人
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("query_type", 0, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Getremainmoney", strJson) = False Then Exit Function
    With objPatiFee
        .剩余款 = Val(Nvl(objServiceCall.GetJsonNodeValue("output.remain_money")))
        .担保额 = Val(Nvl(objServiceCall.GetJsonNodeValue("output.guarantee_money")))
        .预结费用 = Val(Nvl(objServiceCall.GetJsonNodeValue("output.expected_money")))
        .预交余额 = Val(Nvl(objServiceCall.GetJsonNodeValue("output.prepay_money")))
        .未结费用 = Val(Nvl(objServiceCall.GetJsonNodeValue("output.nobalance_money")))
    End With
    zlExseSvr_GetRemainMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExseSvr_GetRemainMoneyByBatch(ByVal str病人IDs As String, ByVal byt费用来源 As Byte, _
    ByRef cllPatiFee As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:批量获取病人的相关费用信息
    '入参:
    '       str病人IDs-病人ID1,病人ID2,....
    '       byt费用来源-0-不区分，1-门诊，2-住院
    '出参:
    '       cllPatiFee-集合，成员：clsPatiFeeinfor=Key(_病人ID)
    '返回:获取成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, strJson As String, i As Long
    Dim clldata As Collection, cllTemp As Collection, clsPatiFeeInfo As clsPatiFeeinfor
    
    On Error GoTo errHandle
    Set cllPatiFee = New Collection
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Exsesvr_Getremainmoney
    '  --获取病人费用余额
    '  --input      获取病人费用余额
    '  --  pati_id                 N  1  病人ID
    '  --  pati_pageid             N  1  主页ID
    '  --  insure_account_balance  N  1  医保账户余额
    '  --  query_type              N  0  查询方式；0-按单病人查询，1-批量查询病人余额，2-批量查询，担保额和适用病人信息
    '  --  pati_ids                C  0  批量查询病人关键信息拼串，两种格式：1-病人ID1:主页ID1,病人ID2:主页ID2,....；2-病人ID1,病人ID2,....
    '  --  fee_source              N  1  费用来源：0-不区分，1-门诊，2-住院；查询方式=1且仅按病人ID查询时有效
    '  --output
    '  --  code                    C  1  应答码：0-失败；1-成功
    '  --  message                 C  1  应答消息
    '  --  remain_money            N     剩余款
    '  --  guarantee_money         N     担保额
    '  --  expected_money          N     预结费用
    '  --  prepay_money            N  0  预交余额
    '  --  nobalance_money         N  0  未结费用金额
    '  --  item_list[]当传入批量病人信息时才返回，该列表可以不返回
    '  --       pati_id            N 1 病人id
    '  --       pati_pageid        N 1 主页id
    '  --       prepay_money       N 0 预交余额
    '  --       nobalance_money    N 0 未结费用金额
    '  --       remain_money       N 1 剩余款
    '  --       guarantee_money    N 1 担保额
    '  --       pati_type          C 1 适用病人
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_ids", str病人IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("fee_source", byt费用来源, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Getremainmoney", strJson) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        Set clsPatiFeeInfo = New clsPatiFeeinfor
        With clsPatiFeeInfo
            .病人ID = Val(Nvl(cllTemp("_pati_id")))
            .预交余额 = Val(Nvl(cllTemp("_prepay_money")))
            .未结费用 = Val(Nvl(cllTemp("_nobalance_money")))
            .剩余款 = Val(Nvl(cllTemp("_remain_money")))
            cllPatiFee.Add clsPatiFeeInfo, "_" & .病人ID
        End With
    Next
    zlExseSvr_GetRemainMoneyByBatch = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExsesvr_GetPatiTotalMoney(ByVal byt来源 As Byte, ByVal lng病人ID As Long, ByVal str医嘱IDs As String, ByVal byt划价统计方式 As Byte, _
    ByRef dbl应收合计_Out As Double, ByRef dbl实收合计_Out As Double, _
    Optional ByVal lng就诊ID As Long, Optional ByVal bln当日费用 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的应收及实收金额
    '入参:bytType-0-所有;1-门诊;2-住院
    '     lng病人ID-病人ID
    '     byt划价统计方式:0-不限制;1-不含划价单;2-仅统计划价单
    '     str医嘱IDs-医嘱ID,多个用逗号(传入医嘱时，其他条件无效)
    '出参:dbl应收合计_Out
    '     dbl实收合计_Out
    '返回:获取成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim objService As clsService
    
    On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'input
    '    pati_source N   1   病人来源:0-所有;1-门诊;2-住院
    '    pati_id N   1   病人ID
    '    visit_id    N       就诊ID:住院时，传入主页id,门诊暂传NULL
    '    today_fee   N       当日费用:1-是的;0-不限制
    '    price_tag   N       划价标志:0-不限制;1-不含划价单;2-仅统计划价单
    '    advice_ids  C       医嘱ids:多个用逗号分离,不为空时，表示按医嘱IDs进行统计,其他条件无效
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_source", byt来源, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("visit_id", lng就诊ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("advice_ids", str医嘱IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("price_tag", byt划价统计方式, Json_num)
    strJson = strJson & "," & GetJsonNodeString("today_fee", IIf(bln当日费用, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("zl_Exsesvr_GetPatiTotalMoney", strJson) = False Then Exit Function
    'code    N   1   应答码：0-失败；1-成功
    'message C   1   "应答消息：成功时返回成功信息
 
    'fee_amrcvb  N   1   应收金额
    'fee_ampaib  N   1   实收金额
    dbl应收合计_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.fee_amrcvb")))
    dbl实收合计_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.fee_ampaib")))
    zlExsesvr_GetPatiTotalMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ZlExseSvr_GetOrderFeeState(ByVal str医嘱IDs As String, ByRef bytState_Out As Byte, ByRef bytBillType_Out As Byte, _
    ByRef str未收费期医嘱IDs_out As String, Optional ByVal strNOs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱的收费状态
    '入参:str医嘱IDs-医嘱IDs
    '     strNos-按单据号查找 单
    '
    '出参:bytState_Out-0-未收费,1-完全收费;2-部门收费
    '     bytBillType_Out-0-不存在任何单据;1-收费单;2-记帐单;3-收费和记帐都有
    '     str未收费期医嘱IDs_out-返回未收费医嘱ids
    '返回:获取成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim objService As clsService
    
    On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'input
    '    advice_ids  C   1   医嘱ids
    '    bill_nos    C   1   单据号

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("advice_ids", str医嘱IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("bill_nos", strNOs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_ExseSvr_GetOrderFeeState", strJson) = False Then Exit Function
    ' output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    state   N   1   状态:0-未收费,1-完全收费;2-部门收费
    '    billtype    N   1   单据类型:0-不存在任何单据;1-收费单;2-记帐单;3-收费和记帐都有
    '    advice_ids  C       未收费的医嘱ID:传入医嘱ids时有效

    str未收费期医嘱IDs_out = Trim(Nvl(objServiceCall.GetJsonNodeValue("output.advice_ids")))
    bytBillType_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.billtype")))
    bytState_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.state")))
    ZlExseSvr_GetOrderFeeState = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlExseSvr_GetFeeChargeState(ByVal byt查询方式 As Byte, ByVal strNOs As String, ByRef intState_Out As Integer, _
    Optional ByVal int单据类型 As Integer = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用单据号，获取单据号所对应的收费状态
    '入参:byt查询方式-0-查询收费状态;1-仅查是否存在未收费的
    '     int单据类型-1-收费单据;待以后扩展
    '出参:intState_Out-状态
    '            1.query_mode=0时
    '               状态:-1-不存在单据,0-未收费;1-部分收费/退费;2-全部收费;3-全部退费
    '            2.query_mode=1时
    '               状态:1-存在未收费的;0-不存在未收费
    '返回:获取成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim objService As clsService
    
    On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'input
    '    query_mode  N   1   查询方式:0-查询收费状态;1-仅查是否存在未收费的
    '    bill_nos    C   1   单据号
    '    bill_type   N   1   单据类型:1-收费单据;待以后扩展

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_mode", byt查询方式, Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_type", int单据类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_nos", strNOs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("zl_ExseSvr_GetFeeChargeState", strJson) = False Then Exit Function
    ' output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    state   N   1
    '        1.query_mode=0时
    '           状态:-1-不存在单据,0-未收费;1-部分收费/退费;2-全部收费;3-全部退费
    '        2.query_mode=1时
    '           状态:1-存在未收费的;0-不存在未收费.
    intState_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.state")))
    zlExseSvr_GetFeeChargeState = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExseSvr_GetFeeBalanceState(ByVal byt查询方式 As Byte, ByVal strNOs As String, ByRef intState_Out As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据记帐费用单据号，获取单据号所对应的结帐状态
    '入参:byt查询方式-0-门诊记帐;1-住院记帐
    '
    '出参:intState_Out-状态:-1-不存在记帐单据;0-未结帐;1-部分结帐;2-全部结帐
 
    '返回:获取成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim objService As clsService
    
    On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'input
    '    query_mode  N   1   查询方式:0-门诊记帐;1-住院记帐
    '    bill_nos    C   1   单据号

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_mode", byt查询方式, Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_nos", strNOs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("zl_ExseSvr_GetFeeBalanceState", strJson) = False Then Exit Function
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    state   N   1   状态:-1-不存在记帐单据;0-未结帐;1-部分结帐;2-全部结帐

    intState_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.state")))
    zlExseSvr_GetFeeBalanceState = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExseSvr_GetFeeInfoByBlncId(ByVal lng结帐ID As Long, ByRef rsFeeList_out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID，返回该结帐ID的费用明细数据
    '入参:lng结帐ID-
    '出参: rsFeeList_out-费用明细数据
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, clldata As Collection, cllTemp As Collection
    Dim strJson As String, i As Long
    Dim intMode As Integer, strMessage As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'input
    '    balance_id  N   1   结帐ID

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("balance_id", lng结帐ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_ExseSvr_GetFeeInfoByBlncId", strJson, , "", mlngModule) = False Then Exit Function
    ' output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    fee_list[]  C   1   费用明细列表
    '        fee_no  C   1   费用单据号
    '        serial_num  N   1   序号
    '        receipt_type    C   1   收费类别
    '        fitem_id    N   1   收费细目id
    '        fitem_name  C   1   收费名称
    '        nums    N   1   数次
    '        unit    C   1   计算单位
    '        price   N   1   标准单价
    '        blnc_money  N   1   结帐金额
    '        exedept_name    C   1   执行科室
    '        happen_time     C   1   发生时间:yyyy-mm-dd HH:MM:SS
    intMode = Val(Nvl(objServiceCall.GetJsonNodeValue("output.code")))
    If intMode <> 1 Then
        strMessage = Nvl(objServiceCall.GetJsonNodeValue("output.message"))
        Call MsgBox(strMessage, vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
        Exit Function
    End If
     
 
    
    
    Set clldata = objServiceCall.GetJsonListValue("output.fee_list")
    Set rsFeeList_out = New ADODB.Recordset
    With rsFeeList_out.Fields
        .Append "费用单号", adLongVarChar, 50, adFldIsNullable
        .Append "序号", adBigInt, , adFldIsNullable
        .Append "收费类别", adLongVarChar, 50, adFldIsNullable
        .Append "收费名称", adLongVarChar, 200, adFldIsNullable
        .Append "收费数量", adDouble, , adFldIsNullable
        .Append "结帐金额", adDouble, , adFldIsNullable
        .Append "收费单价", adDouble, , adFldIsNullable
        .Append "计算单位", adLongVarChar, 50, adFldIsNullable
        .Append "执行科室", adLongVarChar, 100, adFldIsNullable
        .Append "发生时间", adLongVarChar, 50, adFldIsNullable
    End With
    
    rsFeeList_out.CursorLocation = adUseClient
    rsFeeList_out.LockType = adLockOptimistic
    rsFeeList_out.CursorType = adOpenStatic
    rsFeeList_out.Open
    
    zlExseSvr_GetFeeInfoByBlncId = True
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsFeeList_out
            .AddNew
            !费用单号 = Nvl(cllTemp("_" & "fee_no"))
            !序号 = Val(Nvl(cllTemp("_" & "serial_num")))
            !收费类别 = Nvl(cllTemp("_" & "receipt_type"))
            !收费名称 = Nvl(cllTemp("_" & "fitem_name"))
            !收费数量 = Val(Nvl(cllTemp("_" & "nums")))
            !结帐金额 = Val(Nvl(cllTemp("_" & "blnc_money")))
            !收费单价 = Val(Nvl(cllTemp("_" & "price")))
            !计算单位 = Nvl(cllTemp("_" & "unit"))
            !执行科室 = Nvl(cllTemp("_" & "exedept_name"))
            !发生时间 = Nvl(cllTemp("_" & "happen_time"))
            .Update
        End With
    Next
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function



Public Function zlExseSvr_GetBalanceInfoById(ByVal lng结帐ID As Long, ByRef rsblncList_out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID，返回该结帐ID的结算明细数据
    '入参:lng结帐ID-
    '出参: rsblncList_out-费用明细数据
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, clldata As Collection, cllTemp As Collection
    Dim strJson As String, i As Long
    Dim intMode As Integer, strMessage As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'input
    '    balance_id  N   1   结帐ID

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("balance_id", lng结帐ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_ExseSvr_GetBalanceInfoById", strJson, , "", mlngModule) = False Then Exit Function
    ' output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    blnc_list[] C   1   结算明细列表
    '        blnc_mode   C   1   结算方式
    '        blnc_no C   1   结算号码
    '        cardtype_id N   1   卡类别id
    '        consumer_no N   1   结算卡序号，即卡消费接口目录.编号
    '        cardno  C   1   卡号
    '        swapno  C   1   交易流水号
    '        swapmemo    C   1   交易说明
    '        memo    C   1   摘要
    '        cprtion_unit    C   1   合作单位
    '        relation_id N   1   关联交易id
 
    intMode = Val(Nvl(objServiceCall.GetJsonNodeValue("output.code")))
    If intMode <> 1 Then
        strMessage = Nvl(objServiceCall.GetJsonNodeValue("output.message"))
        Call MsgBox(strMessage, vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
        Exit Function
    End If
    Set clldata = objServiceCall.GetJsonListValue("output.blnc_list")
    Set rsblncList_out = New ADODB.Recordset
    With rsblncList_out.Fields
        .Append "结算方式", adLongVarChar, 50, adFldIsNullable
        .Append "结算金额", adDouble, , adFldIsNullable
        .Append "结算号码", adLongVarChar, 200, adFldIsNullable
        .Append "卡类别ID", adBigInt, , adFldIsNullable
        .Append "结算卡序号", adBigInt, , adFldIsNullable
        .Append "消费卡", adBigInt, , adFldIsNullable
        .Append "交易流水号", adLongVarChar, 500, adFldIsNullable
        .Append "交易说明", adLongVarChar, 2000, adFldIsNullable
        .Append "刷卡卡号", adLongVarChar, 200, adFldIsNullable
        .Append "合作单位", adLongVarChar, 200, adFldIsNullable
        .Append "摘要", adLongVarChar, 200, adFldIsNullable
        .Append "关联交易ID", adBigInt, , adFldIsNullable
   
    End With
    
    rsblncList_out.CursorLocation = adUseClient
    rsblncList_out.LockType = adLockOptimistic
    rsblncList_out.CursorType = adOpenStatic
    rsblncList_out.Open
    
    zlExseSvr_GetBalanceInfoById = True
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsblncList_out
            .AddNew
                !结算方式 = Nvl(cllTemp("_" & "blnc_mode"))
                !结算金额 = Val(Nvl(cllTemp("_" & "blnc_money")))
                !结算号码 = Nvl(cllTemp("_" & "blnc_no"))
                !卡类别ID = Val(Nvl(cllTemp("_" & "cardtype_id")))
                !结算卡序号 = Val(Nvl(cllTemp("_" & "consumer_no")))
                !消费卡 = IIf(Val(Nvl(!结算卡序号)) <> 0, 1, 0)
                !交易流水号 = Nvl(cllTemp("_" & "swapno"))
                !交易说明 = Nvl(cllTemp("_" & "swapmemo"))
                !刷卡卡号 = Nvl(cllTemp("_" & "cardno"))
                !合作单位 = Nvl(cllTemp("_" & "cprtion_unit"))
                !摘要 = Nvl(cllTemp("_" & "memo"))
                !关联交易ID = Val(Nvl(cllTemp("_" & "relation_id")))
            .Update
        End With
    Next
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function zlExseSvr_GetRgstApptInfo(ByVal strNo As String, ByRef rsApptRgst_out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据挂号单号，获取病人预约信息,并检查是否允许接收
    '入参:strNo-挂号单号
    '出参: rsApptRgst_out-费用明细数据
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, clldata As Collection, cllTemp As Collection
    Dim objService As clsService
    Dim blnContOrder As Boolean
    Dim strJson As String, i As Long
    Dim int原结算模式 As Integer, int有效天数 As Integer
    Dim intMode As Integer, strMessage As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If GetService(objService) = False Then Exit Function
    'input
    '    balance_id  N   1   结帐ID

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("rgst_no", strNo, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_GetRgstApptInfo", strJson, , "", mlngModule) = False Then Exit Function
    ' output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    blnc_list[] C   1   结算明细列表
    '        blnc_mode   C   1   结算方式
    '        blnc_no C   1   结算号码
    '        cardtype_id N   1   卡类别id
    '        consumer_no N   1   结算卡序号，即卡消费接口目录.编号
    '        cardno  C   1   卡号
    '        swapno  C   1   交易流水号
    '        swapmemo    C   1   交易说明
    '        memo    C   1   摘要
    '        cprtion_unit    C   1   合作单位
    '        relation_id N   1   关联交易id
 
    Set clldata = objServiceCall.GetJsonListValue("output.fee_list")
    int原结算模式 = Val(Nvl(objServiceCall.GetJsonNodeValue("output.close_account_type")))
    
    Set rsApptRgst_out = New ADODB.Recordset
    With rsApptRgst_out.Fields
        .Append "挂号ID", adBigInt, , adFldIsNullable
        .Append "病人ID", adBigInt, , adFldIsNullable
        .Append "门诊号", adLongVarChar, 18, adFldIsNullable
        .Append "姓名", adLongVarChar, 100, adFldIsNullable
        .Append "性别", adLongVarChar, 4, adFldIsNullable
        .Append "年龄", adLongVarChar, 20, adFldIsNullable
        .Append "费别", adLongVarChar, 10, adFldIsNullable
        .Append "付款方式", adLongVarChar, 20, adFldIsNullable
        .Append "号别", adLongVarChar, 5, adFldIsNullable
        .Append "发生时间", adDate, , adFldIsNullable
        .Append "预约时间", adDate, , adFldIsNullable
        .Append "挂号金额", adDouble, , adFldIsNullable
        .Append "执行部门ID", adBigInt, , adFldIsNullable
        .Append "号类", adLongVarChar, 10, adFldIsNullable
        .Append "复诊", adBigInt, , adFldIsNullable
        .Append "急诊", adBigInt, , adFldIsNullable
    End With
    
    With rsApptRgst_out
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsApptRgst_out
            .AddNew
                !挂号ID = Val(Nvl(cllTemp("_" & "rgst_id")))
                !病人ID = Val(Nvl(cllTemp("_" & "pati_id")))
                !门诊号 = Nvl(cllTemp("_" & "outpno"))
                !姓名 = Nvl(cllTemp("_" & "pati_name"))
                !性别 = Nvl(cllTemp("_" & "pati_sex"))
                !年龄 = Nvl(cllTemp("_" & "pati_age"))
                !费别 = Nvl(cllTemp("_" & "fee_category"))
                !付款方式 = Nvl(cllTemp("_" & "mdlpay_mode_name"))
                !号别 = Nvl(cllTemp("_" & "num_category"))
                !执行部门ID = Val(Nvl(cllTemp("_" & "exe_deptid")))
                !发生时间 = Nvl(cllTemp("_" & "happen_time"))
                !预约时间 = Nvl(cllTemp("_" & "appt_time"))
                !挂号金额 = Val(Nvl(cllTemp("_" & "fee_ampaid")))
                !号类 = Nvl(cllTemp("_" & "outptyp_name"))
                !复诊 = Val(Nvl(cllTemp("_" & "revst_sign")))
                !急诊 = Val(Nvl(cllTemp("_" & "emg_sign")))
            .Update
        End With
    Next
    
    If int原结算模式 = 1 Then
        int有效天数 = Val(Left(gobjDatabase.GetPara(21, glngSys, , "01") & "1", 1))
        If objService.ZlCISSvr_ChkOutPatiExistOrder(Val(Nvl(cllTemp("_" & "rgst_id"))), int有效天数, blnContOrder) = False Then Exit Function
        If blnContOrder Then
            MsgBox "当前病人的就诊模式为先诊疗后结算,不允许调整该病人的就诊模式!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    zlExseSvr_GetRgstApptInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlExseSvr_GetRgstInfo(ByVal strNo As String, ByRef rsApptRgst_out As ADODB.Recordset, _
                Optional ByVal bln预约接收 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据挂号单号，获取病人预约信息,并检查是否允许接收
    '入参:strNo-挂号单号
    '出参: rsApptRgst_out-费用明细数据
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, clldata As Collection, cllTemp As Collection
    Dim objService As clsService
    Dim blnContOrder As Boolean
    Dim strJson As String, i As Long
    Dim int原结算模式 As Integer, int有效天数 As Integer
    Dim intMode As Integer, strMessage As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If GetService(objService) = False Then Exit Function
    'input
    '    balance_id  N   1   结帐ID

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("rgst_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("appt_recv", IIf(bln预约接收, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_GetRgstInfo", strJson, , "", mlngModule) = False Then Exit Function
    ' output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '    blnc_list[] C   1   结算明细列表
    '        blnc_mode   C   1   结算方式
    '        blnc_no C   1   结算号码
    '        cardtype_id N   1   卡类别id
    '        consumer_no N   1   结算卡序号，即卡消费接口目录.编号
    '        cardno  C   1   卡号
    '        swapno  C   1   交易流水号
    '        swapmemo    C   1   交易说明
    '        memo    C   1   摘要
    '        cprtion_unit    C   1   合作单位
    '        relation_id N   1   关联交易id
 
    Set clldata = objServiceCall.GetJsonListValue("output.fee_list")
    
    Set rsApptRgst_out = New ADODB.Recordset
    With rsApptRgst_out.Fields
        .Append "挂号ID", adBigInt, , adFldIsNullable
        .Append "病人ID", adBigInt, , adFldIsNullable
        .Append "门诊号", adLongVarChar, 18, adFldIsNullable
        .Append "姓名", adLongVarChar, 100, adFldIsNullable
        .Append "性别", adLongVarChar, 4, adFldIsNullable
        .Append "年龄", adLongVarChar, 20, adFldIsNullable
        .Append "费别", adLongVarChar, 10, adFldIsNullable
        .Append "付款方式", adLongVarChar, 20, adFldIsNullable
        .Append "号别", adLongVarChar, 5, adFldIsNullable
        .Append "发生时间", adDate, , adFldIsNullable
        .Append "预约时间", adDate, , adFldIsNullable
        .Append "挂号金额", adDouble, , adFldIsNullable
        .Append "执行部门ID", adBigInt, , adFldIsNullable
        .Append "号类", adLongVarChar, 10, adFldIsNullable
        .Append "复诊", adBigInt, , adFldIsNullable
        .Append "急诊", adBigInt, , adFldIsNullable
    End With
    
    With rsApptRgst_out
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsApptRgst_out
            .AddNew
                !挂号ID = Val(Nvl(cllTemp("_" & "rgst_id")))
                !病人ID = Val(Nvl(cllTemp("_" & "pati_id")))
                !门诊号 = Nvl(cllTemp("_" & "outpatient_num"))
                !姓名 = Nvl(cllTemp("_" & "pati_name"))
                !性别 = Nvl(cllTemp("_" & "pati_sex"))
                !年龄 = Nvl(cllTemp("_" & "pati_age"))
                !费别 = Nvl(cllTemp("_" & "fee_category"))
                !付款方式 = Nvl(cllTemp("_" & "mdlpay_mode_name"))
                !号别 = Nvl(cllTemp("_" & "num_category"))
                !执行部门ID = Val(Nvl(cllTemp("_" & "exe_deptid")))
                !发生时间 = Nvl(cllTemp("_" & "happen_time"))
                !预约时间 = Nvl(cllTemp("_" & "appt_time"))
                !挂号金额 = Val(Nvl(cllTemp("_" & "fee_ampaid")))
                !号类 = Nvl(cllTemp("_" & "outptyp_name"))
                !复诊 = Val(Nvl(cllTemp("_" & "revst_sign")))
                !急诊 = Val(Nvl(cllTemp("_" & "emg_sign")))
            .Update
        End With
    Next
    
    zlExseSvr_GetRgstInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function zlExseSvr_RgstApptReceive(ByVal strNo As String, ByVal bln挂号划价 As Boolean, _
                ByVal str接诊诊室 As String, ByVal str接收时间 As String, _
                ByVal lng病人ID As Long, ByVal str门诊号 As String, ByVal bln在院 As Boolean, _
                ByVal lng结帐ID As Long, ByVal lng关联ID As Long, _
                Optional ByVal str冲预交病人IDs As String, Optional ByVal cllBalance As Collection, _
                Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 医生站预约挂号接收
    ' 入参 : cllBalance(Collect)三方卡交易信息
    '           |-Array:成员(支付金额,结算方式,卡类别ID,支付卡号)
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/12/30 09:57
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim i As Integer, intReturn As Integer
    Dim strJson As String, strData As String
    
    On Error GoTo errHand
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
'    --input
    '    --rgst_no           C   1   挂号单号
    '    --outp_room_name    C   1   接诊诊室
    '    --recv_time         C   1   接收时间
    '    --pati_id           N   1   病人id
    '    --outpatient_num    C   1   门诊号
    '    --prepay_pati_ids   C       冲预交病人ids
    '    --balance_info[]    C       结算信息列表
        '    --blnc_money    N   1   结帐金额
        '    --blnc_mode     C   1   结算方式名称
        '    --cardtype_id   N   1   卡类别id
        '    --pay_cardno    C   1   支付卡号

'    --output
    '    --code              C   1   应答码：0-失败；1-成功
    '    --message           C   1   应答消息：失败时返回具体的错误信息
    '    --blnc_id           N       结帐ID：三方卡支付时返回
    '    --relation_id       N       关联交易ID：三方卡支付时返回
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("reg_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("outp_room_name", str接诊诊室, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("recv_time", str接收时间, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("outpatient_num", str门诊号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_inhospital", IIf(bln在院, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("blnc_id", lng结帐ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("relation_id", lng关联ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pricing_sign", IIf(bln挂号划价, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("operator_code", UserInfo.编号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_name", UserInfo.姓名, Json_Text)
    
    If str冲预交病人IDs <> "" Then
        strJson = strJson & "," & GetJsonNodeString("prepay_pati_ids", str冲预交病人IDs, Json_Text)
    End If
    
    If Not cllBalance Is Nothing Then
        For i = 1 To cllBalance.Count
            If IsArray(cllBalance(i)) Then
                Select Case cllBalance(i)(0)
                    Case "支付金额"
                        strData = strData & "," & GetJsonNodeString("blnc_money", cllBalance(i)(1), Json_num)
                    Case "结算方式"
                        strData = strData & "," & GetJsonNodeString("blnc_mode", cllBalance(i)(1), Json_Text)
                    Case "卡类别ID"
                        strData = strData & "," & GetJsonNodeString("cardtype_id", cllBalance(i)(1), Json_num)
                    Case "支付卡号"
                        strData = strData & "," & GetJsonNodeString("pay_cardno", cllBalance(i)(1), Json_Text)
                End Select
            End If
        Next
        If strData <> "" Then
            strJson = strJson & ",""balance_info"":{" & Mid(strData, 2) & "}"
        End If
    End If
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_ExseSvr_RgstApptReceive", strJson, , "", mlngModule, False) = False Then Exit Function
    intReturn = Val(Nvl(objServiceCall.GetJsonNodeValue("output.code")))
    If intReturn = 0 Then
        strErrMsg = Nvl(objServiceCall.GetJsonNodeValue("output.message"))
        Exit Function
    End If
    
    zlExseSvr_RgstApptReceive = True
    Exit Function
errHand:
    strErrMsg = Err.Description
End Function

Public Function zlExseSvr_UpdRgstBalanceInfo(ByVal strBalance As String, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 医生站预约挂号接收-更新三方支付信息
    ' 入参 : strBalance-Json串
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/12/30 09:57
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim i As Integer, intReturn As Integer
    Dim strJson As String, strData As String
    
    On Error GoTo errHand
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
'    --input
'    --  rgst_no            C  1 挂号no
'    --  blnc_id            N  1 结帐ID
'    --  relation_id        N  1 关联交易ID
'    --  pati_inhospital    N  1 在院病人
'    --  totalmoney         N  1 支付总金额
'    --  cardtype_id        N  1 支付卡类别ID
'
'    --  balance_list[]
'    --     blnc_mode       N  1 结算方式
'    --     swapmoney       C  1 结算金额
'    --     swapno          C  1 交易流水号
'    --     swapmemo        C  1 交易说明
'    --     blnc_no         C  1 结算号码
'    --     blnc_memo       C  1 结算摘要
'    --     card_no         C  1 支付卡号
'    --     cardtype_id     N    卡类别ID
'    --  otherswap_list[]   C    其他交易信息
'    --     swap_name       C  1  交易名称
'    --     swap_note       C  1  交易内容
'
'    --  output
'    --    code             N 1 应答吗：0-失败；1-成功
'    --    message          C 1 应答消息：失败时返回具体的错误信息
  
    
    If objServiceCall.CallService("zl_ExseSvr_UpdRgstBalanceInfo", strBalance, , "", mlngModule, False) = False Then Exit Function
    intReturn = Val(Nvl(objServiceCall.GetJsonNodeValue("output.code")))
    If intReturn = 0 Then
        strErrMsg = Nvl(objServiceCall.GetJsonNodeValue("output.message"))
        Exit Function
    End If
    
    zlExseSvr_UpdRgstBalanceInfo = True
    Exit Function
errHand:
    strErrMsg = Err.Description
End Function

Public Function Zl_Exsesvr_Updatedepositinvinf(ByVal strNo As String, ByVal lng领用ID As Long, ByVal str发票号 As String, _
     ByVal str使用人 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新病人预交记录的发票信息
    '入参:strNO-预交单据号
    '       lng领用ID-领用id
    '       str发票号-发票号
    '出参:blnExists_Out- 是否存在划价单
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, clldata As Collection, cllTemp As Collection
    Dim strJson As String, i As Long
    Dim intMode As Integer, strMessage As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '     入参 json
    '   --     fun_oper          N 1 操作类型:1-发出;2-重打；3-补打;4-红票打印
    '  --      deposit_no        C 1 预交单号
    '  --      recv_id              N 1 领用id
    '  --      inv_no               C 1 当前发票号或开始使用发票号
    '  --      inv_usenums     N 1 发票使用数量
    '  --      use_time           C 1 票据使用时间:yyyy-mm-dd hh24:mi:ss
    '  --      inv_user            C 1 发票使用人

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fun_oper", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("deposit_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("recv_id", lng领用ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_no", str发票号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("inv_usenums", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_user", str使用人, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Updatedepositinvinf", strJson, , "", mlngModule) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：成功时返回成功信息，失败时返回具体的错误信息

    intMode = Val(Nvl(objServiceCall.GetJsonNodeValue("output.code")))
    If intMode = 0 Then
        strMessage = Nvl(objServiceCall.GetJsonNodeValue("output.message"))
        Call MsgBox(strMessage, vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
        Exit Function
    End If

    Zl_Exsesvr_Updatedepositinvinf = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetBarCode(ByVal strNo As String, ByVal bytBillType As Byte, _
    ByRef rsBarCodeData_out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卫材商品条码和内部条码
    '入参:
    '   bytBillType-1-收费单;2-记帐单;3-自动记帐单;
    '   strNo-单据号
    '出参:
    '   rsBarCodeData_out-卫材条码记录集
    '返回:成功返回true,否则返回False
    '编制:焦博
    '日期:2020-02-28 16:54:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService
    
    On Error GoTo errHandle
    
    '获取费用服务接口
    If GetService(objService) = False Then Exit Function
      
    '获取卫材条码数据
    If objService.zlStuffSvr_GetExecutedNum(strNo, bytBillType, rsBarCodeData_out, True) = False Then Exit Function
    
    zlGetBarCode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

