Attribute VB_Name = "mdlPubServiceExse"

Option Explicit

'*********************************************************************************************************************************************
'功能:所有涉及调用费用的相关服务
'接口说明:
'    0.zl_ExseSvr_GetNextNo-获取下一张单据号
'    1.Zl_Exsesvr_Updatecardinvoice-卡费按门诊医疗票据使用时进行门诊医疗票据的相关更新操作。
'    2.zl_ExseSvr_GetReceiveInvoice-获取票据领用信息
'    3.zl_ExseSvr_GetNextInvoice-根据当前发票号及票据使用明细，获取一个有效的发票号
'    4.zl_ExseSvr_GetFullNo-自动补齐单据号(费用)
'    5.Zl_Exsesvr_GetBillOperControls-获取单据操作控制数据
'    6.zl_ExseSvr_GetBillTotalMoney-获取单据的应收或实收总额
'    7.zl_ExseSvr_GetBillInfoByNo-获取单据的应收或实收总额
'    8.zl_ExseSvr_GetPatiInvoiceClass-获取票据使用类别
'    9.zl_ExseSvr_InvoiceClassUsed-检查指定票种是否启用了使用类别
'    10.zl_ExseSvr_Actualmoney-根据费别计算打折金额
'    11.Zl_Exsesvr_GetDepositDetail-查询指定病人的预交收支明细情况
'    12.zl_ExseSvr_BillInHistory-检查指定费用单据是否进入后备表空间
'    13.zl_ExseSvr_GetCardFeeInfoByNo- 根据指定的卡费单据，获取费用及结算及预交信息
'    14.Zl_Exsesvr_CheckCardnoIsUsed-检查指定卡号是否在票据使用明细中存在，存在时，返回领用ID
'    15.Zl_Exsesvr_Addcardfeeinfo-增加卡费及预交数据
'    16.Zl_Exsesvr_UpdCardFeeBlncInfo:完成卡费收费结算
'    17.zl_ExseSvr_GetBillStatuByNo-获取指定收费单据的异常、收费及结帐等状态.
'    18.zl_ExseSvr_DelCardFeeCheck-退卡及退病历费合法性检查
'    19.Zl_Exsesvr_Delcardfeeinfo:退卡费及预交及病历费
'    20.zl_ExseSvr_GetRelatedTransInfo:根据关联交易ID,获取交易信息
'    21.Zl_Exsesvr_Getbalanceinfo-根据单据等条件获取结算信息
' 二、挂号相关服务
'    1.Zl_Exsesvr_Getapptregisterinfo-根据挂号单获取预约挂号信息

'编制:刘兴洪
'日期:2019*10*31 14:47:18
'*********************************************************************************************************************************************
Public gobjServiceCall As Object
Private mlngErrNum As Long, mstrSource As String, mstrErrMsg As String

 
Private Function GetServiceCall(ByRef objServiceCall_Out As Object, Optional blnShowErrMsg As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共服务对象
    '出参:objServiceCall_Out-返回公共服务对象
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
     
    If Not gobjServiceCall Is Nothing Then Set objServiceCall_Out = gobjServiceCall: GetServiceCall = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
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
    If gobjServiceCall.InitService(gcnOracle, gstrDBUser, glngSys, glngModul) = False Then
        
        Set gobjServiceCall = Nothing: Exit Function
    End If
    Set objServiceCall_Out = gobjServiceCall
    GetServiceCall = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function Zl_Exsesvr_Updatecardinvoice(ByVal int操作类型 As Integer, ByVal str费用单号 As String, _
    ByVal lng领用ID As Long, ByVal str发票号 As String, ByVal str使用人 As String, Optional ByVal str使用时间 As String, _
    Optional ByVal int使用数量 As Integer = 1, Optional ByRef strUseFact As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:卡费按门诊医疗票据使用时进行门诊医疗票据的相关更新操作
    '入参:int操作类型-1-发卡；2-退卡；3-重打；4-补打；5-换卡
    '     str费用单号-费用单据号
    '     str使用时间-格式:yyyy-mm-dd hh24:mi:ss
    '     blnShowErrMsg-是否显示错误信息
    '出参:strUseFact-返回保存后使用的发票号,多个用逗号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
  
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '    input           取病人费用余额
    '        fun_oper    N   1   操作类型:1-发卡；2-退卡；3-重打；4-补打；5-换卡
    '        fee_no  C   1   费用单号
    '        recv_id N   1   领用id
    '        inv_no  C   1   当前发票号或开始使用发票号
    '        inv_usenums N   1   发票使用数量
    '        use_time    C   1   票据使用时间:yyyy-mm-dd hh24:mi:ss
    '        inv_user    C   1   发票使用人
    If str使用时间 = "" Then str使用时间 = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fun_oper", int操作类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_no", str费用单号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("recv_id", lng领用ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_no", str发票号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("inv_usenums", int使用数量, Json_num)
    strJson = strJson & "," & GetJsonNodeString("use_time", str使用时间, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("inv_user", str使用人, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Updatecardinvoice"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   "应答消息：
    '    成功时返回成功信息
    '    失败时返回具体的错误信息 ""
    '    inv_outnos  C   1   门诊医疗票据:使用的门诊医疗票据,多个用逗号返回
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能更新发票使用信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strUseFact = objServiceCall.GetJsonNodeValue("output.inv_outnos")
    Zl_Exsesvr_Updatecardinvoice = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
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
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
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
    If ErrCenter() = 1 Then
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
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
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
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zl_ExseSvr_GetFullNo(ByVal int序号 As Integer, ByVal strInputNo As String, ByVal lng科室ID As Long, ByRef strFullNo_Out As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动补齐单据号(费用)
    '入参:int序号-号码控制表中的序号
    '     strInputNO-输入的单据号
    '     blnShowErrMsg-是否显示错误信息
    '出参:strFullNo_Out-返回完整的单据号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    strFullNo_Out = strInputNo
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetFullNo
    '    input
    '        item_num    N   1   项目序号
    '        input_no    C   1   输入的单据号
    '        dept_id N       科室ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("item_num", int序号, Json_num)
    strJson = strJson & "," & GetJsonNodeString("input_no", strInputNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("dept_id", lng科室ID, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetFullNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '        code    N   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：失败时返回具体的错误信息
    '        full_no C       补齐后的单据号
    '
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能自动补齐单据信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strFullNo_Out = objServiceCall.GetJsonNodeValue("output.full_no")
    zl_ExseSvr_GetFullNo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Exsesvr_GetBillOperControls(ByVal int单据 As Integer, ByVal lng人员ID As Long, ByRef bln存在控制_Out As Boolean, ByRef int时间限制_Out As Integer, _
    ByRef int他人单据_Out As Integer, ByRef dbl金额上限_Out As Double, Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据操作控制数据
    '入参:int单据-单据:1-挂号单据,2-收费单,3-划价单,4-门诊记帐,5-住院记帐,6-预交款,7-结帐单据,8-就诊卡
    '     lng人员ID-人员ID
    '     blnShowErrMsg-是否显示错误信息
    '出参:int时间限制_out-返回时间限制(0(NULL)-不限制,n-n天内)
    '     bln存在控制_out-是否存在单据控制:true-存在，false-不存在
    '     dbl金额上限_Out
    '     int他人单据_out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetBillOperControls
    '    input           病人转病区费用的转入，转出处理
    '        bill_type   N   1   单据类型:1-挂号单据,2-收费单,3-划价单,4-门诊记帐,5-住院记帐,6-预交款,7-结帐单据,8-就诊卡
    '        operator_id N   1   人员ID

    strJson = strJson & "" & GetJsonNodeString("bill_type", int单据, Json_num)
    strJson = strJson & "," & GetJsonNodeString("operator_id", lng人员ID, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Exsesvr_GetBillOperControls"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：失败时返回具体的错误信息
    '    is_exist    N   1   存在控制数据:1-存在;0-不存在
    '    time_limit  N   1   0(NULL)-不限制,n-n天内
    '    other_bill  N   1   是否允许对其它单据进行操作
    '    uplimit_money   N   1   金额上限

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取有效的单据控制信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    bln存在控制_Out = Val(objServiceCall.GetJsonNodeValue("output.is_exist")) = 1
    int时间限制_Out = Val(objServiceCall.GetJsonNodeValue("output.time_limit"))
    int他人单据_Out = Val(objServiceCall.GetJsonNodeValue("output.other_bill"))
    dbl金额上限_Out = Val(objServiceCall.GetJsonNodeValue("output.uplimit_money"))
    Zl_Exsesvr_GetBillOperControls = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zl_ExseSvr_GetBillTotalMoney(ByVal bln门诊 As Boolean, ByVal int记录性质 As Integer, ByVal str单据号 As String, _
    ByRef dbl应收金额_Out As Double, ByRef dbl实收金额_Out As Double, Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long, Optional ByVal lng病人ID As Long, Optional ByVal str记录状态s As String = "0,1") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据的应收或实收总额
    '入参:bln门诊-是否门诊
    '     int记录性质-:1-收费单;3-自动记帐单；2 -记帐记录；4-挂号记录 ;5-就诊卡
    '     lng病人id-多病人单时，病人ID有效
    '     blnShowErrMsg-是否显示错误信息
    '出参:dbl应收金额_Out-应收合计
    '     dbl实收金额_Out-实收合计
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetBillTotalMoney
    '    input
    '        fee_origin  N   1   费用来源:1-门诊;2-住院
    '        bill_type   N   1   单据类型:1-收费单;3-自动记帐单；2 -记帐记录；4-挂号记录 ;5-就诊卡
    '        fee_no  C   1   单据号:费用单据号
    '        pati_id N   1   病人id
    '        rec_status  C       记录状态s:可以多个状态,比如:0,1


    strJson = strJson & "" & GetJsonNodeString("fee_origin", IIf(bln门诊, 1, 2), Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_type", int记录性质, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_no", str单据号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rec_status", str记录状态s, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetBillTotalMoney"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '        code    C   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息： 成功时返回成功信息 ,失败时返回具体的错误信息 ""
    '        fee_amrcvb  N   1   应收金额
    '        fee_ampaib  N   1   实收金额
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取费用单号为" & str单据号 & "的应收或实收总额，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
     
    dbl应收金额_Out = Val(objServiceCall.GetJsonNodeValue("output.fee_amrcvb"))
    dbl实收金额_Out = Val(objServiceCall.GetJsonNodeValue("output.fee_ampaib"))
    zl_ExseSvr_GetBillTotalMoney = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function






Public Function zl_ExseSvr_GetBillInfoByNo(ByVal bln门诊 As Boolean, ByVal int单据类型 As Integer, ByVal str单据号 As String, _
    ByRef str操作员姓名_Out As String, ByRef str登记时间_Out As String, ByRef lng病人id_out As Long, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据的应收或实收总额
    '入参:bln门诊-是否门诊
    '     int单据类型-:1-收费单;3-自动记帐单；2 -记帐记录；4-挂号记录 ;5-就诊卡;-1-结帐单;-2-预交单;-3-补充结算
    '     lng病人id-多病人单时，病人ID有效
    '     blnShowErrMsg-是否显示错误信息
    '出参:str操作员姓名_Out
    '     str登记时间_Out
    '     lng病人id_Out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetBillInfoByNo
    ' input
    '    fee_origin  N   1   费用来源:1-门诊;2-住院 ;bill_type>0时有效
    '    bill_type   N   1   单据类型:1-收费单;3-自动记帐单；2 -记帐记录；4-挂号记录 ;5-就诊卡;-1-结帐单;-2-预交单;-3-补充结算
    '    bill_no C   1   单据号
    
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_origin", IIf(bln门诊, 1, 2), Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_type", int单据类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_no", str单据号, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetBillInfoByNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   "应答消息：
    '    operator_name   C   1   操作员姓名
    '    create_time C   1   登记时间:yyyy-mm-dd hh24:mi:ss
    '    pati_id N   1   病人id

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取费用单号为" & str单据号 & "的登记时间及操作员相关信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    str操作员姓名_Out = Nvl(objServiceCall.GetJsonNodeValue("output.operator_name"))
    str登记时间_Out = Nvl(objServiceCall.GetJsonNodeValue("output.create_time"))
    lng病人id_out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.pati_id")))
    zl_ExseSvr_GetBillInfoByNo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zl_ExseSvr_GetFeeBillByCardNo(ByVal lng病人ID As Long, ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
    ByRef strCardFeeNo_out As String, ByRef strPriceBillNo_out As String, ByRef int收费标志 As Integer, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID及卡号,获取费用的单据信息
    '入参:lng病人ID
    '     lngCardTypeID-卡类别ID
    '     strCardNo-卡号
    '     blnShowErrMsg-是否显示错误信息
    '出参:str操作员姓名_Out
    '     str登记时间_Out
    '     lng病人id_Out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetFeeBillByCardNo
    'input
    '    pati_id N   1   病人id
    '    cardtype_id N   1   卡类别id
    '    cardno  C   1   卡号
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardtype_id", lngCardTypeID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardno", strCardNo, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetFeeBillByCardNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   "应答消息：
    '    operator_name   C   1   操作员姓名
    '    create_time C   1   登记时间:yyyy-mm-dd hh24:mi:ss
    '    pati_id N   1   病人id

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取卡号为" & strCardNo & "的费用单据信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strCardFeeNo_out = Nvl(objServiceCall.GetJsonNodeValue("output.feeno"))
    strPriceBillNo_out = Nvl(objServiceCall.GetJsonNodeValue("output.priceno"))
    int收费标志 = Val(Nvl(objServiceCall.GetJsonNodeValue("output.charge_sign")))
    zl_ExseSvr_GetFeeBillByCardNo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zl_ExseSvr_GetPatiInvoiceClass(ByVal lng病人ID As Long, ByVal lng主页id As Long, ByVal int险类 As Integer, ByRef str使用类别_Out As String, _
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
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetPatiInvoiceClass
    ' input
    '    pati_id N   1   病人id
    '    pati_pageid N   1   主页id
    '    insure_type N   1   险类

    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页id, Json_num)
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
    If ErrCenter() = 1 Then
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
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
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
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_ExseSvr_Actualmoney(ByVal str费别 As String, ByVal lng收费细目id As Long, _
    ByVal lng收入项目id As Long, ByVal dbl应收金额 As Double, ByRef dbl实收金额_Out As Double, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional ByVal dbl数量 As Double, Optional ByVal dbl成本价 As Long, Optional ByVal lng医嘱id As Long) As Boolean
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
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
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
    strJson = strJson & "," & GetJsonNodeString("order_id", lng医嘱id, Json_num, True)
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
    If ErrCenter() = 1 Then
        Resume
    End If
End Function




Public Function Zl_Exsesvr_GetDepositDetail(ByVal lng病人ID As Long, ByVal str开始时间 As String, ByVal str终止时间 As String, _
    ByVal int查询类型 As Integer, ByRef cllDpstDetail_Out As Collection, Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查询指定病人的预交收支明细情况
    '入参:lng病人id-  病人id
    '     str开始时间- 格式:yyyy-mm-dd hh24:mi:ss
    '     str终止时间-格式:yyyy-mm-dd hh24:mi:ss
    '     int查询类型-0-所有,1-门诊;2-住院

    '     blnShowErrMsg-是否显示错误信息
    '     dbl数量\dbl成本价\lng医嘱id:药品及卫材才传入
    '出参:cllDpstDetail_Out-预交收支明细数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Err = 0: On Error GoTo errHandle
    'Zl_Exsesvr_GetDepositDetail
    '    input
    '        pati_id N   1   病人id
    '        begin_time  C   1   开始时间:yyyy-mm-dd hh24:mi:ss
    '        end_time    C   1   终止时间:yyyy-mm-dd hh24:mi:ss
    '        type_sign   N   1   类型标志:0-所有,1-门诊;2-住院
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("begin_time", str开始时间, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("end_time", str终止时间, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("type_sign", int查询类型, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Exsesvr_GetDepositDetail"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '
    '    output
    '        code    C   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：
    '        item_list C
    '            business_type   C   1   业务类型:期初，充值，收费用、结帐等
    '            happen_time C   1   发生时间:yyyy-mm-dd hh24:mi:ss
    '            earlystage  N   1   期初余额
    '            recharge    N   1   本期充值
    '            consume N   1   本期消费

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取指定条件下的预交收支明细数据，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllDpstDetail_Out = objServiceCall.GetJsonListValue("output.item_list")
    Zl_Exsesvr_GetDepositDetail = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function zl_ExseSvr_BillInHistory(ByVal strNO As String, ByVal int单据类型 As Integer, _
    ByVal bln门诊 As Boolean, ByRef bln存在后备表_Out As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定费用单据是否进入后备表空间
    '入参:strNO- 单据号
    '     int单据类型-1-收费单,2-预交单,3-结帐单,4-挂号单,5-就诊卡单据,6-记帐单据;7-自动记帐单
    '     bln门诊-是否门诊
    '     blnShowErrMsg-是否显示错误信息
    '出参:bln存在后备表_Out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_BillInHistory
    '    input
    '            bill_no C   1   单据号
    '            bill_type   C   1   单据类型:1-收费单,2-预交单,3-结帐单,4-挂号单,5-就诊卡单据,6-记帐单据;7-自动记帐单
    '            outpati_flag    N       门诊标志：1-门诊，2-住院
    '

    strJson = strJson & "" & GetJsonNodeString("bill_no", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("bill_type", int单据类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("outpati_flag", IIf(bln门诊, 1, 2), Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_BillInHistory"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    code    N   1   应答码：0-失败；1-成功
    '    message C   1   "应答消息：失败时返回具体的错误信息
    '    exits_history   C   1   存在历史后备表:1-存在;1-不存在


    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能判断是否有效的历史转出数据，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln存在后备表_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.exits_history"))) = 1
    zl_ExseSvr_BillInHistory = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_Exsesvr_BillIsPrintInvoice(ByVal strNO As String, ByVal int单据类型 As Integer, _
    ByVal int票种 As Integer, ByRef bln是否打印_Out As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定的单据是否已经打印了票据
    '入参:strNO- 单据号
    '     int单据类型-1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
    '     int票种-1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    '     blnShowErrMsg-是否显示错误信息
    '出参:bln是否打印_Out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'Zl_Exsesvr_Billisprintinvoice
    '    input
    '        fee_no  C   1   单据号
    '        bill_type   N   1   单据类型:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
    '        inv_type    N   1   票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡


    strJson = strJson & "" & GetJsonNodeString("bill_no", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("bill_type", int单据类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_type", int票种, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Exsesvr_Billisprintinvoice"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    '    printed N   1   是否打印:1-已打印;0-未打印

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能判断是否打印过票据，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln是否打印_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.printed"))) = 1
    zl_Exsesvr_BillIsPrintInvoice = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_ExseSvr_GetCardFeeInfoByNo(ByVal strNO As String, ByVal int查询类型 As Integer, _
    ByRef cllFeeData_out As Collection, ByRef cllPriceBill_out As Collection, ByRef cllBalance_Out As Collection, ByRef cllDeposit_Out As Collection, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional bln是否包含预交 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的卡费单据，获取费用及结算及预交信息
    '入参:strNO- 单据号
    '     int查询类型：0-读取正常单据:1-读取作废单据;2-剩余费用单据
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
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetCardFeeInfoByNo
    '    input
    '       fee_no  C   1   单据号:费用单据号
    '       query_type  N   查询类型：0-读取正常单据:1-读取作废单据;2-剩余费用单据
    '       query_deposit N 1 是否包含预交:1-包含预交信息，0-不包含预交信息
    strJson = strJson & "" & GetJsonNodeString("fee_no", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("query_type", int查询类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("query_deposit", IIf(bln是否包含预交, 1, 0), Json_num)
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
    '        original_money N,1 原始金额,求剩余款数时返回
    '        original_id N 1 原结帐ID:冲销时返回
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
            strErrmsg_Out = "不能获取单据号为【" & strNO & "】的单据信息，请检查！"
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
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_ExseSvr_GetMrbkFeeInfo(ByVal lng病人ID As Long, ByVal str单据号 As String, int记录状态 As Integer, _
     ByRef cllFeeData_out As Collection, Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id获取病历费相关信息
    '入参:lng病人ID-病人ID
    '     str单据号
    '     int记录状态-(1,3)-原始记录,2-销帐记录
    '     blnShowErrMsg-是否显示错误信息
    '出参:cllFeeData_out-返回费用数据集
    '
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Collection
    Dim objServiceCall As Object
    Dim intReturn As Integer
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetMrbkFeeInfo
    '    input
    '        pati_id N   1   病人id
    '        fee_no  C       单据号:病历费所涉及的单据号
    '        rec_status  N   1   记录状态:1-原始记录;2-销帐数据
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rec_status", int记录状态, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_no", str单据号, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetMrbkFeeInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    
    ' output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   "应答消息
    '    fee_list[]  C       费用明细数据
    '        fee_no  C   1   单据号
    '        fee_num N   1   序号
    '        pati_id N   1   病人id
    '        pati_name   C   1   姓名
    '        pati_sex    C   1   性别
    '        pati_age    C   1   年龄
    '        fee_category    C   1   费别
    '        fee_status  N   1   费用状态:1-异常状态;0-正常费用
    '        rec_status  N   1   记录状态:1-正常记录;2-销帐记录;3-补销帐的记录
    '        charge_sign        N       1       收费标志:0-现收;1-记帐;2-划价单
    '        fee_ampaid  N   1   实收金额
    '        happen_time C   1   发生时间:yyyy-mm-dd hh24:mi:ss
    '        operator_name   C   1   操作员姓名
    '        memo    C   1   摘要
    '        pricebill_no    C   1   划价单号
    '        price_charged   N   1   划价已收费:1-划价单已经在收费窗口收费;0-未收费
    '        balance_info  C       结算信息列表
    '            blnc_mode   C   1   结算方式名称
    '            balance_id  N   1   结帐ID：查询作废的单据时为冲销ID
    '            blnc_money  N   1   结帐金额
    '            pay_cardno  N   1   支付卡号
    '            pay_swapno  C   1   交易流水号
    '            pay_swapmemo    C   1   交易说明
    '            relation_id N   1   关联交易id
    '            cardtype_id N   1   卡类别id
    '            consume_card    N   1   是否消费卡:1-是;0-不是
    '            blnc_nature N   1   结算性质:1-现金结算方式,2-其他非医保结算 , 8-结算卡结算 ,9-误差费
    '            blnc_statu  N   1   结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
    '            consume_card_id N   1   消费卡id
    '            blnc_no C   1   结算号码
    '            blnc_memo   C   1   摘要
    '            original_id N   1   原结帐ID:冲销时返回
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取指定病人的病历费,请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set cllFeeData_out = objServiceCall.GetJsonListValue("output.fee_list")
    
    For i = 1 To cllFeeData_out.Count
        Set cllTemp = objServiceCall.GetJsonListValue("output.fee_list[" & i - 1 & "].balance_info")
        If Not cllTemp Is Nothing Then
            Call RemoveCollectionItemFromKey(cllFeeData_out(i), "_balance_info")
            cllFeeData_out(i).Add cllTemp, "_balance_info"
         End If
    Next
    
    
    zl_ExseSvr_GetMrbkFeeInfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardFeeDataFromColl(ByVal strNO As String, ByVal cllCardFee As Collection, _
    ByRef rsCardFee_Out As Recordset, Optional ByRef objBalanceItems_Out As clsBalanceItems, Optional ByRef dblMoney_Out As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据服务返回的集合,按记录集方式返回信息
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
    Dim i As Long, bln记帐 As Boolean
    
    On Error GoTo errHandle
    
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    dblMoney_Out = 0
    Set rsCardFee_Out = New ADODB.Recordset
    With rsCardFee_Out
        If .State = adStateOpen Then .Close
        
        .Fields.Append "单据号", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "费用id", adBigInt, , adFldIsNullable
        .Fields.Append "序号", adBigInt, , adFldIsNullable
        .Fields.Append "病人id", adBigInt, , adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "费别", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "收费项目id", adBigInt, , adFldIsNullable
        .Fields.Append "收入项目id", adBigInt, , adFldIsNullable
        .Fields.Append "数次", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "应收金额", adDouble, , adFldIsNullable
        .Fields.Append "实收金额", adDouble, , adFldIsNullable
        
        .Fields.Append "开单人", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "操作员编号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "操作员姓名", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "登记时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "发生时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "记录状态", adBigInt, , adFldIsNullable
        
        .Fields.Append "是否病历费", adBigInt, , adFldIsNullable
        .Fields.Append "发票号", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "是否记帐", adBigInt, , adFldIsNullable
        .Fields.Append "费用状态", adBigInt, , adFldIsNullable
        .Fields.Append "卡类别ID", adBigInt, , adFldIsNullable
        .Fields.Append "卡号", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "是否挂号发卡", adBigInt, , adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    If cllCardFee Is Nothing Then Exit Function
    '    fee_id  N   1   费用id
    '    fee_num N   1   序号
    '    pati_id N   1   病人id
    '    pati_name   C   1   姓名
    '    pati_sex    C   1   性别
    '    pati_age    C   1   年龄
    '    fee_category    C   1   费别
    '    item_id N   1   收费项目id
    '    income_item_id  N   1   收入项目id
    '    quantity    N   1   数次
    '    fee_amrcvb  N   1   应收金额
    '    fee_ampaid  N   1   实收金额
    '    placer  C   1   开单人
    '    operator_code   C   1   操作员编号
    '    operator_name   C   1   操作员姓名
    '    create_time D   1   登记时间
    '    happen_time D   1   发生时间
    '    rec_status  N   1   记录状态
    '    mrbkfee_sign N   1   是否病历费:1-是病历费;0-不是病历费
    '    invoice_no  N   1   发票号
    '    kpbooks_sign N   1   记帐标志:1-是记帐;0-现收
    '    fee_status   N   1   费用状态:1-异常状态;0-正常费用
    '    cardtype_id N   1   卡类别ID
    '    card_no C   1   卡号
    '    sendcard_reg    N   1   是否挂挂号同步发卡:1-是挂号同时发卡;0-非挂号同时发卡

    For i = 1 To cllCardFee.Count
        Set cllTemp = cllCardFee(i)
        
        If Not bln记帐 Then bln记帐 = Val(Nvl(cllTemp("_kpbooks_sign"))) = 1
        With rsCardFee_Out
            .AddNew
            !单据号 = strNO
            !费用id = Val(Nvl(cllTemp("_fee_id")))
            !序号 = Val(Nvl(cllTemp("_fee_num")))
            !病人ID = Val(Nvl(cllTemp("_pati_id")))
            !姓名 = Nvl(cllTemp("_pati_name"))
            !性别 = Nvl(cllTemp("_pati_sex"))
            !年龄 = Nvl(cllTemp("_pati_age"))
            !费别 = Nvl(cllTemp("_fee_category"))
            !收费项目id = Val(Nvl(cllTemp("_item_id")))
            !收入项目ID = Val(Nvl(cllTemp("_income_item_id")))
            !数次 = Val(Nvl(cllTemp("_quantity")))
            !应收金额 = Val(Nvl(cllTemp("_fee_amrcvb")))
            !实收金额 = Val(Nvl(cllTemp("_fee_ampaid")))
            !开单人 = Nvl(cllTemp("_placer"))
            !操作员编号 = Nvl(cllTemp("_operator_code"))
            !操作员姓名 = Nvl(cllTemp("_operator_name"))
            !登记时间 = Nvl(cllTemp("_create_time"))
            !发生时间 = Nvl(cllTemp("_happen_time"))
            !记录状态 = Val(Nvl(cllTemp("_rec_status")))
            
            !是否病历费 = Val(Nvl(cllTemp("_mrbkfee_sign")))
            !发票号 = Nvl(cllTemp("_invoice_no"))
            !是否记帐 = Val(Nvl(cllTemp("_kpbooks_sign")))
            !费用状态 = Val(Nvl(cllTemp("_fee_status")))
            !卡类别ID = Val(Nvl(cllTemp("_cardtype_id")))
            !卡号 = Nvl(cllTemp("_card_no"))
            !是否挂号发卡 = Val(Nvl(cllTemp("_sendcard_reg")))
            .Update
            dblMoney_Out = RoundEx(dblMoney_Out + Val(Nvl(rsCardFee_Out!实收金额)), 5)
        End With
    Next
    If bln记帐 Then
        objBalanceItems_Out.类型 = gEM_记帐单
    End If
    objBalanceItems_Out.结算金额 = dblMoney_Out
    objBalanceItems_Out.单据号 = strNO
    Set cllTemp = Nothing
    zlGetCardFeeDataFromColl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardFeeBalanceFromColl(ByVal cllCardFeeBalance As Collection, ByRef rsBalance_Out As Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据服务返回的集合,按记录集方式返回结算信息
    '入参:cllCardFeeBalance-当前集合
    '
    '出参:rsBalance_out-返回的卡费结算信息集合
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long
    On Error GoTo errHandle
    
    Set rsBalance_Out = New ADODB.Recordset
    With rsBalance_Out
        If .State = adStateOpen Then .Close
        .Fields.Append "结算方式", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "结帐ID", adBigInt, , adFldIsNullable
        .Fields.Append "结帐金额", adDouble, , adFldIsNullable
        
        .Fields.Append "卡号", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "交易流水号", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "交易说明", adLongVarChar, 4000, adFldIsNullable
        .Fields.Append "关联交易id", adBigInt, , adFldIsNullable
        
        .Fields.Append "卡类别id", adBigInt, , adFldIsNullable
        .Fields.Append "是否消费卡", adSmallInt, , adFldIsNullable
        .Fields.Append "结算性质", adSmallInt, , adFldIsNullable
        .Fields.Append "结算状态", adSmallInt, , adFldIsNullable
        .Fields.Append "消费卡id", adBigInt, , adFldIsNullable
         
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    If cllCardFeeBalance Is Nothing Then Exit Function
    '    blnc_mode   C   1   结算方式名称
    '    balance_id  N   1   结帐ID
    '    blnc_money  N   1   结帐金额
    '    pay_cardno  N   1   支付卡号
    '    pay_swapno  C   1   交易流水号
    '    pay_swapmemo    C   1   交易说明
    '    relation_id N   1   关联交易id
    '    cardtype_id N   1   卡类别id
    '    consume_card    N   1   是否消费卡:1-是;0-不是
    '    blnc_nature N   1   结算性质:1-现金结算方式,2-其他非医保结算 , 8-结算卡结算 ,9-误差费
    '    blnc_statu  N   1   结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
    '    consume_card_id N   1   消费卡id
    For i = 1 To cllCardFeeBalance.Count
        Set cllTemp = cllCardFeeBalance(i)
        With rsBalance_Out
            .AddNew
                !结算方式 = Nvl(cllTemp("_blnc_mode"))
                !结帐id = Val(Nvl(cllTemp("_balance_id")))
                !结帐金额 = Val(Nvl(cllTemp("_blnc_money")))
                !卡号 = Nvl(cllTemp("_pay_cardno"))
                !交易流水号 = Nvl(cllTemp("_pay_swapno"))
                !交易说明 = Nvl(cllTemp("_pay_swapmemo"))
                !关联交易ID = Val(Nvl(cllTemp("_relation_id")))
                !卡类别ID = Val(Nvl(cllTemp("_cardtype_id")))
                !是否消费卡 = Val(Nvl(cllTemp("_consume_card")))
                !结算性质 = Val(Nvl(cllTemp("_blnc_nature")))
                !结算状态 = Val(Nvl(cllTemp("_blnc_statu")))
                !消费卡ID = Val(Nvl(cllTemp("_consume_card_id")))
            .Update
        End With
    Next
    zlGetCardFeeBalanceFromColl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceItemsFromCardFeeColl(ByVal strNO As String, ByVal cllCardFeeBalance As Collection, ByRef objBalanceItems_Out As clsBalanceItems, _
    Optional ByVal bln查看作废 As Boolean, Optional blnDelFee As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据服务返回的集合,按记录集方式返回结算信息
    '入参:cllCardFeeBalance-当前集合
    '     strNo-费用单据号
    '     bln查看作废-当前查阅的是作废单据
    '     blnDelFee-当前为退费操作
    '出参:objBalanceItems_Out-返回的卡费结算信息集合
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long
    Dim objItem As clsBalanceItem, objCard As Card
    Dim dbl误差费 As Double
    On Error GoTo errHandle
    
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    
    If cllCardFeeBalance Is Nothing Then Exit Function
    If objBalanceItems_Out.类型 = gEM_记帐单 Then zlGetBalanceItemsFromCardFeeColl = True: Exit Function
    
    '    blnc_mode   C   1   结算方式名称
    '    balance_id  N   1   结帐ID
    '    blnc_money  N   1   结帐金额
    '    pay_cardno  N   1   支付卡号
    '    pay_swapno  C   1   交易流水号
    '    pay_swapmemo    C   1   交易说明
    '    relation_id N   1   关联交易id
    '    cardtype_id N   1   卡类别id
    '    consume_card    N   1   是否消费卡:1-是;0-不是
    '    blnc_nature N   1   结算性质:1-现金结算方式,2-其他非医保结算 , 8-结算卡结算 ,9-误差费
    '    blnc_statu  N   1   结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
    '    consume_card_id N   1   消费卡id
    '    blnc_no C   1   结算号码
    '    blnc_memo   C   1   摘要
    
    objBalanceItems_Out.结算金额 = 0
    For i = 1 To cllCardFeeBalance.Count
        Set cllTemp = cllCardFeeBalance(i)
        Set objItem = New clsBalanceItem
        Set objCard = zlGetCardFromCardType(Val(Nvl(cllTemp("_cardtype_id"))), Val(Nvl(cllTemp("_consume_card"))) = 1, Nvl(cllTemp("_blnc_mode")))
        If Val(Nvl(cllTemp("_blnc_nature"))) = 9 Then
            dbl误差费 = RoundEx(dbl误差费 + Val(Nvl(cllTemp("_blnc_money"))), 6)
        Else
            With objItem
                Set .objCard = objCard
                .单据号 = strNO
                .单据性质 = 5   ' 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡
                .结算方式 = Nvl(cllTemp("_blnc_mode"))
                .结算金额 = Val(Nvl(cllTemp("_blnc_money")))
                .关联交易ID = Val(Nvl(cllTemp("_relation_id")))
                .交易流水号 = Nvl(cllTemp("_pay_swapno"))
                .交易说明 = Nvl(cllTemp("_pay_swapmemo"))
                .结算号码 = Nvl(cllTemp("_blnc_no"))
                .结算性质 = Val(Nvl(cllTemp("_blnc_nature")))
                .结算摘要 = Nvl(cllTemp("_blnc_memo"))
                .卡号 = Nvl(cllTemp("_pay_cardno"))
                
                .卡类别ID = Val(Nvl(cllTemp("_cardtype_id")))
                .消费卡ID = Val(Nvl(cllTemp("_consume_card_id")))
                .消费卡 = Val(Nvl(cllTemp("_consume_card"))) = 1
                .是否密文 = objCard.卡号密文规则 <> ""
                .原始金额 = .结算金额
                .未退金额 = .结算金额
                .是否允许编辑 = False
                .是否允许删除 = False
                .校对标志 = Val(Nvl(cllTemp("_blnc_statu")))
                .是否结算 = .校对标志 = 2 Or .校对标志 = 0
                .密码 = ""
                .帐户余额 = 0
                If .卡类别ID = 0 Then   '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                      .结算类型 = 0
                ElseIf .卡类别ID <> 0 And .消费卡 = False Then
                      .结算类型 = 3
                ElseIf .卡类别ID <> 0 And .消费卡 Then
                      .结算类型 = 5
                Else
                     .结算类型 = 0
                End If
                .是否退款 = blnDelFee
                If bln查看作废 Then
                    .冲销ID = Val(Nvl(cllTemp("_balance_id")))
                    .结算ID = Val(Nvl(cllTemp("_original_id"))) '原结帐ID
                   
                Else
                    .结算ID = Val(Nvl(cllTemp("_balance_id")))
                    .冲销ID = Val(Nvl(cllTemp("_original_id"))) '原结帐ID
                End If
                .是否预交 = False
            End With
            objBalanceItems_Out.AddItem objItem
            objBalanceItems_Out.结算金额 = RoundEx(objBalanceItems_Out.结算金额 + objItem.结算金额, 6)
            objBalanceItems_Out.单据号 = objItem.单据号
            If objItem.卡类别ID <> 0 Then
                objBalanceItems_Out.类型 = IIf(objItem.消费卡, gEM_消费卡, gEM_一卡通)
            Else
                objBalanceItems_Out.类型 = gEM_普通结算
            End If
        End If
    Next
    
    objBalanceItems_Out.误差费 = dbl误差费
    objBalanceItems_Out.未退金额 = objBalanceItems_Out.结算金额
    objBalanceItems_Out.原始金额 = objBalanceItems_Out.结算金额 '暂定为未退部分
    
    zlGetBalanceItemsFromCardFeeColl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceItemsFromDepositColl(ByVal cllDeposits As Collection, ByRef objBalanceItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据服务返回的集合,按记录集方式返回结算信息
    '入参:cllDeposits-当前集合
    '
    '出参:objBalanceItems-返回的卡费结算信息集合
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long
    Dim objItem As clsBalanceItem, objCard As Card
    
    On Error GoTo errHandle
    
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    
    If cllDeposits Is Nothing Then zlGetBalanceItemsFromDepositColl = True: Exit Function
    If cllDeposits.Count = 0 Then zlGetBalanceItemsFromDepositColl = True: Exit Function
    
    '    deposit_id  N   1   预交id
    '    deposit_no  C   1   预交单据号
    '    deposit_money   N   1   预交金额
    '    blnc_mode   C   1   结算方式
    '    pay_cardno  N   1   支付卡号
    '    pay_swapno  C   1   交易流水号
    '    pay_swapmemo    C   1   交易说明
    '    relation_id N   1   关联交易id
    '    cardtype_id N   1   卡类别id
    '    consume_card    N   1   是否消费卡:1-是;0-不是
    '    blnc_nature N   1   结算性质
    '    blnc_statu  N   1   结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
    '    consume_card_id N   1   消费卡id
    '    blnc_no C   1   结算号码
    '    blnc_memo   C   1   摘要

     Set cllTemp = cllDeposits
     Set objItem = New clsBalanceItem
     Set objCard = zlGetCardFromCardType(Val(Nvl(cllTemp("_cardtype_id"))), Val(Nvl(cllTemp("_consume_card"))) = 1, Nvl(cllTemp("_blnc_mode")))
     With objItem
         .单据性质 = 1 ' 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡
         .单据号 = Nvl(cllTemp("_deposit_no"))
         .结算方式 = Nvl(cllTemp("_blnc_mode"))
         .预交ID = Val(Nvl(cllTemp("_deposit_id")))
         .结算金额 = Val(Nvl(cllTemp("_deposit_money")))
         .关联交易ID = Val(Nvl(cllTemp("_relation_id")))
         .交易流水号 = Nvl(cllTemp("_pay_swapno"))
         .交易说明 = Nvl(cllTemp("_pay_swapmemo"))
         .结算号码 = Nvl(cllTemp("_blnc_no"))
         .结算性质 = Val(Nvl(cllTemp("_blnc_nature")))
         .结算摘要 = Nvl(cllTemp("_blnc_memo"))
         .卡号 = Nvl(cllTemp("_pay_cardno"))
         
         .卡类别ID = Val(Nvl(cllTemp("_cardtype_id")))
         .消费卡ID = Val(Nvl(cllTemp("_consume_card_id")))
         .消费卡 = Val(Nvl(cllTemp("_consume_card"))) = 1
         .是否密文 = objCard.卡号密文规则 <> ""
         .原始金额 = .结算金额
         .未退金额 = .结算金额
         .是否退款 = .结算金额 < 0
         .是否允许编辑 = False
         .是否允许删除 = False
         .校对标志 = Val(Nvl(cllTemp("_blnc_statu")))
         .是否结算 = .校对标志 = 2 Or .校对标志 = 0
         .密码 = ""
         .帐户余额 = 0
         If .卡类别ID = 0 Then   '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
               .结算类型 = 0
         ElseIf .卡类别ID <> 0 And .消费卡 = False Then
               .结算类型 = 3
         ElseIf .卡类别ID <> 0 And .消费卡 Then
               .结算类型 = 5
         Else
              .结算类型 = 0
         End If
         .结算ID = Val(Nvl(cllTemp("_deposit_id")))
         .冲销ID = 0
         .是否预交 = False
     End With
    objBalanceItems_Out.AddItem objItem
    objBalanceItems_Out.结算金额 = RoundEx(objBalanceItems_Out.结算金额 + objItem.结算金额, 6)
 
    zlGetBalanceItemsFromDepositColl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetDepsoitFromColl(ByVal cllDeposit As Collection, ByRef rsDeposit_Out As Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据服务返回的集合,按记录集方式返回结算信息
    '入参:cllDeposit-当前集合
    '
    '出参:rsBalance_out-返回的卡费结算信息集合
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long
    On Error GoTo errHandle
    
    Set rsDeposit_Out = New ADODB.Recordset
    With rsDeposit_Out
        If .State = adStateOpen Then .Close
        .Fields.Append "预交单据号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "预交金额", adDouble, , adFldIsNullable
        .Fields.Append "结算方式", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "卡号", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "交易流水号", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "交易说明", adLongVarChar, 4000, adFldIsNullable
        .Fields.Append "关联交易id", adBigInt, , adFldIsNullable
        
        .Fields.Append "卡类别id", adBigInt, , adFldIsNullable
        .Fields.Append "是否消费卡", adSmallInt, , adFldIsNullable
        .Fields.Append "结算性质", adSmallInt, , adFldIsNullable
        .Fields.Append "结算状态", adSmallInt, , adFldIsNullable
        .Fields.Append "消费卡id", adBigInt, , adFldIsNullable
         
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    
    If cllDeposit Is Nothing Then Exit Function
    '    deposit_no  C   1   预交单据号
    '    deposit_money   N   1   预交金额
    '    blnc_mode   C   1   结算方式
    '    pay_cardno  N   1   支付卡号
    '    pay_swapno  C   1   交易流水号
    '    pay_swapmemo    C   1   交易说明
    '    relation_id N   1   关联交易id
    '    cardtype_id N   1   卡类别id
    '    consume_card    N   1   是否消费卡:1-是;0-不是
    '    blnc_nature N   1   结算性质
    '    blnc_statu  N   1   结算状态:1-未调用接口;2-接口调用成功,但还未收费完成,0-正常结算
    '    consume_card_id N   1   消费卡id

    For i = 1 To cllDeposit.Count
        Set cllTemp = cllDeposit(i)
        With rsDeposit_Out
            .AddNew
                
                !预交单据号 = Nvl(cllTemp("_deposit_no"))
                !预交金额 = Val(Nvl(cllTemp("_deposit_money")))
                !结算方式 = Nvl(cllTemp("_blnc_mode"))
                !卡号 = Nvl(cllTemp("_pay_cardno"))
                !交易流水号 = Nvl(cllTemp("_pay_swapno"))
                !交易说明 = Nvl(cllTemp("_pay_swapmemo"))
                !关联交易ID = Val(Nvl(cllTemp("_relation_id")))
                !卡类别ID = Val(Nvl(cllTemp("_cardtype_id")))
                !是否消费卡 = Val(Nvl(cllTemp("_consume_card")))
                !结算性质 = Val(Nvl(cllTemp("_blnc_nature")))
                !结算状态 = Val(Nvl(cllTemp("_blnc_statu")))
                !消费卡ID = Val(Nvl(cllTemp("_consume_card_id")))
            .Update
        End With
    Next
    zlGetDepsoitFromColl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Public Function Zl_Exsesvr_UpdatePatiBaseInfo(ByVal lng病人ID As Long, Optional ByVal cllUpdateInfo As Collection, _
    Optional ByVal str操作员姓名 As String, Optional ByVal str操作员编码 As String, Optional blnShowErrMsg As Boolean, _
    Optional ByRef str就诊ID As String = "", Optional int场合 As Integer = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新门诊费用信息
    '入参:cllPatiUpdate-修改的病人信息:array(名称,值)
    '                名称包含（姓名,,性别,年龄,病案号(门诊号))
    '     blnShowErrMsg-是否显示错误信息
    '     int场合-场合:1-门诊;2-住院
    '     str就诊ID-主页ID不传为"",否则传入值
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    Dim strJsonPatiInfo  As String
    
    If cllUpdateInfo Is Nothing Then Exit Function
    If cllUpdateInfo.Count = 0 Then Exit Function
    If lng病人ID = 0 Then Exit Function
    
    If blnShowErrMsg Then On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "连接费用域服务失败，无法获取有效的病人信息!"
        If Not blnShowErrMsg Then Err.Raise -1001, strErrMsg, strErrMsg: Exit Function
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strJsonPatiInfo = ""
    For i = 1 To cllUpdateInfo.Count
        varTemp = cllUpdateInfo(i)
        Select Case UCase(varTemp(0))
        Case "姓名", "病人姓名"
            strJsonPatiInfo = strJsonPatiInfo & "," & GetJsonNodeString("pati_name", Trim(varTemp(1)), Json_Text)
        Case "性别"
            strJsonPatiInfo = strJsonPatiInfo & "," & GetJsonNodeString("pati_sex", Trim(varTemp(1)), Json_Text)
        Case "年龄"
            strJsonPatiInfo = strJsonPatiInfo & "," & GetJsonNodeString("pati_age", Trim(varTemp(1)), Json_Text)
        Case "门诊号"
            strJsonPatiInfo = strJsonPatiInfo & "," & GetJsonNodeString("outpatient_num", Trim(varTemp(1)), Json_Text)
        Case Else
        End Select
    Next
    If strJsonPatiInfo = "" Then Exit Function
    strJsonPatiInfo = Mid(strJsonPatiInfo, 2)
    
    strJsonPatiInfo = GetNodeString("update_info") & ":{" & strJsonPatiInfo & "}"
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num, True)
    If str就诊ID <> "" Then
        strJson = strJson & "," & GetJsonNodeString("visit_id", Val(str就诊ID), Json_num, True)
    End If
    strJson = strJson & "," & GetJsonNodeString("occasion", int场合, Json_num, True)
    
    strJson = strJson & "," & GetJsonNodeString("operator_name", str操作员姓名, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_code", str操作员编码, Json_Text)
    strJson = strJson & "," & strJsonPatiInfo
    'Zl_Exsesvr_UpdatePatiBaseInfo
    '    pati_id N   1   病人id
    '    visit_id    N   1   就诊id 门诊病人为挂号ID;住院病人为主页ID;为0说明是外来或非挂号就诊的病人(就诊id_In为空时,不更改该病人的费用部分的业务数据)
    '    occasion    N       场合,1-门诊;2-住院
    '    update_info
    '        outpatient_num  C       门诊号
    '        pati_name   C       姓名
    '        pati_age        C       年龄
    '        pati_sex        C       性别
    '        pati_birthdate  C       出生日期
    '        explain C       说明

   ' strJson = Mid(strJson, 2)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Exsesvr_UpdatePatiBaseInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, blnShowErrMsg, , , , blnShowErrMsg) = False Then Exit Function
    '    output
    '        code    C   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：失败时返回具体的错误信息
    Zl_Exsesvr_UpdatePatiBaseInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Exsesvr_Addcardfeeinfo(ByVal int操作状态 As Integer, _
    cllCardData As Collection, ByRef lng结帐ID_Out As Long, lng预交ID_Out As Long, _
    Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加卡费及预交数据
    '入参:int操作状态-操作状态:0-正常的预交款或卡费缴款;1-保存为未生效的预交款或异常的卡费;2-保存为记帐单;3-保存为划价单
    '     cllData: 卡数据对象
    '          |--billinfo:(结算合计,操作员编号,操作员姓名,登记时间),Key="_billinfo"
    '          |--patinfo:(病人ID,主页ID,病人姓名,性别,年龄,门诊号,住院号,付款方式编号,费别,险类),Key="_patinfo"
    '          |--cardinfo:发卡信息(卡号,卡类别ID,发卡方式(0-发卡,1-补卡,2-换卡),卡号重用,领用id,原卡号),key="_cardinfo"
    '          |--cardfeelists:key="_cardfeelists"
    '               |---cardfeelist:(卡费单据号,序号,价格父号,从属父号,收费类别,收费细目id,收入项目id,标准单价,收据费目,应收金额,实收金额,病人科室id,开单部门id,病人病区id,
    '                                 执行部门id,是否病历费,保险编码,保险项目否,统筹金额,摘要,发卡卡号,发卡卡类别ID,发卡方式(0-发卡,1-补卡,2-换卡)) ,Key="_" & 序号
    
    '          |--balanceinfo:(结算方式,结算号码,卡类别id,结算卡序号,消费卡ID,支付卡号,交易流水号,交易说明,合作单位,是否电子票据) Key="_balanceinfo"
    '          |--depositinfo:(预交单据号,发票号,预交类别,主页id,缴款科室id,缴款金额,缴款单位,单位开户行,摘要,领用id,预交电子票据),Key="_depositinfo",无预交时，不传入
    '          以上，格式为:,格式：array(名称,值)
    '          int操作状态=2-保存为记帐单;3-保存为划价单 的，则无"balanceinfo"和"depositinfo"节点
    '     blnShowErrMsg-是否显示错误信息
    '     blnDepositStartEinvoice-是否启用预交电子票据
    '     blnStartEinvoice-是否启用卡费电子票据
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
        strErrMsg = "连接费用域服务失败，无法获取有效的病人信息!"
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
        Case "年龄"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_age", Trim(varTemp(1)), Json_Text)
        Case "费别"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_category", Trim(varTemp(1)), Json_Text)
        Case "险类"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_type", Val(varTemp(1)), Json_num, True)
        Case "门诊号"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("outpatient_num", Trim(varTemp(1)), Json_Text)
        Case "住院号"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("inpatient_num", Trim(varTemp(1)), Json_Text)
        Case "付款方式名称"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("mdlpay_name", Trim(Trim(varTemp(1))), Json_Text)
 
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
    '        cardno_old  C   1   原卡卡号:换卡时，需要传入原卡号
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
        Case "原卡号"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno_old", Trim(varTemp(1)), Json_Text)
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
            Case "加班标志"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("overtime_flag", Val(Trim(varTemp(1))), Json_num)
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
        '    start_einv  N   1   是否启用电子票据:1-启用;0-不启用

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
            Case "是否电子票据"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("start_einv", Val(varTemp(1)), Json_num, True)
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
            '    start_einv  N   1   是否启用电子票据:1-启用;0-不启用

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
                Case "预交电子票据"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("start_einv", Val(varTemp(1)), Json_num, True)
                Case "领用ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("recv_id", Val(Trim(varTemp(1))), Json_num, True)
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
    Zl_Exsesvr_Addcardfeeinfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function Zl_Exsesvr_CheckCardnoIsUsed(ByVal lng卡类别ID As Long, ByVal strCardNo As String, ByRef bln是否存在_Out As Boolean, ByRef lng领用ID_Out As Long, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定卡号是否在票据使用明细中存在，存在时，返回领用ID
    '入参:lng卡类别ID- 卡号
    '     strCardNo-卡号
    '     blnShowErrMsg-是否显示错误信息
    
    '出参:bln是否存在_Out-卡号存在，返回 true,否则返回False
    '     lng领用ID_Out-存在卡号时，返回原领用ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    bln是否存在_Out = False
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'Zl_Exsesvr_Checkcardnoisused
    'input
    '    cardtype_id N   1   卡类别id
    '    cardno  C   1   卡号


    strJson = strJson & "" & GetJsonNodeString("cardtype_id", lng卡类别ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardno", strCardNo, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Exsesvr_Checkcardnoisused"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    '    isexsit N   1   是否存在:1-存在;0-不存在
    '    recv_id N   1   领用id

  
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    bln是否存在_Out = objServiceCall.GetJsonNodeValue("output.isexsit") = 1
    lng领用ID_Out = objServiceCall.GetJsonNodeValue("output.recv_id")
    Zl_Exsesvr_CheckCardnoIsUsed = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
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
    '              |-预交单号,预交ID,收费单号,结帐ID,操作员编号,操作员姓名,收款时间,病人ID,是否电子票据,是否预交电子票据)
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
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
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
    '    fee_einvoice    N   1   卡费或病历费是否启用电子票据:1-启用;0-不启用
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
    '        deposit_einvoice    N       预交启用电子票据:1-启用;0-不启用
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
        Case "是否电子票据"
            strJson = strJson & "," & GetJsonNodeString("fee_einvoice", Val(varTemp(1)), Json_num, True)
        Case "是否预交电子票据"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("deposit_einvoice", Val(varTemp(1)), Json_num, True)
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
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("send_mode", decode(Val(varTemp(1)), 1, 0, 11, 0, 2, 2, 3, 1, 3), Json_num, True)
            Case "卡号"
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
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zl_ExseSvr_GetNextNo(ByVal int序号 As Integer, ByRef strNo_Out As String, Optional ByVal lng科室ID As Long, _
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
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'input
    '    item_num    N   1   项目序号
    '    dept_id N       科室ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("item_num", int序号, Json_num)
    strJson = strJson & "," & GetJsonNodeString("recv_id", lng科室ID, Json_num)
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
    strNo_Out = objServiceCall.GetJsonNodeValue("output.next_no")
    zl_ExseSvr_GetNextNo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zl_ExseSvr_GetBillStatuByNo(ByVal strNO As String, ByVal int单据性质 As Integer, _
    ByRef int收费状态_Out As Integer, ByRef int异常状态_Out As Integer, Optional ByRef int结帐标志_Out As Integer, _
    Optional ByRef int预交消费标志_Out As Integer, Optional blnShowErrMsg As Boolean = True, Optional ByRef strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号获该单据的收费、异常及结帐等状态
    '入参:strNO-单据号
    '     int单据性质-单据性质:1-收费单;2-记帐单;3-自动记帐单;4-挂号单;5-就诊卡;6-预交单
    '出参:int收费状态_Out:0-未收费或划价;1-已收费或已记帐;2-已全退或全销帐;3-部分退费或部分销帐
    '    int异常状态_Out-:0-正常数据;1-收款发生异常;2-退款发生异常
    '    int结帐标志_Out:针对记帐单有效;0-未结帐;1-已经结帐
    '    int预交消费标志_Out:1-发生了消费;0-未发生消费
    '     strErrMsg_Out-错误信息
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
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'input
    '    fee_no  C   1   单据号
    '    bill_prop   N   1   记录性质:1-收费单;2-记帐单;3-自动记帐单;4-挂号单;5-就诊卡;6-预交单
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_no", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("bill_prop", int单据性质, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetBillStatuByNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    output
    '        code    N   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：失败时返回具体的错误信息
    '        statu   N   1   收费状态:0-未收费或划价;1-已收费或已记帐;2-已全退或全销帐;3-部分退费或部分销帐
    '        err_sign    N   1   异常标志:0-正常数据;1-收款发生异常;2-退款发生异常
    '        blnc_sign   N   1   结帐标志:针对记帐单有效;0-未结帐;1-已经结帐
    '        consumeed_sign  N   1   预交消费标志:1-发生了消费;0-未发生消费


    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取单据号，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    int收费状态_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.statu")))
    int异常状态_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.err_sign")))
    int结帐标志_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.blnc_sign")))
    int预交消费标志_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.consumeed_sign")))
    zl_ExseSvr_GetBillStatuByNo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zl_ExseSvr_UpdateCardFeeInfo(ByVal int操作标志 As Integer, ByVal strCardFeeNo As String, ByVal cllSendCard As Collection, _
    ByVal str操作员姓名 As String, ByVal str操作员编号 As String, ByVal str登记时间 As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改卡费结算数据
    '入参:int操作标志-0-只修改住院费用记录;1-修改费用记录及票据使用明细
    '     cllSendCard -返回发卡数据(卡类别ID,变动类型,卡号,原卡号,IC卡号,密码,加密密码,终止使用时间,卡费,病历费,摘要,卡号重用,领用ID),格式:array(名称,值),"_名称"
    '     str登记时间-格式:yyyy-mm-dd hh24:mi:ss
    '     blnShowErrMsg-是否显示错误信息
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strJsonTemp As String, strServiceName    As String, varTemp As Variant
    Dim objServiceCall As Object, i As Long
    Dim intReturn As Integer
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Err = 0: On Error GoTo errHandle:
    'zl_ExseSvr_UpdateCardFeeInfo
    ' input
    ' input
    '    oper_fun    N   1   操作标志:0-只修改住院费用记录;1-修改费用记录及票据使用明细
    '    fee_no  C   1   费用单号：本次要调整的费用单据
    '    operator_name   C   1   操作员姓名
    '    operator_code   C   1   操作员编号
    '    create_time C   1   登记时间:yyyy-mm-dd hh:mi:ss
    '    sendcard_info 发卡信息
    '        send_mode   N   1   发卡方式;0-发卡,1-补卡,2-换卡;3-退卡
    '        cardtype_id C   1   卡类别id
    '        cardno  C   1   卡号:本次发放或绑定或补卡的卡号
    '        recv_id N   1   领用id:票据领用ID(卡号)
    '        cardno_reusing  N   1   卡号重用:1-卡号允许重复使用用;0-不允许重复使用
    '        cardno_old  C   1   原卡卡号:换卡时，需要传入原卡号

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_fun", int操作标志, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_no", strCardFeeNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_name", str操作员姓名, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_code", str操作员编号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("create_time", str登记时间, Json_Text)
    strJsonTemp = ""
    If Not cllSendCard Is Nothing Then
        For i = 1 To cllSendCard.Count
            varTemp = cllSendCard(i)
            Select Case UCase(varTemp(0))
            Case "卡类别ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num, True)
            Case "变动类型"
                ''0-不确定，1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("send_mode", decode(Val(varTemp(1)), 1, 0, 3, 1, 2, 2, 4, 3, 0), Json_num, True)
            Case "卡号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
            Case "原卡号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno_old", Trim(varTemp(1)), Json_Text)
            Case "IC卡号"
            Case "密码"
            Case "加密密码"
            Case "终止使用时间"
            Case "卡费"
            Case "病历费"
            Case "摘要"
            Case "卡号重用"
            Case "领用ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("recv_id", Val(varTemp(1)), Json_num, True)
            End Select
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonTemp = "," & GetNodeString("sendcard_info") & ":{" & strJsonTemp & "}"
        End If
    End If
    strJson = strJson & strJsonTemp
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_UpdateCardFeeInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, blnShowErrMsg, , , , Not blnShowErrMsg) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    zl_ExseSvr_UpdateCardFeeInfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Exsesvr_CardfeeIsBalance(ByVal strCardFeeNo As String, ByVal intQueryType As Boolean, _
    ByRef str结帐单号_Out As String, ByRef bln结帐_Out As Boolean, _
    Optional blnShowErrMsg As Boolean = True, Optional ByRef strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断卡费用是否已经结帐
    '入参:strCardFeeNo-卡费单据号
    '     intQueryType-0-读取卡费;1-病历费;2-卡费及病历费
    '出参:strErrMsg_Out-错误信息
    '     str结帐单号_Out-返回结帐单据号
    '     bln结帐_Out-是否已经结帐:已经结帐返回true,否则返回False
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
  
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'input
    '   cardfee_no  C   1   卡费对应的费用单据号
    '   strCardFeeNo  N   1   读取卡费标志:0-读取卡费,1-病历费;2-卡费或病历费用
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("cardfee_no", strCardFeeNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("strCardFeeNo", intQueryType, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_CardfeeIsBalance"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    'code    C   1   应答码：0-失败；1-成功
    'message C   1   应答消息：失败时返回具体的错误信息
    'isbalanced  N   1   是否已经结帐:1-已结结帐;0-未结帐
    'blnc_no C   1   结帐单据号

    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能有效读取卡费单据，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln结帐_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.isbalanced"))) = 1
    str结帐单号_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.blnc_no"))) = 1
    Zl_Exsesvr_CardfeeIsBalance = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zl_ExseSvr_DelCardFeeCheck(ByVal int退费操作 As Integer, ByVal strCardFeeNo As String, ByVal bln异常重退 As Boolean, _
    ByVal cllBalanceInf As Collection, Optional ByVal strDepositNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退卡及退病历费前数据检查
    '入参:int退费操作-0-仅退卡费;1-仅退病历费;2-病历费及卡费
    '   cllBalanceInf-(退款金额,结算方式,卡类别ID,结算卡序号,是否全退),array(名称,值 ),"_名称"
    '出参:strErrMsg-错误信息
    '     str结帐单号_Out-返回结帐单据号
    '     bln结帐_Out-是否已经结帐:已经结帐返回true,否则返回False
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Collection, cllTemp As Collection, varTemp As Variant
    Dim strErrMsg As String, strJsonBalance As String
    Dim objServiceCall As Object
    Dim intReturn As Integer
  
    If cllBalanceInf Is Nothing Then
        strErrMsg = "未传入必要的检查条件，不能进行" & decode(int退费操作, 1, "退病历费", "退卡") & "，请检查!"
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If GetServiceCall(objServiceCall) = False Then
        strErrMsg = "连接费用域服务失败，无法检查" & decode(int退费操作, 1, "退病历费", "退卡") & "的有效性,请与系统管理员联系!"
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'input
    '    cardfee_no  C   1   卡费单号
    '    deposit_no  C   1   预交单据号
    '    reretruned  N   1   是否异常重退:1-是异常重退;0-非异常重退
    '    delfee_sign N   1   退费标志：0-仅退卡费;1-仅退病历费;2-病历费及卡费
    '    balance_info    C       退款方式
    '        delmoney    N   1   本次退款金额
    '        pay_mode    C   1   结算方式
    '        cardtype_id N   1   卡类别id
    '        consumer_no N   1   结算卡序号，即卡消费接口目录.编号
    '        must_allreturn  N   1   是否全退:1-必须全退;0-允许部分退
    
    
    strJson = "": strJsonBalance = ""
    If Not cllBalanceInf Is Nothing Then
        
        For i = 1 To cllBalanceInf.Count
            varTemp = cllBalanceInf(i)
            Select Case UCase(varTemp(0))
            Case "退款金额"
                 strJsonBalance = strJsonBalance & "," & GetJsonNodeString("delmoney", Val(varTemp(1)), Json_num)
            Case "结算方式"
                strJsonBalance = strJsonBalance & "," & GetJsonNodeString("deposit_no", Trim(varTemp(1)), Json_Text)
            Case "卡类别ID"
                 strJsonBalance = strJsonBalance & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num)
            Case "结算卡序号"
                 strJsonBalance = strJsonBalance & "," & GetJsonNodeString("consumer_no", Val(varTemp(1)), Json_num)
            Case "是否全退"
                 strJsonBalance = strJsonBalance & "," & GetJsonNodeString("must_allreturn", Val(varTemp(1)), Json_num)
            End Select
        Next
    End If
    

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("delfee_sign", int退费操作, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardfee_no", strCardFeeNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("reretruned", IIf(bln异常重退, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("deposit_no", strDepositNo, Json_Text)
    If strJsonBalance <> "" Then
        strJsonBalance = Mid(strJsonBalance, 2)
        strJsonBalance = "," & GetNodeString("balance_info") & ":{" & strJsonBalance & "}"
    End If
    strJson = strJson & strJsonBalance
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_DelCardFeeCheck"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False, , , , False) = False Then Exit Function
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：失败时返回具体的错误信息
    '    tip_list[]  C   1   提示列表:主要是可能存在多个提示询问方式，所以用列表,禁止时，返回一条信息
    '        tip_mode    C   1   控制方式:1-提示询问;2-禁止
    '        tip_message C   1   提示信息

    If intReturn = 0 Then
        strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg = "" Then
            strErrMsg = "发生不可预知的错误"
        End If
        MsgBox strErrMsg & ",不能进行" & decode(int退费操作, 1, "退病历费", "退卡") & "操作!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set clldata = objServiceCall.GetJsonListValue("output.tip_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            strErrMsg = Nvl(cllTemp("_tip_message"))
              
              Select Case Val(Nvl(cllTemp("_tip_mode")))
              Case 1  '提示
                    If MsgBox(strErrMsg & ",你是否真的要" & decode(int退费操作, 1, "退病历费", "退卡") & "操作?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                      Exit Function
                    End If
              Case 2  '禁止
                  MsgBox strErrMsg & ",不能进行" & decode(int退费操作, 1, "退病历费", "退卡") & "操作!", vbInformation + vbOKOnly, gstrSysName
                  Exit Function
              End Select
        Next
    End If
    zl_ExseSvr_DelCardFeeCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function Zl_Exsesvr_Delcardfeeinfo(ByVal int操作状态 As Integer, cllDelFeeData As Collection, ByRef lng结帐ID_Out As Long, lng预交ID_Out As Long, _
    Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退卡费、病历费及预交数据
    '入参:int操作状态-操作状态:0-正常的预交款或卡费的退款记录;1-保存为异常的退款记录;2-作废异常数据
    '     cllDelFeeData-退费数据
    '        |-(卡费单号,预交单号,是否退卡费,是否退病历费,操作员姓名,操作员编号,退费时间,结算信息) array(名称,值) ,"_名称)
    '        |-结算信息:(退款金额,结算方式,结算号码,卡类别id,结算卡序号,支付卡号,交易流水号,交易说明,合作单位,关联交易ID,结算摘要) Key="_结算信息"
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
        strErrMsg = "连接费用域服务失败，无法获取有效的病人信息!"
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
                    Case "结算摘要"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("memo", Trim(varTemp(1)), Json_Text)
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
    Zl_Exsesvr_Delcardfeeinfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
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
            !id = Val(Nvl(cllTemp("_recv_id")))
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
    If rsInvoice_Out.RecordCount <> 0 Then rsInvoice_Out.MoveFirst
    zlGetReceiveInvoiceRecFromCollect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
Public Function zl_ExseSvr_GetRelatedTransInfo(ByVal str关联交易Ids As String, ByRef rsSwap_Out As ADODB.Recordset, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据关联交易ID,获取交易信息
    '入参:str关联交易Ids-多个用逗号分离
 
    '出参: rsSwap_Out-返回本次交易信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    If blnShowErrMsg Then On Error GoTo errHandle:
    
    Set rsSwap_Out = New ADODB.Recordset
    With rsSwap_Out
        If .State = adStateOpen Then .Close
        .Fields.Append "关联交易ID", adBigInt, , adFldIsNullable
        .Fields.Append "卡类别ID", adBigInt, , adFldIsNullable
        .Fields.Append "结算方式", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "交易流水号", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "交易说明", adLongVarChar, 4000, adFldIsNullable
        .Fields.Append "原始金额", adDouble, , adFldIsNullable
        .Fields.Append "已退金额", adDouble, , adFldIsNullable
        .Fields.Append "剩余未退金额", adDouble, , adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    'zl_ExseSvr_GetRelatedTransInfo
    '   input
    '        related_ids    C   1   关联交易ID:多个用逗号分离
 
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("related_ids", str关联交易Ids, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetRelatedTransInfo"
    
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：失败时返回具体的错误信息
    '    swap_list[]   C   1   交易信息列表
    '       related_id N   1   关联交易ID
    '       cardtype_id N   1   卡类别ID
    '       blnc_mode   C   1   结算方式
    '       swapno  C   1   交易流水号
    '       swapmemo    C   1   交易说明
    '       original_money  N   1   原始金额
    '       return_money    N   1   已退金额

 

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "未找到符合条件的交易信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set clldata = objServiceCall.GetJsonListValue("output.swap_list")
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then Exit Function
 
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsSwap_Out
            .AddNew
            !关联交易ID = Val(Nvl(cllTemp("_related_id")))
            !卡类别ID = Val(Nvl(cllTemp("_cardtype_id")))
            !结算方式 = Nvl(cllTemp("_blnc_mode"))
            !交易流水号 = Nvl(cllTemp("_swapno"))
            !交易说明 = Nvl(cllTemp("_swapmemo"))
            !原始金额 = Val(Nvl(cllTemp("_original_money")))
            !已退金额 = Val(Nvl(cllTemp("_return_money")))
            !剩余未退金额 = RoundEx(Val(Nvl(!原始金额)) - Val(Nvl(!已退金额)), 5)
            .Update
        End With
    Next
    zl_ExseSvr_GetRelatedTransInfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function Zl_Exsesvr_Geteinvoicesinfo(ByVal frmMain As Object, ByVal strNO As String, ByRef cllPati_Out As Collection, _
    ByRef lngEInvoiceID_out As Long, ByRef strInvoiceNO_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取电子票据信息
    '入参:
    '出参:cllPati_Out-返回病人集
    '     lngEInvoiceID_out-返回电子票据ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object, intReturn As Integer
    Dim blnShowErrMsg As Boolean, strErrmsg_Out As String
    
    blnShowErrMsg = True
    
    Dim cllTemp As Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '--    query_type  N    查询范围:0-所有;1-只查询有效的电子票据;2-查询原始电子票据信息
    '--    occasion  N  1  结算场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡,6-补充医保结算
    '--    fee_nos  C    query_type=2时有效:单据号:结算场合=2时，为预交NO, 结算id未传入，该节点必传
    '--    balance_id  N    结算ID：结算场合=2时，为预交ID
    '--    read_oldbill  N  1  是否只读取原始单据的电子票据:1-是;2-否
    '--    invoice_type  N    票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", 2, Json_num)
    strJson = strJson & "," & GetJsonNodeString("occasion", 5, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_nos", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("invoice_type", 5, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Geteinvoicesinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, True) = False Then Exit Function
    '    output
    '--    pati_info  C    病人信息
    '--      pati_id  N  1  病人ID
    '--      pati_pageid  N    主页ID
    '--      pati_name  C  1  姓名
    '--      pati_sex  C  1  性别
    '--      pati_age  C  1  年龄
    '--      outpatient_num  C  1  门诊号
    '--      inpatient_num  C  1  住院号
    '--    einvoice_info  C    电子票据信息:query_type=2时返回
    '--      einv_id  N  1  电子票据ID
    '--      paper_nos  C  1  未回收的纸质发票信息,多个用逗号返回
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取单据号，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllTemp = objServiceCall.GetJsonListValue("output.data.pati_info")
    Set cllPati_Out = New Collection
    cllPati_Out.Add Val(Nvl(cllTemp("_pati_id"))), "_病人ID"
    cllPati_Out.Add Val(Nvl(cllTemp("_pati_pageid"))), "_主页ID"
    
    cllPati_Out.Add Trim(Nvl(cllTemp("_pati_name"))), "_姓名"
    cllPati_Out.Add Trim(Nvl(cllTemp("_pati_sex"))), "_性别"
    cllPati_Out.Add Trim(Nvl(cllTemp("_pati_age"))), "_年龄"
    
    cllPati_Out.Add Trim(Nvl(cllTemp("_outpatient_num"))), "_门诊号"
    cllPati_Out.Add Trim(Nvl(cllTemp("_inpatient_num"))), "_住院号"
    cllPati_Out.Add 0, "_险类"
    
    Set cllTemp = objServiceCall.GetJsonListValue("output.data.einvoice_info")
    
    lngEInvoiceID_out = Val(Nvl(cllTemp("_einv_id")))
    strInvoiceNO_Out = Trim(Nvl(cllTemp("_paper_nos")))
   
    Zl_Exsesvr_Geteinvoicesinfo = True
    Exit Function
errHandle:
 
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

 
Public Function zl_ExseSvr_GetUseBillInfo(ByVal strNO As String, ByRef rsInvoice_Out As ADODB.Recordset, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取纸质票据使用信息
    '入参:费用单据号

    '出参: rsInvoice_Out-返回本次交易信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    If blnShowErrMsg Then On Error GoTo errHandle:
        
    Set rsInvoice_Out = New ADODB.Recordset
    With rsInvoice_Out
        If .State = adStateOpen Then .Close
        .Fields.Append "ID", adBigInt, , adFldIsNullable
        .Fields.Append "票据号", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "使用原因", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "使用时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "使用人", adLongVarChar, 100, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    'zl_ExseSvr_GetUseBillInfo
    '   input
    '        occasion    N   1   业务场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
    '        inv_type    N   1   票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    '        fee_nos C   1   费用单据号,多个用逗号分离
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("occasion", 5, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_type", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_nos", strNO, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetUseBillInfo"
    
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：失败时返回具体的错误信息
    '    data[]  C   1   使用明细数据
    '        use_id  N   1   使用id
    '        invoice_no  C   1   发票号
    '        use_note    C   1   使用原因
    '        use_time    C   1   使用时间:yyyy-mm-dd hh24:mi:ss
    '        inv_user    C   1   发票使用人
    '        recv_id C   1   领用ID


    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "未找到符合条件的交易信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set clldata = objServiceCall.GetJsonListValue("output.data")
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then Exit Function
 
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsInvoice_Out
            .AddNew
            !id = Val(Nvl(cllTemp("_use_id")))
            !票据号 = Trim(Nvl(cllTemp("_invoice_no")))
            !使用原因 = Nvl(cllTemp("_use_note"))
            !使用时间 = Mid(Nvl(cllTemp("_use_time")), 6)
            !使用人 = Nvl(cllTemp("_inv_user"))
            .Update
        End With
    Next
    zl_ExseSvr_GetUseBillInfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function Zl_Exsesvr_Getbalanceinfo(ByVal strNos As String, ByRef cllSwapData_Out As Collection, ByRef blnStartEinvoice_out As Boolean, _
    Optional ByVal strInvoiceNO As String, Optional ByVal lng领用ID As Long, Optional ByVal byt场合 As Byte = 5, _
    Optional ByRef dblEInvoice_Out As Double, Optional ByRef lng原结算ID_Out As Long, Optional str登记时间 As String) As Boolean
 
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取电子票据信息
    '入参:
    '出参:cllPati_Out-返回病人集
    '     lngEInvoiceID_out-返回电子票据ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object, intReturn As Integer
    Dim blnShowErrMsg As Boolean, strErrmsg_Out As String
    Dim cllPati As Collection, cllBalanceInfo As Collection
    
    blnShowErrMsg = True
    
    Dim cllTemp As Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Err = 0: On Error GoTo errHandle
    
    '       query_type  N  1  查询范围:0-返回剩余金额;1-仅返回原始结算信息
    '--    occasion  N  1  结算场合:1-收费,2-预交(包含押金),3-结帐(暂无用),4-挂号,5-就诊卡,6-补充医保结算
    '--    fee_nos  C    query_type=2时有效:单据号:结算场合=2时，为预交NO, 结算id未传入，该节点必传
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", 0, Json_num)
    strJson = strJson & "," & GetJsonNodeString("occasion", byt场合, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_nos", strNos, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Getbalanceinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, True) = False Then Exit Function
    '  --  output
    '  --    code  C  1  应答码：0-失败；1-成功
    '  --    message  C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '  --    data  C    结算信息
    '  --      pati_info  C    病人信息
    '  --        pati_id  N  1  病人ID
    '  --        pati_pageid  N    主页ID
    '  --        pati_name  C  1  姓名
    '  --        pati_sex  C  1  性别
    '  --        pati_age  C  1  年龄
    '  --        outpatient_num  C  1  门诊号
    '  --        inpatient_num  C  1  住院号
    '  --        insurance_type  N  1  险类
    '  --      balance_info  C    结算信息
    '  --        invoice_no  C  1  发票号
    '  --        balance_oldid  N  1  原结算ID
    '  --        create_time  C  1  收费时间:yyyy-mm-dd hh:mi:ss
    '  --        total  N  1  结算总额
    '  --        balance_unit  N  1  是否合约单位结算
    '  --        balance_type  N  1  "预交时，预交类别:1-门诊;2-住院 ;3-门诊和住院;结帐时：结帐类型:1-门诊;2-住院 ;3-门诊和住院;
    '  --        start_einv  N  1  是否启用电子票据
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If blnShowErrMsg And strErrmsg_Out <> "" Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllTemp = objServiceCall.GetJsonListValue("output.data.pati_info")
    Set cllPati = New Collection
    '1.创建病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类）,key("_",节点名称)
     cllPati.Add Val(Nvl(cllTemp("_pati_id"))), "_病人ID"
    cllPati.Add Val(Nvl(cllTemp("_pati_pageid"))), "_主页ID"
    
    cllPati.Add Trim(Nvl(cllTemp("_pati_name"))), "_姓名"
    cllPati.Add Trim(Nvl(cllTemp("_pati_sex"))), "_性别"
    cllPati.Add Trim(Nvl(cllTemp("_pati_age"))), "_年龄"
    
    cllPati.Add Trim(Nvl(cllTemp("_outpatient_num"))), "_门诊号"
    cllPati.Add Trim(Nvl(cllTemp("_inpatient_num"))), "_住院号"
    cllPati.Add Val(Nvl(cllTemp("_insurance_type"))), "_险类"
    
    
   
    Set cllTemp = objServiceCall.GetJsonListValue("output.data.balance_info")

    dblEInvoice_Out = RoundEx(Val(Nvl(cllTemp("_total"))), 6)
    
     
    blnStartEinvoice_out = Val(Nvl(cllTemp("_start_einv"))) = 1
    lng原结算ID_Out = Val(Nvl(cllTemp("_balance_oldid")))


    '2.创建结算信息:(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID,合约单位结帐,结帐类型)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add strInvoiceNO, "_发票号"
    cllBalanceInfo.Add lng原结算ID_Out, "_结算ID"
    cllBalanceInfo.Add 0, "_冲销ID"
    cllBalanceInfo.Add strNos, "_单据号"
    If str登记时间 <> "" Then
        cllBalanceInfo.Add Nvl(cllTemp("_create_time")), "_登记时间"
    Else
        cllBalanceInfo.Add str登记时间, "_登记时间"
    End If
    
    cllBalanceInfo.Add 0, "_是否补结算"
    cllBalanceInfo.Add 0, "_是否部分退款"
    cllBalanceInfo.Add UserInfo.编号, "_操作员编号"
    cllBalanceInfo.Add UserInfo.姓名, "_操作员姓名"
    cllBalanceInfo.Add dblEInvoice_Out, "_结算金额"
    cllBalanceInfo.Add lng领用ID, "_领用ID"
    cllBalanceInfo.Add Val(Nvl(cllTemp("_balance_unit"))), "_合约单位结帐"
    cllBalanceInfo.Add Val(Nvl(cllTemp("_balance_type"))), "_结算类型"
 
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
      
   
    Zl_Exsesvr_Getbalanceinfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function Zl_Exsesvr_GetbalanceinfoFromNos(ByVal strNos As String, Optional ByVal byt场合 As Byte = 5, _
    Optional ByRef dblEInvoice_Out As Double, Optional ByRef lng原结算ID_Out As Long, _
    Optional ByRef blnStartEinvoice_out As Boolean) As Boolean
 
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取电子票据信息
    '入参:strNos-单据号
    '出参:lng原结算ID_Out-返回原结账ID
    '     dblEInvoice_Out-返回原始金额
    '     blnStartEinvoice_Out-是否启用电子票据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object, intReturn As Integer
    Dim blnShowErrMsg As Boolean, strErrmsg_Out As String
    Dim cllTemp As Collection
    
    blnShowErrMsg = True
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "连接费用域服务失败，无法获取有效的病人信息!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '   query_type  N  1  查询范围:0-返回剩余金额;1-仅返回原始结算信息
    '--    occasion  N  1  结算场合:1-收费,2-预交(包含押金),3-结帐(暂无用),4-挂号,5-就诊卡,6-补充医保结算
    '--    fee_nos  C    query_type=2时有效:单据号:结算场合=2时，为预交NO, 结算id未传入，该节点必传
    strJson = ""
    
    strJson = strJson & "" & GetJsonNodeString("query_type", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("occasion", byt场合, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_nos", strNos, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Getbalanceinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, True) = False Then Exit Function
    '  --  output
    '  --    code  C  1  应答码：0-失败；1-成功
    '  --    message  C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
    '  --    data  C    结算信息
    '  --      balance_info  C    结算信息
    '  --        invoice_no  C  1  发票号
    '  --        balance_oldid  N  1  原结算ID
    '  --        create_time  C  1  收费时间:yyyy-mm-dd hh:mi:ss
    '  --        total  N  1  结算总额
    '  --        balance_unit  N  1  是否合约单位结算
    '  --        balance_type  N  1  "预交时，预交类别:1-门诊;2-住院 ;3-门诊和住院;结帐时：结帐类型:1-门诊;2-住院 ;3-门诊和住院;
    '  --        start_einv  N  1  是否启用电子票据
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If blnShowErrMsg And strErrmsg_Out <> "" Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
     
    
    Set cllTemp = objServiceCall.GetJsonListValue("output.data.balance_info")

    dblEInvoice_Out = RoundEx(Val(Nvl(cllTemp("_total"))), 6)
    blnStartEinvoice_out = Val(Nvl(cllTemp("_start_einv"))) = 1
    lng原结算ID_Out = Val(Nvl(cllTemp("_balance_oldid")))
    Zl_Exsesvr_GetbalanceinfoFromNos = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

