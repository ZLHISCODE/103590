Attribute VB_Name = "mdlPubServerExse"
Option Explicit

'*********************************************************************************************************************************************
'功能:所有涉及调用临床的相关服务
'接口说明:
'    1.zl_ExseSvr_GetPatiSurplusInfo-获取病人余额信息
'    2.zl_ExseSvr_GetConsumerCardType-获取消费卡类别信息接口
'    3.Zl_Exsesvr_GetConsumerCardInfo根据卡号和接口编号获取消费卡信息
'编制:刘兴洪
'日期:2019*10*31 14:47:18
'*********************************************************************************************************************************************

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
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.aveErrLog
End Function


Public Function zl_ExseSvr_GetPatiSurplusInfo(ByVal str病人Ids As String, ByRef cllSurplusData_Out As Collection, _
    Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByVal strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人费用余额信息
    '入参:str病人Ids-病人ID,多个用逗号分离
    '     blnNotShowErrMsg-是否显示错误信息框,true-不显示;false-显示
    '出参:cllSurplusData_Out-返回病人信息集
    '     strErrMsg_out-不显示时，返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    On Error GoTo errHandle
    
    Set cllSurplusData_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接临床域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    

    'zl_ExseSvr_GetPatiSurplusInfo
    'input           获取病人当日发生的费用总额
    '    pati_ids    C       病人ID,多个用逗号分离
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_ids", str病人Ids, Json_Text)
     strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetPatiSurplusInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   应答消息：失败时返回具体的错误信息
    '    surplus_list[]  C   1   余额列表
    '        pati_Id N       病人ID
    '        outdpst_surplus N   1   门诊预交余额
    '        indpst_surplus  N   1   住院预交余额
    '        outfee_surplus  N   1   门诊费用余额
    '        infee_surplus   N   1   住院费用余额
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "未找到符合条件的病人信息，请检查！"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set cllSurplusData_Out = objServiceCall.GetJsonListValue("output.surplus_list", "pati_id")
    zl_ExseSvr_GetPatiSurplusInfo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zl_ExseSvr_GetConsumerCardType(ByRef cllTypesData_out As Collection, Optional ByVal blnOnlyStart As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡类别信息服务接口
    '入参:blnOnlyStart-只获取启用的卡类别
    '出参:cllTypesData_out-返回卡类别信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:47:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    
    On Error GoTo errHandle
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接费用域服务失败，无法获取有效的消费卡类别！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetConsumerCardType
    '  --功能:获取消费卡类别
    '  --入参：Json_In:格式
    '  --    input
    '  --      enabled                N    是否启用:1-已启用;0-所有
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("enabled", IIf(blnOnlyStart, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetConsumerCardType"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                      N   1   应答码：0-失败；1-成功
    '  --    message                   C   1   应答消息：失败时返回具体的错误信息
    '  --    type_list[]               C  1  支持的卡类别列表
    '  --          cardtype_id       N  1  id
    '  --          cardtype_num      N  1  编号
    '  --          cardtype_name     C  1  名称
    '  --          cardtype_stname   C  1  短名
    '  --          prefix_text         C  1  前缀文本
    '  --          cardno_len          N  1  卡号长度
    '  --          default             N  1  缺省标志
    '  --          fixed               N  1  是否固定:1-是系统固定;0-不是系统固定
    '  --          strict              N  1  是否严格控制:1-是严格控制;0-不是严格控制
    '  --          self_make           N  1  是否自制:1-是的;0-不是
    '  --          allow_return_cash   N  1  是否退现:1-允许;0-不允许
    '  --          must_all_return     N   1   是否全退:1-必需全退;0-允许部分退
    '  --          specpati            N   1   特定病人
    '  --          component           C   1   部件
    '  --          memo                C   1   备注
    '  --          blnc_mode           C   1   结算方式
    '  --          blnc_nature         N   1   结算性质
    '  --          pwdtxt           N   1   是否密文
    '  --          enabled             N   1   是否启用:1-已启用;0-未启用
    '  --          pwd_len             N   1   密码长度
    '  --          pwd_len_limit       N   1   密码长度限制:0-不作限制;1-固定密码长度;-n表示密码必须输入好多个位密码以上,但不能超过密码长度
    '  --          pwd_rule            N   1   密码规则:０-数字和字符组成;1-仅为数字组成
    '  --          readcard_nature     C   1   读卡性质,医疗卡读卡方式：第一位为:是否刷卡;第二位为是否扫描;第三位是否接触式读卡;第四位是否非接触式读卡。例如刷卡：'1000'
    '  --          keyboard_mode       N   1   键盘控制方式:：0-禁止使用软键盘;1-使用数字软键盘 ,2-使用字符软键盘
    '  --          def_delcash         N   1   是否缺省退现:允许退现时,默认是否退现
    Set cllData = objServiceCall.GetJsonListValue("output.type_list")
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    Set cllTypesData_out = cllData
    zl_ExseSvr_GetConsumerCardType = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_Exsesvr_GetConsumerCardInfo(ByVal lngCardTypeID As String, ByVal strCardNo As String, ByRef dbl余额_out As Double, _
    Optional lng消费卡ID_Out As Long, Optional strPwd_Out As String, Optional str限制类别_Out As String, Optional str场合_Out As String, _
    Optional lng病人ID_Out As Long, Optional bln特定病人_Out As Boolean, Optional strErrMsg_Out As String, Optional bln有效性检查 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡号和接口编号获取消费卡信息
    '入参:strCardNO-消费卡号
    '       lngCardTypeID-消费卡接口序号
    '出参:lng消费卡ID_Out-消费卡id;strPwd_Out-消费卡密码;dbl余额_out-可用余额;str限制类别_Out-限制类型
    '       str场合_Out-应用场合;lng病人ID_Out-病人id;bln特定病人_Out-是否是特定病人
    '       strErrMsg_out-不显示时，返回错误信息
    '返回:成功返回true,否则返回False
    '编制:焦博
    '日期:2019-12-5 10:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall) = False Then
        strErrMsg_Out = "连接费用域服务失败，无法获取有效的病人信息！"
        Exit Function
    End If
    
    'Zl_Exsesvr_Getconsumercardinfo
    '   input
    '       cardno                    C 1 卡号
    '       cardtype_num         N 1 接口编号
    '       check_valid          N 1 检查卡号的有效性,1-检查;0-不检查
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("cardno", strCardNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("cardtype_num", lngCardTypeID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("check_valid", IIf(bln有效性检查, 1, 0), Json_num)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Getconsumercardinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
'    --  output
'    --    code                  N 1 应答码：0-失败；1-成功
'    --    message            C 1 应答消息：失败时返回具体的错误信息
'    --    card_id               N 1 消费卡id
'    --    card_pwd           C 1 密码
'    --    surplus               N 1 余额
'    --    limit_type           N 1 限制类别
'    --    occasion             N 1 应用场合
'    --    pati_id                N 1 病人ID
'    --    specpati              N 1 是否特定病人
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "未找到符合条件的消费卡信息，请检查！"
        End If
        MsgBox strErrMsg_Out, vbOKOnly, gstrSysName
        
        Exit Function
    End If
    lng消费卡ID_Out = Val(NVL(objServiceCall.GetJsonNodeValue("output.card_id")))
    strPwd_Out = NVL(objServiceCall.GetJsonNodeValue("output.card_pwd"))
    dbl余额_out = Val(objServiceCall.GetJsonNodeValue("output.surplus"))
    str限制类别_Out = NVL(objServiceCall.GetJsonNodeValue("output.limit_type"))
    str场合_Out = NVL(objServiceCall.GetJsonNodeValue("output.occasion"))
    lng病人ID_Out = Val(NVL(objServiceCall.GetJsonNodeValue("output.pati_id")))
    bln特定病人_Out = Val(NVL(objServiceCall.GetJsonNodeValue("output.specpati"))) = 1
    zl_Exsesvr_GetConsumerCardInfo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

