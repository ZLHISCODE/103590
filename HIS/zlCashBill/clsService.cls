VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'一、公共接口
'   1.zlInitCommon-初始化接口
'二、病人域相关服务接口
 
'   2.
'四、内部接口及函数
'   1.GetServiceCall-获取服务对象
'   2.GetCardTypeRecStru-创建获取的卡类别信息集
'编制:刘兴洪
'日期:2019-11-25 11:11:39
'*********************************************************************************************************************************************
Private mlngErrNum As Long, mstrSource As String, mstrErrMsg As String
Private mobjServiceCall As Object
Private mlngModule As Long
Private mrsCardType As ADODB.Recordset  '卡类别集

Public Function zlInitCommon(ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngModule-模块号
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlngModule = lngModule
    If GetServiceCall(mobjServiceCall) = False Then Set mobjServiceCall = Nothing: Exit Function
    zlInitCommon = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
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
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zl_PatiSvr_GetCardTypes(ByRef rsCardType_Out As ADODB.Recordset, _
    Optional bln允许发卡 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡服务
    '入参:
    '出参:rsCardType_Out-返回的卡集合
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 16:53:48
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    
    On Error GoTo errHandle
    
    If Not mrsCardType Is Nothing Then
        mrsCardType.Filter = 0
        If mrsCardType.RecordCount <> 0 Then mrsCardType.MoveFirst
        Set rsCardType_Out = mrsCardType: zl_PatiSvr_GetCardTypes = True: Exit Function
    End If
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的卡类别！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'zl_PatiSvr_GetCardTypes
    'input
    '    cardtype_id C       卡类别id:NULL表示不按卡类别ID查找
    '    query_type  N   1   查询类型:0-所有信息;1-基本信息(返回:id,编码，名称,卡号长度,前缀文本,是否启用,结算方式,是否全退,是否退现)
    '    cert_cardtype   N       只读取证件作为卡类别的医疗卡类别:1-只读取是否证件=1的医疗卡;0-全部读取
    '    dffective_cardtype  N       只读取有效的卡类别:1-只读取有效的卡类别;0-全部读取
    
    strJson = strJson & "" & GetJsonNodeString("cardtype_id", "", Json_num)
    strJson = strJson & "," & GetJsonNodeString("query_type", "", Json_num)
    strJson = strJson & "," & GetJsonNodeString("cert_cardtype", 0, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("dffective_cardtype", 1, Json_num, True)
    strJson = "{""input"":{" & strJson & "}}"
    
    'output
    '    cardtype_id N   1   ID
    '    cardtype_code   C   1   编码
    '    cardtype_name   C   1   名称
    '    cardtype_stname C   1   短名
    '    prefix_text C   1   前缀文本
    '    cardno_len  N   1   卡号长度
    '    default    N   1   缺省标志
    '    fixed N   1   是否固定:1-是系统固定;0-不是系统固定
    '    strict   N   1   是否严格控制:1-是严格控制;0-不是严格控制
    '    self_make N   1   是否自制:1-是的;0-不是
    '    exist_account  N   1   是否存在帐户:1-存在帐户;0-不存在账户
    '    allow_return_cash    N   1   是否退现:1-允许;0-不允许
    '    must_all_return   N   1   是否全退:1-必需全退;0-允许部分退
    '    component   C   1   部件
    '    memo    C   1   备注
    '    spec_item   C   1   特定项目
    '    blnc_mode   C   1   结算方式
    '    blnc_nature N   1   结算性质
    '    cardno_pwdtxt   C   1   卡号密文:卡号从第几位至第几位显示密文,格式为:S-N:S表示从第几位开始,至第几位结束.比如:3-10,表示从3位到10位用密文*表示:12********3323主要是适应不同类别的医疗卡
    '    allow_repeat_use N   1   是否重复使用:1-允许;0-不允许
    '    enabled    N   1   是否启用:1-已启用;0-未启用
    '    pwd_len N   1   密码长度
    '    pwd_len_limit   N   1   密码长度限制:0-不作限制;1-固定密码长度;-n表示密码必须输入好多个位密码以上,但不能超过密码长度
    '    pwd_rule    N   1   密码规则:０-数字和字符组成;1-仅为数字组成
    '    allow_vaguefind    N   1   是否模糊查找:1-支持模糊查找;0-不支持
    '    pwd_require    N   1   密码输入限制:0-不限制;1-不输入,提醒;2-不输入禁止;缺省为不限制
    '    default_pwd  N   1   是否缺省密码:1-以身份证后N(以密码长度为准)位作为缺省密码;0-无缺省密码
    '    allow_makecard N   1   是否制卡:1-是;0-否
    '    allow_sendcard N   1   是否发卡:1-是;0-否
    '    allowwritecard    N   1   是否写卡:1-是;0-否
    '    insurance_type  N   1   险类
    '    sendcard_nature N   1   发卡性质:0-不限制;1-同一病人只能发一张卡;2-同一病人允许发多张卡，但需提示;缺省为0
    '    allow_transfer N   1   是否转帐及代扣:1-支持转帐及代扣;0-不支持
    '    readcard_nature C   1   读卡性质,医疗卡读卡方式：第一位为:是否刷卡;第二位为是否扫描;第三位是否接触式读卡;第四位是否非接触式读卡。例如刷卡：'1000'
    '    keyboard_mode   N   1   键盘控制方式:：0-禁止使用软键盘;1-使用数字软键盘 ,2-使用字符软键盘
    '    advsend_buildqrcode N   1   是否医嘱发送调用条码生成接口:1-发送调用生成二维码接口;0-不调用
    '    holding_pay   N   1   是否持卡消费:1-是;0-否
    '    cert_cardtype    N   1   是否证件类型的医疗卡:0-不是；1-是
    '    verfycard    N   1   是否退款验卡
    '    sendcard_sign   N   1   发卡控制:0或NULL-发卡时，卡号必须达到卡号长度;1-发卡时，允许卡号小于等于卡号长度,发卡时，小于卡号长度时，不提示操作员;2-发卡时，允许卡号小于等于卡号长度,小于时，提示操作员。
    '    enterkey_enabled N   1   设备是否启用回车:医疗卡对应的刷卡设备是否启用了回车，如果启用了回车，则卡号长度默认增加一位来屏蔽回车
    '    def_return_cash N   1   是否缺省退现:允许退现时,默认是否退现
    '    balalone N   1   是否独立结算:1-独立结算;0-非独立结算
    '    discern_rule    N   1   卡号识别规则:1-全部转换为大写;0-不区分大小写
    '    def_valid_time  C   1   缺省有效时间:NULL时，表示不限制;非空时，格式为:时间+单位(天，月),比如：3天,3月
    '    scanpay  N   1   是否支持扫码付:是否支持扫码付,支持时，会调用“zlReadQRCode部件”

    strServiceName = "zl_PatiSvr_GetCardTypes"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    Set cllData = objServiceCall.GetJsonListValue("output.type_list")
    If cllData Is Nothing Then Exit Function
    If cllData.Count = 0 Then Exit Function
  
    If cllData Is Nothing Then Exit Function
    If cllData.Count = 0 Then Exit Function
    
    'output
    '    cardtype_id N   1   ID
    '    cardtype_code   C   1   编码
    '    cardtype_name   C   1   名称
    '    cardtype_stname C   1   短名
    '    prefix_text C   1   前缀文本
    '    cardno_len  N   1   卡号长度
    '    default    N   1   缺省标志
    '    fixed N   1   是否固定:1-是系统固定;0-不是系统固定
    '    strict   N   1   是否严格控制:1-是严格控制;0-不是严格控制
    '    self_make N   1   是否自制:1-是的;0-不是
    '    exist_account  N   1   是否存在帐户:1-存在帐户;0-不存在账户
    '    allow_return_cash    N   1   是否退现:1-允许;0-不允许
    '    must_all_return   N   1   是否全退:1-必需全退;0-允许部分退
    '    component   C   1   部件
    '    memo    C   1   备注
    '    spec_item   C   1   特定项目
    '    blnc_mode   C   1   结算方式
    '    blnc_nature N   1   结算性质
    '    cardno_pwdtxt   C   1   卡号密文:卡号从第几位至第几位显示密文,格式为:S-N:S表示从第几位开始,至第几位结束.比如:3-10,表示从3位到10位用密文*表示:12********3323主要是适应不同类别的医疗卡
    '    allow_repeat_use N   1   是否重复使用:1-允许;0-不允许
    '    enabled    N   1   是否启用:1-已启用;0-未启用
    '    pwd_len N   1   密码长度
    '    pwd_len_limit   N   1   密码长度限制:0-不作限制;1-固定密码长度;-n表示密码必须输入好多个位密码以上,但不能超过密码长度
    '    pwd_rule    N   1   密码规则:０-数字和字符组成;1-仅为数字组成
    '    allow_vaguefind    N   1   是否模糊查找:1-支持模糊查找;0-不支持
    '    pwd_require    N   1   密码输入限制:0-不限制;1-不输入,提醒;2-不输入禁止;缺省为不限制
    '    default_pwd  N   1   是否缺省密码:1-以身份证后N(以密码长度为准)位作为缺省密码;0-无缺省密码
    '    allow_makecard N   1   是否制卡:1-是;0-否
    '    allow_sendcard N   1   是否发卡:1-是;0-否
    '    allow_writecard    N   1   是否写卡:1-是;0-否
    '    insurance_type  N   1   险类
    '    sendcard_nature N   1   发卡性质:0-不限制;1-同一病人只能发一张卡;2-同一病人允许发多张卡，但需提示;缺省为0
    '    allow_transfer N   1   是否转帐及代扣:1-支持转帐及代扣;0-不支持
    '    readcard_nature C   1   读卡性质,医疗卡读卡方式：第一位为:是否刷卡;第二位为是否扫描;第三位是否接触式读卡;第四位是否非接触式读卡。例如刷卡：'1000'
    '    keyboard_mode   N   1   键盘控制方式:：0-禁止使用软键盘;1-使用数字软键盘 ,2-使用字符软键盘
    '    advsend_buildqrcode N   1   是否医嘱发送调用条码生成接口:1-发送调用生成二维码接口;0-不调用
    '    holding_pay   N   1   是否持卡消费:1-是;0-否
    '    cert_cardtype    N   1   是否证件类型的医疗卡:0-不是；1-是
    '    verfycard    N   1   是否退款验卡
    '    sendcard_sign   N   1   发卡控制:0或NULL-发卡时，卡号必须达到卡号长度;1-发卡时，允许卡号小于等于卡号长度,发卡时，小于卡号长度时，不提示操作员;2-发卡时，允许卡号小于等于卡号长度,小于时，提示操作员。
    '    enterkey_enabled N   1   设备是否启用回车:医疗卡对应的刷卡设备是否启用了回车，如果启用了回车，则卡号长度默认增加一位来屏蔽回车
    '    def_return_cash N   1   是否缺省退现:允许退现时,默认是否退现
    '    balalone N   1   是否独立结算:1-独立结算;0-非独立结算
    '    discern_rule    N   1   卡号识别规则:1-全部转换为大写;0-不区分大小写
    '    def_valid_time  C   1   缺省有效时间:NULL时，表示不限制;非空时，格式为:时间+单位(天，月),比如：3天,3月
    '    scanpay  N   1   是否支持扫码付:是否支持扫码付,支持时，会调用“zlReadQRCode部件”
    If GetCardTypeRecStru(rsCardType_Out) = False Then
        Set mrsCardType = Nothing: Set rsCardType_Out = Nothing:
     
        Exit Function
    End If
    
    For i = 1 To cllData.Count
        Set cllTemp = cllData(i)
        With rsCardType_Out
        
            If Val(Nvl(cllTemp("_enabled"))) = 1 And _
                (bln允许发卡 And (Val(cllTemp("_self_make")) = 1 Or Val(cllTemp("_allow_sendcard")) = 1) Or bln允许发卡 = False) Then
                .AddNew
                    !ID = cllTemp("_cardtype_id")
                    !编码 = cllTemp("_cardtype_code")
                    !名称 = cllTemp("_cardtype_name")
                    !短名 = cllTemp("_cardtype_stname")
                    
                    !前缀文本 = cllTemp("_prefix_text")
                    !卡号长度 = cllTemp("_cardno_len")
                    !缺省标志 = cllTemp("_default")
                    !是否固定 = cllTemp("_fixed")
                    
                    !是否严格控制 = cllTemp("_strict")
                    !是否自制 = cllTemp("_self_make")
                    !是否存在帐户 = cllTemp("_exist_account")
                    !是否退现 = cllTemp("_allow_return_cash")
                    !是否缺省退现 = cllTemp("_def_return_cash")
                    !是否全退 = cllTemp("_must_all_return")
                    !部件 = cllTemp("_component")
                    !备注 = cllTemp("_memo")
                    
                    !特定项目 = cllTemp("_spec_item")
                    !结算方式 = cllTemp("_blnc_mode")
                    !卡号密文 = cllTemp("_cardno_pwdtxt")
                      
                    !是否重复使用 = cllTemp("_allow_repeat_use")
                    !是否启用 = cllTemp("_enabled")
                    
                    !密码长度 = cllTemp("_pwd_len")
                    !密码长度限制 = cllTemp("_pwd_len_limit")
                    !密码规则 = cllTemp("_pwd_rule")
                    !密码输入限制 = cllTemp("_pwd_require")
                    
                    !是否模糊查找 = cllTemp("_allow_vaguefind")
                    !是否缺省密码 = cllTemp("_default_pwd")
                    
                    
                    
                    !是否制卡 = cllTemp("_allow_makecard")
                    !是否发卡 = cllTemp("_allow_sendcard")
                    !是否写卡 = cllTemp("_allow_writecard")
                    !发卡控制 = cllTemp("_sendcard_sign")
                    
                    !结算性质 = cllTemp("_blnc_nature")
                    !险类 = cllTemp("_insurance_type")
                    !发卡性质 = cllTemp("_sendcard_nature")
                    !是否转帐及代扣 = cllTemp("_allow_transfer")
                    !读卡性质 = cllTemp("_readcard_nature")
                    !键盘控制方式 = cllTemp("_keyboard_mode")
                    
                    !是否持卡消费 = cllTemp("_holding_pay")
                    !是否证件 = cllTemp("_cert_cardtype")
                    
                    !发送调用接口 = cllTemp("_advsend_buildqrcode")
                    !是否退款验卡 = cllTemp("_verfycard")
                                    
                    !是否独立结算 = cllTemp("_balalone")
                    !缺省有效时间 = cllTemp("_def_valid_time")
                    !卡号识别规则 = cllTemp("_discern_rule")
                    !是否支持扫码付 = cllTemp("_scanpay")
                    !是否启用回车 = cllTemp("_enterkey_enabled")
                .Update
            End If
       End With
    Next
    If rsCardType_Out.RecordCount <> 0 Then rsCardType_Out.MoveFirst
    Set mrsCardType = rsCardType_Out
    zl_PatiSvr_GetCardTypes = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetCardTypeRecStru(ByRef rsCardType As ADODB.Recordset) As Boolean

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡类别结构
     '出参:rsCardType-返回的记录集结构
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-24 18:10:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsCardType = New ADODB.Recordset
    With rsCardType
        If .State = 1 Then .Close
        'adBigInt
        .Fields.Append "ID", adVarNumeric, 18, adFldIsNullable
        .Fields.Append "编码", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "名称", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "短名", adLongVarChar, 50, adFldIsNullable
        
        .Fields.Append "前缀文本", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "卡号长度", adSmallInt, 20, adFldIsNullable
        .Fields.Append "缺省标志", adSmallInt, , adFldIsNullable
        .Fields.Append "是否固定", adSmallInt, , adFldIsNullable
        .Fields.Append "是否严格控制", adSmallInt, , adFldIsNullable
        .Fields.Append "是否自制", adSmallInt, , adFldIsNullable
        .Fields.Append "是否存在帐户", adSmallInt, , adFldIsNullable
        
        .Fields.Append "是否退现", adSmallInt, , adFldIsNullable
        .Fields.Append "是否缺省退现", adSmallInt, , adFldIsNullable
        .Fields.Append "是否全退", adSmallInt, , adFldIsNullable
        .Fields.Append "是否重复使用", adSmallInt, , adFldIsNullable
        .Fields.Append "发卡性质", adSmallInt, , adFldIsNullable
        .Fields.Append "密码长度", adSmallInt, , adFldIsNullable
        .Fields.Append "密码长度限制", adSmallInt, , adFldIsNullable
        .Fields.Append "密码规则", adSmallInt, , adFldIsNullable
        .Fields.Append "部件", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "备注", adLongVarChar, 300, adFldIsNullable
        .Fields.Append "特定项目", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "结算方式", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "是否启用", adSmallInt, , adFldIsNullable
        .Fields.Append "卡号密文", adLongVarChar, 50, adFldIsNullable
        
        .Fields.Append "密码输入限制", adSmallInt, , adFldIsNullable
        .Fields.Append "是否缺省密码", adSmallInt, , adFldIsNullable
        .Fields.Append "是否模糊查找", adSmallInt, , adFldIsNullable
        .Fields.Append "是否制卡", adSmallInt, , adFldIsNullable
        .Fields.Append "是否写卡", adSmallInt, , adFldIsNullable
        .Fields.Append "是否发卡", adSmallInt, , adFldIsNullable
        .Fields.Append "发卡控制", adSmallInt, , adFldIsNullable
        
        
        .Fields.Append "结算性质", adSmallInt, , adFldIsNullable
        .Fields.Append "是否证件", adSmallInt, , adFldIsNullable
        
        
        .Fields.Append "是否转帐及代扣", adSmallInt, , adFldIsNullable
        .Fields.Append "读卡性质", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "是否持卡消费", adSmallInt, , adFldIsNullable
        .Fields.Append "发送调用接口", adSmallInt, , adFldIsNullable
        .Fields.Append "是否退款验卡", adSmallInt, , adFldIsNullable
        .Fields.Append "是否独立结算", adSmallInt, , adFldIsNullable
        .Fields.Append "缺省有效时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "卡号识别规则", adSmallInt, , adFldIsNullable
        .Fields.Append "是否支持扫码付", adSmallInt, , adFldIsNullable
        .Fields.Append "险类", adSmallInt, , adFldIsNullable
        .Fields.Append "键盘控制方式", adSmallInt, , adFldIsNullable
        .Fields.Append "是否启用回车", adSmallInt, , adFldIsNullable
    
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    GetCardTypeRecStru = True
End Function
 

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mrsCardType = Nothing
    Set mobjServiceCall = Nothing
    
End Sub
