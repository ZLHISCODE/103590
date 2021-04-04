Attribute VB_Name = "mdlPubServerPati"
Option Explicit
'*********************************************************************************************************************************************
'功能:所有涉及调用病人的相关服务
'接口说明:
'    1.zl_PatiSvr_GetPatiInfsByRange-批量获取病人信息服务
'    2.zl_PatiSvr_GetCardTypes-获取医疗卡类别信息集
'    3.zl_PatiSvr_GetPatiID:获取指定条件的病人Ids
'    4.zl_PatiSvr_GetPatiInfo:获取病人信息详细服务接口
'    5.zl_PatiSvr_GetPatiExtendInfo:获取病人信息从表信息服务接口
'编制:刘兴洪
'日期:2019*10*31 14:47:18
'*********************************************************************************************************************************************
Public gobjServiceCall  As Object
 

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
    Call gobjComLib.SaveErrLog
End Function


Public Function zl_PatiSvr_GetPatiInfsByRange(ByVal intQueryStatus As Integer, ByVal cllFilter As Collection, _
    ByRef cllPatiInfos_out As Collection, Optional ByVal str病人Ids As String, Optional ByRef str病区IDs As String, _
    Optional ByVal blnExpendInfo As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息集
    '入参:intQueryStatus-查询类型(0-仅门诊;1-在院 ;2-门诊及在院)
    '     cllFilter-过滤条件
    '     str病人Ids-病人ID
    '     rsPatiPage-主页信息
    '     str病区IDs-当前病区Ids
    '出参:cllPatiInfos_out-返回的数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim lng病人ID As Long
    
    On Error GoTo errHandle
    Set cllPatiInfos_out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
   
    'zl_PatiSvr_GetPatiInfsByRange
    '  --功能:获取病人信息
    '  --入参：Json_In:格式
    '  --    input
    '  --      query_type        N 1 0：查询基本信息；1：查询基本信息+扩展信息
    '  --      pati_ids          C   病人IDs:多个用逗号
    '  --      pati_name         C   姓名:可以代%分号表表按姓名匹配
    '  --      pati_sex          C   性别,按病人姓名查找有效
    '  --      birthdate_start   C   开始出生日期
    '  --      birthdate_end     C   终止出生日期
    '  --      outpatient_num    C   门诊号
    '  --      pati_idcard       C   身份证号
    '  --      fee_category      C   费别
    '  --      pati_sex          C   性别
    '  --      pati_area         C   区域
    '  --      insurance_num     C   医保号
    '  --      vcard_no          C   就诊卡号
    '  --      iccard_no         C   Ic卡号
    '  --      wardarea_ids      C   病区ids：多个用逗号
    '  --      qurey_Max         N   查询的最大记录数，为0或NULL时表示不限制
    '  --      qrspt_statu       N   查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
    '  --      visit_star_time   C   就诊开始时间:yyyy-mm-dd hh24:mi:ss
    '  --      visit_end_time    C   就诊结束时间:yyyy-mm-dd hh24:mi:ss
    '  --      create_start_time C   开始登记时间:yyyy-mm-dd hh24:mi:ss
    '  --      create_end_time   C   终止登记时间:yyyy-mm-dd hh24:mi:ss
    '  --      module            N   模块号:调用Zl_Custom_Patiids_Get(根据身份证返回病人id)服务时需传入
    '  --      only_ctorg_pati   N   只查询合约单位的病人
    '  --      ctt_unit_id       N   合同单位id,只查询只查询合约单位的病人时有效
    '  --      default_cardtype_id N   缺省卡类别id
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", IIf(blnExpendInfo, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("qrspt_statu", intQueryStatus, Json_num)
    For i = 1 To cllFilter.count
        Select Case cllFilter(i)(0)
        Case "登记时间"
            strJson = strJson & "," & GetJsonNodeString("create_start_time", cllFilter(i)(1), Json_Text)
            strJson = strJson & "," & GetJsonNodeString("create_end_time", cllFilter(i)(2), Json_Text)
        Case "就诊时间"
            strJson = strJson & "," & GetJsonNodeString("visit_start_time", cllFilter(i)(1), Json_Text)
            strJson = strJson & "," & GetJsonNodeString("visit_end_time", cllFilter(i)(2), Json_Text)
        Case "病人ID"
            lng病人ID = cllFilter(i)(1)
        Case "姓名"
            strJson = strJson & "," & GetJsonNodeString("pati_name", cllFilter(i)(1), Json_Text)
        Case "就诊卡号"
            strJson = strJson & "," & GetJsonNodeString("vcard_no", cllFilter(i)(1), Json_Text)
        Case "门诊号"
            strJson = strJson & "," & GetJsonNodeString("outpatient_num", Trim(cllFilter(i)(1)), Json_Text)
        Case "医保号"
            strJson = strJson & "," & GetJsonNodeString("insurance_num", cllFilter(i)(1), Json_Text)
        Case "身份证号"
            strJson = strJson & "," & GetJsonNodeString("pati_idcard", cllFilter(i)(1), Json_Text)
        Case "IC卡号"
            strJson = strJson & "," & GetJsonNodeString("iccard_no", cllFilter(i)(1), Json_Text)
        Case "性别"
            strJson = strJson & "," & GetJsonNodeString("pati_sex", cllFilter(i)(1), Json_Text)
        Case "区域"
            strJson = strJson & "," & GetJsonNodeString("pati_area", cllFilter(i)(1), Json_Text)
        Case "费别"
            strJson = strJson & "," & GetJsonNodeString("fee_category", cllFilter(i)(1), Json_Text)
        Case "查询的最大记录数", "最大记录数", "行数"
            strJson = strJson & "," & GetJsonNodeString("qurey_Max", Val(cllFilter(i)(1)), Json_num, True)
        Case "住院号"
            If gobjOneDataObject.zlGetPatiIDFromInpatientNum(cllFilter(i)(1), lng病人ID) = False Then Exit Function
        Case "缺省卡类别ID"
            strJson = strJson & "," & GetJsonNodeString("default_cardtype_id", cllFilter(i)(1), Json_Text)
        Case "仅合约单位病人"
            strJson = strJson & "," & GetJsonNodeString("only_ctorg_pati", cllFilter(i)(1), Json_num)
        Case "合同单位ID"
            strJson = strJson & "," & GetJsonNodeString("ctt_unit_id", cllFilter(i)(1), Json_num)
        Case "场合"
            strJson = strJson & "," & GetJsonNodeString("occasion", cllFilter(i)(1), Json_num)
        End Select
    Next
    
    If str病区IDs <> "" Then
        strJson = strJson & "," & GetJsonNodeString("wardarea_ids", str病区IDs, Json_Text)
    End If
    
    If lng病人ID <> 0 Then
        If InStr("," & str病人Ids & ",", "," & lng病人ID & ",") = 0 Then str病人Ids = IIf(str病人Ids <> "", ",", "") & lng病人ID
    End If
    If str病人Ids <> "" Then
        strJson = strJson & "," & GetJsonNodeString("pati_ids", str病人Ids, Json_Text)
    End If
    
    strJson = "{""input"":{" & strJson & "}}"
    
    strServiceName = "zl_PatiSvr_GetPatiInfsByRange"
    If objServiceCall.CallService(strServiceName, strJson, , "zl_PatiSvr_GetPatiInfsByRange", glngModul) = False Then Exit Function
    
    '  --出参      json
    '  --output
    '  -- code                   N 1 应答码：0-失败；1-成功
    '  -- message                C 1 应答消息： 失败时返回具体的错误信息
    '  -- pati_list[]                病人信息列表
    '  --   pati_id              N 1 病人id
    '  --   pati_pageid          N 1 主页id：病人信息.主页ID
    '  --   pati_name            C 1 姓名
    '  --   pati_sex             C 1 性别
    '  --   pati_age             C 1 年龄
    '  --   pati_birthdate       C 1 出生日期：yyyy-mm-dd hh24:mi:ss
    '  --   fee_category         C 1 费别
    '  --   outpatient_num       C 1 门诊号
    '  --   inpatient_num        C 1 住院号
    '  --   inp_times            N 1 住院次数
    '  --   pati_nation          C 1 民族
    '  --   pati_idcard          C 1 身份证号
    '  --   vcard_no             C 1 就诊卡号
    '  --   phone_number         C 1 手机号
    '  --   pati_education       C 1 学历
    '  --   ocpt_name            C 1 职业
    '  --   pati_identity        C 1 身份
    '  --   country_name         C 1 国籍
    '  --   pat_home_addr        C 1 家庭地址
    '  --   pati_area            C 1 区域
    '  --   emp_name             C 1 工作单位名称
    '  --   pati_bed             C 1 当前床号
    '  --   is_inhspt            N 1 是否在院：1-在院；0-不在院
    '  --   pati_type            C 1 病人类型(普通，医保，留观)
    '  --   insurance_type       C 1 险类
    '  --   pati_wardarea_id     N 1 当前病区id
    '  --   pati_wardarea_name   C 1 当前病区名称
    '  --   pati_dept_id         N 1 当前科室id
    '  --   pati_dept_name       C 1 当前科室名称
    '  --   adta_time            C 1 入院时间:yyyy-mm-dd hh24:mi:ss
    '  --   adtd_time            C 1 出院时间:yyyy-mm-dd hh24:mi:ss
    '  --   create_time          C 1 登记时间:yyyy-mm-dd hh24:mi:ss
    '  --   medc_card_no         C   医疗卡号
    Set cllData = objServiceCall.GetJsonListValue("output.pati_list")
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    
    Set cllPatiInfos_out = cllData
    zl_PatiSvr_GetPatiInfsByRange = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_PatiSvr_GetCardTypes(ByRef cllCardTypes_out As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡服务
    '入参:
    '出参:cllCardTypes_out-返回的卡集合
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 16:53:48
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    
    On Error GoTo errHandle
    
    Set cllCardTypes_out = New Collection
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
    strJson = strJson & "," & GetJsonNodeString("dffective_cardtype", 0, Json_num, True)
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
    If cllData.count = 0 Then Exit Function
    Set cllCardTypes_out = cllData
    zl_PatiSvr_GetCardTypes = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_PatiSvr_GetPatiID(ByVal cllFindCons As Collection, ByVal cllOtherFindCons As Collection, _
    ByRef cllPatiDatas_Out As Collection, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg As String, _
    Optional ByVal bln检查使用时间 As Boolean = True, Optional ByVal bln检查停用或挂失 As Boolean = True, _
    Optional ByRef intCardStatus As Integer, Optional ByVal blnNotReturnFalse As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人ID信息
    '       cllFindCons-查找条件(array(接点名称,接点值))
    '                接点名称包含:卡类别ID,卡号,二维码,社区名,社区号)
    '       cllOtherFindCons-其他查找条件:array(查询的名称,查询的内容)
    '                   查询的名称:如:门诊号,就诊卡号，身份证号等
    '       blnNotShowErrMsg-不显示错误的提示信息
    '      bln检查使用时间-按卡类别查找有效
    '      bln检查停用或挂失-按卡类别查找有效
    '      blnNotReturnFalse-当服务返回的集合为空时，不返回false
    '出参:strErrMsg-返回的错误信息
    '        lng病人ID-返回的病人ID
    '        cllPatiDatas_Out-返回病人信息数据
    '返回:查找成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-19 09:36:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim strCardTypes  As String, strComminuty As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    Dim strOthers As String
    Dim strCardNo As String
    Dim strRQCode As String
     On Error GoTo errHandle
     
    Set cllPatiDatas_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
 
    'zl_PatiSvr_GetPatiID
    '  --功能:根据指定条件获取病人信息的病人ID
    '  --入参：Json_In:格式
    '  --    input
    '  --          card_find             C
    '  --              cardtype_id       N  1  医疗卡类别ID:=0时，表示模糊查找
    '  --              card_no           C  1  卡号
    '  --              qrcode            C     二维码
    '  --              is_check_usetime  N  1  是否检查使用时间:1-检查;0-不检查
    '  --              is_check_stop     N  1  是否检查停用或挂失:1-检查;0-不检查
    '  --          comminuty_find        C
    '  --            comminuty_num       N  1  社区序号
    '  --            comminuty_code      C     社区号
    '  --          other_cons_find       C
    '  --            find_name           C  1  查找的名称
    '  --            find_text           C  1  查找的文本
    strJson = ""
    If Not cllFindCons Is Nothing Then
        strCardTypes = "": strComminuty = ""
        For i = 1 To cllFindCons.count
            Select Case cllFindCons(i)(0)
            Case "卡类别ID"
                strCardTypes = strCardTypes & "," & GetJsonNodeString("cardtype_id", Val(cllFindCons(i)(1)), Json_num)
            Case "卡号"
                strCardTypes = strCardTypes & "," & GetJsonNodeString("card_no", cllFindCons(i)(1), Json_Text)
                strCardNo = cllFindCons(i)(1)
            Case "二维码"
                If cllFindCons(i)(1) <> "" Then
                    strCardTypes = strCardTypes & "," & GetJsonNodeString("qrcode", cllFindCons(i)(1), Json_Text)
                    strRQCode = cllFindCons(i)(1)
                End If
            Case "社区序号", "社区"
                strComminuty = strComminuty & "," & GetJsonNodeString("comminuty_num", cllFindCons(i)(1), Json_num)
            Case "社区号"
                strComminuty = strComminuty & "," & GetJsonNodeString("comminuty_code", cllFindCons(i)(1), Json_Text)
            End Select
        Next
        
        If strCardTypes <> "" Then
            strCardTypes = strCardTypes & "," & GetJsonNodeString("is_check_usetime", IIf(bln检查使用时间, 1, 0), Json_num)
            strCardTypes = strCardTypes & "," & GetJsonNodeString("is_check_stop", IIf(bln检查停用或挂失, 1, 0), Json_num)
            strCardTypes = Mid(strCardTypes, 2)
            strCardTypes = "," & GetNodeString("card_find") & ":{" & strCardTypes & "}"
        End If
        
        If strComminuty <> "" Then
            strComminuty = Mid(strComminuty, 2)
            strComminuty = "," & GetNodeString("comminuty_find") & ":{" & strComminuty & "}"
        End If
        strJson = strJson & strCardTypes & strComminuty
    End If
    
    If Not cllOtherFindCons Is Nothing Then
        strOthers = ""
        For i = 1 To cllOtherFindCons.count
            strOthers = strOthers & ",{" & GetJsonNodeString("find_name", cllOtherFindCons(i)(0), Json_Text)
            strOthers = strOthers & "," & GetJsonNodeString("find_text", cllOtherFindCons(i)(1), Json_Text) & "}"
        Next
        If strOthers <> "" Then
            strOthers = Mid(strOthers, 2)
            strJson = strJson & "," & GetNodeString("other_cons_find") & ":" & strOthers & ""
        End If
    End If
    
    If strJson = "" Then
        strJson = strJson & "," & GetNodeString("card_find") & ":{" & GetJsonNodeString("cardtype_id", 0, Json_num) & "}"
    End If
    
    strJson = Mid(strJson, 2)
    strJson = "{""input"":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetPatiID"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    
    'output
    '    code    N   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    '    pati_list[] C   1   病人列表,模糊查找时，可能存在多个
    '        cardtype_id N   1   卡类别ID
    '        pati_id N   1   病人ID:未找到时也成功，返回0
    '        card_pwd    C   1   密码
    '        pati_pageid N       主页ID
    '        card_status N   1   当前卡状态。0-正常有效卡;1-已挂失; 2-补卡停用;3-失效卡（病认医疗卡信息.终止使用时间到期时返回该状态，仅本服务使用）

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If blnNotReturnFalse Then zl_PatiSvr_GetPatiID = intReturn = 1
    If intReturn = 0 Then
        strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        intCardStatus = Val(NVL(objServiceCall.GetJsonNodeValue("output.card_status")))
        If strErrMsg = "" Then
            If strCardNo = "" Then
                strErrMsg = "未找到符合条件的病人信息，请检查！"
            Else
                strErrMsg = "未找到卡号为" & strCardNo & "的病人，请检查该卡是否有效卡！"
            End If
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    intCardStatus = objServiceCall.GetJsonNodeValue("output.card_status")
    Set cllData = objServiceCall.GetJsonListValue("output.pati_list")
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    
    Set cllPatiDatas_Out = cllData
    zl_PatiSvr_GetPatiID = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zl_PatiSvr_GetPatiInfo(ByVal lng病人ID As Long, _
    ByVal cllOtherFindCons As Collection, ByRef cllPatiDatas_Out As Collection, _
    Optional ByVal int查询类型 As Integer = 0, _
    Optional ByVal bln包含家属 As Boolean, _
    Optional ByVal bln包含过敏药物 As Boolean, _
    Optional ByVal bln包含免疫信息 As Boolean, _
    Optional ByVal bln包含卡信息 As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg As String, _
    Optional ByVal bln包含医保密码 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人详细信息服务接口
    '入参:cllOtherFindCons-其他查找条件(array(查询名称,查询值)
    '             查询名称:病人IDS,姓名,性别,出生日期等,见query_cons_list[]列表中的描述部分
    '      int查询类型-0-基本;1-基本+联系人;2-所有
    '出参:cllPatiDatas_Out-返回病人信息集
    '
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 18:02:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Collection, cllExpend As Collection, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strJsonTemp As String
    
    On Error GoTo errHandle
    
    Set cllPatiDatas_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not cllOtherFindCons Is Nothing Then
        For i = 1 To cllOtherFindCons.count
            Select Case UCase(cllOtherFindCons(i)(0))
            Case "病人IDS"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_ids", cllOtherFindCons(i)(1), Json_Text)
            Case "姓名"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_name", cllOtherFindCons(i)(1), Json_Text)
            Case "门诊号"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("outpatient_num", Trim(cllOtherFindCons(i)(1)), Json_Text)
            Case "身份证号", "二代身份证", "身份证"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_idcard", cllOtherFindCons(i)(1), Json_Text)
            Case "联系人身份证"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("contacts_idcard", cllOtherFindCons(i)(1), Json_Text)
            Case "医保号", "医保证号"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_num", cllOtherFindCons(i)(1), Json_Text)
            Case "医疗卡类别ID", "卡类别ID"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(cllOtherFindCons(i)(1)), Json_num, True)
            Case "卡号"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("card_no", cllOtherFindCons(i)(1), Json_Text)
            Case "二维码"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("qrcode", cllOtherFindCons(i)(1), Json_Text)
            Case "IC卡号", "IC", "IC卡"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("iccard_no", cllOtherFindCons(i)(1), Json_Text)
            Case "查询住院状态", "住院状态"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("qrspt_statu", Val(cllOtherFindCons(i)(1)), Json_num, True)
            Case "手机号", "手机"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("phone_number", cllOtherFindCons(i)(1), Json_Text)
            Case "就诊卡号", "就诊卡"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("visit_card", cllOtherFindCons(i)(1), Json_Text)
            Case Else
                strErrMsg = "目前暂不不支持按类别为【" & UCase(cllOtherFindCons(i)(0)) & "】来查找病人！"
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
     
    'zl_PatiSvr_GetPatiInfo
    '  --功能:获取病人信息
    '  --入参：Json_In:格式
    '  --    input
    '  --      pati_id           N 1 病人id  病人ID<>0时，查询列表中的条件无效
    '  --      query_type        N 1 查询类型:如：0-基本;1-基本+联系人;3-所有
    '  --      query_card        N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
    '  --      query_family      N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
    '  --      query_drug        N 1 是否包含过敏药物:1-包含，0-不包含
    '  --      query_immune      N 1 是否包含免疫修:1-包含;0-不包含
    '  --      query_insurance_pwd C  是否包含医保密码:1-包含;0-不包含
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
    '  --        pati_bed        C   当前床号
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_type", int查询类型, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_family", IIf(bln包含家属, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_drug", IIf(bln包含过敏药物, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_immune", IIf(bln包含免疫信息, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_card", IIf(bln包含卡信息, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_insurance_pwd", IIf(bln包含医保密码, 1, 0), Json_num)
    strJson = strJson & "," & GetNodeString("query_cons_list") & ":{" & strJsonTemp & "}"
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetPatiInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    
    
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
    '    pati_email           C   1   email
    '    pati_qq              C   1   qq
    '    card_captcha         C   1  卡验证码
    '    insurance_pwd        C       医保密码
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
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg = "" Then
            strErrMsg = "未找到符合条件的病人信息，请检查！"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllData = objServiceCall.GetJsonListValue("output.pati_list")
    If bln包含家属 Or bln包含过敏药物 Or bln包含免疫信息 Or bln包含卡信息 Then
        For i = 1 To cllData.count
            If bln包含家属 Then
                cllData(i).Remove "_family_list"
                Set cllExpend = objServiceCall.GetJsonListValue("output.pati_list[" & i - 1 & "].family_list")
                If Not cllExpend Is Nothing Then
                    cllData(i).Add cllExpend, "_family_list"
                End If
            End If
            
            If bln包含过敏药物 Then
                cllData(i).Remove "_drug_list"
                Set cllExpend = objServiceCall.GetJsonListValue("output.pati_list[" & i - 1 & "].drug_list")
                If Not cllExpend Is Nothing Then
                    cllData(i).Add cllExpend, "_drug_list"
                End If
            End If
            
            If bln包含免疫信息 Then
                cllData(i).Remove "_immune_list"
                Set cllExpend = objServiceCall.GetJsonListValue("output.pati_list[" & i - 1 & "].immune_list")
                If Not cllExpend Is Nothing Then
                    cllData(i).Add cllExpend, "_immune_list"
                End If
            End If
            
            If bln包含卡信息 Then
                cllData(i).Remove "_card_list"
                Set cllExpend = objServiceCall.GetJsonListValue("output.pati_list[" & i - 1 & "].card_list")
                If Not cllExpend Is Nothing Then
                    cllData(i).Add cllExpend, "_card_list"
                End If
            End If
        Next
    End If
    'Set clldata = objServiceCall.GetJsonListValue("output.pati_list[0].drug_list")
    
'    If cllData Is Nothing Then
'        strErrMsg = "未找到符合条件的病人信息，请检查！"
'        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If
'    If cllData.count = 0 Then
'        strErrMsg = "未找到符合条件的病人信息，请检查！"
'        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If
    
    Set cllPatiDatas_Out = cllData
    zl_PatiSvr_GetPatiInfo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_PatiSvr_GetPatiExtendInfo(ByVal lng病人ID As Long, ByVal str信息名集 As String, ByRef cllPatiData_Out As Collection, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg As String, Optional ByVal lng就诊ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息从表信息服务接口
    '入参:str信息名集-多个用逗号分离,如：医学警示,联系人2,联系人3等
    '
    '出参:cllPatiData_Out-返回病人从表信息数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 20:10:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    
     
    On Error GoTo errHandle
    
   
    Set cllPatiData_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
       
    'zl_PatiSvr_GetPatiExtendInfo
    'input
    '    pati_id N   1   病人id
    '    info_names  C   1   信息名：多个用逗号
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("info_names", str信息名集, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("visit_id", lng就诊ID, Json_num, True)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetPatiExtendInfo"
   
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    
    'output
    '    code    N   1   应答码：0-失败；1-成功
    '    message C   1   "应答消息：成功时返回成功信息,失败时返回具体的错误信息 ""
    '    slave_list[]    C       从表项信息列表
    '       info_name   C   1   信息名
    '        info_value  N   1   信息值
    '        visit_id        n 1 就诊ID
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg = "" Then
            strErrMsg = "未找到符合条件的病人信息，请检查！"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set cllPatiData_Out = objServiceCall.GetJsonListValue("output.slave_list")
    zl_PatiSvr_GetPatiExtendInfo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Patisvr_GetPatiCardInfo(ByVal strCardTypeIDs As String, ByVal str病人ID As String, _
                Optional ByVal intQueryType As Integer = 1, Optional ByVal blnOnlyCardTypeID As Boolean, _
                Optional ByVal strCardTypes As String, Optional ByRef cllData As Collection, _
                Optional ByVal blnCertCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从给定卡类别中检索指定病人持有有效卡的卡类别
    '入参: strCardTypeIDs 给定卡类别，多个用逗号分隔
    '      intQueryType:查询基本类型:0-只获取病人ID,1-只获取卡类别ID;2-包含病人基本信息;3-所有
    '返回:返回病人持有有效卡的卡类别，多个用逗号分隔
    '编制:刘兴洪
    '日期:2018-12-03 15:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllTemp As Collection
    Dim objServiceCall As Object
    Dim strCards As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
   
    'Zl_Patisvr_Getpaticardinfo
    'input
    '    pati_ids       C   1   病人id，多个用英文的逗号分隔
    '    cardtype_ids   C   1   卡类别IDs,多个用逗号分离
    '    card_no        C       卡号
    '    cert_cardtype  N   1   只读取证件作为卡类别的医疗卡类别:1-只读取是否证件=1的医疗卡;0-全部读取
    '    query_type     N   1   查询基本类型:0-只获取病人ID,1-只获取卡类别ID;2-包含病人基本信息;3-所有
    '    dffective_cardtype  N       只读取有效的卡类别:1-只读取有效的卡类别;0-全部读取
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_ids", str病人ID, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("cardtype_ids", strCardTypeIDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("query_type", intQueryType, Json_num)
    strJson = strJson & "," & GetJsonNodeString("dffective_cardtype", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cert_cardtype", IIf(blnCertCard, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    strServiceName = "Zl_Patisvr_Getpaticardinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    'output
    '    code    N   1   应答码：0-失败；1-成功
    '    message C   1   "应答消息：成功时返回成功信息
    '    失败时返回具体的错误信息 ""
    '    card_list[] C       病人卡信息列表
    '        pati_id N   1   病人id
    '        pati_name   C       姓名
    '        pati_sex    C       性别
    '        pati_age    C       年龄
    '        pati_birthdate  C       出生日期：yyyy-mm-dd hh24:mi:ss
    '        outpatient_num  C       门诊号
    '        pati_idcard C       身份证号
    '        cardtype_id N   1   卡类别ID
    '        card_no C   1   卡号
    '        card_qrcode C   1   二维码
    '        card_passwod    C   1   密码
    '        cardtype_name   C   1   卡类别名称
    '        cardtype_cardlen    N   1   卡号长度
    '        card_statu  N   1   状态:0-正常有效卡;1-已挂失; 2-补卡停用
    '        loscard_creator C   1   挂失人
    '        loscard_time    C   1   挂失时间:yyyy-mm-dd hh24:mi:ss
    '        loscard_mode    C   1   挂失方式
    '        sendcard_oper   C   1   发卡人
    '        end_time    C   1   终止使用时间:yyyy-mm-dd hh24:mi:ss
    Set cllData = objServiceCall.GetJsonListValue("output.card_list")
    If cllData Is Nothing Then Exit Function
    
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        If InStr(strCards & ",", "," & cllTemp("_cardtype_id") & ",") = 0 Then
            strCards = strCards & "," & cllTemp("_cardtype_id")
        End If
    Next
    strCardTypes = Mid(strCards, 2)
    Zl_Patisvr_GetPatiCardInfo = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function


Public Function zl_PatiSvr_GetInsureByPatiID(lng病人ID As Long, Optional ByRef int险类_Out As Integer, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断医保病人是否存在未结费用
    '入参:lng病人ID
    '     blnNotShowErrMsg-是否不显示错误信息
    '出参:int险类_Out-险类
    '     strErrMsg_out-返回的错误信息值
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 16:40:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData  As Collection, cllTemp  As Collection
    Dim objServiceCall As Object
    Dim intReturn As Integer
 
    
    On Error GoTo ErrH
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
     
    'zl_PatiSvr_GetInsureByPatiID
    '    入参 json
    '    input
    '    pati_id N   1   病人id
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_ids", lng病人ID, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetInsureByPatiID"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, Not blnNotShowErrMsg) = False Then Exit Function
    '    出参 json
    '    output
    '    code    N   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    '    insurance_type  N   1   险类
    '
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
                strErrMsg_Out = "未找到符合条件的病人信息，请检查！"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    int险类_Out = objServiceCall.GetJsonNodeValue("output.insurance_type")
    zl_PatiSvr_GetInsureByPatiID = True
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function


Public Function zl_PatiSvr_CheckOutNoIsExist(ByVal lng病人ID As Long, ByVal str门诊号 As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查门诊号是否被他人使用
    ' 入参 : str门诊号-传入检查的门诊号
    ' 出参 : blnUsedByOther:T:被别人使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim intReturn As Integer
    
    On Error GoTo ErrHand
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("outpatient_num", str门诊号, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_CheckOutNoIsExist", strJson, , , glngModul, Not blnNotShowErrMsg) = False Then Exit Function
    '    出参 json
    '    output
    '    code    N   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    '    insurance_type  N   1   险类
    '
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "未找到符合条件的病人信息，请检查！"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    blnUsedByOther = Val(objServiceCall.GetJsonNodeValue("output.isexist")) <> 0
    
    zl_PatiSvr_CheckOutNoIsExist = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_PatiSvr_PhoneNumberExist(ByVal lng病人ID As Long, ByVal str手机号 As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查手机号是否被他人使用
    ' 入参 :
    ' 出参 : blnUsedByOther:T:被别人使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim intReturn As Integer
    
    On Error GoTo ErrHand
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("phone_number", str手机号, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_PhoneNumberExist", strJson, , , glngModul, Not blnNotShowErrMsg) = False Then Exit Function
    '    出参 json
    '    output
    '    code    N   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    '    insurance_type  N   1   险类
    '
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "未找到符合条件的病人信息，请检查！"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    blnUsedByOther = Val(objServiceCall.GetJsonNodeValue("output.exist")) <> 0
    
    zl_PatiSvr_PhoneNumberExist = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_PatiSvr_CheckInsNoIsExist(ByVal str医保号 As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查医保号是否被他人使用
    ' 入参 :
    ' 出参 : blnUsedByOther:T:被别人使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim intReturn As Integer
    
    On Error GoTo ErrHand
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = GetJsonNodeString("insurance_num", str医保号, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_CheckInsNoIsExist", strJson, , , glngModul, Not blnNotShowErrMsg) = False Then Exit Function
    '    出参 json
    '    output
    '    code    N   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    '    insurance_type  N   1   险类
    '
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    blnUsedByOther = Val(objServiceCall.GetJsonNodeValue("output.exist")) <> 0
    
    zl_PatiSvr_CheckInsNoIsExist = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Patisvr_GetPatiFamilyMember(ByVal byt查询类型 As Byte, ByVal lng病人ID As Long, _
    ByRef str家属IDs As String, Optional ByRef rs家属信息 As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID，获取该病人的家属成员信息
    '入参:
    '   byt查询类型=查询类型：0-只返回家属成员病人id；1-查询家属成员的基本信息
    '出参:
    '   str家属IDs=家属成员病人id,多个英文逗号分隔
    '   rs家属信息=家属成员的基本信息,仅 查询类型=1 是有效,字段：病人ID,关系,姓名,性别,年龄,出生日期,民族,身份证号
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, cllData As Collection
    Dim cllTemp As Collection, i As Integer
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    str家属IDs = ""
    If byt查询类型 = 1 Then
        Set rs家属信息 = New ADODB.Recordset
        With rs家属信息.fields
            .Append "病人ID", adBigInt, 18, adFldIsNullable
            .Append "关系", adVarChar, 30, adFldIsNullable
            .Append "姓名", adVarChar, 100, adFldIsNullable
            .Append "性别", adVarChar, 4, adFldIsNullable
            .Append "年龄", adVarChar, 20, adFldIsNullable
            .Append "出生日期", adVarChar, 30, adFldIsNullable
            .Append "民族", adVarChar, 20, adFldIsNullable
            .Append "身份证号", adVarChar, 18, adFldIsNullable
        End With
        rs家属信息.CursorLocation = adUseClient
        rs家属信息.LockType = adLockOptimistic
        rs家属信息.CursorType = adOpenStatic
        rs家属信息.Open
    End If
    
    If lng病人ID = 0 Then Zl_Patisvr_GetPatiFamilyMember = True: Exit Function
    'Zl_Patisvr_Getpatifamilymember
    '  --功能：根据病人ID，获取该病人的家属成员信息
    '  --入参：Json_In:格式
    '  --  input
    '  --    pati_id              N   1  病人ID
    '  --    query_type           N   1  查询类型：0-只返回家属成员病人id；1-查询家属成员的基本信息
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                 N   1   应答吗：0-失败；1-成功
    '  --    message              C   1   应答消息：失败时返回具体的错误信息
    '  --    family_list[]        C       家属成员:病人家属
    '  --      pati_id            N   1   家属病人ID
    '  --      pati_relation      C   1   关系
    '  --      pati_name          C   1   姓名
    '  --      pati_sex           C   1   性别
    '  --      pati_age           C   1   年龄
    '  --      pati_birthdate     C   1   出生日期：yyyy-mm-dd hh24:mi:ss
    '  --      pati_nation        C   1   民族
    '  --      pati_idcard        C   1   身份证号
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", byt查询类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Patisvr_Getpatifamilymember", strJson, , "", glngModul) = False Then Exit Function
    Set cllData = objServiceCall.GetJsonListValue("output.family_list")
    If cllData Is Nothing Then Exit Function
    
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        str家属IDs = str家属IDs & "," & cllTemp("_pati_id")
        If byt查询类型 = 1 Then
            With rs家属信息
                .AddNew
                !病人ID = cllTemp("_pati_id")
                !关系 = cllTemp("_pati_relation")
                !姓名 = cllTemp("_pati_name")
                !性别 = cllTemp("_pati_sex")
                !年龄 = cllTemp("_pati_birthdate")
                !出生日期 = cllTemp("_pati_birthdate")
                !民族 = cllTemp("_pati_nation")
                !身份证号 = cllTemp("_pati_idcard")
            End With
        End If
    Next
    If byt查询类型 = 1 Then
        If rs家属信息.RecordCount > 0 Then rs家属信息.MoveFirst
    End If
    If str家属IDs <> "" Then str家属IDs = Mid(str家属IDs, 2)
    Zl_Patisvr_GetPatiFamilyMember = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function Zl_Patisvr_UpdateOutpatiState(ByVal lng病人ID As Long, ByVal cllUpdateInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新病人就诊状态及部分信息
    '入参:
    '   lng病人ID=病人ID
    '   cllUpdateInfo=更新信息，成员:(手机号(N),费别(N),就诊诊室(N),就诊状态(N),就诊时间(N),门诊号(N))
    '出参:
    '返回:执行成功返回True，失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lng病人ID = 0 Then Exit Function
    If cllUpdateInfo Is Nothing Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Patisvr_Updateoutpatistate
    '  --功能：更新门诊病人就诊状态
    '  --      可以判断结点是存在，如果不存在则不更新，或者更新为原值，目前暂时未用到后续可以基于这个进行扩展
    '  --入参：Json_In:格式
    '  --input
    '  --    pati_id            N 1 病人id
    '  --    phone_number       C   病人手机号
    '  --    fee_category       C   费别
    '  --    visit_room         C   更新的就诊诊室
    '  --    visit_status       N   更新的就诊状态
    '  --    visit_time         C   更新的就诊时间
    '  --    outpatient_num             C   门诊号
    '  --出参: Json_Out,格式如下
    '  --    output
    '  --        code                    N   1   应答吗：0-失败；1-成功
    '  --        message                 C   1   应答消息：失败时返回具体的错误信息
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    If CollectionExitsValue(cllUpdateInfo, "手机号") Then
        strJson = strJson & "," & GetJsonNodeString("phone_number", cllUpdateInfo("手机号"), Json_Text)
    End If
    If CollectionExitsValue(cllUpdateInfo, "费别") Then
        strJson = strJson & "," & GetJsonNodeString("fee_category", cllUpdateInfo("费别"), Json_Text)
    End If
    If CollectionExitsValue(cllUpdateInfo, "就诊诊室") Then
        strJson = strJson & "," & GetJsonNodeString("visit_room", cllUpdateInfo("就诊诊室"), Json_Text)
    End If
    If CollectionExitsValue(cllUpdateInfo, "就诊状态") Then
        strJson = strJson & "," & GetJsonNodeString("visit_status", cllUpdateInfo("就诊状态"), Json_num)
    End If
    If CollectionExitsValue(cllUpdateInfo, "就诊时间") Then
        strJson = strJson & "," & GetJsonNodeString("visit_time", cllUpdateInfo("就诊时间"), Json_Text)
    End If
    If CollectionExitsValue(cllUpdateInfo, "门诊号") Then
        strJson = strJson & "," & GetJsonNodeString("outpatient_num", cllUpdateInfo("门诊号"), Json_Text)
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Patisvr_Updateoutpatistate", strJson, , "", glngModul) = False Then Exit Function
    
    Zl_Patisvr_UpdateOutpatiState = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_PatiSvr_GetPatiIdsByRange(ByVal strCondition As String, ByRef strPatiIds As String, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String, _
    Optional ByVal blnFindByFilter As Boolean, Optional ByVal cllFilter As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据条件值获取符合条件的病人ID
    '入参:
    '   strCondition=可能是就诊卡号、身份证号、IC卡号、门诊号
    '   blnFindByFilter=True:按过滤条件(cllFilter)获取;False:按strCondition获取
    '   cllFilter=过滤条件:Array(Key,Value),Key:合同单位ID
    '出参:
    '返回:执行成功返回True，失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intReturn As Integer
    Dim objServiceCall As Object
    Dim strJson As String, i As Long
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'zl_PatiSvr_GetPatiIdsByRange
    '  --入参：Json_In:格式
    '  --input
    '  --    query_condition C 1 查询条件
    '  --    ctt_unit_id     N 1 合同单位ID，查询指定合同单位的门诊病人
    '  --出参: Json_Out,格式如下
    '  --    output
    '  --        code                    N   1   应答吗：0-失败；1-成功
    '  --        message                 C   1   应答消息：失败时返回具体的错误信息
    '  --        pati_ids                C   1   病人IDs
    strJson = ""
    If blnFindByFilter = False Then
        strJson = strJson & "" & GetJsonNodeString("query_condition", strCondition, Json_Text)
    Else
        For i = 1 To cllFilter.count
            Select Case cllFilter(i)(0)
            Case "合同单位ID"
                strJson = strJson & "" & GetJsonNodeString("ctt_unit_id", cllFilter(i)(1), Json_num)
            End Select
        Next
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_GetPatiIdsByRange", strJson, , "", glngModul) = False Then Exit Function
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "未找到符合条件的病人信息，请检查！"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strPatiIds = objServiceCall.GetJsonNodeValue("output.pati_ids")
    zl_PatiSvr_GetPatiIdsByRange = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_PatiSvr_GetInputItemLength(ByVal strTableItems As String, ByRef cllColumn As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人信息中指定字段的最大长度
    ' 入参 : strTableItems：表1:列1,列2|表2:列1,列2,列3|..
    ' 出参 : cllcolumn (Collect):成员(表名,字段,字段长度)
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/21 20:16
    '---------------------------------------------------------------------------------------
    Dim strJson As String, strSubJson As String, i As Long, strServiceName  As String
    Dim varData As Variant, varTmp As Variant
    Dim objServiceCall As Object
    
    On Error GoTo ErrHandler
    If strTableItems = "" Then Exit Function
    
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Patisvr_Getinsureinfo
    '--input
    '--  item_list[]
    '--  table_name  C   1   表名
    '--  column_name C   1   列名,多个用逗号
    varData = Split(strTableItems, "|")
    
    For i = 0 To UBound(varData)
        varTmp = Split(varData(i) & ":", ":")
        strSubJson = strSubJson & ",{"
        strSubJson = strSubJson & "" & GetJsonNodeString("table_name", varTmp(0), Json_Text)
        strSubJson = strSubJson & "," & GetJsonNodeString("column_name", varTmp(1), Json_Text)
        strSubJson = strSubJson & "}"
    Next
    strJson = GetNodeString("item_list") & ":[" & Mid(strSubJson, 2) & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    strServiceName = "zl_PatiSvr_GetInputItemLength"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    Set cllColumn = objServiceCall.GetJsonListValue("output.item_list")
    
    zl_PatiSvr_GetInputItemLength = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function Zl_Patisvr_Getinsureinfo(ByVal lng病人ID As Long, ByRef insure As Integer, Optional insureName As String, _
                                                              Optional str医保号 As String, Optional str医保密码 As String, Optional str卡号 As String, _
                                                              Optional str登记时间 As String, Optional lng病种id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id和险类，获取病人的保险信息
    '入参: insure :险类
    '返回:返回病人医保信息
    '编制:
    '日期:2019-11-20 19:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object
    Dim bln仅读险类 As Boolean
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    bln仅读险类 = insure = 0
    'Zl_Patisvr_Getinsureinfo
    'input
    '     pati_id        N   1      病人id
    '     insure_type N           险类：未传入险类时，根据病人id查询病人信息中的险类；传入险类时，根据病人id和险类查询保险帐户中的险类、险类名称、医保号、登记时间、医保密码
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("insure_type", insure, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    strServiceName = "Zl_Patisvr_Getinsureinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    ' output
    '    insure_type C   1   险类
    '    insure_name C   1   险类名称
    '    insure_no   C   1   医保号
    '    card_no        C       卡号
    '    pati_create_time    C   1   病人的登记时间:yyyy-mm-dd hh24:mi:ss
    '    insure_pwd  C   1   医保密码
    insure = Val(NVL(objServiceCall.GetJsonNodeValue("output.insure_type")))
    If Not bln仅读险类 Then
        insureName = objServiceCall.GetJsonNodeValue("output.insure_name")
        str医保号 = objServiceCall.GetJsonNodeValue("output.insure_no")
        str医保密码 = objServiceCall.GetJsonNodeValue("output.insure_pwd")
        str卡号 = objServiceCall.GetJsonNodeValue("output.card_no")
        str登记时间 = objServiceCall.GetJsonNodeValue("output.pati_create_time")
        lng病种id = Val(NVL(objServiceCall.GetJsonNodeValue("output.dz_type_id")))
    End If

    Zl_Patisvr_Getinsureinfo = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function Zl_Patisvr_Getdeptfrombad(ByRef str科室IDs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取床位状况记录里的科室id
    '出参: str科室IDs :科室id字符串,用逗号分隔
    '返回:获取科室ids成功返回True,否则返回False
    '编制:焦博
    '日期:2019-12-25 19:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strServiceName As String
    Dim objServiceCall As Object
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    strServiceName = "Zl_Patisvr_Getdeptfrombad"
    If objServiceCall.CallService(strServiceName, "", , "", glngModul) = False Then Exit Function
    
    'Zl_Patisvr_Getdeptfrombad
    ' output
    ' --  code                  C    1  应答码：0-失败；1-成功
    ' --  message               C    1  应答消息
    ' --  dept_ids              C    1  科室ids串，多个用逗号分隔
    str科室IDs = objServiceCall.GetJsonNodeValue("output.dept_ids")

    Zl_Patisvr_Getdeptfrombad = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function Zl_Patisvr_Checkdepositno(ByVal lng病人ID As Long, ByRef strNo As String, Optional intOcc As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断预交NO是否存在"病人结算异常记录"中
    '出参:strNo-预交单据号
    '        intOcc-病人结算异常记录.操作情景
    '返回:传入的NO号存在"病人结算异常记录"中返回true,否则返回False
    '编制:焦博
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo ErrHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Patisvr_Checkdepositerrorno
    '    input
    '    --   pati_id              N 1 病人ID
    '    --   bill_nos             C 1 病人预交记录.NO
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_nos", strNo, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Patisvr_Checkdepositerrorno"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
     ' --  output
      '  --    code              N 1 应答码：0-失败；1-成功
       ' --    message           C 1 应答消息：失败时返回具体的错误信息
       ' --    bill_nos          C 1 有效的Nos,多个用逗号分隔
       ' --    occasion          N 1 场合：1-医疗卡发放;2-病人信息登记（针对只传一个NO是有效）
    strNo = objServiceCall.GetJsonNodeValue("output.bill_nos")
    intOcc = Val(objServiceCall.GetJsonNodeValue("output.occasion"))
    Zl_Patisvr_Checkdepositno = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_Patisvr_Calc_Age(ByVal lng病人ID As Long, ByVal str出生日期 As String, Optional ByVal str计算日期 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据出生日期计算病人年龄
    '入参:str计算日期-指定的计算年龄的日期
    '出参:
    '返回:计算的年龄
    '编制:李南春
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo ErrHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
'    --  input
    '  --    pati_id            N 1 病人ID
    '  --    birthdate          N 0 出生日期
    '  --    calc_date          N 0 计算日期
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               N 1 应答吗：0-失败；1-成功
    '  --    message            C 1 应答消息：失败时返回具体的错误信息
    '  --    age                C 1  返回:1天以内：X小时[X分钟],1天至1月以内：X天[X小时],1月至1岁以内：X月[X天],1岁至儿童年龄上限：X岁[X月],>=儿童年龄上限：X岁
    '  --                            说明:1天以内，是指按出生日期24小时算;1月以内，是指对天计算；比如7.8日出生，8.8日才算1月;1岁以内，也是对天计算。;“以内”都是指“<”。
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    If IsDate(str出生日期) Then
        strJson = strJson & "," & GetJsonNodeString("birthdate", str出生日期, Json_Text)
    End If
    If IsDate(str计算日期) Then
        strJson = strJson & "," & GetJsonNodeString("calc_date", str计算日期, Json_Text)
    End If
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "zl_Patisvr_Calc_Age"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    zl_Patisvr_Calc_Age = objServiceCall.GetJsonNodeValue("output.age")
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_Patisvr_GetPatiPhoto(ByVal lng病人ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人照片
    '入参:
    '出参:
    '返回:返回病人照片Base64
    '编制:李南春
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo ErrHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
'    --  input
    '  --    pati_id            N 1 病人ID
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               N 1 应答吗：0-失败；1-成功
    '  --    message            C 1 应答消息：失败时返回具体的错误信息
    '  --    pati_photo         C 1 编码:base64
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Patisvr_GetPatiPhoto"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    zl_Patisvr_GetPatiPhoto = objServiceCall.GetJsonNodeValue("output.pati_photo")
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_Patisvr_SavePatiPhoto(ByVal lng病人ID As Long, ByVal strPatiPhoto As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存病人照片
    '入参:
    '出参:
    '返回:返回病人照片Base64
    '编制:李南春
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo ErrHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
'    --  input
    '  --    pati_id            N 1 病人ID.
    '  --    pati_photo         C 1 编码:base64
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               N 1 应答吗：0-失败；1-成功
    '  --    message            C 1 应答消息：失败时返回具体的错误信息
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_photo", strPatiPhoto, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "zl_Patisvr_SavePatiPhoto"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    zl_Patisvr_SavePatiPhoto = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function


Public Function zl_Patisvr_DeletePatiPhoto(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除病人照片
    '入参:
    '出参:
    '返回:
    '编制:李南春
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo ErrHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
'    --  input
    '  --    pati_id            N 1 病人ID.
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               N 1 应答吗：0-失败；1-成功
    '  --    message            C 1 应答消息：失败时返回具体的错误信息
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "zl_Patisvr_DeletePatiPhoto"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    zl_Patisvr_DeletePatiPhoto = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_PatiSvr_GetPatiAddrssInfo(ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
                ByVal str地址类别 As String, ByRef cllAddrList As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人结构化地址信息
    ' 入参 : str地址类别:查询的地址类别：1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址；为0时表示查询所有类型的地址信息
    '        多个用逗号分隔，例如："3,4"
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim cllData As Collection
    
    On Error GoTo ErrHand
    Set cllAddrList = New Collection
    
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页Id, Json_num)
    If IsNumeric(str地址类别) Then
        strJson = strJson & "," & GetJsonNodeString("addr_type", Val(str地址类别), Json_num)
    Else
        strJson = strJson & "," & GetJsonNodeString("addr_types", str地址类别, Json_Text)
    End If

    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_GetPatiAddrssInfo", strJson, , , glngModul) = False Then Exit Function
    Set cllAddrList = objServiceCall.GetJsonListValue("output.addr_list")
    
    zl_PatiSvr_GetPatiAddrssInfo = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

