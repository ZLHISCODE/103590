Attribute VB_Name = "mdlSplitService"
Option Explicit

Public Function zlSplitService_AdviceIsExist(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '根据传入医嘱id返回存在费用记录的医嘱id
    'strInput：医嘱id,医嘱id...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("advice_ids", strInput, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_ExseSvr_AdviceIsExist"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“zl_ExseSvr_AdviceIsExist”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.advice_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_AdviceIsExist = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetRemainMoney(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOut As String) As Boolean
    '取病人余额
    'strInput：费用id,费用id...
    'strOut: 余额,担保额
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
    'Zl_Exsesvr_Getremainmoney
'  ---------------------------------------------------------------------------
'  --input      获取病人费用余额
'  --  pati_id                 N  1  病人ID
'  --  pati_pageid             N  1  主页ID
'  --  insure_account_balance  N  1  医保账户余额
'  --  pati_ids                C  0  病案主页关键信息拼串，病人ID:主页ID,....
'  --  query_type              N  0  查询方式 1-批量查询病人余额，2-批量查询，担保额和适用病人信息
'  --output
'  --  code                    C  1  应答码：0-失败；1-成功
'  --  message                 C  1  应答消息
'  --  remain_money            N     剩余款
'  --  guarantee_money         N     担保额
'  --  expected_money          N     预结费用
'  --  prepay_money            N  0  预交余额
'  --  item_list[]当传入批量病人信息时才返回，该列表可以不返回
'  --       pati_id            N 1 病人id
'  --       pati_pageid        N 1 主页id
'  --       remain_money       N 1 剩余款
'  --       guarantee_money    N 1 担保额
'  --       pati_type          C 1 适用病人
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageid", Split(strInput, ",")(1), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_Getremainmoney"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Exsesvr_Getremainmoney”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")

    strOut = objServiceCall.GetJsonNodeValue("output.pati_type")
    strOut = strOut & "," & objServiceCall.GetJsonNodeValue("output.remain_money")
    strOut = strOut & "," & objServiceCall.GetJsonNodeValue("output.guarantee_money")
        
    zlSplitService_GetRemainMoney = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiPageWarnScheme(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOut As String) As Boolean
    '取病人类型
    'strInput：病人id,主页id...
    'strOut: 病人类型
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
'  --获取病人获取病人费用余额
'  ---------------------------------------------------------------------------
'  --input      获取病人费用余额
'  --  pati_id  N  1  病人ID
'  --  pati_pageid  N  1  主页ID
'  --output
'  --  code  C  1  应答码：0-失败；1-成功
'  --  message  C  1  "应答消息：
'  --  pati_type  C    适用病人类型
'  --  remain_money    N    剩余款
'  --  guarantee_money   N   担保额
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageid", Split(strInput, ",")(1), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatpagewarnscheme"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Cissvr_Getpatpagewarnscheme”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")

    strOut = objServiceCall.GetJsonNodeValue("output.pati_type")
        
    zlSplitService_GetPatiPageWarnScheme = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiPage(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String, Optional ByRef colOutListBaby As Collection, _
    Optional ByVal bln含床位信息 As Boolean, Optional ByRef colOutListBad As Collection) As Boolean
    '取病案主页信息
    'strInput：病人id:主页id,...
    '出参:
    '   colOutListBad 床位信息集合，成员：病人ID,病区ID,病区名称,床号,分类编码,分类名称=Key(_病人ID_病区ID_床号)
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    Dim colTmp As New Collection, colBaby As New Collection
    Dim i As Integer, n As Integer
    Dim colBads As Collection
    
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
'  --      pati_natures        C 1 病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人；多个逗号分隔，不传为所有
'  --      rgst_id             N 1 挂号ID,根据挂号ID查询
'  --      is_badinfo          N 1 是否包含床位信息:1-包含;0-不包含
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
'  --      badinfo_list[]              床位信息，[数组]
'  --        wardarea_id       N 1 病区id
'  --        wardarea_name     C 1 病区名称
'  --        bed_no            C 1 床号
'  --        bed_class_code    C 1 分类编码
'  --        bed_class_name    C 1 分类名称
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageids", strInput, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_babyinfo", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_transdeptinfo", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_lastpage", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_badinfo", IIf(bln含床位信息, 1, 0), 1)
    
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
    Set colOutListBad = New Collection
    
    '婴儿数据的病人ID,主页id相同，婴儿序号不同，要么不用key，要么用 病人id+主页id+婴儿序号 作为key
    For n = 1 To colOutlist.count
        Set colBaby = objServiceCall.GetJsonListValue("output.page_list[" & n - 1 & "].baby_list")
        For i = 1 To colBaby.count
            Set colTmp = New Collection
            colTmp.Add colBaby(i)("_pati_id"), "_pati_id"
            colTmp.Add colBaby(i)("_pati_pageid"), "_pati_pageid"
            colTmp.Add colBaby(i)("_baby_num"), "_baby_num"
            colTmp.Add colBaby(i)("_baby_name"), "_baby_name"
            colTmp.Add colBaby(i)("_baby_sex"), "_baby_sex"
            colTmp.Add colBaby(i)("_baby_date"), "_baby_date"
            colOutListBaby.Add colTmp, "_" & colBaby(i)("_pati_id") & "_" & colBaby(i)("_pati_pageid") & "_" & colBaby(i)("_baby_num")
        Next
        
        If bln含床位信息 Then
            '  --      badinfo_list[]              床位信息，[数组]
            '  --        wardarea_id       N 1 病区id
            '  --        wardarea_name     C 1 病区名称
            '  --        bed_no            C 1 床号
            '  --        bed_class_code    C 1 分类编码
            '  --        bed_class_name    C 1 分类名称
            Set colBads = objServiceCall.GetJsonListValue("output.page_list[" & n - 1 & "].badinfo_list")
            For i = 1 To colBads.count
                Set colTmp = New Collection
                colTmp.Add colOutlist(n)("_pati_id"), "病人ID"
                colTmp.Add colBads(i)("_wardarea_id"), "病区ID"
                colTmp.Add colBads(i)("_wardarea_name"), "病区名称"
                colTmp.Add colBads(i)("_bed_no"), "床号"
                colTmp.Add colBads(i)("_bed_class_code"), "分类编码"
                colTmp.Add colBads(i)("_bed_class_name"), "分类名称"
                'colOutListBad 床位信息集合，成员：病人ID,病区ID,病区名称,床号,分类编码,分类名称=Key(_病人ID_病区ID_床号)
                colOutListBad.Add colTmp, "_" & colOutlist(n)("_pati_id") & "_" & colBads(i)("_wardarea_id") & "_" & colBads(i)("_bed_no")
            Next
        End If
    Next
    
    zlSplitService_GetPatiPage = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiName(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal colInput As Collection, _
    ByRef colPati As Collection, Optional ByVal strKey As String, Optional ByVal intQueryType As Integer = 3) As Boolean
    '用于按一定条件单独查询病人信息
    '目前支持的查询条件，一般都是按其中一种查询：病人ID，门诊号，姓名，就诊卡号，医保号，床号
    'colInput：查询条件组合，Json中input各节点作为元素的KEY值，集合某元素为空表示该节点值为空
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    'Zl_Patisvr_Getpatiinfo
'  ---------------------------------------------------------------------------
'  --功能:获取病人信息
'  --入参：Json_In:格式
'  --    input
'  --      pati_id           N 1 病人id  病人ID<>0时，查询列表中的条件无效
'  --      query_type        N 1 查询类型:如：0-基本;1-基本+联系人;2-所有
'  --      query_card        N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
'  --      query_family      N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
'  --      query_drug        N 1 是否包含过敏药物:1-包含，0-不包含
'  --      query_immune      N 1 是否包含免疫修:1-包含;0-不包含
'  --      query_insurance_pwd C  是否包含医保密码:1-包含;0-不包含
'  --      query_cons_list   C 1 查询条件:可以选择一定条件进行查询（是And关系),只有一行
'  --        pati_ids        C   病人IDs:多个用逗号
'  --        pati_name       C   姓名:可以代%分号表表按姓名匹配
'  --        outpatient_num  C   门诊号
'  --        inpatient_num   C   住院号
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
'  --    insurance_pwd        C       医保密码
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

    '入参
    StrJson_In = ""
    If Not IsNull(colInput("pati_id")) Then
        StrJson_In = GetJsonNodeString("pati_id", colInput("pati_id"), 1)
    Else
        StrJson_In = GetJsonNodeString("pati_id", 0, 1)
    End If
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", intQueryType, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_card", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_family", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_drug", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_immune", 0, 1)
    
    If IsNull(colInput("pati_id")) Then
        '可能按其中一种方式查询：门诊号，姓名，就诊卡号，医保号，床号，身份证号
        If Not IsNull(colInput("outpatient_num")) Then strJson_List = strJson_List & "," & GetJsonNodeString("outpatient_num", colInput("outpatient_num"), 0)
        If Not IsNull(colInput("pati_name")) Then strJson_List = strJson_List & "," & GetJsonNodeString("pati_name", colInput("pati_name"), 0)
        If Not IsNull(colInput("pati_vcard_no")) Then strJson_List = strJson_List & "," & GetJsonNodeString("visit_card", colInput("pati_vcard_no"), 0)
        If ExistsColObject(colInput, "insurance_num") Then
            If Not IsNull(colInput("insurance_num")) Then strJson_List = strJson_List & "," & GetJsonNodeString("insurance_num", colInput("insurance_num"), 0)
        End If
        If Not IsNull(colInput("pati_bed")) Then strJson_List = strJson_List & "," & GetJsonNodeString("pati_bed", colInput("pati_bed"), 0)
        If ExistsColObject(colInput, "pati_idcard") Then strJson_List = strJson_List & "," & GetJsonNodeString("pati_idcard", nvl(colInput("pati_idcard")), 0)
        
        strJson_List = strJson_List & "," & GetJsonNodeString("qrspt_statu", 2, 1)
        
        strJson_List = Mid(strJson_List, 2)
    End If
  
    strJson_List = ",""query_cons_list"":{" & strJson_List & "}"
    
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
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiId(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOutPut As String) As Boolean
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
    
    On Error GoTo errHandle
    
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
    strOutPut = objServiceCall.GetJsonNodeValue("output.pati_id")
    
    zlSplitService_GetPatiId = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetBillOperControls(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOutPut As String) As Boolean
    '获取单据操作控制数据
    'strInput：人员id
    'strOutPut：是否控制|时间限制|是否允许他人单据|金额上限
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
    'Zl_Exsesvr_Getbillopercontrols
'  ---------------------------------------------------------------------------
'  --功能:获取单据操作控制数据
'  --入参：Json_In:格式
'  --  input
'  --    bill_type  N  1  单据类型:1-挂号单据,2-收费单,3-划价单,4-门诊记帐,5-住院记帐,6-预交款,7-结帐单据,8-就诊卡,9-处方
'  --    operator_id  N  1  人员ID
'
'  --出参: Json_Out,格式如下
'  --   output
'  --    code  C  1  应答码：0-失败；1-成功
'  --    message  C  1  应答消息：失败时返回具体的错误信息
'  --    is_exist    N  1  存在控制数据:1-存在;0-不存在
'  --    time_limit  N  1  0(NULL)-不限制,n-n天内
'  --    other_bill  N  1  是否允许对其它单据进行操作
'  --    uplimit_money  N  1  金额上线
'
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '入参
    StrJson_In = ""
    
    StrJson_In = StrJson_In & "" & GetJsonNodeString("bill_type", 9, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_id", strInput, 1)
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_Getbillopercontrols"

    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Exsesvr_Getbillopercontrols”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数据
    strOutPut = objServiceCall.GetJsonNodeValue("output.is_exist") & "|" & _
        objServiceCall.GetJsonNodeValue("output.time_limit") & "|" & _
        objServiceCall.GetJsonNodeValue("output.other_bill") & "|" & _
        objServiceCall.GetJsonNodeValue("output.uplimit_money")
    
    zlSplitService_GetBillOperControls = True
    
    Exit Function
errHandle:
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
    
    On Error GoTo errHandle
    
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
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetPati(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '取病人信息
    'strInput：项目名称;项目内容
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
        
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
    
    On Error GoTo errHandle
    
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
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetCardTypes(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection) As Boolean
    '获取医疗卡类别数据
    'strInput：目前不加条件，查询所有的医疗卡类别
    'strOutPut：返回ID，编码，名称
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    'Zl_Patisvr_Getcardtypes
'  ---------------------------------------------------------------------------
'  --功能:获取医疗卡类别数据
'  --入参：Json_In:格式
'  --    input
'  --      cardtype_id          N   卡类别id:NULL表示不按卡类别ID查找
'  --      query_type           N 1 查询类型:0-所有信息;1-基本信息(返回:id,编码，名称,卡号长度,前缀文本,是否启用,结算方式,是否全退,是否退现)
'  --      cert_cardtype        N   只读取证件作为卡类别的医疗卡类别:1-只读取是否证件=1的医疗卡;0-全部读取
'  --      dffective_cardtype   N   只读取有效的卡类别:1-只读取有效的卡类别;0-全部读取
'  --出参: Json_Out,格式如下
'  --  output
'  --    code                   N   1   应答码：0-失败；1-成功
'  --    message                C   1   应答消息：失败时返回具体的错误信息
'  --    type_list[]            C   1   支持的卡类别列表
'  --        cardtype_id        N   1   ID
'  --        cardtype_code      C   1   编码
'  --        cardtype_name      C   1   名称
'  --        cardtype_stname    C   1   短名
'  --        prefix_text        C   1   前缀文本
'  --        cardno_len         N   1   卡号长度
'  --        default            N   1   缺省标志
'  --        fixed              N   1   是否固定:1-是系统固定;0-不是系统固定
'  --        strict             N   1   是否严格控制:1-是严格控制;0-不是严格控制
'  --        self_make          N   1   是否自制:1-是的;0-不是
'  --        exist_account      N   1   是否存在帐户:1-存在帐户;0-不存在账户
'  --        allow_return_cash  N   1   是否退现:1-允许;0-不允许
'  --        must_all_return    N   1   是否全退:1-必需全退;0-允许部分退
'  --        component          C   1   部件
'  --        memo               C   1   备注
'  --        spec_item          C   1   特定项目
'  --        blnc_mode          C   1   结算方式
'  --        blnc_nature        N   1   结算性质
'  --        cardno_pwdtxt      C   1   卡号密文:卡号从第几位至第几位显示密文,格式为:S-N:S表示从第几位开始,至第几位结束.比如:3-10,表示从3位到10位用密文*表示:12********3323主要是适应不同类别的医疗卡
'  --        allow_repeat_use   N   1   是否重复使用:1-允许;0-不允许
'  --        enabled            N   1   是否启用:1-已启用;0-未启用
'  --        pwd_len            N   1   密码长度
'  --        pwd_len_limit      N   1   密码长度限制:0-不作限制;1-固定密码长度;-n表示密码必须输入好多个位密码以上,但不能超过密码长度
'  --        pwd_rule           N   1   密码规则:０-数字和字符组成;1-仅为数字组成
'  --        allow_vaguefind    N   1   是否模糊查找:1-支持模糊查找;0-不支持
'  --        pwd_require        N   1   密码输入限制:0-不限制;1-不输入,提醒;2-不输入禁止;缺省为不限制
'  --        default_pwd        N   1   是否缺省密码:1-以身份证后N(以密码长度为准)位作为缺省密码;0-无缺省密码
'  --        allow_makecard     N   1   是否制卡:1-是;0-否
'  --        allow_sendcard     N   1   是否发卡:1-是;0-否
'  --        allow_writcard     N   1   是否写卡:1-是;0-否
'  --        insurance_type     N   1   险类
'  --        insurance_name     C   1   险类名称
'  --        sendcard_nature    N   1   发卡性质:0-不限制;1-同一病人只能发一张卡;2-同一病人允许发多张卡，但需提示;缺省为0
'  --        allow_transfer     N   1   是否转帐及代扣:1-支持转帐及代扣;0-不支持
'  --        readcard_nature    C   1   读卡性质,医疗卡读卡方式：第一位为:是否刷卡;第二位为是否扫描;第三位是否接触式读卡;第四位是否非接触式读卡。例如刷卡：'1000'
'  --        keyboard_mode      N   1   键盘控制方式:：0-禁止使用软键盘;1-使用数字软键盘 ,2-使用字符软键盘
'  --        advsend_buildqrcode N   1   是否医嘱发送调用条码生成接口:1-发送调用生成二维码接口;0-不调用
'  --        holding_pay         N   1   是否持卡消费:1-是;0-否
'  --        cert_cardtype       N   1   是否证件类型的医疗卡:0-不是；1-是
'  --        verfycard           N   1   是否退款验卡
'  --        sendcard_sign       N   1   发卡控制:0或NULL-发卡时，卡号必须达到卡号长度;1-发卡时，允许卡号小于等于卡号长度,发卡时，小于卡号长度时，不提示操作员;2-发卡时，允许卡号小于等于卡号长度,小于时，提示操作员。
'  --        enterkey_enabled    N   1   设备是否启用回车:医疗卡对应的刷卡设备是否启用了回车，如果启用了回车，则卡号长度默认增加一位来屏蔽回车
'  --        def_return_cash     N   1   是否缺省退现:允许退现时,默认是否退现
'  --        balalone            N   1   是否独立结算:1-独立结算;0-非独立结算
'  --        discern_rule        N   1   卡号识别规则:1-全部转换为大写;0-不区分大小写
'  --        def_valid_time      C   1   缺省有效时间:NULL时，表示不限制;非空时，格式为:时间+单位(天，月),比如：3天,3月
'  --        scanpay             N   1   是否支持扫码付:是否支持扫码付,支持时，会调用“zlReadQRCode部件”
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle

'  --      cardtype_id          N   卡类别id:NULL表示不按卡类别ID查找
'  --      query_type           N 1 查询类型:0-所有信息;1-基本信息(返回:id,编码，名称,卡号长度,前缀文本,是否启用,结算方式,是否全退,是否退现)
'  --      cert_cardtype        N   只读取证件作为卡类别的医疗卡类别:1-只读取是否证件=1的医疗卡;0-全部读取
'  --      dffective_cardtype   N   只读取有效的卡类别:1-只读取有效的卡类别;0-全部读取

    '入参
    StrJson_In = ""
    
    StrJson_In = StrJson_In & "" & GetJsonNodeString("cardtype_id", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", 1, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("cert_cardtype", 0, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("dffective_cardtype", 0, 0)

    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_PatiSvr_GetCardTypes"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“zl_PatiSvr_GetCardTypes”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数据
    Set colOutlist = objServiceCall.GetJsonListValue("output.type_list")
    
    If colOutlist Is Nothing Then zlSplitService_GetCardTypes = False: Exit Function
    
    zlSplitService_GetCardTypes = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallAccountInsert(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strAccInsert As String) As Boolean
    '静配收取材料配置打包费等：产生门诊/住院记录
    Dim i As Integer
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String, strJson_detail As String
    Dim varList As Variant, varNos As Variant
    Dim strPrePati As String, strPati As String, strPatiJson As String
    
    If strAccInsert = "" Then Exit Function
    '服务：Zl_Exsesvr_Newbill
    ' ---------------------------------------------------------------------------
    '  --功能：门诊病人住院病人发送医嘱生成费用单据
    '  --入参：Json_In:格式
    '  --input
    '  --  pati_list[] 病人列表，单个病人时可以无该节点
    '  --    billtype                                            N 1 类型,1-收费单，2-记帐单
    '  --    pati_source                                         N 1 来源，1-门诊，2-住院
    '  --    pati_id                                             N 1 病人id
    '  --    pati_pageid                                         N 1 主页id
    '  --    baby_num                                            N 1 婴儿费
    '  --    sgin_no                                             N 1 标识号，门诊号，住院号
    '  --    bed_num                                             C 1 床号
    '  --    pati_name                                           C 1 姓名
    '  --    pati_sex                                            C 1 性别
    '  --    pati_age                                            C 1 年龄
    '  --    fee_category                                        C 1 费别
    '  --    overtime_sign                                       N 1 加班标志
    '  --    pati_deptid                                         N 1 病人科室id
    '  --    pati_wardarea_id                                    N 1 病人病区id
    '  --    operator_name                                       C 1 操作员姓名
    '  --    operator_code                                       C 1 操作员编号
    '  --    outpati_tag                                         N 1 门诊标志
    '  --    rgst_id                                             N 1 就诊id
    '  --    emg_sign                                            N 1 是否急诊
    '  --    item_list[]明细列表
    '  --        fee_id                                        N 1 费用id
    '  --        fee_no                                        C 1 No
    '  --        serial_num                                    N 1 序号
    '  --        charge_tag                                    N 1 划价
    '  --        placer                                        C 1 开单人
    '  --        plcdept_id                                    N 1 开单部门id
    '  --        sub_serial_num                                N 1 从属父号
    '  --        fitem_id                                      N 1 收费细目id
    '  --        item_type                                     C 1 收费类别
    '  --        unit                                          C 1 计算单位
    '  --        pharmacy_window                               C 1 发药窗口
    '  --        packages_num                                  N 1 付数
    '  --        send_num                                      N 1 数次
    '  --        ext_mark                                      N 1 附加标志
    '  --        exe_deptid                                    N 1 执行部门id
    '  --        price_ftrnum                                  N 1 价格父号
    '  --        income_item_id                                N 1 收入项目id
    '  --        receipt_name                                  C 1 收据费目
    '  --        price                                         N 1 标准单价
    '  --        fee_amrcvb                                    N 1 应收金额
    '  --        fee_ampaib                                    N 1 实收金额
    '  --        happen_time                                   C 1 发生时间
    '  --        create_time                                   C 1 登记时间
    '  --        memo                                          C 1 费用摘要
    '  --        order_id                                      N 1 医嘱序号
    '  --        exe_properties                                N 1 执行性质
    '  --        decoction_method                              C 1 煎法
    '  --        morphology                                    C 1 中药形态
    '  --        bakstuff_batch                                N 1 批次
    '  --        insurance                                     N 1 保险项目否
    '  --        insure_id                                     N 1 保险大类id
    '  --        insure_code                                   C 1 保险编码
    '  --        fee_type                                      C 1 费用类型
    '  --        si_manp_money                                 N 1 统筹金额
    '  --        synchro                                       N 1 更新同步标志
    '  --        effective_time                                N 1 期效
    '  --        receipt_issecret                              N 1 保密
    '  --        takedept_id                                   N 1 领药部门id
    '  --        group_id                                      N 0 医疗小组id
    '  --出参: Json_Out,格式如下
    '  --output
    '  --  code                                                N 1 应答吗：0-失败；1-成功
    '  --  message                                             C 1 应答消息：失败时返回具体的错误信息
    '  ---------------------------------------------------------------------------
    
    '门诊/住院记帐记录_insert
    '主表：rcp_no 0费用记录.NO, pati_id 1病人ID, sgin_no 2标志号, pati_name 3姓名, pati_sex 4性别, pati_age 5年龄,
             '      fee_category 6费别, pati_wardarea_id 7病人病区id, pati_deptid 8病人科室id, bill_deptid  9开单部门id,
             '      placer 10开单人, operator_name 11操作员姓名, operator_code 12操作员编号,
             '      happen_time 13发生时间, create_time 14登记时间, outpati_flag 15门诊标志
                         
    '明细：serial_num 16序号, baby_num 17婴儿费, fitem_id 18收费细目id, item_type 19收费类别, unit 20计算单位
    '      packages_num 21付数, send_num 22数次, income_item_id 23收入项目id, receipt_name 24收据费目, price 25标准单价, fee_amrcvb 26应收金额,
    '      fee_ampaib 27实收金额, exe_deptid 28执行部门id, pati_source 29费用来源, pati_pageid 30主页id, bed_num 31床号
    
    varList = Split(strAccInsert, "|")
    For i = 0 To UBound(varList)
        varNos = Split(varList(i), ",")
        
        If strPrePati <> varNos(1) Then
            If strJson_List <> "" Then
                strPatiJson = IIf(strPatiJson = "", "", strPatiJson & ",") & "{" & strPati & ",""item_list"":[" & strJson_List & "]}"
                strJson_List = ""
            End If
            
            strPrePati = varNos(1)
            
            '取主信息
            strPati = ""
            strPati = strPati & "" & GetJsonNodeString("billtype", 2, 1)
            strPati = strPati & "," & GetJsonNodeString("pati_source", varNos(29), 1)
            strPati = strPati & "," & GetJsonNodeString("pati_id", varNos(1), 1)
            strPati = strPati & "," & GetJsonNodeString("pati_pageid", varNos(30), 1)
            strPati = strPati & "," & GetJsonNodeString("baby_num", varNos(17), 1)
      
            strPati = strPati & "," & GetJsonNodeString("sgin_no", varNos(2), 1)
            strPati = strPati & "," & GetJsonNodeString("bed_num", varNos(31), 0)
            strPati = strPati & "," & GetJsonNodeString("pati_name", varNos(3), 0)
            strPati = strPati & "," & GetJsonNodeString("pati_sex", varNos(4), 0)
            strPati = strPati & "," & GetJsonNodeString("pati_age", varNos(5), 0)
            
            strPati = strPati & "," & GetJsonNodeString("fee_category", IIf(varNos(6) = "", "普通", varNos(6)), 0)
            strPati = strPati & "," & GetJsonNodeString("overtime_sign", 0, 1)
            strPati = strPati & "," & GetJsonNodeString("pati_deptid", varNos(8), 1)
            strPati = strPati & "," & GetJsonNodeString("pati_wardarea_id", varNos(7), 1)
            strPati = strPati & "," & GetJsonNodeString("operator_name", varNos(11), 0)
    
            strPati = strPati & "," & GetJsonNodeString("operator_code", varNos(12), 0)
            strPati = strPati & "," & GetJsonNodeString("outpati_tag", varNos(15), 1)
            strPati = strPati & "," & GetJsonNodeToNull("rgst_id")
            strPati = strPati & "," & GetJsonNodeToNull("emg_sign")
        End If
        
        '取明细信息
        strJson_detail = ""
        '  --        fee_id                                        N 1 费用id
        '  --        fee_no                                        C 1 No
        '  --        serial_num                                    N 1 序号
        '  --        charge_tag                                    N 1 划价
        '  --        placer                                        C 1 开单人
        strJson_detail = strJson_detail & "" & GetJsonNodeToNull("fee_id")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("fee_no", varNos(0), 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("serial_num", varNos(16), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("charge_tag")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("placer", "", 0)
        
        '  --        plcdept_id                                    N 1 开单部门id
        '  --        sub_serial_num                                N 1 从属父号
        '  --        fitem_id                                      N 1 收费细目id
        '  --        item_type                                     C 1 收费类别
        '  --        unit                                          C 1 计算单位
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("plcdept_id")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("sub_serial_num")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("fitem_id", varNos(18), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("item_type", varNos(19), 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("unit", IIf(varNos(20) = "", "次", varNos(20)), 0)
        
        '  --        pharmacy_window                               C 1 发药窗口
        '  --        packages_num                                  N 1 付数
        '  --        send_num                                      N 1 数次
        '  --        ext_mark                                      N 1 附加标志
        '  --        exe_deptid                                    N 1 执行部门id
        strJson_detail = strJson_detail & "," & GetJsonNodeString("pharmacy_window", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("packages_num", varNos(21), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("send_num", varNos(22), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("ext_mark")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("exe_deptid", varNos(28), 1)
                
        '  --        price_ftrnum                                  N 1 价格父号
        '  --        income_item_id                                N 1 收入项目id
        '  --        receipt_name                                  C 1 收据费目
        '  --        price                                         N 1 标准单价
        '  --        fee_amrcvb                                    N 1 应收金额
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("price_ftrnum")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("income_item_id", varNos(23), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("receipt_name", varNos(24), 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("price", varNos(25), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("fee_amrcvb", varNos(26), 1)
        
        '  --        fee_ampaib                                    N 1 实收金额
        '  --        happen_time                                   C 1 发生时间
        '  --        create_time                                   C 1 登记时间
        '  --        memo                                          C 1 费用摘要
        '  --        order_id                                      N 1 医嘱序号
        strJson_detail = strJson_detail & "," & GetJsonNodeString("fee_ampaib", varNos(27), 1)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("happen_time", varNos(13), 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("create_time", varNos(14), 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("memo", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("order_id")
        
        '  --        exe_properties                                N 1 执行性质
        '  --        decoction_method                              C 1 煎法
        '  --        morphology                                    C 1 中药形态
        '  --        bakstuff_batch                                N 1 批次
        '  --        insurance                                     N 1 保险项目否
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("exe_properties")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("decoction_method", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("morphology", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("bakstuff_batch")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("insurance")
        
        '  --        insure_id                                     N 1 保险大类id
        '  --        insure_code                                   C 1 保险编码
        '  --        fee_type                                      C 1 费用类型
        '  --        si_manp_money                                 N 1 统筹金额
        '  --        synchro                                       N 1 更新同步标志
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("insure_id")
        strJson_detail = strJson_detail & "," & GetJsonNodeString("insure_code", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeString("fee_type", "", 0)
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("si_manp_money")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("synchro")
        
        '  --        effective_time                                N 1 期效
        '  --        receipt_issecret                              N 1 保密
        '  --        takedept_id                                   N 1 领药部门id
        '  --        group_id                                      N 0 医疗小组id
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("effective_time")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("receipt_issecret")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("takedept_id")
        strJson_detail = strJson_detail & "," & GetJsonNodeToNull("group_id")
        
        strJson_List = IIf(strJson_List = "", "", strJson_List & ",") & "{" & strJson_detail & "}"
    Next
    
    If strJson_List <> "" Then
        strPatiJson = IIf(strPatiJson = "", "", strPatiJson & ",") & "{" & strPati & ",""item_list"":[" & strJson_List & "]}"
    End If
    StrJson_In = "{""input"":{""pati_list"":[" & strPatiJson & "]}}"
        
    '调用服务
    If objServiceCall.CallService("Zl_Exsesvr_Newbill", StrJson_In, , "", lngMode, False) = False Then
        MsgBox "调用“Zl_Exsesvr_Newbill”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    If strJson_Out = "0" Then
        MsgBox objServiceCall.GetJsonNodeValue("output.message"), vbInformation, gstrSysName
        Exit Function
    End If

    zlSplitService_CallAccountInsert = True
End Function



Public Function zlSplitService_CallAccountDel_Check(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef strMsg As String, Optional ByRef int执行状态 As Integer = 1) As Boolean
    '门诊/住院记录销账检查
    'strInput:服务需要的入参格式：no,已结禁止销帐(1),医保禁止部分销帐(0),操作状态(1),费用来源|序号,销帐数量;序号,销帐数量...|费用id,已发数量;费用id,已发数量|病人id,审核标志,住院状态,病案编目日期;病人id,审核标志,住院状态,病案编目日期...
    '         一级用"|"，除第1部分外的二，三级再用";"，","分隔
    Dim arrPart As Variant
    Dim i As Integer
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String, strJson_Part As String
    Dim strService As String
    Dim strPart1 As String, strPart2 As String, strPart3 As String, strPart4 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String, strJson_In_Part4 As String
    
On Error GoTo errHandle
    
    '服务：Zl_ExseSvr_DelBill_Check 门诊/住院记录通用服务
'  ---------------------------------------------------------------------------
'  --功能：针对指定单据指定行行进行销帐
'  --入参：Json_In:格式
'  --  input
'  --      fee_no                  C   1   费用单据号
'  --      fee_bill_type           N   1   单据性质:2-门诊记帐单,3-自动记帐单
'  --      balance_ban_writeoffs   N   1   已结禁止销帐:如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
'  --      part_ban_writeoffs      N   1   禁止部分销帐:1-不允许；0-允许
'  --      fee_origin              N   1   费用来源（1-门诊记帐，2-住院记帐）
'  --      item_list[]             本次销帐列表
'  --          serial_num          N   1   序号
'  --          quantity            N   1   销帐数量(为零时，按序号直接销帐)
'  --      excute_list[]           单据已执行列表(药品、卫材费用),即使已执行数为0也要传入
'  --          fee_id              N   1   费用ID
'  --          sended_num          N   1   已发数量
'  --      advice_excute_list[]    单据已执行列表(医嘱费用),即使已执行数为0也要传入
'  --          advice_id           N   1   医嘱ID
'  --          fee_item_id         N   1   收费细目ID
'  --          execute_num         N   1   已执行数
'  --      pati_list[]             病人信息，仅审核这些病人的费用
'  --          pati_id             N   1   病人ID
'  --          fee_audit_status    N   1   费用审核标志:0或空-未审核;1-已审核或开始审核(结合参数:病人审核方式来控制);2-完成审核,结合结帐权限[禁止未审核病人结帐]进行管理控制
'  --          si_inp_status       N   1   住院状态:0-正常住院;1-尚未入科;2-正在转科或正在转病区;3-已预出院
'  --          catalog_date        C   0   病案编目日期：yyyy-mm-dd hh24:mi:ss
'  --出参: Json_Out,格式如下
'  --  output
'  --      code                    N   1   应答吗：0-失败；1-成功
'  --      message                 C   1   应答消息：失败时返回具体的错误信息
'  --      item_list[]                         单据数据列表
'  --          serial_num          N   1   序号
'  --          quantity            N   1   销帐数量
'  --          execute_tag         N   1   执行状态：0-未执行;1-已执行;2-部分执行
'  ---------------------------------------------------------------------------

    If strInput = "" Then zlSplitService_CallAccountDel_Check = True: Exit Function
    
    strJson_List = ""
    
    arrPart = Split(strInput, "|")
    
    strPart1 = arrPart(0)
    strPart2 = arrPart(1)
    strPart3 = arrPart(2)
    strPart4 = arrPart(3)
        
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
    
    strJson_List = ""
    arrPart = Split(strPart4, ";")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("pati_id", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("fee_audit_status", Split(arrPart(i), ",")(1), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("si_inp_status", Split(arrPart(i), ",")(2), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("catalog_date", Split(arrPart(i), ",")(3), 0)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    strJson_In_Part4 = ",""pati_list"":[" & strJson_List & "]"
    
    '汇总
    StrJson_In = "{""input"":{" & strJson_In_Part1 & strJson_In_Part2 & strJson_In_Part3 & strJson_In_Part4 & "}}"
    
    strService = "Zl_ExseSvr_DelBill_Check"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_ExseSvr_DelBill_Check”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        strMsg = objServiceCall.GetJsonNodeValue("output.message")
        zlSplitService_CallAccountDel_Check = False
        Exit Function
    End If
    
    int执行状态 = objServiceCall.GetJsonNodeValue("output.item_list[0].execute_tag")

    zlSplitService_CallAccountDel_Check = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallCancelAccCheck(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String) As Boolean
    '销账核查
    'strInput:核查人,核查时间,申请类别|费用id,申请时间;费用id,申请时间;...
    Dim arrPart As Variant
    Dim i As Integer
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String, strJson_Part As String
    Dim strService As String
    Dim strPart1 As String, strPart2 As String
    
On Error GoTo errHandle
    
    '服务：Zl_Exsesvr_Cancelacc_Check
'  ---------------------------------------------------------------------------
'  --input     费用销帐记录核查
'  --  check_people  C 1 核查人
'  --  check_time    D 1 核查时间
'  --  request_type  N   申请类别：0-未执行;1-已执行;非药品和卫材固定存为0
'  --  rcpdtl_list     [数组]每个处方明细信息
'  --    rcpdtl_id     N 1 处方明细id(费用id)
'  --    request_time  D 1 申请时间
'  --output
'  --  code         C 1 应答码：0-失败；1-成功
'  --  message      C 1 应答消息：
'  ---------------------------------------------------------------------------

    If strInput = "" Then zlSplitService_CallCancelAccCheck = True: Exit Function
    
    strJson_List = ""
    
    strPart1 = Split(strInput, "|")(0)
        
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("check_people", Split(strPart1, ",")(0), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("check_time", Split(strPart1, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("request_type", 1, 1)

    strJson_List = ""
    strPart2 = Split(strInput, "|")(1)
    arrPart = Split(strPart2, ";")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("rcpdtl_id", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("request_time", Split(arrPart(i), ",")(1), 0)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    strJson_List = ",""rcpdtl_list"":[" & strJson_List & "]"
       
    '汇总
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"
    
    strService = "Zl_Exsesvr_Cancelacc_Check"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_Exsesvr_Cancelacc_Check”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallCancelAccCheck = False
        Exit Function
    End If

    zlSplitService_CallCancelAccCheck = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function zlSplitService_CallAccountDel(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strAccDel As String) As Boolean
    '发药/退药调用服务：门诊/住院记录销账
    'strAccDel:操作员姓名,操作员编号,操作时间||费用来源;记录性质;费用单据号;序号串;操作状态|费用来源;记录性质;费用单据号;序号串;操作状态...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim strPart As String
    
    On Error GoTo errHandle
    
    '服务：Zl_Exsesvr_Delbill 门诊/住院记录通用服务
    
'  ---------------------------------------------------------------------------
'  --功能：门诊医嘱作废，住院医嘱回退发送，删除费用单据
'  --入参：Json_In:格式
'  --input
'  --   operator_name         C  1 操作员姓名【记帐单删除时传入】
'  --   operator_code         C  1 操作员编号【记帐单删除时传入】
'  --   operator_time         C  1 操作时间:yyyy-mm-dd hh:mi:ss【记帐单删除时传入】
'  --   del_list  直接删除的单据列表
'  --             fee_source          N 1 费用来源:1-门诊费用记录;2-住院费用记录
'  --             fee_bill_type       N 1 记录性质，1-收费单，2-记帐单
'  --             fee_no              C 1 费用单据号
'  --             serial_num          C 1 序号串，记帐单格式: 序号1:数量:执行状态1,序号2:数量2:执行状态2,...，收费单格式：1,2,3,4,5...
'  --             oper_status         N 1 操作状态，住院记帐单删时才传入，0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
'  --出参: Json_Out,格式如下
'  --output
'  --    code          N 1 应答吗：0-失败；1-成功
'  --    message       C 1 应答消息：失败时返回具体的错误信息
'  ---------------------------------------------------------------------------
    
    If strAccDel = "" Then zlSplitService_CallAccountDel = True: Exit Function
    
    strPart = Split(strAccDel, "||")(0)
    
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("operator_name", Split(strPart, ",")(0), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_code", Split(strPart, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_time", Split(strPart, ",")(2), 0)
    
    strPart = Split(strAccDel, "||")(1)
    strJson_List = ""
    If strPart <> "" Then
        arrInput = Split(strPart, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("fee_source", Split(arrInput(i), ";")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_bill_type", Split(arrInput(i), ";")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_no", Split(arrInput(i), ";")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_num", Split(arrInput(i), ";")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("oper_status", Split(arrInput(i), ";")(4), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""del_list"":[" & strJson_List & "]"
    End If
    
    '汇总
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"
    
    strService = "Zl_Exsesvr_Delbill"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_Exsesvr_Delbill”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallAccountDel = False
        Exit Function
    End If

    zlSplitService_CallAccountDel = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallAccountVerify(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strMain As String, ByVal strInput As String) As Boolean
    '功能：发药后服务：审核划价记账单
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

Public Function zlSplitService_CallAdviceIsInvalid(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, strOut As String, Optional ByRef colOutlist As Collection) As Boolean
    '组织医嘱作废数据：医嘱id,医嘱id...
    '返回值，已作废的医嘱ID串，医嘱id,医嘱id...
    Dim arrLongString As Variant
    Dim i As Integer, n As Integer
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String, strAdvices As String
       
    'zl_CisSvr_AdviceIsInvalid
'  ---------------------------------------------------------------------------
'  --功能：根据医嘱ID查询医嘱状态
'  --入参：Json_In:格式
'  --input
'  --   advice_ids           C  1  多个医嘱ID，用,分隔
'  --出参: Json_Out,格式如下
'  --output
'  --    code                 N  1  应答码：0-失败；1-成功
'  --    message              C  1  应答消息
'  --    advice_ids           C  1  医嘱ID（已作废的）
'  ---------------------------------------------------------------------------
    strOut = ""
    Set colOutlist = New Collection
    
    arrLongString = GetArrayByStr(strInput, 3900, ",")
    For i = 0 To UBound(arrLongString)
        '入参
        StrJson_In = ""
        StrJson_In = StrJson_In & "" & GetJsonNodeString("advice_ids", arrLongString(i), 0)
        StrJson_In = "{""input"":{" & StrJson_In & "}}"
        strService = "zl_CisSvr_AdviceIsInvalid"
    
        '调用服务
        If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
            MsgBox "调用“zl_CisSvr_AdviceIsInvalid”失败！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '返回出参
        strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
        
        '返回数据
        strAdvices = objServiceCall.GetJsonNodeValue("output.advice_ids")
        
        If strAdvices <> "" Then
            For n = 0 To UBound(Split(strAdvices, ","))
                colOutlist.Add strAdvices
            Next
            strOut = IIf(strOut = "", "", strOut & ",") & strAdvices
        End If
    Next

    zlSplitService_CallAdviceIsInvalid = True
End Function



Public Function zlSplitService_CallAuditContent(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strMain As String, ByVal strInput As String) As Boolean
    '静配医嘱审核
    'strMain：审核人
    'strInput：医嘱审核内容，格式化串： ID1,操作1,说明1||ID2,操作2,说明2…
    Dim strJson As String, StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    '服务：zl_CISSvr_AuditDrugOrder
    '  ---------------------------------------------------------------------------
    '  --入参      json
    '  --input     静配医嘱审核
    '  --  auditor        C  1  审核人
    '  --  audit_content  C  1  医嘱审核内容，格式化串：ID1,操作1,说明1||ID2,操作2,说明2…
    '  --出参      json
    '  --output
    '  --  code  C 1 应答码：0-失败；1-成功
    '  --  message C 1 应答消息：
    '  ---------------------------------------------------------------------------

    
    If strInput = "" Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("auditor", strMain, 0)
    strJson = strJson & "," & GetJsonNodeString("audit_content", strInput, 0)

    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "zl_CISSvr_AuditDrugOrder"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“zl_CISSvr_AuditDrugOrder”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallAuditContent = False
        Exit Function
    End If

    zlSplitService_CallAuditContent = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallCancelAccAudit(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strAccAudit As String, ByVal strAccReAudit As String, ByVal strAccDel As String, ByVal strExeSta As String, ByVal strCheck As String) As Boolean
    '退药同时销帐时调用的其他服务：合并销帐申请审核，重审销帐审核，门诊/住院记录删除，更新费用记录状态，销帐检查等功能
    '合并为一个服务可以有效控制发药/退药事务
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim arrList As Variant, arrPart As Variant
    Dim strPart1 As String, strPart2 As String, strPart3 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String
    Dim n As Integer
    Dim strJson_Part As String, strCheckJson_In As String
    
    On Error GoTo errHandle
    
    '1. 销帐申请审核
    'strAccAudit：处方明细id(费用id),申请时间,审核人,审核时间,销账审核状态|...
'    rcpdtl_list         [数组]每个处方明细信息
'        rcpdtl_id   N   1   处方明细id(费用id)
'        request_time    D   1   申请时间
'        auditor C   1   审核人
'        audit_time  D   1   审核时间
'        cancel_status   N   1   销账审核状态
'        auto_stuff_return   N       自动退料（默认传1）
'        request_type    N       申请类别（默认传1）

    
    strJson_List = ""
    If strAccAudit <> "" Then
        arrInput = Split(strAccAudit, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcpdtl_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("request_time", Split(arrInput(i), ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("auditor", Split(arrInput(i), ",")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("audit_time", Split(arrInput(i), ",")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("cancel_status", Split(arrInput(i), ",")(4), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""rcpdtl_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '2. 重审销帐申请：处方明细id(费用id),申请时间,审核人,审核时间,销账审核状态|...
'    retrial_rcpdtl_list            [数组]每个处方明细信息
'        rcpdtl_id   N   1   处方明细id(费用id)
'        request_time    D   1   申请时间
'        auditor         C   1   审核人
'        audit_time  D   1   审核时间
'        oper_type   N   1   操作类型:0-审核拒绝 1-取消拒绝
'        auto_stuff_return   N       自动退料（默认传1）
'        request_type    N       申请类别（默认传1）

    
    strJson_List = ""
    If strAccReAudit <> "" Then
        arrInput = Split(strAccReAudit, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcpdtl_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("request_time", Split(arrInput(i), ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("auditor", Split(arrInput(i), ",")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("audit_time", Split(arrInput(i), ",")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("oper_type", Split(arrInput(i), ",")(4), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""retrial_rcpdtl_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '3. 门诊/住院记录删除： no;费用序号(格式如"1,3,5,7,8",为空表示审核所有未审核的行);操作员编码;操作员姓名;记录性质;操作状态;输液检查;登记时间,费用来源|...
'  --  rcp_list         [数组]每个销账处方信息
'  --    rcp_no              N  1  处方no
'  --    serial_nums         C  1  格式为：序号1:数量1:执行状态1,序号2:数量2:执行状态2,...序号n:数量n:执行状态n  如:1:2:1,2:10:1,3:2:1
'  --    operator_code       C  1  操作员编码
'  --    operator_name       C  1  操作员姓名
'  --    fee_properties      N     记录性质
'  --    operator_status     N     操作状态：0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
'  --    pivas_flag          N     是否输液配药检查：0-医嘱调用，不检查药品是否进入输液配药中心；1-非医嘱调用，检查药品是否进入配药中心
'  --    create_time         D     登记时间
'  --    fee_origin          N     费用来源(1-门诊费用，2-住院费用)
    
    strJson_List = ""
    If strAccDel <> "" Then
        arrInput = Split(strAccDel, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), ";")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_nums", Split(arrInput(i), ";")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_code", Split(arrInput(i), ";")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_name", Split(arrInput(i), ";")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_properties", Split(arrInput(i), ";")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("operator_status", Split(arrInput(i), ";")(5), 1)
            strJson = strJson & "," & GetJsonNodeString("pivas_flag", Split(arrInput(i), ";")(6), 1)
            strJson = strJson & "," & GetJsonNodeString("create_time", Split(arrInput(i), ";")(7), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_origin", Split(arrInput(i), ";")(8), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""rcp_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '4. 更新费用执行状态：费用ID串;执行状态|费用ID串;执行状态...
'    bill_status_list            [数组]更新费用执行状态所需信息(执行状态相同的拼成一个费用id传)
'        detail_ids  C   1   处方明细id串(费用id串),支持多个id，用“,”分隔
'        exe_status  N   1   执行状态(0-未执行;1-完全执行;2-部分执行)
    
    strJson_List = ""
    If strExeSta <> "" Then
        If InStr(strExeSta, "||") > 0 Then
            strExeSta = Split(strExeSta, "||")(1)
        End If
        
        If strExeSta <> "" Then
            '组织入参,detail_ids节点值可能超长，需要在外部传入strExeSta值时先进行分解
            arrInput = Split(strExeSta, "|")
            
            For i = 0 To UBound(arrInput)
                strJson = ""
                strJson = strJson & "" & GetJsonNodeString("detail_ids", Split(arrInput(i), ";")(0), 0)
                strJson = strJson & "," & GetJsonNodeString("exe_status", Split(arrInput(i), ";")(1), 1)
                 
                If strJson_List = "" Then
                    strJson_List = "{" & strJson & "}"
                Else
                    strJson_List = strJson_List & "," & "{" & strJson & "}"
                End If
            Next
            
            strJson_List = ",""bill_status_list"":[" & strJson_List & "]"
        End If
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '5.销帐检查
'  --  feecheck_list    [数组]费用销帐检查所需要信息
'  --    fee_no                     C   1   费用单据号
'  --    balance_ban_writeoffs      N   1   已结禁止销帐:如果已结帐单据禁止销帐,或是医保记帐的单据。则在原始单据行中只取未结帐部分
'  --    part_ban_writeoffs         N   1   禁止部分销帐:1-不允许；0-允许
'  --    oper_type                  N 1 操作状态_In:0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
'  --    fee_origin                 N 1            费用来源（1-门诊记帐，2-住院记帐）
'  --    item_list[]                 本次销帐列表
'  --        serial_num              N   1   序号
'  --        quantity                N   1   销帐数量(为零时，按序号直接销帐)
'  --    excute_list[]               药品及卫材所对应已执行列表
'  --        fee_id                  N   1   费用ID
'  --        sended_num              N   1   已发数量

    'strCheck 格式：no,已结禁止销帐(1),医保禁止部分销帐(0),操作状态(1),费用来源|序号,销帐数量;序号,销帐数量...|费用id,已发数量;费用id,已发数量...||...
    arrList = Split(strCheck, "||")
    
    For n = 0 To UBound(arrList)
        arrPart = Split(arrList(n), "|")
        
        strPart1 = arrPart(0)
        strPart2 = arrPart(1)
        strPart3 = arrPart(2)
            
        strJson_In_Part1 = ""
        strJson_In_Part1 = strJson_In_Part1 & "" & GetJsonNodeString("fee_no", Split(strPart1, ",")(0), 0)
        strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("balance_ban_writeoffs", Split(strPart1, ",")(1), 1)
        strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("part_ban_writeoffs", Split(strPart1, ",")(2), 1)
        strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("oper_type", Split(strPart1, ",")(3), 1)
        strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("fee_origin", Split(strPart1, ",")(4), 1)
    
        strJson_List = ""
        arrPart = Split(strPart2, ";")
        For i = 0 To UBound(arrPart)
            strJson_Part = ""
            strJson_Part = strJson_Part & "" & GetJsonNodeString("serial_num", Split(arrPart(i), ",")(0), 1)
            strJson_Part = strJson_Part & "," & GetJsonNodeString("quantity", zlStr.FormatEx(Split(arrPart(i), ",")(1), 5, False), 1)
            
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
        
        If strCheckJson_In = "" Then
            strCheckJson_In = "{" & strJson_In_Part1 & strJson_In_Part2 & strJson_In_Part3 & "}"
        Else
            strCheckJson_In = strCheckJson_In & "," & "{" & strJson_In_Part1 & strJson_In_Part2 & strJson_In_Part3 & "}"
        End If
    Next
    
    strCheckJson_In = ",""feecheck_list"":[" & strCheckJson_In & "]"
    StrJson_In = StrJson_In & strCheckJson_In
     
    '汇总
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "zl_ExseSvr_BillChargeOffExtend"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“zl_ExseSvr_BillChargeOffExtend”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallCancelAccAudit = False
        Exit Function
    End If
    
    zlSplitService_CallCancelAccAudit = True
    Exit Function
errHandle:
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
    
    On Error GoTo errHandle
    
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
errHandle:
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
    
    On Error GoTo errHandle
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
        If colOutlist.count > 0 Then
            '控制方式,提示信息
            strCheckMsg = colOutlist(1)(1) & "," & colOutlist(1)(2)
        End If
    End If
    
    zlSplitService_WriteOffCheck = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetCloseAccount(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal colInput As Collection, _
    ByRef colCloseAccount As Collection) As Boolean
    '取销帐记录
    'colInput：查询条件组合，Json中input各节点作为元素的KEY值，集合某元素为空表示该节点值为空
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    'Zl_Exsesvr_Getchargeoffinfo
    '  --根据条件查询费用销帐信息
    '  --入参      json
    '  --  input      根据条件查询费用销帐信息
    '  --    audit_dept_id       N    审核部门ID(药房)
    '  --    request_begin_time  D    申请开始时间
    '  --    request_end_time    D    申请结束时间
    '  --    audit_begin_time    D    审核开始时间
    '  --    audit_end_time      D    审核结束时间
    '  --    cancel_status       N  1 状态
    '  --    request_dept_id     N    申请部门ID
    '  --    request_operator    C    申请人
    '  --    pati_id             N    病人ID
    '  --    cancel_condition    C    销账条件
    '  --    cancel_check        N    核查（选择参数【销账申请需要核查】时传入，0-未核查 1-已核查）
    '  --    rcpdtl_id          C     处方明细id,[数组]：[1,2,3]
    '  --    request_dept_ids   C     申请部门id串，用于批量查询
    '  --    item_ids           C     收费细目id串,用于批量查询
    '  --    request_type       N     申请类别：-1-不区分;0-未执行;1-已执行
    '  --出参      json
    '  -- output
    '  --   code     C  1   应答码：0-失败；1-成功
    '  --   message  C  1   应答消息：
    '  --   fee_cancel_list      [数组]满足条件的每个费用销帐记录
    '  --     rcpdtl_id          N    处方明细id(费用id)
    '  --     request_type       N    申请类别
    '  --     item_id            N    收费细目id
    '  --     request_dept_id    N    申请部门id
    '  --     request_dept       C    申请部门
    '  --     audit_dept_id      N    审核部门id
    '  --     quantity           N    数量
    '  --     request_operator   C    申请人
    '  --     request_time       D    申请时间
    '  --     auditor            C    审核人
    '  --     audit_time         D    审核时间
    '  --     cancel_status      N    状态
    '  --     cancel_reason      C    销帐原因
    '  --     checker            C    核查人
    '  --     price_retail       N    零售价
    '  --     advice_id          N    医嘱id
    '  --     pati_id            N    病人ID
    '  --     pati_name          C    病人姓名
    '  --     inpatient_num      C    住院号
    '  --     pati_pageid        N    主页id
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
    If ExistsColObject(colInput, "request_type") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_type", colInput("request_type"), 1)
    
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
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function zlSplitService_GetFee(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String, Optional ByVal strInputByNO As String) As Boolean
    '取费用信息
    'strInput：费用id,费用id...
    'strInputByNO：no,记录性质|...
    Dim StrJson_In As String
    Dim strService As String
    
    On Error GoTo errHandle
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
    '  --  fee_status          N    费用状态
    '  -------------------------------------------------------------------------------------------------
    
    If strInput = "" And strInputByNO = "" Then Exit Function
    
    '入参
    StrJson_In = ""
    If strInput <> "" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_ids", strInput, 0)
    ElseIf strInputByNO <> "" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("bill_nos", strInputByNO, 0)
    End If
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_GetBillDetailInfo"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, , "", lngMode) = False Then Exit Function
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.fee_list", strKey)
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetFee = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiDiagnose(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '取诊断信息
    'strInput：病人id,主页id...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    If strInput = "" Then Exit Function
    
    '入参
    StrJson_In = ""
    If strInput <> "" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ",")(0), 1)
        StrJson_In = StrJson_In & "," & GetJsonNodeString("visit_id", Split(strInput, ",")(1), 1)
    End If
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_CisSvr_GetPatiDiagnose"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“zl_CisSvr_GetPatiDiagnose”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.diagnose_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetPatiDiagnose = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetDiagInfo(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '取诊断信息：门诊
    'strInput：医嘱id,医嘱id...
    '         这里是主医嘱id
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    'Zl_Cissvr_Getdiaginfo
'  ---------------------------------------------------------------------------
'  --功能:获取病人诊断信息
'  --入参：Json_In:格式
'  -- input
'  --   advice_ids           C 1  医嘱ids,医嘱id拼串
'  --   query_type           N 1 查询方式1-按指定条件查询,2-仅按病人id,主页id查询诊断
'  --   pati_info            C 0  病人id等其他信息
'  --     pati_id            N 1 病人id
'  --     pati_pageid        N 1 主页id
'  --     diag_types         C 1  诊断类型:0-所有类型,1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
'  --                            可以为多个诊断类型，用逗号分离,如:2,12
'  --     rec_source         N 1 记录来源:1-病历；2-入院登记；3-首页整理;4-病案;NULL-不作限制
'  --     diag_num           N 1 诊断次序:NULL表示不作限制
'  --     code_type          C 1  编码类别:ICD-11的编码编码类别为'E',空时表示读取ICD-10等
'  --     input_num          C 1  录入次序:启用了ICD-11编码录入后，诊断的录入次序
'  --     rec_sources        C 1 记录来源拼串
'
'  --出参: Json_Out,格式如下
'  --  output
'  --    code                N  1 应答码：0-失败；1-成功
'  --    message             C  1 应答消息：失败时返回具体的错误信息
'  --    diag_list     [数组]
'  --      diag_type         N 1 诊断类型:1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;
'  --                                     12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
'  --      diag_num          N 1 诊断序号
'  --      code_num          N 1 编码序号
'  --      dz_id             N 1 疾病ID
'  --      dz_code           C 1 疾病编码
'  --      diag_note         C 1 诊断描述
'  --      recoder           C 1 记录人
'  --      rec_time          C 1 记录时间:yyyy-mm-dd hh24:mi:ss
'  --      adtd_rsn          C 1 出院情况:治愈、好转、未愈、死亡、其他
'  --      diag_id           N 1 诊断id
'  --      diag_rec_id       N 1 诊断记录ID:病人诊断记录.ID
'  --      diag_doubt        N 1 是否疑诊
'  --      advice_id         N   医嘱id(根据医嘱ids查询时才返回)
'  --      advice_main_id    N   组医嘱id(根据医嘱ids查询时才返回)
'  --      advice_related_id N   相关id(根据医嘱ids查询时才返回)
'  --      rec_source        N   记录来源:1-病历；2-入院登记；3-首页整理;4-病案;NULL-不作限制
'  ---------------------------------------------------------------------------
  
    On Error GoTo errHandle
    
    If strInput = "" Then Exit Function
    
    '入参
    StrJson_In = ""
    If strInput <> "" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("advice_ids", strInput, 0)
    End If
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getdiaginfo"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Cissvr_Getdiaginfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.diag_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetDiagInfo = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_GetAccWarnLine(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOut As String) As Boolean
    '获取记账报警线
    'strInput：科室或病区id,病人类型
    'strOut: 报警方法,报警值,报警标志1,报警标志2,报警标志3
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    Dim colOutlist As New Collection
    
    'Zl_Exsesvr_Getwarnline
'  ---------------------------------------------------------------------------
'  --功能：获取记帐报警线，过滤病人用于排开欠费病人
'  --入参：Json_In:格式
'  --  input
'  --     pati_scheme  C 1 适用病人
'  --     wardarea_id  N 1 病区id
'  --     query_type   N 1 查询方式
'  --                     0-仅根据 病区id / 适用病人 查找，返回一个值
'  --                     1-按病区id 查找，返回列表
'  --                     2-获取所有报警线设置
'  --                     3-根据病区id，适用病人查找，返回报警方法，报警值，报警标志
'  --出参: Json_Out,格式如下
'  --  output
'  --    code                N 1 应答吗：0-失败；1-成功
'  --    message             C 1 应答消息：失败时返回具体的错误信息
'  --      alarm_value       N 1 报警值
'  --      item_list[]
'  --        pati_scheme     C 1 适用病人
'  --        alarm_way       N 1 报警方法
'  --        alarm_value     N 1 报警值
'  --        alarm_one       C 1 报警标志1
'  --        alarm_two       C 1 报警标志2
'  --        alarm_three     C 1 报警标志3
'  --        wardarea_id     N 1 病区id
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_scheme", Split(strInput, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("wardarea_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", 3, 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_Getwarnline"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Exsesvr_Getwarnline”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    Set colOutlist = objServiceCall.GetJsonListValue("output.item_list")
    
    If Not colOutlist Is Nothing Then
        If colOutlist.count > 0 Then
            strOut = colOutlist(1)("_pati_scheme")
            strOut = strOut & "," & colOutlist(1)("_alarm_way")
            strOut = strOut & "," & colOutlist(1)("_alarm_value")
            strOut = strOut & "," & colOutlist(1)("_alarm_one")
            strOut = strOut & "," & colOutlist(1)("_alarm_two")
            strOut = strOut & "," & colOutlist(1)("_alarm_three")
        End If
    End If
        
    zlSplitService_GetAccWarnLine = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetTodayMoney(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByVal bln含划价 As Boolean, ByRef strOut As String) As Boolean
    '获取记账报警线
    'strInput：病人id
    'strOut: 总额
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
     'Zl_Exsesvr_Getpatitotalmoney
'  ---------------------------------------------------------------------------
'  --功能:根据病人ID,主页ID或医嘱id，获取应收、实收总额
'  --入参：Json_In:格式
'  --  input
'  --    pati_source N 1 病人来源:0-所有;1-门诊;2-住院
'  --    pati_id N 1 病人ID
'  --    visit_id  N   就诊ID:住院时，传入主页id,门诊暂传NULL
'  --    advice_ids  C   医嘱ids:多个用逗号分离
'  --    today_fee N   是否当日费用:1-是的;0-不限制
'  --    price_tag N   划价标志:0-不限制;1-不含划价单;2-仅统计划价单
'  --出参: Json_Out,格式如下
'  --  output
'  --    code        C  1  应答码：0-失败；1-成功
'  --    message     C  1  "应答消息：
'  --    fee_amrcvb  N  1  应收金额
'  --    fee_ampaib  N  1  实收金额
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_source", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_id", strInput, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("today_fee", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("price_tag", IIf(bln含划价, 0, 1), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_Getpatitotalmoney"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Exsesvr_Getpatitotalmoney”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")

    strOut = objServiceCall.GetJsonNodeValue("output.fee_ampaib")
        
    zlSplitService_GetTodayMoney = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetFeeNO(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByVal lng药房id As Long, ByRef colOutlist As Collection) As Boolean
    '取费用NO整体信息
    'strInput：记录性质,NO|...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    'Zl_Exsesvr_Getbillinfo
    '  --功能：获取和药品发药业务相关的费用信息，主要用于界面显示
    '  --入参：json格式
    '  --Input
    '  --   pharmacy_id：库房id
    '  --   fee_nos：费用no，支持多个no，格式： 记录性质,no,…
    '  --出参：json格式
    '  --Json_Out
    '  --  code          C   1   应答码：0-失败；1-成功
    '  --  message       C   1   应答消息：
    '  --  fee_list      C       [数组]每个费用NO信息
    '  --    fee_properties      N 记录性质
    '  --    bill_no             C 费用no
    '  --    real_amount         N 实收金额
    '  --    rcp_type            N 收费类别(按整个NO来说，1-西药，2-中药，3-混合)
    '  --    iden_id             C 标识号
    '  --    placer              C 开单人
    '  --    bill_deptid         N 开单部门id
    '  --    create_time         D 登记时间
    '  --    pati_bed            C 当前床号
    '  --    operator_name       C 操作员姓名
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pharmacy_id", lng药房id, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("fee_nos", strInput, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
        
    strService = "zl_ExseSvr_GetBillInfo"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“zl_ExseSvr_GetBillInfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.fee_list")
    If colOutlist Is Nothing Then Exit Function

    zlSplitService_GetFeeNO = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetNOByInvoice(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
     ByRef strOutPut As String) As Boolean
    '通过票据号取费用NO
    'strInput：票据号
    'strOutput：NO1,NO2...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
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
    
    On Error GoTo errHandle
    
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
    strOutPut = objServiceCall.GetJsonNodeValue("output.rcp_nos")

    zlSplitService_GetNOByInvoice = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetNextNO(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef strNos_Out As String, Optional ByVal lngCount As Long = 1) As Boolean
    '获取费用业务NO
    '入参：
    '   strInput：序号|科室ID
    '   lngCount = 获取NO个数
    '出参：
    '   strNos_Out = NO,多个逗号分隔
    Dim StrJson_In As String, strJson_Out As String
    
    On Error GoTo errHandle
    'Zl_Exsesvr_Getnextno
    '  --功能：功能：根据特定规则产生新的号码
    '  --入参：Json_In:格式
    '  --  input
    '  --    item_num            N   1   项目序号
    '  --    dept_id             N   0   科室ID
    '  --    quantity            N   0   所需no号的个数，如果只取一个该参不传或都传0
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N   1   应答码：0-失败；1-成功
    '  --    message             C   1   应答消息：失败时返回具体的错误信息
    '  --    next_no             C   1   下一个号码,quantity>1 时，表示取多个单据号,用逗号分离
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("item_num", Val(Split(strInput, ",")(0)), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("dept_id", Val(Split(strInput, ",")(1)), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("quantity", lngCount, 1)
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    If objServiceCall.CallService("zl_ExseSvr_GetNextNo", StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“zl_ExseSvr_GetNextNo”获取费用单据号失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    strNos_Out = objServiceCall.GetJsonNodeValue("output.next_no")
    If strNos_Out = "" Then
        MsgBox "调用“zl_ExseSvr_GetNextNo”获取费用单据号失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    zlSplitService_GetNextNO = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetAllergy(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
   ByRef colOutlist As Collection) As Boolean
    '取病人药物过敏记录
    'strInput：病人id,标识id
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    'Zl_Cissvr_Getpatiallergyinfo
'    -------------------------------------------------------------------------------------------------
'    --功能：获取病人过敏信息
'--input      获取病人过敏信息
'--  pati_id  N  1  病人id
'--  visit_id  N    标识号：挂号id（门诊），主页id（住院）
'--output
'--  code  C  1  应答码：0-失败；1-成功
'--  message  C  1  "应答消息：
'--  allergy_list  C    过敏信息，[数组]
'--     drug_name  C    药物名称
'--     allergy_time  D    过敏时间
'--     allergy_info  C    过敏反应
'    -------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("visit_id", Split(strInput, ",")(1), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
        
    strService = "Zl_Cissvr_Getpatiallergyinfo"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Cissvr_Getpatiallergyinfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.allergy_list")
    
    If colOutlist Is Nothing Then Exit Function

    zlSplitService_GetAllergy = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPativitalsigns(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
   ByRef colOutlist As Collection) As Boolean
    '取病人生命体征信息
    'strInput：病人id,标识id,门诊标志
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
'  --获取病人生命体征信息
'  ---------------------------------------------------------------------------
'  --input      获取病人生命体征信息
'  --  pati_id  N  1  病人ID
'  --  visit_id  N  1  就诊id ，门诊病人为挂号ID;住院病人为主页ID;
'  --  outpati_flag  N    门诊标志：1-门诊，2-住院
'  --output
'  --  code  C  1  应答码：0-失败；1-成功
'  --  message  C  1  "应答消息：
'  --  pativital_list      体征信息，包括项目，数值，单位。[数组]
'  --     pativital_item  C    项目
'  --     pativital_value  C    值
'  --     pativital_unit  C    单位
'  ---------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("visit_id", Split(strInput, ",")(1), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("outpati_flag", Split(strInput, ",")(2), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
        
    strService = "Zl_Cissvr_Getpativitalsigns"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“Zl_Cissvr_Getpativitalsigns”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.pativital_list")
    
    If colOutlist Is Nothing Then Exit Function

    zlSplitService_GetPativitalsigns = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_GetAdvice(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String, Optional ByVal bytQueryType As Byte = 0) As Boolean
    '取医嘱信息
    'strInput：医嘱id,医嘱id...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
            
'      --功能:获取医嘱信息
'  --入参：Json_In:格式
'  -- input
'  --   query_type                   N 0 查询类型：0:查询基本信息；1:查询基本信息+扩展信息
'  --   advice_ids                   C 0 多个医嘱ID，可能是药嘱，也可能是主医嘱（给药途径）,用“,”分隔
'  --   rgst_no                      C 0 挂号单号:挂号单或病人ID或医嘱ID必传其中一个条件
'  --   pati_id                      N 0 病人ID:挂号单或病人ID或医嘱ID必传其中一个条件
'  --   pati_pageid                  N 0 主页Id
  
    On Error GoTo errHandle
    
    If strInput = "" Then zlSplitService_GetAdvice = True: Exit Function
  
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", bytQueryType, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("advice_ids", strInput, 0)
'    strJson_In = strJson_In & "," & GetJsonNodeString("advice_ids", strPatis, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_CisSvr_GetAdviceInfo"
    
    '测试
    'StrJson_In = "{""input"":{""advice_id"":""521""}}"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“zl_CisSvr_GetAdviceInfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.advice_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetAdvice = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetAdviceSend(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '组织明细数据集：医嘱发送信息
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    '入参
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("advice_ids", strInput, 0)
'    strJson_In = strJson_In & "," & GetJsonNodeString("pati_list", strPatis, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_CisSvr_GetAdviceSendInfo"

    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "调用“zl_CisSvr_GetAdviceSendInfo”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '返回数组
    Set colOutlist = objServiceCall.GetJsonListValue("output.advice_send_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function

    zlSplitService_GetAdviceSend = True
End Function

Public Function zlSplitService_CheckAdviceaffirm(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOutPut As String, ByRef strErrMsg_Out As String) As Boolean
    '返回指定病人ID，主页ID，挂号ID的医嘱发送信息
    'strInput：病人id,主页ID,挂号ID,挂号单号|...
    'strOutPut：医嘱id,发送号|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colOutlist As Collection, colOrderList As Collection
    Dim arrOrder As Variant
    
    On Error GoTo ErrHandler
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
     
    If strInput = "" Then zlSplitService_CheckAdviceaffirm = True: Exit Function
    
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
    StrJson_In = """pati_list"":[" & strJson_List & "]"
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_CISSvr_GetAffirmErrorData"
    
    If objServiceCall.CallService(strService, StrJson_In, , "", lngMode, False, , , , True) = False Then Exit Function
    
    Set colOutlist = objServiceCall.GetJsonListValue("output.pati_bill_list")
    '只需要返回内容：医嘱id,发送号|...
    For i = 0 To colOutlist.count - 1
        '循环取子节点数组
        Set colOrderList = objServiceCall.GetJsonListValue("output.pati_bill_list[" & i & "].order_list")
        
        For Each arrOrder In colOrderList
            strOutPut = IIf(strOutPut = "", "", strOutPut & "|") & arrOrder("_advice_id") & "," & arrOrder("_send_no")
        Next
    Next

    zlSplitService_CheckAdviceaffirm = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

Public Function zlSplitService_CheckErrData(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strType As String, ByVal strInputByid As String, ByVal strInputByNO As String, _
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
    If strInputByid = "" And strInputByNO = "" Then zlSplitService_CheckErrData = True: Exit Function
    
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_type", strType, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("rcpdtl_ids", strInputByid, 0)
    If strInputByNO <> "" Then
        arrInput = Split(strInputByNO, "|")
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
    If strInputByid <> "" Then
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

Public Function zlSplitService_CallUpdateSynchrosign(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal intType As Integer, ByVal strInputByid As String, ByRef strErrMsg_Out As String) As Boolean
    '更新费用记费同步标记
    '按处方NO进行更新
    'intType：0-更新记费同步标志，1-更新转费同步标志
    'strInputByid：处方明细ID，多个英文逗号分隔
    Dim StrJson_In As String
    
    On Error GoTo ErrHandler
    strErrMsg_Out = ""
    If strInputByid = "" Then zlSplitService_CallUpdateSynchrosign = True: Exit Function
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
    StrJson_In = StrJson_In & "," & GetJsonNodeString("detail_ids", strInputByid, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Sync_Update", StrJson_In, , "", lngMode, False, , , , True) = False Then Exit Function
    
    zlSplitService_CallUpdateSynchrosign = True
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
    
    On Error GoTo ErrHandler
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
ErrHandler:
    strErrMsg_Out = err.Description
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
    
    On Error GoTo errHandle
    
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
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallExseData(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strExeDept As String, ByVal strSendWin As String, ByVal strExePeople As String, _
    ByVal strExeSta As String, Optional ByVal strAccDel As String) As Boolean
    '发药/退药调用服务：更新费用字段（执行部门，发药窗口，执行人, 执行状态）根据传参更新对应内容
    Dim arrInput As Variant, strPart As String
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim str执行人 As String, str执行时间 As String
    
    On Error GoTo errHandle
    
    'Zl_ExseSvr_BillInforUpdate 通用服务
    '1. 更新执行部门
    '2. 更新发药窗口
    '3. 更新执行人
    '4. 更新执行状态
    
    '1. 更新执行部门
    '更新执行部门：当前药房id,单据性质,no,原药房id,门诊标志,填制日期|...
    'strExeDept：306,1,T0000312,310,2,2018-03-28 13:05:22|...
    'bill_dept_list          [数组]更新费用执行部门所需信息
    '    pharmacy_id N   1   当前药房id
    '    bill_type   C   1   单据性质(1-收费处方,2-记账处方)
    '    rcp_no  C   1   NO
    '    pharmacy_id_old C   1   原药房id
    '    outpati_flag    N       门诊标志:1-门诊，2-住院
    '    write_time  N       填制日期
    
    strJson_List = ""
    If strExeDept <> "" Then
        arrInput = Split(strExeDept, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("pharmacy_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("bill_type", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("rcp_no", Split(arrInput(i), ",")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("pharmacy_id_old", Split(arrInput(i), ",")(3), 1)
            strJson = strJson & "," & GetJsonNodeString("outpati_flag", Split(arrInput(i), ",")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("write_time", Split(arrInput(i), ",")(5), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_dept_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '2. 更新发药窗口：no,单据类型,执行部门id,发药窗口|...
    'bill_win_list           [数组]更新费用发药窗口所需信息
    '    rcp_no  C   1   NO
    '    fee_properties  N   1   记录性质
    '    fee_exe_deptid  N   1   执行部门id
    '    pharmacy_window C   1   发药窗口
    
    strJson_List = ""
    If strSendWin <> "" Then
        arrInput = Split(strSendWin, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), ",")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_properties", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_exe_deptid", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("pharmacy_window", Split(arrInput(i), ",")(3), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_win_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '3. 更新费用执行人：no,单据类型,执行部门id,执行人|...
    'bill_people_list            [数组]更新费用执行人所需信息
    '    rcp_no  C   1   NO
    '    bill_type   N   1   单据类型(1-收费，2-记账）
    '    exe_deptid  N   1   执行部门id
    '    exe_people  C   1   执行人
    
    strJson_List = ""
    If strExePeople <> "" Then
        arrInput = Split(strExePeople, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), ",")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("bill_type", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_deptid", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_people", Split(arrInput(i), ",")(3), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_people_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '4. 更新费用执行状态：执行人;发药时间||费用ID串;执行状态|费用ID串;执行状态
    'strExeSta：张三;2018-09-10 10:43:22||6341798,6341800,634180;1|6341818;2
    '    bill_status_list            [数组]更新费用执行状态所需信息(执行状态相同的拼成一个费用id传)
    '        detail_ids  C   1   处方明细id串(费用id串),支持多个id，用“,”分隔
    '        exe_status  N   1   执行状态(0-未执行;1-完全执行;2-部分执行)
    '        exe_people  C   1   执行人(退药的情况传空：全退执行人会改为null,部分退不修改)
    '        give_time   C   1   发药时间(退药的情况传空：与执行人同理)
    
    strJson_List = ""
    If strExeSta <> "" Then
        str执行人 = Split(Split(strExeSta, "||")(0), ";")(0)
        str执行时间 = Split(Split(strExeSta, "||")(0), ";")(1)
        strExeSta = Split(strExeSta, "||")(1)
        
        If strExeSta <> "" Then
            '组织入参,detail_ids节点值可能超长，需要在外部传入strExeSta值时先进行分解
            arrInput = Split(strExeSta, "|")
            
            For i = 0 To UBound(arrInput)
                strJson = ""
                strJson = strJson & "" & GetJsonNodeString("detail_ids", Split(arrInput(i), ";")(0), 0)
                strJson = strJson & "," & GetJsonNodeString("exe_status", Split(arrInput(i), ";")(1), 1)
                strJson = strJson & "," & GetJsonNodeString("exe_people", str执行人, 0)
                strJson = strJson & "," & GetJsonNodeString("give_time", str执行时间, 0)
                
                If strJson_List = "" Then
                    strJson_List = "{" & strJson & "}"
                Else
                    strJson_List = strJson_List & "," & "{" & strJson & "}"
                End If
            Next
            
            strJson_List = ",""bill_status_list"":[" & strJson_List & "]"
        End If
    End If
    StrJson_In = StrJson_In & strJson_List

    '5. 门诊/住院记录删除： 操作员姓名,操作员编号,操作时间||费用来源;记录性质;费用单据号;序号串;操作状态|费用来源;记录性质;费用单据号;序号串;操作状态...
    '  --  bill_rcp_list     [数组]需删除划价单的单据信息
    '  --    rcp_no          N   1     处方no
    '  --    serial_nums     C   1     格式如"1,3,5,7,8",或"1:2:33456,3:2,5:2,7:2,8:2",冒号前面的数字表示行号,中间的数字表示退的数量,后面的数字表示配药记录的ID,目前仅在销帐审核时才传入
    '  --    operator_code   C   1     操作员编码
    '  --    operator_name   C   1     操作员姓名
    '  --    fee_properties  N         记录性质
    '  --    operator_status N         操作状态：0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
    '  --    create_time     D         登记时间
    '  --    fee_origin      N         费用来源(默认=2：1-门诊费用，2-住院费用)
    
    strJson_List = ""
    If strAccDel <> "" Then
        strPart = Split(strAccDel, "||")(0)
        arrInput = Split(Split(strAccDel, "||")(1), "|")
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("fee_origin", Split(arrInput(i), ";")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_properties", Split(arrInput(i), ";")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("rcp_no", Split(arrInput(i), ";")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_nums", Split(arrInput(i), ";")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_status", Split(arrInput(i), ";")(4), 1)
            
            strJson = strJson & "," & GetJsonNodeString("operator_code", Split(strPart, ",")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_name", Split(strPart, ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("create_time", Split(strPart, ",")(2), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_rcp_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    If StrJson_In = "" Then zlSplitService_CallExseData = True: Exit Function
    
    '汇总
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_ExseSvr_BillInforUpdate"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_ExseSvr_BillInforUpdate”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallExseData = False
        Exit Function
    End If

    zlSplitService_CallExseData = True

    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_CallRetunDrug(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strAccDel As String, ByVal strExeSta As String) As Boolean
    '退药同时清除划价单时调用的其他服务：门诊/住院记录删除，更新费用记录状态等功能
    '合并为一个服务可以有效控制发药/退药事务
    'strAccDel：rcp_no|serial_nums|operator_code|operator_name|fee_properties|operator_status|create_time|fee_origin||...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    'Zl_ExseSvr_BillInforUpdate 通用服务
    
    '1. 门诊/住院记录删除： no;费用序号(格式如"1,3,5,7,8",为空表示审核所有未审核的行);操作员编码;操作员姓名;记录性质;操作状态;登记时间;费用来源|...
'  --  bill_rcp_list     [数组]需删除划价单的单据信息
'  --    rcp_no          N   1     处方no
'  --    serial_nums     C   1     格式如"1,3,5,7,8",或"1:2:33456,3:2,5:2,7:2,8:2",冒号前面的数字表示行号,中间的数字表示退的数量,后面的数字表示配药记录的ID,目前仅在销帐审核时才传入
'  --    operator_code   C   1     操作员编码
'  --    operator_name   C   1     操作员姓名
'  --    fee_properties  N         记录性质
'  --    operator_status N         操作状态：0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程);2-表示转病区费用
'  --    create_time     D         登记时间
'  --    fee_origin      N         费用来源(默认=2：1-门诊费用，2-住院费用)
    
    strJson_List = ""
    If strAccDel <> "" Then
        arrInput = Split(strAccDel, "||")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), "|")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_nums", Split(arrInput(i), "|")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_code", Split(arrInput(i), "|")(2), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_name", Split(arrInput(i), "|")(3), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_properties", Split(arrInput(i), "|")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("operator_status", Split(arrInput(i), "|")(5), 1)
            strJson = strJson & "," & GetJsonNodeString("create_time", Split(arrInput(i), "|")(6), 0)
            strJson = strJson & "," & GetJsonNodeString("fee_origin", Split(arrInput(i), "|")(7), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_rcp_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '2. 更新费用执行状态：费用ID串;执行状态|费用ID串;执行状态...
'  --  bill_status_list   [数组]更新费用执行状态所需信息(执行状态相同的拼成一个费用id传)
'  --    detail_ids      C   1     处方明细id串(费用id串),支持多个id，用“,”分隔
'  --    exe_status      N   1     执行状态
'  --    exe_people      N   1     执行人
'  --    give_time       C   1     发药时间
    
    strJson_List = ""
    If strExeSta <> "" Then
        strExeSta = Split(strExeSta, "||")(1)
        
        If strExeSta <> "" Then
            '组织入参,detail_ids节点值可能超长，需要在外部传入strExeSta值时先进行分解
            arrInput = Split(strExeSta, "|")
            
            For i = 0 To UBound(arrInput)
                strJson = ""
                strJson = strJson & "" & GetJsonNodeString("detail_ids", Split(arrInput(i), ";")(0), 0)
                strJson = strJson & "," & GetJsonNodeString("exe_status", Split(arrInput(i), ";")(1), 1)
                 
                If strJson_List = "" Then
                    strJson_List = "{" & strJson & "}"
                Else
                    strJson_List = strJson_List & "," & "{" & strJson & "}"
                End If
            Next
            
            strJson_List = ",""bill_status_list"":[" & strJson_List & "]"
        End If
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '汇总
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_ExseSvr_BillInforUpdate"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_ExseSvr_BillInforUpdate”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallRetunDrug = False
        Exit Function
    End If
    
    zlSplitService_CallRetunDrug = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
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


Public Function GetJsonNodeToNull(ByVal strNodeName As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将节点值设为null，只针对数字型
    '入参:strNodeName-接点名
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    
    strJson = Chr(34) & strNodeName & Chr(34)
    strJson = strJson & ":null"
    
    GetJsonNodeToNull = strJson
End Function

Public Function zlSplitService_CallIsCloseAcc(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef inState As Integer) As Boolean
    '发药/退药调用服务：查询是否已结帐
    'strInput：费用来源|No
    'inState：返回结帐状态
    Dim strJson As String, StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
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
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CheckExeItemValied(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef strOutPut As String) As Boolean
    '先诊疗后结算方式，检查执行项目的合法性
    'strInput：病人id|挂号id|收费类别
    'strOutPut：检查标志|消息内容
    Dim strJson As String, StrJson_In As String, strJson_Out As String
    Dim strService As String
    
    On Error GoTo errHandle
    
    '服务：Zl_Exsesvr_CheckExeItemValied
'  -------------------------------------------------------------------------------------------------
'  --功能：先诊疗后结算方式，检查执行项目的合法性
'  --input
'  --  pati_id      N   1   病人id
'  --  register_id   N   1   挂号id
'  --  receipt_type  C   1   收费类别
'  --output
'  --  code          C   1   应答码：0-失败；1-成功
'  --  message       C   1   应答消息：
'  --  check_flag   N   0   检查标志：0-不检查或合法，1-提醒 ，2-拒绝
'  --  check_msg    C   0   提醒或拒绝的内容提示
'  -------------------------------------------------------------------------------------------------
    
    If strInput = "" Then Exit Function

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", Split(strInput, "|")(0), 1)
    strJson = strJson & "," & GetJsonNodeString("register_id", Split(strInput, "|")(1), 1)
    strJson = strJson & "," & GetJsonNodeString("receipt_type", Split(strInput, "|")(2), 0)

    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "Zl_Exsesvr_CheckExeItemValied"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“zlSplitService_CheckExeItemValied”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CheckExeItemValied = False
        Exit Function
    Else
        strOutPut = objServiceCall.GetJsonNodeValue("output.check_flag") & "|" & objServiceCall.GetJsonNodeValue("output.check_msg")
    End If

    zlSplitService_CheckExeItemValied = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_CallSetWindows(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String) As Boolean
    '发药窗口调整调用服务：设置费用药品单据的发药窗口
    'strInput：库房id,旧窗口,新窗口|单据性质,no;单据性质,no...
    'dblSumAmount：返回结帐金额
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim strWins As String, strNOs As String
    
    On Error GoTo errHandle
    
    '服务：Zl_Exsesvr_Setsendwin
'  --功能：设置费用药品单据的发药窗口
'  --入参：Json_In:格式
'  --  input
'  --    pharmacy_id              N   1  库房id
'  --    pharmacy_window_old      C   1  旧发药窗口
'  --    pharmacy_window_new      C   1  新发药窗口
'  --    bill_list[]
'  --      billtype               N   1 单据类型:1-收费处方;2-记帐处方
'  --      rcp_no                 C   1 处方No
'  --出参: Json_Out,格式如下
'  --  output
'  --     code                   N   1 应答吗：0-失败；1-成功
'  --     message                C   1 应答消息：失败时返回具体的错误信息
    
    If strInput = "" Then Exit Function
    
    '窗口列表
    strWins = Split(strInput, "|")(0)
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pharmacy_id", Split(strWins, ",")(0), 1)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_window_old", Split(strWins, ",")(1), 0)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_window_new", Split(strWins, ",")(2), 0)
    
    '单据列表
    strNOs = Split(strInput, "|")(1)
    arrInput = Split(strNOs, ";")
    
    For i = 0 To UBound(arrInput)
        StrJson_In = ""
        StrJson_In = StrJson_In & "" & GetJsonNodeString("billtype", Split(arrInput(i), ",")(0), 1)
        StrJson_In = StrJson_In & "," & GetJsonNodeString("rcp_no", Split(arrInput(i), ",")(1), 0)
            
        If strJson_List = "" Then
            strJson_List = "{" & StrJson_In & "}"
        Else
            strJson_List = strJson_List & "," & "{" & StrJson_In & "}"
        End If
    Next
    
    strJson_List = """bill_list"":[" & strJson_List & "]"

    StrJson_In = "{""input"":{" & strJson & "," & strJson_List & "}}"
    
    strService = "Zl_Exsesvr_Setsendwin"
    
    '调用服务
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "调用“Zl_Exsesvr_Setsendwin”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallSetWindows = False
        Exit Function
    End If

    zlSplitService_CallSetWindows = True
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallPatiIsOut(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal lngPaitId As Long, ByVal lngPageId As Long, ByRef intOutSign As Integer) As Boolean
    '发药/退药调用服务：查询病人是否已出院
    '病人信息：病人id，主页id
    'intOutSign：0-未出院，1-出院
    Dim strJson As String, StrJson_In As String, strJson_Out As String
    
    On Error GoTo errHandle
    If lngPaitId = 0 Then Exit Function
    
    'Zl_Cissvr_Patiisout
    '  --功能：查询病人是否已经出院
    '  --input      查询病人是否已经出院
    '  --  pati_id               N  1  病人id
    '  --  pati_pageid           N  1  主页id
    '  --  query_type            N  1  查询类型：0-单个病人查询；1-多个病人批量查询
    '  --  pati_pageids          C  1  格式：病人ID:主页ID,病人ID:主页ID,...
    '  --output
    '  --  code                  C  1  应答码：0-失败；1-成功
    '  --  message               C  1  应答消息：失败时返回具体的错误信息
    '  --  pati_outsign          N  1  出院标记：0-未出院，1-出院；query_type=0时返回
    '  --  item_list[]           出院标记列表，query_type=1时返回
    '  --    pati_id             N  1  病人id
    '  --    pati_outsign        N  1  出院标记：0-未出院，1-出院
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lngPaitId, 1)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lngPageId, 1)
    StrJson_In = "{""input"":{" & strJson & "}}"
    
    '调用服务
    If objServiceCall.CallService("zl_CisSvr_PatiIsOut", StrJson_In, strJson_Out, "", lngMode, False) = False Then
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
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallPatiIsOutByBatch(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strPaitPages As String, ByRef cllPatiIsOut As Collection) As Boolean
    '批量查询病人是否已出院
    '入参:
    '   strPaitPages=病人信息：病人id:主页id,病人id:主页id,...
    '出参:
    '   cllPatiIsOut=病人出院情况(病人ID,出院标志)；其中，出院标志：0-未出院，1-出院
    Dim strJson As String, StrJson_In As String, strJson_Out As String
    
    On Error GoTo errHandle
    If strPaitPages = "" Then Exit Function
    
    Set cllPatiIsOut = New Collection
    'Zl_Cissvr_Patiisout
    '  --功能：查询病人是否已经出院
    '  --input      查询病人是否已经出院
    '  --  pati_id               N  1  病人id
    '  --  pati_pageid           N  1  主页id
    '  --  query_type            N  1  查询类型：0-单个病人查询；1-多个病人批量查询
    '  --  pati_pageids          C  1  格式：病人ID:主页ID,病人ID:主页ID,...
    '  --output
    '  --  code                  C  1  应答码：0-失败；1-成功
    '  --  message               C  1  应答消息：失败时返回具体的错误信息
    '  --  pati_outsign          N  1  出院标记：0-未出院，1-出院；query_type=0时返回
    '  --  item_list[]           出院标记列表，query_type=1时返回
    '  --    pati_id             N  1  病人id
    '  --    pati_outsign        N  1  出院标记：0-未出院，1-出院
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", 1, 1)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", strPaitPages, 0)
    StrJson_In = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_CisSvr_PatiIsOut", StrJson_In, , "", lngMode, False) = False Then
        MsgBox "调用“zl_CisSvr_PatiIsOut”失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '返回出参
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    If strJson_Out = "0" Then
        MsgBox objServiceCall.GetJsonNodeValue("output.message"), vbInformation, gstrSysName
        zlSplitService_CallPatiIsOutByBatch = False: Exit Function
    End If
    
    Set cllPatiIsOut = objServiceCall.GetJsonListValue("output.item_list", "pati_id")
    If cllPatiIsOut Is Nothing Then Exit Function

    zlSplitService_CallPatiIsOutByBatch = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CheckExistNoSendStuff(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal lng病人id As Long, ByVal lng库房ID As Long, ByRef blnExist As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查某病人在指定库房是否存在未发料
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim StrJson_In As String
    
    On Error GoTo ErrHandler
    'Zl_Stuffsvr_Checkpatiexecute
    '  --功能：根据病人信息或获取未发料数据
    '  --入参：Json_In:格式
    '  --  input
    '  --     check_type         N 1 检查方式:0-按如下参数值进行检查；1-仅按病人ID和发料库房进行检查
    '  --     pati_id            N 1 病人ID
    '  --     pati_pageid        N 1 主页ID
    '  --     baby_num           N 1 婴儿序号:-1表示不区分;0-母亲的;>0具体婴儿费用
    '  --     fee_source         N 1 费用来源:1-门诊;2-住院;4-体检
    '  --     stuff_nos             处方单据号，数组如：["A0001","A0002"]
    '  --     warehouse_id       N 0 库房ID
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                    N 1 应答吗：0-失败；1-成功
    '  --    message                 C 1 应答消息：失败时返回具体的错误信息
    '  --    isexist                 N 1 是否存在: 1-存在;0-不存在
    '  --    stuff_notsend_infor     C 1 未发料信息,isexist=1时返回
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("check_type", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_id", lng病人id, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("warehouse_id", lng库房ID, 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Stuffsvr_Checkpatiexecute", StrJson_In, "", "", lngMode) = False Then Exit Function
    
    blnExist = Val(nvl(objServiceCall.GetJsonNodeValue("output.isexist"))) = 1
    
    zlSplitService_CheckExistNoSendStuff = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetRequestCancel(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection) As Boolean
    '取销帐申请记录
    'strInput：费用id,费用id...
    Dim StrJson_In As String
    
    'Zl_Exsesvr_Getrequestcancel
    '  --查询销帐申请记录
    '  --入参      json
    '  --  input      查询是否存在销帐申请记录
    '  --    query_type          N 1 查询方式:0-根据费用ID查询
    '  --    rcpdtl_id           C 0 单据明细id,[数组]：[1,2,3]
    '  --    request_type        N 0 申请类别
    '  --    cancel_status       N 1 申请状态
    '  --出参      json
    '  -- output
    '  --   code     C  1   应答码：0-失败；1-成功
    '  --   message  C  1   应答消息：
    '  --   fee_cancel_list      [数组]满足条件的每个费用销帐记录
    '  --     rcpdtl_id          N    处方明细id(费用id)
    '  --     apply_type         N    申请类别:对药品和卫材有效:0-未执行;1-已执行;非药品和卫材固定存为0
    '  --     apply_time         N    申请时间:yyyy-mm-dd hh24:mi:ss
    '  --     aplnt_name         N    申请人
    '  --     apply_dept_id      N    申请部门id
    '  --     apply_dept_name    N    申请部门名称
    '  --     audit_dept_id      N    审核部门id;
    '  --     audit_dept_name    N    审核部门名称
    '  --     bill_no            N    费用单据号
    '  --     item_id            N    收费细目id
    '  --     item_name          N    收费项目名称
    '  --     quantity           N    数量
    
    On Error GoTo ErrHandler
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 0, 1)
    StrJson_In = StrJson_In & ",""rcpdtl_id"":" & "[" & strInput & "]"
    StrJson_In = StrJson_In & "," & GetJsonNodeString("request_type", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_status", 0, 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Getrequestcancel", StrJson_In, "", "", lngMode) = False Then Exit Function
    Set colOutlist = objServiceCall.GetJsonListValue("output.fee_cancel_list")
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetRequestCancel = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetExseBillByTime(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal colQueryCons As Collection, ByRef strBill_Out As String, ByRef strErrMsg_Out As String) As Boolean
    '按时间范围获取费用单据
    '入参：
    '   colQueryCons = 查询条件，成员(Key)：查询方式,费用来源,开始时间,结束时间,执行部门IDS,不含执行部门IDS
    '                           其中，费用来源：0-不区分;1-门诊;2-住院
    '出参:
    '   strBill_Out = 单据信息，格式：单据1:NO,单据1:NO,...；其中，单据：8-收费处方发药；9-记帐单处方发药；10-记帐表处方发药
    Dim StrJson_In As String
    Dim colOutlist As Collection, colTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    strBill_Out = "": strErrMsg_Out = ""
    'Zl_Exsesvr_Getbillbytime
    '  --功能：按时间范围获取费用单据
    '  --入参：json格式
    '  --  input
    '  --    query_type          N 0 查询方式:0-获取药品医嘱费用单据
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
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", colQueryCons("查询方式"), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("fee_source", colQueryCons("费用来源"), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("start_time", Format(colQueryCons("开始时间"), "yyyy-MM-dd HH:mm:ss"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("end_time", Format(colQueryCons("结束时间"), "yyyy-MM-dd HH:mm:ss"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("exe_deptids", colQueryCons("执行部门IDS"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("excp_exe_deptids", colQueryCons("不含执行部门IDS"), 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Getbillbytime", StrJson_In, "", "", lngMode, False, , , , True) = False Then Exit Function
    
    strBill_Out = objServiceCall.GetJsonNodeValue("output.bill_nos")
    
    If strBill_Out <> "" Then
        '单据类型转换：8-收费处方发药；9-记帐单处方发药；10-记帐表处方发药
        strBill_Out = "," & strBill_Out
        strBill_Out = Replace(strBill_Out, ",1:", ",8:")
        strBill_Out = Replace(strBill_Out, ",2:", ",9:")
        strBill_Out = Replace(strBill_Out, ",3:", ",10:")
        strBill_Out = zlStr.TrimEx(strBill_Out, ",")
    End If
    
    zlSplitService_GetExseBillByTime = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

