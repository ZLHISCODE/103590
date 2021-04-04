Attribute VB_Name = "mdlPubServiceCis"
Option Explicit


'*********************************************************************************************************************************************
'功能:所有涉及调用临床的相关服务
'接口说明:
'    1.zl_CisSvr_ExistAdvice-判断指定的挂号单或病人是否存在医嘱数据
'    2.zl_CisSvr_UpdateOutMedRecord-更新门诊病案记录
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
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zl_CisSvr_ExistAdvice(ByVal str挂号单 As String, ByVal lng病人ID As Long, ByRef bln存在医嘱_Out As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional lngModule As Long, Optional str主页id As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定的挂号单或病人是否存在医嘱数据
    '入参:str挂号单-指定的挂号单
    '     lng病人ID-病人id
    '     blnShowErrMsg-是否显示错误信息
    '     str主页id-主页id=""时，表示不按主页id查找;0时表示只查主页id为零的医嘱,>0表示查询指定主页的医嘱
    '出参:bln存在医嘱_Out
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
    
    'zl_CisSvr_ExistAdvice
    '  --功能:判断是否存在医嘱数据或判断指定挂号单是否已经开医嘱
    '  --入参：Json_In:格式
    '  -- input
    '  --   pati_id              N 1 病人ID
    '  --   pati_pageid          N   主页Id
    '  --   rgst_no              C 1 挂号单，多个用逗号分隔
    '  --   only_valid           N   只检查没有作废的医嘱
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               C 1 应答码：0-失败；1-成功
    '  --    message            C 1 应答消息：失败时返回具体的错误信息
    '  --    exist              N 1 是否存在，1-存在;0-不存在
    strJson = strJson & "" & GetJsonNodeString("rgst_no", str挂号单, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    If str主页id <> "" Then
        strJson = strJson & "" & GetJsonNodeString("pati_pageid", Val(str主页id), Json_num)
    End If
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_CisSvr_ExistAdvice"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能获取指定条件下的医嘱数据，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln存在医嘱_Out = Val(NVL(objServiceCall.GetJsonNodeValue("output.exist"))) = 1
    zl_CisSvr_ExistAdvice = True
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




Public Function zl_CisSvr_UpdateOutMedRecord(ByVal cllOutMedRec As Collection, Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新门诊病案记录
    '入参:cllOutMedRec-门诊病案数据集:array(名称,值)
    '                名称包含（病人id,.病案号(门诊号),建立日期,病案类别,存储状态,存放位置)
    '     blnShowErrMsg-是否显示错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    
    If cllOutMedRec Is Nothing Then Exit Function
    If cllOutMedRec.Count = 0 Then Exit Function
    If blnShowErrMsg Then On Error GoTo errHandle:
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "连接费用域服务失败，无法获取有效的病人信息!"
        Err.Raise -1001, strErrMsg, strErrMsg
        Exit Function
    End If
    
    For i = 1 To cllOutMedRec.Count
        varTemp = cllOutMedRec(i)
        Select Case UCase(varTemp(0))
        Case "病人ID"
            strJson = strJson & "," & GetJsonNodeString("pati_id", Val(varTemp(1)), Json_num)
        Case "病案号", "门诊号"
            strJson = strJson & "," & GetJsonNodeString("mr_no", Trim(varTemp(1)), Json_Text)
        Case "建立日期", "建档日期", "登记日期", "登记时间"
            strJson = strJson & "," & GetJsonNodeString("create_date", Trim(varTemp(1)), Json_Text)
        Case "病案类别"
            strJson = strJson & "," & GetJsonNodeString("mr_type", Trim(varTemp(1)), Json_Text)
        Case "存储状态"
            strJson = strJson & "," & GetJsonNodeString("strgloc_status", Trim(varTemp(1)), Json_Text)
        Case "存放位置"
            strJson = strJson & "," & GetJsonNodeString("strgloc", Trim(varTemp(1)), Json_Text)
        Case Else
        End Select
    Next
    If strJson = "" Then Exit Function
    
    'zl_CisSvr_UpdateOutMedRecord
    '    input
    '       pati_id N   1   病人id
    '       mr_no   C   1   病案号（门诊号）
    '       create_date C   1   建立日期
    '       mr_type C   1   病案类别
    '       strgloc_status  C   1   存储状态
    '       strgloc C   1   存放位置
    
    strJson = Mid(strJson, 2)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_CisSvr_UpdateOutMedRecord"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, blnShowErrMsg, , , , blnShowErrMsg) = False Then Exit Function
    '    output
    '        code    C   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：失败时返回具体的错误信息
    zl_CisSvr_UpdateOutMedRecord = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_CIsSvr_GetPatiPageInfo(ByVal int查询类别 As Integer, ByVal strPatiInfo As String, ByRef cllPatiPages_Out As Collection, _
    Optional ByVal bln含婴儿信息 As Boolean = False, Optional ByVal bln含转科信息 As Boolean = False, Optional bln取最后一次住院 As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID或主页ID，查询病案信息
    '入参:int查询类别-:0-基本信息;1-基本信息的展;2-仅取主页
    '     strPatiInfo-病人信息,格式两种:一种是:病人id:主页ID,…;一种：病人id,…
    '     blnShowErrMsg-是否显示错误信息
    '出参:cllPatiPages_Out-返回主页信息
    '     strErrmsg_Out-返回错误信息
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
    
    'zl_CIsSvr_GetPatiPageInfo
    ' input
    '    query_type  C   1   查询类型:0-基本信息;1-基本信息的展;2-仅取主页
    '    pati_pageids    C   1   病人信息,格式两种:一种是:病人id:主页ID,…;一种：病人id,…
    '    is_babyinfo N   1   是否包含婴儿信息:1-包含;0-不包含
    '    is_transdeptinfo    N   1   是否包含转科信息:1-包含;0-不包含
    '    is_lastpage N   1   是否取最后一次住院
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", int查询类别, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", strPatiInfo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("is_babyinfo", IIf(bln含婴儿信息, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("is_transdeptinfo", IIf(bln含转科信息, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("is_lastpage", IIf(bln取最后一次住院, 1, 0), Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_CIsSvr_GetPatiPageInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, blnShowErrMsg) = False Then Exit Function
    '出参            json    基本    扩展    只取主页
    'output
    '    code    N   1   应答码：0-失败；1-成功  √  √  √
    '    message C   1   应答消息： 失败时返回具体的错误信息 √  √  √
    '    page_list[]     1   数据组
    '        pati_id N   1   病人id  √  √  √
    '        pati_pageid N   1   主页id  √  √  √
    '        pati_name   C   1   姓名    √  √  √
    '        pati_sex    C   1   性别    √  √
    '        pati_age    C   1   年龄    √  √
    '        fee_category    C   1   费别    √  √
    '        mdlpay_mode_name    C   1   医疗付款方式名称        √
    '        mdlpay_mode_code    C   1   医疗付款方式编码        √
    '        pati_bed    C   1   当前床号        √
    '        pati_type   C   1   病人类型(普通，医保，留观)      √
    '        pati_education  C   1   学历        √
    '        ocpt_name   C   1   职业        √
    '        country_name    C   1   国籍        √
    '        pati_marital_cstatus    C   1   婚姻状况        √
    '        pati_nature N   1   病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人   √  √
    '        audit_sign  N   1   审核标志:病案主页.审核标志  √  √
    '        si_inp_status   N   1   住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)  √  √
    '        pati_wardarea_id    N   1   当前病区id      √
    '        pati_wardarea_name  C   1   当前病区名称        √
    '        pati_dept_id    N   1   当前科室id      √
    '        pati_dept_name  C   1   当前科室名称        √
    '        adta_time   C   1   入院时间:yyyy-mm-dd hh24:mi:ss  √  √
    '        adtd_time   C   1   出院时间:yyyy-mm-dd hh24:mi:ss  √  √
    '        insurance_type  N   1   险类        √
    '        scheme_type C   1   适用病人:Zl_Patiwarnscheme      √
    '        garnt_money N   1   担保额:Zl_Patientsurety     √
    '        catalog_date    C   1   编目日期:yyyy-mm-dd hh24:mi:ss      √
    '        baby_list[]     1   婴儿信息，[数组]    is_babyinfo=1
    '            pati_id N   1   病人id
    '            pati_pageid N   1   主页id
    '            baby_num    N   1   婴儿序号
    '            baby_name   C   1   婴儿姓名
    '            baby_sex    C   1   婴儿性别
    '            baby_date   C   1   出生时间
    '        trans_list[]    C       转科列表信息    is_transdeptinfo=1
    '            start_reason    C   1   开始原因
    '            start_time  C   1   开始时间:yyyy-mm-dd hh24:mi:ss
    '            dept_name   C   1   科室名称
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "不能根据条件获取病案信息，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set cllPatiPages_Out = objServiceCall.GetJsonListValue("output.page_list")
    zl_CIsSvr_GetPatiPageInfo = True
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



