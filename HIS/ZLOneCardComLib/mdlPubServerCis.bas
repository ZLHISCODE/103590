Attribute VB_Name = "mdlPubServerCis"
Option Explicit

'*********************************************************************************************************************************************
'功能:所有涉及调用临床的相关服务
'接口说明:
'    1.zl_CisSvr_GetPatPageInfByRange-批量获取病案信息服务
'    2.zl_CisSvr_GetPatiID:根据床号等获取病人ID
'    3.zl_CIsSvr_GetPatiPageInfo-获取病人病案信息
'    4.zl_CisSvr_UpdateOutMedRecord-修改门诊病案记录信息
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
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_CisSvr_GetPatPageInfByRange(ByVal intQueryStatus As Integer, ByVal cllFilter As Collection, Optional ByVal str病人Ids As String, Optional ByRef str病区IDs As String, _
    Optional ByRef cllPatiPages_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一条范围条件，获取病人病案信息
    '入参:intQueryStatus-查询类型(0-在院病人;1-出院病人;2-在院或出院 )
    '     cllFilter-过滤条件
    '     str病人Ids-多个用逗号:病人ID或病人ID:主页ID
    '     rsPatiPage-主页信息
    '     str病区IDs-当前病区Ids
    '出参:rsPatiPageInfo_Out-返回的病人信息集
    '     strPatiIds_Out-返回当前所涉及的病人IDs
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim lng病人ID As Long, strErrMsg As String
    
    
    On Error GoTo errHandle
    
    Set cllPatiPages_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接临床域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'zl_CisSvr_GetPatPageInfByRange
    '    input
    '        query_type  N   1   查询类型:0-基本;1-基本扩展
    '        wararea_ids C       病区ids:多个用逗号
    '        pati_ids    C       病人ids:多个用逗号分离
    '        pati_pageids    C       主页IDs:病人id:主页id,…
    '        adta_start_time C       入院开始时间:yyyy-mm-dd hh24:mi:ss
    '        adta_end_time   C       入院结束时间:yyyy-mm-dd hh24:mi:ss
    '        adtd_start_time C       出院开始时间:yyyy-mm-dd hh24:mi:ss
    '        adtd_end_time   C       出院结束时间:yyyy-mm-dd hh24:mi:ss
    '        fee_category    C       费别
    '        inp_status  N       住院状态:0-在院病人;1-出院病人;2-在院或出院
    '        pati_natures    C       "病人性质：多个用逗号分离
    '        0-普通住院病人,1-门诊留观病人,2-住院留观病人
    '        NULL-表示不区分"

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", intQueryStatus, Json_num)
    If InStr(str病人Ids, ":") > 0 Then
        strJson = strJson & "," & GetJsonNodeString("pati_pageids", str病人Ids, Json_Text)  '按病人id+主页id传入
    Else
        strJson = strJson & "," & GetJsonNodeString("pati_ids", str病人Ids, Json_Text)
    
    End If
    strJson = strJson & "," & GetJsonNodeString("wararea_ids", str病区IDs, Json_Text)
    
    For i = 1 To cllFilter.count
    
        Select Case cllFilter(i)(0)
        Case "入院日期"
            strJson = strJson & "," & GetJsonNodeString("adta_start_time", cllFilter(i)(1), Json_Text)
            strJson = strJson & "," & GetJsonNodeString("adta_end_time", cllFilter(i)(2), Json_Text)
        Case "出院日期"
            strJson = strJson & "," & GetJsonNodeString("adtd_start_time", cllFilter(i)(1), Json_Text)
            strJson = strJson & "," & GetJsonNodeString("adtd_end_time", cllFilter(i)(2), Json_Text)
        Case "费别"
            strJson = strJson & "," & GetJsonNodeString("fee_category", cllFilter(i)(1), Json_Text)
        Case "病人性质"
            strJson = strJson & "," & GetJsonNodeString("pati_natures", cllFilter(i)(1), Json_Text)
        End Select
    Next
    
    strJson = "{""input"":{" & strJson & "}}"
    strServiceName = "zl_CisSvr_GetPatPageInfByRange"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    '出参            json    基本    扩展
    'output
    '    code    N       应答码：0-失败；1-成功  √  √
    '    message C       应答消息： 失败时返回具体的错误信息 √  √
    '    page_list[]         数据组  √  √
    '        pati_id N       病人id  √  √
    '        pati_pageid N       主页id  √  √
    '        pati_name   C       姓名    √  √
    '        pati_sex    C       性别    √  √
    '        pati_age    C       年龄    √  √
    '        inpatient_num   C       住院号  √  √
    '        pati_bed    C       当前床号    √  √
    '        insurance_type  N       险类    √  √
    '        fee_category    C       费别    √  √
    '        pati_type   C       病人类型(普通，医保，留观)  √  √
    '        adta_time   C       入院时间:yyyy-mm-dd hh24:mi:ss  √  √
    '        adtd_time   C       出院时间:yyyy-mm-dd hh24:mi:ss  √  √
    '        si_inp_status   N       住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)  √  √
    Set cllData = objServiceCall.GetJsonListValue("output.page_list")
    
    If cllData Is Nothing Then
        strErrMsg = "未找到符合条件的病人信息，请检查！"
         MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If cllData.count = 0 Then
          strErrMsg = "未找到符合条件的病人信息，请检查！"
          MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
          Exit Function
    End If
    Set cllPatiPages_Out = cllData
    zl_CisSvr_GetPatPageInfByRange = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zl_CisSvr_GetPatiID(cllFindCons As Collection, _
    Optional ByRef lng主页ID_out As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据床号及住院号获取病人ID
    '入参:
    '   cllFindCons-查找条件:array(查询的名称,查询的内容)
    '               查询的名称:如:住院号,留观号,(病区ID,床号)等
    '出参:
    '   lng主页ID_out-返回当前病人的主页ID
    '返回:成功返回病人ID,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHandle
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'zl_CisSvr_GetPatiID
    '  --根据床号、住院呈获取病人ID及主页ID
    '  --input
    '  --   wardarea_id          N 1 当前病区id
    '  --   pati_bed             C 1 当前床号
    '  --   inpatient_num        C 1 住院号
    '  --   obsv_no              C 1 留观号
    '  --output
    '  --    code                N 1 应答码：0-失败；1-成功
    '  --    message             C 1 应答消息： 失败时返回具体的错误信息
    '  --    pati_id             N 1 病人ID:未找到时也成功，返回0
    '  --    pati_pageid         N   主页ID
    strJson = ""
    For i = 1 To cllFindCons.count
        Select Case cllFindCons(i)(0)
        Case "病区ID"
            strJson = strJson & "," & GetJsonNodeString("wardarea_id", cllFindCons(i)(1), Json_num)
        Case "床号"
            strJson = strJson & "," & GetJsonNodeString("pati_bed", cllFindCons(i)(1), Json_Text)
        Case "住院号"
            strJson = strJson & "," & GetJsonNodeString("inpatient_num", cllFindCons(i)(1), Json_Text)
        Case "留观号"
            strJson = strJson & "," & GetJsonNodeString("obsv_no", cllFindCons(i)(1), Json_Text)
        End Select
    Next
    strJson = "{""input"":{" & Mid(strJson, 2) & "}}"
  
    strServiceName = "zl_CisSvr_GetPatiID"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    lng主页ID_out = Val(NVL(objServiceCall.GetJsonNodeValue("output.pati_pageid")))
    zl_CisSvr_GetPatiID = Val(NVL(objServiceCall.GetJsonNodeValue("output.pati_id")))
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_CIsSvr_GetPatiPageInfo(ByVal int查询类型 As Integer, ByVal str病人主页IDs As String, ByRef cllPatiPage_Out As Variant, _
    Optional ByRef bln仅取最后住院 As Boolean, Optional bln含婴儿信息 As Boolean, Optional ByRef bln含转科信息 As Boolean, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID或主页id信息来获取病案主页信息
    '入参:int查询类型-0-只获取基本信息;1-获取基本信息+扩展信息;2-仅获取取主页ID字段
    '     str病人主页IDs-两种格式:
    '           1.病人id1:主页id1,病人id2:主页id2...
    '           2.病人id1,病人id2,...病人idn
    '      bln仅取最后住院:主读取病人最后一次的病案,(str病人主页IDs第二种格式有效)
    '      bln含婴儿信息:是否包含婴儿信息
    '      bln含转科信息:是否包转科信息
    '出参:cllPatiPageInfo_Out-返回的病案信息集
    '     strErrMsg_Out-返回的错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-09-19 15:50:18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer

    On Error GoTo errHandle
    
    Set cllPatiPage_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    'input
    '    query_type  C   1   查询类型:0-基本信息;1-基本信息的展;2-仅取主页
    '    pati_pageids    C   1   病人信息,格式两种:一种是:病人id:主页ID,…;一种：病人id,…
    '    is_lastpage N   1   是否取最后一次住院
    '    is_babyinfo N   1   是否包含婴儿信息:1-包含;0-不包含
    '    is_transdeptinfo    N   1   是否包含转科信息:1-包含;0-不包含

    strJson = strJson & "" & GetJsonNodeString("query_type", int查询类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", str病人主页IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("is_lastpage", IIf(bln仅取最后住院, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("is_babyinfo", IIf(bln含婴儿信息, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("is_transdeptinfo", IIf(bln含转科信息, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    strServiceName = "zl_CIsSvr_GetPatiPageInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    '    出参            json    基本    扩展    只取主页
    '    output
    '        code    N   1   应答码：0-失败；1-成功  √  √  √
    '        message C   1   应答消息： 失败时返回具体的错误信息 √  √  √
    '        page_list[]     1   数据组
    '        pati_id N   1   病人id  √  √  √
    '        pati_pageid N   1   主页id  √  √  √
    '        pati_name   C   1   姓名    √  √  √
    '        pati_sex    C   1   性别    √  √
    '        pati_age    C   1   年龄    √  √
    '        fee_category    C   1   费别    √  √
    '        mdlpay_mode_name    C   1   医疗付款方式名称        √
    '        mdlpay_mode_code    C   1   医疗付款方式编码        √
    '        pati_bed    C   1   当前床号
    '        pati_type   C   1   病人类型(普通，医保，留观)
    '        pati_education  C   1   学历
    '        ocpt_name   C   1   职业
    '        country_name    C   1   国籍
    '        pati_marital_cstatus    C   1   婚姻状况
    '        pati_nature N   1   病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人   √
    '        audit_sign  N   1   审核标志:病案主页.审核标志  √  √
    '        si_inp_status   N   1   住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)  √  √
    '        pati_wardarea_id    N   1   当前病区id      √
    '        pati_deptid N   1   当前科室id      √
    '        pati_wardarea_id    N   1   当前病区id      √
    '        pati_wardarea_name  C   1   当前病区名称        √
    '        pati_dept_id    N   1   当前科室id      √
    '        pati_dept_name  C   1   当前科室名称        √
    '        adta_time   C   1   入院时间:yyyy-mm-dd hh24:mi:ss  √  √
    '        adtd_time   C   1   出院时间:yyyy-mm-dd hh24:mi:ss  √  √
    '        insurance_type  N   1   险类        √
    '        scheme_type C   1   适用病人:Zl_Patiwarnscheme      √
    '        garnt_money     1   担保额:Zl_Patientsurety     √
    '        catalog date    C   1   编目日期:yyyy-mm-dd hh24:mi:ss      √
    '        baby_list[]     1   婴儿信息，[数组]    is_babyinfo=1
    '            pati_id N   1   病人id
    '            pati_pageid N   1   主页id
    '            baby_num    N   1   婴儿序号
    '            baby_name   C   1   婴儿姓名
    '            baby_sex    C   1   婴儿性别
    '            baby_date   D   1   出生时间
    '        trans_list[]    C       转科列表信息    is_transdeptinfo=1
    '            start_reason    C   1   开始原因
    '            start_time  C   1   开始时间:yyyy-mm-dd hh24:mi:ss
    '            dept_name   C   1   科室名称
    
        
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "未找到符合条件的病案信息，请检查！"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    Set cllData = objServiceCall.GetJsonListValue("output.page_list")
'    If clldata Is Nothing Then
'            strErrMsg_Out = "未找到符合条件的病案信息，请检查！"
'        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If
'
'    If clldata.count = 0 Then
'        strErrMsg_Out = "未找到符合条件的病案信息，请检查！"
'        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If
    Set cllPatiPage_Out = cllData
    zl_CIsSvr_GetPatiPageInfo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function zl_CisSvr_UpdateOutMedRecord(ByVal cllOutMedRec As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新门诊病案记录
    '入参:cllOutMedRec-门诊病案数据集:array(名称,值)
    '                名称包含（病人id,.病案号(门诊号),建立日期,病案类别,存储状态,存放位置)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    
    If cllOutMedRec Is Nothing Then Exit Function
    If cllOutMedRec.count = 0 Then Exit Function
    
    On Error GoTo errHandle
    If GetServiceCall(objServiceCall, False) = False Then
        strErrMsg = "连接费用域服务失败，无法获取有效的病人信息!"
        Err.Raise -1001, strErrMsg, strErrMsg
        Exit Function
    End If
    
    For i = 1 To cllOutMedRec.count
        varTemp = cllOutMedRec(i)
        Select Case UCase(varTemp(0))
        Case "病人ID"
            strJson = strJson & "," & GetJsonNodeString("pati_id", Val(varTemp(1)), Json_num)
        Case "病案号"
            strJson = strJson & "," & GetJsonNodeString("mr_no", varTemp(1), Json_Text, True)
        Case "门诊号"
            strJson = strJson & "," & GetJsonNodeString("outpatient_num", varTemp(1), Json_Text, True)
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
    '       mr_no   N   1   病案号（门诊号）
    '       create_date C   1   建立日期
    '       mr_type C   1   病案类别
    '       strgloc_status  C   1   存储状态
    '       strgloc C   1   存放位置
    
    strJson = Mid(strJson, 2)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_CisSvr_UpdateOutMedRecord"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    '    output
    '        code    C   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：失败时返回具体的错误信息
    zl_CisSvr_UpdateOutMedRecord = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function Zl_CisSvr_PatiIsInhospital(ByVal lng病人ID As Long, ByRef blnInhospital As Boolean, _
                Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查病人是否在院就诊
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 14:35
    '---------------------------------------------------------------------------------------
    Dim intReturn As Integer
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHandle
    blnInhospital = False
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_CisSvr_Patiisinhospital
    '    input
    '        pati_id            N 1 病人ID
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_CisSvr_PatiIsInhospital"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, Not blnNotShowErrMsg, strErrMsg) = False Then Exit Function
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
         
    blnInhospital = Val(objServiceCall.GetJsonNodeValue("output.inhouspital")) = 1
    Zl_CisSvr_PatiIsInhospital = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function


Public Function zl_Cissvr_Existadvice(ByVal lng病人ID As Long, ByVal str挂号单 As String, ByRef blnHavAdvice As Boolean, _
                Optional ByVal lng主页Id As Long, Optional ByVal blnOnlyValid As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查挂号单是否发生了医嘱
    ' 入参 : str挂号单-多个单据号间用逗号分隔
    '        blnOnlyValid-是否只检查有效医嘱
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 19:48
    '---------------------------------------------------------------------------------------
    
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'zl_CisSvr_GetPatiID
    '    input
    '    --   pati_id              N 1 病人ID
    '    --   pati_pageid          N   主页Id
    '    --   rgst_no              C 1 挂号单，多个用逗号分隔
    '    --   only_valid           N   只检查没有作废的医嘱
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rgst_no", str挂号单, Json_Text)
    If lng主页Id <> 0 Then
        strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页Id, Json_num)
    End If
    strJson = strJson & "," & GetJsonNodeString("only_valid", IIf(blnOnlyValid, 1, 0), Json_num)
    
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Cissvr_Existadvice"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    blnHavAdvice = Val(objServiceCall.GetJsonNodeValue("output.exist")) = 1
         
    zl_Cissvr_Existadvice = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Zl_Cissvr_GetPatiVitalSigns(ByVal lng病人ID As Long, ByVal lng挂号ID As String, _
                ByRef cllVital As Collection, Optional ByVal blnOutPati As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人生命体征信息
    ' 入参 : blnOutPati-门诊病人
    ' 出参 : cllVital:体征信息(Collect)(项目,值,单位)
    ' 返回 : 返回病人的体征信息，包括项目，数值，单位
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 19:48
    '---------------------------------------------------------------------------------------
    
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接病人域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Cissvr_GetPatiVitalSigns
    '    input
    '    --   pati_id              N 1 病人ID
    '    --   visit_id             N 1 挂号ID
    '    --   outpati_flag         N   门诊标志：1-门诊，2-住院
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("visit_id", lng挂号ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("outpati_flag", IIf(blnOutPati, 1, 0), Json_num)
    
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Cissvr_GetPatiVitalSigns"
    If objServiceCall.CallService(strServiceName, strJson, strServiceName, "", glngModul) = False Then Exit Function
    Set cllVital = objServiceCall.GetJsonListValue("output.pativital_list")
    
    Zl_Cissvr_GetPatiVitalSigns = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Zl_Cissvr_Checkdepositno(ByVal lng病人ID As Long, ByRef strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断预交NO是否存在"病人结算异常记录"中
    '入参:strNo-预交单据号
    '返回:传入的NO号存在"病人结算异常记录"中返回true,否则返回False
    '编制:焦博
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接临床域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Cissvr_Checkdepositerrorno
    '    input
    '    --   pati_id              N 1 病人ID
    '    --   bill_nos             C 1 病人预交记录.NO
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_nos", strNo, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Cissvr_Checkdepositerrorno"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    strNo = objServiceCall.GetJsonNodeValue("output.bill_nos")
    Zl_Cissvr_Checkdepositno = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Zl_Cissvr_GetPatiBaseInfo(ByVal lng病人ID As Long, Optional ByVal lng主页Id As Long = -1, _
                Optional ByRef cllPati_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人住院信息
    '入参:lng病人ID；lng主页ID
    '返回:cllPati_Out 病人信息
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接临床域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Cissvr_Checkdepositerrorno
'      input
'  --    query_type        N 1 查询方式-- 1-通过病人ID+主页ID查询病人信息,2-通过医嘱ID获取病人基本信息 ,3-通过挂号单获取病人基本信息
'  --    pati_id           N   病人id--
'  --    page_id           N   主页id--
'  --    advice_id         N   医嘱ID--
'  --    pati_type         N   0-住院病人 1-门诊病人
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    If lng主页Id <> -1 Then
        strJson = strJson & "," & GetJsonNodeString("page_id", lng主页Id, Json_num)
    End If
    strJson = strJson & "," & GetJsonNodeString("query_type", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_type", 0, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Cissvr_Getpatibaseinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    Set cllPati_Out = objServiceCall.GetJsonListValue("output.page_list")
    Zl_Cissvr_GetPatiBaseInfo = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zl_Cissvr_GetInpatiState(ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
                Optional ByVal intPatiType As Integer, _
                Optional ByRef cllPati_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人住院状态
    '入参:lng病人ID；lng主页ID
    '     intPatiType:病人性质 0-普通住院病人 1-门诊留观病人 2-住院留观病人
    '返回:cllPati_Out 病人信息
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHand
    
    Set cllPati_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接临床域服务失败，无法获取有效的病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Cissvr_Checkdepositerrorno
'      input
'    pati_id       N   1   病人ID
'    pati_pageid   N   1   主页id
'    pati_type     N   1   病人性质 0-普通住院病人 1-门诊留观病人 2-住院留观病人

    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页Id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_type", intPatiType, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    strServiceName = "zl_Cissvr_GetInpatiState"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    If objServiceCall.GetJsonNodeValue("output.pati_type") = "" Then Exit Function
    If Val(objServiceCall.GetJsonNodeValue("output.pati_type")) <> intPatiType Then Exit Function
    cllPati_Out.Add Val(NVL(objServiceCall.GetJsonNodeValue("output.pati_state"))), "病人状态"
    cllPati_Out.Add NVL(objServiceCall.GetJsonNodeValue("output.out_time")), "出院日期"
    zl_Cissvr_GetInpatiState = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function
