Attribute VB_Name = "mdlPubServicePati"
Option Explicit

'*********************************************************************************************************************************************
'功能:所有涉及调用费用的相关服务
'接口说明:
'    1.Zl_病人结算异常记录_Modify-病人异常数据增加,修改及删除
'    2.zl_PatiSvr_NewPatiArchives-新建病人档案
'    3.zl_Patisvr_GetNextNo-获取门诊号及病人ID信息
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

Public Function zl_Patisvr_GetNextNo(ByVal int序号 As Integer, ByRef strNo_Out As String, Optional ByVal lng科室ID As Long, _
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
    strServiceName = "Zl_Patisvr_GetNextNo"
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
    zl_Patisvr_GetNextNo = True
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
 
Public Function Zl_病人结算异常记录_Modify(ByVal int操作状态 As Integer, ByVal cllSaveData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:异常结算数据修正
    '入参:int操作状态-操作状态:0-新增记录,1-更新状态及更新交易说明，2-删除异常数据
    '    cllSaveData-格式为Array(保存项名称,保存项值)
    '          保存的项目名称包含: 异常ID,操作场景,作废标志,业务id,是否病历费,病人id,主页id,姓名,性别,年龄,门诊号,住院号,预交单号,预交金额,医疗卡单号,卡费,发卡类别id,发卡类别名称,发卡卡号,同步状态,交易信息)
    '          其中交易信息为Json串，格式如下
    '           {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]}
    '     blnShowErrMsg-是否显示错误信息
    '     str主页id-主页id=""时，表示不按主页id查找;0时表示只查主页id为零的医嘱,>0表示查询指定主页的医嘱
    '出参:strErrmsg_Out-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objExpense As Object
    If zlGetPubExpenseObject(objExpense) = False Then Exit Function
    Zl_病人结算异常记录_Modify = objExpense.Zl病人结算异常记录_Modify(int操作状态, cllSaveData)
End Function


Public Function Zl_医疗卡变动_Insert_Check(ByVal cllCheckData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医疗卡变动前检查数据的合法性
    '入参:cllCheckData-格式为Array(保存项名称,保存项值)
    '                   保存的项目名称包含:操作状态,病人ID,卡类别ID,卡号,新卡号,异常状态
    '                    操作状态:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int操作状态 As Integer, lng病人id As Long, lng卡类别ID As Long, str卡号 As String, str新卡号 As String, int异常状态 As Integer
    Dim varRetrun As Variant, strErrMsg As String
    Dim i As Long
    
    On Error GoTo errHandle
    If cllCheckData Is Nothing Then
        strErrMsg = "未传入需要检查的必要条件，请检查!"
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    For i = 1 To cllCheckData.Count
        Select Case UCase(cllCheckData(i)(0))
        Case "操作状态"
            int操作状态 = Val(cllCheckData(i)(1))
        Case "病人ID"
            lng病人id = Val(cllCheckData(i)(1))
        Case "卡类别ID"
            lng卡类别ID = Val(cllCheckData(i)(1))
        Case "卡号"
            str卡号 = Trim(cllCheckData(i)(1))
        Case "新卡号"
            str新卡号 = Trim(cllCheckData(i)(1))
        Case "异常状态"
            int异常状态 = Val(cllCheckData(i)(1))
        End Select
    Next
    
    '  Zl_医疗卡变动_Insert_Check
    '操作状态_In  Integer,
    '卡类别id_In  医疗卡类别.Id%Type,
    '卡号_In      病人医疗卡变动.卡号%Type,
    '新卡号_In    病人医疗卡变动.卡号%Type,
    '病人id_In    病人医疗卡变动.病人id%Type,
    '应签码_Out   Out Integer,
    '应答信息_Out Out Varchar2
    '
    varRetrun = zlDatabase.CallProcedure("Zl_医疗卡变动_Insert_Check", "医疗卡变动检查", int操作状态, lng卡类别ID, str卡号, str新卡号, lng病人id, int异常状态, Empty, Empty)
    
    If varRetrun(0) <> 1 Then
        strErrMsg = varRetrun(1)
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Zl_医疗卡变动_Insert_Check = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function




Public Function zl_PatiSvr_NewPatiArchives(ByVal cllUpdBasePati As Collection, ByVal cllUpdContacts As Collection, _
    ByVal cllUpdCommunity As Collection, ByVal cllUpdVisit As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医疗卡变动前检查数据的合法性
    '入参:cllUpdBasePati-修改病人基本信息:病人ID,姓名,性别,年龄,出生日期(yyyy-mm-dd hh24:mi:ss),身份证号,病人类型,门诊号,就诊卡号,卡验证码,费别,医疗付款方式名称,籍贯,国籍,民族,婚姻状况,学历,职业,身份,工作单位,
    '            单位邮编,单位电话,单位开户行,单位帐号,合同单位ID,家庭地址,家庭电话,家庭地址邮编,区域,出生地点,户口地址,户口地址邮编,监护人,手机号,
    '            医保号,Ic卡号,登记时间,操作员姓名,身份证签约,签约密码,险类,其他证件）
    '    cllUpdContacts-修改病人联系人信息:(联系人姓名,联系人身份证号,联系人电话,联系人关系,联系人地址)
    '    cllUpdCommunity-修改社区信息:社区序号,社区号码,社区操作类型
    '    cllUpdVisit-更新就诊信息:就诊状态,就诊诊室,就诊时间
    '   以上集合格式:Array(保存项名称,保存项值)
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String, strJsonTemp As String
    Dim clldata As Variant, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    

    On Error GoTo errHandle
    If GetServiceCall(objServiceCall, True) = False Then Exit Function
    If Not cllUpdBasePati Is Nothing Then
        strJson = ""
        For i = 1 To cllUpdBasePati.Count
            varTemp = cllUpdBasePati(i)
            Select Case varTemp(0)
            Case "病人ID"
                strJson = strJson & "," & GetJsonNodeString("pati_id", Val(varTemp(1)), Json_num, True)
            Case "姓名"
                strJson = strJson & "," & GetJsonNodeString("pati_name", Trim(varTemp(1)), Json_Text)
            Case "性别"
                strJson = strJson & "," & GetJsonNodeString("pati_sex", Trim(varTemp(1)), Json_Text)
            Case "年龄"
                strJson = strJson & "," & GetJsonNodeString("pati_age", Trim(varTemp(1)), Json_Text)
            Case "出生日期" ':yyyy-mm-dd hh24:mi:ss
                strJson = strJson & "," & GetJsonNodeString("pati_birthdate", Trim(varTemp(1)), Json_Text)
            Case "身份证号"
                strJson = strJson & "," & GetJsonNodeString("pati_idcard", Trim(varTemp(1)), Json_Text)
            Case "病人类型"
                strJson = strJson & "," & GetJsonNodeString("pati_type", Trim(varTemp(1)), Json_Text)
            Case "门诊号"
                strJson = strJson & "," & GetJsonNodeString("outpatient_num", Val(varTemp(1)), Json_num, True)
            Case "就诊卡号"
                strJson = strJson & "," & GetJsonNodeString("vcard_no", Trim(varTemp(1)), Json_Text)
            Case "卡验证码"
                strJson = strJson & "," & GetJsonNodeString("vcard_pwd", Trim(varTemp(1)), Json_Text)
            Case "费别"
                strJson = strJson & "," & GetJsonNodeString("fee_category", Trim(varTemp(1)), Json_Text)
            Case "医疗付款方式名称"
                strJson = strJson & "," & GetJsonNodeString("mdlpay_mode_name", Trim(varTemp(1)), Json_Text)
            Case "籍贯"
                strJson = strJson & "," & GetJsonNodeString("native_place", Trim(varTemp(1)), Json_Text)
            Case "国籍"
                strJson = strJson & "," & GetJsonNodeString("country_name", Trim(varTemp(1)), Json_Text)
            Case "民族"
                strJson = strJson & "," & GetJsonNodeString("nation_name", Trim(varTemp(1)), Json_Text)
            Case "婚姻状况"
                strJson = strJson & "," & GetJsonNodeString("mari_status", Trim(varTemp(1)), Json_Text)
            Case "学历"
                strJson = strJson & "," & GetJsonNodeString("edu_name", Trim(varTemp(1)), Json_Text)
            Case "职业"
                strJson = strJson & "," & GetJsonNodeString("ocpt_name", Trim(varTemp(1)), Json_Text)
            Case "身份"
                strJson = strJson & "," & GetJsonNodeString("pati_identity", Trim(varTemp(1)), Json_Text)
            Case "工作单位"
                strJson = strJson & "," & GetJsonNodeString("emp_name", Trim(varTemp(1)), Json_Text)
            Case "单位邮编"
                strJson = strJson & "," & GetJsonNodeString("emp_postcode", Trim(varTemp(1)), Json_Text)
            Case "单位电话"
                strJson = strJson & "," & GetJsonNodeString("emp_phno", Trim(varTemp(1)), Json_Text)
            Case "单位开户行"
                strJson = strJson & "," & GetJsonNodeString("emp_bank_name", Trim(varTemp(1)), Json_Text)
            Case "单位帐号"
                strJson = strJson & "," & GetJsonNodeString("emp_bank_accnum", Trim(varTemp(1)), Json_Text)
            Case "合同单位ID"
                strJson = strJson & "," & GetJsonNodeString("ctt_unit_id", Val(varTemp(1)), Json_num, True)
            Case "家庭地址"
                strJson = strJson & "," & GetJsonNodeString("pat_home_addr", Trim(varTemp(1)), Json_Text)
            Case "家庭电话"
                strJson = strJson & "," & GetJsonNodeString("pat_home_phno", Trim(varTemp(1)), Json_Text)
            Case "家庭地址邮编"
                strJson = strJson & "," & GetJsonNodeString("pat_home_postcode", Trim(varTemp(1)), Json_Text)
            Case "区域"
                strJson = strJson & "," & GetJsonNodeString("region", Trim(varTemp(1)), Json_Text)
            Case "出生地点"
                strJson = strJson & "," & GetJsonNodeString("pat_baddr", Trim(varTemp(1)), Json_Text)
            Case "户口地址"
                strJson = strJson & "," & GetJsonNodeString("pat_hous_addr", Trim(varTemp(1)), Json_Text)
            Case "户口地址邮编"
                strJson = strJson & "," & GetJsonNodeString("pat_hous_postcode", Trim(varTemp(1)), Json_Text)
            Case "监护人"
                strJson = strJson & "," & GetJsonNodeString("pat_grdn_name", Trim(varTemp(1)), Json_Text)
            Case "手机号"
                strJson = strJson & "," & GetJsonNodeString("phone_number", Trim(varTemp(1)), Json_Text)
            Case "医保号"
                strJson = strJson & "," & GetJsonNodeString("insurance_num", Trim(varTemp(1)), Json_Text)
            Case "IC卡号"
                strJson = strJson & "," & GetJsonNodeString("iccard_no", Trim(varTemp(1)), Json_Text)
            Case "登记时间" ':yyyy-mm-dd hh24:mi:ss
                strJson = strJson & "," & GetJsonNodeString("create_time", Trim(varTemp(1)), Json_Text)
            Case "操作员姓名 "
                strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
            Case "身份证签约"
                strJson = strJson & "," & GetJsonNodeString("idcard_sign", Val(varTemp(1)), Json_num, True)
            Case "签约密码"
                strJson = strJson & "," & GetJsonNodeString("idcard_sign_pwd", Trim(varTemp(1)), Json_Text)
            Case "险类"
                strJson = strJson & "," & GetJsonNodeString("insurance_type", Trim(varTemp(1)), Json_Text)
            Case "其他证件"
                strJson = strJson & "," & GetJsonNodeString("cert_no_other", Trim(varTemp(1)), Json_Text)
            Case ""
            End Select
        Next
    End If
    
    If Not cllUpdContacts Is Nothing Then
        '更新联系人
        strJsonTemp = ""
        For i = 1 To cllUpdContacts.Count
            varTemp = cllUpdContacts(i)
            Select Case UCase(varTemp(0))
            Case "联系人姓名"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("name", varTemp(1), Json_Text)
            Case "联系人身份证号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("idcard", varTemp(1), Json_Text)
            Case "联系人电话"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("phone", varTemp(1), Json_Text)
            Case "联系人关系"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("relation", varTemp(1), Json_Text)
            Case "联系人地址"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("address", varTemp(1), Json_Text)
            End Select
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonTemp = "," & GetNodeString("contacts") & ":{" & strJsonTemp & "}"
            strJson = strJson & strJsonTemp
        End If
    End If
    
    If Not cllUpdCommunity Is Nothing Then
        '更新社区
        strJsonTemp = ""
        For i = 1 To cllUpdCommunity.Count
            varTemp = cllUpdCommunity(i)
            Select Case UCase(varTemp(0))
            Case "社区序号"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("num", Val(varTemp(1)), Json_num, True)
            Case "社区号码"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("code", varTemp(1), Json_Text)
            Case "社区操作类型"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("oper_type", Val(varTemp(1)), Json_num, True)
            End Select
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonTemp = "," & GetNodeString("community_info") & ":{" & strJsonTemp & "}"
            strJson = strJson & strJsonTemp
        End If
    End If
    
    If Not cllUpdVisit Is Nothing Then
        '更新社区
        strJsonTemp = ""
        For i = 1 To cllUpdVisit.Count
            varTemp = cllUpdVisit(i)
            Select Case UCase(varTemp(0))
            Case "就诊状态"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("statu", Val(varTemp(1)), Json_num, True)
            Case "就诊诊室"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("room", varTemp(1), Json_Text)
            Case "就诊时间"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("time", varTemp(1), Json_Text)
            End Select
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonTemp = "," & GetNodeString("visit_info") & ":{" & strJsonTemp & "}"
            strJson = strJson & strJsonTemp
        End If
    End If
    
    If strJson = "" Then
        strErrMsg = "未传入需要检查的必要条件，请检查!"
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_PatiSvr_NewPatiArchives
    'input
    '   pati_id N   1   病人id
    '   pati_name   C   1   姓名
    '   pati_sex    C   1   性别
    '   pati_age    C   1   年龄
    '   pati_birthdate  C   1   出生日期:yyyy-mm-dd hh24:mi:ss
    '   pati_idcard C   1   身份证号
    '   pati_type   C   1   病人类型
    '   outpatient_num  N   1   门诊号
    '   vcard_no    C   1   就诊卡号
    '   vcard_pwd   C   1   卡验证码
    '   fee_category    C   1   费别
    '   mdlpay_mode_name    C   1   医疗付款方式名称
    '   native_place    C   1   籍贯
    '   country_name    C   1   国籍
    '   nation_name C   1   民族
    '   mari_status C   1   婚姻状况
    '   edu_name    C   1   学历
    '   ocpt_name   C   1   职业
    '   pati_identity   C   1   身份
    '   emp_name    C   1   工作单位
    '   emp_postcode    C   1   单位邮编
    '   emp_phno    C   1   单位电话
    '   emp_bank_name   C   1   单位开户行
    '   emp_bank_accnum C   1   单位帐号
    '   ctt_unit_id N   1   合同单位id
    '   pat_home_addr   C   1   家庭地址
    '   pat_home_phno   C   1   家庭电话
    '   pat_home_postcode   C   1   家庭地址邮编
    '   region  C   1   区域
    '   pat_baddr   C   1   出生地点
    '   pat_hous_addr   C   1   户口地址
    '   pat_hous_postcode   C   1   户口地址邮编
    '   pat_grdn_name   C   1   监护人
    '   phone_number    C   1   手机号
    '   insurance_num   C   1   医保号
    '   iccard_no   C   1   Ic卡号
    '   create_time C   1   登记时间:yyyy-mm-dd hh24:mi:ss
    '   operator_name   C   1   操作员姓名
    '   idcard_sign N       身份证签约
    '   idcard_sign_pwd C       签约密码
    '   insurance_type  C   1   险类
    '   cert_no_other   C   1   其他证件
    '   contacts    C       更新联系人信息节点
    '       name    C   1   联系人姓名
    '       idcard  C   1   联系人身份证号
    '       phone   C   1   联系人电话
    '       relation    C   1   联系人关系
    '       address C       联系人地址
    '   community_info  C       社区信息节点
    '       num N   1   社区序号
    '       code    C   1   社区号码
    '       oper_type   N   1   社区操作类型
    '   visit_info  C       就诊信息节点
    '       statu   N       更新的就诊状态
    '       room    C       更新的就诊诊室
    '       time    C       就诊时间:yyyy-mm-dd hh24:mi:ss
    '   addr_list[] C       地址信息列表
    '       oper_fun    N   1   操作功能:1-新增,修改   2-删除
    '       type    C   1   地址类别
    '       state   C   1   地址_省
    '       city    C   1   地址_市
    '       county  C   1   地址_县
    '       township    C   1   地址_乡
    '       other   C   1   地址_其他
    '       code    C   1   区划代码
    '   ext_list[]  C       病人信息从项列表
    '       info_name   C   1   信息名
    '       upd_info_value  N   1   修改的信息值
    '   cert_list[]         证件列表(主要是当成绑卡处理)
    '       cert_name   C   1   证件名称
    '       cert_no C   1   证号号码
    '   allergic_drugs_list[]           病人过敏药物列表:有数据时，是先删除过敏药物插入的方式
    '       pat_algc_cadn_id    N   1   过敏药品ID
    '       pat_algc_cadn   C   1   过敏药物名称
    '       allergy_info    C   1   过每药物反应
    '       immune_list[]   C       病人免疫列表
    '       vaccinate_time  C   1   接种时间:yyyy-mm-dd hh24:mi:ss
    '       vaccinate_name  C   1   接种名称
    '   card_property_list[]    C       医疗卡属性列表
    '       cardtype_id N   1   医疗卡类别ID
    '       card_no C   1   卡号
    '       info_name   C   1   信息名
    '       info_value  N   1   信息值
    '
    strJson = Mid(strJson, 2)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_NewPatiArchives"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, True, , , , True) = False Then Exit Function
    '    output
    '        code    C   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：失败时返回具体的错误信息
    zl_PatiSvr_NewPatiArchives = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
