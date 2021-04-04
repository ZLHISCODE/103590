Attribute VB_Name = "mdlExseSvr"
Option Explicit

'*********************************************************************************************************************************************
'功能:费用相关服务接口处理
'函数:
' 一、公共部分
'    1.GetJsonNodeString:根据节点名或节点值获取Json串
'    2.GetNodeString:格式化节点名称
'    3.zlGetPubExseSvrObject:获取药品公共服务对象
'    4.zlGetRecipe_ID:获取处方ID
' 二、业务处理部分
'    1.zlHospitalization_Charge_Verfiy_isValied:住院记帐审核合法性检查
'    2.zlUpdateExcuteStatu:修改病人费用记录的执行状态
'    3.zlExcuteBillVerfiy:执行记帐单审核操作
'    4.SaveBill_NewRecipeBill-记帐插入合法性检查
'出参:
'编制:刘兴洪
'日期:2019*08*08 19:20:59
'*********************************************************************************************************************************************
Private mobjPubExseSvr As clsExpenceSvr '费用相关服务接口

Public Function zlGetPubExseSvrObject(ByRef objPubExseSvr As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品及卫材公共服务对象
    '出参:objPubExseSvr-返回药品及卫材公共服务对象
    '返回:获取返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-30 10:08:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If Not mobjPubExseSvr Is Nothing Then Set objPubExseSvr = mobjPubExseSvr: zlGetPubExseSvrObject = True: Exit Function
    
    
    Err = 0: On Error Resume Next
    Set mobjPubExseSvr = CreateObject("zlPublicExpense.clsExpenceSvr")
    If Err <> 0 Then
        MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense.clsExpenceSvr)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If mobjPubExseSvr.zlInitCommon(glngSys, glngModul, gcnOracle, gstrDBUser) = False Then
         MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)初始化失败，请与系统管理员联系！", vbExclamation, gstrSysName
         Exit Function
    End If
    Set objPubExseSvr = mobjPubExseSvr
    zlGetPubExseSvrObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Set mobjPubExseSvr = Nothing
End Function

Public Function zlReadInPatientDelBillData(ByVal int单据 As Integer, ByVal strNO As String, _
    Optional ByVal bln住院单位 As Boolean, Optional ByVal bln已结禁止销帐 As Boolean, _
    Optional ByVal bln禁止部分销帐 As Boolean, Optional ByVal str登记时间 As String, _
    Optional ByVal str收费类别s As String, Optional ByVal str排除收费类别s As String, _
    Optional ByRef rsBillData_out As ADODB.Recordset, Optional ByRef rsIncome_out As ADODB.Recordset, _
    Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long, _
    Optional ByVal str病人IDs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取需要销帐的单据数据
    '入参:
    '   int单据-1-收费单;2-记帐单;3-自动记帐单;
    '   strNo-单据号
    '   str登记时间-单据时间
    '   lngModule-模块号
    '   bln禁止部分销帐-
    '   str收费类别s-多个用逗号分离,如"5,6,7"
    '   str排除收费类别s-多个用逗号分离,如"5,6,7"
    '   str病人IDs-多病人单所销帐的病人IDs,如"1,2,3"
    '出参:rsBillData_out-单据数据
    '     rsIncome_out-收入汇总集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-16 20:04:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    If objPubExseSvr.zlReadInPatientDelBillData(int单据, strNO, bln住院单位, _
        bln已结禁止销帐, bln禁止部分销帐, str登记时间, str收费类别s, str排除收费类别s, _
        rsBillData_out, rsIncome_out, strFrmCaption, lngModule, str病人IDs) = False Then Exit Function
    
    zlReadInPatientDelBillData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlMzTOZyExceptionUpdate(ByVal strNos As String, Optional ByVal lngModule As Long, _
    Optional ByVal bln销账申请 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:当前单据是否门诊转住院异常及修正门诊转住院异常
    '入参:
    '     strNos-单据号，格式:A001,A002,...
    '     lngModule-模块号
    '     bln销账申请-是否销帐申请(true-销帐申请,false-销帐)
    '出参:strErrMsg_Out-返回错误信息(参数blnShowMsg=false时)
    '返回:不存在异常且成功修正异常返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objPubExseSvr As clsExpenceSvr, strErrMsg As String
    Dim strSubTable As String, varPara() As Variant

    On Error GoTo errHandle
    If zlGetVarBoundSQL(1, strNos, strSubTable, varPara, 0) = False Then Exit Function
    strSQL = _
        " Select Distinct a.No" & _
        " From 住院费用记录 A, (" & strSubTable & ") B" & _
        " Where NO = b.Column_Value And 记录性质 = 2 " & _
        "   And Not Exists(Select 1 From 病人费用异常记录 c Where c.费用id = a.id And c.产生环节 = 0 And c.同步标志 = 1) " & _
        "   And Exists(Select 1 From 病人费用异常记录 c Where c.费用id = a.id And c.产生环节 = 2)"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "检查转费同步标志", varPara)
    If rsTemp.EOF Then zlMzTOZyExceptionUpdate = True: Exit Function
    
    strNos = ""
    Do While Not rsTemp.EOF
        strNos = strNos & "," & Nvl(rsTemp!NO)
        rsTemp.MoveNext
    Loop
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    '存在时，试着同步关联
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    If objPubExseSvr.zlAdjustFeeData(strNos, True, False, strErrMsg) = False Then
        If strErrMsg = "" Then
            Call MsgBox("单据[" & strNos & "]为门诊费用转住院单据，目前处于数据异常状态，禁止进行销帐" & IIf(bln销账申请, "申请", "") & "！", vbInformation + vbOKOnly, gstrSysName)
        Else
            Call MsgBox("单据[" & strNos & "]为门诊费用转住院单据，目前处于数据异常状态，禁止进行销帐" & IIf(bln销账申请, "申请", "") & "，详细错误原因如下: " & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
        End If
        Exit Function
    End If
    zlMzTOZyExceptionUpdate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlExcuteBillVerfiy(ByVal strNO As String, ByVal str序号 As String, ByVal dt审核时间 As Date, _
    ByVal lngModule As Long, Optional ByVal strInsure As String, Optional ByRef blnAutoSendDrug As Boolean, _
    Optional ByVal bln是否多病人单 As Boolean, Optional ByVal lng病人ID As Long, _
    Optional ByVal byt费用来源 As Byte = 2, Optional ByRef blnStuffSync As Boolean, Optional ByRef blnDrugSync As Boolean, _
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
    '   byt费用来源-费用来源：1-门诊，2-住院
    '   blnAutoSendStuff-费用来源=2时有效，是否自动发药卫材，True=根据 Zl_住院记帐记录_Verify_Check 返回控制自动发料，false=不自动发料
    '出参:
    '   blnAutoSendDrug-True:已经自动发放药品;false-未自动发放成功药品
    '   blnStuffSync-卫材数据是否已同步
    '   blnDrugSync-药品数据是否已同步
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
    
    On Error GoTo ErrHandler
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    
    If byt费用来源 = 2 Then
        If strInsure = "0" Then strInsure = ""
        If objPubExseSvr.zlExcute_InBillVerfiy(strNO, str序号, dt审核时间, lngModule, _
            strInsure, blnAutoSendDrug, bln是否多病人单, lng病人ID, blnStuffSync, blnDrugSync, blnAutoSendStuff) = False Then Exit Function
    Else
        'strNos-单据信息, 格式：NO1:序号1,序号2,...|NO1:序号1,序号2,...|...
        If objPubExseSvr.zlVerfyBillingPriceBill(1, strNO & ":" & str序号, Format(dt审核时间, "yyyy-MM-dd HH:mm:ss")) = False Then Exit Function
        '药品已收费状态确认
        blnDrugSync = objPubExseSvr.zlDrugOutRecipeAffirm(strNO, 1, 2)
        '卫材已收费状态确认
        blnStuffSync = objPubExseSvr.zlStuffOutBillAffirm(strNO, 1, 2)
    End If
    zlExcuteBillVerfiy = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlUpdateExcuteStatu(ByVal cllData As Collection, _
    Optional ByVal int费用来源 As Integer = 2, Optional ByVal blnAutoCalc As Boolean = False, Optional ByVal str执行情况 As String, _
    Optional str操作员姓名 As String, Optional str操作时间 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改病人费用记录的执行状态
    '入参:cllData-更新的数据集(array(费用ids,执行状态,已发料)
    '     int费用来源-1-门诊;2-住院;0-不区分门诊或住院
    '     blnAutoCalc-是否自动计算执行状态
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-10 16:42:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
    
    On Error GoTo ErrHandler
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    zlUpdateExcuteStatu = objPubExseSvr.zlUpdateExcuteStatu(cllData, int费用来源, blnAutoCalc, str执行情况, str操作员姓名, str操作时间)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlUpdateExcuteStautsFromFeeIDs(ByVal strFeeIds As String, ByVal byt类别 As Byte, Optional byt费用来源 As Byte = 0, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用ID,自动修正执行状态
    '入参:strFeeIds-所涉及的费用ID
    '     byt类别-所涉及的类别:0-药品;1-卫材,2-所有
    '     byt费用来源-1-门诊;2-住院;0-不区分门诊或住院
    '返回:修正成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-24 11:49:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnExistStuff As Boolean, blnExistDrug As Boolean
    Dim cllUpdate As Collection, varData As Variant, rsRecipe As ADODB.Recordset '处方单号,药品ID,处方明细ID,已发数量,商品条码,内部条码
    Dim dbl已发数量 As Double, i As Long
    
    On Error GoTo errHandle
    blnExistDrug = byt类别 = 0 Or byt类别 = 2
    blnExistStuff = byt类别 = 1 Or byt类别 = 2
    If strFeeIds = "" Then zlUpdateExcuteStautsFromFeeIDs = True: Exit Function
    If mdlDrugAndStuffSvr.zlGetDrugAndStuff_ExcuteNum("", rsRecipe, blnExistDrug, blnExistStuff, , lngModule, strFeeIds) = False Then Exit Function
    
    varData = Split(strFeeIds, ",")
    Set cllUpdate = New Collection
    'cllData-更新的数据集(array(费用ids,执行状态,已发料)
    For i = 0 To UBound(varData)
        rsRecipe.Filter = "处方明细ID=" & Val(varData(i))
        dbl已发数量 = 0
        If Not rsRecipe.EOF Then dbl已发数量 = Val(Nvl(rsRecipe!已发数量))
        
        cllUpdate.Add Array(Val(varData(i)), 0, dbl已发数量)
    Next
    
    If zlUpdateExcuteStatu(cllUpdate, byt费用来源, True) = False Then Exit Function
    zlUpdateExcuteStautsFromFeeIDs = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlSaveBill_NewRecipeBill(ByVal strNO As String, ByVal cllRcpBillData As Collection, _
    ByVal strFrmCaption As String, ByVal bln划价 As Boolean, ByVal intBillType As Integer, _
    Optional ByVal lngModule As Long, Optional bln门诊 As Boolean, Optional ByRef blnSendMateria As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行记帐保存操作
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
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,[cllBillDetails(collect)])=cllBillists(_单据号)
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
    '  bln划价-是否划价
    '出参:blnSendMateria:记帐后自动发药
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-13 22:52:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHave As Boolean, cllDrawDeptFeeIds As Collection
    Dim objExpenceSvr As clsExpenceSvr
    
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then Exit Function
    
    blnHave = False
    If Not cllRcpBillData Is Nothing Then blnHave = cllRcpBillData.Count <> 0
    If Not blnHave Then
         MsgBox "不存在需要保存的记帐数据，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '更新病人主页信息：费用审核标志、住院状态
    If UpdatePatiPageInfo(cllRcpBillData, intBillType, lngModule) = False Then Exit Function
    
    '1.先进行数据合法性检查
    If objExpenceSvr.zlExcute_SaveRecipeBill_Check(cllRcpBillData, intBillType, _
        cllDrawDeptFeeIds, bln划价, bln门诊, lngModule) = False Then Exit Function
    
    '2.数据保存
    If objExpenceSvr.zlExcute_SaveRecipeBill(cllRcpBillData, strFrmCaption, intBillType, bln划价, _
        cllDrawDeptFeeIds, bln门诊, , blnSendMateria, lngModule) = False Then Exit Function
    
    zlSaveBill_NewRecipeBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiPageInfo(ByVal byt查询类型 As Byte, _
    ByVal str病人信息 As String, ByRef rsPatiPageInfo As ADODB.Recordset, _
    Optional ByVal bln包含婴儿信息 As Boolean, _
    Optional ByRef rsBabyInfo As ADODB.Recordset, Optional ByVal lngModule As Long, _
    Optional ByVal bln取最大主页ID As Boolean = True, _
    Optional ByVal str病人性质 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查询病人主页信息
    '入参:
    '   byt查询类型 查询类型:0-基本信息;1-基本信息的扩展;2-仅取主页
    '   str病人信息 病人信息,格式两种:一种是:病人id:主页ID,…;一种：病人id,…
    '   bln包含婴儿信息 是否包含婴儿信息
    '   bln取最大主页ID 是否取最后一次住院
    '   str病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人；多个逗号分隔，不传为所有
    '出参:
    '   rsPatiPageInfo 病人病案主页信息：病人id,主页id,姓名,性别,年龄,费别,病人性质,审核标志,住院状态,入院时间,出院时间,住院医师,
    '                                   医疗付款方式名称,医疗付款方式编码,当前病区id,当前病区名称,当前科室id,当前科室名称,
    '                                   险类,前床号,病人类型,[学历,职业,国籍,婚姻状况,编目日期,病人备注]
    '   rsBabyInfo 婴儿信息：病人id,主页id,序号,姓名,性别,出生时间
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim cllBabyInfo As Collection
    
    On Error GoTo ErrHandler
    If str病人信息 = "" Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    
    zlGetPatiPageInfo = objService.zlCIsSvr_GetPatiPageInfo(byt查询类型, str病人信息, _
        rsPatiPageInfo, bln包含婴儿信息, cllBabyInfo, lngModule, bln取最大主页ID, str病人性质)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function UpdatePatiPageInfo(cllRcpBillData As Collection, ByVal intBillType As Integer, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新病人主页信息：费用审核标志、住院状态
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
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,[cllBillDetails(collect)])=cllBillists(_单据号)
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
    '出参:
    '返回:更新成功返回True,失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiInfo As Collection, cllBillLists As Collection
    Dim cllBillDetails As Collection, cllBillDetail As Collection
    Dim rsPatiPageInfo As ADODB.Recordset
    Dim p As Long, i As Long, str病人IDs As String, str病人信息 As String

    On Error GoTo ErrHandler
    '收集病人信息
    str病人信息 = ""
    If intBillType = 3 Then '记帐表
        str病人IDs = ""
        Set cllBillLists = cllRcpBillData("_cllBillLists")
        For p = 1 To cllBillLists.Count
            Set cllBillDetails = cllBillLists(p)("_cllBillDetails")
            For i = 1 To cllBillDetails.Count
                Set cllBillDetail = cllBillDetails(i)
                If InStr("," & str病人IDs & ",", "," & cllBillDetail("病人ID") & ",") = 0 Then
                    str病人信息 = str病人信息 & "," & cllBillDetail("病人ID") & ":" & cllBillDetail("主页ID")
                    str病人IDs = str病人IDs & "," & cllBillDetail("病人ID")
                End If
            Next
        Next
    Else
        Set cllPatiInfo = cllRcpBillData("_patiinfor")
        str病人信息 = str病人信息 & "," & cllPatiInfo("病人ID") & ":" & cllPatiInfo("主页ID")
    End If
    If str病人信息 <> "" Then str病人信息 = Mid(str病人信息, 2)

    '获取病人主页信息
    If zlGetPatiPageInfo(0, str病人信息, rsPatiPageInfo, False, , lngModule) = False Then Exit Function

    '更新病人主页信息
    If intBillType = 3 Then '记帐表
        Set cllBillLists = cllRcpBillData("_cllBillLists")
        For p = 1 To cllBillLists.Count
            Set cllBillDetails = cllBillLists(p)("_cllBillDetails")
            For i = 1 To cllBillDetails.Count
                Set cllBillDetail = cllBillDetails(i)
                rsPatiPageInfo.Filter = "病人ID=" & cllBillDetail("病人ID")
                If rsPatiPageInfo.EOF Then
                    MsgBox "获取病人【" & cllBillDetail("姓名") & "】的病人信息失败！", vbInformation, gstrSysName
                    Exit Function
                End If
                If CollectionExitsValue(cllBillDetail, "费用审核标志") Then cllBillDetail.Remove "费用审核标志"
                cllBillDetail.Add Nvl(rsPatiPageInfo!审核标志), "费用审核标志"
                If CollectionExitsValue(cllBillDetail, "住院状态") Then cllBillDetail.Remove "住院状态"
                cllBillDetail.Add Nvl(rsPatiPageInfo!住院状态), "住院状态"
            Next
        Next
    Else
        Set cllPatiInfo = cllRcpBillData("_patiinfor")
        If rsPatiPageInfo.EOF Then
            MsgBox "获取病人【" & cllPatiInfo("姓名") & "】的病人信息失败！", vbInformation, gstrSysName
            Exit Function
        End If
        If CollectionExitsValue(cllPatiInfo, "费用审核标志") Then cllPatiInfo.Remove "费用审核标志"
        cllPatiInfo.Add Nvl(rsPatiPageInfo!审核标志), "费用审核标志"
        If CollectionExitsValue(cllPatiInfo, "住院状态") Then cllPatiInfo.Remove "住院状态"
        cllPatiInfo.Add Nvl(rsPatiPageInfo!住院状态), "住院状态"
    End If

    UpdatePatiPageInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Public Function zlExcute_DelRecipeBill(ByVal cllBillLists As Collection, _
    ByVal strFrmCaption As String, ByVal intBillType As Integer, _
    Optional ByVal bln划价 As Boolean, Optional ByVal lngModule As Long, Optional bln门诊 As Boolean, _
    Optional ByVal cllPro As Collection, Optional ByVal cllPatients As Collection) As Boolean
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
    Dim blnHave As Boolean
    Dim objExpenceSvr As clsExpenceSvr
    
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then Exit Function
    
    blnHave = False
    If Not cllBillLists Is Nothing Then blnHave = cllBillLists.Count <> 0
    If Not blnHave Then
         MsgBox "不存在需要销帐的数据，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '1.先进行销帐数据的有效性进行检查
    If objExpenceSvr.zlExcute_DelRecipeBill_Check(cllBillLists, intBillType, lngModule, _
        bln划价, bln门诊, cllPatients) = False Then Exit Function
    
    '2.再进行销帐处理
    If objExpenceSvr.zlExcute_DelRecipeBill(cllBillLists, strFrmCaption, _
        intBillType, bln划价, lngModule, bln门诊, cllPro, cllPatients) = False Then Exit Function
    zlExcute_DelRecipeBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl病人费用销帐_Insert(ByVal cllApplyDatas As Collection, ByRef str申请费用ids_Out As String, _
    Optional ByVal strFrmCaption As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人费用销帐申请
    '入参:
    '   cllApplyDatas-销账申请数据集(RowData=collect:(费用ID,单据号,收费细目ID,申请科室ID,审核科室ID,申请数量,申请人,申请时间,申请类别,收费类别,销帐原因))
    '出参:
    '   str申请费用ids_Out-本次申请所涉及的费用ID(主要是后继审核要用)
    '返回:申请返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    zl病人费用销帐_Insert = objPubExseSvr.zl病人费用销帐_Insert(cllApplyDatas, str申请费用ids_Out, strFrmCaption, lngModule)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_病人费用销帐_Audit_Check(ByVal cllAuditDatas As Collection, _
    Optional ByVal strFrmCaption As String, Optional lngModule As Long, _
    Optional ByRef rsRecipe_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人费用销帐申请
    '入参:
    '   cllAuditDatas-销账审核数据集(RowData=collect:(费用ID,单据号,收费类别,卫材是否自动退料,申请时间,申请类别))
    '          成员值说明:1.卫生材料是否自动退料:1-自动退料;0-不自动退料
    '                     2.申请类别:0-未发药(料);1-已发药(料);其他为0
    '出参:
    '   rsRecipe_Out-药品及卫生材料的已执行数量的数据集
    '返回:审核成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    Zl_病人费用销帐_Audit_Check = objPubExseSvr.Zl_病人费用销帐_Audit_Check(cllAuditDatas, strFrmCaption, lngModule, rsRecipe_Out)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_病人费用销帐_Cancel_Check(ByVal cllAuditDatas As Collection, _
    Optional ByVal strFrmCaption As String, Optional lngModule As Long, _
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
    '---------------------------------------------------------------------------------------------------------------------------------------------    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    Zl_病人费用销帐_Cancel_Check = objPubExseSvr.Zl_病人费用销帐_Cancel_Check(cllAuditDatas, strFrmCaption, lngModule, rsSendDatas)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetDeptName(ByVal lngDeptID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取部门名称
    '入参:lngDeptID-部门ID(或病区ID)
    '返回:返回取部门名称
    '编制:刘兴洪
    '日期:2015-07-15 17:52:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select 名称 From 部门表 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取部门名称", lngDeptID)
    If rsTemp.EOF = False Then zlGetDeptName = Nvl(rsTemp!名称)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsExistInfusion_FromNo(ByVal strNO As String, str序号s As String, _
    Optional blnIsExist_Out As Boolean, Optional ByVal int记录性质 As Integer = 2, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查药品是否进入输入液配药中心
    '入参:strNo-单据号
    '     str序号s-可以为空,为空时，表示所有数据
    '出参:blnIsExist_Out-是否存在，存在返回true,否则返回False
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-22 16:30:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str费用IDs As String
    Dim strWhere As String
    
    On Error GoTo errHandle
    blnIsExist_Out = False
    If str序号s <> "" Then strWhere = strWhere & " And instr([3],','||序号||',')>0 "
    strSQL = "Select ID" & _
            " From 住院费用记录" & _
            " Where NO=[1] and 记录性质 =[2]  And 记录状态 in (0,1,3)" & _
            "       And 收费类别 in ('5','6','7') " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取费用ID", strNO, int记录性质, "," & str序号s & ",")
    If rsTemp.EOF Then zlIsExistInfusion_FromNo = True: Exit Function
    
    Do While Not rsTemp.EOF
        str费用IDs = str费用IDs & "," & rsTemp!ID
        rsTemp.MoveNext
    Loop
    str费用IDs = Mid(str费用IDs, 2)
    If mdlDrugAndStuffSvr.zlPivasSvr_Isexsitinfusion(str费用IDs, blnIsExist_Out, lngModule) = False Then Exit Function
    
    zlIsExistInfusion_FromNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlFeeExecute_Check(ByVal strNO As String, ByVal str序号 As String, ByVal int费用来源 As Integer, ByVal int单据性质 As Integer, _
    Optional ByRef strSendStuffFeeIDs_Out As String, _
    Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    '功能: 执行登记合法性检查
    '入参:
    '   strNO-单据号
    '   str序号-序号
    '   int费用来源-1-门诊;2-住院
    '   int单据性质-1-收费;2-记帐;3-自动记帐
    '出参:
    '   strSendStuffFeeIDs_Out-本次执行所涉及自动发放卫材所对应的费用Ids,多个用逗号分离
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    
    '执行登记前检查
    If objPubExseSvr.zlFeeExecute_Check(strNO, str序号, int费用来源, int单据性质, _
        strSendStuffFeeIDs_Out, strFrmCaption, lngModule, 1) = False Then Exit Function
    zlFeeExecute_Check = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlFeeUnExecute_Check(ByVal strNO As String, ByVal str序号 As String, ByVal int费用来源 As Integer, _
    ByVal int单据性质 As Integer, ByRef strStuffFeeIDs_Out As String, _
    ByVal strFrmCaption As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    '功能: 取消执行登记前检查
    '入参:
    '   strNO-单据号
    '   str序号-序号
    '   int费用来源-1-门诊;2-住院
    '   int单据性质-1-收费;2-记帐;3-自动记帐
    '出参:
    '   strStuffFeeIDs_Out-本次执行所涉及卫材所对应的费用Ids,多个用逗号分离
    '返回:
    '---------------------------------------------------------------------------------------
    Dim objPubExseSvr As clsExpenceSvr
     
    On Error GoTo errHandle
    If zlGetPubExseSvrObject(objPubExseSvr) = False Then Exit Function
    
    '执行登记前检查
    If objPubExseSvr.zlFeeUnExecute_Check(strNO, str序号, int费用来源, int单据性质, strStuffFeeIDs_Out, _
         strFrmCaption, lngModule) = False Then Exit Function
    zlFeeUnExecute_Check = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetMaxBedLen(Optional ByVal lng部门ID As Long, Optional ByVal bln科室 As Boolean) As Integer
'功能：获取指定部门的床位号的最大长度
'参数：lng部门ID=病区ID或科室ID,为0表示所有病区或科室
    Dim objService As zlPublicExpense.clsService
    Dim lngBedNoMaxLen As Long, bln按病区查询 As Boolean
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    
    bln按病区查询 = (Not bln科室 Or lng部门ID = 0)
    If objService.ZlCissvr_GetMaxBedLen(lngBedNoMaxLen, bln按病区查询, lng部门ID, lng部门ID) = False Then Exit Function
    
    zlGetMaxBedLen = lngBedNoMaxLen
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCheck医生下达出院医嘱(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医生是否下达了出院医嘱
    '入参:
    '出参:
    '   blnExistOutAdvice=是否存在出院医嘱
    '   lngOutAdviceId=已经开了出院医嘱的，返回医嘱的ID
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim blnExistOutAdvice  As Boolean
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then
        zlCheck医生下达出院医嘱 = True: Exit Function '保守策略
    End If
    
    If objService.ZlCissvr_ExistOutAdvice(lng病人ID, lng主页ID, blnExistOutAdvice) = False Then
        zlCheck医生下达出院医嘱 = True: Exit Function '保守策略
    End If
    
    zlCheck医生下达出院医嘱 = blnExistOutAdvice
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetPatiInfoByPage(objPati As clsPatientInfo, _
    Optional ByVal lng主页ID As Long, Optional ByVal bln包含婴儿信息 As Boolean, _
    Optional ByRef rsBabyInfo As ADODB.Recordset, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从病案主页中获取病人信息
    '入参:
    '   objPati-已有病人信息
    '   lng主页ID-主页ID，为0时，取最后一次住院的
    '   bln包含婴儿信息 是否包含婴儿信息
    '出参:
    '   objPati-返回病人信息对象
    '   rsBabyInfo 婴儿信息：病人id,主页id,序号,姓名,性别,出生时间
    '返回:成功返回True，否则返回False
    '说明:如果传入 objPati 不为Nothing，则进行信息合并
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objService As zlPublicExpense.clsService
    Dim cllBabyInfo As Collection
    
    On Error GoTo errHandle
    If objPati Is Nothing Then Exit Function
    If objPati.病人ID = 0 Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    
    If lng主页ID = 0 Then lng主页ID = objPati.主页ID '病人信息中的主页ID为最后一次主页ID
    If objService.zlCIsSvr_GetPatiPageInfo(1, objPati.病人ID & ":" & lng主页ID, _
        rsTemp, bln包含婴儿信息, cllBabyInfo, lngModule) = False Then Exit Function
    If rsTemp Is Nothing Then zlGetPatiInfoByPage = True: Exit Function
    If rsTemp.EOF Then zlGetPatiInfoByPage = True: Exit Function
    
    If objPati Is Nothing Then Set objPati = New clsPatientInfo
    If bln包含婴儿信息 And Not cllBabyInfo Is Nothing Then
        If cllBabyInfo.Count > 0 Then Set rsBabyInfo = cllBabyInfo("_" & objPati.病人ID)
    End If
    
    With objPati
        .主页ID = Nvl(rsTemp!主页ID)
        .姓名 = Nvl(rsTemp!姓名)
        .性别 = Nvl(rsTemp!性别)
        .年龄 = Nvl(rsTemp!年龄)
        .费别 = Nvl(rsTemp!费别)
        .医疗付款方式 = Nvl(rsTemp!医疗付款方式名称)
        .医疗付款方式编码 = Nvl(rsTemp!医疗付款方式编码)
        .险类 = Val(Nvl(rsTemp!险类))
        .险类名称 = GetInsureName(Val(Nvl(rsTemp!险类)))
        .病人类型 = Nvl(rsTemp!病人类型)
        .当前病区ID = Val(Nvl(rsTemp!当前病区ID))
        .当前病区名称 = Nvl(rsTemp!当前病区名称)
        .当前科室ID = Val(Nvl(rsTemp!当前科室ID))
        .当前科室名称 = Nvl(rsTemp!当前科室名称)
        .床号 = Nvl(rsTemp!当前床号)
        .住院号 = Nvl(rsTemp!住院号)
        .病人性质 = Val(Nvl(rsTemp!病人性质))
        .入院日期 = Nvl(rsTemp!入院时间)
        .出院日期 = Nvl(rsTemp!出院时间)
        .住院医师 = Nvl(rsTemp!住院医师)
        .病人备注 = Nvl(rsTemp!病人备注)
        .住院状态 = Val(Nvl(rsTemp!住院状态))
        .审核标志 = Val(Nvl(rsTemp!审核标志))
        .编目日期 = Nvl(rsTemp!编目日期)
        .医保号 = Nvl(rsTemp!医保号)
    End With
    zlGetPatiInfoByPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiInfo(ByVal lng病人ID As Long, _
    Optional ByVal lng主页ID As Long, Optional ByVal lngModule As Long, _
    Optional ByVal bln包含婴儿信息 As Boolean, _
    Optional ByRef rsBabyInfo As ADODB.Recordset, _
    Optional ByVal blnNotShowErrMsg As Boolean) As clsPatientInfo
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息，先从病人信息中获取，再从病案主页中获取进行合并
    '入参:
    '   objPati-已有病人信息
    '   lng主页ID-主页ID，为0时，取最后一次住院的；为-1时取门诊病人信息
    '   bln包含婴儿信息 是否包含婴儿信息
    '出参:
    '   objPati-返回病人信息对象
    '   rsBabyInfo 婴儿信息：病人id,主页id,序号,姓名,性别,出生时间
    '返回:成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPati As clsPatientInfo
    Dim lngTmp As Long
    
    On Error GoTo errHandle
    '读取病人信息
    If gobjSquare.objOneCardComLib.zlGetPatiInforFromPatiID(lng病人ID, objPati, , , , , , , , , , , blnNotShowErrMsg) = False Then Exit Function
    If objPati Is Nothing Then Exit Function
    lngTmp = objPati.主页ID
    '2.读取病案主页
    If lng主页ID = 0 Then lng主页ID = objPati.主页ID
    If lng主页ID > 0 Then
        If zlGetPatiInfoByPage(objPati, lng主页ID, bln包含婴儿信息, rsBabyInfo, lngModule) = False Then Exit Function
    End If
    If lngTmp <> objPati.主页ID Then objPati.在院 = False
    Set zlGetPatiInfo = objPati
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetMultiPatiInfo(ByVal str病人信息 As String, _
    Optional ByRef cllBabyInfo As Collection, Optional ByVal lngModule As Long) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取多个病人信息，先从病人信息中获取，再从病案主页中获取进行合并
    '入参:
    '   str病人信息=病人ID和主页信息，格式：病人ID1:主页ID,病人ID2:主页ID,...；其中,主页ID=0时，取最后一次住院的；为-1时取门诊病人信息
    '出参:
    '   cllBabyInfo=婴儿信息,成员:ADODB.Recordset=cllBabyInfo(_病人ID)，成员字段：病人id,主页id,序号,姓名,性别,出生时间
    '返回:病人信息集，成员：clsPatinetInfo=cllPatis(_病人ID)
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection, varData As Variant, i As Long
    Dim str病人IDs As String, lng病人ID As Long, lng主页ID As Long
    Dim rsTemp As ADODB.Recordset, objPati As clsPatientInfo
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    
    varData = Split(str病人信息, ",")
    For i = 0 To UBound(varData)
        lng病人ID = Split(varData(i) & ":", ":")(0)
        str病人IDs = str病人IDs & "," & lng病人ID
    Next
    If str病人IDs = "" Then Exit Function
    
    str病人IDs = Mid(str病人IDs, 2)
    '读取病人信息
    If gobjSquare.objOneCardComLib.zlGetMultiPatiInforFromPatiID(str病人IDs, cllPati) = False Then Exit Function
    If cllPati Is Nothing Then Exit Function
    If cllPati.Count = 0 Then Exit Function
    
    Set zlGetMultiPatiInfo = cllPati
    '2.读取病案主页
    str病人信息 = ""
    For i = 0 To UBound(varData)
        lng病人ID = Split(varData(i) & ":", ":")(0)
        lng主页ID = Val(Split(varData(i) & ":", ":")(1))
        If lng主页ID = 0 Then '主页ID=0时，取最后一次住院的
            Set objPati = cllPati("_" & lng病人ID)
            lng主页ID = objPati.主页ID
        End If
        If lng主页ID > 0 Then '主页ID=-1时，仅取门诊病人信息
            str病人信息 = str病人信息 & "," & lng病人ID & ":" & lng主页ID
        End If
    Next
    If str病人信息 = "" Then Exit Function
    
    str病人信息 = Mid(str病人信息, 2)
    If objService.zlCIsSvr_GetPatiPageInfo(1, str病人信息, rsTemp, True, cllBabyInfo, lngModule) = False Then Exit Function
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.EOF Then Exit Function
    
    For i = 0 To UBound(varData)
        lng病人ID = Split(varData(i) & ":", ":")(0)
        rsTemp.Filter = "病人ID=" & lng病人ID
        If Not rsTemp.EOF Then
            Set objPati = cllPati("_" & lng病人ID)
            With objPati
                .主页ID = Nvl(rsTemp!主页ID)
                .姓名 = Nvl(rsTemp!姓名)
                .性别 = Nvl(rsTemp!性别)
                .年龄 = Nvl(rsTemp!年龄)
                .费别 = Nvl(rsTemp!费别)
                .医疗付款方式 = Nvl(rsTemp!医疗付款方式名称)
                .医疗付款方式编码 = Nvl(rsTemp!医疗付款方式编码)
                .险类 = Val(Nvl(rsTemp!险类))
                .险类名称 = GetInsureName(Val(Nvl(rsTemp!险类)))
                .病人类型 = Nvl(rsTemp!病人类型)
                .当前病区ID = Val(Nvl(rsTemp!当前病区ID))
                .当前病区名称 = Nvl(rsTemp!当前病区名称)
                .当前科室ID = Val(Nvl(rsTemp!当前科室ID))
                .当前科室名称 = Nvl(rsTemp!当前科室名称)
                .床号 = Nvl(rsTemp!当前床号)
                .住院号 = Nvl(rsTemp!住院号)
                .病人性质 = Val(Nvl(rsTemp!病人性质))
                .入院日期 = Nvl(rsTemp!入院时间)
                .出院日期 = Nvl(rsTemp!出院时间)
                .住院医师 = Nvl(rsTemp!住院医师)
                .病人备注 = Nvl(rsTemp!病人备注)
                .住院状态 = Val(Nvl(rsTemp!住院状态))
                .审核标志 = Val(Nvl(rsTemp!审核标志))
                .编目日期 = Nvl(rsTemp!编目日期)
                .医保号 = Nvl(rsTemp!医保号)
            End With
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlGetPatiPageInfByRange(ByVal cllFilter As Collection, _
    ByRef rsPatiPageInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long, Optional ByVal bln包含婴儿信息 As Boolean, _
    Optional ByRef cllBabyInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查询病人主页信息
    '入参:
    '   cllFilter 查询条件集:成员(Array(Key,Value),Array(Key,Value),,...)
    '       Key:病区IDS,科室IDS,病人IDS,主页IDS,入院开始时间,入院结束时间,出院开始时间,出院结束时间,
    '           费别,住院状态,病人性质,姓名,站点编号,查询转科病人,最后一次住院,险类,病区站点编号
    '       住院状态:0-在院病人;1-出院病人;2-在院或出院
    '       病人性质：多个用逗号分0-普通住院病人,1-门诊留观病人,2-住院留观病人，NULL-表示不区分
    '       姓名:可以代%分号表表按姓名匹配
    '       已出院天数，住院状态为1和2时有效
    '       站点编号:科室对应的站点编号
    '       险类:>0:指定险类医保病人,0:医保和普通病人,-1:普通病人,-2:医保病人
    '   bln包含婴儿信息 是否包含婴儿信息
    '出参:
    '   rsPatiPageInfo 病人病案主页信息：病人ID,主页ID,姓名,性别,年龄,住院号,床号,险类,费别,病人类型,医保号,
    '                                   入院时间,出院时间,住院状态,病人性质,当前病区ID,当前病区名称,当前科室ID,当前科室名称,
    '                                   医疗付款方式名称,医疗付款方式编码,住院医师,病人备注,编目日期,护理等级,
    '                                   数据转出,审核标志,审核人,预出院时间,上次催款金额
    '       住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)
    '       病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
    '       数据转出:0-未转出，1-已转出
    '       审核标志:0或空-未审核,1-已审核或开始审核;2-完成审核
    '   cllBabyInfo 婴儿信息,成员：ADODB.Recordset=cllBabyInfo(_病人ID_主页ID)，字段：病人id,主页id,序号,姓名,性别,出生时间
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    ZlGetPatiPageInfByRange = objService.ZlCissvr_GetPatiPageInfByRange(cllFilter, rsPatiPageInfo, lngModule, _
        bln包含婴儿信息, cllBabyInfo)
End Function

Public Function ZlGetBabyData(ByVal lng病人ID As Long, _
    ByVal lng主页ID As Long, ByRef rsBabyInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查询病人主页信息
    '入参:
    '出参:
    '   cllBabyInfo 婴儿信息,字段：序号,姓名
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    If lng病人ID = 0 Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    ZlGetBabyData = objService.ZlCissvr_GetBabyData(lng病人ID, lng主页ID, rsBabyInfo, lngModule)
End Function


Public Function ZLGetAdviceIDs(ByVal str医嘱ID As String) As String
    '读取一组医嘱包含的医嘱记录ID串
    '入参:
    '   str医嘱IDs 医嘱ID,多个英文逗号分隔
    Dim objService As zlPublicExpense.clsService
    Dim rsAdviceData As ADODB.Recordset, str医嘱IDs As String
    
    If str医嘱ID = "" Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAllGroupAdviceIDs(str医嘱ID, rsAdviceData) = False Then Exit Function
    If rsAdviceData.EOF Then Exit Function
    
    str医嘱IDs = ""
    Do While Not rsAdviceData.EOF
        str医嘱IDs = str医嘱IDs & "," & Nvl(rsAdviceData!医嘱ID)
        rsAdviceData.MoveNext
    Loop
    ZLGetAdviceIDs = Mid(str医嘱IDs, 2)
End Function

Public Function ZlGetPatiIdFromPatiPage(ByRef lng病人ID As Long, Optional ByVal lng病区ID As Long, _
    Optional ByVal str住院号 As String, Optional ByVal str床号 As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据床号、住院号获取病人ID及主页ID
    '入参:
    '   str住院号、lng病区id与str床号-二者至少传一个
    '出参:
    '   str病人信息:病人ID:主页ID
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    ZlGetPatiIdFromPatiPage = objService.zlCisSvr_GetPatiID(lng病人ID, lng病区ID, str住院号, str床号, lngModule)
End Function

Public Function ZlGetInDeptInfor(ByVal byt执行功能 As Byte, ByVal bln在院病人 As Boolean, _
    Optional ByVal byt查找方式 As Byte, Optional ByVal bln所有病区 As Boolean, Optional ByVal str病区IDs As String, _
    Optional ByVal str服务对象 As String, Optional ByVal byt病人来源 As Byte = 2, Optional ByVal lngModule As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取所有有病人的住院科室
    '入参:
    '   byt执行功能=0-获取所有有病人的住院科室 1-通过科室id/病区id查找所有病人的入院科室或者病区 2-加载站点
    '   bln在院病人=是否仅取存在在院病人的科室\病区
    '   byt查找方式=0-按科室查找 1-按病区查找
    '   bln所有病区=是否所有病区
    '   str病区ids=非所有病区时有效
    '   str服务对象=科室服务对象，多个逗号分隔,如:1,2,3（1-门诊,2-住院,3-门诊和住院）
    '出参:
    '   rsDept 科室信息,字段：ID,编码,名称,简码
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim rsDept As ADODB.Recordset
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetInDeptInfor(byt执行功能, byt病人来源, gstrNodeNo, _
        bln在院病人, rsDept, byt查找方式, bln所有病区, str病区IDs, str服务对象, lngModule) = False Then Exit Function
    Set ZlGetInDeptInfor = rsDept
End Function

Public Function zlCheckPatiIsDeath(ByVal lng病人ID As Long, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id检查病人是否已经死亡
    '入参:
    '出参:
    '返回:已死亡返回True,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDeath As Boolean
    Dim objService As zlPublicExpense.clsService
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_PatiIsDead(lng病人ID, blnDeath, lngModule) = False Then Exit Function
    zlCheckPatiIsDeath = blnDeath
End Function

Public Function ZlGetPatiChangeStopInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng病区ID As Long, ByVal lng科室ID As Long, ByVal str终止原因 As String, _
    ByRef str终止原因_Out As String, ByRef str终止时间_Out As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id、主页id、科室及病区id，获取病人变动的终止信息(终止时间、终止原因等）
    '入参:
    '   str终止原因=终止原因:多个用逗号分离,如:3,15,10,1
    '出参:
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetPatiChangeStopInfo(lng病人ID, lng主页ID, _
        lng病区ID, lng科室ID, str终止原因, str终止原因_Out, str终止时间_Out, lngModule) = False Then Exit Function
    ZlGetPatiChangeStopInfo = True
End Function

Public Function zlCheckPatiIsMemo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id及主页id检查是否存在备注信息
    '入参
    '出参:
    '   blnIsExist_Out-存在的，返回true,否则返回False
    '返回:成功返回true,否则返回False
    '问题:34763
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim blnIsExis As Boolean
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlCissvr_PatiExistMemo(lng病人ID, lng主页ID, blnIsExis, lngModule) = False Then Exit Function
    zlCheckPatiIsMemo = blnIsExis
End Function

Public Function zlGetPatiPageExtendInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str信息名 As String, Optional ByVal lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id及主页id获取病案主页从表信息
    '入参
    '   str信息名=信息名：多个用逗号
    '出参:
    '返回:返回信息值
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim cllValue As Collection '信息值集合，成员:Array(信息名,信息值)=cllValue(信息名)
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlCissvr_GetPatiPageExtendInfo(lng病人ID, lng主页ID, str信息名, cllValue, lngModule) = False Then Exit Function
    If cllValue Is Nothing Then Exit Function
    If cllValue.Count = 0 Then Exit Function
    zlGetPatiPageExtendInfo = cllValue(1)(str信息名)
End Function

Public Function ZlGetAdviceInfoByPati(ByVal lng病人ID As Long, ByVal ln主页ID As Long, _
    ByRef rsAdviceData As ADODB.Recordset, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱ID获取医嘱信息
    '入参:
    '出参:
    '   rsAdviceData 医嘱信息：医嘱ID,医嘱内容,挂号单号,诊疗项目ID,诊疗类别,操作类型,皮试结果,医嘱期效,
    '               【完整数据还包含:相关ID,序号,婴儿序号,医嘱状态,开嘱医生,医生嘱托,开嘱时间,开嘱科室ID,
    '                 毒理分类,紧急标志,总量,天数,单量,计算单位,执行频次,用法,执行时间方案,
    '                 开始执行时间,执行终止时间,执行科室ID,执行科室名称,执行性质,上次执行时间,执行标记,
    '                 校对护士,校对时间,停嘱医生,停嘱时间,停嘱护士,确认停嘱时间,
    '                 费用审核,审查结果,审核状态,试管编码,屏蔽打印,是否签名,报告ID,查阅状态】
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim cllFilter As Collection
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    Set cllFilter = New Collection
    cllFilter.Add Array("病人ID", lng病人ID)
    cllFilter.Add Array("主页ID", ln主页ID)
    If objService.ZlCissvr_GetAdviceInfo(cllFilter, rsAdviceData, 1, lngModule) = False Then Exit Function
    ZlGetAdviceInfoByPati = True
End Function

Public Function ZLGetAdviceSendInfo(ByVal byt查询类型 As Byte, ByVal strValues As String, _
    ByRef rsAdviceSendData As ADODB.Recordset, Optional ByVal bln包含相关医嘱 As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱发送信息
    '入参:
    '   byt查询类型=0-按医嘱id串（advice_ids）查询;1-按医嘱ID+医嘱发送号查询;2-按医嘱ID+记录性质+NO查询;3-仅按医嘱发送号查询
    '   strValues=查询值，具体如下:
    '               byt查询类型=0:医嘱ID串,格式：医嘱ID,医嘱ID,...
    '               byt查询类型=1:单据信息,格式：医嘱ID:NO:记录性质,医嘱ID:NO:记录性质,...
    '               byt查询类型=2:医嘱发送信息,格式:医嘱ID:发送号,医嘱ID:发送号,...
    '               byt查询类型=3:医嘱发送号,格式:发送号,发送号,...
    '   bln包含相关医嘱=是否包含相关医嘱ID
    '出参:
    '   rsAdviceSendData 医嘱发送信息：医嘱ID,发送号,挂号单号,病人ID,主页ID,病人科室ID,开嘱科室ID,
    '                                 病人来源,诊疗类别,计价性质,相关id,门诊记帐,医嘱内容,样本条码,
    '                                 No,记录性质,首次时间,末次时间,执行状态,发送时间,医嘱期效
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAdviceSendInfo(byt查询类型, strValues, rsAdviceSendData, bln包含相关医嘱, lngModule) = False Then Exit Function
    ZLGetAdviceSendInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlStopAutoAccount(ByVal lng病人ID As Long, ByVal str主页IDS As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:停止自动记帐
    '入参:
    '   str主页IDs=主页ID,多个逗号分隔
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCisSvr_UpddteAutoAccountSign(lng病人ID, str主页IDS, True, lngModule) = False Then Exit Function
    ZlStopAutoAccount = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlRestoreAutoAccount(ByVal strNO As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:恢复自动记帐
    '入参:
    '   strNO=结帐单号
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng病人ID As Long, str主页IDS As String
    
    On Error GoTo ErrHandler
    strSQL = "Select 病人id, 住院次数" & _
            " From 病人结帐记录" & _
            " Where NO = [1] And 记录状态 = 3 And 结帐类型 = 2 And 中途结帐 = 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询结帐住院次数", strNO)
    If rsTemp.EOF Then Exit Function
    
    lng病人ID = Val(Nvl(rsTemp!病人ID))
    str主页IDS = Nvl(rsTemp!住院次数)
    If str主页IDS Then ZlRestoreAutoAccount = True: Exit Function
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCisSvr_UpddteAutoAccountSign(lng病人ID, str主页IDS, False, lngModule) = False Then Exit Function
    ZlRestoreAutoAccount = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetMedicalGroupID(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng开单科室ID As Long, ByVal str开单人 As String, _
    ByVal dt发生时间 As Date, Optional ByVal lngModule As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据条件获取对应的医疗小组ID
    '入参:
    '   dt发生时间=费用发生时间
    '出参:
    '返回:获取到的医疗小组ID
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim lng组id As Long
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetMedicalGroupID(lng病人ID, lng主页ID, _
        lng开单科室ID, str开单人, dt发生时间, lng组id, lngModule) = False Then Exit Function
        
    ZlGetMedicalGroupID = lng组id
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCheckNotExcuteItem(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal byt婴儿序号 As Byte, ByVal byt费用来源 As String, _
    ByRef strNotExcuteInfo As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人信息获取医技未执行的项目
    '入参:
    '   byt婴儿序号=婴儿序号:-1表示不区分;0-母亲的;>0具体婴儿费用
    '   byt费用来源=1-门诊;2-住院;4-体检
    '出参:
    '   strNotExcuteInfo=未执行的项目信息
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim lng组id As Long
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_CheckNotExcuteItem(lng病人ID, lng主页ID, -1, byt费用来源, strNotExcuteInfo, lngModule) = False Then Exit Function
        
    zlCheckNotExcuteItem = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckNotExcuteItemValied(ByVal str姓名 As String, ByVal lng病人ID As Long, _
    ByVal lng主页ID As Long, ByVal int门诊标志 As Integer, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查门诊费用未执行项目是否合法
    '入参:
    '   int门诊标志-1-门诊;2-住院
    '   lngModule -调用模块号
    '返回:合法返回返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If int门诊标志 = 1 Then
        If gTy_System_Para.TY_Balance.byt门诊检查未执行 = 0 Then CheckNotExcuteItemValied = True: Exit Function
    Else
        If gTy_System_Para.TY_Balance.byt检查未执行 = 0 Then CheckNotExcuteItemValied = True: Exit Function
    End If
    
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_CheckNotExcuteItem(lng病人ID, lng主页ID, -1, 2, strInfo) = False Then Exit Function
    If strInfo = "" Then CheckNotExcuteItemValied = True: Exit Function
        
    If gTy_System_Para.TY_Balance.byt检查未执行 = 1 And int门诊标志 <> 1 _
        Or gTy_System_Para.TY_Balance.byt门诊检查未执行 = 1 And int门诊标志 = 1 Then
        If MsgBox("发现病人" & str姓名 & "存在尚未执行完成的内容：" & _
            vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    Else
        MsgBox "发现病人" & str姓名 & "存在尚未执行完成的内容：" & vbCrLf & vbCrLf & strInfo & _
            vbCrLf & vbCrLf & "不允许" & IIf(int门诊标志 <> 2, "门诊", "出院") & "结帐.", vbInformation, gstrSysName
        Exit Function
    End If
    CheckNotExcuteItemValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlAuditAdviceCharge(ByVal lng医嘱ID As Long, ByVal bln审核 As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对指定医嘱进行费用审核完成
    '入参:
    '   lng医嘱ID=医嘱ID
    '   bln审核=是否审核:True-审核;False-取消审核
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCisSvr_AuditAdviceCharge(lng医嘱ID, bln审核, lngModule) = False Then Exit Function
    ZlAuditAdviceCharge = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetAdviceDefinedInfo(ByRef rsAdviceDefinedInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱内容定义的相关信息
    '入参:
    '出参:
    '   rsAdviceDefinedInfo 医嘱内容定义信息：诊疗类别,医嘱内容
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAdviceDefinedInfo(rsAdviceDefinedInfo, lngModule) = False Then Exit Function
    ZlGetAdviceDefinedInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetAdviceOperMaxTime(ByVal lng医嘱ID As Long, ByVal byt操作类型 As Byte, _
    Optional ByVal lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱操作最后一次的时间
    '入参:
    '   lng医嘱ID=医嘱ID
    '   byt操作类型=操作类型:1-新开；2-校对疑问；3-校对通过；4-作废；5-重整；
    '                       6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果；
    '                       11-审核通过；12-审核未通过；13-实习医师停嘱后待审核；14-血库接收；15-血库审核通过；
    '                       16-血库配血拒绝；17-血库停止配血；18-输血初审通过待签发；9-输血初审回退；20-输血医嘱标记未用
    '出参:
    '返回:医嘱操作的最后一次时间
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim rsAdviceOper As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAdviceOperInfo(lng医嘱ID, byt操作类型, rsAdviceOper, True, lngModule) = False Then Exit Function
    If Not rsAdviceOper.EOF Then ZlGetAdviceOperMaxTime = Nvl(rsAdviceOper!操作时间)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetAdviceOperLastNotes(ByVal lng医嘱ID As Long, ByVal byt操作类型 As Byte, _
    Optional ByVal lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱操作最后一次的说明
    '入参:
    '   byt操作类型=操作类型:1-新开；2-校对疑问；3-校对通过；4-作废；5-重整；
    '                       6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果；
    '                       11-审核通过；12-审核未通过；13-实习医师停嘱后待审核；14-血库接收；15-血库审核通过；
    '                       16-血库配血拒绝；17-血库停止配血；18-输血初审通过待签发；9-输血初审回退；20-输血医嘱标记未用
    '出参:
    '返回:医嘱操作最后一次的说明
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim rsAdviceOper As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If lng医嘱ID = 0 Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAdviceOperInfo(lng医嘱ID, byt操作类型, rsAdviceOper, True, lngModule) = False Then Exit Function
    If Not rsAdviceOper.EOF Then ZlGetAdviceOperLastNotes = Nvl(rsAdviceOper!操作说明)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetMultiAdviceOperLastNotes(ByVal str医嘱IDs As String, ByVal byt操作类型 As Byte, _
    ByRef rsAdviceOper As ADODB.Recordset, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:批量获取医嘱操作最后一次的说明
    '入参:
    '   str医嘱IDs=医嘱id,多个用逗号分隔
    '   byt操作类型=操作类型:1-新开；2-校对疑问；3-校对通过；4-作废；5-重整；
    '                       6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果；
    '                       11-审核通过；12-审核未通过；13-实习医师停嘱后待审核；14-血库接收；15-血库审核通过；
    '                       16-血库配血拒绝；17-血库停止配血；18-输血初审通过待签发；9-输血初审回退；20-输血医嘱标记未用
    '出参:
    '   rsAdviceOper=医嘱操作信息：医嘱ID,操作时间,操作说明
    '返回:医嘱操作最后一次的说明
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If str医嘱IDs = "" Then Exit Function
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetAdviceOperInfo(str医嘱IDs, byt操作类型, rsAdviceOper, True, lngModule) = False Then Exit Function
    ZlGetMultiAdviceOperLastNotes = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlUpdatePatiAuditInfo(ByVal lng病人ID As Long, _
    ByVal lng主页ID As Long, ByVal byt审核标记 As Byte, ByVal bln取消审核 As Boolean, _
    Optional ByVal str审核人 As String, Optional ByVal str审核说明 As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新病人审核信息
    '入参:
    '   byt审核标记=审核标记：0或空-未审核,1-已审核或开始审核;2-完成审核
    '   bln取消审核=是否取消审核：1-取消审核,0-审核
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCisSvr_UpdatePatiAuditInfo(lng病人ID, lng主页ID, _
        byt审核标记, bln取消审核, str审核人, str审核说明, lngModule) = False Then Exit Function
    ZlUpdatePatiAuditInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlUpdateInpatientExtendInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str更新信息 As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改病案主页从表相关信息
    '入参:
    '   str更新信息=格式：信息名:信息值,信息名:信息值,...
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCisSvr_UpdateInpatientExtendInfo(lng病人ID, lng主页ID, str更新信息, lngModule) = False Then Exit Function
    ZlUpdateInpatientExtendInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetGroupAdviceInfo(ByVal str医嘱IDs As String, ByRef rsAdvice As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一并给药的医嘱内容
    '入参:
    '   str医嘱IDs=医嘱id,多个用逗号分隔
    '出参:
    '   rsAdvice=医嘱操作信息：医嘱ID,医嘱内容
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetGroupAdviceInfo(str医嘱IDs, rsAdvice, lngModule) = False Then Exit Function
    ZlGetGroupAdviceInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlUpdateAdviceExeStatus(ByVal strNO As String, ByVal str序号s As String, _
    ByVal byt记录性质 As Byte, ByVal byt费用来源 As Byte, ByVal blnCancelExe As Boolean, _
    Optional ByVal str执行人 As String, Optional ByVal str执行时间 As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新医嘱发送的执行状态
    '入参:
    '   strNo：费用单据号
    '   str序号s：费用序号
    '   byt记录性质-1-收费;2-记帐;3-自动记帐
    '   byt费用来源-费用来源：1-门诊，2-住院
    '   blnCancelExe-是否是取消执行
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '   处理医嘱的执行状态,如果同一张医嘱发送单都执行完了,才更新执行状态为已执行
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim cllAdviceDatas As Collection, cllItem As Collection
    '   cllAdviceDatas(collect)-数据集，格式如下
    '     |-cllAdviceData(collect)每行明细数据集
    '        |-成员(医嘱ID,费用单号,单据性质,执行状态,执行人,执行时间,原执行状态)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mdlDrugAndStuffSvr.zlGetServiceObject(objService) = False Then Exit Function
    
    '不含药品和跟踪在用的卫材
    strSQL = _
        " Select 1 As 类型,医嘱序号" & _
        " From 住院费用记录 A" & _
        " Where a.No = [1] And a.记录性质 = [2] And a.记录状态 In (0, 1, 3) And a.医嘱序号 Is Not Null" & _
        "       And (Instr(',' || [3] || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 Or [3] Is Null)" & _
        "       And Exists(Select 1 From 住院费用记录" & _
        "                  Where NO = a.No And 记录性质 = a.记录性质 And 医嘱序号 = a.医嘱序号" & _
        "                        And 执行状态 = 0 And 记录状态 In (0, 1, 3))" & _
        "       And a.收费类别 Not In ('5', '6', '7')" & _
        "       And Not Exists(Select 1 From 材料特性 B Where a.收费细目id = b.材料id And a.收费类别 = '4'  And Nvl(b.跟踪在用, 0) = 1)"

    strSQL = strSQL & " Union All" & _
        " Select 2 As 类型,医嘱序号" & _
        " From 住院费用记录 A" & _
        " Where a.No = [1] And a.记录性质 = [2] And a.记录状态 In (0, 1, 3) And a.医嘱序号 Is Not Null" & _
        "       And (Instr(',' || [3] || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 Or [3] Is Null)" & _
        "       And Not Exists(Select 1 From 住院费用记录" & _
        "                      Where NO = a.No And 记录性质 = a.记录性质 And 医嘱序号 = a.医嘱序号" & _
        "                            And 执行状态 = 0 And 记录状态 In (0, 1, 3))" & _
        "       And a.收费类别 Not In ('5', '6', '7')" & _
        "       And Not Exists(Select 1 From 材料特性 B Where a.收费细目id = b.材料id And a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1)"
    
    If byt费用来源 <> 2 Then strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    If blnCancelExe Then strSQL = Replace(strSQL, "And 执行状态 = 0", "And 执行状态 In(1,2)")
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱费用", strNO, byt记录性质, str序号s)
    If rsTemp.EOF Then ZlUpdateAdviceExeStatus = True: Exit Function
    
    Set cllAdviceDatas = New Collection
    Do While Not rsTemp.EOF
        Set cllItem = New Collection
        cllItem.Add Val(Nvl(rsTemp!医嘱序号)), "医嘱ID"
        cllItem.Add strNO, "费用单号"
        cllItem.Add byt记录性质, "单据性质"
        cllItem.Add str执行人, "执行人"
        cllItem.Add str执行时间, "执行时间"
        If blnCancelExe Then
            cllItem.Add IIf(Val(Nvl(rsTemp!类型)) = 1, 3, 0), "执行状态"
            cllItem.Add IIf(Val(Nvl(rsTemp!类型)) = 1, "1", "1,3"), "原执行状态"
        Else
            cllItem.Add IIf(Val(Nvl(rsTemp!类型)) = 1, 3, 1), "执行状态"
            cllItem.Add IIf(Val(Nvl(rsTemp!类型)) = 1, "0", "0,3"), "原执行状态"
        End If
        cllAdviceDatas.Add cllItem
        rsTemp.MoveNext
    Loop
    If objService.ZlCisSvr_UpdateAdviceExeStatus(cllAdviceDatas, lngModule) = False Then Exit Function
    ZlUpdateAdviceExeStatus = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
