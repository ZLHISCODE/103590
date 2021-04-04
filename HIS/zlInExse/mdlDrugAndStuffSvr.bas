Attribute VB_Name = "mdlDrugAndStuffSvr"
Option Explicit
'*********************************************************************************************************************************************
'功能:药品及卫材相关处理
'函数:
'    1.zlGetPublicDrugObjct:获取公共药品对象
'    2.zlGetServiceObject:获取药品公共服务对象
'    3.zlGetStock:获取指定药品或卫生材料在指定库房中的可用库存数
'    4.zlGetMultiStock:获取指定药品或卫生材料在多个库房中的可用库存数
'    5.zlCheckWaitSendDrugAndSutff:检查未发药品及数据(不存在时，返回true,否则返回false)
'    6.zlDrugSvr_RecipeAffirm:收费、记帐等划价审核药品处方确认
'    7.zlStuffSvr_BillAffirm:收费、记帐等划价审核卫材处方确认
'    8.zlGetDrugSendWindows:获取发药窗口
'出参:
'编制:刘兴洪
'日期:2019*08*08 19:20:59
'*********************************************************************************************************************************************
Private mobjService  As zlPublicExpense.clsService    '药品及卫材相关服务处理
Private mobjPublicDrug As Object '药品公共部件,105875

Public Function zlGetPublicDrugObjct(Optional ByRef objPubDrug As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:公共药品部件
    '出参:objPubDrug-返回公共药品相关对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 21:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjPublicDrug Is Nothing Then Set objPubDrug = mobjPublicDrug: zlGetPublicDrugObjct = True: Exit Function

    Err = 0: On Error Resume Next
    Set mobjPublicDrug = CreateObject("zlPublicDrug.clsPublicDrug")
    If Err <> 0 Then
        MsgBox "药品公共部件（zlPublicDrug）创建失败，请与系统管员联系！", vbInformation, gstrSysName
        Exit Function
    End If
    Err = 0: On Error GoTo errHandle
    'Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If mobjPublicDrug.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
        MsgBox "药品公共部件（zlPublicDrug）初始化失败，请与系统管员联系！", vbInformation, gstrSysName
        Set mobjPublicDrug = Nothing: Exit Function
    End If
    zlGetPublicDrugObjct = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Set mobjPublicDrug = Nothing
End Function

Public Function zlGetServiceObject(Optional ByRef objService As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共服务对象
    '出参:objService-返回公共服务对象
    '返回:获取返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-30 10:08:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If Not mobjService Is Nothing Then Set objService = mobjService: zlGetServiceObject = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjService = CreateObject("zlPublicExpense.clsService")
    If Err <> 0 Then
        MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense.clsService)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
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
    If mobjService.zlInitCommon(glngSys, glngModul, gcnOracle, gstrDBUser) = False Then
         MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)初始化失败，请与系统管理员联系！", vbExclamation, gstrSysName
         Exit Function
    End If
    Set objService = mobjService
    zlGetServiceObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Set mobjService = Nothing
End Function

Public Function zlCheckPriceAdjustBySellFromBillDetails(ByVal objBillDetails As BillDetails) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据明细对象进行零差价检查
    '入参:objBill-费用单据对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 21:18:15
    '说明:零差价检查,105875
    '
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubDrug As Object
    Dim i As Long
    
    On Error GoTo errHandle
    If zlGetPublicDrugObjct(objPubDrug) = False Then Exit Function
    If objBillDetails Is Nothing Then Exit Function
    

    'Private Function zlCheckPriceAdjustBySell(ByVal lng药品id As Long, ByVal lng药房id As Long) As Boolean
    '零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
    '定价药品：售价是固定的，比较所有药房的成本价，如果存在不一致的就不能销售出库
    '时价药品：比较药房库存记录的零售价和成本价，如果存在不一致的就不能销售出库
    '销售出库时只判断药房
    '返回：True-正常进行销售出库；false-不能进行销售出库
    For i = 1 To objBillDetails.Count
        With objBillDetails(i)
            If InStr(",5,6,7,", .收费类别) > 0 Then
                If objPubDrug.zlCheckPriceAdjustBySell(.收费细目ID, .执行部门ID) = False Then
                    Exit Function
                End If
            End If
        End With
    Next
    zlCheckPriceAdjustBySellFromBillDetails = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetStock(ByVal bln卫材 As Boolean, ByVal lng收费细目ID As Long, ByVal lng库房ID As Long, _
    Optional ByVal lng批次 As Long = -1, Optional ByVal lngMoudle As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定药品或卫生材料在指定库房中的可用库存数(以零售单位)
    '入参:bln卫材-是否卫生材料
    '     lng收费细目ID-药品ID或卫材ID
    '     lng库房ID-lng库房ID
    '     lng批次-批次(获取库存(不分批或分批),药房不分批(批次=0,这里为药房)不管效期)
    '     lngMoudle-当前调用的模块号
    '出参:
    '返回:返回药品或卫材库存
    '编制:刘兴洪
    '日期:2019-08-08 19:25:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl库存 As Double
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    If bln卫材 Then
        If objService.zlStuffSvr_GetStock(lng收费细目ID, lng库房ID, lng批次, dbl库存, lngMoudle) = False Then Exit Function
    Else
        If objService.zlDrugSvr_GetStock(lng收费细目ID, lng库房ID, lng批次, dbl库存, lngMoudle) = False Then Exit Function
    End If
    zlGetStock = dbl库存
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetMultiStock(ByVal lng收费细目ID As Long, ByVal str库房Ids As String, Optional ByVal bln卫材 As Boolean = False, Optional lngModule As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定药品或卫生材料在多个库房中的可用库存数(以零售单位)
    '入参:lng收费细目ID-药品ID或卫材ID
    '     str库房IDs-库房ID:多个用逗号
    '     lngModule-模块号
    '出参:
    '返回:返回库存数
    '编制:刘兴洪
    '日期:2019-08-08 21:32:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl库存 As Double
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    If zlGetServiceObject(objService) = False Then Exit Function
    If bln卫材 Then
        If objService.zlStuffSvr_GetMultiStock(lng收费细目ID, str库房Ids, dbl库存, lngModule) = False Then Exit Function
    Else
        If objService.zlDrugSvr_GetMultiStock(lng收费细目ID, str库房Ids, dbl库存, lngModule) = False Then Exit Function
    End If
    zlGetMultiStock = dbl库存
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetStockInfo(lng药品ID As Long, bln药房 As Boolean, bln药库 As Boolean, Optional ByVal dbl换算系数 As Double, Optional lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品在各个药房，药库的库存信息
    '入参:objPati-病人信息集
    '     "bln药房/bln药库"至少要有一个设置为真
    '     dbl换算系数-换算系数
    '     lngModule-模块号
    '出参:
    '返回:成功返回库存信息
    '编制:刘兴洪
    '日期:2019-08-08 22:28:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str库房性质 As String, strStockInfor As String
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    
    If bln药房 And bln药库 Then
        str库房性质 = "中药房,西药房,成药房,中药库,西药库,成药库"
    ElseIf bln药房 Then
        str库房性质 = "中药房,西药房,成药房"
    ElseIf bln药库 Then
        str库房性质 = "中药库,西药库,成药库"
    End If
  
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlDrugSvr_GetStockInfo(lng药品ID, str库房性质, dbl换算系数, strStockInfor, lngModule) = False Then Exit Function
    zlGetStockInfo = strStockInfor
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlCheckValidity(ByVal lng材料ID As Long, ByVal lng库房ID As Long, ByVal dbl数量 As Double, _
    Optional ByVal blnAsk As Boolean = True, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查卫生材料的灭菌效期是否过期
    '入参:objPati-病人信息集
    '     blnAsk=表示是否询问是否继续,否则为提醒
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 23:21:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
  
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlStuffSvr_CheckValidity(lng材料ID, lng库房ID, dbl数量, blnAsk, lngModule) = False Then Exit Function
    zlCheckValidity = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Public Function zlCheckWaitSendDrugAndSutff(ByVal str姓名 As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    Optional ByVal int检查离院带药 As Integer = 0, Optional int门诊标志 As Integer = 1, _
    Optional ByVal int婴儿序号 As Integer = -1, Optional ByVal strNos As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人在药房是否还有未发药的药品或卫材
    '入参:lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     int检查离院带药-离院带药
    '     int门诊标志-1-门诊;2-住院
    '     lngModule -调用模块号
    '返回:不存在时，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotSendInfor As String
    Dim objExpenceSvr As clsExpenceSvr
    
    On Error GoTo errHandle
    
    If int门诊标志 = 1 Then
        If gTy_System_Para.TY_Balance.byt门诊检查未发药 = 0 Then zlCheckWaitSendDrugAndSutff = True: Exit Function
    Else
        If gTy_System_Para.TY_Balance.byt检查未发药 = 0 Then zlCheckWaitSendDrugAndSutff = True: Exit Function
    End If
    
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then
        MsgBox "费用公共部件(zlpubExpence)失败，请与管理员联系!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If objExpenceSvr.zlGetWaitExcuteDrugAndStuff(lng病人ID, lng主页ID, strNotSendInfor, int检查离院带药, int婴儿序号, int门诊标志, strNos, lngModule) = False Then Exit Function
    
    If strNotSendInfor = "" Then zlCheckWaitSendDrugAndSutff = True: Exit Function
    
    If gTy_System_Para.TY_Balance.byt检查未发药 = 1 And int门诊标志 <> 1 _
        Or gTy_System_Para.TY_Balance.byt门诊检查未发药 = 1 And int门诊标志 = 1 Then
        If MsgBox("发现病人" & str姓名 & strNotSendInfor & vbCrLf & vbCrLf & "要继续结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        zlCheckWaitSendDrugAndSutff = True: Exit Function
    End If
    MsgBox "发现病人" & str姓名 & strNotSendInfor & vbCrLf & vbCrLf & "不允许" & IIf(int门诊标志 <> 2, "门诊", "出院") & "结帐。", vbInformation, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Public Function zlStuffSvr_AutoBatchSendStuff(ByVal strNO As String, ByVal str费用IDs As String, ByVal strCurDate As String, ByVal str操作员姓名 As String, ByVal str操作员编号 As String, _
    Optional intBillType As Integer = 2, Optional lngModule As Long, Optional strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用卫材自动发料服务
    '入参:strNo-单据号
    '     strCurDate-当前操作时间
    '     str费用Ids-费用Ids
    '     str操作员姓名-操作员姓名
    '     str操作员编号-操作员编号
    '     lngModule -调用模块号
    '     intBilltype-单据类型(1-收费，2-记帐,3-记帐表)
    '     intSendMode-发药方式( 1-处方发药;2-批量发药;3-部门发药)
    '出参:strErrMsg_Out-发生错误时，返回错误信息
     '返回:自动发料成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '说明:
    '   不加错误捕获，由上级过程捕获(避免事务站用）
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objService As zlPublicExpense.clsService
    If zlGetServiceObject(objService) = False Then Exit Function
    zlStuffSvr_AutoBatchSendStuff = objService.zlStuffSvr_AutoSendStuffFromNo(strNO, str费用IDs, _
        strCurDate, str操作员姓名, str操作员编号, intBillType, lngModule, , strErrMsg_out)
    
End Function

Public Function zlDrugSvr_RecipeAffirm(ByVal cllRecipeData As Collection, ByVal str操作时间 As String, _
    ByVal str审核人编号 As String, ByVal str审核人姓名 As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费、记帐等划价审核药品处方确认
    '入参:
    '   cllRecipeData-处方数据,每项是数组:
    '       Array(单据类型,单据号,费用IDs,发药窗口,是否自动发放,自动发放明细IDs)
    '           单据类型:1-收费处方发药;2-记帐单处方发药;3-记帐表处方发药
    '           发药窗口:发药窗口1:药房ID1|…|发药窗口n:药房Idn
    '   str操作时间-当前操作的时间,审核时，代表审核时间
    '   lngModule-模块号
    '出参:
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService

    On Error GoTo errHandle
    If cllRecipeData.Count = 0 Then zlDrugSvr_RecipeAffirm = True: Exit Function '无药品相关处理，直接返回true
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlDrugSvr_RecipeAffirm(cllRecipeData, 0, str操作时间, str审核人编号, str审核人姓名, lngModule) = False Then Exit Function
    zlDrugSvr_RecipeAffirm = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlStuffSvr_BillAffirm(ByVal cllRecipeData As Collection, ByVal str操作时间 As String, _
    ByVal str审核人编号 As String, ByVal str审核人姓名 As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费、记帐等划价审核卫材处方确认
    '入参:
    '   cllRecipeData-处方数据,每项是数组:
    '       Array(单据类型,单据号,费用IDs,是否自动发放,自动发放明细IDs)
    '           单据类型:1 -收费处方发药;2-记帐单处方发药;3-记帐表处方发药
    '   str操作时间-当前操作的时间,审核时，代表审核时间
    '   lngModule-模块号
    '出参:
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    
    If cllRecipeData.Count = 0 Then zlStuffSvr_BillAffirm = True: Exit Function '无卫材相关处理，直接返回true
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlStuffSvr_BillAffirm(cllRecipeData, 0, str操作时间, str审核人编号, str审核人姓名, lngModule) = False Then Exit Function
    zlStuffSvr_BillAffirm = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetDrugSendWindows(ByVal lng病人ID As Long, ByVal int挂号有效天数 As Integer, ByVal str药房Ids As String, _
    ByRef rsSendWindows_out As ADODB.Recordset, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前单据的发药窗口
    '入参:objPati-病人信息集
    '    str药房Ids-药房ID1,缺省发药窗口|药房ID2,缺省发药窗口2|...
    '出参:rsSendWindows_out-药房ID,发药窗口
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-16 11:40:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    On Error GoTo errHandle
        
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlDrugSvr_GetSendWindows(2, lng病人ID, int挂号有效天数, str药房Ids, rsSendWindows_out, lngModule) = False Then Exit Function
    
    zlGetDrugSendWindows = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 


Public Function zlGetDrugAndStuff_ExcuteNum(ByVal strNos As String, ByRef rsRecipe As ADODB.Recordset, ByVal blnExistDrug As Boolean, _
    ByVal blnExistStuff As Boolean, Optional ByVal bytBillType As Byte = 2, Optional lngModule As Long, Optional ByVal str费用IDs As String, _
    Optional blnAppend As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取准退数量
    '入参:strNos:单据号，多个用英文逗号分隔
    '   bytBillType:1-收费单,2-记帐单
    '   blnExistDrug:是否含药品
    '   blnExistStuff:是否含卫材
    '   str费用IDs-多个用逗号，传入后，前面的单据号及单据类型就无效
    '   blnAppend-是否自动追加数据
    '出参: rsRecipe:处方卫材数据：处方单号,药品ID,处方明细ID,已发数量,商品条码,内部条码
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-16 14:48:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objExpenceSvr As clsExpenceSvr
    On Error GoTo errHandle
        
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then Exit Function
    If objExpenceSvr.zlGetDrugStuff_ExecutedNum(strNos, bytBillType, rsRecipe, _
        blnExistDrug, blnExistStuff, blnAppend, str费用IDs, , , lngModule) = False Then Exit Function
    zlGetDrugAndStuff_ExcuteNum = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

 
Public Function zlExcute_SendDrug(strNO As String, strTime As String, Optional ByVal intBillType As Integer = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行药品发放操作
    '入参:strNO-单据号
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 15:04:40
    '说明：普通发药时为病人科室，急诊、医技则为开单科室。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, dtCurdate As Date
    Dim objService As zlPublicExpense.clsService, objExpenceSvr As clsExpenceSvr
    Dim blnTrans As Boolean
    Dim str费用IDs As String, cllFeeIds As Collection, str禁止发药费用IDs As String
    On Error GoTo errHandle
        
    If zlGetServiceObject(objService) = False Then Exit Function
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then Exit Function
  
    strSQL = _
    " Select  b.ID, b.执行部门ID,c.同步标志 As 记费同步标志, c1.同步标志 As 作废同步标志 " & _
    " From 住院费用记录 B, 病人费用异常记录 C, 病人费用异常记录 C1" & _
    " Where b.NO=[1] And b.记录性质=2 And  b.记录状态 in (0,1,3) And b.价格父号 is null And nvl(b.执行状态,0)<>1 And B.登记时间+0=[3]" & _
    "       And instr(',5,6,7,',','||b.收费类别||',')>0 " & _
    "       And b.ID = c.费用ID(+) And c.产生环节(+) = 0" & _
    "       And b.ID = c1.费用ID(+) And c1.产生环节(+) = 1"
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, 2, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, 2)
    End If
    
    If rsTmp.RecordCount > 0 Then '排除进入输液配置中心的药品
        If mobjService.zlPivasSvr_Getinfusion_Record(strNO, str禁止发药费用IDs) = False Then Exit Function
    End If
    
    With rsTmp
        Do While Not .EOF
            If Val(Nvl(!记费同步标志)) <> 0 Or Val(Nvl(!作废同步标志)) <> 0 Then
                MsgBox "单据“" & strNO & "”为异常单据，正被其他人占用，请稍后再试!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            If Nvl(!执行部门ID) = "" Then
                MsgBox "单据“" & strNO & "”中存在未确定执行药房，不能在这里发药。", vbInformation, gstrSysName
                Exit Function
            End If
            
            If InStr("," & str禁止发药费用IDs & ",", "," & Val(Nvl(!ID)) & ",") = 0 Then
                str费用IDs = IIf(str费用IDs = "", "", str费用IDs & ",") & Val(Nvl(!ID))
            End If
            .MoveNext
        Loop
    End With
    
    If str费用IDs = "" Then
        MsgBox "单据""" & strNO & """当前内容中没有可以发放的药品！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Set cllFeeIds = New Collection  'array(费用ids,执行状态)
    cllFeeIds.Add Array(str费用IDs, 1)
    
    dtCurdate = zlDatabase.Currentdate
    
    gcnOracle.BeginTrans: blnTrans = True
    
    If objExpenceSvr.zlUpdateExcuteStatu(cllFeeIds) = False Then Exit Function
    If objService.zlDrugsvr_AutoSendDrugFromNo(strNO, Format(dtCurdate, "yyyy-mm-dd HH:MM:SS"), UserInfo.姓名, UserInfo.编号, intBillType) = False Then Exit Function
   
    gcnOracle.CommitTrans: blnTrans = False
    
    zlExcute_SendDrug = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Public Function zlGetPrice(ByVal objBillDetail As BillDetail, ByVal dbl数量 As Double, ByRef dbl成本价 As Double, _
    Optional ByVal lngRow As Long, Optional ByVal lngModule As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取价格
    '入参:单据明细对象
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 16:02:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim dblPrice As Double, dbl空缺数量 As Double
        
    On Error GoTo errHandle
    dbl成本价 = 0
    If zlGetServiceObject(objService) = False Then Exit Function
  
  
    '获取药品/卫材价格
    If objBillDetail Is Nothing Then Exit Function
    
    With objBillDetail
        If .收费类别 = "4" Then
            If objService.zlStuffSvr_GetPrice(.收费细目ID, .执行部门ID, dbl数量, .Detail.批次, "", _
                  dblPrice, dbl成本价, dbl空缺数量, lngModule) = False Then
                '获取价格失败
                MsgBox IIf(lngRow > 0, "第" & lngRow & "行", "") & "卫生材料""" & .Detail.名称 & """获取价格失败！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If objService.zlDrugSvr_GetPrice(.收费细目ID, .执行部门ID, dbl数量, .Detail.批次, "", _
                  dblPrice, dbl成本价, dbl空缺数量, lngModule) = False Then
                '获取价格失败
                MsgBox IIf(lngRow > 0, "第" & lngRow & "行", "") & "药品""" & .Detail.名称 & """获取价格失败！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If dbl空缺数量 <> 0 And .Detail.变价 Then
            '数量未分解完毕
            If InStr(",5,6,7,", .收费类别) > 0 Then
                MsgBox IIf(lngRow > 0, "第" & lngRow & "行", "") & "时价药品""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
            Else
                MsgBox IIf(lngRow > 0, "第" & lngRow & "行", "") & "时价卫生材料""" & .Detail.名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    End With
    zlGetPrice = dblPrice
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlAutoSplitSpeci(ByVal lng药名ID As Long, ByVal byt中药形态 As Byte, _
    ByVal int付数 As Integer, ByVal dbl数量 As Double, ByVal lng药房ID As Long, _
    Optional ByVal byt场合 As Byte = 1, Optional lngModule As Long) As String
    '功能:针对中草药根据药名（品种）来自动分配药品(自动分解多个规格)
    '入参:
    '   byt中药形态 0-散装;1-中药饮片;2-免煎剂
    '   byt场合 1-门诊 ，2-住院
    '出参:
    '返回:格式：药品id,数量;药品id,数量;...(散装只选择一个规格)
    '               不能完全分配时返回:剂量为6和10的情况下,17克的分配=23755,6;23756,10|1
    '               不能分配时返回空,例如:剂量为6和10的情况下,3克的分配
    '编制:刘兴洪
    '日期:2019-08-18 15:44:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
        
    If zlGetServiceObject(objService) = False Then Exit Function

    If objService.zlDrugSvr_AutoSplitSpeci(lng药名ID, byt中药形态, int付数, dbl数量, lng药房ID, strData, byt场合, lngModule) = False Then Exit Function
    zlAutoSplitSpeci = strData
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ZlDrugsvr_GetStockByDrugName(ByVal lng药名ID As Long, ByVal lng药房ID As Long, int显示单位 As Integer, _
ByRef cllStockData_Out As Collection, ByVal int场合 As Integer, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据药名,获取指定库存信息
    '入参:lng药名ID-
    '     lng药房ID-药房ID,=0时表示不区分药房
    '     int场合-1-门诊;2-住院
    '     int显示单位-0-售价单位;1-门诊单位;2-住院单位
    '出参:cllStockData_Out-返回数据集，每个成员的key(pharmacy_id,drug_id,stock)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 17:30:45
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
        
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlDrugsvr_GetStockByDrugName(lng药名ID, lng药房ID, int显示单位, cllStockData_Out, int场合, lngModule) = False Then Exit Function
    ZlDrugsvr_GetStockByDrugName = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetStockCheck(ByVal bytType As Byte, Optional lngModule As Long) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品或卫材出库检查的集合
    '入参:bytType:0-药品，1-卫材
    '返回:返回库存检查方式
    '编制:刘兴洪
    '日期:2019-08-18 20:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllStockCheck As Collection, colStock As Collection
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
       
    If zlGetServiceObject(objService) = False Then GoTo ReInit
    If bytType = 1 Then
        If objService.zlStuffSvr_GetStockCheck(cllStockCheck, lngModule) = False Then GoTo ReInit
    Else
        If objService.zlDrugSvr_GetStockCheck(cllStockCheck, lngModule) = False Then GoTo ReInit
    End If
    
    Err = 0: On Error Resume Next
    cllStockCheck.Add 0, "_0"
    Set zlGetStockCheck = cllStockCheck
    On Error GoTo 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
ReInit:
    Set colStock = New Collection
    colStock.Add 0, "_0" '避免出错
    Set zlGetStockCheck = colStock
End Function


Public Function zlHaveNOAuditing(ByVal lng病人ID As Long, Optional ByVal str主页IDS As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断病人未结费用中是否存在未审核记帐费用
    '入参:str主页Ids-多个用逗号分离
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-19 09:20:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, str费用IDs As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim objService As zlPublicExpense.clsService
    Dim strNos As String, rsAdvice As ADODB.Recordset
    
    On Error GoTo errHandle
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.zlDrugSvr_GetRefuseSendList(lng病人ID, str主页IDS, str费用IDs, lngModule) = False Then Exit Function
 
    strSQL = _
        " Select Distinct a.医嘱序号,a.No,a.记录性质 From 住院费用记录 A" & _
        " Where 记帐费用=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And 病人ID=[1]" & _
                IIf(str费用IDs <> "", " And instr([3] ,','||ID||',')=0 ", "") & _
                IIf(str主页IDS <> "", " And a.主页ID In(Select /*+cardinality(j,10)*/ Column_Value From Table(f_num2list([2])) J)", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, str主页IDS, str费用IDs)
    If rsTmp.RecordCount = 0 Then Exit Function
    
    '已经存在非医嘱未审核的费用就不用再检查医嘱了
    strNos = ""
    Do While rsTmp.EOF
        If Val(Nvl(rsTmp!医嘱序号)) = 0 Then
            zlHaveNOAuditing = True: Exit Function
        Else
            strNos = strNos & "," & Nvl(rsTmp!医嘱序号) & ":" & Nvl(rsTmp!NO) & ":" & Nvl(rsTmp!记录性质)
        End If
        rsTmp.MoveNext
    Loop
    
    strNos = Mid(strNos, 2)
    If ZLGetAdviceSendInfo(1, strNos, rsAdvice) = False Then zlHaveNOAuditing = True: Exit Function   '保守策略
    
    rsAdvice.Filter = "执行状态<>2" '0-未执行;1-完全执行;2-拒绝执行;3-正在执行(今后可能分解为若干实际步骤)
    zlHaveNOAuditing = Not rsAdvice.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlCheckChargeAudit(ByVal lng病人ID As Long, ByVal bln出院 As Boolean, _
    Optional blnSaveCheck As Boolean = False, _
    Optional ByVal str主页IDS As String = "", Optional bln已选中途结帐_out As Boolean, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记帐审核检查
    '入参:
    '出参:bln已选中途结帐_out -结帐选择了中途结帐
    '返回:检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-04 15:04:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'bytAuditing:0-不检查,1-检查并提示,2-检查并禁止
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strWhere As String, str费用IDs As String
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    bln已选中途结帐_out = False
    If zlHaveNOAuditing(lng病人ID, str主页IDS, lngModule) = False Then zlCheckChargeAudit = True: Exit Function
    
    Select Case gTy_System_Para.TY_Balance.bytAuditing
    Case 1
        '在读取病人信息时,已经提示了
        If Not blnSaveCheck Then
            If MsgBox("该病人还存在未审核的记帐费用，要结帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    Case 2
        If blnSaveCheck Then
            If bln出院 Then
                MsgBox "该病人还存在未审核的记帐费用,不能出院结帐！", vbInformation, gstrSysName
                Exit Function
            End If
            '在读取病人信息时,已经提示了
        Else
            If MsgBox("该病人还存在未审核的记帐费用，不能出院结帐！" & vbCrLf & vbCrLf & _
                "是否对该病人进行中途结帐？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                bln已选中途结帐_out = True
        End If
    Case Else
    End Select
    zlCheckChargeAudit = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function zlPivasSvr_Getinfusion_RecordFeeids(ByVal strNos As String, ByRef str费用IDs_out As String, Optional ByVal lng病人ID As Long, Optional ByVal str主页IDS As String, _
    Optional ByVal str费用Ids_in As String = "", Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的拒发药清单
    '入参:lng病人ID-按病人ID查(str主页Ids-多个用逗号)
    '     strNos-按单据查
    '     str费用Ids_in-按费用ID查
    '     以上条件，必有一个
    '出参:str费用IDs_out-返回所有涉及进入输液配药记录中的费用IDs
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
     
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    zlPivasSvr_Getinfusion_RecordFeeids = objService.zlPivasSvr_Getinfusion_Record(strNos, str费用IDs_out, lng病人ID, str主页IDS, str费用Ids_in, lngModule)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlPivasSvr_Isexsitinfusion(ByVal str费用IDs As String, Optional blnIsExist_Out As Boolean, Optional lngModule As Long, _
    Optional lng医嘱ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查药品是否进行输液配药记录
    '入参: str费用Ids_in-按费用ID查
    '      lng医嘱ID-如果传入医嘱id,则按医嘱ID查证
    '出参:blnIsExist_Out-是否存在的，返回true,否则返回False
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
     
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    zlPivasSvr_Isexsitinfusion = objService.zlPivasSvr_Isexsitinfusion(str费用IDs, blnIsExist_Out, lngModule, lng医嘱ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ZlStuffSvr_AutoReturnStuff(ByVal strAutoStuffDatas As String, _
    Optional ByVal strTittle As String, Optional lngModule As Long, Optional rsSendDatas As ADODB.Recordset, _
    Optional ByVal bln门诊 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动退料操作
    '入参:strAutoStuffDatas-自动退料数据,格式:费用ID1:数量,费用ID2:数量2,...
    '    rsSendDatas-卫材已经执行数量
    '    blnReturnAll-是否将所涉及的费用ID所对应的剩余未退的全退（全退时，strAutoStuffDatas参数不传数量)
    '    bln不启用事务-内部不开启事务
    '出参:strErrMsg_out-失败时，返回错误信息
    '返回:退料成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService, strErrMsg_out As String
    Dim blnTrans As Boolean, varData As Variant, strFeeIds As String, lng费用ID As Long
    Dim i As Long, strSQL As String, varParaValue()  As Variant, rsTemp As ADODB.Recordset
    Dim cllExcuteUpdate As Collection, int执行状态 As Integer, dbl执行数 As Double
    Dim strSubTable As String, strFeeTable As String
    Dim strAutoStuff As String, dbl退料数 As Double
    
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    '为空时，直接返回true
    If strAutoStuffDatas = "" Then ZlStuffSvr_AutoReturnStuff = True: Exit Function
    
    
    If rsSendDatas Is Nothing Then
        '重新获取
        strFeeIds = ""
        varData = Split(strAutoStuffDatas, ",")
        For i = 0 To UBound(varData)
            lng费用ID = Split(varData(i) & ":", ":")(0)
            If InStr(strFeeIds & ",", "," & lng费用ID & ",") > 0 Then
                 strFeeIds = strFeeIds & "," & lng费用ID
            End If
        Next
        If objService.zlStuffSvr_GetExecutedNum("", 2, rsSendDatas, , strFeeIds, lngModule, True, strErrMsg_out) = False Then Exit Function
    End If
    
    '入参:
    '   bytType: 0-Num2List;1-Str2List;2-Num2List2;3-Str2List2
    '   strValues: bytType=0,1时,多个用","分离
    '              bytType=2,3时,列之间用":"分离,行之间用","分离:如:张三:22,李四:22
    '   lngStep: 步长(即绑定变量从好多开始)

    If zlGetVarBoundSQL(2, strAutoStuffDatas, strSubTable, varParaValue, 0) = False Then Exit Function
    
    strFeeTable = IIf(bln门诊, "门诊费用记录", "住院费用记录")
    
    strSQL = "With 销帐信息 As (" & strSubTable & ")" & vbCrLf & _
    "   Select  A.No,a.序号,a.收费细目ID,max(Decode(a.记录状态,2,0,a.ID)) as 费用id, " & vbCrLf & _
    "           sum(decode(a.记录状态,2,0,1)*nvl(a.付数,1)*nvl(a.数次,0)) as 原始数量," & vbCrLf & _
    "           sum(nvl(a.付数,1)*nvl(a.数次,0)) as 剩余数量," & vbCrLf & _
    "           max(B.申请数量) as 申请数量 " & vbCrLf & _
    "   From " & strFeeTable & " a ," & _
    "         ( Select /*+CARDINALITY(M,10)*/ distinct J.No,J.序号,C2 as 申请数量 From " & strFeeTable & " J,销帐信息  M  Where J.ID=M.C1 ) B " & vbCrLf & _
    "   Where a.NO=B.NO and a.序号=b.序号 And a.价格父号 is null " & _
    "   Group by A.NO,A.序号,a.收费细目ID "
    
    strSQL = "" & _
    "    Select A.No,A.序号,A.费用ID,A.原始数量,A.剩余数量,A.申请数量,C.名称 as 收费项目" & vbCrLf & _
    "    From (" & strSQL & ") A,收费项目目录 C" & vbCrLf & _
    "    Where  a.收费细目ID=C.ID" & vbCrLf & _
    "    Order by a.no,a.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, strTittle, varParaValue)
    
    '先检查
    Set cllExcuteUpdate = New Collection
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        strAutoStuff = ""
        Do While Not .EOF
            
            If Val(Nvl(!剩余数量)) = 0 Then
                strErrMsg_out = "【" & !名称 & "】已经没有可退数了，可能是因为并发原因被他人销帐，不能再退!"
                MsgBox strErrMsg_out, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            If Val(Nvl(!剩余数量)) < Val(Nvl(!申请数量)) Then
                strErrMsg_out = "【" & !名称 & "】的卫生材料的可销帐数量小于申请数量，不能再销帐!" & vbCrLf & _
                       "剩余数量:" & FormatEx(Val(Nvl(!剩余数量)), 4) & vbCrLf & _
                       "申请数量:" & FormatEx(Val(Nvl(!申请数量)), 4) & vbCrLf
                MsgBox strErrMsg_out, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            
            dbl执行数 = 0
            rsSendDatas.Filter = "处方明细ID=" & Val(Nvl(!费用ID))
            If rsSendDatas.EOF = False Then dbl执行数 = Val(Nvl(rsSendDatas!已发数量))
            rsSendDatas.Filter = 0
            
            '计算本次自动退料数量,规则如下:
            '1.申请数量>(剩余数量-已执行数量) 则本次自动退料=申请数量-(剩余数量-已执行数量)
            dbl退料数 = Val(Nvl(!申请数量)) - (Val(Nvl(!剩余数量)) - dbl执行数)
        
            If dbl执行数 < dbl退料数 And dbl退料数 <> 0 Then
                strErrMsg_out = "【" & !名称 & "】的卫生材料的本次退料数量大于了已发料数量,不能进行自动退料操作!" & vbCrLf & _
                       "已发料数量:" & FormatEx(dbl执行数, 4) & vbCrLf & _
                       "申请退料数量:" & FormatEx(dbl退料数, 4) & vbCrLf
                MsgBox strErrMsg_out, vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            If dbl退料数 > 0 Then
                If RoundEx(dbl执行数 - dbl退料数, 5) > 0 Then
                    '部分执行
                    int执行状态 = 2
                Else
                    int执行状态 = 0
                End If
                cllExcuteUpdate.Add Array(Val(Nvl(!费用ID)), int执行状态)
                strAutoStuff = strAutoStuff & "," & Val(Nvl(!费用ID)) & ":" & dbl退料数
            End If
            .MoveNext
        Loop
    End With
    If strAutoStuff = "" Then ZlStuffSvr_AutoReturnStuff = True: Exit Function
    strAutoStuff = Mid(strAutoStuff, 2)
 
    '自动退料
    gcnOracle.BeginTrans: blnTrans = True
    '更新的数据集(array(费用ids,执行状态))
    If mdlExseSvr.zlUpdateExcuteStatu(cllExcuteUpdate, IIf(bln门诊, 1, 2)) = False Then gcnOracle.RollbackTrans: Exit Function
    If objService.ZlStuffSvr_AutoReturnStuff(strAutoStuff, strErrMsg_out, strTittle, lngModule) = False Then
        gcnOracle.RollbackTrans: blnTrans = False
        If strErrMsg_out <> "" Then MsgBox strErrMsg_out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    ZlStuffSvr_AutoReturnStuff = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ZlStuffSvr_AutoReturnStuffFromFeeIds(ByVal strAutoStuffDatas As String, _
    Optional ByVal strTittle As String, Optional lngModule As Long, _
    Optional ByVal bln门诊 As Boolean, Optional strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用ID自动退料操作
    '入参:strAutoStuffDatas-自动退料数据,格式:费用ID1,费用ID2,...
    '出参:strErrMsg_out-失败时，返回错误信息
    '返回:退料成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim varData As Variant
    Dim i As Long, cllExcuteUpdate As Collection
    Dim strAutoStuff As String, blnTrans As Boolean
    
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    '为空时，直接返回true
    If strAutoStuffDatas = "" Then ZlStuffSvr_AutoReturnStuffFromFeeIds = True: Exit Function
    
    varData = Split(strAutoStuffDatas, ",")
    Set cllExcuteUpdate = New Collection
    For i = 0 To UBound(varData)
        cllExcuteUpdate.Add Array(varData(i), 0)
        strAutoStuff = strAutoStuff & "," & varData(i) & ":"    '接口要求，必传:号，表示按剩余全退
    Next
    If strAutoStuff <> "" Then strAutoStuff = Mid(strAutoStuff, 2)
     
     '更新的数据集(array(费用ids,执行状态))
    gcnOracle.BeginTrans: blnTrans = True
    If mdlExseSvr.zlUpdateExcuteStatu(cllExcuteUpdate, IIf(bln门诊, 1, 2)) = False Then gcnOracle.RollbackTrans: Exit Function
    If objService.ZlStuffSvr_AutoReturnStuff(strAutoStuff, strErrMsg_out, strTittle, lngModule) = False Then gcnOracle.RollbackTrans: Exit Function
    gcnOracle.CommitTrans: blnTrans = False
    ZlStuffSvr_AutoReturnStuffFromFeeIds = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    strErrMsg_out = Err.Description
End Function


Public Function zlExecuteUpdateSyncSymbol(ByVal str费用IDs As String, ByVal byt标志类型 As Byte, _
    ByVal byt门诊标志 As Byte, ByVal byt原标志值 As Byte, Optional ByVal byt新标志值 As Byte = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新费用记录中的药品/卫材同步标志
    '入参:
    '   str费用IDs 需要更新的费用ID，多个用英文逗号分隔
    '   byt标志类型 标志类型：0-记费同步标志,1-作废同步标志
    '   byt门诊标志 门诊标志：1-门诊，2-住院
    '出参:
    '返回:
    '  记费同步标志：NULL或0-正常，1-未生成处方单/发料单，2-未更新药品/卫材收费状态
    '  作废同步标志：NULL或0-正常，1-药品/卫材已作废但费用未作废(禁止发放/退回)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objExpenceSvr As clsExpenceSvr
    On Error GoTo errHandle
    
    If zlGetPubExseSvrObject(objExpenceSvr) = False Then Exit Function
     
    If objExpenceSvr.zlExecuteUpdateSyncSymbol(str费用IDs, byt标志类型, byt门诊标志, byt原标志值, byt新标志值) = False Then Exit Function
    zlExecuteUpdateSyncSymbol = True
    Exit Function
errHandle:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

