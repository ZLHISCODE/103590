Attribute VB_Name = "mdlRecipeAudit"
Option Explicit

Public gobjRecipeAuditEx As Object

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
'返回：True成功；False失败
    Dim rsTmp As ADODB.Recordset
    
    UserInfo.姓名 = UserInfo.用户名
    Set rsTmp = SYS.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.编号 = rsTmp!编号
            UserInfo.部门ID = zlCommFun.NVL(rsTmp!部门ID, 0)
            UserInfo.简码 = zlCommFun.NVL(rsTmp!简码)
            UserInfo.姓名 = zlCommFun.NVL(rsTmp!姓名)
            UserInfo.用户名 = rsTmp!用户名
            GetUserInfo = True
        End If
        rsTmp.Close
    End If
End Function

Public Function AuditDrug(ByVal fldItem As Fields, ByVal lngPatientID As Long, _
    ByVal bytMedicalClass As Byte, ByVal lngBillID As Long, _
    ByVal strSubmitID As String, ByRef strMedicalID As String) As Byte
'功能：审查药嘱
'参数：
'  fldItem：处方审查项目
'  lngPatientID：病人ID
'  bytMedicalClass：病人临床类别；1-门诊；2-住院
'  lngBillID：门诊病人为挂号ID；住院病人为主页ID
'  strSubmitID：详见clsBusiness.AutoAudit函数的strSubmitID参数说明（给药途径医嘱ID）
'  strMedicalID(实参)：不合格医嘱ID；如果无值表示整批医嘱（药品医嘱ID）
'返回：0-异常/未知；1-合格；2-不合格
    
    Dim strID As String, strIDs As String, strReturn As String, strErr As String
    Dim objRecipeAuditEx As Object
    Dim l As Long

    On Error GoTo errHandle
    
    strIDs = GetMedicalID(strSubmitID)  '将相关ID转换成医嘱ID
    strID = strSubmitID                 '
    
    Select Case UCase(fldItem!编码)
        Case "A01"          '皮试
            strMedicalID = RAI_AllergicTest(lngPatientID, bytMedicalClass, lngBillID, strIDs)
            AuditDrug = IIf(strMedicalID = "", 1, 2)
            
        Case "A02"
        Case "A03"
        Case "A04"
        Case "A05", "2-7"  '重复给药
            strMedicalID = RAI_RepeatDrug(lngPatientID, bytMedicalClass, lngBillID, strID)
            AuditDrug = IIf(strMedicalID = "", 1, 2)
            
        Case "A06"
        Case "A07"
        Case "1-4"         '新生儿、婴幼儿未写明日、月龄
            If bytMedicalClass = 2 Then
                '只有住院存在这种情况
                AuditDrug = RAI_InfantAge(lngPatientID, lngBillID)
            Else
                AuditDrug = 1
            End If
            
        Case "1-9"         '处方修改未签名或药品超量未注明原因
            strMedicalID = RAI_OverloadExplain(lngPatientID, bytMedicalClass, lngBillID, strID)
            AuditDrug = IIf(strMedicalID = "", 1, 2)
            
        Case "1-10"         '未写临床诊断或书写不全
            AuditDrug = RAI_Diagnosis(lngPatientID, bytMedicalClass, lngBillID, strID)
            
        Case "1-14"         '未按抗菌药物管理开具
            strMedicalID = RAI_AntibiosisManage(lngPatientID, bytMedicalClass, lngBillID, strID)
            AuditDrug = IIf(strMedicalID = "", 1, 2)
            
        Case "C01"          'PASS结果
            strMedicalID = RAI_PASS(lngPatientID, bytMedicalClass, lngBillID, NVL(fldItem!PASS结果), strID)
            AuditDrug = IIf(strMedicalID = "", 1, 2)
            
        Case Else           '自定义审查项目
        
            If fldItem!类别 = 4 Then
                If gobjRecipeAuditEx Is Nothing Then
                    On Error Resume Next
                    Set gobjRecipeAuditEx = CreateObject("zlRecipeAuditEx.clsRecipeAuditEx")
                    If gobjRecipeAuditEx Is Nothing Then
                        Err.Clear
                        gstrErrInfo = gstrErrInfo & vbCr & "创建“zlRecipeAuditEx”部件失败"
                        Exit Function
                    End If
                    Err.Clear: On Error GoTo errHandle
                End If
                If gobjRecipeAuditEx.Init(gcnOracle, bytMedicalClass, strSubmitID) Then
                    If gobjRecipeAuditEx.Check(UCase(fldItem!编码), strReturn, strErr) Then
                        '合格
                        AuditDrug = 1
                    Else
                        '有返回医嘱ID为不合格，反之没有检查出结果，即未审查
                        AuditDrug = IIf(strReturn = "", 0, 2)
                    End If
                End If
            End If
            
    End Select
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "AuditDrug：" & vbCr & Err.Description
End Function

Private Function RAI_PASS(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strAuditPASS As String, ByVal strID As String) As String
'功能：PASS结果
'参数：
'  lngPatientID：病人ID
'  bytMedicalClass：病人临床类别；1-门诊；2-住院
'  lngBillID：门诊病人为挂号ID；住院病人为主页ID
'  strAuditPASS：要检查的PASS审查结果值
'  strID：医嘱ID；     格式：给药途径医嘱ID[,给药途径医嘱ID]
'返回：不合格的医嘱ID。 格式：医嘱ID[;医嘱ID]

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    If strAuditPASS = "" Then
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    strSQL = "Select a.Id " & vbCr & _
             "From 病人医嘱记录 A, Table(f_Num2list([1], ',')) B, Table(f_Num2list([2], ';')) C " & vbCr & _
             "Where a.相关Id = b.Column_Value And a.审查结果 = c.Column_Value And a.病人id = [3] And a.诊疗类别 in ('5','6','7') "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取医嘱的PASS审查结果值", strID, strAuditPASS, lngPatientID)
    With rsTemp
        Do While .EOF = False
            strReturn = strReturn & zlStr.FormatString("[1];", !ID)
            .MoveNext
        Loop
        If strReturn <> "" Then strReturn = Left(strReturn, Len(strReturn) - 1)
        .Close
    End With
    
    RAI_PASS = strReturn
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_PASS：" & vbCr & Err.Descriptions
End Function

Private Function RAI_AntibiosisManage(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strID As String) As String
'功能：抗菌药品是否按管理开具
'参数：
'  lngPatientID：病人ID
'  bytMedicalClass：病人临床类别；1-门诊；2-住院
'  lngBillID：门诊病人为挂号ID；住院病人为主页ID
'  strID：给药途径医嘱ID；   格式：给药途径医嘱ID[,给药途径医嘱ID]
'返回：不合格的医嘱ID；      格式：医嘱ID[;医嘱ID]

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '检查是否有限制使用的抗菌药物
    strSQL = "Select a.Id " & vbCr & _
             "From 病人医嘱记录 A, 药品特性 B, Table(f_Num2list([1], ',')) C " & vbCr & _
             "Where a.诊疗项目id = b.药名id And a.相关Id = c.Column_Value And b.抗生素 > 1 " & vbCr & _
             "    And a.病人id = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取抗菌药物", strID, lngPatientID)
    With rsTemp
        Do While .EOF = False
            strReturn = strReturn & zlStr.FormatString(";[1]", !ID)
            .MoveNext
        Loop
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        .Close
    End With

    '有限制使用的抗菌药药嘱
    If strReturn <> "" Then
        If Val(zlDatabase.GetPara("抗菌药物分级管理", glngSys)) <> 1 Then
            '未启用抗菌药物分级管理，根据strReturn确定是否合格
            RAI_AntibiosisManage = strReturn
        Else
            '已启用抗菌药物分级管理，表示合格
            RAI_AntibiosisManage = ""
        End If
    End If
    
    Exit Function

errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_AntibiosisManage：" & vbCr & Err.Descriptions
End Function

Private Function RAI_Diagnosis(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strID As String) As Byte
'功能：未写临床诊断或书写不全
'参数：
'  lngPatientID：病人ID
'  bytMedicalClass：病人临床类别；1-门诊；2-住院
'  lngBillID：门诊病人为挂号ID；住院病人为主页ID
'  strID：给药途径医嘱ID；  格式：给药途径医嘱ID[,给药途径医嘱ID]
'返回：0-异常/未知；1-合格；2-不合格

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "Select Sum(Rec) Rec " & vbNewLine & _
             "From (Select Count(1) Rec " & vbNewLine & _
             "      From 病人诊断记录 C, Table(f_Num2list([1], ',')) E " & vbNewLine & _
             "      Where c.医嘱id = e.Column_Value And c.病人id = [2] And c.诊断描述 Is Not Null And Rownum < 2 " & vbNewLine & _
             "      Union All " & vbNewLine & _
             "      Select Count(1) Rec" & vbNewLine & _
             "      From 病人诊断记录 C, 病人诊断医嘱 D, Table(f_Num2list([1], ',')) E " & vbNewLine & _
             "      Where e.Column_Value = d.医嘱id And d.诊断id = c.Id And c.病人id = [2] And c.诊断描述 Is Not Null And Rownum < 2 " & vbNewLine & _
             "      Union All " & vbNewLine & _
             "      Select Count(1) Rec From 病人诊断记录 Where 诊断描述 Is Not Null And 病人id = [2] And 主页id = [3] " & vbNewLine & _
             ") A "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取医嘱对应诊断描述", strID, lngPatientID, lngBillID)
    If rsTemp!Rec <= 0 Then
        RAI_Diagnosis = 2
    Else
        RAI_Diagnosis = 1
    End If
    rsTemp.Close
    
    Exit Function

errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_Diagnosis：" & vbCr & Err.Description
End Function

Private Function RAI_OverloadExplain(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strID As String) As String
'功能：药品超量说明
'参数：
'  lngPatientID：病人ID
'  bytMedicalClass：病人临床类别；1-门诊；2-住院
'  lngBillID：门诊病人为挂号ID；住院病人为主页ID
'  strID：给药途径医嘱ID；  格式：给药途径医嘱ID[,给药途径医嘱ID]
'返回：不合格的医嘱ID       格式：医嘱ID[;医嘱ID]

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "Select a.ID " & vbCr & _
             "From 病人医嘱记录 A, 药品特性 B, Table(f_Num2list([1], ',')) C " & vbCr & _
             "Where a.诊疗项目id = b.药名id And a.相关Id = c.Column_Value And a.诊疗类别 In ('5', '6', '7') " & vbCr & _
             "  And Nvl(a.总给予量, 0) > Nvl(b.处方限量, 0) And Nvl(b.处方限量, 0) > 0 " & vbCr & _
             "  And a.超量说明 is null And a.病人ID = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取药品超量", strID, lngPatientID)
    With rsTemp
        Do While .EOF = False
            strReturn = strReturn & zlStr.FormatString("[1];", !ID)
            .MoveNext
        Loop
        If strReturn <> "" Then strReturn = Left(strReturn, Len(strReturn) - 1)
        .Close
    End With
    
    RAI_OverloadExplain = strReturn
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_OverloadExplain：" & vbCr & Err.Descriptions
End Function

Private Function RAI_AllergicTest(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strID As String) As String
'功能：药品过敏试验
'参数：
'  lngPatientID：病人ID
'  bytMedicalClass：病人临床类别；1-门诊；2-住院
'  lngBillID：门诊病人为挂号ID；住院病人为主页ID
'  strID：医嘱ID；      格式：医嘱ID[,医嘱ID]
'返回：不合格的医嘱ID    格式：医嘱ID[;医嘱ID]

    Dim arrID As Variant
    Dim l As Long
    Dim strReturn As String
    Dim intResult As Integer
    
    On Error GoTo errHandle
    
    If gobjPubAdvice Is Nothing Then Exit Function
    
    arrID = Split(strID, ",")
    For l = LBound(arrID) To UBound(arrID)
        '调用zlPublicAdvice.CheckAdviceSkinResult的皮试函数
        intResult = gobjPubAdvice.CheckAdviceSkinResult(Val(arrID(l)))
        '-1表示无需皮试或免试；0表示还未标记皮试结果或未下达皮试医嘱；1表示阴性；2表示阳性
        If intResult = 0 Or intResult = 2 Then
            strReturn = strReturn & arrID(l) & ";"
        End If
    Next
    If strReturn <> "" Then strReturn = Left(strReturn, Len(strReturn) - 1)
    
    RAI_AllergicTest = strReturn
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_AllergicTest：" & vbCr & Err.Description
End Function

Private Function RAI_RepeatDrug(ByVal lngPatientID As Long, ByVal bytMedicalClass As Byte, _
    ByVal lngBillID As Long, ByVal strID As String) As String
'功能：检查重复给药
'参数：
'  lngPatientID：病人ID
'  bytMedicalClass：病人临床类别；1-门诊；2-住院
'  lngBillID：门诊病人为挂号ID；住院病人为主页ID
'  strID：医嘱ID；      格式：给药途径医嘱ID[,给药途径医嘱ID]
'返回：不合格的医嘱ID    格式：医嘱ID[;医嘱ID]

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    
    '检查当前传入医嘱的药名是否重复
    strSQL = "Select Max(a.Id) 医嘱id, a.诊疗项目id " & vbCr & _
             "From 病人医嘱记录 A, Table(f_Num2list([1], ',')) B" & vbCr & _
             "Where a.相关Id = b.Column_Value And a.诊疗类别 In ('5', '6', '7') " & vbCr & _
             "Group By a.诊疗项目id " & vbCr & _
             "Having Count(a.诊疗项目id) > 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取重复药嘱", strID)
    Do While rsTemp.EOF = False
        strReturn = strReturn & zlStr.FormatString("[1];", rsTemp!医嘱ID)
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If strReturn <> "" Then
        strReturn = Left(strReturn, Len(strReturn) - 1)
        RAI_RepeatDrug = strReturn
        Exit Function
    End If
    
    strReturn = ""
    
    '再检查当前传入医嘱的药品与“发送”通过的药品（24小时内）是否重复
    strSQL = "Select b.医嘱id " & vbCr & _
             "From (Select Distinct a.诊疗项目id " & vbCr & _
             "      From 病人医嘱记录 A, 病人医嘱发送 B " & IIf(bytMedicalClass = 1, ", 病人挂号记录 C ", " ") & vbCr & _
             "      Where a.Id = b.医嘱id " & IIf(bytMedicalClass = 1, " And a.挂号单 = c.No ", " ") & "And a.诊疗类别 In ('5', '6', '7') " & vbCr & _
             "          And a.病人id = [1]" & IIf(bytMedicalClass = 1, " And c.Id = [2] ", " And a.主页ID = [2] ") & _
             "          And b.发送时间 >= Sysdate - [3] / 24 ) A, " & vbCr & _
             "     (Select a.诊疗项目id, a.Id 医嘱id " & vbCr & _
             "      From 病人医嘱记录 A, Table(f_Num2list([4], ',')) B " & vbCr & _
             "      Where a.相关Id = b.Column_Value And a.诊疗类别 In ('5', '6', '7') ) B " & vbCr & _
             "Where a.诊疗项目id = b.诊疗项目id "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取已发送的重复药嘱", lngPatientID, lngBillID, gintHoursRecipe, strID)
    Do While rsTemp.EOF = False
        strReturn = strReturn & zlStr.FormatString("[1];", rsTemp!医嘱ID)
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If strReturn <> "" Then
        strReturn = Left(strReturn, Len(strReturn) - 1)
        RAI_RepeatDrug = strReturn
        Exit Function
    End If
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_RepeatDrug：" & vbCr & Err.Description
End Function

Private Function RAI_InfantAge(ByVal lngPatientID As Long, ByVal lngMasterPageID As Long) As Byte
'功能：检查新生儿、婴幼儿填写日、月龄（住院类）
'参数：
'  lngPatientID：病人ID
'  lngMasterPageID：住院病人为主页ID
'返回：0-异常/未知；1-合格；2-不合格

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "Select Count(1) Rec " & vbCr & _
             "From 病人新生儿记录 " & vbCr & _
             "Where 病人id = [1] And 主页id = [2] And 出生时间 Is Null "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取新生儿、婴幼儿日月龄", lngPatientID, lngMasterPageID)
    If rsTemp!Rec > 0 Then
        '有记录表明新生儿没有填写出生时间
        RAI_InfantAge = 2
    Else
        '无记录表明没有新生儿记录，或者有填写出生时间
        RAI_InfantAge = 1
    End If
    rsTemp.Close
    
    Exit Function
    
errHandle:
    gstrErrInfo = gstrErrInfo & vbCr & "RAI_InfantAge：" & vbCr & Err.Description
End Function

'Public Function GetStoreID(ByVal lngID As Long) As Long
''功能：通过医嘱ID获取医嘱的执行科室（发药药房）ID
''参数：
''  lngID：医嘱ID
''返回：发药药房ID
'
'    Dim strSQL As String
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo errHandle
'
'    strSQL = "Select 执行科室ID from 病人医嘱记录 where ID = [1] "
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取执行科室ID", lngID)
'    If rsTemp.RecordCount = 1 Then
'        GetStoreID = NVL(rsTemp!执行科室ID, 0)
'    End If
'    rsTemp.Close
'
'errHandle:
'    If ErrCenter = 1 Then Resume
'End Function

Public Function GetCalorie(ByVal lngPatientID As Long, ByVal lngRegisterID As Long, ByVal lngPageID As Long) As String
'功能：获取病人热量需要量
'参数：
'  lngPatientID：病人ID
'  lngRegisterID：挂号单ID
'  lngPageID：主页ID
'返回：热量需要量公式和值（如：66.5 + 13.8 * 61KG + 5.0 * 172CM - 6.8 * 30岁 = 1564.30）

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If lngPageID <= 0 Then
        '门诊
        strSQL = "Select zl_fun_pati_calorie([1], Null, [2]) 热量 From Dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人热量需要量", lngPatientID, lngRegisterID)
    Else
        '住院
        strSQL = "Select zl_fun_pati_calorie([1], [2], Null) 热量 From Dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人热量需要量", lngPatientID, lngPageID)
    End If
    
    If rsTemp.EOF = False Then
        GetCalorie = NVL(rsTemp!热量)
    End If
    
    Exit Function

errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowReason(ByVal frmOwner As Form, ByVal strSQL As String, ByRef blnCancel As Boolean, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：调用理由选择器
'参数：
'  frmOwner：宿主窗体对象
'  strSQL：SQL查询
'  blnCancel（实参）：True选择确认；False选择取消
'返回：已选择的记录

    Dim frmSelector As New frmReasonSelector

    Set ShowReason = frmSelector.ShowMe(frmOwner, strSQL, blnCancel, arrInput)
    
End Function

Public Sub zlRptPrint(ByVal bytMode As Byte, ByVal vsfVar As VSFlexGrid, ByVal strTitle As String)
'-------------------------------------------------
'功能:记录表打印
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'-------------------------------------------------
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    Dim lngRow As Long
    Dim lngColor As Long

    lngColor = vsfVar.GridColor
    vsfVar.GridColor = vbBlack

    lngRow = vsfVar.Row
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTitle
        
    objRow.Add strRange
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(SYS.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfVar
    
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
            Case 1
                 zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    
    vsfVar.Row = lngRow
    vsfVar.GridColor = lngColor
End Sub

Public Function SendMessage(ByVal bytResult As Byte, ByVal lngAuditID As Long, _
    ByVal bytMode As Byte, ByRef objMIP As clsMipModule, _
    Optional ByVal blnSendBeforeAudit As Boolean = False) As Boolean
'功能：发送消息通知开方医生
'参数：
'  bytResult：1-审查合格；2-审查不合格；11-审查合格，自动发送失败
'  lngAuditID：审方ID
'  bytMode：1-门诊；2-住院
'  objMIP：消息平台对象
'  blnSendBeforeAudit：bytMode=1才生效。True-医嘱发送前审方；False-药房配发药前审方
'返回：True成功；False失败

    Const STR_OUT_PATI As String = "ZLHIS_RECIPEAUDIT_001"
    Const STR_IN_PATI  As String = "ZLHIS_RECIPEAUDIT_002"
    
    Dim strOutIn As String, strXML As String, strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objXML As New zl9ComLib.clsXML
    Dim blnMIP As Boolean
    
    If objMIP Is Nothing Then
        blnMIP = False
    Else
        blnMIP = objMIP.IsConnect
    End If

    If bytMode = 1 Then
        strOutIn = STR_OUT_PATI     '门诊
    Else
        strOutIn = STR_IN_PATI      '住院
    End If
    
    On Error GoTo errHandle
    
    'XML结构与内容
    strSQL = "Select c.病人id, c.姓名, c.住院号, c.门诊号, b.病人来源, Decode(b.病人来源, 1, d.Id, 2, b.主页id, Null) 就诊id, " & _
             "    c.当前病区id, e1.名称 当前病区, c.当前科室id, e2.名称 当前科室, c.当前床号, b.开嘱医生, " & _
             "    to_char(a.审查时间, 'yyyy-mm-dd hh24:mi:ss') 审查时间, a.审查人, a.审查结果, a.Ids " & vbNewLine & _
             "From (Select Max(a.医嘱id) 医嘱id, b.审查时间, User 审查人id, b.审查人, b.审查结果," & _
             "          f_List2str(Cast(Collect(Cast(a.医嘱id As Varchar2(20))) As t_Strlist)) Ids " & vbNewLine & _
             "      From 处方审查明细 A, 处方审查记录 B " & vbNewLine & _
             "      Where a.审方id = b.Id And a.审方id = [1] " & vbNewLine & _
             "      Group By b.审查时间, User, b.审查人, b.审查结果" & vbNewLine & _
             "     ) A, 病人医嘱记录 B, 病人信息 C, 病人挂号记录 D, 部门表 E1, 部门表 E2 " & vbNewLine & _
             "Where a.医嘱id = b.Id And b.病人id = c.病人id And b.挂号单 = d.No(+) And c.当前病区id = E1.Id(+) And c.当前病区id = E2.Id(+) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人、医嘱、审方信息", lngAuditID)
    If rsTemp.EOF = False Then
        '组装XML
        
        '病人信息
        objXML.AppendNode "patient_info"
        objXML.AppendData "patient_id", NVL(rsTemp!病人ID)
        objXML.AppendData "patient_name", NVL(rsTemp!姓名)
        objXML.AppendData "in_number", NVL(rsTemp!住院号)
        objXML.AppendData "out_number", NVL(rsTemp!门诊号)
        objXML.AppendNode "patient_info", True
        
        '医嘱信息
        objXML.AppendNode "patient_clinic"
        objXML.AppendData "patient_source", NVL(rsTemp!病人来源)
        objXML.AppendData "clinic_id", NVL(rsTemp!就诊id)
        objXML.AppendData "clinic_area_id", NVL(rsTemp!当前病区id)
        objXML.AppendData "clinic_area_title", NVL(rsTemp!当前病区)
        objXML.AppendData "clinic_dept_id", NVL(rsTemp!当前科室id)
        objXML.AppendData "clinic_dept_title", NVL(rsTemp!当前科室)
        objXML.AppendData "clinic_room", ""
        objXML.AppendData "clinic_bed", NVL(rsTemp!当前床号)
        objXML.AppendNode "patient_clinic", True
        
        '审方信息
        objXML.AppendNode "recipe_audit_info"
        objXML.AppendData "create_doctor_name", NVL(rsTemp!开嘱医生)
        objXML.AppendData "ra_time", NVL(rsTemp!审查时间)
        objXML.AppendData "ra_chemist_id", UserInfo.ID
        objXML.AppendData "ra_chemist_name", NVL(rsTemp!审查人)
        objXML.AppendData "ra_result", NVL(rsTemp!审查结果)
        objXML.AppendData "ra_sent", IIf(bytResult = 11, 1, 0)
        objXML.AppendData "order_ids", NVL(rsTemp!ids)
        objXML.AppendNode "recipe_audit_info", True
        
        strXML = objXML.XmlText
        
        objXML.ClearXmlText
        Set objXML = Nothing
    End If
    rsTemp.Close
    
    If strXML = "" Then
        MsgBox "无数据向医生发送消息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '1.门诊医生工作站“发送”前审方流程，合格与不合格消息都需要发送给医生
    '2.门诊/住院配发药审方流程，审查不合格发消息通知医生
    If blnSendBeforeAudit Then
        '通过临床的消息机制发送
        If zlDatabase.SendMsg(strOutIn, strXML) = False Then
            MsgBox "向医生发送消息失败！", vbInformation, gstrSysName
        End If
        '处方发送前开展审方
        If blnMIP Then
            '通过消息平台发送
            If objMIP.CommitMessage(strOutIn, strXML) = False Then
                MsgBox "向医生发送消息失败！", vbInformation, gstrSysName
            End If
        End If
    Else
        '配发药前开展审方
        If bytResult = 2 Then
            '通过临床的消息机制发送
            If zlDatabase.SendMsg(strOutIn, strXML) = False Then
                MsgBox "向医生发送消息失败！", vbInformation, gstrSysName
            End If
            '通过消息平台发送
            If blnMIP Then
                If objMIP.CommitMessage(strOutIn, strXML) = False Then
                    MsgBox "向医生发送消息失败！", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    
    SendMessage = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Sub PassResultView(ByRef objPASS As Object, ByVal blnOutPatient As Boolean, ByVal lngMedicalID As Long)
'功能：查看指定药嘱的PASS检查结果
'参数：
'  objPASS：PASS接口对象
'  blnOutPatient：True门诊病人；False住院病人
'  lngMedicalID：药嘱ID

    If objPASS Is Nothing Then Exit Sub
    
    Dim strSQL As String, strNO As String
    Dim lngPageID As Long
    Dim lngPatientID As Long
    Dim rsTemp As ADODB.Recordset
    
    '获取病人信息
    On Error GoTo hErr
    strSQL = "Select 病人id, 主页id, 挂号单 From 病人医嘱记录 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "通过医嘱ID，获取病人信息", lngMedicalID)
    If rsTemp.EOF = False Then
        lngPatientID = zl9ComLib.NVL(rsTemp!病人ID, 0)
        If blnOutPatient Then
            strNO = zl9ComLib.NVL(rsTemp!挂号单)
        Else
            lngPageID = zl9ComLib.NVL(rsTemp!主页ID, 0)
        End If
    End If
    rsTemp.Close
    
    '合理用药重新计算后，才能查看检查结果
    On Error Resume Next: Err.Clear
    If blnOutPatient Then
        Call objPASS.zlPassRecipelCheck(lngPatientID, 0, strNO, CStr(lngMedicalID))
    Else
        Call objPASS.zlPassRecipelCheck(lngPatientID, lngPageID, "", CStr(lngMedicalID))
    End If
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        Exit Sub
    End If
    
    '查看合理用药的检查结果
    On Error Resume Next: Err.Clear
    Call objPASS.zlPassShowWarn_YF(CStr(lngMedicalID))
    If Err.Number <> 0 Then Err.Clear
    
    Exit Sub

hErr:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Sub

Public Function GetRecipeAuditBills(ByVal bytType As Byte) As Boolean
'功能：检查最近门诊或住院的“处方审查记录”是否存在未审查的记录
'参数：
'  bytType：0-不区分门诊或住院；1-门诊；2-住院
'返回：True存在未审查的记录；False不存在未审查的记录

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    If bytType = 1 Then
        '门诊
        strSQL = "Select ID From 处方审查记录 Where 状态 = 0 And 提交时间 >= Trunc(Sysdate - [1]) And 挂号Id Is Not Null And Rownum < 2 "
    ElseIf bytType = 2 Then
        '住院
        strSQL = "Select ID From 处方审查记录 Where 状态 = 0 And 提交时间 >= Trunc(Sysdate - [1]) And 主页Id Is Not Null And Rownum < 2 "
    Else
        '不区分门诊或住院
        strSQL = "Select ID From 处方审查记录 Where 状态 = 0 And 提交时间 >= Trunc(Sysdate - [1]) And Rownum < 2 "
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查未审查的处方审查记录", IIf(bytType = 1, 2, 4))
    GetRecipeAuditBills = rsTemp.EOF = False
    rsTemp.Close
    
    Exit Function

hErr:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Function

Private Function GetMedicalID(ByVal strParentID As String) As String
'功能：将给药途径ID转换成医嘱ID
'参数：
'  strParentID：给药途径ID；格式：医嘱ID1[,医嘱ID2[,...]]
'返回：医嘱ID

    Dim strSQL As String, strResult As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    strSQL = "Select a.Id From 病人医嘱记录 A, Table(Cast(f_Str2list([1]) As t_Strlist)) B Where a.相关id = b.Column_Value"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "将给药途径ID转换成医嘱ID", strParentID)
    With rsTemp
        Do While .EOF = False
            strResult = strResult & "," & !ID
            .MoveNext
        Loop
        .Close
        If strResult <> "" Then strResult = Mid(strResult, 2)
    End With
    
    GetMedicalID = strResult
    
    Exit Function
    
hErr:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Function

Public Function GetDiagnose(ByVal lngPatientID As Long, ByVal lngPageID As Long) As String
'功能：获取住院病人的诊断
'参数：
'   lngPatientID：病人ID
'   lngPageID：主页ID
'返回：诊断内容

    Dim strSQL As String, str诊断 As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    strSQL = "Select 是否疑诊, 诊断描述 From 病人诊断记录 Where 病人id = [1] And 主页id = [2] And 诊断描述 Is Not Null Order By Rowid"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取住院病人的诊断描述", lngPatientID, lngPageID)
    With rsTemp
        Do While .EOF = False
            str诊断 = str诊断 & "," & !诊断描述 & IIf(Val(!是否疑诊 & "") = 1, "(疑)", "")
            
            .MoveNext
        Loop
        .Close
    End With
    
    If str诊断 <> "" Then GetDiagnose = Mid(str诊断, 2)
    
    Exit Function

hErr:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Function

Public Sub DispCountNG(ByVal vsfCount As VSFlexGrid, ByVal lblDisp As Label)
'功能：计算控件中不合格的审查项目数量，并显示在Label上
'参数：
'  vsfCount：要计算的控件
'  lblDisp：要显示的控件

    Const STR_ITEMS As String = "审查项目"
    Dim i As Integer, intCount As Integer
    
    With vsfCount
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("药师审查"))) = "不合格" Then
                intCount = intCount + 1
            End If
        Next
    End With
    If intCount > 0 Then
        lblDisp = STR_ITEMS & zlStr.FormatString("（共有[1]项不合格）", intCount)
    Else
        lblDisp = STR_ITEMS
    End If

End Sub

Public Function GetAuditResult(ByVal vsfVar As VSFlexGrid) As Boolean
'功能：检查VSF审查项目中“药量审查”结果是否所有项目都合格
'返回：True所有项目都合格；False有不合格

    Const STR_NAME As String = "药师审查"
    Const STR_PASS As String = "合格"
    Dim i As Integer
    Dim blnAllPass As String
    
    With vsfVar
        blnAllPass = True
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex(STR_NAME)) <> STR_PASS Then
                blnAllPass = False
                Exit For
            End If
        Next
    End With
    
    GetAuditResult = blnAllPass
End Function

