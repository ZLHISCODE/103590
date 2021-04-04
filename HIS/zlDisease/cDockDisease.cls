VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function InitDockDisease(ByVal lngSys As Long, ByVal cnOracle As ADODB.Connection) As Boolean
    On Error GoTo errH
    glngSys = lngSys: Set gcnOracle = cnOracle
    If zlGetComLib = False Then Exit Function
    InitDockDisease = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function SatisfyCondition(ByVal lngPatiID As Long, ByVal lngPageId As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHand
    
    strSQL = "Select b.文件id From 病人诊断记录 A, 疾病报告前提 B,病历文件列表 C Where a.病人id = [1] And a.主页id = [2] And a.疾病id = b.疾病id AND B.文件ID=C.ID AND C.保留=4" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select b.文件id From 病人诊断记录 A, 疾病报告前提 B,病历文件列表 C Where a.病人id = [1] And a.主页id = [2] And a.诊断id = b.诊断id AND B.文件ID=C.ID AND C.保留=4"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询是否符合传染病上报", lngPatiID, lngPageId)
    If rsTemp.RecordCount > 0 Then
        SatisfyCondition = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function EditDiseaseDoc(ByVal frmPar As Object, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long, ByVal strDiseaseIDs As String, ByVal strDiagnoseIDs As String, Optional ByRef strReason As String) As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHand
    If strDiseaseIDs = "" And strDiagnoseIDs = "" Then Exit Function
 
    Set rsTemp = SatisfyEditDiseaseDoc(lngPatiID, lngPageId, lngPatFrom, lngDeptId, strDiseaseIDs, strDiagnoseIDs)
    EditDiseaseDoc = EditDiseaseReport(frmPar, rsTemp, lngPatiID, lngPageId, lngPatFrom, lngDeptId, strDiseaseIDs, strDiagnoseIDs, strReason)
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SatisfyEditDiseaseDoc(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long, ByVal strDiseaseIDs As String, ByVal strDiagnoseIDs As String) As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If strDiseaseIDs = "" And strDiagnoseIDs = "" Then
        Exit Function
    End If

    If strDiseaseIDs <> "" Then
        strSQL = " Union Select 文件ID From 疾病报告前提 Where 疾病ID IN (Select Column_Value From Table(f_Num2list([3])))"
    End If
    If strDiagnoseIDs <> "" Then
        strSQL = strSQL & " Union Select 文件ID From 疾病报告前提 Where 诊断ID IN (Select Column_Value From Table(f_Num2list([4])))"
    End If
    strSQL = Mid(strSQL, 8)
    strSQL = "Select /*+ Rule*/ distinct " & vbNewLine & _
            " a.Id, a.种类, a.编号, a.名称, a.保留, a.说明" & vbNewLine & _
            "From 病历文件列表 A ,(" & strSQL & ") B Where A.ID=B.文件ID  And" & vbNewLine & _
            "(a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 病历应用科室 C Where c.文件id = a.Id And c.科室id = [5]))" & vbNewLine & _
            "Order By a.编号"
            
    Set SatisfyEditDiseaseDoc = zlDatabase.OpenSQLRecord(strSQL, "CheckDiseaseFile", lngPatiID, lngPageId, strDiseaseIDs, strDiagnoseIDs, lngDeptId)

    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function EditDiseaseReport(ByVal frmPar As Object, ByVal rsTemp As ADODB.Recordset, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long, ByVal strDiseaseIDs As String, ByVal strDiagnoseIDs As String, Optional ByRef strReason As String) As Boolean
    Dim lngFileID As Long, strMsg As String
    Dim objDoc As cEPRDocument, objInfection As Object
    Dim strSQL As String
    Dim objDisease As zlRichEPR.cDockDisease
    Dim objDiseaseFileList As Object
    Dim rsDis As ADODB.Recordset
    Dim blnFinded As Boolean
    Dim ObjTabEpr As cTableEPR            '表格式病历编辑器
    Dim ObjTabEprView As cTableEPR
    
    On Error GoTo errHand
    
    If rsTemp Is Nothing Then
        Exit Function
    ElseIf rsTemp.RecordCount = 0 Then
        Exit Function
    End If

    Do Until rsTemp.EOF
        strMsg = strMsg & vbCrLf & "【" & rsTemp!名称 & "】"
        rsTemp.MoveNext
    Loop
    rsTemp.MoveFirst

    If rsTemp.RecordCount = 1 Then
            If zlDatabase.GetPara("首页整理后编辑疾控报告卡", glngSys, 1277, "0") = 1 And rsTemp!保留 = 4 Then  '＝0表示提醒 =1表示弹出编辑
            Else
                If MsgBox("根据病人的诊断信息，以下疾病证明报告还没有填写：" & vbCrLf & vbCrLf & Mid(strMsg, 3) & vbCrLf & vbCrLf & "是否立即填写？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If CheckDisNum(lngPatiID, lngPageId, lngPatFrom, rsDis, glngOpenedID) Then
                        strReason = frmReason.ShowMe(frmPar, Nvl(rsDis!处理情况说明))
                        Call SaveReason(strReason, rsDis!id, rsDis!记录状态)
                    End If
                    Exit Function
                End If
            End If
Process:    lngFileID = Val(rsTemp!id & "")
        If rsTemp!保留 = 4 Then '新编辑器
            Set objInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "传染病报告卡", True)
            If Not objInfection Is Nothing Then
                objInfection.Init gcnOracle, glngSys
                objInfection.OpenDoc frmPar, cprEM_新增, lngPatiID, lngPageId, lngPatFrom, 0, lngDeptId, lngFileID
            End If
        ElseIf rsTemp!保留 = 2 Then '表格编辑器
            Set ObjTabEprView = New cTableEPR
            Call ObjTabEprView.InitTableEPR(gcnOracle, glngSys, gstrDBOwer)
            
            If Not ObjTabEpr Is Nothing Then
                blnFinded = ObjTabEpr.Showfrm(lngFileID, lngPatiID, lngPageId, lngPatFrom, lngDeptId)
            End If
            If Not blnFinded Then
                Set ObjTabEpr = New cTableEPR
                ObjTabEpr.InitOpenEPR frmPar, cprEM_新增, cprET_单病历编辑, lngFileID, True, 0, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, GetPrivFunc(glngSys, IIf(lngPatFrom = cprPF_住院, 1251, 1250)), , , , 1
            End If
        Else
            Set objDoc = New cEPRDocument
            Call objDoc.InitEPRDoc(cprEM_新增, cprET_单病历编辑, lngFileID, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, False)
            Call objDoc.ShowEPREditor(frmPar, , vbModal)
        End If
    ElseIf rsTemp.RecordCount > 1 Then
        Set objDisease = New zlRichEPR.cDockDisease

        If objDisease.zlGetFrmDiseaseFileList.ShowMe(frmPar, rsTemp, lngFileID) Then
            rsTemp.Filter = "ID=" & lngFileID
            GoTo Process
        Else
            If CheckDisNum(lngPatiID, lngPageId, lngPatFrom, rsDis, glngOpenedID) Then
                strReason = frmReason.ShowMe(frmPar, Nvl(rsDis!处理情况说明))
                Call SaveReason(strReason, rsDis!id, rsDis!记录状态)
            End If
        End If
        Set objDisease = Nothing
    End If
    EditDiseaseReport = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function RelateFeedback(ByRef frmParent As Object, ByVal lngID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngFrom As Long, ByVal isRelated As Boolean) As Boolean
'功能：传染病报告卡，关联阳性结果反馈单，或者取消关联
'参数：lngID -报告ID
'      lngPatiID -病人ID
'      lngPageID -病人主页ID或者挂号ID
'      lngFrom-病人来源，1-门诊，2-住院
'      isRelated  true-关联；false-取消关联
    Dim strSQL As String
    Dim rsDisease As ADODB.Recordset
    Dim strIDs As String
    Dim arrayID() As String
    Dim i As Long
    Dim frmFBList As frmFeedBackList

On Error GoTo errHand
  
    If isRelated Then   '关联
        If lngFrom = 1 Then
            strSQL = "select rowNum as NO,a.ID,c.名称 as 科室, a.登记时间 from  疾病阳性记录 A ,病人挂号记录 B ,部门表 C where A.文件ID is NULL  and A.挂号单 = B.NO and A.病人ID = B.病人ID and A.记录状态 <> 3 and A.登记科室ID = C.ID  and A.病人ID = [1] and B.ID = [2]"
        ElseIf lngFrom = 2 Then
            strSQL = "select rowNum as NO,a.ID ,c.名称 as 科室,a.登记时间 from  疾病阳性记录 A ,部门表 C  where A.文件ID is NULL  and A.记录状态 <> 3  and A.登记科室ID = C.ID and A.病人ID = [1] and A.主页ID = [2] "
        End If
        Set rsDisease = zlDatabase.OpenSQLRecord(strSQL, "查询该报告对应的阳性结果反馈单", lngPatiID, lngPageId)
        If rsDisease.RecordCount = 1 Then
            strSQL = "Zl_疾病阳性检测记录_Update(2," & rsDisease!id & "," & lngID & ",NULL,NULL,NULL,NULL)"
            Call zlDatabase.ExecuteProcedure(strSQL, "关联疾病报告和阳性结果反馈单")
        ElseIf rsDisease.RecordCount > 1 Then
            Set frmFBList = New frmFeedBackList
            If frmFBList.ShowMe(frmParent, rsDisease, strIDs) Then
                If strIDs <> "" Then
                    arrayID = Split(strIDs, ",")
                    For i = LBound(arrayID) To UBound(arrayID)
                        If Val(arrayID(i)) <> 0 Then
                            strSQL = "Zl_疾病阳性检测记录_Update(2," & arrayID(i) & "," & lngID & ",NULL,NULL,NULL,NULL)"
                            Call zlDatabase.ExecuteProcedure(strSQL, "关联疾病报告和阳性结果反馈单")
                        End If
                    Next
                End If
            End If
            If Not frmFBList Is Nothing Then
                Unload frmFBList
                Set frmFBList = Nothing
            End If
        End If
    Else '取消关联
        strSQL = "Zl_疾病阳性检测记录_Update(3, NULL " & "," & lngID & ",NULL,NULL,NULL,NULL)"
        Call zlDatabase.ExecuteProcedure(strSQL, "取消疾病报告和阳性结果反馈单的关联")
    End If
    
    RelateFeedback = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

