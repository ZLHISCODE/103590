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
Private mfrmDisease As frmDockDisease
Private mObjTabEpr As cTableEPR            '表格式病历编辑器
Private mObjTabEprView As cTableEPR
Private mfrmDiseaseFileList As frmDiseaseFileList
Private mfrmEPRFileApplyTo As frmEPRFileApplyTo
Private mfrmEPRFileDisease As frmEPRFileDisease
Private WithEvents mfrmEPRFileContent As frmEPRFileContent
Attribute mfrmEPRFileContent.VB_VarHelpID = -1
Private WithEvents mfrmEPRFileRequest As frmEPRFileRequest
Attribute mfrmEPRFileRequest.VB_VarHelpID = -1

Private mfrmEPRFileEdit As frmEPRFileEdit
Private mfrmElementChange As frmElementChange
Private mfrmFileExportOrImport As frmFileExportOrImport
Private mfrmPrivacyProtect As frmPrivacyProtect

Public Event EPRFileContentDblClick()
Public Event EPRFileRequestDblClick(lngWhere As Integer)

Private Sub Class_initialize()
    Set mfrmDisease = New frmDockDisease
    Set mfrmDiseaseFileList = New frmDiseaseFileList
    Set mfrmEPRFileEdit = New frmEPRFileEdit
    Set mfrmElementChange = New frmElementChange
    Set mfrmFileExportOrImport = New frmFileExportOrImport
    Set mfrmPrivacyProtect = New frmPrivacyProtect
    Set mfrmEPRFileContent = New frmEPRFileContent
    Set mfrmEPRFileRequest = New frmEPRFileRequest
    Set mfrmEPRFileApplyTo = New frmEPRFileApplyTo
    Set mfrmEPRFileDisease = New frmEPRFileDisease
End Sub

Private Sub Class_Terminate()
    Unload mfrmDisease
    Set mfrmDisease = Nothing
    Unload mfrmDiseaseFileList
    Set mfrmDiseaseFileList = Nothing
    Unload mfrmEPRFileEdit
    Set mfrmEPRFileEdit = Nothing
    Unload mfrmElementChange
    Set mfrmElementChange = Nothing
    Unload mfrmFileExportOrImport
    Set mfrmFileExportOrImport = Nothing
    Unload mfrmPrivacyProtect
    Set mfrmPrivacyProtect = Nothing
    Unload mfrmEPRFileContent
    Set mfrmEPRFileContent = Nothing
    Unload mfrmEPRFileRequest
    Set mfrmEPRFileRequest = Nothing
    Unload mfrmEPRFileApplyTo
    Set mfrmEPRFileApplyTo = Nothing
    Unload mfrmEPRFileDisease
    Set mfrmEPRFileDisease = Nothing
End Sub
Public Function zlGetForm() As Object
    Set zlGetForm = mfrmDisease
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object)
    Call mfrmDisease.zlDefCommandBars(CommandBars)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmDisease.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmDisease.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngFrom As Long, ByVal lngDeptId As Long, ByVal blnMoved As Boolean, Optional ByVal blnEdit As Boolean, Optional ByVal intState As Integer) As Long
'lngPageId 住院传主页ID，门诊传挂号ID
    zlRefresh = mfrmDisease.zlRefresh(lngPatiID, lngPageId, lngFrom, lngDeptId, blnMoved, blnEdit, intState)
End Function


Public Function SatisfyCondition(ByVal lngPatiID As Long, ByVal lngPageId As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    strSQL = "Select id From 电子病历记录 Where 病人id =[1] And 主页id =[2] And 编辑方式 In (0,1,2)  And 病历种类 = 5"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询当前是否已有传染病记录", lngPatiID, lngPageId)
    If rsTemp.RecordCount > 0 Then
        SatisfyCondition = True
    Else
        gstrSQL = "Select b.文件id From 病人诊断记录 A, 疾病报告前提 B,病历文件列表 C Where a.病人id = [1] And a.主页id = [2] And a.疾病id = b.疾病id AND B.文件ID=C.ID AND C.种类=5" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select b.文件id From 病人诊断记录 A, 疾病报告前提 B,病历文件列表 C Where a.病人id = [1] And a.主页id = [2] And a.诊断id = b.诊断id AND B.文件ID=C.ID AND C.种类=5"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询是否符合传染病上报", lngPatiID, lngPageId)
        If rsTemp.RecordCount > 0 Then
            SatisfyCondition = True
        End If
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Public Function EditDiseaseDoc(ByVal frmPar As Object, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long, ByVal strDiseaseIDs As String, ByVal strDiagnoseIDs As String, Optional ByRef blnNo As Boolean) As Boolean
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHand
    
    If strDiseaseIDs = "" And strDiagnoseIDs = "" Then
        Exit Function
    End If
    
    If InStr(";" & GetPrivFunc(glngSys, 1249) & ";", ";病历书写;") <= 0 Then
        Exit Function
    End If
    
    Set rsTemp = SatisfyEditDiseaseDoc(lngPatiID, lngPageId, lngDeptId, strDiseaseIDs, strDiagnoseIDs)
     
    If rsTemp Is Nothing Then
        Exit Function
    ElseIf rsTemp.RecordCount = 0 Then
        Exit Function
    End If
   
    EditDiseaseDoc = EditDiseaseReport(frmPar, rsTemp, lngPatiID, lngPageId, lngPatFrom, lngDeptId, blnNo)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SatisfyEditDiseaseDoc(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal strDiseaseIDs As String, ByVal strDiagnoseIDs As String) As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHand
    If strDiseaseIDs = "" And strDiagnoseIDs = "" Then
        Exit Function
    End If
        
    gstrSQL = ""
    If strDiseaseIDs <> "" Then
        gstrSQL = " Union Select 文件ID From 疾病报告前提 Where 疾病ID IN (Select Column_Value From Table(f_Num2list([3])))"
    End If
    If strDiagnoseIDs <> "" Then
        gstrSQL = gstrSQL & " Union Select 文件ID From 疾病报告前提 Where 诊断ID IN (Select Column_Value From Table(f_Num2list([4])))"
    End If
    gstrSQL = "(" & Mid(gstrSQL, 8) & ") Minus Select 文件ID From 电子病历记录 Where 病人ID=[1] And 主页ID=[2] And 病历种类=5"
    gstrSQL = "Select /*+ Rule*/" & vbNewLine & _
            " a.Id, a.种类, a.编号, a.名称, a.保留, a.说明" & vbNewLine & _
            "From 病历文件列表 A ,(" & gstrSQL & ") B Where A.ID=B.文件ID  And" & vbNewLine & _
            "(a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 病历应用科室 C Where c.文件id = a.Id And c.科室id = [5]))" & vbNewLine & _
            "Order By a.编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "CheckDiseaseFile", lngPatiID, lngPageId, strDiseaseIDs, strDiagnoseIDs, lngDeptId)
    
    Set SatisfyEditDiseaseDoc = rsTemp
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function EditDiseaseReport(ByVal frmPar As Object, ByVal rsTemp As ADODB.Recordset, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long, Optional ByRef blnNo As Boolean) As Boolean
    Dim lngFileID As Long, strMsg As String
    Dim objDoc As cEPRDocument, objInfection As Object
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
                blnNo = True
                Exit Function
            End If
        End If
Process:    lngFileID = Val(rsTemp!ID & "")
        If rsTemp!保留 = 4 Then         '新编辑器
            Set objInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "传染病报告卡", True)
            If Not objInfection Is Nothing Then
                objInfection.Init gcnOracle, glngSys
                objInfection.OpenDoc frmPar, cprEM_新增, lngPatiID, lngPageId, lngPatFrom, 0, lngDeptId, lngFileID
            End If
        ElseIf rsTemp!保留 = 2 Then     '表格编辑器
            Dim bFinded As Boolean
    
            Set mObjTabEprView = New cTableEPR
            Call mObjTabEprView.InitTableEPR(gcnOracle, glngSys, gstrDbOwner)
             

            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(lngFileID, lngPatiID, lngPageId, lngPatFrom, lngDeptId)
            End If
            If Not bFinded Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.InitOpenEPR frmPar, cprEM_新增, cprET_单病历编辑, lngFileID, True, 0, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, GetPrivFunc(glngSys, IIf(lngPatFrom = cprPF_住院, 1251, 1250)), , InStr(gstrPrivsEpr, "病历打印") > 0, Val(gstrESign)
            End If

        Else
            Set objDoc = New cEPRDocument
            Call objDoc.InitEPRDoc(cprEM_新增, cprET_单病历编辑, lngFileID, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, False)
            Call objDoc.ShowEPREditor(frmPar, , vbModal)
        End If
    ElseIf rsTemp.RecordCount > 1 Then
        If frmDiseaseFileList.ShowMe(frmPar, rsTemp, lngFileID) Then
            rsTemp.Filter = "ID=" & lngFileID
            GoTo Process
        Else
            blnNo = True
        End If
    End If
    
    EditDiseaseReport = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
'-0-小(缺省)，1-大
    Call mfrmDisease.SetFontSize(bytSize)
End Sub

Public Function zlGetFrmDiseaseFileList() As Object
'功能：获取窗口对象
    Set zlGetFrmDiseaseFileList = mfrmDiseaseFileList
End Function

Public Function zlGetFrmEPRFileEdit() As Object
'功能：获取窗口对象
    Set zlGetFrmEPRFileEdit = mfrmEPRFileEdit
End Function

Public Function zlGetFrmElementChange() As Object
'功能：获取窗口对象
    Set zlGetFrmElementChange = mfrmElementChange
End Function

Public Function zlGetFrmFileExportOrImport() As Object
'功能：获取窗口对象
    Set zlGetFrmFileExportOrImport = mfrmFileExportOrImport
End Function

Public Function zlGetFrmPrivacyProtect() As Object
'功能：获取窗口对象
    Set zlGetFrmPrivacyProtect = mfrmPrivacyProtect
End Function

Public Function zlGetFrmEPRFileContent() As Object
'功能：获取窗口对象
    Set zlGetFrmEPRFileContent = mfrmEPRFileContent
End Function

Private Sub mfrmEPRFileContent_DblClick()
    RaiseEvent EPRFileContentDblClick
End Sub

Public Function zlGetFrmEPRFileRequest() As Object
'功能：获取窗口对象
    Set zlGetFrmEPRFileRequest = mfrmEPRFileRequest
End Function

Private Sub mfrmEPRFileRequest_DblClick(lngWhere As zlEnumDClick)
    Dim i As Integer
    i = CInt(lngWhere)
    RaiseEvent EPRFileRequestDblClick(i)
End Sub

Public Function zlGetFrmEPRFileApplyTo() As Object
'功能：获取窗口对象
    Set zlGetFrmEPRFileApplyTo = mfrmEPRFileApplyTo
End Function

Public Function zlGetFrmEPRFileDisease() As Object
'功能：获取窗口对象
    Set zlGetFrmEPRFileDisease = mfrmEPRFileDisease
End Function
