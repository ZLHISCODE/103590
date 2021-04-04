VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmChildDocumentView 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3525
      Index           =   0
      Left            =   210
      ScaleHeight     =   3525
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   405
      Width           =   4815
      Begin XtremeSuiteControls.TabControl tbcSub 
         Height          =   1830
         Left            =   585
         TabIndex        =   1
         Top             =   630
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
End
Attribute VB_Name = "frmChildDocumentView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mfrmMain As Object
Private mlngKey As Long
Private mlngReferKey As Long
Private mblnReading As Boolean
Private mstrSQL As String
Private mbytMode As Byte
Private mstrObject As String
Private mstrParam As String
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlngRecordKey As Long
Private mblnPrinted As Boolean
Private mlngNo As Long
Private mblnNewTends As Boolean
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1

Private mclsArchiveMedRec As zlMedRecPage.clsArchive
Private mfrmArchiveMedRec As Object

'Private mfrmChildMedrec As frmChildMedrec
Private WithEvents mclsInAdvices As zlCISKernel.clsDockInAdvices
Attribute mclsInAdvices.VB_VarHelpID = -1
Private WithEvents mclsDockAduits As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1
Private WithEvents mclsPath As zlCISPath.clsDockPath
Attribute mclsPath.VB_VarHelpID = -1
Private WithEvents mclsTendsNew As zl9TendFile.clsTendFile    '新版护士工作站
Attribute mclsTendsNew.VB_VarHelpID = -1
Private mblnTrans As Boolean
Private mobjRichEMR As Object
Private mobjPACSDoc As Object
Public Event AfterDataChanged()

Public Function zlPrintDocument(ByVal cbrControls As CommandBars, ByVal bytMode As Byte, Optional ByVal strPrintDeviceName As String, Optional ByVal lngNo As Long, Optional ByVal blnDoctorAdvice As Boolean = False) As Boolean
    Dim varParam As Variant
    Dim bytRet As Byte
    Dim intSel As Integer
    Dim strSQL As String
    Dim strName As String
    Dim lng科室ID As Long
    
    Set mobjReport = New clsReport
   
    mblnPrinted = False
    
    Select Case mstrObject
    Case "首页记录"
    
        If mobjReport Is Nothing Then Set mobjReport = New clsReport
        
        lng科室ID = GetlngID(mlng病人ID, mlng主页ID)
        
        Select Case Val(zlDatabase.GetPara("病案首页标准", glngSys, 1261, "0"))
        Case 0 '卫生部标准
            If Have部门性质(lng科室ID, "中医科") Then
                strName = "ZL1_INSIDE_1261_4"
            Else
                strName = "ZL1_INSIDE_1261_1"
            End If
        Case 1    '四川省标准
            If Have部门性质(lng科室ID, "中医科") Then
                strName = "ZL1_INSIDE_1261_6"
            Else
                strName = "ZL1_INSIDE_1261_5"
            End If
        Case 2    '云南省标准
            If Have部门性质(lng科室ID, "中医科") Then
                strName = "ZL1_INSIDE_1261_8"
            Else
                strName = "ZL1_INSIDE_1261_7"
            End If
        End Select
        
        
        Call mobjReport.ReportOpen(gcnOracle, ParamInfo.系统号, strName, mfrmMain, "病人id=" & mlng病人ID, "主页id=" & mlng主页ID, bytMode)
        If mblnPrinted Then Call RecordEprPrintInfo(2, "首页记录", mlngNo, mlng病人ID, mlng主页ID)
        
    Case "住院医嘱"
        If blnDoctorAdvice Then
            '打印长嘱
            Call gobjKernel.zlPrintAdvice(Me, mlng病人ID, mlng主页ID, 0, 0)
            
            '打印临嘱
            Call gobjKernel.zlPrintAdvice(Me, mlng病人ID, mlng主页ID, 0, 1)
        Else
            Call mobjReport.ReportOpen(gcnOracle, ParamInfo.系统号, "ZL1_INSIDE_1560", mfrmMain, "病人id=" & mlng病人ID, "主页id=" & mlng主页ID, bytMode)
        End If
        
        If mblnPrinted Then Call RecordEprPrintInfo(2, "住院医嘱", mlngNo, mlng病人ID, mlng主页ID)
        
    Case "住院病历"
                
        If mstrParam <> "" Then
            If IsNumeric(Split(mstrParam, ";")(0)) Then
                Call mclsDockAduits.zlPrintDocument(3, bytMode)
                If mblnPrinted Then Call RecordEprPrintInfo(1, mlngRecordKey, mlngNo)
            Else
                If Not mobjRichEMR Is Nothing Then
                    Call mobjRichEMR.zlPrintDoc(bytMode = 1)
                End If
            End If
        End If
        
    Case "护理病历"
                
        If mstrParam <> "" Then
            If IsNumeric(Split(mstrParam, ";")(0)) Then
                Call mclsDockAduits.zlPrintDocument(3, bytMode)
                If mblnPrinted Then Call RecordEprPrintInfo(1, mlngRecordKey, mlngNo)
            Else
                If Not mobjRichEMR Is Nothing Then
                    mobjRichEMR.zlPrintDoc (False)
                End If
            End If
        End If
        
    Case "知情文件"

        If mstrParam <> "" Then
            If IsNumeric(Split(mstrParam, ";")(0)) Then
                Call mclsDockAduits.zlPrintDocument(3, bytMode)
                If mblnPrinted Then Call RecordEprPrintInfo(1, mlngRecordKey, mlngNo)
            Else
                If Not mobjRichEMR Is Nothing Then
                    Call mobjRichEMR.zlPrintDoc(bytMode = 1)
                End If
            End If
        End If
        
    Case "疾病证明"
        
         If mstrParam <> "" Then
            If IsNumeric(Split(mstrParam, ";")(0)) Then
                Call mclsDockAduits.zlPrintDocument(3, bytMode)
                If mblnPrinted Then Call RecordEprPrintInfo(1, mlngRecordKey, mlngNo)
            Else
                If Not mobjRichEMR Is Nothing Then
                    Call mobjRichEMR.zlPrintDoc(bytMode = 1)
                End If
            End If
        End If
        
    Case "医嘱报告"

        If mstrParam <> "" Then
                    If Split(mstrParam, ";")(0) <> 0 Then
                                Call mclsDockAduits.zlPrintDocument(4, bytMode, mlngRecordKey, strPrintDeviceName)
                                If mblnPrinted Then Call RecordEprPrintInfo(1, mlngRecordKey, mlngNo)
            Else
                If Not mobjPACSDoc Is Nothing Then
                    Call mobjPACSDoc.PrintReport(Split(mstrParam, ";")(2), strPrintDeviceName)
                End If
            End If
        End If
        
    Case "护理记录"
        
        If mstrParam <> "" Then
            mblnNewTends = Get新版护理(mlng病人ID, mlng主页ID)
            If mblnNewTends = False Then
                varParam = Split(mstrParam, ";")
                If UBound(varParam) >= 1 Then
                    If Val(varParam(1)) = -1 Then
                        bytRet = mclsDockAduits.zlPrintDocument(1, bytMode, , strPrintDeviceName)
                        
                        If bytRet = 2 Or bytMode = 2 Then
                           Call RecordEprPrintInfo(2, "体温单", mlngNo, mlng病人ID, mlng主页ID)
                        End If
                        
                    Else
                        Call mclsDockAduits.zlPrintDocument(2, bytMode, , strPrintDeviceName)
                        If bytMode = 2 Then
                            Call RecordEprPrintInfo(3, Val(varParam(3)), mlngNo, mlng病人ID, mlng主页ID)
                        End If
                    End If
                End If
            Else
                '新版护理打印、预览
                '此参数保存 保留
                varParam = Split(mstrParam, ";")
                    If UBound(varParam) >= 1 Then
                    
                    Select Case Val(varParam(1))
                        Case -1 '体温单
                            intSel = 1
                        Case 1  '产程图
                            intSel = 3
                        Case Else '记录单
                            intSel = 2
                    End Select
                    Call mclsTendsNew.zlPrintTendFile(intSel, bytMode, strPrintDeviceName)
                End If
            End If
        End If
    Case "临床路径"
        If bytMode = 1 Then
            Call mclsPath.zlExecuteCommandBars(cbrControls.FindControl(, conMenu_File_Preview))
        Else
            Call mclsPath.zlExecuteCommandBars(cbrControls.FindControl(, conMenu_File_Print))
        End If
    End Select
    
    
End Function


Public Function zlPrintSet(ByVal Control As CommandBarControl, Optional ByVal blnDoctorAdvice As Boolean = False) As Boolean
    Set mobjReport = New clsReport
    Select Case mstrObject
    Case "首页记录"
        Call mobjReport.ReportPrintSet(gcnOracle, ParamInfo.系统号, "ZL1_INSIDE_1261_1", Me)
    Case "住院医嘱"
        Call mobjReport.ReportPrintSet(gcnOracle, ParamInfo.系统号, "ZL1_INSIDE_1560", Me)
    Case "住院病历"
    Case "护理病历"
    Case "知情文件"
    Case "疾病证明"
    Case "医嘱报告"
    Case "护理记录"
    Case "临床路径"
        Control.ID = 101
        Call mclsPath.zlExecuteCommandBars(Control)
    End Select
    
End Function

Public Function zlInitData(ByVal frmMain As Object) As Boolean
    Set mfrmMain = frmMain
    zlInitData = InitControl
End Function

Public Function zlRefresh(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strObject As String, ByVal strParam As String, ByVal strCaption As String, ByVal blnDataMoved As Boolean) As Boolean
    Dim varParam As Variant
    Dim intSel As Integer
    
    mstrObject = strObject
    mstrParam = strParam
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlngNo = 0
 
    Select Case strObject
    Case "首页记录"
        
        Call mclsArchiveMedRec.zlRefresh(1, lng病人ID, lng主页ID, False)
        Call ShowArchiveTab(strObject, strCaption)
        
    Case "住院医嘱"
        
        Call mclsInAdvices.zlRefresh(lng病人ID, lng主页ID, 0, 0, 1, blnDataMoved, 0, 0)
        Call ShowArchiveTab(strObject, strCaption)
                
    Case "住院病历"
                
        If strParam <> "" Then
            If IsNumeric(Split(strParam, ";")(0)) Then
                varParam = Split(strParam, ";")
                mclsDockAduits.ParentForm mfrmMain
                mlngRecordKey = Val(varParam(0))
                Call mclsDockAduits.zlRefresh(2, Val(varParam(0)), , , , , , , blnDataMoved)
                Call ShowArchiveTab(strObject, strCaption)
            ElseIf strParam <> "" Then
                'EMR病历预览
                If Not mobjRichEMR Is Nothing Then
                    If InStr(strParam, "|") > 0 Then
                        Call mobjRichEMR.zlShowDoc(Split(strParam, "|")(0), Split(strParam, "|")(1))
                    Else
                        Call mobjRichEMR.zlShowDoc(strParam, "")
                    End If
                End If
                Call ShowArchiveTab("电子病历", strCaption)
            End If
        End If
        
    Case "护理病历"
                
        If strParam <> "" Then
            If IsNumeric(Split(strParam, ";")(0)) Then
                varParam = Split(strParam, ";")
                mclsDockAduits.ParentForm mfrmMain
                mlngRecordKey = Val(varParam(0))
                Call mclsDockAduits.zlRefresh(4, Val(varParam(0)), , , , , , , blnDataMoved)
                Call ShowArchiveTab("住院病历", strCaption)
            ElseIf strParam <> "" Then
                'EMR病历预览
                If Not mobjRichEMR Is Nothing Then
                    If InStr(strParam, "|") > 0 Then
                        Call mobjRichEMR.zlShowDoc(Split(strParam, "|")(0), Split(strParam, "|")(1))
                    Else
                        Call mobjRichEMR.zlShowDoc(strParam, "")
                    End If
                End If
                Call ShowArchiveTab("电子病历", strCaption)
            End If
        End If
        
    Case "知情文件"
        
        If strParam <> "" Then
            If IsNumeric(Split(strParam, ";")(0)) Then
                varParam = Split(strParam, ";")
                mclsDockAduits.ParentForm mfrmMain
                mlngRecordKey = Val(varParam(0))
                Call mclsDockAduits.zlRefresh(6, Val(varParam(0)), , , , , , , blnDataMoved)
                Call ShowArchiveTab("住院病历", strCaption)
            ElseIf strParam <> "" Then
                'EMR病历预览
                If Not mobjRichEMR Is Nothing Then
                    If InStr(strParam, "|") > 0 Then
                        Call mobjRichEMR.zlShowDoc(Split(strParam, "|")(0), Split(strParam, "|")(1))
                    Else
                        Call mobjRichEMR.zlShowDoc(strParam, "")
                    End If
                End If
                Call ShowArchiveTab("电子病历", strCaption)
            End If
        End If
        
    Case "疾病证明"
        
        If strParam <> "" Then
            If IsNumeric(Split(strParam, ";")(0)) Then
                varParam = Split(strParam, ";")
                mclsDockAduits.ParentForm mfrmMain
                mlngRecordKey = Val(varParam(0))
                Call mclsDockAduits.zlRefresh(5, Val(varParam(0)), , , , , , , blnDataMoved)
                Call ShowArchiveTab("住院病历", strCaption)
            ElseIf strParam <> "" Then
                'EMR病历预览
                If Not mobjRichEMR Is Nothing Then
                    If InStr(strParam, "|") > 0 Then
                        Call mobjRichEMR.zlShowDoc(Split(strParam, "|")(0), Split(strParam, "|")(1))
                    Else
                        Call mobjRichEMR.zlShowDoc(strParam, "")
                    End If
                End If
                Call ShowArchiveTab("电子病历", strCaption)
            End If
        End If
    
    Case "医嘱报告"

        If strParam <> "" Then
            varParam = Split(strParam, ";")
            mclsDockAduits.ParentForm mfrmMain
            mlngRecordKey = Val(varParam(0))
            If mlngRecordKey <> 0 Then
                                Call mclsDockAduits.zlRefresh(7, Val(varParam(0)), , , , , , , blnDataMoved)
                                Call ShowArchiveTab("住院病历", strCaption)
            Else
                Call mobjPACSDoc.zlDocRefresh(varParam(2)) '新PACS报告编辑器仅有病人医嘱报告记录 参数='0;医嘱ID;检查报告ID'
                Call ShowArchiveTab("检查报告", strCaption)
            End If
        End If
    
    Case "护理记录"
        If strParam <> "" Then
            mblnNewTends = Get新版护理(lng病人ID, lng主页ID)
            If mblnNewTends = False Then
                varParam = Split(strParam, ";")
                If UBound(varParam) >= 1 Then
                    If Val(varParam(1)) = -1 Then
                        mlngRecordKey = Val(varParam(0))
                        Call mclsDockAduits.zlRefreshTendBody(lng病人ID, lng主页ID, Val(Split(varParam(0), "_")(0)), Val(varParam(4)), blnDataMoved)
                        Call ShowArchiveTab("体温记录单", strCaption)
                    Else
                        mlngRecordKey = Val(varParam(3))
                        Call mclsDockAduits.zlRefresh(3, Val(varParam(3)), lng病人ID, lng主页ID, Val(Split(varParam(0), "_")(0)), CStr(varParam(2)), , Val(varParam(4)), blnDataMoved)
                        Call ShowArchiveTab("护理记录单", strCaption)
                    End If
                End If
            Else
                '此参数保存 保留
                varParam = Split(strParam, ";")
                If UBound(varParam) >= 1 Then
                    Select Case Val(varParam(1))
                        Case -1 '体温单
                            intSel = 0
                        Case 1  '产程图
                            intSel = 2
                        Case Else '记录单
                            intSel = 1
                    End Select
                    Call mclsTendsNew.zlRefreshTendFile(mlng病人ID, lng主页ID, Val(varParam(4)), Val(varParam(0)), False, False, intSel, Val(varParam(3)), 1)
                    Call ShowArchiveTab("新版护理", strCaption)
                End If
            End If
        End If
    Case "临床路径"
        Call mclsPath.zlRefreshReadOnly(lng病人ID, lng主页ID)
        Call ShowArchiveTab(strObject, strCaption)
        
    Case Else

    End Select

    zlRefresh = True
    
End Function

Public Function zlMediAudit(ByVal CommandBar As CommandBar) As Boolean
    '******************************************************************************************************************
    '功能：调用药嘱审查
    '参数：
    '返回：
    '******************************************************************************************************************
    Call mclsInAdvices.zlPopupCommandBars(CommandBar)
End Function

Public Function zlMediAuditShell(ByVal Control As CommandBarControl) As Boolean
    '******************************************************************************************************************
    '功能：调用药嘱审查执行
    '参数：
    '返回：
    '******************************************************************************************************************
    Call mclsInAdvices.zlExecuteCommandBars(Control)
End Function

Public Function GetTbcStatus() As Boolean
    GetTbcStatus = tbcSub.Item(2).Selected
End Function

Private Sub ShowArchiveTab(ByVal strShow As String, ByVal strCaption As String)
'功能：切换显示不同的档案页面
    Dim i As Long

    For i = 0 To tbcSub.ItemCount - 1
        If tbcSub(i).Tag = strShow Then
            tbcSub(i).Caption = strCaption
            tbcSub(i).Visible = True
            tbcSub(i).Selected = True
        Else
            If tbcSub(i).Visible Then
                tbcSub(i).Visible = False
            End If
        End If
    Next
End Sub

Private Function InitControl() As Boolean
Dim objTab As TabControlItem

    On Error GoTo errHand
    If Not gobjEmr Is Nothing Then
        If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
            Set gobjEmr = Nothing
        Else
            Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "新版病历", False)
            If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
        End If
    End If
    Set mobjPACSDoc = DynamicCreate("zlPublicPACS.clsPublicPacs", "新版PACS编辑器", False)
    If Not mobjPACSDoc Is Nothing Then
        Call mobjPACSDoc.InitInterface(gcnOracle, gstrDBUser)
    End If

    Call TabControlInit(tbcSub)
    With tbcSub
        .PaintManager.BoldSelected = True
        
'        Set mfrmChildMedrec = New frmChildMedrec
        
        '初始CISJOB首页接口
        Set mclsArchiveMedRec = New zlMedRecPage.clsArchive
        Call mclsArchiveMedRec.InitArchiveMedRec(gcnOracle, glngSys)
        Set mfrmArchiveMedRec = mclsArchiveMedRec.zlGetForm(1)
        
        Set mclsInAdvices = New zlCISKernel.clsDockInAdvices
        Set mclsDockAduits = New zlRichEPR.clsDockAduits
        Set mclsPath = New zlCISPath.clsDockPath
        Set mclsTendsNew = New zl9TendFile.clsTendFile
        
        Call mclsInAdvices.zlDefCommandBars(Me, Nothing, 2)
        Call mclsTendsNew.InitTendFile(gcnOracle, glngSys)
        
        Call FormSetCaption(mclsDockAduits.zlGetFormTendBody, False, False)

        Set objTab = .InsertItem(0, "首页记录", mfrmArchiveMedRec.hWnd, 0): objTab.Tag = "首页记录"
        Set objTab = .InsertItem(1, "住院病历", mclsDockAduits.zlGetFormEPR.hWnd, 0): objTab.Tag = "住院病历"
        Set objTab = .InsertItem(2, "住院医嘱", mclsInAdvices.zlGetForm.hWnd, 0): objTab.Tag = "住院医嘱"
        Set objTab = .InsertItem(3, "体温记录单", mclsDockAduits.zlGetFormTendBody.hWnd, 0): objTab.Tag = "体温记录单"
        Set objTab = .InsertItem(4, "护理记录单", mclsDockAduits.zlGetFormTendFile.hWnd, 0): objTab.Tag = "护理记录单"
        Set objTab = .InsertItem(5, "临床路径", mclsPath.zlGetForm.hWnd, 0): objTab.Tag = "临床路径"
        Set objTab = .InsertItem(6, "新版护理", mclsTendsNew.zlGetfrmInTendFile.hWnd, 0): objTab.Tag = "新版护理"
        If Not mobjRichEMR Is Nothing Then
            Set objTab = .InsertItem(7, "电子病历", mobjRichEMR.zlGetForm.hWnd, 0): objTab.Tag = "电子病历"
        End If
        If Not mobjPACSDoc Is Nothing Then
            Set objTab = .InsertItem(8, "检查报告", mobjPACSDoc.zlDocGetForm.hWnd, 0): objTab.Tag = "检查报告"
        End If
        .Item(0).Selected = True
    End With

    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'######################################################################################################################

Private Sub Form_Resize()
    On Error Resume Next
    
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        
    Set mfrmMain = Nothing
    Set mobjReport = Nothing
    Unload mclsInAdvices.zlGetForm: Set mclsInAdvices.zlGetForm = Nothing
    
    Set mclsInAdvices = Nothing
    Unload mclsDockAduits.zlGetFormEPR: Set mclsDockAduits.zlGetFormEPR = Nothing
    Unload mclsDockAduits.zlGetFormTendBody: Set mclsDockAduits.zlGetFormTendBody = Nothing
    Unload mclsDockAduits.zlGetFormTendFile: Set mclsDockAduits.zlGetFormTendFile = Nothing
    Set mclsDockAduits = Nothing
    Unload mclsPath.zlGetForm:  Set mclsPath.zlGetForm = Nothing
    Set mclsPath = Nothing
    Unload mclsTendsNew.zlGetfrmInTendFile: Set mclsTendsNew.zlGetfrmInTendFile = Nothing
    Set mclsTendsNew = Nothing
    Unload mobjRichEMR.zlGetForm: Set mobjRichEMR.zlGetForm = Nothing
    Set mobjRichEMR = Nothing
    Unload mobjPACSDoc.zlDocGetForm: Set mobjPACSDoc.zlDocGetForm = Nothing
    Set mobjPACSDoc = Nothing
    
    Set mfrmArchiveMedRec = Nothing
    Set mclsArchiveMedRec = Nothing
End Sub

Private Sub mclsDockAduits_AfterEprPrint(ByVal lngRecordId As Long)
    mblnPrinted = True
End Sub

Private Sub mclsDockAduits_AfterTendPrint(ByVal lngFileID As Long)
    
    Call RecordEprPrintInfo(3, lngFileID, mlngNo, mlng病人ID, mlng主页ID)
    
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    mblnPrinted = True
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    tbcSub.Move 0, 0, picPane(Index).Width, picPane(Index).Height
End Sub
