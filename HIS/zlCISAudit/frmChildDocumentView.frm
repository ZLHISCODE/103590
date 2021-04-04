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
   StartUpPosition =   3  '����ȱʡ
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
Private mlng����ID As Long
Private mlng��ҳID As Long
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
Private WithEvents mclsTendsNew As zl9TendFile.clsTendFile    '�°滤ʿ����վ
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
    Dim lng����ID As Long
    
    Set mobjReport = New clsReport
   
    mblnPrinted = False
    
    Select Case mstrObject
    Case "��ҳ��¼"
    
        If mobjReport Is Nothing Then Set mobjReport = New clsReport
        
        lng����ID = GetlngID(mlng����ID, mlng��ҳID)
        
        Select Case Val(zlDatabase.GetPara("������ҳ��׼", glngSys, 1261, "0"))
        Case 0 '��������׼
            If Have��������(lng����ID, "��ҽ��") Then
                strName = "ZL1_INSIDE_1261_4"
            Else
                strName = "ZL1_INSIDE_1261_1"
            End If
        Case 1    '�Ĵ�ʡ��׼
            If Have��������(lng����ID, "��ҽ��") Then
                strName = "ZL1_INSIDE_1261_6"
            Else
                strName = "ZL1_INSIDE_1261_5"
            End If
        Case 2    '����ʡ��׼
            If Have��������(lng����ID, "��ҽ��") Then
                strName = "ZL1_INSIDE_1261_8"
            Else
                strName = "ZL1_INSIDE_1261_7"
            End If
        End Select
        
        
        Call mobjReport.ReportOpen(gcnOracle, ParamInfo.ϵͳ��, strName, mfrmMain, "����id=" & mlng����ID, "��ҳid=" & mlng��ҳID, bytMode)
        If mblnPrinted Then Call RecordEprPrintInfo(2, "��ҳ��¼", mlngNo, mlng����ID, mlng��ҳID)
        
    Case "סԺҽ��"
        If blnDoctorAdvice Then
            '��ӡ����
            Call gobjKernel.zlPrintAdvice(Me, mlng����ID, mlng��ҳID, 0, 0)
            
            '��ӡ����
            Call gobjKernel.zlPrintAdvice(Me, mlng����ID, mlng��ҳID, 0, 1)
        Else
            Call mobjReport.ReportOpen(gcnOracle, ParamInfo.ϵͳ��, "ZL1_INSIDE_1560", mfrmMain, "����id=" & mlng����ID, "��ҳid=" & mlng��ҳID, bytMode)
        End If
        
        If mblnPrinted Then Call RecordEprPrintInfo(2, "סԺҽ��", mlngNo, mlng����ID, mlng��ҳID)
        
    Case "סԺ����"
                
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
        
    Case "������"
                
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
        
    Case "֪���ļ�"

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
        
    Case "����֤��"
        
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
        
    Case "ҽ������"

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
        
    Case "�����¼"
        
        If mstrParam <> "" Then
            mblnNewTends = Get�°滤��(mlng����ID, mlng��ҳID)
            If mblnNewTends = False Then
                varParam = Split(mstrParam, ";")
                If UBound(varParam) >= 1 Then
                    If Val(varParam(1)) = -1 Then
                        bytRet = mclsDockAduits.zlPrintDocument(1, bytMode, , strPrintDeviceName)
                        
                        If bytRet = 2 Or bytMode = 2 Then
                           Call RecordEprPrintInfo(2, "���µ�", mlngNo, mlng����ID, mlng��ҳID)
                        End If
                        
                    Else
                        Call mclsDockAduits.zlPrintDocument(2, bytMode, , strPrintDeviceName)
                        If bytMode = 2 Then
                            Call RecordEprPrintInfo(3, Val(varParam(3)), mlngNo, mlng����ID, mlng��ҳID)
                        End If
                    End If
                End If
            Else
                '�°滤���ӡ��Ԥ��
                '�˲������� ����
                varParam = Split(mstrParam, ";")
                    If UBound(varParam) >= 1 Then
                    
                    Select Case Val(varParam(1))
                        Case -1 '���µ�
                            intSel = 1
                        Case 1  '����ͼ
                            intSel = 3
                        Case Else '��¼��
                            intSel = 2
                    End Select
                    Call mclsTendsNew.zlPrintTendFile(intSel, bytMode, strPrintDeviceName)
                End If
            End If
        End If
    Case "�ٴ�·��"
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
    Case "��ҳ��¼"
        Call mobjReport.ReportPrintSet(gcnOracle, ParamInfo.ϵͳ��, "ZL1_INSIDE_1261_1", Me)
    Case "סԺҽ��"
        Call mobjReport.ReportPrintSet(gcnOracle, ParamInfo.ϵͳ��, "ZL1_INSIDE_1560", Me)
    Case "סԺ����"
    Case "������"
    Case "֪���ļ�"
    Case "����֤��"
    Case "ҽ������"
    Case "�����¼"
    Case "�ٴ�·��"
        Control.ID = 101
        Call mclsPath.zlExecuteCommandBars(Control)
    End Select
    
End Function

Public Function zlInitData(ByVal frmMain As Object) As Boolean
    Set mfrmMain = frmMain
    zlInitData = InitControl
End Function

Public Function zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strObject As String, ByVal strParam As String, ByVal strCaption As String, ByVal blnDataMoved As Boolean) As Boolean
    Dim varParam As Variant
    Dim intSel As Integer
    
    mstrObject = strObject
    mstrParam = strParam
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlngNo = 0
 
    Select Case strObject
    Case "��ҳ��¼"
        
        Call mclsArchiveMedRec.zlRefresh(1, lng����ID, lng��ҳID, False)
        Call ShowArchiveTab(strObject, strCaption)
        
    Case "סԺҽ��"
        
        Call mclsInAdvices.zlRefresh(lng����ID, lng��ҳID, 0, 0, 1, blnDataMoved, 0, 0)
        Call ShowArchiveTab(strObject, strCaption)
                
    Case "סԺ����"
                
        If strParam <> "" Then
            If IsNumeric(Split(strParam, ";")(0)) Then
                varParam = Split(strParam, ";")
                mclsDockAduits.ParentForm mfrmMain
                mlngRecordKey = Val(varParam(0))
                Call mclsDockAduits.zlRefresh(2, Val(varParam(0)), , , , , , , blnDataMoved)
                Call ShowArchiveTab(strObject, strCaption)
            ElseIf strParam <> "" Then
                'EMR����Ԥ��
                If Not mobjRichEMR Is Nothing Then
                    If InStr(strParam, "|") > 0 Then
                        Call mobjRichEMR.zlShowDoc(Split(strParam, "|")(0), Split(strParam, "|")(1))
                    Else
                        Call mobjRichEMR.zlShowDoc(strParam, "")
                    End If
                End If
                Call ShowArchiveTab("���Ӳ���", strCaption)
            End If
        End If
        
    Case "������"
                
        If strParam <> "" Then
            If IsNumeric(Split(strParam, ";")(0)) Then
                varParam = Split(strParam, ";")
                mclsDockAduits.ParentForm mfrmMain
                mlngRecordKey = Val(varParam(0))
                Call mclsDockAduits.zlRefresh(4, Val(varParam(0)), , , , , , , blnDataMoved)
                Call ShowArchiveTab("סԺ����", strCaption)
            ElseIf strParam <> "" Then
                'EMR����Ԥ��
                If Not mobjRichEMR Is Nothing Then
                    If InStr(strParam, "|") > 0 Then
                        Call mobjRichEMR.zlShowDoc(Split(strParam, "|")(0), Split(strParam, "|")(1))
                    Else
                        Call mobjRichEMR.zlShowDoc(strParam, "")
                    End If
                End If
                Call ShowArchiveTab("���Ӳ���", strCaption)
            End If
        End If
        
    Case "֪���ļ�"
        
        If strParam <> "" Then
            If IsNumeric(Split(strParam, ";")(0)) Then
                varParam = Split(strParam, ";")
                mclsDockAduits.ParentForm mfrmMain
                mlngRecordKey = Val(varParam(0))
                Call mclsDockAduits.zlRefresh(6, Val(varParam(0)), , , , , , , blnDataMoved)
                Call ShowArchiveTab("סԺ����", strCaption)
            ElseIf strParam <> "" Then
                'EMR����Ԥ��
                If Not mobjRichEMR Is Nothing Then
                    If InStr(strParam, "|") > 0 Then
                        Call mobjRichEMR.zlShowDoc(Split(strParam, "|")(0), Split(strParam, "|")(1))
                    Else
                        Call mobjRichEMR.zlShowDoc(strParam, "")
                    End If
                End If
                Call ShowArchiveTab("���Ӳ���", strCaption)
            End If
        End If
        
    Case "����֤��"
        
        If strParam <> "" Then
            If IsNumeric(Split(strParam, ";")(0)) Then
                varParam = Split(strParam, ";")
                mclsDockAduits.ParentForm mfrmMain
                mlngRecordKey = Val(varParam(0))
                Call mclsDockAduits.zlRefresh(5, Val(varParam(0)), , , , , , , blnDataMoved)
                Call ShowArchiveTab("סԺ����", strCaption)
            ElseIf strParam <> "" Then
                'EMR����Ԥ��
                If Not mobjRichEMR Is Nothing Then
                    If InStr(strParam, "|") > 0 Then
                        Call mobjRichEMR.zlShowDoc(Split(strParam, "|")(0), Split(strParam, "|")(1))
                    Else
                        Call mobjRichEMR.zlShowDoc(strParam, "")
                    End If
                End If
                Call ShowArchiveTab("���Ӳ���", strCaption)
            End If
        End If
    
    Case "ҽ������"

        If strParam <> "" Then
            varParam = Split(strParam, ";")
            mclsDockAduits.ParentForm mfrmMain
            mlngRecordKey = Val(varParam(0))
            If mlngRecordKey <> 0 Then
                                Call mclsDockAduits.zlRefresh(7, Val(varParam(0)), , , , , , , blnDataMoved)
                                Call ShowArchiveTab("סԺ����", strCaption)
            Else
                Call mobjPACSDoc.zlDocRefresh(varParam(2)) '��PACS����༭�����в���ҽ�������¼ ����='0;ҽ��ID;��鱨��ID'
                Call ShowArchiveTab("��鱨��", strCaption)
            End If
        End If
    
    Case "�����¼"
        If strParam <> "" Then
            mblnNewTends = Get�°滤��(lng����ID, lng��ҳID)
            If mblnNewTends = False Then
                varParam = Split(strParam, ";")
                If UBound(varParam) >= 1 Then
                    If Val(varParam(1)) = -1 Then
                        mlngRecordKey = Val(varParam(0))
                        Call mclsDockAduits.zlRefreshTendBody(lng����ID, lng��ҳID, Val(Split(varParam(0), "_")(0)), Val(varParam(4)), blnDataMoved)
                        Call ShowArchiveTab("���¼�¼��", strCaption)
                    Else
                        mlngRecordKey = Val(varParam(3))
                        Call mclsDockAduits.zlRefresh(3, Val(varParam(3)), lng����ID, lng��ҳID, Val(Split(varParam(0), "_")(0)), CStr(varParam(2)), , Val(varParam(4)), blnDataMoved)
                        Call ShowArchiveTab("�����¼��", strCaption)
                    End If
                End If
            Else
                '�˲������� ����
                varParam = Split(strParam, ";")
                If UBound(varParam) >= 1 Then
                    Select Case Val(varParam(1))
                        Case -1 '���µ�
                            intSel = 0
                        Case 1  '����ͼ
                            intSel = 2
                        Case Else '��¼��
                            intSel = 1
                    End Select
                    Call mclsTendsNew.zlRefreshTendFile(mlng����ID, lng��ҳID, Val(varParam(4)), Val(varParam(0)), False, False, intSel, Val(varParam(3)), 1)
                    Call ShowArchiveTab("�°滤��", strCaption)
                End If
            End If
        End If
    Case "�ٴ�·��"
        Call mclsPath.zlRefreshReadOnly(lng����ID, lng��ҳID)
        Call ShowArchiveTab(strObject, strCaption)
        
    Case Else

    End Select

    zlRefresh = True
    
End Function

Public Function zlMediAudit(ByVal CommandBar As CommandBar) As Boolean
    '******************************************************************************************************************
    '���ܣ�����ҩ�����
    '������
    '���أ�
    '******************************************************************************************************************
    Call mclsInAdvices.zlPopupCommandBars(CommandBar)
End Function

Public Function zlMediAuditShell(ByVal Control As CommandBarControl) As Boolean
    '******************************************************************************************************************
    '���ܣ�����ҩ�����ִ��
    '������
    '���أ�
    '******************************************************************************************************************
    Call mclsInAdvices.zlExecuteCommandBars(Control)
End Function

Public Function GetTbcStatus() As Boolean
    GetTbcStatus = tbcSub.Item(2).Selected
End Function

Private Sub ShowArchiveTab(ByVal strShow As String, ByVal strCaption As String)
'���ܣ��л���ʾ��ͬ�ĵ���ҳ��
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
            Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "�°没��", False)
            If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
        End If
    End If
    Set mobjPACSDoc = DynamicCreate("zlPublicPACS.clsPublicPacs", "�°�PACS�༭��", False)
    If Not mobjPACSDoc Is Nothing Then
        Call mobjPACSDoc.InitInterface(gcnOracle, gstrDBUser)
    End If

    Call TabControlInit(tbcSub)
    With tbcSub
        .PaintManager.BoldSelected = True
        
'        Set mfrmChildMedrec = New frmChildMedrec
        
        '��ʼCISJOB��ҳ�ӿ�
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

        Set objTab = .InsertItem(0, "��ҳ��¼", mfrmArchiveMedRec.hWnd, 0): objTab.Tag = "��ҳ��¼"
        Set objTab = .InsertItem(1, "סԺ����", mclsDockAduits.zlGetFormEPR.hWnd, 0): objTab.Tag = "סԺ����"
        Set objTab = .InsertItem(2, "סԺҽ��", mclsInAdvices.zlGetForm.hWnd, 0): objTab.Tag = "סԺҽ��"
        Set objTab = .InsertItem(3, "���¼�¼��", mclsDockAduits.zlGetFormTendBody.hWnd, 0): objTab.Tag = "���¼�¼��"
        Set objTab = .InsertItem(4, "�����¼��", mclsDockAduits.zlGetFormTendFile.hWnd, 0): objTab.Tag = "�����¼��"
        Set objTab = .InsertItem(5, "�ٴ�·��", mclsPath.zlGetForm.hWnd, 0): objTab.Tag = "�ٴ�·��"
        Set objTab = .InsertItem(6, "�°滤��", mclsTendsNew.zlGetfrmInTendFile.hWnd, 0): objTab.Tag = "�°滤��"
        If Not mobjRichEMR Is Nothing Then
            Set objTab = .InsertItem(7, "���Ӳ���", mobjRichEMR.zlGetForm.hWnd, 0): objTab.Tag = "���Ӳ���"
        End If
        If Not mobjPACSDoc Is Nothing Then
            Set objTab = .InsertItem(8, "��鱨��", mobjPACSDoc.zlDocGetForm.hWnd, 0): objTab.Tag = "��鱨��"
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
    
    Call RecordEprPrintInfo(3, lngFileID, mlngNo, mlng����ID, mlng��ҳID)
    
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    mblnPrinted = True
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    tbcSub.Move 0, 0, picPane(Index).Width, picPane(Index).Height
End Sub
