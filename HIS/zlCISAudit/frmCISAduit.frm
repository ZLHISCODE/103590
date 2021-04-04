VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISAduit 
   Caption         =   "���Ӳ������"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14370
   Icon            =   "frmCISAduit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   14370
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   0
      Left            =   270
      ScaleHeight     =   3585
      ScaleWidth      =   3570
      TabIndex        =   1
      Top             =   705
      Width           =   3570
      Begin XtremeSuiteControls.TabControl tbcTask 
         Height          =   1830
         Left            =   345
         TabIndex        =   3
         Top             =   615
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   9720
      TabIndex        =   0
      Top             =   105
      Width           =   1125
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8745
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22437
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmCISAduit.frx":6852
      Left            =   4020
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmCISAduit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���弶��������
'######################################################################################################################
Private mstrPrivs As String
Private mblnStartUp As Boolean
Private mblnAllowClose As Boolean
Private mblnAllowModify As Boolean
Private mstrCondition As String
Private mstrFindKey As String
Private mlngTmp As Long
Private mobjFindKey As CommandBarControl
Private mobjPrintView As CommandBarControl
Private mobjPrintPatient As CommandBarControl
Private mobjPrint As CommandBarControl
Private mobjPrint1 As CommandBarControl
Private mintIndex As Integer
Private mlngModul As Long
Private mrsCondition As ADODB.Recordset
Private mblnMuliSelect As Boolean
Private mblnMediAudit  As Boolean
Private mblnMediAuditPass As Boolean
Private mblnTrans As Boolean
Private Type SELECTEDPERSON
    δ�������� As Long
    δ�鵵���� As Long
    �Ѿܾ����� As Long
    �ѹ鵵���� As Long
End Type
Private mblnAudit       As Boolean              '�Ƿ���Ҫ��˺���ܹ鵵
Private mblnAuditEnter  As Boolean              '�Ƿ�����¼��������
Private mblnDoctorAdvice As Boolean             '����ҽ����������ҽ����
Private mSelectedPerson As SELECTEDPERSON
Private mstrFindDeal As String
Private WithEvents mfrmChildPatientAduit As frmChildPatient '��Ժ������Ϣ
Attribute mfrmChildPatientAduit.VB_VarHelpID = -1
Private WithEvents mfrmChildPatientIn As frmChildPatient    '��Ժ������Ϣ
Attribute mfrmChildPatientIn.VB_VarHelpID = -1
Private WithEvents mfrmChildQuestion As frmChildQuestion
Attribute mfrmChildQuestion.VB_VarHelpID = -1
Private WithEvents mfrmChildDocumentView As frmChildDocumentView
Attribute mfrmChildDocumentView.VB_VarHelpID = -1
Private WithEvents mfrmChildDocumentScaleView As frmChildDocumentView
Attribute mfrmChildDocumentScaleView.VB_VarHelpID = -1

Private Const conMenu_Manage_magnify = 3946 '�Ŵ����
Private mblnShowDept As Boolean             '�Ƿ���ʾͣ�ò���
Public Property Let ShowDept(ByVal blnData As Boolean)
    mblnShowDept = blnData
End Property

Public Property Get ShowDept() As Boolean
    ShowDept = mblnShowDept
End Property
Public Property Let AllowModify(ByVal blnData As Boolean)
    mblnAllowModify = blnData
    If blnData = False Then mfrmChildQuestion.AllowModify = blnData
End Property

Public Property Get AllowModify() As Boolean
 AllowModify = mblnAllowModify
End Property
Public Property Get ģ���() As Long
    ģ��� = mlngModul
End Property

Private Property Let DataChanged(ByVal blnData As Boolean)
    mfrmChildQuestion.DataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmChildQuestion Is Nothing) Then
        DataChanged = mfrmChildQuestion.DataChanged
    End If
End Property

Private Function GetChildPatient(ByVal intIndex As Integer) As frmChildPatient
    Select Case intIndex
        Case 0
            Set GetChildPatient = mfrmChildPatientAduit
        Case 1
            Set GetChildPatient = mfrmChildPatientIn
    End Select
End Function

Private Function CountSelected() As Boolean
    '******************************************************************************************************************
    '���ܣ�ͳ��ѡ�еĸ���
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim lngCount As Long

    mSelectedPerson.δ�������� = 0
    mSelectedPerson.δ�鵵���� = 0
    mSelectedPerson.�ѹ鵵���� = 0
    mSelectedPerson.�Ѿܾ����� = 0
     
    Select Case mintIndex
        Case 0
            GetChildPatient(mintIndex).labSelect.Caption = "��Ժ"
            With GetChildPatient(mintIndex).VsfBody
                If GetChildPatient(mintIndex).VsfBody.Rows = 2 And GetChildPatient(mintIndex).VsfBody.RowData(1) = "" Then Exit Function
                If .ColIndex("ѡ��") = -1 Then Exit Function
                If .ColIndex("����״ֵ̬") = -1 Then Exit Function
                For lngLoop = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(lngLoop, .ColIndex("ѡ��")))) = 1 Then
                        If Val(.TextMatrix(lngLoop, .ColIndex("����״ֵ̬"))) = "10" Then
                            mSelectedPerson.δ�������� = mSelectedPerson.δ�������� + 1
                        ElseIf Val(.TextMatrix(lngLoop, .ColIndex("����״ֵ̬"))) = "2" Then
                            mSelectedPerson.�Ѿܾ����� = mSelectedPerson.�Ѿܾ����� + 1
                        End If
                        
                        If .TextMatrix(lngLoop, .ColIndex("����״ֵ̬")) = 3 Then
                            mSelectedPerson.δ�鵵���� = mSelectedPerson.δ�鵵���� + 1
                        End If
                        
                        If .TextMatrix(lngLoop, .ColIndex("����״ֵ̬")) = 5 Then
                            mSelectedPerson.�ѹ鵵���� = mSelectedPerson.�ѹ鵵���� + 1
                        End If
                    End If
                Next
            End With
            GetChildPatient(mintIndex).labNum.Caption = mSelectedPerson.δ�������� + mSelectedPerson.�Ѿܾ����� + mSelectedPerson.δ�鵵���� + mSelectedPerson.�ѹ鵵����
        Case 1
            GetChildPatient(mintIndex).labSelect.Visible = False
            GetChildPatient(mintIndex).labNum.Visible = False
            With GetChildPatient(mintIndex).VsfBody
                If .Rows = 1 Then
                    GetChildPatient(mintIndex).LabStatus.Caption = ""
                Else
                    If .ColIndex("����") <> -1 Then
                        GetChildPatient(mintIndex).LabStatus.Caption = "������" & .TextMatrix(.Row, .ColIndex("����")) & "    סԺ�ţ�" & .TextMatrix(.Row, .ColIndex("סԺ��"))
                    End If
                End If
            End With
    End Select
    mblnMuliSelect = (mSelectedPerson.δ�������� > 0 Or mSelectedPerson.δ�鵵���� > 0 Or mSelectedPerson.�ѹ鵵���� > 0 Or mSelectedPerson.�Ѿܾ����� > 0)
    CountSelected = True

End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    
    Call CommandBarInit(cbsMain)
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
    
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�ļ�
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_magnify, "��(&O)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_File_Print * 100, "�嵥���(&P)")
        Call NewCommandBar(objControl, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
        Call NewCommandBar(objControl, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Call NewCommandBar(objControl, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Call NewCommandBar(objControl, xtpControlButton, conMenu_File_Excel, "�����&Excel")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_File_BillPrint * 100, "����ĵ�")
        Set mobjPrintView = NewCommandBar(objControl, xtpControlButton, conMenu_File_BillPrintView, "Ԥ���ĵ�(&E)")
        Set mobjPrint = NewCommandBar(objControl, xtpControlButton, conMenu_File_BillPrint, "��ӡ�ĵ�(&T)")
        Set mobjPrint1 = NewCommandBar(objControl, xtpControlButton, conMenu_File_MedRecSetup, "��ӡ����")
    
    Set mobjPrintPatient = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BatPrint, "�������(&B)", True)
    Call NewCommandBar(objMenu, xtpControlButton, conMenu_File_BatPrint * 100, "�����PDF(&E)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "��������(&M)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    
    '�༭
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Plan, "��������(&B)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "�Զ����(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit * 10, "�������(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Audit, "���鵵(&D)", True)
    
'    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Edit_Untread, "���˲���(&U)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "���˳���(&1)", , , "���˳���")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "���˹鵵(&2)", , , "���˹鵵")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Pause, "�������(&P)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Reuse, "�������(&R)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SelAll, "ȫ��ѡ��(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ClsAll, "ȡ��ѡ��(&C)")
        
    '�鿴
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_ShowStoped, "��ʾͣ�ò���(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SafeKeep, "���鿴(&F)...", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Column, "ѡ������(&H)...", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_ReportView, "���Ĳ���(&V)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Filter, "����(&F)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
    
'     Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "����", -1, False)
'    objMenu.ID = conMenu_Edit_MediAudit
'    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Edit_MediAudit, "ҩ�����")
    '����
    '------------------------------------------------------------------------------------------------------------------
    Call CreateHelpMenu(cbsMain)
    
    '���˵��Ҳ�Ĳ���
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    
    mstrFindKey = GetPara("��λ����", mlngModul, "����")
    If mstrFindKey = "" Then mstrFindKey = "����"
    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.����", , , "����")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.סԺ��", , , "סԺ��")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.����", , , "����")
'''    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&4.���￨��", , , "���￨��")
    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, ""): cbrCustom.Handle = txtLocation.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "ǰһ��"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "��һ��"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon

    Set objPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationMethod, "ѡ��")
    objPopup.IconId = conMenu_View_LocationMethod
    objPopup.Flags = xtpFlagRightAlign
  
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_LocationMethod, "&1.�����Ҷ�λ", , , "�����Ҷ�λ")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_LocationMethod, "&2.���Ҳ�ѡ��", , , "���Ҳ�ѡ��")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_LocationMethod, "&3.���Ҳ���ѡ", , , "���Ҳ���ѡ")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_LocationMethod, "&4.���Ҳ���ѡ", , , "���Ҳ���ѡ")

    
    '����������:������������
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("��׼", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_magnify, "��")
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Plan, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Audit, "�Զ�")

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Audit, "�鵵")
        
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Pause, "���", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Reuse, "���")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SafeKeep, "�鿴���", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_ReportView, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "����", True)
    
    Set objControl = NewToolBar(objBar, xtpControlButtonPopup, conMenu_Edit_MediAudit, "ҩ�����", True)
'    objControl.ID = 0
'    Set objPopup = objControl.Add(xtpControlPopup, conMenu_Edit_MediAudit, "ҩ�����1", , False)
'    Set objPopup = NewToolBar(objControl, xtpControlSplitButtonPopup, conMenu_Edit_MediAudit, "ҩ�����")
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")

    
    '����Ŀ����:���������������Ѵ���
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF6, conMenu_Edit_Audit                 '�Զ�
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save           '����
        .Add 0, vbKeyF3, conMenu_Manage_Plan                '��������
        .Add 0, vbKeyF12, conMenu_File_Parameter            '��������
        .Add 0, vbKeyF11, conMenu_Manage_ReportView         '����
        .Add 0, vbKeyF10, conMenu_File_BatPrint            '��ӡ�������е���
        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add FCONTROL, vbKeyE, conMenu_File_BillPrintView   'Ԥ���ĵ�
        .Add FCONTROL, vbKeyT, conMenu_File_BillPrint       '��ӡ�ĵ�
        .Add FCONTROL, vbKeyZ, conMenu_File_MedRecSetup     '��ӡ����
        .Add FCONTROL, vbKeyF, conMenu_View_Filter          '����
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll          'ȫѡ
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_ClsAll       'ȫ��
        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      'ǰһ��
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '��һ��
    End With

End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop         As Integer
    Dim intRow          As Integer
    Dim rs              As New ADODB.Recordset
    Dim rsSQL           As New ADODB.Recordset
    Dim strTmp          As String
    Dim strSQL          As String
    Dim strNow          As String
    Dim strNote         As String
    Dim strDept         As String
    Dim strMsg          As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        
        Set mfrmChildPatientAduit = New frmChildPatient
        Set mfrmChildPatientIn = New frmChildPatient
        Set mfrmChildDocumentView = New frmChildDocumentView
        Set mfrmChildQuestion = New frmChildQuestion
        Set mfrmChildDocumentScaleView = New frmChildDocumentView
        
        Call mfrmChildPatientAduit.zlInitData(Me, 1, mstrPrivs)
        Call mfrmChildPatientIn.zlInitData(Me, 3, mstrPrivs)
        Call mfrmChildDocumentView.zlInitData(Me)
        Call mfrmChildDocumentScaleView.zlInitData(Me)

        Call mfrmChildQuestion.InitData(Me, mlngModul, IsPrivs(mstrPrivs, "��鲡��"), mblnAuditEnter, mstrPrivs)
        
        '��ʼ�˵���������
        '--------------------------------------------------------------------------------------------------------------
        Call InitCommandBar
        
        '����ͣ������
        '--------------------------------------------------------------------------------------------------------------
        Dim objPane As Pane
        Set objPane = dkpMain.CreatePane(1, 100, 200, DockLeftOf, Nothing): objPane.Title = "�����б�": objPane.Options = PaneNoCaption
        Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, Nothing): objPane.Title = "��������": objPane.Options = PaneNoCaption
        Set objPane = dkpMain.CreatePane(3, 300, 100, DockRightOf, objPane): objPane.Title = "��������": objPane.Options = PaneNoCloseable  'Or PaneNoFloatable
        
        dkpMain.SetCommandBars cbsMain
        Call DockPannelInit(dkpMain)
        Call TabControlInit(tbcTask)
        With tbcTask
            .PaintManager.BoldSelected = True
            
            .InsertItem 0, "��Ժ����", mfrmChildPatientAduit.hWnd, 4
            .InsertItem 1, "��Ժ����", mfrmChildPatientIn.hWnd, 3
                                                
            If IsPrivs(mstrPrivs, "������鲡��") = False Then .Item(0).Visible = False
            If IsPrivs(mstrPrivs, "���ĳ�鲡��") = False Then .Item(1).Visible = False
            
            If .Item(0).Visible Then
                .Item(0).Selected = True
            ElseIf .Item(1).Visible Then
                .Item(1).Selected = True
            End If
            
        End With
            
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
                                
        '��������������Ŀ�������г�ʼ��
        Call ParamCreate(mrsCondition)
        Call ParamAdd(mrsCondition, "�ύ����", 1)
        Call ParamAdd(mrsCondition, "���մ���", 1)
        Call ParamAdd(mrsCondition, "�ܾ�����", 1)
        Call ParamAdd(mrsCondition, "�������", 1)
        Call ParamAdd(mrsCondition, "��鷴��", 1)
        Call ParamAdd(mrsCondition, "�������", 1)
        
        Call ParamAdd(mrsCondition, "��ǰ����", "")
        Call ParamAdd(mrsCondition, "��Ժ���", "")
        
        Call ParamAdd(mrsCondition, "��������", 0)
        Call ParamAdd(mrsCondition, "ҽ������", "")
        
        Call ParamAdd(mrsCondition, "��鿪ʼʱ��", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "������ʱ��", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "�鵵��ʼʱ��", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "�鵵����ʱ��", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    
        Call ParamAdd(mrsCondition, "��Ժ����", 0)
        Call ParamAdd(mrsCondition, "��Ժ��ʼʱ��", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "��Ժ����ʱ��", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        
        Call ParamAdd(mrsCondition, "ҽ����ʼʱ��", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "ҽ������ʱ��", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "סԺҽʦ", "")
        Call ParamAdd(mrsCondition, "��������", "")
        Call ParamAdd(mrsCondition, "�������", "")
        Call ParamAdd(mrsCondition, "ҩƷ��Ϣ", "")
                
        '��ȡȱʡʱ�䷶Χ
        strTmp = GetPara("���ȱʡ��Χ", mlngModul, "��  ��")
        If strTmp = "" Then strTmp = "��  ��"
        Call ParamWrite(mrsCondition, "��鿪ʼʱ��", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "������ʱ��", GetDateTime(strTmp, 2))
        
        strTmp = GetPara("�鵵ȱʡ��Χ", mlngModul, "��  ��")
        If strTmp = "" Then strTmp = "��  ��"
        Call ParamWrite(mrsCondition, "�鵵��ʼʱ��", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "�鵵����ʱ��", GetDateTime(strTmp, 2))
        
        strTmp = GetPara("��Ժȱʡ��Χ", mlngModul, "��  ��")
        If strTmp = "" Then strTmp = "��  ��"
        Call ParamWrite(mrsCondition, "��Ժ��ʼʱ��", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "��Ժ����ʱ��", GetDateTime(strTmp, 2))
        
        '�¼�����
        strTmp = GetPara("ҽ��ȱʡ��Χ", mlngModul, "��  ��")
        If strTmp = "" Then strTmp = "��  ��"
        Call ParamWrite(mrsCondition, "ҽ����ʼʱ��", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "ҽ������ʱ��", GetDateTime(strTmp, 2))
        
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
        If tbcTask.Enabled <> Not DataChanged Then
            
            tbcTask.Enabled = Not DataChanged
            
            mfrmChildPatientAduit.Enabled = Not DataChanged
            mfrmChildPatientIn.Enabled = Not DataChanged
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��״̬"
        
        With GetChildPatient(mintIndex).VsfBody
            If Val(.TextMatrix(.Row, .ColIndex("����id"))) = 0 Then
                strTmp = "��ǰ״̬�»�û���κβ��ˣ�"
            Else
                strTmp = "��ǰ״̬�¹��� " & .Rows - 1 & " �����ˣ�"
            End If
        End With
        
        stbThis.Panels(2).Text = strTmp
        
        With GetChildPatient(0).VsfBody
            If tbcTask.ItemCount > 0 Then
                If Val(.TextMatrix(.Row, .ColIndex("����id"))) > 0 Then
                    tbcTask.Item(0).Caption = "��Ժ����(" & .Rows - 1 & ")"
                Else
                    tbcTask.Item(0).Caption = "��Ժ����"
                End If
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
        Call ExecuteCommand("��ȡ��Ժ����")
        Call ExecuteCommand("��ȡ��Ժ����")
        
        Call ExecuteCommand("��ȡ���˲���")
        Call ExecuteCommand("��ȡ������¼")
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��ָ������"
        
        Select Case tbcTask.Selected.Index
            Case 0
                If UBound(varParam) >= 1 Then
                    Call mfrmChildPatientAduit.zlRefreshData(mrsCondition, Val(varParam(0)), Val(varParam(1)))
                Else
                    Call mfrmChildPatientAduit.zlRefreshData(mrsCondition)
                End If
            Case 1
                If UBound(varParam) >= 1 Then
                    Call mfrmChildPatientIn.zlRefreshData(mrsCondition, Val(varParam(0)), Val(varParam(1)))
                Else
                    Call mfrmChildPatientIn.zlRefreshData(mrsCondition)
                End If
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��Ժ����"
    
        Call GetChildPatient(0).zlRefreshData(mrsCondition)
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��Ժ����"
    
        Call GetChildPatient(1).zlRefreshData(mrsCondition)
    '------------------------------------------------------------------------------------------------------------------

    Case "��ȡ���˲���"
        
        Select Case mintIndex
            Case 0
                Call mfrmChildPatientAduit.zlShowDocument
            Case 1
                Call mfrmChildPatientIn.zlShowDocument
        End Select
        
        If mfrmChildQuestion.CurrentPatient Then
            Call mfrmChildQuestion.RefreshData(GetChildPatient(mintIndex).Depts, mrsCondition, mblnAuditEnter)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ������¼"
        strDept = GetChildPatient(mintIndex).Depts
        Call mfrmChildQuestion.RefreshData(strDept, mrsCondition, mblnAuditEnter)
    '------------------------------------------------------------------------------------------------------------------
    Case "��鱨��"
        
        With GetChildPatient(mintIndex).VsfBody
            If Not mblnMuliSelect Then
                strMsg = "ȷ�ϳ������²�����?" & vbCrLf & vbCrLf
                strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(.Row, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(.Row, .ColIndex("סԺ��"))
            ElseIf mSelectedPerson.δ�������� <= 5 Then
                strMsg = "ȷ�ϳ���" & mSelectedPerson.δ�������� & "���ݲ�����?" & vbCrLf & vbCrLf
                For intRow = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 And Val(.TextMatrix(intRow, .ColIndex("����״ֵ̬"))) = "10" And .TextMatrix(intRow, .ColIndex("���ʱ��")) = "" Then
                         strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(intRow, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(intRow, .ColIndex("סԺ��")) & vbCrLf
                    End If
                Next
            Else
                strMsg = "ȷ�ϳ���" & mSelectedPerson.δ�������� & "���ݲ�����"
            End If
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                If strNow = "" Then strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                If mblnMuliSelect Then
                    For intRow = 1 To .Rows - 1
                        If Val(.TextMatrix(intRow, .ColIndex("ID"))) > 0 And Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 And (Val(.TextMatrix(intRow, .ColIndex("����״ֵ̬"))) = 10 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 1) And .TextMatrix(intRow, .ColIndex("���ʱ��")) = "" Then
                            strTmp = strTmp & "," & Val(.TextMatrix(intRow, .ColIndex("ID")))
                            If Len(strTmp) > (40000 - 20) Then
                                strTmp = Mid(strTmp, 2)
                                strSQL = "zl_�����ύ��¼_SeReceive('" & strTmp & "','" & UserInfo.���� & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                                Call SQLRecordAdd(rsSQL, strSQL)
                                strTmp = ""
                            End If
                        End If
                    Next
                    
                    If Len(strTmp) > 0 Then
                        strTmp = Mid(strTmp, 2)
                        strSQL = "zl_�����ύ��¼_SeReceive('" & strTmp & "','" & UserInfo.���� & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                        Call SQLRecordAdd(rsSQL, strSQL)
                        strTmp = ""
                    End If
                Else
                    
                    If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And (Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 10 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 1) And .TextMatrix(.Row, .ColIndex("���ʱ��")) = "" Then
                        strSQL = "zl_�����ύ��¼_SeReceive('" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & "','" & UserInfo.���� & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                        Call SQLRecordAdd(rsSQL, strSQL)
                    End If
                        
                End If
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            
            GoTo endHand
        End With
                        
    '------------------------------------------------------------------------------------------------------------------
    Case "���鵵"
    
        With GetChildPatient(mintIndex).VsfBody
            If mintIndex = 0 Then '���ý��չ鵵
                If Not mblnMuliSelect Then
                    strMsg = "ȷ�Ϲ鵵���²�����?" & vbCrLf & vbCrLf
                    strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(.Row, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(.Row, .ColIndex("סԺ��"))
                ElseIf mSelectedPerson.δ�������� + mSelectedPerson.δ�鵵���� <= 5 Then
                    strMsg = "ȷ�Ϲ鵵��" & mSelectedPerson.δ�������� + mSelectedPerson.δ�鵵���� & "���ݲ�����?" & vbCrLf & vbCrLf
                    For intRow = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 And (Val(.TextMatrix(intRow, .ColIndex("����״ֵ̬"))) = "10" Or Val(.TextMatrix(intRow, .ColIndex("����״ֵ̬"))) = "3") And .TextMatrix(intRow, .ColIndex("���ʱ��")) = "" Then
                             strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(intRow, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(intRow, .ColIndex("סԺ��")) & vbCrLf
                        End If
                    Next
                Else
                    strMsg = "ȷ�Ϲ鵵��" & mSelectedPerson.δ�������� + mSelectedPerson.δ�鵵���� & "���ݲ�����"
                End If
            End If
            
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                
                If strNow = "" Then strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                
                If mblnMuliSelect Then
                    For intRow = 1 To .Rows - 1
                        If Val(.TextMatrix(intRow, .ColIndex("ID"))) > 0 And Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 And (Val(.TextMatrix(intRow, .ColIndex("����״ֵ̬"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 10) And .TextMatrix(.Row, .ColIndex("���ʱ��")) = "" Then
                            strTmp = strTmp & "," & Val(.TextMatrix(intRow, .ColIndex("ID")))
                            If Len(strTmp) > (4000 - 20) Then
                                strTmp = Mid(strTmp, 2)
                                strSQL = "zl_�����ύ��¼_Archive('" & strTmp & "','" & UserInfo.���� & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                                Call SQLRecordAdd(rsSQL, strSQL)
                                strTmp = ""
                            End If
                        End If
                    Next
                    
                    If Len(strTmp) > 0 Then
                        strTmp = Mid(strTmp, 2)
                        strSQL = "zl_�����ύ��¼_Archive('" & strTmp & "','" & UserInfo.���� & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                        Call SQLRecordAdd(rsSQL, strSQL)
                        strTmp = ""
                    End If
                Else
                    
                    If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And (Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 10) And .TextMatrix(.Row, .ColIndex("���ʱ��")) = "" Then
                        strSQL = "zl_�����ύ��¼_Archive('" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & "','" & UserInfo.���� & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                        Call SQLRecordAdd(rsSQL, strSQL)
                    End If
                        
                End If
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "���˽���"
        
        With GetChildPatient(mintIndex).VsfBody
            If Not mblnMuliSelect Then
                strMsg = "ȷ�ϻ��˳������²�����?" & vbCrLf & vbCrLf
                strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(.Row, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(.Row, .ColIndex("סԺ��"))
            ElseIf mSelectedPerson.δ�鵵���� <= 5 Then
                strMsg = "ȷ�ϻ��˳���" & mSelectedPerson.δ�鵵���� & "���ݲ�����?" & vbCrLf & vbCrLf
                For intRow = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 And Val(.TextMatrix(intRow, .ColIndex("����״ֵ̬"))) = "3" And .TextMatrix(intRow, .ColIndex("���ʱ��")) = "" Then
                         strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(intRow, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(intRow, .ColIndex("סԺ��")) & vbCrLf
                    End If
                Next
            Else
                strMsg = "ȷ�ϻ��˳���" & mSelectedPerson.δ�鵵���� & "���ݲ�����"
            End If
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                If mblnMuliSelect Then
                    For intRow = 1 To .Rows - 1
                        If Val(.TextMatrix(intRow, .ColIndex("ID"))) > 0 And Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 And Val(.TextMatrix(intRow, .ColIndex("����״ֵ̬"))) = 3 And .TextMatrix(intRow, .ColIndex("���ʱ��")) = "" Then
                            If Val(.TextMatrix(intRow, .ColIndex("��������"))) > 0 Then
                                '����Ƿ��з�����¼
                                strMsg = "��ǰ�������з�����¼���ܻ���!" & vbCrLf & vbCrLf
                                strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(intRow, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(intRow, .ColIndex("סԺ��"))
                                
                                Call MsgBox(strMsg, vbQuestion, ParamInfo.ϵͳ����)
                                ExecuteCommand = False
                                GoTo endHand
                            Else
                                strTmp = strTmp & "," & Val(.TextMatrix(intRow, .ColIndex("ID")))
                                If Len(strTmp) > (40000 - 20) Then
                                    strTmp = Mid(strTmp, 2)
                                    strSQL = "zl_�����ύ��¼_SeUnReceive('" & strTmp & "')"
                                    Call SQLRecordAdd(rsSQL, strSQL)
                                    strTmp = ""
                                End If
                            End If
                        End If
                    Next
                    
                    If Len(strTmp) > 0 Then
                        strTmp = Mid(strTmp, 2)
                        strSQL = "zl_�����ύ��¼_SeUnReceive('" & strTmp & "')"
                        Call SQLRecordAdd(rsSQL, strSQL)
                        strTmp = ""
                    End If
                Else
                    If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 3 And .TextMatrix(.Row, .ColIndex("���ʱ��")) = "" Then
                        If Val(.TextMatrix(.Row, .ColIndex("��������"))) > 0 Then
                            strMsg = "��ǰ�������з�����¼���ܻ���!" & vbCrLf & vbCrLf
                            strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(.Row, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(.Row, .ColIndex("סԺ��"))
                            
                            Call MsgBox(strMsg, vbQuestion, ParamInfo.ϵͳ����)
                            ExecuteCommand = False
                            GoTo endHand
                        Else
                            strSQL = "zl_�����ύ��¼_SeUnReceive('" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & "')"
                            Call SQLRecordAdd(rsSQL, strSQL)
                        End If
                    End If
                End If
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "���˹鵵"
    
        With GetChildPatient(mintIndex).VsfBody
            If Not mblnMuliSelect Then
                strMsg = "ȷ�ϻ��˹鵵���²�����?" & vbCrLf & vbCrLf
                strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(.Row, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(.Row, .ColIndex("סԺ��"))
            ElseIf mSelectedPerson.�ѹ鵵���� <= 5 Then
                strMsg = "ȷ�ϻ��˹鵵��" & mSelectedPerson.�ѹ鵵���� & "���ݲ�����" & vbCrLf & vbCrLf
                For intRow = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 And Val(.TextMatrix(intRow, .ColIndex("����״ֵ̬"))) = "5" And .TextMatrix(intRow, .ColIndex("���ʱ��")) = "" Then
                         strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(intRow, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(intRow, .ColIndex("סԺ��")) & vbCrLf
                    End If
                Next
            Else
                strMsg = "ȷ�ϻ��˹鵵��" & mSelectedPerson.�ѹ鵵���� & "���ݲ�����"
            End If
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                If mblnMuliSelect Then
                    For intRow = 1 To .Rows - 1
                        If Val(.TextMatrix(intRow, .ColIndex("ID"))) > 0 And Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 And Val(.TextMatrix(intRow, .ColIndex("����״ֵ̬"))) = 5 And .TextMatrix(intRow, .ColIndex("���ʱ��")) = "" Then
                            strTmp = strTmp & "," & Val(.TextMatrix(intRow, .ColIndex("ID")))
                            If Len(strTmp) > (40000 - 20) Then
                                strTmp = Mid(strTmp, 2)
                                strSQL = "zl_�����ύ��¼_UnArchive('" & strTmp & "')"
                                Call SQLRecordAdd(rsSQL, strSQL)
                                strTmp = ""
                            End If
                        End If
                    Next
                    
                    If Len(strTmp) > 0 Then
                        strTmp = Mid(strTmp, 2)
                        strSQL = "zl_�����ύ��¼_UnArchive('" & strTmp & "')"
                        Call SQLRecordAdd(rsSQL, strSQL)
                        strTmp = ""
                    End If
                Else
                    If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 5 And .TextMatrix(.Row, .ColIndex("���ʱ��")) = "" Then
                        strSQL = "zl_�����ύ��¼_UnArchive('" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & "')"
                        Call SQLRecordAdd(rsSQL, strSQL)
                    End If
                End If
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��没��"
    
        With GetChildPatient(mintIndex).VsfBody
            If Val(.TextMatrix(.Row, .ColIndex("����id"))) = 0 Then GoTo endHand
            
            If MsgBox("���Ƿ����Ҫ��浱ǰѡ�в��˵ĵ��Ӳ�����", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                
                If frmPubNoteEdit.ShowNoteEdit(Me, "����������", strNote) Then
                    strSQL = "zl_��������¼_Lock(" & Val(.TextMatrix(.Row, .ColIndex("����id"))) & "," & Val(.TextMatrix(.Row, .ColIndex("��ҳid"))) & ",'" & UserInfo.���� & "',Sysdate,'" & strNote & "')"
                    
                    Call SQLRecordAdd(rsSQL, strSQL)
                    ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
                End If
            End If
            GoTo endHand
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ⲡ��"
    
        With GetChildPatient(mintIndex).VsfBody
            If Val(.TextMatrix(.Row, .ColIndex("����id"))) = 0 Then GoTo endHand
            
            If MsgBox("���Ƿ����Ҫ��⵱ǰѡ�в��˵ĵ��Ӳ�����", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then

                strSQL = "zl_��������¼_UnLock(" & Val(.TextMatrix(.Row, .ColIndex("����id"))) & "," & Val(.TextMatrix(.Row, .ColIndex("��ҳid"))) & ")"
                
                Call SQLRecordAdd(rsSQL, strSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "���鿴"
        
        Call frmCISAuditSafeKeep.zlInitData(Me, mlngModul)
        frmCISAuditSafeKeep.Show vbModal, Me
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
        
        mrsCondition.Filter = ""
        ExecuteCommand = frmCISAduitFilter.ShowPara(Me, mrsCondition)

        GoTo endHand
        
    '--------------------------------------------------------------------------------------------------------------
    Case "ȫ��ѡ��"
        With GetChildPatient(mintIndex).VsfBody
            If .ColIndex("ѡ��") = -1 Then Exit Function
            .Cell(flexcpText, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 1
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "ȫ��ȡ��"
        With GetChildPatient(mintIndex).VsfBody
            If .ColIndex("ѡ��") = -1 Then Exit Function
            .Cell(flexcpText, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 0
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "ǰһ��"
        
        With GetChildPatient(mintIndex).VsfBody
            If .Row > 1 Then
                .Row = .Row - 1
                .ShowCell .Row, .Col
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "��һ��"
        
        With GetChildPatient(mintIndex).VsfBody
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            End If
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case "������Ϣ����"
        
        GetChildPatient(mintIndex).zlColumnSelect
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ע���"
        
        mstrFindDeal = Trim(zlDatabase.GetPara("�鵽����", ParamInfo.ϵͳ��, mlngModul, "�����Ҷ�λ"))
        
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
            'ʹ�ø��Ի�����
            mstrFindKey = Trim(GetPara("��λ����", mlngModul, "����"))
            
            Call RestoreWinState(Me, App.ProductName)
            
            dkpMain.LoadStateFromString (GetRegister(˽��ģ��, Me.Name & "\��������\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString))
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "дע���"
        
        Call zlDatabase.SetPara("�鵽����", mstrFindDeal, ParamInfo.ϵͳ��, mlngModul)
        
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
            'ʹ�ø��Ի�����
            Call SaveWinState(Me, App.ProductName)
            Call SetPara("��λ����", mstrFindKey, mlngModul)

        End If
        
        With GetChildPatient(mintIndex)
            If .cboDept.ListIndex >= 0 Then
                Call SetPara("�ϴ�״̬", mintIndex & ";" & .cboDept.ItemData(.cboDept.ListIndex) & ";" & .VsfBody.TextMatrix(.VsfBody.Row, .VsfBody.ColIndex("����id")) & ";" & .VsfBody.TextMatrix(.VsfBody.Row, .VsfBody.ColIndex("��ҳid")) & ";" & .tvw.SelectedItem.Key, ģ���)
            Else
                Call SetPara("�ϴ�״̬", mintIndex & ";0;" & .VsfBody.TextMatrix(.VsfBody.Row, .VsfBody.ColIndex("����id")) & ";" & .VsfBody.TextMatrix(.VsfBody.Row, .VsfBody.ColIndex("��ҳid")) & ";" & .tvw.SelectedItem.Key, ģ���)
            End If
        End With
        
        Call SetRegister(˽��ģ��, Me.Name & "\��������\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    End Select

    ExecuteCommand = True

    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:

End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BillPrintView                    'Ԥ����ǰ�ĵ�
        
        If Not mfrmChildDocumentView Is Nothing Then
        
            Call mfrmChildDocumentView.zlPrintDocument(cbsMain, 1, , , mblnDoctorAdvice)
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BillPrint                    '��ӡ��ǰ�ĵ�
        
        If Not mfrmChildDocumentView Is Nothing Then
        
            Call mfrmChildDocumentView.zlPrintDocument(cbsMain, 2, , , mblnDoctorAdvice)
            
        End If
    Case conMenu_File_MedRecSetup '��ӡ���õ�ǰ�ĵ�
        If Not mfrmChildDocumentView Is Nothing Then
        
            Call mfrmChildDocumentView.zlPrintSet(Control, mblnDoctorAdvice)
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_magnify '�Ŵ�
    
        Call GetChildPatient(mintIndex).FileBatPrint
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BatPrint
        Call frmCISAduitPDF.ShowMe(Me, GetChildPatient(mintIndex).VsfBody, mintIndex, mblnDoctorAdvice, False)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BatPrint * 100
        Call frmCISAduitPDF.ShowMe(Me, GetChildPatient(mintIndex).VsfBody, mintIndex, mblnDoctorAdvice, True)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Parameter
        If frmCISAduitPara.ShowEdit(Me, mstrPrivs) Then
            If zlDatabase.GetPara("סԺҽ����ӡ", ParamInfo.ϵͳ��, ParamInfo.ģ���, "����ҽ����", , IsPrivs(mstrPrivs, "��������")) = "����ҽ����" Then
                mblnDoctorAdvice = False
            Else
                mblnDoctorAdvice = True
            End If
        
            Call GetChildPatient(mintIndex).zlRefreshStruct
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Plan
        
        If ExecuteCommand("��鱨��") Then
            'ˢ�µ�ǰ���˵���ʾ��Ϣ
            Call ExecuteCommand("��ȡ��Ժ����")
            Call ExecuteCommand("��ȡ���˲���")
            Call ExecuteCommand("ˢ��״̬")
            If mblnMuliSelect Then
                MsgBox "���γɹ������󡿲�����" & mSelectedPerson.δ�������� & "����", vbInformation, ParamInfo.��Ʒ����
            End If
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh
        
        Call ExecuteCommand("ˢ������")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_ReportView      '������ҳ
        Call RecordLook
    Case conMenu_Edit_Audit '�Զ����
        With frmCISAduitAuto
            If mintIndex = 0 Then
                .lng�ύId = mfrmChildQuestion.�ύId
            Else
                .lng�ύId = -1
            End If
            .lng����ID = GetChildPatient(mintIndex).VsfBody.TextMatrix(GetChildPatient(mintIndex).VsfBody.Row, GetChildPatient(mintIndex).VsfBody.ColIndex("����ID"))
            .lng��ҳID = GetChildPatient(mintIndex).VsfBody.TextMatrix(GetChildPatient(mintIndex).VsfBody.Row, GetChildPatient(mintIndex).VsfBody.ColIndex("��ҳID"))
            .lng����ID = GetChildPatient(mintIndex).VsfBody.TextMatrix(GetChildPatient(mintIndex).VsfBody.Row, GetChildPatient(mintIndex).VsfBody.ColIndex("��Ժ����ID"))
            
            .strLink = IIf(mintIndex = 1, "1", "2") '1Ϊ��� 2Ϊ���
            .strTreeSelect = mfrmChildPatientAduit.tvw.SelectedItem.Key
            .Show vbModal
            If .blnCancel Then
                Set frmCISAduitAuto = Nothing
                Exit Sub
            End If
        End With
        Set frmCISAduitAuto = Nothing
        Call ExecuteCommand("��ȡ������¼")
    Case conMenu_Edit_Audit * 10
        If frmCISAduitAutos.ShowMe(Me, IIf(mintIndex = 1, 1, 2), GetChildPatient(mintIndex).VsfBody) Then
        Call ExecuteCommand("��ȡ������¼")
        End If
    Case conMenu_Manage_Audit
    
        If ExecuteCommand("���鵵") Then
            Call ExecuteCommand("��ȡ��Ժ����")
            Call ExecuteCommand("��ȡ���˲���")
            Call ExecuteCommand("ˢ��״̬")
            If mblnMuliSelect Then
                If mintIndex = 0 Then
                    MsgBox "���γɹ����鵵��������" & mSelectedPerson.δ�鵵���� & "����", vbInformation, ParamInfo.��Ʒ����
                ElseIf mintIndex = 1 Then
                    MsgBox "���γɹ����鵵��������" & mSelectedPerson.δ�鵵���� & "����", vbInformation, ParamInfo.��Ʒ����
                End If
            End If
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Untread
        
        Select Case Control.Caption
            Case "���˳���(&1)"
            
                If ExecuteCommand("���˽���") Then
                    Call ExecuteCommand("��ȡ��Ժ����")
                    Call ExecuteCommand("��ȡ���˲���")
                    Call ExecuteCommand("ˢ��״̬")
                    If mblnMuliSelect Then
                        MsgBox "���γɹ������˽��ա�������" & mSelectedPerson.δ�鵵���� & "����", vbInformation, ParamInfo.��Ʒ����
                    End If
                End If
            Case "���˹鵵(&2)"
                If ExecuteCommand("���˹鵵") Then
                    Call ExecuteCommand("��ȡ��Ժ����")
                    Call ExecuteCommand("��ȡ���˲���")
                    Call ExecuteCommand("ˢ��״̬")
                    If mblnMuliSelect Then
                        MsgBox "���γɹ������˹鵵��������" & mSelectedPerson.�ѹ鵵���� & "����", vbInformation, ParamInfo.��Ʒ����
                    End If
                End If
        End Select
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Pause
    
        If ExecuteCommand("��没��") Then
            'ˢ�µ�ǰ���˵���ʾ��Ϣ
            Call ExecuteCommand("ˢ��ָ������")
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Reuse
    
        If ExecuteCommand("��ⲡ��") Then
            'ˢ�µ�ǰ���˵���ʾ��Ϣ
            Call ExecuteCommand("ˢ��ָ������")
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SafeKeep
        
        Call ExecuteCommand("���鿴")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SelAll
        
        Call ExecuteCommand("ȫ��ѡ��")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ClsAll
    
        Call ExecuteCommand("ȫ��ȡ��")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationMethod
    
        mstrFindDeal = Control.Parameter
        cbsMain.RecalcLayout
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter '����
        
        If ExecuteCommand("��������") Then
            Call ExecuteCommand("��ȡ��Ժ����")
            Call ExecuteCommand("��ȡ��Ժ����")
            Call ExecuteCommand("��ȡ���˲���")
            Call ExecuteCommand("��ȡ������¼")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Column
        
        Call ExecuteCommand("������Ϣ����")
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward
        Call ExecuteCommand("ǰһ��")
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward
        Call ExecuteCommand("��һ��")
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        
        mobjFindKey.Execute
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
    
        Call LocationObj(txtLocation)
        
    '-----------------------------------------------------------------------------------------------------------------
    
    Case conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 22
        '������ҩ���
        Call mfrmChildDocumentView.zlMediAuditShell(Control)
    Case conMenu_View_ShowStoped
        Control.Checked = Not Control.Checked
        Me.ShowDept = Control.Checked
        If mintIndex = 0 Then
        Call GetChildPatient(mintIndex).InitData(Me.ShowDept)
        With GetChildPatient(mintIndex).VsfBody
            mfrmChildQuestion.AllowModify = Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 4 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 6
        End With
        
        Else
            mfrmChildQuestion.AllowModify = True
             Call GetChildPatient(mintIndex).InitData(False)
        End If
    Call ExecuteCommand("��ȡ���˲���")
    Call ExecuteCommand("ˢ��״̬")
     '-----------------------------------------------------------------------------------------------------------------
    Case Else
    
        If Control.ID > 400 And Control.ID < 500 Then
            With GetChildPatient(mintIndex).VsfBody
                Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me, "ID=" & Val(.TextMatrix(.Row, .ColIndex("ID"))), "����id=" & Val(.TextMatrix(.Row, .ColIndex("����id"))), "��ҳid=" & Val(.TextMatrix(.Row, .ColIndex("��ҳid"))))
            End With
        Else
             '��ҵ���޹صĹ��ܣ������Ĺ���
            Select Case mintIndex
                Case 0
                    Call CommandBarExecutePublic(Control, Me, GetChildPatient(mintIndex).VsfBody, "��Ժ�����б��嵥")
                Case 1
                    Call CommandBarExecutePublic(Control, Me, GetChildPatient(mintIndex).VsfBody, "��Ժ�����б��嵥")
            End Select
        End If
        
    End Select
    Call CountSelected

End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    If CommandBar.Parent.ID = conMenu_Edit_MediAudit Then '����ҩ�����
        Call mfrmChildDocumentView.zlMediAudit(CommandBar)
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHand

    With GetChildPatient(mintIndex).VsfBody
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel  'Ԥ��,��ӡ,�����Excel
            Control.Enabled = (Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_BillPrintView, conMenu_File_BillPrint, conMenu_File_BatPrint, conMenu_File_MedRecSetup, conMenu_File_BatPrint * 100

            Control.Visible = IsPrivs(mstrPrivs, "��ӡԤ���ĵ�")
            Control.Enabled = Control.Visible

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_MediAudit                         'ҩ�����
            
            Control.Visible = mfrmChildDocumentView.GetTbcStatus And mblnMediAudit And mblnMediAuditPass
        Case conMenu_Manage_Plan                            '��������

            Control.Visible = IsPrivs(mstrPrivs, "��������")
            If mblnMuliSelect Then
                Control.Enabled = (Control.Visible And mintIndex = 0 And mSelectedPerson.δ�������� > 0) And .TextMatrix(.Row, .ColIndex("���ʱ��")) = ""
            Else
                If mintIndex <> 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = Control.Visible And mintIndex = 0 And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And (Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 10 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 1) And .TextMatrix(.Row, .ColIndex("���ʱ��")) = ""
                End If
            End If
            If Me.AllowModify = False Then Control.Enabled = Me.AllowModify
        '--------------------------------   ------------------------------------------------------------------------------
        Case conMenu_Edit_Audit                 '�Զ����
            Control.Visible = IsPrivs(mstrPrivs, "�鵵����")
            If mintIndex = 1 Then
                Control.Enabled = (Control.Visible And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0)
            Else
                Control.Enabled = (Control.Visible And (Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) > 10) And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0)
            End If
            If Me.AllowModify = False Then Control.Enabled = Me.AllowModify
        Case conMenu_Edit_Audit * 10            '�������
            Control.Enabled = (Control.Visible And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0)
        Case conMenu_Manage_Audit                           '�鵵����

            Control.Visible = IsPrivs(mstrPrivs, "�鵵����")

            If mblnMuliSelect Then
                Control.Enabled = (Control.Visible And mintIndex = 0 And Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) <> 4 And Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) <> 5 And (mSelectedPerson.δ�������� + mSelectedPerson.δ�鵵����) > 0) And IIf(Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 10, Not mblnAudit, True) And .TextMatrix(.Row, .ColIndex("���ʱ��")) = ""
            Else
                If Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) > 10 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 5 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 6 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (Control.Visible And mintIndex = 0 And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) <> 2 And Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) <> 4 And Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) <> 5) And IIf(Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 10, Not mblnAudit, True) And .TextMatrix(.Row, .ColIndex("���ʱ��")) = ""
                End If
            End If
            If Me.AllowModify = False Then Control.Enabled = Me.AllowModify
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Untread                           '���˽���/�ܾ�/�鵵

            Select Case Control.Caption
                Case "���˳���(&1)"
                    Control.Visible = IsPrivs(mstrPrivs, "���˳���")
                Case "���˹鵵(&2)"
                    Control.Visible = IsPrivs(mstrPrivs, "���˹鵵")
            End Select

            If mblnMuliSelect Then
                Select Case Control.Caption
                    Case "���˳���(&1)"
                        Control.Enabled = (Control.Visible And mintIndex = 0 And mSelectedPerson.δ�鵵���� > 0)
                    Case "���˹鵵(&2)"
                        Control.Enabled = (Control.Visible And mintIndex = 0 And mSelectedPerson.�ѹ鵵���� > 0)
                End Select
            Else
                Select Case Control.Caption
                    Case "���˳���(&1)"
                    
                        If .ColIndex("����״ֵ̬") = -1 Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = (Control.Visible And mintIndex = 0 And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 3) And .TextMatrix(.Row, .ColIndex("���ʱ��")) = ""
                        End If
    
                    Case "���˹鵵(&2)"
                        If .ColIndex("����״ֵ̬") = -1 Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = (Control.Visible And mintIndex = 0 And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 5) And .TextMatrix(.Row, .ColIndex("���ʱ��")) = ""
                        End If
                End Select
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Pause                             '�������

            Control.Visible = IsPrivs(mstrPrivs, "��没��")
            Control.Enabled = (DataChanged = False And Control.Visible And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And .TextMatrix(.Row, .ColIndex("���ʱ��")) = "")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Reuse                             '�������

            Control.Visible = IsPrivs(mstrPrivs, "��ⲡ��")
            Control.Enabled = (DataChanged = False And Control.Visible And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And .TextMatrix(.Row, .ColIndex("���ʱ��")) <> "")
        
        '------------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SafeKeep
            
            Control.Visible = IsPrivs(mstrPrivs, "���鿴")
            
        '------------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SelAll, conMenu_Edit_ClsAll                       'ȫѡ��ȫ��

            Control.Enabled = (DataChanged = False And Control.Visible And mintIndex <> 3 And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0)
        Case conMenu_Manage_ReportView
            If Val(.TextMatrix(.Row, .ColIndex("ID"))) = 0 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Filter, conMenu_View_Refresh

            Control.Enabled = (DataChanged = False And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Forward

            Control.Enabled = (.Row > 1 And DataChanged = False)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Backward

            Control.Enabled = (.Row < .Rows - 1 And DataChanged = False)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationMethod               '
            Control.Checked = (mstrFindDeal = Control.Parameter)
            Control.Enabled = (DataChanged = False)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationItem        '
            Control.Checked = (mstrFindKey = Control.Parameter)
            Control.Enabled = (DataChanged = False)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Location, conMenu_View_Column
             Control.Enabled = (DataChanged = False)
        '--------------------------------------------------------------------------------------------------------------
        Case Else
            Call CommandBarUpdatePublic(Control, Me)
        End Select
    End With

    '------------------------------------------------------------------------------------------------------------------
errHand:
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = picPane(0).hWnd
        Case 2
            Item.Handle = mfrmChildDocumentView.hWnd
        Case 3
            Item.Handle = mfrmChildQuestion.hWnd
    End Select
End Sub

Private Sub Form_Activate()
    Dim varTmp As Variant
    Dim intPos As Integer
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    DoEvents
    mblnAudit = (zlDatabase.GetPara("���ղ��ܹ鵵", ParamInfo.ϵͳ��, ParamInfo.ģ���, "0", , IsPrivs(mstrPrivs, "��������")) = 1)
    mblnAuditEnter = (zlDatabase.GetPara("��������¼��������", ParamInfo.ϵͳ��, ParamInfo.ģ���, "0", , IsPrivs(mstrPrivs, "��������")) = 1)
    If zlDatabase.GetPara("סԺҽ����ӡ", ParamInfo.ϵͳ��, ParamInfo.ģ���, "����ҽ����", , IsPrivs(mstrPrivs, "��������")) = "����ҽ����" Then
        mblnDoctorAdvice = False
    Else
        mblnDoctorAdvice = True
    End If
    
    If ExecuteCommand("��ʼ����") = False Then GoTo errHand
    
    Call ExecuteCommand("ˢ������")
    
    mblnAllowClose = True
    
    varTmp = Split(GetPara("�ϴ�״̬", ģ���, "0"), ";")
    If Val(varTmp(0)) >= 0 And Val(varTmp(0)) <= 2 And tbcTask.ItemCount > Val(varTmp(0)) Then
        
        If tbcTask.Item(Val(varTmp(0))).Visible Then
            tbcTask.Item(Val(varTmp(0))).Selected = True
        
            If UBound(varTmp) > 2 Then
                Call GetChildPatient(mintIndex).zlLocationPatient(3, , , , Val(varTmp(2)), Val(varTmp(3)), Val(varTmp(1)))
            End If
            If UBound(varTmp) > 3 Then
                If varTmp(4) <> "" Then
                    intPos = InStr(varTmp(4), "K")
                    If InStr(varTmp(4), "K") > 0 Then
                        Call GetChildPatient(mintIndex).zlLocationDocument(Val(varTmp(2)), Val(varTmp(3)), Val(Mid(varTmp(4), 2, intPos - 2)), Mid(varTmp(4), intPos + 1))
                    ElseIf InStr(varTmp(4), "R") Then
                        Call GetChildPatient(mintIndex).zlLocationDocument(Val(varTmp(2)), Val(varTmp(3)), Val(Mid(varTmp(4), 2)), "")
                    Else
                        Call GetChildPatient(mintIndex).zlLocationDocument(Val(varTmp(2)), Val(varTmp(3)), 2, varTmp(4))
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub

    '------------------------------------------------------------------------------------------------------------------
errHand:
    mblnAllowClose = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mblnAllowClose = False
    mblnShowDept = False
    mstrPrivs = UserInfo.ģ��Ȩ��
    mlngModul = ParamInfo.ģ���
    
    '����Ƿ���к�����ҩȨ��
    If InStrRev(GetPrivFunc(100, 1253), "������ҩ���") > 0 Then
        mblnMediAudit = True
    Else
        mblnMediAudit = False
    End If
    
    '����Ƿ������˺�����ҩ�ӿ� 0-δ���� 1-��ͨ�ӿ� 2-��ͨ�ӿ�
    '����ֻ�ж���ͨ�ӿ��Ƿ���
    If Val(zlDatabase.GetPara(30, glngSys)) = 1 Then
        mblnMediAuditPass = True
    Else
        mblnMediAuditPass = False
    End If
    
    Call ExecuteCommand("��ʼ�ؼ�")
    Call ExecuteCommand("��ע���")

    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1261_1", "ZL1_INSIDE_1254_1", "ZL1_INSIDE_1254_2")
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 1, 250, 100, 500, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 3, 250, 100, 400, Me.ScaleHeight)
    dkpMain.RecalcLayout
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = Not mblnAllowClose
    
    If Cancel = False Then

        
        Call ExecuteCommand("дע���")
        

        If Not (mfrmChildDocumentView Is Nothing) Then Unload mfrmChildDocumentView
        If Not (mfrmChildDocumentScaleView Is Nothing) Then Unload mfrmChildDocumentScaleView
        If Not (mfrmChildQuestion Is Nothing) Then Unload mfrmChildQuestion
        If Not (mfrmChildPatientAduit Is Nothing) Then Unload mfrmChildPatientAduit
        If Not (mfrmChildPatientIn Is Nothing) Then Unload mfrmChildPatientIn
        If Not (frmChildScale Is Nothing) Then Unload frmChildScale
        Set frmChildScale = Nothing
        
        Set mrsCondition = Nothing
    End If

End Sub

Private Sub mfrmChildPatientAduit_AfterDeptChanged()
     Call ExecuteCommand("ˢ��״̬")
End Sub

Private Sub mfrmChildPatientAduit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call CountSelected
End Sub

Private Sub mfrmChildPatientAduit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    If Button = 2 Then
        Set cbrPopupBar = CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        
        cbrPopupBar.ShowPopup
    End If
End Sub

Private Sub mfrmChildPatientAduit_AfterDocumentChanged(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strObject As String, ByVal strParam As String, ByVal strCaption As String, ByVal lng�ύId As Long, ByVal blnDataMove As Boolean, ByVal blnScale As Boolean)
     If blnScale Then
        
        Call mfrmChildDocumentScaleView.zlRefresh(lng����ID, lng��ҳID, strObject, strParam, strCaption, blnDataMove)
        Call frmChildScale.zlInitData(mfrmChildDocumentScaleView)
        frmChildScale.Show
        
'        Set mfrmChildDocumentScaleView = New frmChildDocumentView
'        Call mfrmChildDocumentScaleView.zlInitData(Me)
    Else
        Call mfrmChildDocumentView.zlRefresh(lng����ID, lng��ҳID, strObject, strParam, strCaption, blnDataMove)
        
        mobjPrintView.Caption = "Ԥ��""" & mfrmChildPatientAduit.Title & """(&E)"
        mobjPrint.Caption = "��ӡ""" & mfrmChildPatientAduit.Title & """(&T)"
        mobjPrint1.Caption = "��ӡ����""" & mfrmChildPatientAduit.Title & """(&T)"
        
        With GetChildPatient(mintIndex).VsfBody
            mobjPrintPatient.Caption = "��ӡ""" & .TextMatrix(.Row, .ColIndex("����")) & """�ĵ���(&B)"
        End With
        cbsMain.RecalcLayout
        
        If Not (mfrmChildQuestion Is Nothing) Then
            Call mfrmChildQuestion.SetParamter(lng����ID, lng��ҳID, strObject, strParam, lng�ύId)
            If mfrmChildQuestion.CurrentPatient Then
                If mintIndex = 0 Then
                    With GetChildPatient(mintIndex).VsfBody
                        mfrmChildQuestion.AllowModify = (Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 4 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 6) And zlCommFun.NVL(.TextMatrix(.Row, .ColIndex("���ʱ��")) = "" And Me.AllowModify)
                    End With
                Else
                    mfrmChildQuestion.AllowModify = True
                End If
                
                If strParam = "" And (strObject = "סԺ����" Or strObject = "������" Or strObject = "�����¼" Or strObject = "ҽ������" Or strObject = "����֤��" Or strObject = "֪���ļ�") Then
                    '��ˢ����ϸ����
                Else
                    Call mfrmChildQuestion.RefreshData(GetChildPatient(mintIndex).Depts, mrsCondition, mblnAuditEnter)
                End If
            End If
        End If
    End If
End Sub

Private Sub mfrmChildPatientAduit_StatusChanged()
     Call ExecuteCommand("ˢ��״̬")
End Sub

Private Sub mfrmChildPatientIn_AfterDeptChanged()
     Call ExecuteCommand("ˢ��״̬")
End Sub

Private Sub mfrmChildPatientIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    If Button = 2 Then
        Set cbrPopupBar = CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        cbrPopupBar.ShowPopup
    End If
End Sub

Private Sub mfrmChildPatientIn_AfterDocumentChanged(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strObject As String, ByVal strParam As String, ByVal strCaption As String, ByVal lng�ύId As Long, ByVal blnDataMove As Boolean, ByVal blnScale As Boolean)
    If blnScale Then
        '��鲡������ --ZQ
'        Set mfrmChildDocumentScaleView = New frmChildDocumentView
'        Call mfrmChildDocumentScaleView.zlInitData(Me)
  
        Call mfrmChildDocumentScaleView.zlRefresh(lng����ID, lng��ҳID, strObject, strParam, strCaption, blnDataMove)
        Call frmChildScale.zlInitData(mfrmChildDocumentScaleView)
        frmChildScale.Show
        
'        Unload frmChildScale
'        Set frmChildScale = Nothing
'        Set mfrmChildDocumentScaleView = Nothing
    Else
        Call mfrmChildDocumentView.zlRefresh(lng����ID, lng��ҳID, strObject, strParam, strCaption, blnDataMove)
        
        mobjPrintView.Caption = "Ԥ��""" & mfrmChildPatientIn.Title & """(&E)"
        mobjPrint.Caption = "��ӡ""" & mfrmChildPatientIn.Title & """(&T)"
        mobjPrint1.Caption = "��ӡ����""" & mfrmChildPatientIn.Title & """(&T)"
        With GetChildPatient(mintIndex).VsfBody
            mobjPrintPatient.Caption = "��ӡ""" & .TextMatrix(.Row, .ColIndex("����")) & """�ĵ���(&B)"
        End With
        cbsMain.RecalcLayout
        
        If Not (mfrmChildQuestion Is Nothing) Then
            Call mfrmChildQuestion.SetParamter(lng����ID, lng��ҳID, strObject, strParam)
            If mfrmChildQuestion.CurrentPatient Then
                Call mfrmChildQuestion.RefreshData(GetChildPatient(mintIndex).Depts, mrsCondition, mblnAuditEnter)
            End If
        End If
    End If
End Sub

Private Sub mfrmChildPatientIn_StatusChanged()
     Call ExecuteCommand("ˢ��״̬")
End Sub

Private Sub mfrmChildQuestion_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildQuestion_AfterDeleteQuestion(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    '
    Call ExecuteCommand("ˢ��ָ������", lng����ID, lng��ҳID)
End Sub

Private Sub mfrmChildQuestion_AfterQuestionType(ByVal blnQuestionType As Boolean)
    'blnQuestionType=True Ժ������ =Flase �Ƽ�����
    If blnQuestionType Then
        If ObjPtr(dkpMain.Panes(2)) > 0 Then
            dkpMain.Panes(3).Title = "Ժ�����ⷴ��"
        End If
    Else
        If ObjPtr(dkpMain.Panes(2)) > 0 Then
            dkpMain.Panes(3).Title = "�Ƽ����ⷴ��"
        End If
    End If
End Sub

Private Sub mfrmChildQuestion_AfterSaveQuestion(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    '
    Call ExecuteCommand("ˢ��ָ������", lng����ID, lng��ҳID)
End Sub

Private Sub mfrmChildQuestion_LocationDocument(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal byt�������� As Byte, ByVal lng�ļ�id As Long, ByVal lngҽ��id As Long, ByVal lng����ID As Long)
    
    '������Ϣ��λ��ָ�����˵�ָ������������ȥ
    Dim rs          As ADODB.Recordset
    Dim rsTmp       As ADODB.Recordset
'    lng����id , lng��ҳID
    On Error GoTo errHand
    Set rs = gclsPackage.GetDocumentLocation(lng����ID, lng��ҳID)
    If Not rs.BOF Then
        '��������
        gstrSQL = "select b.���� || '-' || b.���� from ������ҳ a, ���ű� b where a.��Ժ����id=b.id And ����ID = [1] and ��ҳID = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, lng��ҳID)
        If Not (rsTmp.EOF Or rsTmp.BOF) Then
            If zlCommFun.NVL(rs("״̬").Value, 0) = 0 Then
                mintIndex = 1
            ElseIf zlCommFun.NVL(rs("����״ֵ̬").Value, 0) = 5 Then
                mintIndex = 0
            ElseIf zlCommFun.NVL(rs("����״ֵ̬").Value, 0) = 1 Or zlCommFun.NVL(rs("����״ֵ̬").Value, 0) = 2 Then
                mintIndex = 0
            Else
                If zlCommFun.NVL(rs("����״ֵ̬").Value, 0) > 10 Then
                    mintIndex = 1
                Else
                    mintIndex = 0
                End If
            End If
            If GetChildPatient(mintIndex).cboDept.Text <> rsTmp.Fields(0) Then
                Call GetChildPatient(mintIndex).cboDeptRefresh(rsTmp.Fields(0))
            End If
        End If
        
        If zlCommFun.NVL(rs("״̬").Value, 0) = 0 Then
                
            If tbcTask.Item(1).Selected = False And tbcTask.Item(1).Visible Then tbcTask.Item(1).Selected = True
            Call mfrmChildPatientIn.zlLocationDocument(lng����ID, lng��ҳID, byt��������, lng�ļ�id & "," & lngҽ��id & "," & lng����ID)
        
        Else
        
            Select Case zlCommFun.NVL(rs("����״ֵ̬").Value, 0)
                Case 1              'δ�ύ����״̬
                
                Case 10, 2           '����״̬
                    If tbcTask.Item(0).Selected = False And tbcTask.Item(0).Visible Then tbcTask.Item(0).Selected = True
                    Call mfrmChildPatientAduit.zlLocationDocument(lng����ID, lng��ҳID, byt��������, lng�ļ�id & "," & lngҽ��id & "," & lng����ID)
                Case 3, 4       '���״̬
                    
                    If tbcTask.Item(0).Selected = False And tbcTask.Item(0).Visible Then tbcTask.Item(0).Selected = True
                    Call mfrmChildPatientAduit.zlLocationDocument(lng����ID, lng��ҳID, byt��������, lng�ļ�id & "," & lngҽ��id & "," & lng����ID)
                    
                Case 5              '�鵵״̬
                    
                    If tbcTask.Item(0).Selected = False And tbcTask.Item(0).Visible Then tbcTask.Item(0).Selected = True
                    Call mfrmChildPatientAduit.zlLocationDocument(lng����ID, lng��ҳID, byt��������, lng�ļ�id & "," & lngҽ��id & "," & lng����ID)
                
            End Select
            
        End If
                
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'�Զ�����̻���
'######################################################################################################################
Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
        Case 0
            tbcTask.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    End Select
End Sub

Private Sub tbcTask_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

    mintIndex = Item.Index
    Call CountSelected
    
    If mintIndex = 0 Then
        Call GetChildPatient(mintIndex).InitData(Me.ShowDept)
        With GetChildPatient(mintIndex).VsfBody
            mfrmChildQuestion.AllowModify = Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 4 Or Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 6
        End With
        
    Else
        mfrmChildQuestion.AllowModify = True
         Call GetChildPatient(mintIndex).InitData(False)
    End If
    Call ExecuteCommand("��ȡ���˲���")
    Call ExecuteCommand("ˢ��״̬")
    
End Sub

Private Sub txtLocation_GotFocus()
    Call zlControl.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtLocation.Text = "" Then Exit Sub
        Call GetChildPatient(mintIndex).zlLocationPatient(2, mstrFindKey, CheckSpecialSign(txtLocation.Text), , , , , mstrFindDeal)
        
        Call LocationObj(txtLocation)
        Call CountSelected
    Else
        If InStr(":��;��?��''||", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub


'��ⳤ���Ƿ񳬹�����(�ֽ���)
Private Function ChkStrUniCode(mStr As String, mLen As Long) As String
    Dim strL        As String
On Error GoTo ErrH
    mStr = ConvertString(mStr)
    If mLen <= 0 Then
        ChkStrUniCode = mStr
        Exit Function
    Else
        strL = StrConv(mStr, vbFromUnicode)
        strL = LeftB(strL, mLen)
        ChkStrUniCode = StrConv(strL, vbUnicode)
    End If
    Exit Function
ErrH:
    Err.Clear
    ChkStrUniCode = ""
    Exit Function
End Function

'����Ƿ�����������(�Żش���ֵ)
Private Function CheckSpecialSign(ByVal mStr As String) As String
    Dim i As Integer
    If mStr = "" Then Exit Function
    If InStrRev(mStr, "'", -1) > 0 Then
        CheckSpecialSign = Replace(mStr, "'", "''")
    Else
        CheckSpecialSign = mStr
    End If
End Function

'==============================================================================
'=���ܣ� �鿴��ҳ
'==============================================================================
Private Sub RecordLook()
    
    On Error GoTo ErrH
    With GetChildPatient(mintIndex).VsfBody
        If .Row < 1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("����id"))) = 0 Then GoTo ErrH
        Call frmArchiveView.ShowArchive(Me, Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), False)
        
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'''Private Sub ����_Click() '���Զ������Ӳ����������ģ��
'''    '��ʼ�� zlRichEPR.clsChildQuestion �ӿ���
'''    Dim clsChildQuestion As zlRichEPR.clsChildQuestion
'''    Set clsChildQuestion = New zlRichEPR.clsChildQuestion
'''
'''    Call clsChildQuestion.zlOpenQuestion(Me, 727848, 1)
'''    Set clsChildQuestion = Nothing
'''End Sub
