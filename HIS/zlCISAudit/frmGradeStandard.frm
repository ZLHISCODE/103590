VERSION 5.00
Begin VB.Form frmGradeStandard 
   Caption         =   "���ֱ�׼ά��"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11295
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "frmGradeStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''///////////////////////////////////////////////////////////////////////////////
''
''       ģ�飺���ֱ�׼ά��
''       ���ܣ��������ֱ�׼��¼�롢�޸ġ�ɾ������ӡ��ѡ�õȡ�
''       ��д������ΰ
''       ���ڣ�2005��1��5��
''
''///////////////////////////////////////////////////////////////////////////////
'
'
'Option Explicit
'
'
'Private mstrPrivs As String
'Private mblnStartUp As Boolean
'Private mblnAllowClose As Boolean
'Private mlngModul As Long
'
'Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
'
'Private m_lngOldRow As Long
'Private m_lngCurRow As Long
'Private m_lngCurID As Long
'Private m_lngCurFAID As Long
'Private m_lngCurSJID As Long     '��¼��ǰ��¼ID,����ID,�ϼ�ID
'Private m_strTreeKey As String
'Private m_lngOldSJID As Long
'
'Private Function InitCommandBar() As Boolean
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim objMenu As CommandBarPopup
'    Dim objBar As CommandBar
'    Dim objPopup As CommandBarPopup
'    Dim objControl As CommandBarControl
'    Dim cbrCustom As CommandBarControlCustom
'
'    '------------------------------------------------------------------------------------------------------------------
'    '��ʼ����
'
'    Call CommandBarInit(cbsMain)
'
'    '------------------------------------------------------------------------------------------------------------------
'    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
'
'    cbsMain.ActiveMenuBar.Title = "�˵�"
'    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
'
'    '�ļ�
'    '------------------------------------------------------------------------------------------------------------------
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
'    objMenu.ID = conMenu_FilePopup
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
'
'    '�༭
'    '------------------------------------------------------------------------------------------------------------------
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
'    objMenu.ID = conMenu_EditPopup
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewKind, "���ӷ���(&N)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyKind, "�޸ķ���(&F)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteKind, "ɾ������(&L)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Import, "���뷽��(&P)...")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Select, "ѡ�÷���(&S)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewParent, "������Ŀ(&X)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Insert, "������Ŀ(&R)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyParent, "�޸���Ŀ(&G)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteParent, "ɾ����Ŀ(&C)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "���ӱ�׼(&A)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Append, "�����׼(&I)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "�޸ı�׼(&M)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ����׼(&D)")
'
'    '�鿴
'    '------------------------------------------------------------------------------------------------------------------
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
'    objMenu.ID = conMenu_ViewPopup
'    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
'    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
'    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
'    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
'
'
'    '����
'    '------------------------------------------------------------------------------------------------------------------
'    Call CreateHelpMenu(cbsMain)
'
'    '����������:������������
'    '------------------------------------------------------------------------------------------------------------------
'    Set objBar = cbsMain.Add("��׼", xtpBarTop)
'    objBar.ContextMenuPresent = False
'    objBar.ShowTextBelowIcons = False
'    objBar.EnableDocking xtpFlagStretched
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
'
'    Set objPopup = NewToolBar(objBar, xtpControlPopup, conMenu_Edit_NewKind * 10# + 1, "����", True, , , , objControl.Index + 1)
'    objPopup.ID = conMenu_Edit_NewKind
'    objPopup.IconId = conMenu_Edit_NewParent
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_NewKind, "���ӷ���(&A)")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_ModifyKind, "�޸ķ���(&D)")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_DeleteKind, "ɾ������(&D)")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Select, "ѡ�÷���(&S)")
'
'    Set objPopup = NewToolBar(objBar, xtpControlPopup, conMenu_Edit_NewParent * 10# + 1, "��Ŀ", True, , , , objControl.Index + 1)
'    objPopup.ID = conMenu_Edit_NewParent
'    objPopup.IconId = conMenu_Edit_NewParent
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_NewParent, "������Ŀ(&A)")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸���Ŀ(&D)")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_DeleteParent, "ɾ����Ŀ(&D)")
'
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "����", True)
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "�޸�")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ��")
'
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")
'
'    '����Ŀ����:���������������Ѵ���
'    '------------------------------------------------------------------------------------------------------------------
'    With cbsMain.KeyBindings
'        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
'        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
'        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
'    End With
'
'End Function
'
'
'Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim intLoop As Integer
'    Dim intRow As Integer
'    Dim Rs As New ADODB.Recordset
'    Dim rsSQL As New ADODB.Recordset
'    Dim strTmp As String
'    Dim strSql As String
'
'    On Error GoTo errHand
'
'    Call SQLRecord(rsSQL)
'
'    Select Case strCommand
'    '------------------------------------------------------------------------------------------------------------------
'    Case "��ʼ�ؼ�"
'
'
'        '��ʼ�˵���������
'        '--------------------------------------------------------------------------------------------------------------
'        Call InitCommandBar
'
'        '����ͣ������
'        '--------------------------------------------------------------------------------------------------------------
'        Dim objPane As Pane
'        Set objPane = dkpMain.CreatePane(1, 100, 200, DockLeftOf, Nothing): objPane.Title = "���ַ���": objPane.Options = PaneNoCaption
'        Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, Nothing): objPane.Title = "���ֱ�׼": objPane.Options = PaneNoCaption
'        Set objPane = dkpMain.CreatePane(3, 100, 100, DockBottomOf, objPane): objPane.Title = "��Ŀ��Ϣ": objPane.Options = PaneNoCaption
'
'        dkpMain.SetCommandBars cbsMain
'        Call DockPannelInit(dkpMain)
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "��ʼ����"
'
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "�ؼ�״̬"
'
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "ˢ��״̬"
'
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "ˢ������"
'
'        '���Tree
'        Call FillTree
'
'        '����б�
'        Call Fill���
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "��ע���"
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "дע���"
'
'    End Select
'
'    ExecuteCommand = True
'
'    Exit Function
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call SaveErrLog
'
'End Function
'
'Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    Dim objControl As CommandBarControl
'    Dim lngLoop As Long
'
'    Select Case Control.ID
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_File_Parameter
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case Else
'
'        If Control.ID > 400 And Control.ID < 500 Then
'            Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me, "���ַ���=" & m_lngCurFAID)
'        Else
'             '��ҵ���޹صĹ��ܣ������Ĺ���
'            Call CommandBarExecutePublic(Control, Me, fgMain, "���ַ�����׼����")
'        End If
'
'    End Select
'End Sub
'
'Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
'    If stbThis.Visible Then Bottom = stbThis.Height
'End Sub
'
'Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    On Error GoTo errHand
'
'    With fgMain
'        Select Case Control.ID
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel
'            Control.Enabled = .Row > 0
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_EditPopup, conMenu_Edit_NewKind, conMenu_Edit_NewParent, conMenu_Edit_Insert, conMenu_Edit_NewItem, conMenu_Edit_Append
'            Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_Edit_ModifyKind, conMenu_Edit_DeleteKind, conMenu_Edit_Select
'            Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
'            Control.Enabled = Control.Visible And Not (tvw����.SelectedItem Is Nothing)
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_Edit_Import
'            Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
'            If tvw����.SelectedItem Is Nothing Then
'                Control.Enabled = False
'            Else
'                Control.Enabled = Control.Visible And tvw����.Nodes.Count > 1
'            End If
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_Edit_ModifyParent, conMenu_Edit_DeleteParent
'            Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
'            Control.Enabled = Control.Visible And .Row > 0
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_Edit_Modify, conMenu_Edit_Delete
'            Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
'            Control.Enabled = Control.Visible And .Row > 0
'        '--------------------------------------------------------------------------------------------------------------
'        Case Else
'            Call CommandBarUpdatePublic(Control, Me)
'        End Select
'    End With
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'End Sub
'
'Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
'    Select Case Item.ID
'    Case 1
'        Item.Handle = picPane(0).Hwnd
'    Case 2
'        Item.Handle = picPane(1).Hwnd
'    Case 3
'        Item.Handle = picPane(2).Hwnd
'    End Select
'End Sub
'
'Private Sub fgMain_Click()
'    fgMain_SelChange
'End Sub
'
'Private Sub fgMain_DblClick()
'    If fgMain.MouseRow = 0 Then Exit Sub
'    Call fgMain_KeyPress(13)
'End Sub
'
'Private Sub fgMain_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'
'        If IsPrivs(mstrPrivs, "��ɾ��") Then
'            If fgMain.Row > 0 And fgMain.TextMatrix(fgMain.Row, 2) <> "" Then
'                Call mnuEditModBZ_Click
'            End If
'        End If
'    End If
'End Sub
'
''Private Sub fgMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    If InStr(gstrPrivs, "��ɾ��") = 0 Then Exit Sub
''    If Button = vbRightButton Then
''        If fgMain.MouseRow = -1 And fgMain.Rows >= 1 Then
''            fgMain.Row = fgMain.Rows - 1
''        ElseIf fgMain.MouseRow = 0 And fgMain.Rows > 1 Then
''            fgMain.Row = 1
''        Else
''            fgMain.Row = fgMain.MouseRow
''        End If
''        fgMain.Col = fgMain.MouseCol
''
''        m_lngCurSJID = IIf(Len(fgMain.Cell(flexcpText, fgMain.Row, 5)) = 0, 0, Val(fgMain.Cell(flexcpText, fgMain.Row, 5)))      '��ȡID
''
''        PopupMenu mnuShortEdit
''    End If
''End Sub
'
'Private Sub fgMain_SelChange()
'    Dim lngID As Long
'    m_lngCurRow = fgMain.Row
'    If m_lngCurRow < 0 Then m_lngCurSJID = 0: m_lngCurID = 0: Exit Sub
'    m_lngCurID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 4)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 4)))    '��ȡID
'    m_lngCurSJID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 5)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 5)))     '��ȡID
'    m_lngCurFAID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 6)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 6)))     '��ȡID
'    If m_lngCurSJID = 0 Then
'        lngID = m_lngCurID
'    Else
'        lngID = m_lngCurSJID
'    End If
'
'    Show����Ҫ�� lngID, fgMain.Cell(flexcpText, m_lngCurRow, 0), fgMain.Cell(flexcpText, m_lngCurRow, 1)
'    m_lngOldRow = m_lngCurRow
'    SetMenu
'End Sub
'
'Private Sub Form_Activate()
'    If mblnStartUp = False Then Exit Sub
'    mblnStartUp = False
'    DoEvents
'
'    If ExecuteCommand("��ʼ����") = False Then GoTo errHand
'
'    Call ExecuteCommand("ˢ������")
'
'    mblnAllowClose = True
'    Exit Sub
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mblnAllowClose = True
'    Unload Me
'End Sub
'
'Private Sub Form_Initialize()
'    Call InitCommonControls
'End Sub
'
'Private Sub Form_Load()
'
'    mblnStartUp = True
'    mblnAllowClose = False
'
'    mstrPrivs = UserInfo.ģ��Ȩ��
'    mlngModul = ParamInfo.ģ���
'
'    Call ExecuteCommand("��ʼ�ؼ�")
'    Call ExecuteCommand("��ע���")
'
'    Call RestoreWinState(Me, App.ProductName)
'    Call zlCommFun.SetWindowsInTaskBar(Me.Hwnd, gblnShowInTaskBar)
'
'
'    Me.KeyPreview = True
'    m_lngOldRow = -1
'    m_lngCurRow = -1
'    m_lngCurID = -1
'    m_lngOldSJID = -1
'
'    'Ȩ�޿���
''    Call Ȩ�޿���
''    '���Tree
''    Call FillTree
''
''    '����б�
''    Call Fill���
'
'    '�ָ�����λ��
'
''    RestoreWinState Me, App.ProductName
''    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
''    Call SetMenu
'
'    picFAXX.Picture = imgClose.Picture
'End Sub
'
'Private Sub Form_Resize()
'    On Error Resume Next
'
'    Call SetPaneRange(dkpMain, 1, 100, 100, 250, Me.ScaleHeight)
'    Call SetPaneRange(dkpMain, 3, 100, 100, Me.ScaleWidth, 200)
'    dkpMain.RecalcLayout
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    m_strTreeKey = ""
'    SaveWinState Me, App.ProductName
'End Sub
'
'Private Sub picFAXX_Click()
'    If picFAXX.Tag = "" Then
'        picFAXX.Tag = "Opened"
'        picFAXX.Picture = imgOpen.Picture
'        pic������Ϣ.Height = 340
'    Else
'        picFAXX.Tag = ""
'        picFAXX.Picture = imgClose.Picture
'        pic������Ϣ.Height = 1695
'    End If
'    picFAXX.Refresh
'    Call picPane_Resize(0)
'End Sub
'
'Private Sub mnuEditDelBZ_Click()
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'
'    On Error GoTo errHandle
'
'    Dim intIndex As Long
'
'    If m_lngCurID < 1 Then Exit Sub
'    If MsgBox("��ȷ��Ҫɾ���������ֱ�׼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'    gstrSQL = "ZL_�������ֱ�׼_Delete(" & CStr(m_lngCurID) & ",1)"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'    Call Fill���
'    Call SetMenu
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub
'
'Private Sub mnuEditDelFA_Click()
'    'ɾ�����ַ���
'    On Error GoTo errHandle
'    Dim intIndex As Long
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'
'    If m_lngCurFAID < 1 Then Exit Sub
'
'    If MsgBox("��ȷ��Ҫɾ������������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'
'    gstrSQL = "ZL_�������ַ���_Delete(" & CStr(m_lngCurFAID) & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'    Call FillTree
'    Call Fill���
'    Call SetMenu
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub
'
'Private Sub mnuEditDelXM_Click()
'     m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'
'    On Error GoTo errHandle
'    Dim intIndex As Long
'
'    If m_lngCurID < 1 Then Exit Sub
'
'    If MsgBox("��ȷ��Ҫɾ������������Ŀ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'
'    If m_lngCurSJID = 0 Then
'        gstrSQL = "ZL_�������ֱ�׼_Delete(" & CStr(m_lngCurID) & ",0)"
'    Else
'        gstrSQL = "ZL_�������ֱ�׼_Delete(" & CStr(m_lngCurSJID) & ",0)"
'    End If
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'    Call Fill���
'    Call SetMenu
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub
'
'Private Sub mnuEditEmportFA_Click()
'    '�������з���
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'    If m_lngCurFAID <= 0 Then Exit Sub
'
'    Dim lID As Long     'ѡ�еķ���ID
'    Dim lNewID As Long
'
'    Dim f As New frmѡ�����ַ���
'    f.FillCmbSelFA m_lngCurFAID
'    f.Show 1
'    lID = f.ID_From
'
'    'ִ�е������������
'    'ԴIDΪ��lID   Ŀ��IDΪ�� m_lngCurFAID
'    Dim Rs As New ADODB.Recordset, lng�ܷ� As Double, rsTmp As New ADODB.Recordset
'    Dim strT As String
'    gstrSQL = "select * from �������ֱ�׼ where �ϼ�ID is null and ����ID=" & lID & " order by �ϼ����,���,ID"
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    zlCommFun.ShowFlash "���Ժ�ϵͳ���ڵ������ַ�������", Me
'    DoEvents
'    On Error GoTo LL
'    gcnOracle.BeginTrans
'    Do While Not Rs.EOF
'        '�ҵ�����Ŀ�������Ŀ
'        lNewID = zlDatabase.GetNextId("�������ֱ�׼")
'        gstrSQL = "ZL_�������ֱ�׼_Insert" & _
'            "(" & lNewID & _
'            "," & NVL(Rs("�ϼ�ID"), "NULL") & _
'            "," & m_lngCurFAID & _
'            ",'" & NVL(Rs("����")) & _
'            "','" & NVL(Rs("����")) & _
'            "'," & NVL(Rs("��׼��ֵ"), "NULL") & _
'            ",'" & NVL(Rs("ȱ�ݵȼ�")) & _
'            "','" & NVL(Rs("���ֵ�λ")) & "',0)"
'        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'        '��һ�������¼���Ŀ��ѭ�����֮��
'        gstrSQL = "select * from �������ֱ�׼ where �ϼ�ID=" & Rs("ID") & " and ����ID=" & lID & " order by �ϼ����,���,ID"
'        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
'        Do While Not rsTmp.EOF
'            gstrSQL = "ZL_�������ֱ�׼_Insert" & _
'                "(" & zlDatabase.GetNextId("�������ֱ�׼") & _
'                "," & lNewID & _
'                "," & m_lngCurFAID & _
'                ",'" & NVL(rsTmp("����")) & _
'                "','" & NVL(rsTmp("����")) & _
'                "'," & NVL(rsTmp("��׼��ֵ"), "NULL") & _
'                ",'" & NVL(rsTmp("ȱ�ݵȼ�")) & _
'                "','" & NVL(rsTmp("���ֵ�λ")) & "',0)"
'            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'            rsTmp.MoveNext
'        Loop
'        Rs.MoveNext
'    Loop
'
'    'ˢ�½����
'    gcnOracle.CommitTrans
'
'    Call Fill���
'    zlCommFun.StopFlash
'    Exit Sub
'LL:
'    gcnOracle.RollbackTrans
'    zlCommFun.StopFlash
'End Sub
'
'Private Sub mnuEditInsBZ_Click()
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'    Dim f As New frm���ֱ�׼�༭
'    If m_lngCurSJID < 1 Then 'Ϊ����������
'        f.ShowForm "����", m_lngCurFAID, m_lngCurID, m_lngCurSJID
'    Else
'        f.ShowForm "����", m_lngCurFAID, m_lngCurSJID, m_lngCurID
'    End If
'
'    Call ˢ�·�����Ϣ
'    If f.Moded Then
'        Call Fill���
'    End If
'End Sub
'
'Private Sub mnuEditInsXM_Click()
'    '����������Ŀ
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'    Dim f As New frm���ֱ�׼�༭
'    If m_lngCurSJID < 1 Then 'Ϊ����������
'        f.ShowForm "����", m_lngCurFAID, 0, m_lngCurID
'    Else
'        f.ShowForm "����", m_lngCurFAID, 0, m_lngCurSJID
'    End If
'    Call ˢ�·�����Ϣ
'    If f.Moded Then
'        Call Fill���
'    End If
'End Sub
'
'Private Sub mnuEditModBZ_Click()
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'    '�޸����ֱ�׼
'    If m_lngCurID < 1 Then Exit Sub
'    Dim f As New frm���ֱ�׼�༭
'    If fgMain.Col < 2 Then  'һ����Ŀ
'        If m_lngCurSJID < 1 Then
'            f.ShowForm "�޸�", m_lngCurFAID, , m_lngCurID
'        Else
'            f.ShowForm "�޸�", m_lngCurFAID, , m_lngCurSJID
'        End If
'    Else                    '����Ŀ
'        f.ShowForm "�޸�", m_lngCurFAID, m_lngCurSJID, m_lngCurID
'    End If
'    Call ˢ�·�����Ϣ
'    If f.Moded Then
'        Call Fill���
'    End If
'End Sub
'
'Private Sub mnuEditModFA_Click()
'    '�޸����ַ���
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'    If m_lngCurFAID < 1 Then Exit Sub
'    Dim f As New frm���ַ����༭, lng�ܷ� As Double
'    f.ShowForm m_lngCurFAID   '�޸ģ�����ID
'    Call ˢ�·�����Ϣ
'    If f.Moded Then
'        Call FillTree
'        '����б�
'        Call Fill���
'    End If
'End Sub
'
'Private Sub mnuEditModXM_Click()
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'    If m_lngCurID < 1 Then Exit Sub
'    Dim f As New frm���ֱ�׼�༭
'    If m_lngCurSJID < 1 Then
'        f.ShowForm "�޸�", m_lngCurFAID, , m_lngCurID
'    Else
'        f.ShowForm "�޸�", m_lngCurFAID, , m_lngCurSJID
'    End If
'    Call ˢ�·�����Ϣ
'    If f.Moded Then
'        Call Fill���
'    End If
'End Sub
'
'Private Sub mnuEditNewBZ_Click()
'    '�����¼����ֱ�׼
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'    Dim f As New frm���ֱ�׼�༭
'    If m_lngCurSJID < 1 Then 'Ϊ����������
'        f.ShowForm "����", m_lngCurFAID, m_lngCurID
'    Else
'        f.ShowForm "����", m_lngCurFAID, m_lngCurSJID
'    End If
'
'    Call ˢ�·�����Ϣ
'    If f.Moded Then
'        Call Fill���
'    End If
'
'End Sub
'
'Private Sub mnuEditNewFA_Click()
'    '�������ַ���
'    Dim f As New frm���ַ����༭
'    f.ShowForm   '����
'    Call ˢ�·�����Ϣ
'    If f.Moded Then
'        Call FillTree
'        '����б�
'        Call Fill���
'    End If
'End Sub
'
'Private Sub ˢ�·�����Ϣ()
'    Dim Rs As New ADODB.Recordset, lng�ܷ� As Double
'    gstrSQL = "select * from �������ַ��� where ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    If Not Rs.EOF Then
'        lbl�������� = Rs("����")
'        lbl���� = "����:" & Rs("����")
'        lbl��ֵ = "��ֵ:" & Rs("��ֵ")
'        lbl��ֵ = "��ֵ:" & Rs("��ֵ")
'        lbl�ܷ� = "�ܷ�:" & Rs("�ܷ�")
'        lng�ܷ� = Rs("�ܷ�")
'    Else
'        lbl�������� = ""
'        lbl���� = ""
'        lbl��ֵ = ""
'        lbl��ֵ = ""
'        lbl�ܷ� = ""
'    End If
'
'    Rs.Close
'    gstrSQL = "select sum(��׼��ֵ) from �������ֱ�׼ where �ϼ�ID is null and ����ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    If Not Rs.EOF Then
'        If Abs(lng�ܷ� - Rs.Fields(0)) > 0.01 Then
'            lbl�ܷ� = lbl�ܷ� + "����Ŀ������Ϊ:" & Rs.Fields(0)
'            lbl�ܷ�.ForeColor = vbRed
'        Else
'            lbl�ܷ�.ForeColor = vbBlack
'        End If
'    Else
'        lbl�ܷ�.ForeColor = vbRed
'    End If
'End Sub
'
'Private Sub mnuEditNewXM_Click()
'    '����������Ŀ
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
'    Dim f As New frm���ֱ�׼�༭
'    f.ShowForm "����", m_lngCurFAID
'    Call ˢ�·�����Ϣ
'    If f.Moded Then
'        Call Fill���
'    End If
'End Sub
'
'Private Sub mnuEditSelFA_Click()
'    On Error GoTo errHandle
'    Dim intIndex As Long, bln��ʹ�� As Boolean
'
'    If m_lngCurFAID < 1 Then Exit Sub
'    If MsgBox("ע�⣺���ְַ���ѡ����һ���ǳ����ص����飬ͨ����Ҫ������ģ�" & vbCrLf & "��ȷ��ѡ�ñ����ַ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'
'    Dim rsTemp As New ADODB.Recordset
'    gstrSQL = "select count(*) from �������ֽ�� where ����ID=(select ID from �������ַ��� where ����='סԺ' and ѡ��=1)"
'    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
'    If rsTemp(0).Value > 0 Then
'        'Ĭ��סԺ�����Ѿ�ʹ��
'        If MsgBox("ע�⣺ϵͳĬ�����ְַ�����ʹ�õ��У��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'    End If
'    rsTemp.Close
'
'    gstrSQL = "ZL_�������ַ���_ѡ��(" & CStr(m_lngCurFAID) & ",1)"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'    Call FillTree
'    Call SetMenu
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub
'
'Private Sub mnuFileSetup_Click()
'    frm���ֱ�׼��������.Show 1
'End Sub
'
'Private Sub mnuHelpAbout_Click()
'    '���ڶԻ���
'    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
'End Sub
'
'Private Sub mnuShortMenuBZ_Click(Index As Integer)
'    '�����˵�����
'    Select Case Index
'        Case 1
'            mnuEditNewBZ_Click
'        Case 2
'            mnuEditInsBZ_Click
'        Case 3
'            mnuEditModBZ_Click
'        Case 4
'            mnuEditDelBZ_Click
'    End Select
'End Sub
'
'Private Sub mnuShortMenuFA_Click(Index As Integer)
'    '�����˵�����
'    Select Case Index
'        Case 1
'            mnuEditNewFA_Click
'        Case 2
'            mnuEditModFA_Click
'        Case 3
'            mnuEditDelFA_Click
'        Case 4
'            mnuEditSelFA_Click
'        Case 5
'            mnuEditEmportFA_Click
'    End Select
'End Sub
'
'Private Sub mnuShortMenuXM_Click(Index As Integer)
'    Select Case Index
'        Case 1
'            mnuEditNewXM_Click
'        Case 2
'            mnuEditInsXM_Click
'        Case 3
'            mnuEditModXM_Click
'        Case 4
'            mnuEditDelXM_Click
'    End Select
'End Sub
'
'Private Sub mnuShortMnuXM_Click(Index As Integer)
'    Select Case Index
'        Case 1
'            mnuEditNewXM_Click
'        Case 2
'            mnuEditInsXM_Click
'        Case 3
'            mnuEditModXM_Click
'        Case 4
'            mnuEditDelXM_Click
'    End Select
'End Sub
'
'Private Sub mnuViewRefresh_Click()
'    'ˢ��TreeView
'    Call FillTree
'End Sub
'
'Private Sub mnuFileExit_Click()
'    '�رմ���
'    Unload Me
'End Sub
'
'Private Sub mnuFileExcel_Click()
'    '�����Excel
'    '1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    subPrint 3
'End Sub
'
'Private Sub mnufilepre_Click()
'    'Ԥ��
'    '1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    subPrint 2
'End Sub
'
'Private Sub mnuFilePrint_Click()
'    '��ӡ
'    '1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    subPrint 1
'End Sub
'
'Private Sub mnufileset_Click()
'    '��ӡ���� ��zlPrintMethod��
'    zlPrintSet
'End Sub
'
'Private Sub picFAXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If X >= 0 And X <= picFAXX.ScaleWidth And Y >= 0 And Y <= picFAXX.ScaleHeight Then
'        SetCapture picFAXX.Hwnd
'        '������룡����
'        picFAXX.Line (0, 0)-(picFAXX.ScaleWidth - Screen.TwipsPerPixelX, picFAXX.ScaleHeight - Screen.TwipsPerPixelY), vbBlue, B
'    Else
'        '����Ƴ�������
'        picFAXX.Cls
'        ReleaseCapture
'    End If
'End Sub
'
'Private Sub picPane_Resize(Index As Integer)
'    Select Case Index
'    Case 0
'
'        pic������Ϣ.Move 135, picPane(Index).ScaleHeight - pic������Ϣ.Height - 270, picPane(Index).ScaleWidth - 270
'        picTree.Move 135, 135, pic������Ϣ.Width, Abs(picPane(Index).ScaleHeight - pic������Ϣ.Height - 270 * 2)
'        picTree.Cls
'        picTree.PaintPicture imgBGBlue.Picture, 0, 0, picTree.Width, 360, 0, 0, imgBGBlue.Width, 360
'        picTree.PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, picTree.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
'        picTree.PaintPicture imgBGBlue.Picture, picTree.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picTree.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
'        picTree.PaintPicture imgBGBlue.Picture, 0, picTree.ScaleHeight - Screen.TwipsPerPixelY, picTree.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
'
'        tvw����.Move Screen.TwipsPerPixelX * 4, 390, Abs(picTree.ScaleWidth - 8 * Screen.TwipsPerPixelX), Abs(picTree.ScaleHeight - 390 - Screen.TwipsPerPixelY * 4)
'
'        pic������Ϣ.Cls
'        pic������Ϣ.PaintPicture imgBG.Picture, 0, 0, pic������Ϣ.Width, 360, 0, 0, imgBG.Width, 360
'        pic������Ϣ.PaintPicture imgBG.Picture, 0, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
'        pic������Ϣ.PaintPicture imgBG.Picture, pic������Ϣ.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, imgBG.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
'        pic������Ϣ.PaintPicture imgBG.Picture, 0, pic������Ϣ.ScaleHeight - Screen.TwipsPerPixelY, pic������Ϣ.Width, Screen.TwipsPerPixelY, 0, imgBG.Height - Screen.TwipsPerPixelY, imgBG.Width, Screen.TwipsPerPixelY
'        picFAXX.Move pic������Ϣ.ScaleWidth - picFAXX.Width - 100
'
'        Refresh
'
'    Case 1
'        fgMain.Move 0, 0, picPane(Index).Width, picPane(Index).Height
'    Case 2
'        lblInfo.Move lblInfo.Left, lblInfo.Top, Abs(picPane(Index).ScaleWidth - 2 * lblInfo.Left), Abs(picPane(Index).ScaleHeight - lblInfo.Top)
'    End Select
'End Sub
'
''Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''    '��������ť�¼�
''    Select Case Button.Key
''        Case "FA"
''            PopupMenu mnuShortFA, , Button.Left + 45, Button.Top + Button.Height + 45
''        Case "XM"
''            PopupMenu mnuShortXM, , Button.Left + 45, Button.Top + Button.Height + 45
''        Case "NewBZ"
''            mnuEditNewBZ_Click
''        Case "ModBZ"
''            mnuEditModBZ_Click
''        Case "DelBZ"
''            mnuEditDelBZ_Click
''        Case "Quit"
''            mnuFileExit_Click
''        Case "Print"
''            mnuFilePrint_Click
''        Case "Preview"
''            mnufilepre_Click
''        Case "Help"
''            mnuHelpTitle_Click
''    End Select
''
''End Sub
'
''Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
''    If ButtonMenu.Parent.Key = "FA" Then
''        Select Case ButtonMenu.Index
''        Case 1
''            mnuEditNewFA_Click
''        Case 2
''            mnuEditModFA_Click
''        Case 3
''            mnuEditDelFA_Click
''        Case 4
''            mnuEditSelFA_Click
''        End Select
''    Else
''        Select Case ButtonMenu.Index
''        Case 1
''            mnuEditNewXM_Click
''        Case 2
''            mnuEditModXM_Click
''        Case 3
''            mnuEditDelXM_Click
''        End Select
''    End If
''End Sub
'
'Private Sub picTree_DblClick()
'    If Left(tvw����.SelectedItem.Key, 4) = "Root" Then Exit Sub
'    mnuEditModFA_Click
'End Sub
'
'Private Sub picTree_KeyPress(KeyAscii As Integer)
'    If IsNumeric(Mid(m_strTreeKey, 2)) Then
'        If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call mnuEditModFA_Click
'    End If
'End Sub
'
''Private Sub tvw����_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    If InStr(gstrPrivs, "��ɾ��") = 0 Then Exit Sub
''    If Button = vbRightButton Then
''        PopupMenu mnuShortFA
''    End If
''End Sub
'
'Private Sub tvw����_NodeClick(ByVal Node As MSComctlLib.Node)
'On Error Resume Next
'    If m_strTreeKey = Node.Key Then Exit Sub     '�����ظ�ˢ��
'    m_strTreeKey = Node.Key
'    m_lngCurFAID = Val(Mid(m_strTreeKey, 2))
'    Dim Rs As New ADODB.Recordset, lng�ܷ� As Double
'    gstrSQL = "select * from �������ַ��� where ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    If Not Rs.EOF Then
'        lbl�������� = Rs("����")
'        lbl���� = "����:" & Rs("����")
'        lbl��ֵ = "��ֵ:" & Rs("��ֵ")
'        lbl��ֵ = "��ֵ:" & Rs("��ֵ")
'        lbl�ܷ� = "�ܷ�:" & Rs("�ܷ�")
'        lng�ܷ� = Rs("�ܷ�")
'    Else
'        lbl�������� = ""
'        lbl���� = ""
'        lbl��ֵ = ""
'        lbl��ֵ = ""
'        lbl�ܷ� = ""
'    End If
'
'    Rs.Close
'    gstrSQL = "select sum(��׼��ֵ) from �������ֱ�׼ where �ϼ�ID is null and ����ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    If Not Rs.EOF Then
'        If Abs(lng�ܷ� - Rs.Fields(0)) > 0.01 Then
'            lbl�ܷ� = lbl�ܷ� + "����Ŀ������Ϊ:" & Rs.Fields(0)
'            lbl�ܷ�.ForeColor = vbRed
'        Else
'            lbl�ܷ�.ForeColor = vbBlack
'        End If
'    Else
'        lbl�ܷ�.ForeColor = vbRed
'    End If
'    '����б�
'    Call Fill���
'End Sub
'
'Private Sub subPrint(bytMode As Byte)
'    '-------------------------------------------------
'    '����:�����ݱ���д�ӡ,Ԥ���������EXCEL
'    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    '-------------------------------------------------
'    Dim objPrint As New zlPrint1Grd
'    Dim objAppRow As zlTabAppRow
'    Dim bytR As Byte
'
'    Set objPrint.Body = fgMain
'    objPrint.Title.Text = tvw����.SelectedItem.Text
'    objPrint.Title.Font.Name = "����_GB2312"
'    objPrint.Title.Font.Size = 18
'    objPrint.Title.Font.Bold = True
'
'    Set objAppRow = New zlTabAppRow
'    Dim Rs As New ADODB.Recordset
'    gstrSQL = "select * from �������ַ��� where ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    If Not Rs.EOF Then
'        objAppRow.Add "�ܷ�:" & NVL(Rs("�ܷ�"), 0)
'        objAppRow.Add "�׼�������:" & NVL(Rs("��ֵ"), 0)
'        objAppRow.Add "�Ҽ�������:" & NVL(Rs("��ֵ"), 0)
'
'        objPrint.UnderAppRows.Add objAppRow
'    End If
'
'    Set objAppRow = New zlTabAppRow
'    objAppRow.Add "��ӡ�ˣ�" & gstrUserName
'    objAppRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
'    objPrint.BelowAppRows.Add objAppRow
'
'    If bytMode = 1 Then
'        bytR = zlPrintAsk(objPrint)
'        If bytR <> 0 Then zlPrintOrView1Grd objPrint, bytR
'    Else
'        zlPrintOrView1Grd objPrint, bytMode
'    End If
'
'End Sub
'
'
'Private Sub FillTree()
'    '����:װ�����ַ��� Ŀǰֻ����סԺ����
'    Dim rsTemp As New ADODB.Recordset
'    Dim nod As Node, i As Long, FirstKey As String
'    rsTemp.CursorLocation = adUseClient
'
'    fgMain.Tag = ""
'    'Tree�ĳ�ʼ��
'    tvw����.Nodes.Clear
'    '��Ӹ��ڵ�
''    Set nod = tvw����.Nodes.Add(, , "Root", "���ַ����б�", "Root", "Root")
''    nod.Expanded = True
'
'    'ע����ø�ʽ���ȸ�ֵgstrSQL,Ȼ������ݼ�
'    gstrSQL = "select ID,����,ѡ�� from �������ַ��� where ����='סԺ' Order by ѡ�� desc,����,����ʱ��"
'    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
'
'    i = 1
'    Do Until rsTemp.EOF
'        '����ӽڵ�
''        Set nod = tvw����.Nodes.Add("Root", tvwChild, "A" & rsTemp("ID"), rsTemp("����"), "Child", "Child")
'        Set nod = tvw����.Nodes.Add(, , "A" & rsTemp("ID"), rsTemp("����"), IIf(rsTemp("ѡ��") = 1, "RootSel", "Root"), IIf(rsTemp("ѡ��") = 1, "RootSel", "Root"))
'        If rsTemp("ѡ��") = 1 Then
'            nod.Bold = True
'        Else
'            nod.Bold = False
'        End If
'        If i = 1 Then FirstKey = nod.Key
'        If FirstKey = nod.Key Then i = 2
'        If FirstKey = "" And i = 1 Then FirstKey = nod.Key: i = 2
'        rsTemp.MoveNext
'    Loop
''    '��Ӹ��ڵ�
''    Set nod = tvw����.Nodes.Add(, tvwNext, "RootMZ", "�������ַ���", "B", "B")
''    nod.Expanded = True
''
''    'ע����ø�ʽ���ȸ�ֵgstrSQL,Ȼ������ݼ�
''    gstrSQL = "select ID,����,ѡ�� from �������ַ��� where ����='����' Order by ѡ�� desc,����,����ʱ��"
''    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
''    i = 1
''
''    Do Until rsTemp.EOF
''        '����ӽڵ�
''        Set nod = tvw����.Nodes.Add("RootMZ", tvwChild, "B" & rsTemp("ID"), IIf(rsTemp("ѡ��") = 1, "��", "") + rsTemp("����"), "C", "C")
''        rsTemp.MoveNext
''    Loop
'    If i = 1 Then m_strTreeKey = FirstKey   'm_strTreeKey��Ϊ�գ�������û���ҵ���
'    Dim v As Variant
'    For Each v In tvw����.Nodes
'        If v.Key = FirstKey Then
'            '����ѡ��
'            v.Selected = True
'            v.EnsureVisible
'            If picTree.Visible = True Then picTree.SetFocus
'        End If
'    Next
'    tvw����_NodeClick tvw����.SelectedItem
'
'End Sub
'
'Public Sub Fill���()
'    '����:װ���Ӧ���������ֱ�׼
'    Dim rsTemp As New ADODB.Recordset
'    With fgMain
'        .Redraw = flexRDNone
'        .Rows = 1
'        .Clear
'        Dim i As Long
'        .Cell(flexcpText, 0, 0) = "��Ŀ"
'        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
'        .Cell(flexcpText, 0, 1) = "��׼��ֵ"
'        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
'        .Cell(flexcpText, 0, 2) = "ȱ������"
'        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
'        .Cell(flexcpText, 0, 3) = "���ֱ�׼"
'        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
'        .Cell(flexcpText, 0, 4) = "ID"
'        .Cell(flexcpText, 0, 5) = "�ϼ�ID"
'        .Cell(flexcpText, 0, 6) = "����ID"
'        .Cell(flexcpText, 0, 7) = "���"
'        rsTemp.CursorLocation = adUseClient
'
'        'ȷ����������
'        If tvw����.SelectedItem Is Nothing Then .Redraw = flexRDDirect: Exit Sub
'        With tvw����.SelectedItem
'            Select Case Left(.Key, 1)
'                Case "A", "B"
'                    m_lngCurFAID = Val(Mid(.Key, 2))
'                    gstrSQL = "select * from �������ֱ�׼��ͼ Where ����='��' and ����ID=" & CStr(Mid(.Key, 2))
'                Case Else
'                    Call SetMenu
'                    fgMain.Redraw = flexRDDirect
'                    Exit Sub
'            End Select
'        End With
'        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
'
'        .FocusRect = flexFocusSolid
'        '��������
'        .Cols = 8
'        .Rows = rsTemp.RecordCount + 1
'        i = 1
'        Do Until rsTemp.EOF
'            .Cell(flexcpText, i, 0) = NVL(rsTemp.Fields("��Ŀ"))
'            .Cell(flexcpAlignment, i, 0) = flexAlignCenterCenter
'            .Cell(flexcpText, i, 1) = IIf(IsNull(rsTemp.Fields("��׼��ֵ")), " ", Format(rsTemp.Fields("��׼��ֵ"), "####��"))
'            .Cell(flexcpAlignment, i, 1) = flexAlignCenterCenter
'            .Cell(flexcpText, i, 2) = NVL(rsTemp.Fields("ȱ������"))
'            .Cell(flexcpAlignment, i, 2) = flexAlignLeftTop
'            .Cell(flexcpText, i, 3) = IIf(IsNull(rsTemp.Fields("�۷ֱ�׼")), "", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "�׼�", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "�Ҽ�", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "����", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "������", rsTemp.Fields("�۷ֱ�׼"))))))
'            .Cell(flexcpAlignment, i, 3) = flexAlignCenterCenter
'            .Cell(flexcpText, i, 4) = NVL(rsTemp.Fields("ID"), 0)
'            .Cell(flexcpText, i, 5) = NVL(rsTemp.Fields("�ϼ�ID"), 0)
'            .Cell(flexcpText, i, 6) = NVL(rsTemp.Fields("����ID"), 0)
'            .Cell(flexcpText, i, 7) = NVL(rsTemp.Fields("���"), 0)
'            rsTemp.MoveNext
'            i = i + 1
'        Loop
'
'
'        '�Զ�����
'        .WordWrap = True
'        '�ϲ���Ԫ��
'        .MergeCells = 2
'        .MergeCol(.ColIndex("��Ŀ")) = True
'        .MergeCol(.ColIndex("��׼��ֵ")) = True
'        '��������
'        .ColAlignment(.ColIndex("��Ŀ")) = flexAlignLeftCenter
'        .ColAlignment(.ColIndex("��׼��ֵ")) = flexAlignCenterCenter
'        .ColAlignment(.ColIndex("���ֱ�׼")) = flexAlignCenterCenter
'        '���ص�Ԫ��
'        .ColWidth(.ColIndex("ID")) = 0
'        .ColWidth(.ColIndex("�ϼ�ID")) = 0
'        .ColWidth(.ColIndex("����ID")) = 0
'        .ColWidth(.ColIndex("���")) = 0
'        '�������
'        .ColWidth(.ColIndex("��Ŀ")) = 1500
'        .ColWidth(.ColIndex("��׼��ֵ")) = 850
'        .ColWidth(.ColIndex("ȱ������")) = 3700
'        .ColWidth(.ColIndex("���ֱ�׼")) = 1100
'        '�и�����
''        .RowHeightMin = 300
'        '���������
''        .ColWidthMax = 7000
'        '�Զ���Ӧ�иߡ��п�
'        .AutoSizeMode = flexAutoSizeRowHeight
'        .AutoSize .ColIndex("ȱ������")
'        .SelectionMode = flexSelectionListBox
'        .AllowBigSelection = False
'        .Redraw = flexRDBuffered
'        'ѡ����ǰ����
'        If m_lngOldRow > 0 And m_lngOldRow < i Then
'            .Row = m_lngOldRow
'            .Col = 2
'            .ShowCell m_lngOldRow, 2
'            On Error Resume Next
'            If .Visible = True Then .SetFocus
'            fgMain_SelChange
'        ElseIf fgMain.Tag = "" And i > 1 And .Rows > 1 Then
'            m_lngOldRow = 1
'            fgMain.Tag = "ѡ�е�һ��"
'            .Row = 1
'            .Col = 2
'            .ShowCell m_lngOldRow, 2
'            On Error Resume Next
'            If .Visible = True Then .SetFocus
'            fgMain_SelChange
'        Else
'            lblInfo = "������"
'        End If
'
'    End With
'
'    Call SetMenu
'    Call ˢ�·�����Ϣ
'End Sub
'
'Private Sub SetMenu()
'    '����:�����޸ĺ�ɾ����ť����Чֵ
'    '���û��ѡ����������Ӧ��ť
'
'    Dim blnModBZ As Boolean, blnModFA As Boolean
'    If fgMain.Rows <= 1 Then    '������
'        fgMain.WallPaper = imgBG_fg(0).Picture
'    Else
'        fgMain.WallPaper = LoadPicture("")
'    End If
'    If IsNumeric(Mid(m_strTreeKey, 2)) Then '����Ϊ��
'        blnModFA = True
'    Else
'        blnModFA = False
'    End If
'    If m_lngCurRow < 1 Or fgMain.Rows <= 1 Then  '��׼Ϊ��
'        blnModBZ = False
'    Else
'        blnModBZ = True
'    End If
'
'    Dim rsTemp As New ADODB.Recordset
'    gstrSQL = "select count(*) from �������ֽ�� where ����ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
'    If rsTemp(0).Value > 0 Then
'        '�÷����Ѿ�ʹ��
'        fgMain.WallPaper = imgBG_fg(1).Picture
'        blnModBZ = False
'        blnModFA = False
'    End If
'    rsTemp.Close
'
''    mnuShortMenuFA(5).Enabled = IIf(fgMain.Rows > 1, False, blnModFA)    ' ����Ϊ�հ׷���ʱ���ܵ��룡
''    mnuEditEmportFA.Enabled = IIf(fgMain.Rows > 1, False, blnModFA)
''
''    Toolbar1.Buttons("FA").ButtonMenus(2).Enabled = blnModFA
''    Toolbar1.Buttons("FA").ButtonMenus(3).Enabled = blnModFA
''    mnuEditDelFA.Enabled = blnModFA
''    mnuEditModFA.Enabled = blnModFA
''    mnuShortMenuFA(2).Enabled = blnModFA
''    mnuShortMenuFA(3).Enabled = blnModFA
''    Toolbar1.Buttons("NewBZ").Enabled = blnModFA
''    Toolbar1.Buttons("XM").ButtonMenus("NewXM").Enabled = blnModFA
''    mnuEditNewXM.Enabled = blnModFA
''    mnuEditInsXM.Enabled = blnModFA
''    mnuShortMenuXM(1).Enabled = blnModFA
''    mnuShortMenuXM(2).Enabled = blnModFA
''    mnuShortMnuXM(1).Enabled = blnModFA
''    mnuShortMnuXM(2).Enabled = blnModFA
''    mnuEditNewBZ.Enabled = blnModFA
''    mnuEditInsBZ.Enabled = blnModFA
''    mnuShortMenuBZ(1).Enabled = blnModFA
''    mnuShortMenuBZ(2).Enabled = blnModFA
''    If fgMain.Rows > 1 Then
''        mnuEditInsXM.Enabled = blnModFA
''        mnuEditInsBZ.Enabled = blnModFA
''        mnuShortMenuXM(2).Enabled = blnModFA
''        mnuShortMnuXM(2).Enabled = blnModFA
''        mnuShortMenuBZ(4).Enabled = blnModFA
''        mnuShortMenuBZ(3).Enabled = blnModFA
''        mnuShortMenuXM(3).Enabled = blnModFA
''        mnuShortMnuXM(3).Enabled = blnModFA
''        mnuShortMnuXM(2).Enabled = blnModFA
''    Else
''        mnuEditInsXM.Enabled = False
''        mnuEditInsBZ.Enabled = False
''        mnuShortMenuXM(2).Enabled = False
''        mnuShortMnuXM(2).Enabled = False
''        mnuEditNewBZ.Enabled = False
''        mnuEditInsBZ.Enabled = False
''        Toolbar1.Buttons("NewBZ").Enabled = False
''        mnuShortMenuBZ(1).Enabled = False
''        mnuShortMenuBZ(2).Enabled = False
''        mnuShortMenuBZ(4).Enabled = False
''        mnuShortMenuBZ(3).Enabled = False
''        mnuShortMenuXM(3).Enabled = False
''        mnuShortMnuXM(3).Enabled = False
''        mnuShortMnuXM(2).Enabled = False
''    End If
''
''    Toolbar1.Buttons("ModBZ").Enabled = blnModBZ
''    Toolbar1.Buttons("DelBZ").Enabled = blnModBZ
''    Toolbar1.Buttons("XM").ButtonMenus("ModXM").Enabled = blnModBZ
''    Toolbar1.Buttons("XM").ButtonMenus("DelXM").Enabled = blnModBZ
''
''    mnuEditDelBZ.Enabled = blnModBZ
''    mnuEditModBZ.Enabled = blnModBZ
''    mnuEditDelXM.Enabled = blnModBZ
''    mnuEditModXM.Enabled = blnModBZ
''
''    mnuShortMenuBZ(4).Enabled = blnModBZ
''    mnuShortMenuXM(4).Enabled = blnModBZ
''    mnuShortMnuXM(4).Enabled = blnModBZ
''
''    If m_lngCurSJID <= 0 And fgMain.Rows > 1 Then  '���ϼ�����ʾ����������
''        mnuEditModBZ.Enabled = False
''        mnuEditDelBZ.Enabled = False
''        mnuShortMenuBZ(3).Enabled = False
''        mnuShortMenuBZ(4).Enabled = False
''        Toolbar1.Buttons("SplitBZ").Enabled = False
''        Toolbar1.Buttons("ModBZ").Enabled = False
''        Toolbar1.Buttons("DelBZ").Enabled = False
''    Else
''        mnuEditModBZ.Enabled = blnModBZ
''        mnuEditDelBZ.Enabled = blnModBZ
''        mnuShortMenuBZ(3).Enabled = blnModBZ
''        mnuShortMenuBZ(4).Enabled = blnModBZ
''        Toolbar1.Buttons("SplitBZ").Enabled = blnModBZ
''        Toolbar1.Buttons("ModBZ").Enabled = blnModBZ
''        Toolbar1.Buttons("DelBZ").Enabled = blnModBZ
''    End If
''
''    If mnuEditNewXM.Enabled = False And mnuEditInsXM.Enabled = False And mnuEditModXM.Enabled = False And mnuEditDelXM.Enabled = False Then
''        Toolbar1.Buttons("XM").Enabled = False
''    Else
''        Toolbar1.Buttons("XM").Enabled = True
''    End If
''    If fgMain.Rows <= 1 Then
''        mnuShortMenuFA(4).Enabled = False     'ֻҪ��׼���ھ�����ѡ�ã�
''        mnuEditSelFA.Enabled = False
''        Toolbar1.Buttons("FA").ButtonMenus(4).Enabled = False
''    Else
''        If tvw����.Nodes(tvw����.SelectedItem.Index).Image = "RootSel" Then
''            mnuShortMenuFA(4).Enabled = False
''            mnuEditSelFA.Enabled = False
''            Toolbar1.Buttons("FA").ButtonMenus(4).Enabled = False
''        Else
''            mnuShortMenuFA(4).Enabled = True
''            mnuEditSelFA.Enabled = True
''            Toolbar1.Buttons("FA").ButtonMenus(4).Enabled = True
''        End If
''    End If
''
''    '����б�ֵ����1���������ӡ
''    EnablePrint fgMain.Rows > 1
'
'    '��ʾ��¼����Ϣ
'    stbThis.Panels(2).Text = "�б��й���ʾ��" & fgMain.Rows - 1 & "�����ݡ�"
'End Sub
'
'Private Sub Ȩ�޿���()
'    '����:�����е��û�Ȩ�޲���,��ʹһЩ�˵����ť���ɼ�
''    If InStr(gstrPrivs, "��ɾ��") = 0 Then
''        mnuEdit.Visible = False
''        'mnusplit1.Visible = False
''        'mnuFileSetup.Visible = False
''        mnuShortMenuBZ(1).Visible = False
''        mnuShortMenuBZ(2).Visible = False
''        mnuShortMenuBZ(3).Visible = False
''        mnuShortMenuBZ(4).Visible = False
''        mnuShortMenuSplit.Visible = False
''        mnuShortMenuXM(1).Visible = False
''        mnuShortMenuXM(2).Visible = False
''        mnuShortMenuXM(3).Visible = False
''        mnuShortMnuXM(1).Visible = False
''        mnuShortMnuXM(2).Visible = False
''        mnuShortMnuXM(3).Visible = False
''        mnuShortMenuFA(1).Visible = False
''        mnuShortMenuFA(2).Visible = False
''        mnuShortMenuFA(3).Visible = False
''        Toolbar1.Buttons("SplitFA").Visible = False
''        Toolbar1.Buttons("FA").Visible = False
''        Toolbar1.Buttons("FA").ButtonMenus(1) = False
''        Toolbar1.Buttons("FA").ButtonMenus(2) = False
''        Toolbar1.Buttons("FA").ButtonMenus(3) = False
''        Toolbar1.Buttons("FA").ButtonMenus(4) = False
''        Toolbar1.Buttons("SplitXM").Visible = False
''        Toolbar1.Buttons("XM").Visible = False
''        Toolbar1.Buttons("XM").ButtonMenus(1) = False
''        Toolbar1.Buttons("XM").ButtonMenus(2) = False
''        Toolbar1.Buttons("XM").ButtonMenus(3) = False
''        Toolbar1.Buttons("SplitBZ").Visible = False
''        Toolbar1.Buttons("NewBZ").Visible = False
''        Toolbar1.Buttons("ModBZ").Visible = False
''        Toolbar1.Buttons("DelBZ").Visible = False
''    End If
'End Sub
'
''Private Sub EnablePrint(ByVal blnEnabled As Boolean)
''    '����:���ô�ӡ��Ԥ����ť����Чֵ
''    '����:blnEnabled ��Чֵ
''
''    Toolbar1.Buttons("Print").Enabled = blnEnabled
''    Toolbar1.Buttons("Preview").Enabled = blnEnabled
''    mnuFilePre.Enabled = blnEnabled
''    mnuFilePrint.Enabled = blnEnabled
''    mnuFileExcel.Enabled = blnEnabled
''End Sub
'
'Private Sub Show����Ҫ��(lngID As Long, ��Ŀ As String, ��׼��ֵ As String)
'    '������ĿID��ʾ����Ҫ��
'    Dim Rs As New ADODB.Recordset
'    gstrSQL = "select ID,���� as ����Ҫ��,�ϼ�ID from �������ֱ�׼ Where ID=" & CStr(lngID)
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'
'    If Not Rs.EOF Then
'        If m_lngOldSJID > 0 And m_lngOldSJID = lngID Then Exit Sub
'        If IsNull(Rs.Fields("����Ҫ��")) Then
'                lblInfo = "���ƣ�" + ��Ŀ + "  " + IIf(Len(Trim(��׼��ֵ)) = 0, "", "(" + ��׼��ֵ + ")")
'                lblInfo = lblInfo + vbCrLf
'        Else
'            If Len(Rs.Fields("����Ҫ��")) > 0 Then
'                lblInfo = "���ƣ�" + ��Ŀ + "  " + IIf(Len(Trim(��׼��ֵ)) = 0, "", "(" + ��׼��ֵ + ")")
'                lblInfo = lblInfo + vbCrLf + Rs.Fields("����Ҫ��")
'            End If
'        End If
'    Else
'        lblInfo.Caption = "������":
'    End If
'    m_lngOldSJID = m_lngCurSJID
'End Sub
'
'
