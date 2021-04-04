VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFeeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'ֱ��ͨ����ģ�鼤����¼�
Public Event Activate() '�Ӵ��弤��
Public Event RequestRefresh() 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal strText As String) 'Ҫ�����������״̬������
Private WithEvents mfrmFeeQuery As frmFeeQuery
Attribute mfrmFeeQuery.VB_VarHelpID = -1
Private Sub Class_Initialize()
    Set mfrmFeeQuery = New frmFeeQuery
    glngInstanceCount = glngInstanceCount + 1
End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmFeeQuery Is Nothing Then Unload mfrmFeeQuery
    Set mfrmFeeQuery = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub

Private Sub mfrmFeeQuery_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmFeeQuery_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub mfrmFeeQuery_StatusTextUpdate(ByVal strText As String)
    RaiseEvent StatusTextUpdate(strText)
End Sub
Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:���˺�
    '����:2012-06-18 16:50:35
    '����:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmFeeQuery Is Nothing Then Exit Sub
    mfrmFeeQuery.SetFontSize (bytSize)
End Sub
Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
    
    '��ʼ��ȡȨ��,�Լӿ�����ٶ�(���Ǳ�Ҫ��)
    Call GetInsidePrivs(Enum_Inside_Program.p���ò�ѯ, True)
    
    Set zlGetForm = mfrmFeeQuery
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByRef CommandBars As Object, ByVal int���� As Integer, Optional ByVal blnChildToolBar As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
    '���:int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    '       CommandBars=�����ڲ鿴ʱ���Բ���(����Nothing)
    '       blnChildToolBar = True��ʾ�������������Լ��Ĵ����ڲ�
    '����:
    '����:
    '˵��:
    '   �����Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)�����bln�ڲ�������Ϊ�٣������������ϴ������������˵���ȻҪ��������
    '   ����Ҫ���Լ��Ľ����ϴ�������������˶����Լ��������Ѿ����ڹ������ĳ���Ӧ����ؼ����ظ���
    'ע��:
    '         ���ӹ�����ʱע��������ܰ�ť��������Ҫ�ظ�
    '         ����������ģ���޲˵���conMenu_ManagePopup������ӳ����ڴ���ʱ��Ҫ��飬�޴˶���ʱ���ӵ����ѵĲ˵���
    '         ����������ڲ�����������ɾ�����������������
    '         δʹ�ù�������ģ����Ҫ���ӳ�ʼ������������
    '         �������Ĺ���״̬�ı仯��ͨ�����������zlUpdateCommandBars��ͳһ����
    '����:���˺�
    '����:2010-10-29 15:14:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmFeeQuery.zlDefCommandBars(frmParent, CommandBars, int����, blnChildToolBar)
End Sub

Public Sub zlExecuteCommandBarsDirect(ByRef Control As Object, ByRef frmParent As Object, _
    ByRef strPrivs As String, ByVal bln���� As Boolean, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal bytӤ�� As Byte, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal lngǰ��ID As Long, _
    ByVal lng�������ID As Long, ByVal int���� As Integer, ByVal bln���� As Boolean, _
    ParamArray arrPar() As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ����Ӧģ��ľ��幦�ܣ���ִ��ҽ��ģ����¿�ҽ������
    '���:int����-������(int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS))
    '       bln����-��ת�ƣ�ת�����Ĳ��˽��в���
    '       lng����ID-������Ϊtrueʱ,���뱾����Ҫ���ѵĿ���ID
    '       lng����ID-������Ϊtrueʱ,���뱾����Ҫ���ѵĲ���ID
    '����:
    '����:
    '����:���˺�
    '����:2010-12-12 09:57:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objControl As CommandBarControl, bln��Ժ As Boolean, bln���� As Boolean, rsTemp As ADODB.Recordset
    Select Case Control.ID
        Case conMenu_File_PrintBedCard: Call zlPrintBedCard(frmParent, lng����ID, lng��ҳID)    '��ӡ��ͷ��
        Case conMenu_File_PrintSingleBill: Call zlExecPrintSingleBill(frmParent, lng����ID, strPrivs, "")   '��ӡ���Ŵ߿
        Case conMenu_File_PrintDayDetail: Call zlPrintDayDetail(frmParent, 1, lng����ID, lng����ID, False, False) '��ӡһ���嵥
        Case conMenu_File_PrintPageSet: Call zlPrintAccountPage(frmParent)   '��ӡ��ҳ����
        Case conMenu_Edit_PreBalance    'Ԥ����
            If bln���� Then
                '����Ԥ��
                Call zlPreBalanceAll(frmParent, lng����ID)
            Else
                Call zlPreBalance(frmParent, lng����ID, lng��ҳID)
            End If
        Case conMenu_Edit_Billing   '����
        
            On Error GoTo errHandle

            gstrSQL = "Select A.��Ժ����,B.�������  From  ������ҳ A,������� B Where A.����ID=[1] and A.��ҳID=[2]  and a.����id=b.����ID(+) And b.����(+)=1 And b.����(+)=1 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ״̬", lng����ID, lng��ҳID)
            If rsTemp.EOF Then Exit Sub
            bln��Ժ = IIf(Not IsNull(rsTemp!��Ժ����), True, False)
            bln���� = IIf(Val(Nvl(rsTemp!�������)) = 0, True, False)
            '
            Call zlExecBilling(int����, frmParent, lng����ID, lng����ID, bln��Ժ, bln����, "", lng��ҳID, bln����, lng����ID)
        Case conMenu_Edit_Balance
            '����
            Call ExecBalance(frmParent, lng����ID, CStr(lng��ҳID))
        Case conMenu_Edit_Billing_Mulit    '��������
            On Error GoTo errHandle
            gstrSQL = "Select A.��Ժ����,B.�������  From  ������ҳ A,������� B Where A.����ID=[1] and A.��ҳID=[2]  and a.����id=b.����ID(+) And b.����(+)=1 And b.����(+)=1 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ״̬", lng����ID, lng��ҳID)
            If rsTemp.EOF Then Exit Sub
            bln��Ժ = IIf(Not IsNull(rsTemp!��Ժ����), True, False)
            bln���� = IIf(Val(Nvl(rsTemp!�������)) = 0, True, False)
            '
            Call zlExecBilling_Mulit(int����, frmParent, lng����ID, lng����ID, bln��Ժ, bln����, "", lng��ҳID, bln����, lng����ID)
        Case conMenu_Edit_ReBilling '����
        Case conMenu_Edit_CardBackMoney '����˿�
            Call NurseDeposit(frmParent, lng����ID, lng��ҳID)
        Case conMenu_Edit_ReBillingApply, conMenu_Edit_ReBillingAudit   '������������
            Call zlWrite_Off_ApplyAndVerfy(frmParent, lng����ID, lng����ID, Control.ID = conMenu_Edit_ReBillingApply)
        Case conMenu_View_TurnToWardFeeQuery 'ת�������ñ䶯��ѯ
            If CreatePublicExpenseBillOperation() Then
                Call gobjPublicExpenseBillOperation.zlTurnToWard_Fee_Query(frmParent, 3, lng����ID, lng��ҳID)
            End If
        Case conMenu_View_DateType * 10 + 1, conMenu_View_DateType * 10 + 2 'ʱ��ģʽ
        Case conMenu_View_DetailType * 10 To conMenu_View_DetailType * 10 + 9 '��ѯ��ʽ'
        Case conMenu_View_ReBalance '��ʾ��������
        Case conMenu_View_ZeroFee   '��ʾ�����
        Case conMenu_View_CheckFee  '��ʾ������
        Case conMenu_View_ToolBar_Button '������
        Case conMenu_View_ToolBar_Text '��ť����
        Case conMenu_View_ToolBar_Size '��ͼ��
        Case conMenu_View_PatInfor  '�鿴���˿�Ƭ
        Case conMenu_View_Billing   '�鿴���ʵ�
        Case conMenu_View_Refresh
        Case conMenu_Tool_Option    '����ѡ��
        Case Else
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecBalance(frmMain As Object, lng����ID As Long, str��ҳID As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�н��ʲ���
    '����:���˺�
    '����:2015-02-05 12:00:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String
    
    strPrivs = ";" & GetInsidePrivs(Enum_Inside_Program.p���˽���) & ";"
    If Val(zlDatabase.GetPara("���ʽ�����", glngSys, 1137, "1")) = 0 Then
        If frmPatiBalanceTraditional.ShowMe(frmMain, g_Ed_סԺ����, strPrivs, lng����ID, str��ҳID) = False Then Exit Sub
    Else
        If frmPatiBalanceSplit.ShowMe(frmMain, g_Ed_סԺ����, strPrivs, lng����ID, str��ҳID) = False Then Exit Sub
    End If
    
End Sub

Public Sub zlCheckPrivs(ByVal Control As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ɳ������Ƿ�߱���Ȩ��
    '���:Control-���ܰ�ť��IDֵ
    '����:
    '����:����Ȩ��,����True,���򷵻�False
    '����:���˺�
    '����:2010-10-29 15:12:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static strBillingPrivs As String
    Static strQueryPrivs As String
    Static strBalancePrivs As String
    Dim blnVisible As Boolean
    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If strBillingPrivs = "" Then
        strBillingPrivs = ";" & GetInsidePrivs(Enum_Inside_Program.p���ʲ���) & ";"
        If strBillingPrivs = "" Then strBillingPrivs = "-"  '�����ظ�����
    End If
    If strQueryPrivs = "" Then
        strQueryPrivs = GetInsidePrivs(Enum_Inside_Program.p���ò�ѯ)
        If strQueryPrivs = "" Then strQueryPrivs = "-"  '�����ظ�����
    End If
    If strBalancePrivs = "" Then
        strBalancePrivs = GetInsidePrivs(Enum_Inside_Program.p���˽���)
        If strBalancePrivs = "" Then strBalancePrivs = "-"  '�����ظ�����
    End If
    blnVisible = True
    Select Case Control.ID
        Case conMenu_File_PrintPageSet
            blnVisible = InStr(strQueryPrivs, ";������ҳ") > 0
        Case conMenu_File_PrintMultiBill, conMenu_File_PrintSingleBill
            blnVisible = InStr(strQueryPrivs, ";�߿��ӡ") > 0
        Case conMenu_Edit_PreBalanceAll
            blnVisible = InStr(strQueryPrivs, ";Ԥ�����в���") > 0
        Case conMenu_Edit_Billing, conMenu_Edit_Billing_Mulit
            blnVisible = InStr(strBillingPrivs, ";סԺ����;") > 0
        Case conMenu_Edit_ReBilling
            '55380
            blnVisible = InStr(strBillingPrivs, ";ҩƷ����;") > 0 _
                Or InStr(strBillingPrivs, ";��������;") > 0 _
                Or InStr(strBillingPrivs, ";��������;") > 0
        Case conMenu_Edit_Balance
            blnVisible = InStr(strBalancePrivs, ";�����ֽ����;") = 0
        Case conMenu_Edit_ReBillingApply
            blnVisible = (InStr(strBillingPrivs, ";ҩƷ��������;") > 0 _
                Or InStr(strBillingPrivs, ";������������;") > 0 _
                Or InStr(strBillingPrivs, ";������������;") > 0) And InStr(strBillingPrivs, "��������") > 0
        Case conMenu_Edit_ReBillingAudit
            blnVisible = InStr(strBillingPrivs, "�������") > 0
        Case conMenu_Edit_ReBillingButton
            '55380
            blnVisible = InStr(strBillingPrivs, "�������") > 0 _
                Or ((InStr(strBillingPrivs, ";ҩƷ��������;") > 0 _
                    Or InStr(strBillingPrivs, ";������������;") > 0 _
                    Or InStr(strBillingPrivs, ";������������;") > 0) And InStr(strBillingPrivs, "��������") > 0)
        Case conMenu_Edit_CardBackMoney
            blnVisible = InStr(";" & GetInsidePrivs(9000), ";��Ժ��������˿�;") > 0 Or InStr(";" & GetInsidePrivs(9000), ";��Ժ��������˿�;") > 0
    End Select
    Control.Visible = blnVisible
End Sub
Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmFeeQuery.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    
    Call mfrmFeeQuery.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmFeeQuery.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strסԺ�� As String, ByVal lng����ID As Long, _
    ByVal intInsure As Integer, ByVal blnDateMoved As Boolean, ByVal bln��Ժ As Boolean, _
    ByVal bln���� As Boolean, Optional blnOnlyRefreshVar As Boolean, _
    Optional bln���� As Boolean = False, Optional lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ�²��˷�������
    '���:
    '       bln����-��ת�ƣ�ת�����Ĳ��˽��в���
    '       lng����ID-������Ϊtrueʱ,���뱾����Ҫ���ѵĿ���ID
    '       lng����ID-������Ϊtrueʱ,���뱾����Ҫ���ѵĲ���ID
    '����:
    '����:
    '����:���˺�
    '����:2010-12-10 14:24:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmFeeQuery.zlRefresh(lng����ID, lng��ҳID, strסԺ��, lng����ID, intInsure, blnDateMoved, bln��Ժ, bln����, blnOnlyRefreshVar, bln����, lng����ID)
End Sub

Public Sub InitCallByNurse(frmMain As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, lngSys As Long)
'���ܣ��ɻ�ʿ����վ����ʱ��ʼ������
    
    Set gfrmMain = frmMain
    'If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrDBUser = strDBUser
    glngSys = lngSys
    gblnNurseStation = True
    glngModul = 1139
    
    gstrPrivs = ";" & GetPrivFunc(glngSys, glngModul) & ";"
    gstrUnitName = GetUnitName
    
    Call GetUserInfo
    Call InitSysPar
    Call InitLocPar(glngModul)
    Call InitLocPar(Enum_Inside_Program.pסԺ����)
    Call CreateSquareCardObject(frmMain, glngModul)
    Call CreatePublicDrug(lngSys, cnMain, strDBUser)
    Call CreatePublicExpenseObject(glngModul)
    'ҽ���ӿ�
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreatePlugIn(glngModul)
End Sub
Public Function zlPatiPressMoney(ByVal frmMain As Object, _
    ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, _
    ByVal lngMoudle As Long, strDBUser As String, _
    ByVal strPrivs As String, ByVal lng����ID As Long, ByVal str�������� As String, _
    Optional lng����ID As Long = 0, Optional bytPrintModule As Byte = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���벡�˴߿��������
    '���:frmMain-���õĴ���
    '       strDBUser-db�û�
    '       bytPrintModule-2.��ӡ;1-Ԥ��
    '����:
    '����:�����ӡ�ɹ�1�����ϵĲ���,����true,���򷵻�False
    '����:���˺�
    '����:2010-12-16 10:28:25
    '����:36025
    '---------------------------------------------------------------------------------------------------------------------------------------------
    glngModul = lngMoudle: glngSys = lngSys
    Set gcnOracle = cnOracle
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrDBUser = strDBUser
    glngSys = lngSys
        
    Call GetUserInfo
    Call InitSysPar
    Call InitLocPar(glngModul)
    Call CreateSquareCardObject(frmMain, glngModul)

    zlPatiPressMoney = frmPatiPressMoney.zlPatiPressMoney(frmMain, lngMoudle, strPrivs, lng����ID, str��������, lng����ID, bytPrintModule)
End Function
Public Function zlPatiBalance(ByVal frmMain As Object, _
    ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, strDBUser As String, _
    ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strDelTime As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˽���
    '���:frmMain-���õĴ���
    '       strDBUser-db�û�
    '       strDelTime-�������ϵ�ʱ���ת��ʱ��(��Ҫ�Ǵ�����Ԥ�����)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-03-29 13:51:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    glngModul = 1137: glngSys = lngSys: gstrDBUser = strDBUser
    Set gcnOracle = cnOracle
    gstrPrivs = ";" & GetPrivFunc(lngSys, glngModul) & ";"
    If (InStr(gstrPrivs, ";������ý���;") = 0) Or _
        (InStr(gstrPrivs, ";��ͨ���˽���;") = 0 And InStr(gstrPrivs, ";���ս���;") = 0) Then
        Exit Function    'û��Ȩ��,����ִ��
    End If
    If lng��ҳID = 0 And InStr(gstrPrivs, ";������ý���;") = 0 Then
        MsgBox "�㲻�߱�������ý���Ȩ��,����ϵͳ����Ա��ϵ!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    Call GetUserInfo: Call InitSysPar
    Call InitLocPar(glngModul)
    'ҽ���ӿ�
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreatePlugIn(glngModul)
    Call CreateSquareCardObject(frmMain, glngModul)
    
    'IIf(lng��ҳID = 0, 0, 1) '0-����;1-סԺ
    zlPatiBalance = frmPatiBalanceSplit.ShowMe(frmMain, IIf(lng��ҳID = 0, g_Ed_�������, g_Ed_סԺ����), _
        gstrPrivs, lng����ID, CStr(lng��ҳID), , , True)
End Function

Public Function zlPatiBilling(ByVal frmMain As Object, _
    ByVal lng����ID As Long, ByVal lng����ID As Long, lng��ҳID As Long, _
    ByVal lng����ID As Long, ByVal bln���� As Boolean, ByVal lngҽ��ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˼���(����ҽ��),Ŀǰ��ʿվ����
    '���:frmMain-���õĴ���
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:������
    '����:2015-12-17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln��Ժ As Boolean, bln���� As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim bln�������۲��� As Boolean
    
    On Error GoTo errHandle
    gstrSQL = _
        " Select A.��Ժ����,B.�������,a.��������" & _
        " From  ������ҳ A,������� B" & _
        " Where A.����ID=[1] and A.��ҳID=[2]  and a.����id=b.����ID(+) And b.����(+)=1 And b.����(+)=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ״̬", lng����ID, lng��ҳID)
    If rsTemp.EOF Then Exit Function
    bln��Ժ = Not IsNull(rsTemp!��Ժ����)
    bln���� = Val(Nvl(rsTemp!�������)) = 0
    bln�������۲��� = Val(Nvl(rsTemp!��������)) = 1
    
    zlPatiBilling = zlExecBilling(1, frmMain, lng����ID, lng����ID, bln��Ժ, bln����, "", _
        lng��ҳID, bln����, lng����ID, lngҽ��ID, bln�������۲���)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetFormOperation() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ѡ�񣬴���ж��ǰ����
    '����:
    '     �ϴδ�����������ַ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetFormOperation = mfrmFeeQuery.zlGetFormOperation
End Function

Public Function RestoreFormOperation(ByVal strValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ָ��������ѡ�񣬴���ˢ��ǰ����
    '���:
    '     strValue-�ϴδ�����������ַ���
    '����:
    '     True-����ָ��ɹ�;False-����ָ�ʧ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If strValue = "" Then Exit Function
    If mfrmFeeQuery.zlRestoreFormOperation(strValue) = False Then Exit Function
End Function
