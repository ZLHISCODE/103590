VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ��¼���ݿ��û���

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If

    Set gfrmMain = frmMain
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    glngSys = lngSys
    gstrDBOwner = Sys.SystemOwner(lngSys)  ' ȡϵͳ������
    gstrUnitName = GetUnitName
    gstrMatch = IIf(Val(zlDatabase.GetPara("����ƥ��", , , True)) = 0, "%", "")

    'ȡҩƷ�۸�ȼ�
    Call GetPriceClass
    
    '-------------------------------------------------
    On Error Resume Next
    Select Case lngModul
        Case 1023 'ҩƷĿ¼����
            frmMediLists.ShowMe gfrmMain, False
        Case 1050 '���Ʋο�����
            frmCureRefers.Show , gfrmMain
        Case 1051 '������ϲο�
            frmDiagnoses.Show , gfrmMain
        Case 1053 '����������Ŀ
            frmVItemLists.Show , gfrmMain
        Case 1054 '������Ŀ����
            frmClinicLists.Show , gfrmMain
        Case 1055 'ҽ��Ƶ������
            frmAdviceFreq.Show , gfrmMain
        Case 1056 'Ӱ������Ŀ
            frmRadLists.Show , gfrmMain
        Case 1057 'Ӱ���豸Ŀ¼
            frmPacsDev.Show , gfrmMain
        Case 1049 '��鲿λ����
            frmCheckPartList.Show , gfrmMain
        Case 1058   '������������
            frmApparatusList.Show , gfrmMain
'            frmInstrument.Show , gfrmMain
        Case 1059   '������Ŀ����
            frmLabItems.Show , gfrmMain
'            frmVerify.Show , gfrmMain
        Case 1060 '�����ط������
            frmAntiList.Show , gfrmMain
'            FrmAntibioticManager.Show , gfrmMain
        Case 1061 '����ϸ������
            FrmMicrobeList.Show , gfrmMain
'            FrmBacilliManager.Show , gfrmMain
        Case 1062 '�ʿ�Ʒ����
            frmMassResList.Show , gfrmMain
'            FrmQualityResManager.Show , gfrmMain
        Case 1063 '�ʿع������
            frmMassRuleList.Show , gfrmMain
'            FrmQualityRuleManager.Show , gfrmMain
        Case 1065 '���鱨��ģ��
            frmLabRptList.Show , gfrmMain
        Case 1066   '�����ʿع���
            frmAppRuleList.Show , gfrmMain
        Case 1067 '��Ѫ�ܹ��� 070425 add
            frmMiningVessels.Show , gfrmMain
        Case 1068 '��˹��� 070928 add
            frmLabVerifyList.Show , gfrmMain
        Case 1069 '����С�� 090104
            frmLabGroup.Show , gfrmMain
        Case 1077 '������Ϣ����
            frmBaseInfoList.Show , gfrmMain
        Case 1024 '����ҩ����Ȩ
            frmKSSManagement.Show , gfrmMain
        Case 1276 '��׼·���ο�
            frmStPathMain.Show , gfrmMain
        Case 1080 '������Ȩ����
            frmOPSEmpower.Show , gfrmMain
        Case 1081 '�ⷿ��λ
            frmStoreSpace.Show , gfrmMain
        Case 1082 'ҽ����Ȩ����
            frmDoctorManage.Show , gfrmMain
        Case 1085 'ҽ�����Ӱ���������
            frmDocShiftBase.Show , gfrmMain
    End Select
End Sub

Public Sub SetMedList(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, ByVal strDBUser As String)
'���ܣ����������ٴ��Թ�ҩ

    Set gfrmMain = frmMain
'    If Not CheckValid Then Exit Sub    '���ͨѶԭ��
    
    Set gcnOracle = cnMain
    glngSys = lngSys
    glngModul = 1023
    gstrDBUser = strDBUser
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrPrivs = GetPrivFunc(lngSys, 1023)
    
    If gfrmMain Is Nothing Then
        frmMediLists.Show vbModal
    Else
        frmMediLists.ShowMe frmMain, True
    End If
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ��¼���ݿ��û���

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If

    Set gfrmMain = Nothing
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    glngSys = lngSys
    gstrDBOwner = Sys.SystemOwner(lngSys) ' ȡϵͳ������
    gstrUnitName = GetUnitName
    gstrMatch = IIf(Val(zlDatabase.GetPara("����ƥ��", , , True)) = 0, "%", "")
    
    'ȡҩƷ�۸�ȼ�
    Call GetPriceClass
    
    '-------------------------------------------------
    On Error Resume Next
    Select Case lngModul
        Case 1023 'ҩƷĿ¼����
            OS.ShowChildWindow frmMediLists.hwnd, lngMain
        Case 1050 '���Ʋο�����
            OS.ShowChildWindow frmCureRefers.hwnd, lngMain
        Case 1051 '������ϲο�
            OS.ShowChildWindow frmDiagnoses.hwnd, lngMain
        Case 1053 '����������Ŀ
            OS.ShowChildWindow frmVItemLists.hwnd, lngMain
        Case 1054 '������Ŀ����
            OS.ShowChildWindow frmClinicLists.hwnd, lngMain
        Case 1055 'ҽ��Ƶ������
            OS.ShowChildWindow frmAdviceFreq.hwnd, lngMain
        Case 1056 'Ӱ������Ŀ
            OS.ShowChildWindow frmRadLists.hwnd, lngMain
        Case 1057 'Ӱ���豸Ŀ¼
            OS.ShowChildWindow frmPacsDev.hwnd, lngMain
        Case 1049 '��鲿λ����
            OS.ShowChildWindow frmCheckPartList.hwnd, lngMain
        Case 1058   '������������
            OS.ShowChildWindow frmApparatusList.hwnd, lngMain
        Case 1059   '������Ŀ����
            OS.ShowChildWindow frmLabItems.hwnd, lngMain
        Case 1060 '�����ط������
            OS.ShowChildWindow frmAntiList.hwnd, lngMain
        Case 1061 '����ϸ������
            OS.ShowChildWindow FrmMicrobeList.hwnd, lngMain
        Case 1062 '�ʿ�Ʒ����
            OS.ShowChildWindow frmMassResList.hwnd, lngMain
        Case 1063 '�ʿع������
            OS.ShowChildWindow frmMassRuleList.hwnd, lngMain
        Case 1065 '���鱨��ģ��
            OS.ShowChildWindow frmLabRptList.hwnd, lngMain
        Case 1066   '�����ʿع���
            OS.ShowChildWindow frmAppRuleList.hwnd, lngMain
        Case 1067 '��Ѫ�ܹ��� 070425 add
            OS.ShowChildWindow frmMiningVessels.hwnd, lngMain
        Case 1068 '��˹��� 070928 add
            OS.ShowChildWindow frmLabVerifyList.hwnd, lngMain
        Case 1069 '����С�� 090104
            OS.ShowChildWindow frmLabGroup.hwnd, lngMain
        Case 1077 '������Ϣ����
            OS.ShowChildWindow frmBaseInfoList.hwnd, lngMain
        Case 1024 '����ҩ����Ȩ
            OS.ShowChildWindow frmKSSManagement.hwnd, lngMain
        Case 1276 '��׼·���ο�
            OS.ShowChildWindow frmStPathMain.hwnd, lngMain
        Case 1080 '������Ȩ����
            OS.ShowChildWindow frmOPSEmpower.hwnd, lngMain
        Case 1081 '�ⷿ��λ
            frmStoreSpace.Show , gfrmMain
        Case 1082 'ҽ����Ȩ����
            OS.ShowChildWindow frmDoctorManage.hwnd, lngMain
        Case 1085 'ҽ�����Ӱ���������
            frmDocShiftBase.Show , gfrmMain
    End Select
End Sub

Public Sub CallClinicScheme(frmMain As Object, cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strDBUser As String, ByVal int��Χ As Integer)
'���ܣ�����������Ŀ����-���׷�������
'������int��Χ=���׷����Ŀ�ʹ�ó��ϣ�1-����,2-סԺ,3-�����סԺ

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = 1054 '1054-������Ŀ����
    
    'Ȩ���жϺʹ���
    gstrPrivs = ";" & GetPrivFunc(lngSys, 1054) & ";"
    If InStr(gstrPrivs, ";���׷����༭;") = 0 Or InStr(gstrPrivs, ";�������׷���;") = 0 Then
        MsgBox "��û�г��׷����Ĺ�����༭Ȩ�ޣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrPrivs = Replace(gstrPrivs, ";��Ŀ����;", ";")
    gstrPrivs = Replace(gstrPrivs, ";�շ�����;", ";")
    gstrPrivs = Replace(gstrPrivs, ";����������Ŀ;", ";")
    gstrPrivs = Replace(gstrPrivs, ";������ҩ�䷽;", ";")
    
    Call GetUserInfo
    
    '��ʼ�����Ĳ���
    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs)
    
    frmClinicLists.ShowMeWithScheme frmMain, int��Χ
End Sub

Public Sub CallOPSEmpower(frmMain As Object, cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strDBUser As String)
'���ܣ�ҽ��վ����������˹���

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = 1054 '1054-������Ŀ����
    
    'Ȩ���жϺʹ���
    gstrPrivs = GetPrivFunc(lngSys, 1080)
    If gstrPrivs = "" Then
        MsgBox "��û��������Ȩ����Ȩ�ޣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Sub
    End If

    Call GetUserInfo
    
    frmOPSEmpower.Show , frmMain
End Sub

Public Sub CallSetClinicCharge(ByVal lngDeptID As Long, ByVal lngMode As Long, frmMain As Object, cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strDBUser As String, ByVal int���� As Integer, ByVal blnModify As Boolean)
'���ܣ�����������Ŀ ���� �շѶ���
'������ lngDeptID ����ID
'       lngMode=�Ƿ�ģ̬��ʾ����
'       int����= 1-���� ��2-סԺ
    '------------------------------------------------
    Dim lngModul As Long
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    
    If InStr(frmMain.Caption, "����ҽ������վ -") > 0 Then
        lngModul = 1260
    ElseIf InStr(frmMain.Caption, "סԺ��ʿ����վ -") > 0 Then
        lngModul = 1262
    ElseIf InStr(frmMain.Caption, "ҽ������վ -") > 0 Then
        lngModul = 1263
    ElseIf InStr(frmMain.Caption, "�°�סԺ��ʿ����վ") > 0 Then
        lngModul = 1265
    Else
        MsgBox "��ģ�鲻�ܵ��á�������Ŀ�������á����ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul

    gstrDBOwner = Sys.SystemOwner(lngSys) ' ȡϵͳ������
    gstrUnitName = GetUnitName
    gstrMatch = IIf(Val(zlDatabase.GetPara("����ƥ��", , , True)) = 0, "%", "")

    'Ȩ���жϺʹ���
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngMode) & ";"
'    If InStr(gstrPrivs, ";�����շ���Ŀ����;") = 0 Then
'        MsgBox "��û�в����շ���Ŀ����Ȩ�ޣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
'        Exit Sub
'    End If
    
    Call GetUserInfo
    Call frmSetClinicCharge.ShowMe(lngDeptID, lngMode, frmMain, int����, Not blnModify)
End Sub


Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = gobjKernel.CloseWindows() And (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

