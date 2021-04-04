VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'Ҫ���������
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ���ݿ��¼�û���

    Set gcnOracle = cnMain
    Set gfrmMain = frmMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrUnitName = GetUnitName
    
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not CheckValid Then Exit Sub

    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul

    If glngSys Like "8??" Then App.Title = "�ͻ���ϵ����"
    
    '-------------------------------------------------
    '1100       ��Լ��λ����
    '1101       ������Ϣ����
    '1102       ���￨���Ź���
    '1103       Ԥ�������
    '1104       Ԥ��������ձ�
    '1105       ��Լ��λ����
    '1106       ���˷�������
    Call InitLocPar(lngModul)
    Call CreateSquareCardObject(frmMain, lngModul)
    Select Case lngModul
        Case P��Լ��λ���� '��Լ��λ����
            Call InitLocPar(lngModul)
            frmUnit.Show , frmMain
        Case P������Ϣ���� '������Ϣ����
            Call InitSysPar
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            frmManagePatient.Show , frmMain
'        Case P���￨���Ź��� '���￨���Ź���
'            Call InitSysPar
'            frmManageIDCard.Show , frmMain
        Case PԤ�������  'Ԥ�������
            Call InitSysPar
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            frmManageDeposit.Show , frmMain
        Case P���˷�������  '���˷�������
            Call InitLocPar(lngModul)
            Call InitSysPar
            frmManageExamine.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ���ݿ��¼�û���

    Set gcnOracle = cnMain
    Set gfrmMain = Nothing
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrUnitName = GetUnitName
    
    gstrDBUser = strDBUser
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not CheckValid Then Exit Sub

    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
   
    If glngSys Like "8??" Then App.Title = "�ͻ���ϵ����"
    
    '-------------------------------------------------
    '1100       ��Լ��λ����
    '1101       ������Ϣ����
    '1102       ���￨���Ź���
    '1103       Ԥ�������
    '1104       Ԥ��������ձ�
    '1105       ��Լ��λ����
    '1106       ���˷�������
    Call InitLocPar(lngModul)
    
    Select Case lngModul
        Case P��Լ��λ���� '��Լ��λ����
            Call InitLocPar(lngModul)
            Call CreateSquareCardObject(frmUnit, lngModul)
            zlCommFun.ShowChildWindow frmUnit.hWnd, lngMain
        Case P������Ϣ���� '������Ϣ����
            Call InitSysPar
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(frmManagePatient, lngModul)
            zlCommFun.ShowChildWindow frmManagePatient.hWnd, lngMain
'        Case P���￨���Ź��� '���￨���Ź���
'            Call InitSysPar
'            zlCommFun.ShowChildWindow frmManageIDCard.hwnd, lngMain
        Case PԤ������� 'Ԥ�������
            Call InitSysPar
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(frmManageDeposit, lngModul)
            zlCommFun.ShowChildWindow frmManageDeposit.hWnd, lngMain
        Case P���˷������� '���˷�������
            Call InitLocPar(lngModul)
            Call InitSysPar
            Call CreateSquareCardObject(frmManageExamine, lngModul)
            zlCommFun.ShowChildWindow frmManageExamine.hWnd, lngMain
    End Select
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
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Function NurseDeposit(ByVal lngSys As Long, cnMain As ADODB.Connection, frmMain As Object, _
    ByVal strDBUser As String, Optional lng����ID As Long, Optional lng��ҳID As Long, _
    Optional ByVal bln����˿� As Boolean = True, Optional ByVal bytPrepayType As Byte = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ� ��ʿվ��������˿��
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '   strDBUser:��ǰ���ݿ��¼�û���
    '   lng����ID-ȱʡ�Ĳ���ID
    '   lng��ҳID-ȱʡ����ҳID
    '   bytPrepayType-Ԥ������(0-�����סԺ;1-����;2-סԺ)
    '����:
    '����:������
    '����:2016-09-05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNursePrivs As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
    glngSys = lngSys
    gstrPrivs = ";����;סԺԤ��;Ԥ�������˿�;Ԥ���˿�;Ԥ���տ�;�����ɿ�"
    If bytPrepayType <> 2 Then gstrPrivs = gstrPrivs & ";����Ԥ��;"
    strNursePrivs = GetPrivFunc(lngSys, 9000)
    
    Call GetUserInfo
    
    '-------------------------------------------------
    '1103      Ԥ�����տ��
    Call InitLocPar(1103)
    Call InitSysPar
    
    'ҽ���ӿ�
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(frmMain, 1103)
    
    If strNursePrivs = "" Then
        MsgBox "�㲻�߱���ʿվʹ������˿��Ȩ��,����ϵϵͳ����Ա!", vbInformation, gstrSysName
        Exit Function
    Else
        strSQL = "Select 1 From ������ҳ Where ����ID=[1] And ��ҳID=[2] And ��Ժ���� Is Null"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "��Ժ״̬", lng����ID, lng��ҳID)
        If rsTmp.EOF Then
            '��Ժ
            If InStr(";" & strNursePrivs & ";", ";��Ժ��������˿�;") = 0 Then
                MsgBox "�㲻�߱��Գ�Ժ����ʹ������˿��Ȩ��,����ϵϵͳ����Ա!", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            '��Ժ
            If InStr(";" & strNursePrivs & ";", ";��Ժ��������˿�;") = 0 Then
                MsgBox "�㲻�߱�����Ժ����ʹ������˿��Ȩ��,����ϵϵͳ����Ա!", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If

    NurseDeposit = frmDeposit.zlShowEdit(frmMain, 0, IIf(bln����˿�, 3, 0), gstrPrivs, 1103, _
        bytPrepayType, , , , lng����ID, lng��ҳID, 0, True)
End Function

Public Function PlusDeposit(ByVal lngSys As Long, cnMain As ADODB.Connection, frmMain As Object, _
    ByVal strDBUser As String, Optional bytCallObject As Byte = 0, _
    Optional lng����ID As Long, Optional lng��ҳID As Long, _
    Optional dblDefPrePayMoney As Double = 0, _
    Optional bytԤ������ As Integer = 0, _
    Optional blnOneCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ� ����Ԥ�����տ��
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '   strDBUser:��ǰ���ݿ��¼�û���
    '  bytCallObject:���˺����(0-Ԥ�������(ȱʡ��);1-���˷��ò�ѯ����,2-ҽ�ƿ�����)
    '  lng����ID-ȱʡ�Ĳ���ID
    '  lng��ҳID-ȱʡ����ҳID
    '  dblDefPrePayMoney-ȱʡ��Ԥ�����
    '  bytԤ������: 0-��Ԥ����(ȱʡ,���л�����),1-�������(1),2-����״̬(1); 3-����˿�(37770), 4-����תסԺ;5-סԺת����
    '����:
    '����:���˺�
    '����:2011-02-17 17:18:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt���� As Byte
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, PԤ�������) & ";"
    Call GetUserInfo
    If gstrPrivs = ";;" Then
         MsgBox "�㲻�߱���Ԥ�����Ȩ��,���ܽ��н�Ԥ��!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    
    '-------------------------------------------------
    '1103      Ԥ�����տ��
    Call InitLocPar(PԤ�������)
    Call InitSysPar
    
    'ҽ���ӿ�
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(frmMain, PԤ�������)
    If bytԤ������ = 4 Then
        byt���� = 1
    ElseIf bytԤ������ = 5 Then
        byt���� = 2: bytԤ������ = 4
    Else
        byt���� = IIf(lng��ҳID <> 0, 2, 1)
    End If
    '0-��Ԥ����(ȱʡ,���л�����),1-�������(1),2-����״̬(1);3-����˿�(37770)
    PlusDeposit = frmDeposit.zlShowEdit(frmMain, bytCallObject, bytԤ������, gstrPrivs, PԤ�������, byt����, , , , lng����ID, lng��ҳID, dblDefPrePayMoney, , blnOneCard)
End Function

Public Function RefundDeposit(ByVal lngSys As Long, cnMain As ADODB.Connection, frmMain As Object, _
    ByVal strDBUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ� ��������˿��
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '   strDBUser:��ǰ���ݿ��¼�û���
    '����:
    '����:���˺�
    '����:2011-02-17 17:18:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim byt���� As Byte
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, PԤ�������) & ";"
    Call GetUserInfo
    If gstrPrivs = ";;" Then
         MsgBox "�㲻�߱�Ԥ���������Ȩ��,���ܽ�������˿�!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    
    '-------------------------------------------------
    '1103      Ԥ�����տ��
    Call InitLocPar(PԤ�������)
    Call InitSysPar
    
    'ҽ���ӿ�
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(frmMain, PԤ�������)

    RefundDeposit = frmDeposit.zlShowEdit(frmMain, 0, 3, gstrPrivs, PԤ�������)
End Function

Public Function IDCard(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, _
    ByVal strDBUser As String, Optional lngModule As Long = 0, _
    Optional lng�����ID As Long, Optional lng����ID As Long = 0, Optional strPrivs As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ֱ�ӵ��÷�������
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-26 12:17:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCardObject As Object, strExpend As String
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    glngSys = lngSys: glngModul = lngModule: gstrPrivs = strPrivs
    gstrDBUser = strDBUser: Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrUnitName = GetUnitName
    Set objCardObject = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        MsgBox "ҽ�ƿ�����������,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If objCardObject.zlInitComponents(frmMain, lngModule, lngSys, strDBUser, cnMain, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Function
    End If
    '���÷���
    '���:frmMain-���õ�������
    '        lngModule-ģ���
    '        lngCardTypeID-ҽ�ƿ����
    '        lng����ID-����ָ���Ĳ���ID,��ָ�����˽��з���
    '        strPrivs-���Դ���գ������ʱȡҽ�ƿ�����ģ������Ȩ��
    '����:48416
    IDCard = objCardObject.zlSendCard(frmMain, lngModule, lng�����ID, lng����ID, strPrivs)
    Set objCardObject = Nothing
End Function

Private Sub Class_Terminate()
    '62646:������,2013-06-18
    If Forms.Count = 0 Then
        Call CloseSquareCardObject '�رտ�������Ϣ
    End If
End Sub