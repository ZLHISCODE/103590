VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDrugStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, FrmMain As Object, strUsername As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '���أ�
    '------------------------------------------------
    Dim frmTemp As Form
    Dim strTitle As String
    Dim rsTemp As ADODB.Recordset
    Dim strDept As String
    Dim strPrivs As String

    Set gcnOracle = cnMain
    Set gobjBrower = FrmMain
    
    On Error GoTo errHandle
    
    glngSys = lngSys
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("kind"), Default:="")
    
    gstrMatchMethod = GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0)
    
    gbytSimpleCodeTrans = Val(zlDataBase.GetPara("����ƥ�䷽ʽ�л�", , , 1))
    
    gstrUnitName = GetUnitName()
    gstrDbUser = strUsername
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    glngModul = lngModul
    gstrprivs = GetPrivFunc(glngSys, lngModul)
    App.Title = IIf(glngSys \ 100 = 1, "ҩ������", "����ҩ�����")
    If Not CheckValid Then Exit Sub
    
    '��ȡҩ������Ҫʹ�õ�ϵͳ����ֵ
    Call GetSysParms
    
    'ȡҩƷ�۸�ȼ�
    Call GetPriceClass
    
    '-------------------------------------------------
    'PASS�ӿڳ�ʼ��:ֻ��ʼ��һ��
    If IsHavePrivs(gstrprivs, "������ҩ���") Then
        If InStr(",1341,1342,1345,", "," & lngModul & ",") > 0 And gobjPass Is Nothing Then
            On Error Resume Next
            Set gobjPass = CreateObject("zlpassinterface.clspass")
            Error.Clear: On Error GoTo 0
            
            If Not gobjPass Is Nothing Then
                If Not gobjPass.zlPassInit_YF(gcnOracle, glngSys, lngModul) Then
                    Set gobjPass = Nothing
                Else
                    gintPass = PassUseable
                End If
            End If
        End If
    End If
    '------------------------------------------------
    
    '1340    ��ҩ���ڹ���
    '1341    ҩƷ������ҩ
    '1342    ҩƷ���ŷ�ҩ
    '1343    ҩƷ�������
    '1344    Э��ҩƷ���
    '1345    PIVA����
    
    Call setNOtExcetePrice '����ʱ�仹û����Ч�ļ۸�ִ����
    
    Select Case lngModul
    Case 1332 'ҩƷ���
        frmClosingAccount.Show , FrmMain
    Case 1340
        frm��ҩ����.Show , FrmMain
    Case 1341
        '���ݲ�������
        Call LoadBillControl
        
        frmҩƷ������ҩNew.Show , FrmMain
    Case 1342
        '���ݲ�������
        Call LoadBillControl
        
        frm���ŷ�ҩ����New.Show , FrmMain
        If frm���ŷ�ҩ����New.mblnEnter = False Then Unload frm���ŷ�ҩ����New
    Case 1343
        strTitle = "ҩƷ�������"
        For Each frmTemp In Forms
            If frmTemp.Caption = strTitle Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        frmRequestDrugList.ShowList FrmMain
    Case 1344
        
        strTitle = "Э��ҩƷ���"
        For Each frmTemp In Forms
            If frmTemp.Caption = strTitle Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmTemp = New frmMainList
        frmTemp.ShowList lngModul, strTitle, FrmMain
    Case 1345
        '��鲿������
        
        gstrSQL = "Select Distinct A.ID, A.����" & _
            " From ���ű� A, ��������˵�� B " & _
            " Where A.ID = B.����id And B.�������� = '��������' And " & _
            " B.����id In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��') " & _
            " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ��������")
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "���ڲ��Ź�����������Һ�������Ĳ���(�������ҩ����������������)��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��鲿����Ա
        gstrSQL = "Select Distinct P.ID, P.����" & _
            " From ���ű� P " & _
            " Where (P.վ�� = '" & gstrNodeNo & "' Or P.վ�� is Null) And P.ID In (Select Distinct A.����id " & _
            " From ������Ա A, ��������˵�� B " & _
            " Where A.��Աid = [1] And A.����id = B.����id And B.�������� = '��������' And " & _
            " B.����id In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��')) And " & _
            " (P.����ʱ�� Is Null Or P.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ����������Ա", glngUserId)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "�㲻����Һ����������Ա������ʹ�ñ�ģ�飡", vbInformation, gstrSysName
            Exit Sub
        End If

        frmPIVAMain.Show , FrmMain
    Case 1347
        frm�󴦷����.Show , FrmMain
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function AutoSetBatch(ByVal lng�ⷿid As Long, ByVal str���ͺ� As String, cnMain As ADODB.Connection) As Boolean
    '�Զ�����������������
    'lng�ⷿid���������Ĳ���id
    'str���ͺţ�ҽ�����ͺ�
'
End Function
Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, strUsername As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   lngMain:��������
    '���أ�
    '------------------------------------------------
    Dim frmTemp As Form
    Dim strTitle As String
    Dim rsTemp As ADODB.Recordset
    Dim strDept As String
    Dim strPrivs As String

    Set gcnOracle = cnMain
    
    glngSys = lngSys
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("kind"), Default:="")
    
    gstrMatchMethod = GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0)
    
    gbytSimpleCodeTrans = Val(zlDataBase.GetPara("����ƥ�䷽ʽ�л�", , , 1))
    
    gstrUnitName = GetUnitName()
    gstrDbUser = strUsername
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    glngModul = lngModul
    gstrprivs = GetPrivFunc(glngSys, lngModul)
    App.Title = IIf(glngSys \ 100 = 1, "ҩ������", "����ҩ�����")
    If Not CheckValid Then Exit Sub
    
    '��ȡҩ������Ҫʹ�õ�ϵͳ����ֵ
    Call GetSysParms
    
    '-------------------------------------------------
    'PASS�ӿڳ�ʼ��:ֻ��ʼ��һ��
    If IsHavePrivs(gstrprivs, "������ҩ���") Then
        If InStr(",1341,1342,1345,", "," & lngModul & ",") > 0 And gobjPass Is Nothing Then
            On Error Resume Next
            Set gobjPass = CreateObject("zlpassinterface.clspass")
            Error.Clear: On Error GoTo 0
            
            If Not gobjPass Is Nothing Then
                If Not gobjPass.zlPassInit_YF(gcnOracle, glngSys, lngModul) Then
                    Set gobjPass = Nothing
                Else
                    gintPass = PassUseable
                End If
            End If
        End If
    End If
    '------------------------------------------------
    
    '1340    ��ҩ���ڹ���
    '1341    ҩƷ������ҩ
    '1342    ҩƷ���ŷ�ҩ
    '1343    ҩƷ�������
    '1344    Э��ҩƷ���
    '1345    PIVA����
    
    Call setNOtExcetePrice '����ʱ�仹û����Ч�ļ۸�ִ����
    
    Select Case lngModul
    Case 1332 'ҩƷ���
        OS.ShowChildWindow frmClosingAccount.hWnd, lngMain
    Case 1340
        OS.ShowChildWindow frm��ҩ����.hWnd, lngMain
    Case 1341
        '���ݲ�������
        Call LoadBillControl
        
        OS.ShowChildWindow frmҩƷ������ҩNew.hWnd, lngMain
    Case 1342
        '���ݲ�������
        Call LoadBillControl
        
        OS.ShowChildWindow frm���ŷ�ҩ����New.hWnd, lngMain
        If frm���ŷ�ҩ����New.mblnEnter = False Then Unload frm���ŷ�ҩ����New
    Case 1343
        strTitle = "ҩƷ�������"
        For Each frmTemp In Forms
            If frmTemp.Caption = strTitle Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        frmRequestDrugList.ShowList lngMain
    Case 1344
        
        strTitle = "Э��ҩƷ���"
        For Each frmTemp In Forms
            If frmTemp.Caption = strTitle Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmTemp = New frmMainList
        frmTemp.ShowList lngModul, strTitle, lngMain
    Case 1345
        '��鲿������
        
        gstrSQL = "Select Distinct A.ID, A.����" & _
            " From ���ű� A, ��������˵�� B " & _
            " Where A.ID = B.����id And B.�������� = '��������' And " & _
            " B.����id In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��') " & _
            " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ��������")
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "���ڲ��Ź�����������Һ�������Ĳ���(�������ҩ����������������)��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��鲿����Ա
        gstrSQL = "Select Distinct P.ID, P.����" & _
            " From ���ű� P " & _
            " Where (P.վ�� = '" & gstrNodeNo & "' Or P.վ�� is Null) And P.ID In (Select Distinct A.����id " & _
            " From ������Ա A, ��������˵�� B " & _
            " Where A.��Աid = [1] And A.����id = B.����id And B.�������� = '��������' And " & _
            " B.����id In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��')) And " & _
            " (P.����ʱ�� Is Null Or P.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ����������Ա", glngUserId)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "�㲻����Һ����������Ա������ʹ�ñ�ģ�飡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        OS.ShowChildWindow frmPIVAMain.hWnd, lngMain
    Case 1347
        OS.ShowChildWindow frm�󴦷����.hWnd, lngMain
    End Select
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.count = 0)
End Function

'Private Sub Class_Terminate()
'   '�ڵ���̨�˳�ʱ����,��ʱע��PASS
'    If gintPass = 1 Then
'        Call PassQuit: gintPass = 0
'    End If
'End Sub