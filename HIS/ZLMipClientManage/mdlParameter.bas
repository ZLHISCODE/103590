Attribute VB_Name = "mdlParameter"
Option Explicit
Private mstr�û��� As String
Private mstr������ As String

Public Sub UpdateParameters()
'���ܣ���ԭ������ע������ֵ������������
    Dim rsTmp As New ADODB.Recordset
    Dim rsSys As New ADODB.Recordset
    Dim rsUpgrade As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select User as �û���,SYS_CONTEXT('USERENV','TERMINAL') as ������ From Dual"
    Call zlDataBase.OpenRecordset(rsTmp, strSQL, "UpdateParameters")
    mstr�û��� = rsTmp!�û���: mstr������ = rsTmp!������
    
    strSQL = "Select Trunc(���/100) as ϵͳ From zlSystems Where �汾�� Like '10.%'"
    Call zlDataBase.OpenRecordset(rsSys, strSQL, "UpdateParameters")
    
    strSQL = "Select Trunc(ϵͳ/100) as ϵͳ From zlUpgrade" & _
        " Where ϵͳ Is Not Null And ԭʼ�汾 Like '10.%'" & _
        " And ԭʼ�汾>='10.24.0' And Substr(Ŀ��汾,1,5)>Substr(ԭʼ�汾,1,5)"
    Call zlDataBase.OpenRecordset(rsUpgrade, strSQL, "UpdateParameters")
    
    On Error GoTo 0
    
    '�����׼��Ĳ���ֵ����
    '-----------------------------------------------------------------
    rsSys.Filter = "ϵͳ=1": rsUpgrade.Filter = "ϵͳ=1"
    If Not rsSys.EOF And rsUpgrade.EOF Then
        '�ҺŹ���
        Call UpdateParameterValue(rsSys!ϵͳ, 1111, "����ģ��\zl9RegEvent", "ȱʡ���ʽ", "ȱʡ���ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1111, "����ģ��\zl9RegEvent", "ȱʡ�ѱ�", "ȱʡ�ѱ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1111, "����ģ��\zl9RegEvent", "ȱʡ�Ա�", "ȱʡ�Ա�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1111, "����ģ��\zl9RegEvent", "ȱʡ���㷽ʽ", "ȱʡ���㷽ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1111, "����ģ��\zl9RegEvent", "�Һſ���", "�Һſ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1111, "����ģ��\zl9RegEvent", "���ùҺ�Ʊ������", "���ùҺ�Ʊ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1111, "����ģ��\zl9RegEvent", "���þ��￨����", "���þ��￨����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1111, "˽��ģ��\" & mstr�û��� & "\zl9RegEvent", "��ǰ�Һ�Ʊ�ݺ�", "��ǰ�Һ�Ʊ�ݺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1111, "˽��ģ��\" & mstr�û��� & "\zl9RegEvent\frmRegist", "ˢ�·�ʽ", "ˢ�·�ʽ")
        '����������
        Call UpdateParameterValue(rsSys!ϵͳ, 1113, "����ģ��\zl9RegEvent", "�������", "�������")
        '�����շ�
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "����ģ��\zl9OutExse", "�շ����", "�շ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "˽��ģ��\" & mstr�û��� & "\zl9OutExse", "ȱʡ�ѱ�", "ȱʡ�ѱ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "˽��ģ��\" & mstr�û��� & "\zl9OutExse\frmManageCharge", "ˢ�·�ʽ", "ˢ�·�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "˽��ģ��\" & mstr�û��� & "\zl9OutExse", "��ҩ�Զ�����", "��ҩ�Զ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "˽��ģ��\" & mstr�û��� & "\zl9OutExse", "��ҩ�Զ����볤��", "��ҩ�Զ����볤��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "˽��ģ��\" & mstr�û��� & "\zl9OutExse", "ȱʡ���㷽ʽ", "ȱʡ���㷽ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "����ģ��\zl9OutExse", "�����շ�Ʊ������", "�����շ�Ʊ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "����ģ��\zl9OutExse", "�ҺŹ����շ�Ʊ��", "�ҺŹ����շ�Ʊ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "����ģ��\zl9OutExse", "�ֹ�����", "�ֹ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "����ģ��\zl9OutExse", "LED��ʾ�շ���ϸ", "LED��ʾ�շ���ϸ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "����ģ��\zl9OutExse", "LED��ʾ��ӭ��Ϣ", "LED��ʾ��ӭ��Ϣ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "˽��ģ��\" & mstr�û��� & "\zl9OutExse", "��ǰ�շ�Ʊ�ݺ�", "��ǰ�շ�Ʊ�ݺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1121, "˽��ģ��\" & mstr�û��� & "\zl9OutExse", "�˷Ѻ�������ģʽ", "�˷Ѻ�������ģʽ")
        
        '���ﻮ��
        Call UpdateParameterValue(rsSys!ϵͳ, 1120, "����ģ��\zl9OutExse", "�շ����", "�շ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1120, "˽��ģ��\" & mstr�û��� & "\zl9OutExse", "ȱʡ�ѱ�", "ȱʡ�ѱ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1120, "˽��ģ��\" & mstr�û��� & "\zl9OutExse\frmManagePrice", "ˢ�·�ʽ", "ˢ�·�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1120, "˽��ģ��\" & mstr�û��� & "\zl9OutExse", "��ҩ�Զ�����", "��ҩ�Զ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1120, "˽��ģ��\" & mstr�û��� & "\zl9OutExse", "��ҩ�Զ����볤��", "��ҩ�Զ����볤��")
        
        '�������
        Call UpdateParameterValue(rsSys!ϵͳ, 1122, "����ģ��\zl9OutExse", "�շ����", "�շ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1122, "˽��ģ��\" & mstr�û��� & "\zl9OutExse\frmManageBilling", "ˢ�·�ʽ", "ˢ�·�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1122, "˽��ģ��\" & mstr�û��� & "\zl9OutExse\frmManageBilling\TabStrip", "ҳ��", "ҳ��")
        
        'סԺ���ʹ���
        Call UpdateParameterValue(rsSys!ϵͳ, 1133, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmManageBilling", "ˢ�·�ʽ", "ˢ�·�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1133, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmManageBilling\TabStrip", "ҳ��", "ҳ��")
        
        '���ҷ�ɢ����
        Call UpdateParameterValue(rsSys!ϵͳ, 1134, "˽��ģ��\" & mstr�û��� & "\zlInExse\frmDeptBilling", "ˢ�·�ʽ", "ˢ�·�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1134, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmDeptBilling\TabStrip", "ҳ��", "ҳ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1134, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmDeptBilling", "��ʾ���˷�ʽ", "��ʾ���˷�ʽ")
        
        'ҽ�����Ҽ���
        Call UpdateParameterValue(rsSys!ϵͳ, 1135, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmTechnoBilling", "ˢ�·�ʽ", "ˢ�·�ʽ")
        
        'ִ�еǼǹ���
        Call UpdateParameterValue(rsSys!ϵͳ, 1142, "����ģ��\zl9InExse", "ҽ��������Դ", "ҽ��������Դ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1142, "����ģ��\zl9InExse", "ҽ�����ﵥ������", "ҽ�����ﵥ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1142, "����ģ��\zl9InExse", "ҽ��סԺ��������", "ҽ��סԺ��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1142, "����ģ��\zl9InExse", "ҽ����쵥������", "ҽ����쵥������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1142, "����ģ��\zl9InExse", "ҽ��ִ�����", "ҽ��ִ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1142, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ʾ����ͷ", "��ʾ����ͷ", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1142, "����ģ��\zl9InExse", "ҽ��������Ŀͬʱѡ��", "������Ŀͬʱѡ��")
        
        'סԺ���ʲ���
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmPatiSelect", "��ʾ���˷�ʽ", "��ʾ���˷�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ʾ��Ժ����", "��ʾ��Ժ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ʾԤ��Ժ����", "��ʾԤ��Ժ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ʾ��Ժ����", "��ʾ��Ժ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmReCharge", "���ÿ�ʼʱ��", "���ÿ�ʼʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmReCharge", "�����ڼ�", "�����ڼ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmReCharge", "��˿�ʼʱ��", "��˿�ʼʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ҩ�Զ�����", "��ҩ�Զ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ҩ�Զ����볤��", "��ҩ�Զ����볤��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "����ģ��\zl9InExse", "�������۲��˼���", "�������۲��˼���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "����ģ��\zl9InExse", "סԺ���۲��˼���", "סԺ���۲��˼���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1150, "����ģ��\zl9InExse", "�շ����", "�շ����")
        
        '���˽��ʹ���
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "����Ʊ������", "����Ʊ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "����ģ��\zl9InExse", "���ý���Ʊ������", "���ý���Ʊ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "����ģ��\zl9InExse", "LED��ʾ��ӭ��Ϣ", "LED��ʾ��ӭ��Ϣ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ǰ����Ʊ�ݺ�", "��ǰ����Ʊ�ݺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmManageBalance", "ˢ�·�ʽ", "ˢ�·�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmManageDue\TabStrip", "ҳ��", "����Ӧ�տ�ҳ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "˽��ģ��\" & mstr�û��� & "\zl9InExse\frmPatiSelect", "��ʾ���˷�ʽ", "��ʾ���˷�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ʾ��Ժ����", "��ʾ��Ժ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ʾԤ��Ժ����", "��ʾԤ��Ժ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ʾ��Ժ����", "��ʾ��Ժ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "��ʾ���岡��", "��ʾ���岡��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1137, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "Ĭ�ϳ�Ժ����", "Ĭ�ϳ�Ժ����", True)
        
        'һ���嵥����
        Call UpdateParameterValue(rsSys!ϵͳ, 1141, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "һ���嵥�����˲���ģʽ", "���˲���ģʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1141, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "һ���嵥������ʱ��", "����ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1141, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "һ���嵥���������", "�������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1141, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "һ���嵥����ʼʱ��", "��ʼʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1141, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "һ���嵥����ʼ���", "��ʼ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1141, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "һ���嵥����ҽ������", "��ҽ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1141, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "һ���嵥��ҽ������", "ҽ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1141, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "һ���嵥����Ժ����", "��Ժ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1141, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "һ���嵥����Ժ����", "��Ժ����")
        
        '���˷��ò�ѯ
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "�嵥����", "�嵥����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "ViewDate", "����ʱ������", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "ViewCancel״̬", "��ʾ��������", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "ViewZero״̬", "��ʾ�����", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "����ģ��\zl9InExse", "��ʾ������", "��ʾ������", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "����ģ��\zl9InExse", "�ֿ�ģʽ", "�ֿ�ģʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "����ģ��\zl9InExse", "����ģʽ", "����ģʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "����״̬", "����״̬")
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "Ƿ�Ѳ�ѯ-�������", "�������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "ViewOwe״̬", "����δ���岡��", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "ViewUnAudit״̬", "����δ��˲���", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1139, "˽��ģ��\" & mstr�û��� & "\zl9InExse", "Ƿ�Ѳ�ѯ-������ʾ", "������ʾ")
        
        'Ʊ��ʹ�ü��
        Call UpdateParameterValue(rsSys!ϵͳ, 1501, "˽��ģ��\" & mstr�û��� & "\zL9CashBill\frmBillSupervise\Menu", "mnuViewAll״̬", "��ʾ�������ü�¼", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1501, "˽��ģ��\" & mstr�û��� & "\zL9CashBill\frmBillSupervise\Menu", "�鿴�˶���Ϣ", "�鿴�˶���Ϣ", True)
        
        '�շѲ�����
        Call UpdateParameterValue(rsSys!ϵͳ, 1500, "˽��ģ��\" & mstr�û��� & "\zL9CashBill\frmCashSupervise\Menu", "mnuViewAll״̬", "��ʾ�����տ�Ա", True)
        
    
        '1260-����ҽ��վ
        Call UpdateParameterValue(rsSys!ϵͳ, 1260, "����ģ��\zl9CISJob", "��������", "��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1260, "����ģ��\zl9CISJob", "�����������", "�����������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1260, "����ģ��\zl9CISJob", "���ﷶΧ", "���ﷶΧ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1260, "����ģ��\zl9CISJob", "�������", "�������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1260, "˽��ģ��\" & mstr�û��� & "\zl9CISJob", "����ҽ��", "����ҽ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1260, "˽��ģ��\" & mstr�û��� & "\zl9CISJob", "���ﲡ�˽������", "���ﲡ�˽������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1260, "˽��ģ��\" & mstr�û��� & "\zl9CISJob", "���ﲡ�˿�ʼ���", "���ﲡ�˿�ʼ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1260, "˽��ģ��\" & mstr�û��� & "\zl9CISJob\frmOutDoctorStation", "ҽ������", "ҽ������")
        
        '1261-סԺҽ��վ
        Call UpdateParameterValue(rsSys!ϵͳ, 1261, "˽��ģ��\" & mstr�û��� & "\zl9CISJob\frmAuditResponse", "��������-������", "�����鷴��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1261, "˽��ģ��\" & mstr�û��� & "\zl9CISJob\frmAuditResponse", "��������-�ύ���", "�ύ��鷴��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1261, "˽��ģ��\" & mstr�û��� & "\zl9CISJob\frmInDoctorStation", "ҽ������", "ҽ������")
        
        '1262-סԺ��ʿվ
        Call UpdateParameterValue(rsSys!ϵͳ, 1262, "˽��ģ��\" & mstr�û��� & "\zl9CISJob\frmAuditResponse", "��������-������", "�����鷴��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1262, "˽��ģ��\" & mstr�û��� & "\zl9CISJob\frmAuditResponse", "��������-�ύ���", "�ύ��鷴��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1262, "˽��ģ��\" & mstr�û��� & "\zl9CISJob\frmInNurseStation", "Filter��ǰ����", "��ǰ��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1262, "˽��ģ��\" & mstr�û��� & "\zl9CISJob\frmInNurseStation", "Filter����ȼ�", "����ȼ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1262, "˽��ģ��\" & mstr�û��� & "\zl9CISJob\frmInNurseStation", "ҽ������", "ҽ������")
        
        '1263-ҽ������վ
        Call UpdateParameterValue(rsSys!ϵͳ, 1263, "����ģ��\zl9CISJob", "��¼ִ�����", "��¼ִ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1263, "˽��ģ��\" & mstr�û��� & "\zl9CISJob\frmTechnicStation", "ҽ������", "ҽ������")
        '����������⴦��
        Set rsTmp = New ADODB.Recordset
        strSQL = "Select A.����ID From ������Ա A,�ϻ���Ա�� B Where A.��ԱID=B.��ԱID And B.�û���=User"
        Call zlDataBase.OpenRecordset(rsTmp, strSQL, "UpdateParameters")
        Do While Not rsTmp.EOF
            Call UpdateParameterValue(rsSys!ϵͳ, 1263, "����ģ��\zl9CISJob\����" & rsTmp!����ID, "ִ�м䷶Χ", "ִ�м䷶Χ")
            rsTmp.MoveNext
        Loop

        '1252-����ҽ���´�
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel", "����ȱʡ��ҩ��", "����ȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel", "����ȱʡ���ϲ���", "����ȱʡ���ϲ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel", "����ȱʡ��ҩ��", "����ȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel", "����ȱʡ��ҩ��", "����ȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockOutAdvice", "FilterAutoHide", "���������Զ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockOutAdvice", "Filter����Ӥ��", "����Ӥ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockOutAdvice", "Filter����ҽ��", "����ҽ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockOutAdvice", "Filter��Ҫ����", "��Ҫ�������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockOutAdvice", "Filterҽ��״̬", "ҽ��״̬����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockOutAdvice", "ҽ�����б�", "ҽ�����б�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmLisView", "���ؼ���ͼ��", "���ؼ���ͼ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel\frmLisRptGeneral", "�鿴����", "�鿴����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel\frmLisRptGeneral", "�鿴��־", "�鿴��־")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel\frmLisRptGeneral", "�鿴��λ", "�鿴��λ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel\frmLisRptGeneral", "�鿴�ο�", "�鿴�ο�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel\frmLisRptGeneral", "�鿴ø��", "�鿴ø��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel\frmLisRptGeneral", "�鿴��ע", "�鿴��ע")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel\frmLisRptGeneral", "�鿴���", "�鿴���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1252, "����ģ��\zlCISKernel\frmLisRptMicrobiology", "�ϴν��", "�ϴν��")
        
        '1253-סԺҽ���´�
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "����ģ��\zlCISKernel", "סԺȱʡ��ҩ��", "סԺȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "����ģ��\zlCISKernel", "סԺȱʡ���ϲ���", "סԺȱʡ���ϲ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "����ģ��\zlCISKernel", "סԺȱʡ��ҩ��", "סԺȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "����ģ��\zlCISKernel", "סԺȱʡ��ҩ��", "סԺȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockInAdvice", "FilterAutoHide", "���������Զ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockInAdvice", "Filter����Ӥ��", "����Ӥ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockInAdvice", "Filter����ҽ��", "����ҽ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockInAdvice", "Filter��Ҫ����", "��Ҫ�������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockInAdvice", "Filterҽ����Ч", "ҽ����Ч����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockInAdvice", "Filterҽ��״̬", "ҽ��״̬����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockInAdvice", "Filter����ҽ��", "����ҽ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1253, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDockInAdvice", "ҽ�����б�", "ҽ�����б�")
        
        '1254-סԺҽ������
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "����ģ��\zlCISKernel", "ȱʡ�������", "ȱʡ�������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "����ģ��\zlCISKernel", "ȱʡ�������", "ȱʡ�������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "����ģ��\zlCISKernel", "ȱʡ����ҩ��", "ȱʡ����ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceOperateCond", "�ϴο�ʼ��ͣ", "�ϴο�ʼ��ͣ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceReport", "���ñ�����", "���ñ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceReport", "���ñ���������", "���ñ���������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceReport", "���ñ������ʱ��", "���ñ������ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceReport", "���ñ���ʼ���", "���ñ���ʼ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceReport", "���ñ���ʼʱ��", "���ñ���ʼʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceReport", "���ñ�����Ч", "���ñ�����Ч")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceRollSendCond", "�����ջز���", "�����ջز���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendDrugCond", "���ƽ���ʱ��", "ҩ���������ƽ���ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendDrugCond", "ҩ�����Ͳ���", "ҩ�����Ͳ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendDrugCond", "ҩ����ҩ;��", "ҩ�����͸�ҩ;��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendDrugCond", "ҩ������ʱ��", "ҩ�����ͽ���ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendDrugCond", "ҩ������ʱ��", "ҩ�����ͽ���ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendDrugCond", "ҩ��ʱ����", "ҩ������ʱ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendDrugCond", "ҩ��ҩ���û�", "ҩ������ҩ���û�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendDrugCond", "ҩ��ҽ����Ч", "ҩ������ҽ����Ч")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendOtherCond", "��ҩ���Ͳ���", "�������Ͳ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendOtherCond", "��ҩ����ʱ��", "�������ͽ���ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendOtherCond", "��ҩ����ʱ��", "�������ͽ���ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendOtherCond", "��ҩʱ����", "��������ʱ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendOtherCond", "��ҩҽ����Ч", "��������ҽ����Ч")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmAdviceSendOtherCond", "��ҩ�������", "���������������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDrugSendQueryCond", "��ҩ��ѯ���", "��ҩ��ѯ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDrugSendQueryCond", "ҩ�Ʋ�ѯ����", "ҩ�Ʋ�ѯ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDrugSendQueryCond", "ҩ�Ʋ�ѯ��Ժ����", "ҩ�Ʋ�ѯ��Ժ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDrugSendQueryCond", "ҩ�Ʋ�ѯ���", "ҩ�Ʋ�ѯ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDrugSendQueryCond", "ҩ�Ʋ�ѯ��Ч", "ҩ�Ʋ�ѯ��Ч")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDrugSendQueryCond", "ҩ�Ʋ�ѯҩ��", "ҩ�Ʋ�ѯҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDrugSendQueryCond", "ҩ�Ʋ�ѯ״̬", "ҩ�Ʋ�ѯ״̬")
        Call UpdateParameterValue(rsSys!ϵͳ, 1254, "˽��ģ��\" & mstr�û��� & "\zlCISKernel\frmDrugSendQueryCond", "ҩ����ҩ;��", "ҩ�Ʋ�ѯ��ҩ;��")
        
        '1255-�����¼����
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "����ģ��\zlRichEPR\���µ���ӡѡ��", "����ӡ�������ͼ��", "����ӡ�������ͼ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\frmCaseTendBodyPrintSet", "������ӡ", "������ӡ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\frmCaseTendBodyPrintSet", "��ӡҳ��", "��ӡҳ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\frmCaseTendBodyPrintSet", "��ʼҳ��", "��ʼҳ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\���µ���ӡѡ��", "��ӡ����", "��ӡ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\frmCaseTendSign", "chkEsign", "��������ǩ��")
        
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ӡ����", "��ӡ��", "���µ���ӡ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ӡ����", "ֽ��", "���µ�ֽ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ӡ����", "���", "���µ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ӡ����", "�߶�", "���µ��߶�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ӡ����", "ֽ��", "���µ�ֽ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ӡ����", "��ֽ", "���µ���ֽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ӡ����", "��߾�", "���µ���߾�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ӡ����", "�ұ߾�", "���µ��ұ߾�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ӡ����", "�ϱ߾�", "���µ��ϱ߾�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1255, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ӡ����", "�±߾�", "���µ��±߾�")
        
        '1257-ҽ�����ѹ���
        Call UpdateParameterValue(rsSys!ϵͳ, 1257, "����ģ��\zlCISKernel", "����ȱʡ��ҩ��", "����ȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1257, "����ģ��\zlCISKernel", "����ȱʡ��ҩ��", "����ȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1257, "����ģ��\zlCISKernel", "����ȱʡ��ҩ��", "����ȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1257, "����ģ��\zlCISKernel", "�շ����", "�շ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1257, "����ģ��\zlCISKernel", "סԺȱʡ��ҩ��", "סԺȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1257, "����ģ��\zlCISKernel", "סԺȱʡ��ҩ��", "סԺȱʡ��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1257, "����ģ��\zlCISKernel", "סԺȱʡ��ҩ��", "סԺȱʡ��ҩ��")

        '1264-������Һ�Ŷ�
        Call UpdateParameterValue(rsSys!ϵͳ, 1264, "����ģ��\zl9Transfusion", "��ʾ��������", "��ʾ��������")
        
        '�����Ӧ����ϵͳ�����ģ��
        '���˺�
        '�������
        Call UpdateParameterValue(rsSys!ϵͳ, 1323, "˽��ģ��\" & mstr�û��� & "\zl9Due\�������", "��ͷ", "�����ͷ�б�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1323, "˽��ģ��\" & mstr�û��� & "\zl9Due\�������", "������ϸ", "������ϸ�б�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1323, "˽��ģ��\" & mstr�û��� & "\zl9Due\�������", "������Ϣ", "���ʽ�б�")
        'Ӧ����ѯ
        Call UpdateParameterValue(rsSys!ϵͳ, 1324, "˽��ģ��\" & mstr�û��� & "\zl9Due\Ӧ�����ѯ", "��λID", "���ѡ��λID")
        Call UpdateParameterValue(rsSys!ϵͳ, 1324, "˽��ģ��\" & mstr�û��� & "\zl9Due\Ӧ�����ѯ", "��Ӧ�����", "�����Ϣ�б�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1324, "˽��ģ��\" & mstr�û��� & "\zl9Due\Ӧ�����ѯ", "Ӧ�����ѯ-������ϸ", "������ϸ�б�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1324, "˽��ģ��\" & mstr�û��� & "\zl9Due\Ӧ�����ѯ", "Ӧ�����ѯ-�Ѹ���ϸ", "�Ѹ���ϸ�б�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1324, "˽��ģ��\" & mstr�û��� & "\zl9Due\Ӧ�����ѯ", "Ӧ�����ѯ-δ����ϸ", "δ����ϸ�б�")
        
        '������������ϵͳ�����ģ��
        '��������Ŀ¼����
        Call UpdateParameterValue(rsSys!ϵͳ, 1711, "˽��ģ��\" & mstr�û��� & "\������ʾģʽ", "��ʾ�¼�", "�����¼�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1711, "����ģ��\zl9Stuff\��������ģʽ", "Ʒ��", "Ʒ������ģʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1711, "����ģ��\zl9Stuff\��������ģʽ", "Ʒ��->���", "Ʒ�ֹ��ģʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1711, "����ģ��\zl9Stuff\��������ģʽ", "���", "�������ģʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1711, "����ģ��\zl9Stuff\�������Ϲ��༭", "ָ�������", "�ϴ�ָ�������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1711, "����ģ��\zl9Stuff\�������Ϲ��༭", "�ӳ���", "�ϴμӳ���")
        '�⹺�⹺���
        Call UpdateParameterValue(rsSys!ϵͳ, 1712, "˽��ģ��\" & mstr�û��� & "\��������\�����⹺��ⵥ\BillEdit", "mshBill���", "�����п�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1712, "˽��ģ��\" & mstr�û��� & "\��������\�����⹺��ⵥ\BillEdit", "mshBill����", "������ͷ�ı�")
        
        '�����Ź���
        Call UpdateParameterValue(rsSys!ϵͳ, 1723, "˽��ģ��\" & mstr�û��� & "\zl9Stuff\δ�����嵥", "���ݸ�ʽ", "���ϵ��ݴ�ӡ��ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1723, "����ģ��\zl9Stuff\���ķ��Ź���", "��ӡ��ʽ", "���ϴ�ӡ���ѷ�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1723, "����ģ��\zl9Stuff\���ķ��Ź���", "ҵ������", "��ѯҵ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1723, "����ģ��\zl9Stuff\�����ݽ�������", "��������", "������ϵ�������")
        
        '-- ������Ŀ����
        Call UpdateParameterValue(rsSys!ϵͳ, 1054, "����ģ��\zl9CISBase\������Ŀ����", "����", "������Ŀ��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1054, "˽��ģ��\" & mstr�û��� & "\zl9CISBase\frmClinicLists", "��ʾͣ����Ŀ", "��ʾͣ����Ŀ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1054, "˽��ģ��\" & mstr�û��� & "\zl9CISBase\frmClinicFind", "ƥ�䷽ʽ", "ƥ�䷽ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1054, "˽��ģ��\" & mstr�û��� & "\zl9CISBase\frmClinicFind", "���ҷ�Χ", "���ҷ�Χ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1054, "˽��ģ��\" & mstr�û��� & "\zl9CISBase\frmClinicFind", "���ִ�Сд", "���ִ�Сд")
        Call UpdateParameterValue(rsSys!ϵͳ, 1054, "˽��ģ��\" & mstr�û��� & "\zl9CISBase\frmClinicFind", "���ұ���", "���ұ���")
        
        '- 1059 ������Ŀ����
        Call UpdateParameterValue(rsSys!ϵͳ, 1059, "����ģ��\zl9CISBase\frmLabItems", "�б�Χ", "�б�Χ")
        '-  1062 �ʿ�Ʒ����
        Call UpdateParameterValue(rsSys!ϵͳ, 1062, "����ģ��\zl9CISBase\frmMassResEdit", "����������", "����������")
        
        '1028 ���鼼ʦ����վ
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "�걾��Χ", "�걾��Χ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "�����շ�Χ", "�����շ�Χ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "���μ��鷶Χ", "���μ��鷶Χ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "�걾������ɹ���", "�걾������ɹ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "���μ��鷶Χָ����ʼ����", "���μ��鷶Χָ����ʼ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "�Զ�ˢ��", "�Զ�ˢ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "���պ���ʱ��", "���պ���ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "������ʾ�շ�", "������ʾ�շ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "��������˫��", "��������˫��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "����걾", "����걾")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "��������Ŀ����", "��������Ŀ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "��ʷ����ʶ��", "��ʷ����ʶ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "����Ӧ��ʾ���", "����Ӧ��ʾ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "���ϴ�����ı걾���ۼ�", "���ϴ�����ı걾���ۼ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "ֻ�ں��յǼ�ʱ��ʾ�ǼǴ���", "ֻ�ں��յǼ�ʱ��ʾ�ǼǴ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "�Ǽ�ʱ����Ҫ������Ŀ", "�Ǽ�ʱ����Ҫ������Ŀ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "�ֹ���Ŀ����Ŀ�ۼӱ걾��", "�ֹ���Ŀ����Ŀ�ۼӱ걾��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "���������ļ�", "���������ļ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "�ļ���ȡ����", "�ļ���ȡ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "�ļ���ȡ��Χ", "�ļ���ȡ��Χ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "�ļ���ȡ��ʼ����", "�ļ���ȡ��ʼ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "�ļ���ȡ��������", "�ļ���ȡ��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMain", "��ս�����־", "��ս�����־")
        
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabFilter", "ʹ����ϲ�ѯ", "ʹ����ϲ�ѯ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabFilter", "��ϲ�ѯ", "��ϲ�ѯ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabFilter", "�Ƿ�ʹ��ʱ��", "�Ƿ�ʹ��ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabMain", "ȱʡ����ID", "ȱʡ����ID")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabMain", "��������", "��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabMain", "����С��", "����С��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabMain", "��ʾ������", "��ʾ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabMain", "���ؼ���ͼ��", "���ؼ���ͼ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabMain", "ͼ����", "ͼ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork", "��ʾ���鱸ע", "��ʾ���鱸ע")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork", "ʹ������ɨ��", "ʹ������ɨ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork", "��������", "��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabTrack", "����������", "����������")
        '--frmAddPatient
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmAddPatient", "ѡ�����", "frmAddPatient_ѡ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmAddPatient", "ѡ������", "frmAddPatient_ѡ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmAddPatient", "ѡ�����", "frmAddPatient_ѡ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmAddPatient", "��������", "frmAddPatient_��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmAddPatient", "����ҽ��", "frmAddPatient_����ҽ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmAddPatient", "ִ�п���", "frmAddPatient_ִ�п���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmAddPatient", "��������", "frmAddPatient_��������")
        '--frmBatchAction
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmBatchAction", "����ӡ���ϲ��걾", "frmBatchAction_����ӡ���ϲ��걾")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmBatchAction", "ͬһ�����˺ϲ�Ϊһ�����浥��ӡ", "frmBatchAction_ͬһ�����˺ϲ�Ϊһ�����浥��ӡ")
        '--frmLabAuditingLand
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabAuditingLand", "ʱ��", "frmLabAuditingLand_ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabAuditingLand", "ʱ��", "frmLabAuditingLand_ʱ��")
        '--frmLabBarCodeBatPrint
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabBarCodeBatPrint", "��������Id", "frmLabBarCodeBatPrint_��������Id")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabBarCodeBatPrint", "�걾ID", "frmLabBarCodeBatPrint_�걾ID")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabBarCodeBatPrint", "�ɼ�����", "frmLabBarCodeBatPrint_�ɼ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabBarCodeBatPrint", "ִ��״̬", "frmLabBarCodeBatPrint_ִ��״̬")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "\zl9LisWork\frmLabBarCodeBatPrint", "�Ƿ���Ϊ���", "frmLabBarCodeBatPrint_�Ƿ���Ϊ���")
        '--frmLabMainFindRePort
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMainFindRePort", "ʹ��ʱ�䷶Χ", "frmLabMainFindRePort_ʹ��ʱ�䷶Χ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLabMainFindRePort", "rptFind", "frmLabMainFindRePort_rptFind")
        '--frmLabMainSizer
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "zl9LisWork\������", "���ﲡ��", "������_���ﲡ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "zl9LisWork\������", "סԺ����", "������_סԺ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "zl9LisWork\������", "�����걾", "������_�����걾")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "zl9LisWork\������", "����걾", "������_����걾")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "zl9LisWork\������", "δ��걾", "������_δ��걾")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "zl9LisWork\������", "��첡��", "������_��첡��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "zl9LisWork\������", "����ҽ��", "������_����ҽ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "zl9LisWork\������", "�����걾", "������_�����걾")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "zl9LisWork\������", "���ﲡ��", "������_���ﲡ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\" & mstr�û��� & "zl9LisWork\������", "סԺ����", "������_סԺ����")
        '--frmLabMB
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "����ID", "frmLabMB_����ID")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "���Զ���", "frmLabMB_���Զ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "ͨѶ��", "frmLabMB_ͨѶ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "������", "frmLabMB_������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "����λ", "frmLabMB_����λ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "ֹͣλ", "frmLabMB_ֹͣλ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "У��λ", "frmLabMB_У��λ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "����", "frmLabMB_����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "�ο�����", "frmLabMB_�ο�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "���Ƶ��", "frmLabMB_���Ƶ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "���ʱ��", "frmLabMB_���ʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "���巽ʽ", "frmLabMB_���巽ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "�հ���ʽ", "frmLabMB_�հ���ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "��ĿID", "frmLabMB_��ĿID")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "���հ׶���", "frmLabMB_���հ׶���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "˽��ģ��\zl9LisWork\frmLabMB", "���Զ���", "frmLabMB_���Զ���")
        '--frmLisStationWrite
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLisStationWrite", "�鿴ԭʼ���", "frmLisStationWrite_�鿴ԭʼ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLisStationWrite", "�鿴�ϴν��", "frmLisStationWrite_�鿴�ϴν��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLisStationWrite", "�鿴��־", "frmLisStationWrite_�鿴��־")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLisStationWrite", "�鿴��λ", "frmLisStationWrite_�鿴��λ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLisStationWrite", "�鿴�ο�", "frmLisStationWrite_�鿴�ο�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLisStationWrite", "�鿴ø��", "frmLisStationWrite_�鿴ø��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLisStationWrite", "�鿴����", "frmLisStationWrite_�鿴����")
        '--frmLisStationWrite2
        Call UpdateParameterValue(rsSys!ϵͳ, 1208, "����ģ��\zl9LisWork\frmLisStationWrite2", "�鿴�ϴν��", "frmLisStationWrite2_�鿴�ϴν��")
        
        '--1211
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "˽��ģ��\" & mstr�û��� & "\frmLabSamplingFilter", "�ɼ�����վ����", "�ɼ�����վ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "����ģ��\zl9LisWork\frmLabSampling", "����", "����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "����ģ��\zl9LisWork\frmLabSampling", "����������ӡ", "����������ӡ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "����ģ��\zl9LisWork\frmLabSampling", "���ɺ���Ϊ�����", "���ɺ���Ϊ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "����ģ��\zl9LisWork\frmLabSampling", "����ɺ��ӡ��ִ��", "����ɺ��ӡ��ִ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "����ģ��\zl9LisWork\frmLabSampling", "��������", "��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "����ģ��\zl9LisWork\frmLabSampling", "���Ҳ��˺����ƶ�", "���Ҳ��˺����ƶ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "˽��ģ��\" & mstr�û��� & "\frmLabSamplingRegister", "�ɼ�����վ�Ǽ�", "�ɼ�����վ�Ǽ�")
        '--1212
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "˽��ģ��\" & mstr�û��� & "\frmLabSampleRegister", "�Ƿ񰴲�����ʾ", "�Ƿ񰴲�����ʾ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "˽��ģ��\" & mstr�û��� & "\frmLabSampleRegisterFilter", "�걾�Ǽǹ���", "�걾�Ǽǹ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1211, "˽��ģ��\" & mstr�û��� & "\frmLabSampleRegister", "����", "����")
        
        '-- 1209 ��ʷ�ʿز�ѯ
        Call UpdateParameterValue(rsSys!ϵͳ, 1209, "����ģ��\zl9LisWork\frmQCHistory", "�����ʿ�Ʒ", "�����ʿ�Ʒ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1209, "����ģ��\zl9LisWork\frmQCHistory", "��ʾ����ʧ����Ŀ", "��ʾ����ʧ����Ŀ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1209, "����ģ��\zl9LisWork\frmQCHistory", "����", "����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1209, "����ģ��\zl9LisWork\frmQCHistory", "����", "����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1209, "����ģ��\zl9LisWork\frmQCHistory", "��Ŀ", "��Ŀ")
        
        '--������Ժ����
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��ʾ��ס����", "��ʾ��ס����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient\frmManageHosReg", "ˢ�·�ʽ", "ˢ�·�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient\frmManageHosReg", "��ʾ���˷�ʽ", "��ʾ���˷�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient\frmHosReg", "����ģʽ", "����ģʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��ǰԤ��Ʊ�ݺ�", "��ǰԤ��Ʊ�ݺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "����", "����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "����", "����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "ѧ��", "ѧ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "����״��", "����״��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "ְҵ", "ְҵ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "���", "���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��������", "��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "���֤��", "���֤��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "�����ص�", "�����ص�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��ͥ��ַ", "��ͥ��ַ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "�����ʱ�", "�����ʱ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��ͥ�绰", "��ͥ�绰")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��ϵ������", "��ϵ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��ϵ�˹�ϵ", "��ϵ�˹�ϵ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��ϵ�˵�ַ", "��ϵ�˵�ַ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��ϵ�˵绰", "��ϵ�˵绰")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "������λ", "������λ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��λ�绰", "��λ�绰")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��λ�ʱ�", "��λ�ʱ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��λ������", "��λ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "��λ�ʺ�", "��λ�ʺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "����ģ��\zl9InPatient", "���þ��￨����", "���þ��￨����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1131, "����ģ��\zl9InPatient", "����Ԥ��Ʊ������", "����Ԥ��Ʊ������")

        '--�����������
        Call UpdateParameterValue(rsSys!ϵͳ, 1132, "˽��ģ��\" & mstr�û��� & "\zl9InPatient", "������Ժ", "������Ժ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1132, "����ģ��\zl9InPatient", "����Ʋ��˿���", "����Ʋ��˿���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1132, "����ģ��\zl9InPatient", "��Ժ����", "��Ժ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1132, "����ģ��\zl9InPatient", "��Ժ����", "��Ժ����")
        
        '--Ԥ�������
        Call UpdateParameterValue(rsSys!ϵͳ, 1103, "˽��ģ��\" & mstr�û��� & "\zl9Patient", "��ǰԤ��Ʊ�ݺ�", "��ǰԤ��Ʊ�ݺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1103, "����ģ��\zl9Patient", "����Ԥ��Ʊ������", "����Ԥ��Ʊ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1103, "����ģ��\zl9Patient", "LED��ʾ��ӭ��Ϣ", "LED��ʾ��ӭ��Ϣ")
        
        '--���￨����
        Call UpdateParameterValue(rsSys!ϵͳ, 1102, "����ģ��\zl9Patient", "���þ��￨����", "���þ��￨����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1102, "����ģ��\zl9Patient", "LED��ʾ��ӭ��Ϣ", "LED��ʾ��ӭ��Ϣ")
        
        '--������Ϣ����
        Call UpdateParameterValue(rsSys!ϵͳ, 1101, "˽��ģ��\" & mstr�û��� & "\zl9Patient", "��ǰԤ��Ʊ�ݺ�", "��ǰԤ��Ʊ�ݺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1101, "˽��ģ��\" & mstr�û��� & "\zl9Patient\frmManagePatient", "��ʾ���˷�ʽ", "��ʾ���˷�ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1101, "˽��ģ��\" & mstr�û��� & "\zl9Patient\frmManagePatient", "��������", "��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1101, "˽��ģ��\" & mstr�û��� & "\zl9Patient\frmPatient", "����ģʽ", "����ģʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1101, "����ģ��\zl9Patient", "���û�Ա������", "���û�Ա������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1101, "����ģ��\zl9Patient", "���þ��￨����", "���þ��￨����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1101, "����ģ��\zl9Patient", "����Ԥ��Ʊ������", "����Ԥ��Ʊ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1101, "����ģ��\zl9Patient", "LED��ʾ��ӭ��Ϣ", "LED��ʾ��ӭ��Ϣ")
        
        '--��Լ��λ����1100
        Call UpdateParameterValue(rsSys!ϵͳ, 1100, "˽��ģ��\" & mstr�û��� & "\zl9Patient\frmUnit\Menu", "mnuViewShowAll״̬", "��ʾ�����¼�", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1100, "˽��ģ��\" & mstr�û��� & "\zl10Patient\frmUnit\Menu", "mnuViewShowStop״̬", "��ʾͣ�õ�λ", True)

        '--���Ӳ�������
        Call UpdateParameterValue(rsSys!ϵͳ, 1070, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\�Զ�����", "AutoSave", "AutoSave", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1070, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\�Զ�����", "UndoLimit", "UndoLimit")
        Call UpdateParameterValue(rsSys!ϵͳ, 1070, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\�Զ�����", "SaveInterval", "aveInterval")
        Call UpdateParameterValue(rsSys!ϵͳ, 1070, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\�Զ�����", "AutoSaveEPR", "AutoSaveEPR", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1070, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\�Զ�����", "SaveIntervalEPR", "SaveIntervalEPR")
        Call UpdateParameterValue(rsSys!ϵͳ, 1070, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\�Զ�����", "AutoPageCount", "AutoPageCount", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1070, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\�Զ�����", "AutoPageNote", "AutoPageNote", True)
        Call UpdateParameterValue(rsSys!ϵͳ, 1070, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��ʷ����", "SharePageCount", "SharePageCount")
        Call UpdateParameterValue(rsSys!ϵͳ, 1070, "˽��ģ��\" & mstr�û��� & "\zlRichEPR\��Ĭ��ӡ", "NoAsk", "NoAsk", True)
        
        '--Ӱ��ҽ������վ
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frm3DSetup", "������ά�ؽ�", "������ά�ؽ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frm3DSetup", "3D����·��", "3D����·��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frm3DSetup", "3D����", "3D����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frm3DSetup", "3D����", "3D����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�����Ǽ�����", "�����Ǽ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�Ǽ�ֱ�Ӽ��", "�Ǽ�ֱ�Ӽ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�������Զ���ӡ���뵥", "�������Զ���ӡ���뵥")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "����ʾ��������", "����ʾ��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "����ʾ��Ӱ�� ", "����ʾ��Ӱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "��ʼ����Զ��򿪱���", "��ʼ����Զ��򿪱���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "����ʱ��Ƭ", "����ʱ��Ƭ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "����ʾ��ȡ���ĵǼ�", "����ʾ��ȡ���ĵǼ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "���˸���", "���˸���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "ֻ����ѡ��ִ�м䲡��", "ֻ����ѡ��ִ�м䲡��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "ִ�м䷶Χ", "ִ�м䷶Χ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�����б��ͷ����", "�����б��ͷ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�����б��ͷ�ֺ�", "�����б��ͷ�ֺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�����б��ͷ����", "�����б��ͷ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�����б��ͷб��", "�����б��ͷб��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�����б���������", "�����б���������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�����б������ֺ�", "�����б������ֺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�����б����ݴ���", "�����б����ݴ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1290, "����ģ��\zl9PACSWork\frmPACSTechnicSetup", "�����б�����б��", "�����б�����б��")

        '--Ӱ��ɼ�����վ
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�����Ǽ�����", "�����Ǽ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�Ǽ�ֱ�Ӽ��", "�Ǽ�ֱ�Ӽ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�������Զ���ӡ���뵥", "�������Զ���ӡ���뵥")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "����ʾ��Ӱ��", "����ʾ��Ӱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "����ʾ��������", "����ʾ��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "��ʼ����Զ��򿪱���", "��ʼ����Զ��򿪱���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "����ʾ��ȡ���ĵǼ�", "����ʾ��ȡ���ĵǼ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "����ʱ��Ƭ", "����ʱ��Ƭ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "���˸���", "���˸���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "ֻ����ѡ��ִ�м䲡��", "ֻ����ѡ��ִ�м䲡��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "ִ�м䷶Χ", "ִ�м䷶Χ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoCapture", "��̤�˿�", "��̤�˿�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoCapture", "��̤�ɼ���ʽ", "��̤�ɼ���ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoCapture", "��̤ʱ����", "��̤ʱ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoCapture", "����ƶ�ʱ��ʾ��ͼ", "����ƶ�ʱ��ʾ��ͼ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoCapture", "�ɼ���ͼ�Ŵ���", "�ɼ���ͼ�Ŵ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�����б��ͷ����", "�����б��ͷ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�����б��ͷ�ֺ�", "�����б��ͷ�ֺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�����б��ͷ����", "�����б��ͷ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�����б��ͷб��", "�����б��ͷб��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�����б���������", "�����б���������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�����б������ֺ�", "�����б������ֺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�����б����ݴ���", "�����б����ݴ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1291, "����ģ��\zl9PACSWork\frmVideoTechnicSetup", "�����б�����б��", "�����б�����б��")
        
        '--Ӱ������վ
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�����Ǽ�����", "�����Ǽ�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�Ǽ�ֱ�Ӽ��", "�Ǽ�ֱ�Ӽ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�������Զ���ӡ���뵥", "�������Զ���ӡ���뵥")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "����ʾ��Ӱ��", "����ʾ��Ӱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "����ʾ��������", "����ʾ��������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "��ʼ����Զ��򿪱���", "��ʼ����Զ��򿪱���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "����ʱ��Ƭ", "����ʱ��Ƭ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "����ʾ��ȡ���ĵǼ�", "����ʾ��ȡ���ĵǼ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "���˸���", "���˸���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "ֻ����ѡ��ִ�м䲡��", "ֻ����ѡ��ִ�м䲡��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "ִ�м䷶Χ", "ִ�м䷶Χ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmVideoCapture", "��̤�˿�", "��̤�˿�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmVideoCapture", "��̤�ɼ���ʽ", "��̤�ɼ���ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmVideoCapture", "��̤ʱ����", "��̤ʱ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmVideoCapture", "����ƶ�ʱ��ʾ��ͼ", "����ƶ�ʱ��ʾ��ͼ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmVideoCapture", "�ɼ���ͼ�Ŵ���", "�ɼ���ͼ�Ŵ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�����б��ͷ����", "�����б��ͷ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�����б��ͷ�ֺ�", "�����б��ͷ�ֺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�����б��ͷ����", "�����б��ͷ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�����б��ͷб��", "�����б��ͷб��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�����б���������", "�����б���������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�����б������ֺ�", "�����б������ֺ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�����б����ݴ���", "�����б����ݴ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1293, "����ģ��\zl9PACSWork\frmPathologyTechnicSetup", "�����б�����б��", "�����б�����б��")
        
        '��������ҩƷĿ¼����
        Call UpdateParameterValue(rsSys!ϵͳ, 1023, "����ģ��\zl9CisBase\ҩƷ����ģʽ", "Ʒ������ģʽ", "Ʒ������ģʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1023, "����ģ��\zl9CisBase\ҩƷ����ģʽ", "�������ģʽ", "�������ģʽ")
        
        'ҩ����ҩ��ҩƷ��ͨ����
        'ҩƷ������ҩ
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "�շѴ�����ʾ��ʽ", "�շѴ�����ʾ��ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "���ʴ�����ʾ��ʽ", "���ʴ�����ʾ��ʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ѯ����", "��ѯ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ӡ�������ʵ�", "��ӡ�������ʵ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ӡ�˷ѵ��ݼ��", "��ӡ�˷ѵ��ݼ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ӡ�ӳ�", "��ӡ�ӳ�")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "ˢ�¼��", "ˢ�¼��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ӡ���", "��ӡ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ʾ����", "��ʾ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ҩ���Զ���ӡ", "��ҩ���Զ���ӡ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "ҩ������", "ҩ������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "�����µ����Ƿ��ӡ", "�����µ����Ƿ��ӡ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ӡָ����ҩ����", "��ӡָ����ҩ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ӡҩƷ��ǩ", "��ӡҩƷ��ǩ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ҩ����", "��ҩ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ҩҩ��", "��ҩҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��ҩ��", "��ҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "��Դ����", "��Դ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "�Զ���ҩ", "�Զ���ҩ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "����ģ��\����\zl9DrugStore\frmҩƷ��ҩ����", "�Զ���ҩʱ��", "�Զ���ҩʱ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1341, "˽��ģ��\" & mstr�û��� & "\zl9DrugStore\frmҩƷ��ҩ����", "��ʾ��С��λ", "��ʾ��С��λ")
        
        'ҩƷ���ŷ�ҩ
        Call UpdateParameterValue(rsSys!ϵͳ, 1342, "����ģ��\����\zl9DrugStore\Frm���ŷ�ҩ����", "��ҩҩ��", "��ҩҩ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 1342, "����ģ��\����\zl9DrugStore\Frm���ŷ�ҩ����", "�Զ���ӡ", "�Զ���ӡ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1342, "˽��ģ��\" & mstr�û��� & "\zl9DrugStore\���ŷ�ҩ����", "��ʾ��С��λ", "��ʾ��С��λ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1342, "˽��ģ��\" & mstr�û��� & "\zl9DrugStore\���ŷ�ҩ����", "�����һ�����ʾ�����嵥", "�����һ�����ʾ�����嵥")
        Call UpdateParameterValue(rsSys!ϵͳ, 1342, "˽��ģ��\" & mstr�û��� & "\zl9DrugStore\���ŷ�ҩ����", "����ģʽ", "����ģʽ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1342, "˽��ģ��\" & mstr�û��� & "\zl9DrugStore\���ŷ�ҩ����", "������", "������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1342, "˽��ģ��\" & mstr�û��� & "\zl9DrugStore\���ŷ�ҩ����", "�������", "�������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1342, "˽��ģ��\" & mstr�û��� & "\zl9DrugStore\���ŷ�ҩ����", "��ֵ����", "��ֵ����")
        
        'ҩƷ����ѯ
        Call UpdateParameterValue(rsSys!ϵͳ, 1309, "˽��ģ��\" & mstr�û��� & "\zl9MediStore\ҩƷ����ѯ", "��λ", "��λ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1309, "˽��ģ��\" & mstr�û��� & "\zl9MediStore\ҩƷ����ѯ", "�Ƿ���ʾ�޿��ҩƷ", "�Ƿ���ʾ�޿��ҩƷ")
        Call UpdateParameterValue(rsSys!ϵͳ, 1309, "˽��ģ��\" & mstr�û��� & "\zl9MediStore\ҩƷ����ѯ", "Ч�ڱ�������", "Ч�ڱ�������")
        Call UpdateParameterValue(rsSys!ϵͳ, 1309, "˽��ģ��\" & mstr�û��� & "\zl9MediStore\ҩƷ����ѯ", "�Ƿ���ʾͣ��ҩƷ", "�Ƿ���ʾͣ��ҩƷ")
        
        '1560-�������
        Call UpdateParameterValue(rsSys!ϵͳ, 1560, "˽��ģ��\" & mstr�û��� & "\zl9CISAudit\frmChildQuestion", "��ǰ����", "��ǰ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1560, "˽��ģ��\" & mstr�û��� & "\zl9CISAudit\frmEPRAuditMan", "��ʾ��ҵ�����", "��ʾ��ҵ�����")
        
        '1561-��������
        Call UpdateParameterValue(rsSys!ϵͳ, 1561, "˽��ģ��\" & mstr�û��� & "\zl9CISAudit\frmSearchPatient", "��������", "��������")
        
        '1562-��������
        Call UpdateParameterValue(rsSys!ϵͳ, 1562, "˽��ģ��\zl9CISAudit\frm��������", "δ����", "δ����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1562, "˽��ģ��\zl9CISAudit\frm��������", "δ���", "δ���")
        Call UpdateParameterValue(rsSys!ϵͳ, 1562, "˽��ģ��\zl9CISAudit\frm��������", "�����", "�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 1562, "˽��ģ��\zl9CISAudit\frm��������", "��λ��Χ", "��λ��Χ")
        
    End If
    
    '��������ϵͳ�Ĳ���ֵ����
    '-----------------------------------------------------------------
    rsSys.Filter = "ϵͳ=4": rsUpgrade.Filter = "ϵͳ=4"
    If Not rsSys.EOF And rsUpgrade.EOF Then
        '��������ϵͳ�����ģ��
        '��������Ŀ¼����
        Call UpdateParameterValue(rsSys!ϵͳ, 603, "˽��ģ��\" & mstr�û��� & "\ZL9Material\����\����Ŀ¼����\��Ƭ", "��Ƭ", "��ʾ���ؿ�Ƭ")
        Call UpdateParameterValue(rsSys!ϵͳ, 603, "˽��ģ��\" & mstr�û��� & "\ZL9Material\����\����Ŀ¼����", "�����¼�����", "�����¼�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 603, "˽��ģ��\" & mstr�û��� & "\ZL9Material\����\����Ŀ¼����", "����ͣ������", "����ͣ������")
        
        '�⹺���
        Call UpdateParameterValue(rsSys!ϵͳ, 309, "˽��ģ��\" & mstr�û��� & "\ZL9Material\�����⹺��ⵥ\BillEdit", "mshBill���", "�����п�")
        Call UpdateParameterValue(rsSys!ϵͳ, 309, "˽��ģ��\" & mstr�û��� & "\ZL9Material\�����⹺��ⵥ\BillEdit", "mshBill����", "������ͷ�ı�")
        '���ù���
        Call UpdateParameterValue(rsSys!ϵͳ, 312, "˽��ȫ��\" & mstr�û��� & "\�������õ�\mshBill", "��˱�־", "��˱�־�п�")
    End If
    
    '�����豸ϵͳ�Ĳ���ֵ����
    '-----------------------------------------------------------------
    rsSys.Filter = "ϵͳ=6": rsUpgrade.Filter = "ϵͳ=6"
    If Not rsSys.EOF And rsUpgrade.EOF Then
        '�����豸ϵͳ�����ģ��
        '�豸Ŀ¼����
        Call UpdateParameterValue(rsSys!ϵͳ, 602, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸Ŀ¼����\��Ƭ", "��ʾ��Ƭ", "��ʾ���ؿ�Ƭ")
        Call UpdateParameterValue(rsSys!ϵͳ, 602, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸Ŀ¼����", "��ʾͣ��", "����ͣ���豸")
        
        '�豸ʹ��״̬����
        Call UpdateParameterValue(rsSys!ϵͳ, 603, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸ʹ��״̬����", "��ʾ��Ƭ", "��ʾ���ؿ�Ƭ")
        '�豸�⹺������
        Call UpdateParameterValue(rsSys!ϵͳ, 616, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸��ѡ��-616", "�豸ѡ����", "�豸��Ϣ�б�")
        Call UpdateParameterValue(rsSys!ϵͳ, 616, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸��ѡ��-616", "����ѡ����", "������Ϣ�б�")
        '�豸��������
        Call UpdateParameterValue(rsSys!ϵͳ, 618, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸���ε�", "��ʾ��Ƭ", "��ʾ��Ƭ��Ϣ")
        '�豸���ù���
        Call UpdateParameterValue(rsSys!ϵͳ, 619, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸���õ�", "��ʾ��Ƭ", "��ʾ��Ƭ��Ϣ")
        '�豸ʹ�ù���
        Call UpdateParameterValue(rsSys!ϵͳ, 624, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸ʹ�ù���", "���������豸", "���������豸")
        '�豸��������
        Call UpdateParameterValue(rsSys!ϵͳ, 625, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸��������", "���������豸", "���������豸")
        '�豸������
        Call UpdateParameterValue(rsSys!ϵͳ, 626, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸������", "���������豸", "���������豸")
        '�豸ά�޹���
        Call UpdateParameterValue(rsSys!ϵͳ, 627, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸ά�޹���", "���������豸", "���������豸")
        '�豸�䶯����
        Call UpdateParameterValue(rsSys!ϵͳ, 628, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸�䶯����", "���������豸", "���������豸")
        '�豸���ʹ���
        Call UpdateParameterValue(rsSys!ϵͳ, 631, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸���ʹ���", "���������豸", "���������豸")
        '�����豸��ó��
        Call UpdateParameterValue(rsSys!ϵͳ, 636, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸���ò�ѯ", "����ID", "�ϴ��豸����ID")
        Call UpdateParameterValue(rsSys!ϵͳ, 636, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸���ò�ѯ", "����ID", "�ϴβ���ID")
        Call UpdateParameterValue(rsSys!ϵͳ, 636, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸���ò�ѯ", "�����¼�����", "�����¼�����")
        Call UpdateParameterValue(rsSys!ϵͳ, 636, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸���ò�ѯ", "�������ż��豸", "ͬʱ���Ʋ��ż��豸����")
        
        '�豸����ѯ
        Call UpdateParameterValue(rsSys!ϵͳ, 637, "˽��ģ��\" & mstr�û��� & "\ZL9Device\Frm�豸����ѯ\Menu", "��ʾ�������", "����ʾ����豸")
        '�豸��ϸ��
        Call UpdateParameterValue(rsSys!ϵͳ, 650, "˽��ģ��\" & mstr�û��� & "\ZL9Device\�豸��ϸ��", "ֻ��ʾ�����ݵĲ���", "ֻ��ʾ�����ݵĲ���")
    End If
    
    
    '������ϵͳ�Ĳ���ֵ����
    '-----------------------------------------------------------------
    rsSys.Filter = "ϵͳ=23": rsUpgrade.Filter = "ϵͳ=23"
    If Not rsSys.EOF And rsUpgrade.EOF Then
        
    End If
    
    '����Ժ��ϵͳ�Ĳ���ֵ����
    '-----------------------------------------------------------------
    If Not rsSys.EOF And rsUpgrade.EOF Then
        Call UpdateParameterValue(rsSys!ϵͳ, 2301, "˽��ģ��\" & mstr�û��� & "\ZL9Device\zl9Infect\��ˮ��Ŀ����", "��ʾͣ��", "��ˮ��Ŀ-��ʾͣ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 2301, "˽��ģ��\" & mstr�û��� & "\ZL9Device\zl9Infect\��ˮ��Ŀ����", "��ʾ�¼�", "��ˮ��Ŀ-��ʾ�¼�")
        Call UpdateParameterValue(rsSys!ϵͳ, 2301, "˽��ģ��\" & mstr�û��� & "\ZL9Device\zl9Infect\������Ŀ����", "��ʾͣ��", "������Ŀ-��ʾͣ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 2301, "˽��ģ��\" & mstr�û��� & "\ZL9Device\zl9Infect\������Ŀ����", "��ʾ�¼�", "������Ŀ-��ʾ�¼�")
        Call UpdateParameterValue(rsSys!ϵͳ, 2301, "˽��ģ��\" & mstr�û��� & "\ZL9Device\zl9Infect\�׸����ع���", "��ʾͣ��", "�׸�����-��ʾͣ��")
        Call UpdateParameterValue(rsSys!ϵͳ, 2301, "˽��ģ��\" & mstr�û��� & "\ZL9Device\zl9Infect\�׸����ع���", "��ʾ�¼�", "�׸�����-��ʾ�¼�")
    End If
    Exit Sub
errH:
    If zlComLib.ErrCenter() = 1 Then Resume
    Call zlComLib.SaveErrLog
End Sub

Private Sub UpdateParameterValue(ByVal intϵͳ As Integer, ByVal intģ�� As Integer, _
    ByVal strPath As String, ByVal strPreName As String, ByVal strNowName As String, Optional ByVal blnTransBool As Boolean)
'���ܣ����¾���ĳ��ע��������ֵ
'������intϵͳ=ϵͳ��,��Ϊ��ǰע�������ǲ������״洢�ģ��������ʱֻ�����׼��ϵͳ��
'      intģ��=ģ���
'      strPath=ԭע�����·��
'      strPreName=ԭע����ŵĲ�����(ע������)
'      strNowName=�µĴ�ŵ����ݿ��еı���������
'      blnTransBool=�Ƿ�ע�����ΪTrue��False��ֵתΪ1��0�浽���ݿ���
    Dim strVal As String, strSQL As String
    
    On Error GoTo errH
    
    strVal = GetSetting("ZLSOFT", strPath, strPreName)
    If strVal = "" Then Exit Sub 'û��ֵʱҲ�����޴˼�����
    If blnTransBool Then
        If UCase(strVal) = "TRUE" Then
            strVal = "1"
        ElseIf UCase(strVal) = "FALSE" Then
            strVal = "0"
        End If
    End If
    
    strSQL = "zl_Parameters_Update('" & strNowName & "','" & Replace(strVal, "'", "''") & "'," & intϵͳ * 100 & "," & intģ�� & ")"
    zlDataBase.ExecuteProcedure strSQL, "UpdateParameterValue"
    
    DeleteSetting "ZLSOFT", strPath, strPreName '������ʱɾ�������
    Exit Sub
errH:
    If zlComLib.ErrCenter = 1 Then Resume
    Call zlComLib.SaveErrLog
End Sub
