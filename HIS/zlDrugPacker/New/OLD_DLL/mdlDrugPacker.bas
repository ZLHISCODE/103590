Attribute VB_Name = "mdlDrugPacker"
Option Explicit

Public Function HIS2Auto_DrugInfo(ByVal objConn As clsConnect, ByVal strContent As String) As String
'���ܣ��ϴ�ҩƷ������Ϣ
'������
'   objConnect�����Ӷ���
'   strContent���Ѹ�ʽ��������
'���أ���Ϣ����
    
    
End Function

Public Function HIS2Auto_Dispense(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'���ܣ����Զ���ϵͳ������ҩ��Ϣ
'������
'   objDevice���豸����
'   strContent���Ѹ�ʽ��������
'���أ�True�ɹ���Falseʧ��

    
End Function

Public Function HIS2Auto_Dispensing(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'���ܣ����Զ���ϵͳ���ͷ�ҩ��Ϣ
'������
'   objDevice���豸����
'   strContent���Ѹ�ʽ��������
'���أ�True�ɹ���Falseʧ��

    HIS2Auto_Dispensing = True

End Function
