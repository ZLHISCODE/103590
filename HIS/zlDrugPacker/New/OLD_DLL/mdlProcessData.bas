Attribute VB_Name = "mdlProcessData"
Option Explicit

Public Sub ProcDrugInfo(ByVal strDrugType As String, ByVal strLinkName As String)
'���ܣ���ȡHISҩƷ������Ϣ
'������
'  strDrugType�����ʹ�
'  strLinkName����������
    
    'ʵ��clsConnect
    
    '��HIS����
    
    '���������Ͳ�ͬ���ֱ��ʽ��Ҫ�ϴ�������
    
    '����mdlDrugPacker.DrugInfo
    
    
    Exit Sub

errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub



Public Sub SetUpload(ByVal bytType As Byte, ByVal varKey As Variant)
'���ܣ���ȡHIS����ϴ���Ϣ
'������
'   bytType��
'       1: ���ﴦ���ϴ� (��ҩ)
'       2: ���﷢ҩ֪ͨ (��ҩ)
'       3: סԺҩƷҽ���ϴ� (�䡢��ҩ)
'   varKey��
'       ��bytType=1ʱ��varKey��ʾ������;�ⷿID;NO����
'       ��ʽ��������;�ⷿID;NO[|����;�ⷿID;NO][|...]��
'       ��bytType=2ʱ��ͬbytType=1
'       ��bytType=3ʱ��varKey��ʾҩƷ�շ�ID��
'       ��ʽ����ҩƷ�շ�ID[|ҩƷ�շ�ID][|...]��

    '��HIS����
    
    '��¼��ȷ��Ҫ�ϴ��豸����
    
    '��ʽ��Ҫ�ϴ�������
    
    '����mdlDrugPacker.Dispense��Dispensing


End Sub

