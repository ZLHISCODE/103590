Attribute VB_Name = "mdlDrugPacker"
Option Explicit

Public Function DrugInfo(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'���ܣ��ϴ�ҩƷ������Ϣ
'������
'   objDevice���豸����
'   strContent���Ѹ�ʽ��������
'���أ�True�ɹ���Falseʧ��
    
    If objDevice Is Nothing Then Exit Function
    If objDevice.Status = False Then Exit Function

    On Error GoTo errHandle
    Select Case objDevice.LinkType
    Case enuLinkType.DB
        'DB�������͵������ϴ�������ô洢���̡�ֱ�ӽ������ݵ�
        DrugInfo = True
    
    Case enuLinkType.WEBServices
        'WebServices���͵������ϴ���ͨ������WebServices���ŵĽӿں����ϴ����ݡ�
        'objDevice.WSConnect.?????
        DrugInfo = True
        
    Case enuLinkType.Directory
        '�ļ����͵������ϴ���
        DrugInfo = True
        
    End Select
    Exit Function
    
errHandle:
    gstrMessage = Err.Description
    gobjComLib.ErrCenter
End Function

Public Function DrugStock(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'���ܣ��ϴ�ҩƷ�����Ϣ
'������
'   objDevice���豸����
'   strContent���Ѹ�ʽ��������
'���أ�True�ɹ���Falseʧ��

    If objDevice Is Nothing Then Exit Function
    If objDevice.Status = False Then Exit Function
    
    Select Case objDevice.LinkType
    Case enuLinkType.DB
        'DB�������͵������ϴ�������ô洢���̡�ֱ�ӽ������ݵ�
        DrugStock = True
        
    Case enuLinkType.WEBServices
        'WebServices���͵������ϴ���ͨ������WebServices���ŵĽӿں����ϴ����ݡ�
        DrugStock = True
        
    Case enuLinkType.Directory
        '�ļ����͵������ϴ���
        DrugStock = True
    
    End Select
    
errHandle:
    gstrMessage = Err.Description
    gobjComLib.ErrCenter
End Function

Public Function Dispense(ByVal objDevice As clsDevice, ByVal strNO As String, ByVal int���� As Integer, ByVal strContent As String) As Boolean
'���ܣ����Զ���ϵͳ������ҩ��Ϣ
'������
'   objDevice���豸����
'   strNO�����ݺ�
'   strContent���Ѹ�ʽ��������
'���أ�True�ɹ���Falseʧ��

    If objDevice Is Nothing Then Exit Function
    If objDevice.Status = False Then Exit Function

    Select Case objDevice.LinkType
    Case enuLinkType.DB
        'DB�������͵������ϴ�������ô洢���̡�ֱ�ӽ������ݵ�
        Dispense = True
        
    Case enuLinkType.WEBServices
        'WebServices���͵������ϴ���ͨ������WebServices���ŵĽӿں����ϴ����ݡ�
        Dispense = True
        
    Case enuLinkType.Directory
        '�ļ����͵������ϴ���
        Dispense = True
    
    End Select
    
    '������ҩ����
    'If SetSendWin(ҩ��ID, ���ݺ�, ����, ��ҩ����) = False Then gstrMessage = "���������ķ�ҩ����ʧ�ܣ�"
    
errHandle:
    gstrMessage = Err.Description
End Function

Public Function Dispensing(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'���ܣ����Զ���ϵͳ���ͷ�ҩ��Ϣ
'������
'   objDevice���豸����
'   strContent���Ѹ�ʽ��������
'���أ�True�ɹ���Falseʧ��

    If objDevice Is Nothing Then Exit Function
    If objDevice.Status = False Then Exit Function
    
    On Error GoTo errHandle
    Select Case objDevice.LinkType
    Case enuLinkType.DB
        'DB�������͵������ϴ�������ô洢���̡�ֱ�ӽ������ݵ�
        Dispensing = True
    
    Case enuLinkType.WEBServices
        'WebServices���͵������ϴ���ͨ������WebServices���ŵĽӿں����ϴ����ݡ�
        Dispensing = True
        
    Case enuLinkType.Directory
        '�ļ����͵������ϴ���
        Dispensing = True
        
    End Select
    
    Exit Function
    
errHandle:
    gstrMessage = Err.Description
End Function
