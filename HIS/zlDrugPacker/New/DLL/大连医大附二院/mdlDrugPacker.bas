Attribute VB_Name = "mdlDrugPacker"
Option Explicit

Public Function DrugInfo(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'���ܣ��ϴ�ҩƷ������Ϣ
'������
'   objDevice���豸����
'   strContent���Ѹ�ʽ��������
'���أ�True�ɹ���Falseʧ��
    
    Dim cmOutside As New ADODB.Command
    Dim strResume As String, strIP As String
    
    strIP = GetLocalIP
    
    On Error GoTo errHandle
    Select Case objDevice.LinkType
    Case enuLinkType.DB
        On Error GoTo errDB
        objDevice.DBConnect.BeginTrans
        Set cmOutside.ActiveConnection = objDevice.DBConnect
        cmOutside.CommandText = strContent
        cmOutside.Execute
        objDevice.DBConnect.CommitTrans
        DrugInfo = True
    
    Case enuLinkType.WEBServices
        'TransConsiData
        '����1������ID
        '����2��ҵ������
        '����3��XML�ı�
        '����4������IP
        '����5��HIS�û����
        '����6��HIS�û�����
        '����7��������Ϣ
        If objDevice.WSConnect.TransConsisData(1, 101, strContent, strIP, gstrUserCode, gstrUserName, strResume) <> 1 Then
            gstrMessage = strResume
        Else
            DrugInfo = True
        End If
        
    Case enuLinkType.Directory
        
        
    End Select
    Exit Function
    
errDB:
    objDevice.DBConnect.RollbackTrans
    Exit Function
    
errHandle:
    gobjComLib.ErrCenter
End Function

Public Function DrugStock(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'���ܣ��ϴ�ҩƷ�����Ϣ
'������
'   objDevice���豸����
'   strContent���Ѹ�ʽ��������
'���أ�True�ɹ���Falseʧ��
    
    Dim strIP As String, strResume As String
    Dim intRetval As Integer
    
    strIP = GetLocalIP
    
    Select Case objDevice.LinkType
    Case enuLinkType.DB
        
    Case enuLinkType.WEBServices
        If objDevice.WSConnect.TransConsisData(1, 102, strContent, strIP, gstrUserCode, gstrUserName, intRetval, strResume) <> 1 Then
            gstrMessage = strResume
        Else
            DrugStock = True
        End If
        
    Case enuLinkType.Directory
    
    End Select
    
End Function

Public Function Dispense(ByVal objDevice As clsDevice, ByVal strNO As String, ByVal int���� As Integer, ByVal strContent As String) As Boolean
'���ܣ����Զ���ϵͳ������ҩ��Ϣ
'������
'   objDevice���豸����
'   strNO�����ݺ�
'   strContent���Ѹ�ʽ��������
'���أ�True�ɹ���Falseʧ��

    Dim strIP As String, strResume As String
    Dim intRetval As Integer
    
    strIP = GetLocalIP
    
    If objDevice.Status Then
        
        On Error GoTo errHandle
    
        Select Case objDevice.LinkType
        Case enuLinkType.DB
            '
        Case enuLinkType.WEBServices
            If objDevice.WSConnect.TransConsisData(1, 201, strContent, strIP, gstrUserCode, gstrUserName, intRetval, strResume) <> 1 Then
                gstrMessage = strResume
            Else
                Dispense = True
            End If
            
            '��ҩ����
            If SetSendWin(objDevice.DeptID, strNO, int����, intRetval) = False Then gstrMessage = "���������ķ�ҩ����ʧ�ܣ�"
            
            
        Case enuLinkType.Directory
        End Select
        
    End If
    
    Exit Function

errHandle:
    gstrMessage = Err.Description
End Function

Public Function Dispensing(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'���ܣ����Զ���ϵͳ���ͷ�ҩ��Ϣ
'������
'   objDevice���豸����
'   strContent���Ѹ�ʽ��������
'���أ�True�ɹ���Falseʧ��

    Dim strIP As String, strResume As String
    Dim cmOutside As New ADODB.Command
    Dim intRetval As Integer
    
    strIP = GetLocalIP
    
    On Error GoTo errHandle
    If objDevice.Status Then
        Select Case objDevice.LinkType
        Case enuLinkType.DB
            Set cmOutside.ActiveConnection = objDevice.DBConnect
            cmOutside.CommandText = strContent
            cmOutside.Execute
            Dispensing = True
        
        Case enuLinkType.WEBServices
            If objDevice.WSConnect.TransConsisData(1, 202, strContent, strIP, gstrUserCode, gstrUserName, intRetval, strResume) <> 1 Then
                gstrMessage = strResume
            Else
                Dispensing = True
            End If
            
        Case enuLinkType.Directory
        End Select
    End If
    Exit Function
    
errHandle:
    gstrMessage = Err.Description
End Function
