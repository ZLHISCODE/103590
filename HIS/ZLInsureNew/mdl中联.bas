Attribute VB_Name = "mdl����"
Option Explicit

Public Function ҽ����ʼ��_����(ByVal int���� As Integer) As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false

    
    'Ϊ�˱�����Ȩ�Ѷ����ӣ��˴����ٽ��жԸ���ҽ�������ݵļ��
    ҽ����ʼ��_���� = True
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long, Optional ByVal int���� As Integer) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    
    Dim strTmpIden As String
    
    strTmpIden = frmIdentify����.ShowCard(bytType, lng����ID, int����)
    ��ݱ�ʶ_���� = strTmpIden
End Function

Public Function �������_����(ByVal lng����ID As Long, ByVal intinsure As Integer) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: bytYear-�������,0-�������,1-�������,2-�������
'����: ���ظ����ʻ����Ľ��
    Dim rsTemp As New ADODB.Recordset

    
    gstrSQL = "select A.�ʻ���� from �����ʻ� A where A.����ID=[1] and A.����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", lng����ID, intinsure)
    
    If rsTemp.EOF Then
        �������_���� = 0
    Else
        �������_���� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
    End If

End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, _
            ByVal curȫ�Է� As Currency, ByVal cur�����Ը� As Currency, ByVal intinsure As Integer, Optional ByRef strAdvance As String = "") As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    Dim curƱ���ܽ�� As Currency
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curDate As Date
On Error GoTo ErrH
'    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
'    gstrSQL = "Select ����ID,���ʽ��  From ���˷��ü�¼ Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9"
'    Call OpenRecordset(rsTemp, "ģ��ҽ��")
'
'    Do Until rsTemp.EOF
'        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
'
'        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
'        rsTemp.MoveNext
'    Loop
    
    '---------------------------------------------------------------------------------------------
    '��д�����
    curDate = zlDatabase.Currentdate
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(intinsure, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    '����:
    '   ����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN
    '   �ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,�������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,֧��˳���_IN,��ҳID_IN,
    '   ��;����_IN,��ע_IN
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & g��������.�������ý�� & "," & g��������.ȫ�Էѽ�� & "," & g��������.�����Ը���� & "," & _
        g��������.����ͳ���� & "," & g��������.ͳ�ﱨ����� & ",0,0," & cur�����ʻ� & "," & g��������.�Żݽ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    �������_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function


Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, ByVal intinsure As Integer, Optional ByRef strAdvance As String = "") As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset
    Dim rs�˷� As New ADODB.Recordset
    Dim lng����ID As Long
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curƱ���ܽ�� As Currency, curȫ�Է� As Currency, cur�����Ը� As Currency, cur����ͳ�� As Currency
    Dim curDate As Date
On Error GoTo ErrH
        
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,�������ý��,ȫ�Ը����,�����Ը����,����ͳ����  From ���ս����¼ Where ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", lng����ID)
        
    lng����ID = rsTemp("����ID")
        
    curƱ���ܽ�� = IIf(IsNull(rsTemp("�������ý��")), 0, rsTemp("�������ý��")) * -1
    curȫ�Է� = IIf(IsNull(rsTemp("ȫ�Ը����")), 0, rsTemp("ȫ�Ը����")) * -1
    cur�����Ը� = IIf(IsNull(rsTemp("�����Ը����")), 0, rsTemp("�����Ը����")) * -1
    cur����ͳ�� = IIf(IsNull(rsTemp("����ͳ����")), 0, rsTemp("����ͳ����")) * -1
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rs�˷� = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", lng����ID)
    
    lng����ID = rs�˷�("����ID")
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(intinsure, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� & "," & curȫ�Է� & "," & cur�����Ը� & "," & _
        cur����ͳ�� & ",0,0,0," & cur�����ʻ� * -1 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")

    ����������_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function �����ʻ�תԤ��_����(lngԤ��ID As Long, cur�����ʻ� As Currency, lng����ID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curDate As Date
    
    '---------------------------------------------------------------------------------------------
    '��д�����
    curDate = zlDatabase.Currentdate
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(intinsure, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(3," & lngԤ��ID & "," & intinsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & cur�����ʻ� & ",0,0,0,0,0,0," & _
        cur�����ʻ� & ",0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    �����ʻ�תԤ��_���� = True
End Function


Public Function �����ʻ�תԤ������_����(lngԤ��ID As Long, cur�����ʻ� As Currency, lng����ID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    Dim rs�˷� As New ADODB.Recordset
    Dim lng����ID As Long
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curƱ���ܽ�� As Currency
    Dim curDate As Date
        
        
    curDate = zlDatabase.Currentdate
    
    '�˷�
    gstrSQL = "select distinct A.ID from ����Ԥ����¼ A,����Ԥ����¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", lngԤ��ID)
    
    lng����ID = rsTemp("ID")
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(intinsure, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & cur�����ʻ� * -1 & ",0,0,0,0,0,0," & _
        cur�����ʻ� * -1 & ",0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")

    �����ʻ�תԤ������_���� = True
    
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false

    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    ��Ժ�Ǽ�_���� = True
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    ��Ժ�Ǽ�_���� = True
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
            'ȡ��Ժ�Ǽ���֤�����ص�˳���
    On Error GoTo errHandle
    
        
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rs������ϸ As Recordset, ByVal intinsure As Integer) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����Ҫ��NO����š�����ID��ҽ����Ŀ���롢�շ�����շ����ơ��������š���񡢲��ء��������۸񡢽�ҽ��,�Ǽ�ʱ��(����ʱ��),Ӥ����,���մ���ID
    Dim rs������� As Recordset     '��ҽ��֧��������ܵõ�
    Dim rs�㷨 As New ADODB.Recordset          '����
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng���� As Long
    Dim lng��ְ As Long, lng����� As Long, lng���� As Long
    Dim dblTemp As Double, lng���� As Long
    
    Dim dbl�����  As Double ''��һ����סԺ�ռ������Ŀ������ܵõ��Ľ��
    Dim dbl�ѱ������ As Double, dbl�ۼƽ��� As Double
    Dim dbl���� As Double, dbl���� As Double, dbl�ֶν��� As Double, dbl�ֶα��� As Double
    
    Dim clsҽ�� As New clsInsure
    Dim bln�����ʻ�֧��ȫ�Է� As Boolean, bln�����ʻ�֧�������Ը� As Boolean, bln�����ʻ�֧������ As Boolean
    Dim curȫ�Է� As Currency, cur�����Ը� As Currency
    Dim blnȫ��ͳ�� As Boolean, bln������ As Boolean, bln�޷ⶥ�� As Boolean
    
    Dim bln������� As Boolean   '�����Թ�ҽ��������ǿ�����㣬��ʹ�ò����ǵڶ��ν��ʡ����ֶμ���Ҳ�Ǵ�ͷ��ʼ
    Dim dbl������ߺ� As Double, dbl��ν���ͳ��� As Double   '�����ָ�ò�����ǰ���ʵ��ۼ�
    Dim dbl�������� As Double, dbl�������� As Double
    Dim dblTemp1 As Double
    
    '������������������������������������������������������������������������������������
    '1����ʼ��һЩ����
    Set gcol������� = New Collection
    With g��������
        .����ID = rs������ϸ("����ID")
        
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", CLng(rs������ϸ("����ID")))
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        .��ҳID = rsTemp("��ҳID")
        .��� = Int(Format(zlDatabase.Currentdate, "yyyy"))
    End With
    
    bln�����ʻ�֧��ȫ�Է� = clsҽ��.GetCapability(support�����ʻ�ȫ�Է�, 0, intinsure)
    bln�����ʻ�֧�������Ը� = clsҽ��.GetCapability(support�����ʻ������Ը�, 0, intinsure)
    bln�����ʻ�֧������ = clsҽ��.GetCapability(support�����ʻ�����, 0, intinsure)

        
    '1.2 �������˵���Ժʱ��
    gstrSQL = "select ��Ժ����,nvl(��Ժ����,to_date('3000-01-01','yyyy-MM-dd')) as ��Ժ���� " & _
              "from ������ҳ where ����ID=[1] and ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", g��������.����ID, g��������.��ҳID)
    If rsTemp("��Ժ����") = CDate("3000-01-01") Then
        g��������.��;���� = 1
    Else
        '��ʾ�ò����Ѿ���Ժ
        g��������.��;���� = 0
    End If

    '1.3 ��������סԺ�ڼ��ۼƽ������
    gstrSQL = "select nvl(sum(A.����),0) as ����,nvl(sum(A.����ͳ����),0) as ����ͳ���� " & _
              "  from ���ս����¼ A,���˽��ʼ�¼ B " & _
              "  Where A.����ID = [1] And A.��ҳID = [2]" & _
              " And A.���� = [3] And A.��¼ID = B.ID "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", g��������.����ID, g��������.��ҳID, intinsure)
    dbl������ߺ� = rsTemp("����")
    dbl��ν���ͳ��� = rsTemp("����ͳ����")
    
    With g��������
        gstrSQL = "select A.����,A.��Ա���,A.��ְ,A.�����," & _
                  "      B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ�" & _
                  " from �����ʻ� A,�ʻ������Ϣ B" & _
                  " where A.����ID=B.����ID(+) and A.����=B.����(+) " & _
                  "     and B.���(+)=[1] and A.����ID=[2] and A.����=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", .���, .����ID, intinsure)
        
        lng���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
        lng��ְ = IIf(IsNull(rsTemp("��ְ")), 1, rsTemp("��ְ"))
        lng���� = IIf(IsNull(rsTemp("�����")), 0, rsTemp("�����"))
        .סԺ���� = IIf(IsNull(rsTemp("סԺ�����ۼ�")), 0, rsTemp("סԺ�����ۼ�"))
        .�ʻ��ۼ����� = IIf(IsNull(rsTemp("�ʻ������ۼ�")), 0, rsTemp("�ʻ������ۼ�"))
        .�ʻ��ۼ�֧�� = IIf(IsNull(rsTemp("�ʻ�֧���ۼ�")), 0, rsTemp("�ʻ�֧���ۼ�"))
        .�ۼƽ���ͳ�� = IIf(IsNull(rsTemp("����ͳ���ۼ�")), 0, rsTemp("����ͳ���ۼ�"))
        .�ۼ�ͳ�ﱨ�� = IIf(IsNull(rsTemp("ͳ�ﱨ���ۼ�")), 0, rsTemp("ͳ�ﱨ���ۼ�"))
    
        
        gstrSQL = "select �����,nvl(ȫ��ͳ��,0) as ȫ��ͳ�� ,nvl(������,0) as ������ ,nvl(�޷ⶥ��,0) as �޷ⶥ�� " & _
                " from ���������" & _
                " where ����=[1] and nvl(����,0)=[2]" & _
                "       and ��ְ=[3] and ����<=[4] and ([4]<=���� or ����=0)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", intinsure, lng����, lng��ְ, lng����)
        If rsTemp.RecordCount = 0 Then
            MsgBox "���ڡ��������������������������õ���", vbInformation, gstrSysName
            Exit Function
        End If
        lng����� = rsTemp("�����")
        blnȫ��ͳ�� = (rsTemp("ȫ��ͳ��") = 1)
        bln������ = (rsTemp("������") = 1)
        bln�޷ⶥ�� = (rsTemp("�޷ⶥ��") = 1)
    End With
    
    '������������������������������������������������������������������������������������
    '2����ͳ��֧����Ŀ�ϼƷ�����������
    '2.1����ʼ����¼��
    Set rs������� = New ADODB.Recordset
    With rs�������
        If .State = adStateOpen Then .Close
        .Fields.Append "���մ���ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 8, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ͳ����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With

    Do Until rs������ϸ.EOF
    'װ����д���¼��������������ʹ��
        If rs������ϸ("������Ŀ��") = 1 Then
            If rs�������.RecordCount = 0 Then
                rs�������.AddNew
                rs�������("���մ���ID") = rs������ϸ("���մ���ID")
                rs�������("����") = rs������ϸ("����")
                rs�������("���") = rs������ϸ("���")
            Else
                rs�������.MoveFirst
                rs�������.Find "���մ���ID=" & rs������ϸ("���մ���ID")
                If rs�������.EOF Then
                    rs�������.AddNew
                    rs�������("���մ���ID") = rs������ϸ("���մ���ID")
                    rs�������("����") = rs������ϸ("����")
                    rs�������("���") = rs������ϸ("���")
                Else
                    rs�������("����") = rs�������("����") + rs������ϸ("����")
                    rs�������("���") = rs�������("���") + rs������ϸ("���")
                End If
            End If
            rs�������.Update
        Else
            curȫ�Է� = curȫ�Է� + rs������ϸ("���")
        End If
            
        dblTemp = dblTemp + rs������ϸ("���")
        rs������ϸ.MoveNext
    Loop
    g��������.�������ý�� = dblTemp
    
    '2.2���������ͳ����
    gstrSQL = "select ID,�㷨,ͳ��ȶ�,��׼����,��׼����,�Ƿ�ҽ�� FROM ����֧������  where ����=[1]"
    Set rs�㷨 = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", intinsure)
    
    dblTemp = 0
    g��������.�Żݽ�� = 0
    If rs�������.RecordCount > 0 Then rs�������.MoveFirst
    Do Until rs�������.EOF
        
        rs�㷨.Filter = "ID=" & rs�������("���մ���ID")
        If rs�㷨.RecordCount > 0 Then
            If rs�㷨("�Ƿ�ҽ��") = 1 Then
            
                '�㷨:1-�ܶ������Ŀ��2-סԺ�պ˶���Ŀ
                Select Case Nvl(rs�㷨!�㷨, 2)
                Case 1      '1-�ܶ������Ŀ
                        If rs�㷨("ͳ��ȶ�") = 0 Then
                            curȫ�Է� = curȫ�Է� + rs�������("���")
                        Else
                            dblTemp = dblTemp + rs�������("���") * rs�㷨("ͳ��ȶ�") / 100
                        End If
                Case 2      '2-סԺ�պ˶���Ŀ
                        If Val(rs�������("����")) > Val(rs�㷨("��׼����")) Then
                            '���סԺ�ճ�����׼��������ô�������� ��׼����*��׼���� +  (����-��׼����)*ͳ��ȶ�
                            '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                            dbl����� = rs�㷨("��׼����") * rs�㷨("��׼����") + _
                                (rs�������("����") - IIf(rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0, 0, rs�㷨("��׼����"))) * rs�㷨("ͳ��ȶ�")
                        Else
                            '���סԺ�յ�����׼��������ô�������� ����*��׼���� ���� ����*ͳ��ȶ�
                            '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                            If rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0 Then
                                dbl����� = rs�������("����") * rs�㷨("ͳ��ȶ�")
                            Else
                                dbl����� = rs�������("����") * rs�㷨("��׼����")
                            End If
                        End If
                        
                        '�ܽ��������С����ȡȫ��������ֻ�����
                        dblTemp = dblTemp + IIf(rs�������("���") < dbl�����, rs�������("���"), dbl�����)
                        
                        If rs�������("���") > dbl����� Then
                            'ȫ������ȫ�Է�
                            curȫ�Է� = curȫ�Է� + rs�������("���") - dbl�����
                        End If
                Case Else   '3-���õ��μ��㷨
                    If Nvl(rs�������!���, 0) = 0 Then
                    Else
                        dblTemp1 = ��ȡ���õ��ζ�_����(Nvl(rs�������!���մ���id, 0), Nvl(rs�������!���, 0))
                        dblTemp = dblTemp + dblTemp1
                        g��������.�Żݽ�� = g��������.�Żݽ�� + (Nvl(rs�������!���, 0) - dblTemp1)
                    End If
                End Select
            Else
                curȫ�Է� = curȫ�Է� + rs�������("���")
            End If
        Else
            curȫ�Է� = curȫ�Է� + rs�������("���")
        End If
        rs�������.MoveNext
    Loop
    g��������.����ͳ���� = dblTemp
    g��������.ȫ�Էѽ�� = curȫ�Է�
    g��������.�����Ը���� = g��������.�������ý�� - curȫ�Է� - dblTemp - g��������.�Żݽ��
    
    '������������������������������������������������������������������������������������
    '3��������ߡ��ⶥ�ߡ�֧������������
    '3.1��������ߡ��ⶥ��
    With g��������
        
        gstrSQL = "select max(decode(A.����,'A',A.���,0)) as ������ ,max(decode(A.����,'1',A.���,0)) as ���� " & _
                  "         ,max(decode(A.����,'" & (.סԺ���� + 1) & "',A.���,0)) as ʵ������,min(A.���) as ������� " & _
                  "  from ����֧���޶� A " & _
                  "  where A.����=[1] and A.����=[2] and A.���=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", intinsure, lng����, .���)
                
        If bln������ Then
            .ʵ������ = 0
            .���� = 0
        Else
            .���� = IIf(IsNull(rsTemp("ʵ������")), 0, rsTemp("ʵ������"))
            If .���� = 0 Then
                'һ�㶼���У����ʵ�ڳ�����סԺ��������ȡ���һ�Σ�Ҳ���ǽ����С��һ�Σ�
                .���� = IIf(IsNull(rsTemp("�������")), 0, rsTemp("�������"))
            End If
            If .���� = 0 Then
                MsgBox "���ڡ���Ƚ�����������ñ���ȵ����ߡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If bln�޷ⶥ�� Then
            .�ⶥ�� = 0
        Else
            .�ⶥ�� = IIf(IsNull(rsTemp("������")), 0, rsTemp("������"))
            If .�ⶥ�� = 0 Then
                MsgBox "���ڡ���Ƚ�����������ñ���ȵķⶥ�ߡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
        '3.2��������ǰ�۳������߽��ó����ε�ʵ������
        If dbl������ߺ� > 0 Then
            '�����ò��˿϶��ж�ν���
            
            If dbl������ߺ� > dbl��ν���ͳ��� Then
                '�ò��˵ı��ν��㻹Ҫ�۳�һ�������߽��
                dbl�������� = dbl������ߺ� - dbl��ν���ͳ���
            Else
                '�����Ѿ�����
                dbl�������� = 0
            End If
            
            If .���� > dbl������ߺ� Then
                '���������ߣ�Ҫ����β�ֵ
                .���� = .���� - dbl������ߺ�
            Else
                '��ǰ�����߽���Ѿ�ȫ��棬���β����ٱ�����
                .���� = 0
            End If
                
            dbl�������� = dbl�������� + .����
        Else
            dbl�������� = .����
        End If
        dbl�������� = dbl��������
    End With
    
    '3.3��ȡ�÷��õ���
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "select B.����,B.����,B.����,A.���� " & _
              "  from ����֧������ A,���շ��õ� B " & _
              "  Where A.���� =[1] And A.���� =[2] And A.��� =[3] And A.��ְ =[4] And A.����� =[5]" & _
              "       and A.����=B.���� and A.����=b.���� and A.����=B.���� " & _
              "  order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", intinsure, lng����, g��������.���, lng��ְ, lng�����)
    If rsTemp.RecordCount = 0 Then
        MsgBox "���ڡ���Ƚ�����������ñ���ȵ�ͳ��֧����������", vbInformation, gstrSysName
        Exit Function
    End If
    
    '������������������������������������������������������������������������������������
    '4������ôν���ɱ����Ľ��
    dbl�ۼƽ��� = 0   '����ֶ��ۼƽ���ͳ��
    dbl�ѱ������ = g��������.�ۼ�ͳ�ﱨ��
    g��������.ͳ�ﱨ����� = 0
    
    If bln������� = True Then
        '�������Ͳ��ÿ�����ǰ�Ľ�����
        dbl��ν���ͳ��� = 0
    End If
    Do Until rsTemp.EOF
        dbl�ֶν��� = 0
        dbl�ֶα��� = 0
        
        If dbl�ѱ������ < g��������.�ⶥ�� Or g��������.�ⶥ�� = 0 Then    'δ�����ⶥ�߻��޷ⶥ��
            '�����Լ�������
            dbl���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
            dbl���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
            If dbl���� = 0 Then
                If g��������.���� > dbl���� Then
                    MsgBox "�ò��˵�ʵ�����߱ȵ�һ�����õ����޻��࣬���鱣�շ��õ���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            If g��������.����ͳ���� + dbl��ν���ͳ��� > dbl���� And (dbl��ν���ͳ��� < dbl���� Or dbl���� = 0) Then
                '�ö���ǰ��δ������ȫ�����������Ҫ����۳��Ľ��
                dblTemp = 0
                If dbl��ν���ͳ��� > dbl���� Then
                    '��ǰ�Ѿ��������
                    dblTemp = dbl��ν���ͳ��� - dbl����
                End If
                
                '����Ҫ�۳�һ�������ߺ��ѽ���������޽����б仯
                If dbl���� + dblTemp + dbl�������� > dbl���� And dbl���� > 0 Then
                    dbl���� = dbl����
                    dbl�������� = dbl�������� - (dbl���� - dbl���� - dblTemp) '�����Ѿ����꣬�����¶ο�
                Else
                    dbl���� = dbl���� + dbl�������� + dblTemp
                    dbl�������� = 0
                End If
                
                If g��������.����ͳ���� + dbl��ν���ͳ��� <= dbl���� Or dbl���� = 0 Then
                    '��ʵ��ֵ����
                    dbl�ֶν��� = g��������.����ͳ���� + dbl��ν���ͳ��� - dbl����
                    
                    '������ڼ������ߡ�����ǰ�Ľ��ʽ����½���ͳ��Ľ����ܴﵽ���ޣ���ֻ��ȡ0
                    If dbl�ֶν��� < 0 Then dbl�ֶν��� = 0
                Else
                    'ȫ�����
                    dbl�ֶν��� = dbl���� - dbl����
                End If
                '����������öεı������
                dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
                dbl�ֶα��� = Val(Format(dbl�ֶν��� * rsTemp("����") / 100, "0.00"))
                
                If dbl�ѱ������ + dbl�ֶα��� > g��������.�ⶥ�� And g��������.�ⶥ�� <> 0 Then
                    '���������˷ⶥ�ߣ����Ҵ��ڷⶥ������
                    dbl�ֶα��� = g��������.�ⶥ�� - dbl�ѱ������
                    
                    '���ƽ���ͳ����
                    If rsTemp("����") <> 0 Then
                        dbl�ֶν��� = dbl�ֶα��� * 100 / rsTemp("����")
                    Else
                        dbl�ֶν��� = 0
                    End If
                End If
                
                '���и�ʽ��
                dbl�ֶν��� = Val(Format(dbl�ֶν���, "0.00"))
                dbl�ֶα��� = Val(Format(dbl�ֶα���, "0.00"))
                
                dbl�ѱ������ = dbl�ѱ������ + dbl�ֶα���
                g��������.ͳ�ﱨ����� = g��������.ͳ�ﱨ����� + dbl�ֶα���
            End If
        End If
        
        '���Ρ�����ͳ���ͳ�ﱨ��������
        lng���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
        dblTemp = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
        dbl�ۼƽ��� = dbl�ֶν��� + dbl�ۼƽ���
            
        gcol�������.Add Array(lng����, dbl�ֶν���, dbl�ֶα���, dblTemp)
        rsTemp.MoveNext
    Loop
    
    g��������.ʵ������ = dbl�������� - dbl��������
    
    With g��������
        '���㳬���Ը�����
        .�����Ը���� = .����ͳ���� - dbl�������� - dbl�ۼƽ���
        If .�����Ը���� < 0 Then .�����Ը���� = 0                   '�������ͳ����������ߣ�Ϊ����
    End With
    
    If blnȫ��ͳ�� = True Then
        סԺ�������_���� = "ҽ������;" & g��������.ͳ�ﱨ����� + g��������.�����Ը���� & ";0"
    Else
        סԺ�������_���� = "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
    End If
    
    '����Ҫ���Ǹ����ʻ���֧����Χ
    With g��������
        dblTemp = 0   '��ʱ�����ʹ�õĸ����ʻ����
        
        If bln�����ʻ�֧��ȫ�Է� = True Then
            dblTemp = dblTemp + .ȫ�Էѽ��
        End If
        
        If bln�����ʻ�֧�������Ը� = True And blnȫ��ͳ�� = False Then
            dblTemp = dblTemp + .�����Ը����
        End If
        
        If bln�����ʻ�֧������ = True Then
            'ֻ��֧������ͳ���δ�����Ĳ���
            dblTemp = dblTemp + .����ͳ���� - .ͳ�ﱨ�����
        Else
            dblTemp = dblTemp + .����ͳ���� - .ͳ�ﱨ����� - .�����Ը���� - .�Żݽ��
        End If
        '�����ʻ�֧����ܳ����ʻ����⵼���˲��ֵ������˵��ʻ����Ϊ�㣬���������Ȼ���ʻ�֧����
        If dblTemp > (g��������.�ʻ��ۼ����� - g��������.�ʻ��ۼ�֧��) Then
            dblTemp = (g��������.�ʻ��ۼ����� - g��������.�ʻ��ۼ�֧��)
        End If
        If .�Żݽ�� <> 0 Then
            סԺ�������_���� = סԺ�������_���� & "|�Żݽ��;" & .�Żݽ�� & ";0" & "|�����ʻ�;" & dblTemp & ";1"
        Else
            סԺ�������_���� = סԺ�������_���� & "|�����ʻ�;" & dblTemp & ";1"
        End If
    End With
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    Dim rsTemp As New ADODB.Recordset
    Dim cur�����ʻ� As Currency
    Dim var������� As Variant
On Error GoTo ErrH
    '������ʻ�֧�����
    gstrSQL = "Select Nvl(��Ԥ��,0) as ��� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ��¼����=2 And ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", lng����ID)
    
    If rsTemp.RecordCount > 0 Then
        cur�����ʻ� = rsTemp("���")
    End If
    
    With g��������
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & .����ID & "," & intinsure & "," & .��� & "," & _
            .�ʻ��ۼ����� & "," & .�ʻ��ۼ�֧�� + cur�����ʻ� & "," & .�ۼƽ���ͳ�� + .����ͳ���� & "," & _
            .�ۼ�ͳ�ﱨ�� + .ͳ�ﱨ����� & "," & .סԺ���� + 1 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
        
        '����
        '   ����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN
        '   �ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,�������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN
        '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,֧��˳���_IN,��ҳID_IN,
        '   ��;����_IN,��ע_IN

        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & intinsure & "," & .����ID & "," & _
            .��� & "," & .�ʻ��ۼ����� & "," & .�ʻ��ۼ�֧�� & "," & .�ۼƽ���ͳ�� & "," & _
            .�ۼ�ͳ�ﱨ�� & "," & .סԺ���� + 1 & "," & .���� & "," & .�ⶥ�� & "," & .ʵ������ & "," & _
            .�������ý�� & "," & .ȫ�Էѽ�� & "," & .�����Ը���� & "," & .����ͳ���� & "," & .ͳ�ﱨ����� & ",0," & _
            .�����Ը���� & "," & cur�����ʻ� & "," & g��������.�Żݽ�� & "," & .��ҳID & "," & .��;���� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
        
        
        For Each var������� In gcol�������
            '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
            gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
                var�������(0) & "," & var�������(1) & "," & var�������(2) & "," & var�������(3) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
        Next
    End With
    
    סԺ����_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_����(lng����ID As Long, ByVal intinsure As Integer) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rs�ʻ� As New ADODB.Recordset, rs������� As New ADODB.Recordset
    Dim lng����ID As Long
    Dim lngסԺ���� As Long, cur�ʻ����� As Currency, cur�ʻ�֧�� As Currency, cur�ۼƽ���ͳ�� As Currency, cur�ۼ�ͳ�ﱨ�� As Currency
    
On Error GoTo ErrH

    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", lng����ID)
    
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "Select * " & _
              "  From ���ս����¼ Where ����=2 and ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", lng����ID)
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "�ò��˵�ҽ���������ݶ�ʧ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
    gstrSQL = "select B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ� " & _
              " from �����ʻ� A,�ʻ������Ϣ B " & _
              " where A.����ID=B.����ID(+) and A.����=B.����(+) and B.���(+)=[1] and A.����ID=[2] and A.����=[3]"
    Set rs�ʻ� = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", Year(zlDatabase.Currentdate), CLng(rsTemp("����ID")), intinsure)
    
    If rs�ʻ�.EOF = False Then
        lngסԺ���� = IIf(IsNull(rs�ʻ�("סԺ�����ۼ�")), 0, rs�ʻ�("סԺ�����ۼ�"))
        cur�ʻ����� = IIf(IsNull(rs�ʻ�("�ʻ������ۼ�")), 0, rs�ʻ�("�ʻ������ۼ�"))
        cur�ʻ�֧�� = IIf(IsNull(rs�ʻ�("�ʻ�֧���ۼ�")), 0, rs�ʻ�("�ʻ�֧���ۼ�"))
        cur�ۼƽ���ͳ�� = IIf(IsNull(rs�ʻ�("����ͳ���ۼ�")), 0, rs�ʻ�("����ͳ���ۼ�"))
        cur�ۼ�ͳ�ﱨ�� = IIf(IsNull(rs�ʻ�("ͳ�ﱨ���ۼ�")), 0, rs�ʻ�("ͳ�ﱨ���ۼ�"))
    End If
    
    '���˴������ݱ���������������ݱ������һ������
    '��˾Ͳ���Ҫ�������������
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & intinsure & "," & rsTemp("���") & "," & _
        cur�ʻ����� & "," & cur�ʻ�֧�� - rsTemp("�����ʻ�֧��") & "," & cur�ۼƽ���ͳ�� - rsTemp("����ͳ����") & "," & _
        cur�ۼ�ͳ�ﱨ�� - rsTemp("ͳ�ﱨ�����") & "," & lngסԺ���� - 1 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    '�������ݣ������˼����ۼ�
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & intinsure & "," & rsTemp("����ID") & "," & _
        rsTemp("���") & "," & cur�ʻ����� & "," & cur�ʻ�֧�� - rsTemp("�����ʻ�֧��") & "," & cur�ۼƽ���ͳ�� - rsTemp("����ͳ����") & "," & _
        cur�ۼ�ͳ�ﱨ�� - rsTemp("ͳ�ﱨ�����") & "," & lngסԺ���� & "," & rsTemp("����") * -1 & "," & rsTemp("�ⶥ��") & "," & rsTemp("ʵ������") * -1 & "," & _
        rsTemp("�������ý��") * -1 & "," & rsTemp("ȫ�Ը����") * -1 & "," & rsTemp("�����Ը����") * -1 & "," & rsTemp("����ͳ����") * -1 & "," & _
        rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") * -1 & "," & rsTemp("�����ʻ�֧��") * -1 & ",''," & _
        IIf(IsNull(rsTemp("��ҳID")), "null", rsTemp("��ҳID")) & "," & rsTemp("��;����") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    
    gstrSQL = "select ����,����ͳ����,ͳ�ﱨ�����,���� from ���ս������ where ����ID=[1]"
    Set rs������� = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", lng����ID)
    
    Do Until rs�������.EOF
        '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
        gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
            rs�������("����") & "," & rs�������("����ͳ����") * -1 & "," & rs�������("ͳ�ﱨ�����") * -1 & "," & rs�������("����") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
        
        rs�������.MoveNext
    Loop
    
    סԺ�������_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ������Ϣ_����(ByVal lngErrCode As Long) As String
'���ܣ����ݴ���ŷ��ش�����Ϣ

End Function

Public Function ��ȡ���õ��ζ�_����(ByVal lng����ID As Long, ByVal dbl��� As Double) As Double
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ݴ���ID�ͽ���������ͳ����
    '--�����:
    '--������:
    '--��  ��:����ͳ����
    '--�� ��:���˺� 20040617
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "" & _
        "   Select " & _
        "       nvl(Sum(decode(sign(" & dbl��� & "-nvl(����,0)),-1,0, " & _
        "              decode(sign(" & dbl��� & "-nvl(����,90009000900099.99)),1, " & _
        "                     decode(nvl(����,0),0," & dbl��� & "-nvl(����,0),����-nvl(����,0)),   " & _
        "                      " & dbl��� & "-nvl(����,0)))*����/100)," & dbl��� & ") as ͳ���  " & _
        "   From ���൵�α��� " & _
        "   Where ����id=" & lng����ID

    zlDatabase.OpenRecordset rsTemp, gstrSQL, "���㱣�մ����ͳ���"
    ��ȡ���õ��ζ�_���� = Nvl(rsTemp!ͳ���, dbl���)
    rsTemp.Close
End Function

Public Function ҽ����Ŀ_����(lngItemID As Long, Optional ByVal intinsure As Integer) As String
    Dim rsTemp As New ADODB.Recordset
    

    gstrSQL = "Select nvl(B.����,'') as ���� From ����֧����Ŀ A,����֧������ B Where A.����=" & intinsure & " and A.����id=B.id and " & "A.�շ�ϸĿid=" & lngItemID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡģ��ҽ����Ŀ�ķ�������"
        
    If rsTemp.RecordCount = 0 Then
        ҽ����Ŀ_���� = ""
        Exit Function
    End If

    ҽ����Ŀ_���� = rsTemp!����
End Function
