Attribute VB_Name = "mdl����"
Option Explicit
'===============================================================================
'��μ�������ҽ������.DOC���˽���ϸ���������Ϊ����
'���ݡ���ְ��Ա�������ԷѶ�
'������Ա��������ǰ�Ϲ��ˣ����﷢����ҽ�Ʒ���ȫ����ͳ��ҽ�ƻ���֧��
'
'���ﲡ�˺�סԺ���˵ı�������һ��
'   1������ʹ�ø����ʻ�֧��
'   2�����ԷѶν�����ڵ��Ը����൱�����ߣ�
'   3�������ԷѶεģ��������Ը�
'       0-5000;         �Ը�10%
'       5000-10000;     �Ը�8%
'       10000-��;       �Ը�2%
'       ���ϱ����ֵ��ۼӼ���
'
'   ����ʹ�ø����ʻ�֧����������ۼƼ��㣬�ԷѶ����Ը��Ľ�
'�����ڣ����ս����¼.�ۼƽ���ͳ��У���������ԷѶΣ������µĽ�
'����������ĵ��������
'
'===============================================================================
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _

Private Const str���ֲ� As String = "'����','��Ⱦ��','ְҵ��','����','�ƻ�����','��֢','��λ֧��','��������'"
Private mCur�����Ը��� As Currency          '��¼���εĸ����Ը���
Private mCur�����Ը���_֧�� As Currency     '��¼����ʵ��֧���ĸ����Ը��β��ֽ��

Public Function ҽ����ʼ��_����(ByVal int���� As Integer) As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false

    
    'Ϊ�˱�����Ȩ�Ѷ����ӣ��˴����ٽ��жԸ���ҽ�������ݵļ��
    ҽ����ʼ��_���� = True
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    
    Dim strTmpIden As String
    
    strTmpIden = frmIdentify����.ShowCard(bytType, lng����ID)
    ��ݱ�ʶ_���� = strTmpIden
End Function

Public Function �������_����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: bytYear-�������,0-�������,1-�������,2-�������
'����: ���ظ����ʻ����Ľ��
    Dim rsTemp As New ADODB.Recordset

    
    gstrSQL = "select A.�ʻ���� from �����ʻ� A where A.����ID=[1] and A.����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID, TYPE_��������)
    
    If rsTemp.EOF Then
        �������_���� = 0
    Else
        �������_���� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
    End If

End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false

    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, _
            ByVal curȫ�Է� As Currency, ByVal cur�����Ը� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
    Dim var������� As Variant
On Error GoTo ErrH
    mCur�����Ը��� = mCur�����Ը��� - mCur�����Ը���_֧��
    With g��������
        cur�����ʻ� = .�����ʻ�֧��
        '�����ʻ������Ϣʱ�����ۼ�ͳ�ﱨ��ʼ��Ϊ�㣨���ֶ�δ�ã�
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�������� & "," & .��� & "," & _
            .�ʻ��ۼ����� & "," & .�ʻ��ۼ�֧�� + cur�����ʻ� & "," & mCur�����Ը��� & "," & _
            .�ۼ�ͳ�ﱨ�� + .ͳ�ﱨ����� & "," & IIf(.��ҳID = 0, "NULL", .��ҳID) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�������� & "," & .����ID & "," & _
            .��� & "," & .�ʻ��ۼ����� & "," & .�ʻ��ۼ�֧�� & "," & .�ۼƽ���ͳ�� & "," & _
            .�ۼ�ͳ�ﱨ�� & "," & IIf(.��ҳID = 0, "NULL", .��ҳID) & "," & .���� & "," & .�ⶥ�� & "," & .ʵ������ & "," & _
            .�������ý�� & "," & .ȫ�Էѽ�� & "," & .�����Ը���� & "," & .����ͳ���� & "," & .ͳ�ﱨ����� & ",0," & _
            .�����Ը���� & "," & cur�����ʻ� & ",'')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        For Each var������� In gcol�������
            '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
            gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
                var�������(0) & "," & var�������(1) & "," & var�������(2) & "," & var�������(3) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        Next
    End With

    �������_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function


Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset
    Dim rs�ʻ� As New ADODB.Recordset
    Dim rs������� As New ADODB.Recordset
    Dim lng����ID As Long
    Dim cur�ʻ����� As Currency, cur�ʻ�֧�� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim lngסԺ���� As Long
    Dim curDate As Date
On Error GoTo ErrH
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "Select * From ���ս����¼ Where ����=1 And ��¼ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "�ò��˵�ҽ���������ݶ�ʧ���������ϡ�"
        Exit Function
    End If
    
    '�ʻ������Ϣ
    gstrSQL = "select B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ� " & _
              " from �����ʻ� A,�ʻ������Ϣ B " & _
              " where A.����ID=B.����ID(+) and A.����=B.����(+) and B.���(+)=" & Year(zlDatabase.Currentdate) & " and A.����ID=" & rsTemp("����ID") & " and A.����=" & TYPE_��������
    Call OpenRecordset(rs�ʻ�, "����ҽ��")
    
    If rs�ʻ�.EOF = False Then
        lngסԺ���� = IIf(IsNull(rs�ʻ�("סԺ�����ۼ�")), 0, rs�ʻ�("סԺ�����ۼ�"))
        cur�ʻ����� = IIf(IsNull(rs�ʻ�("�ʻ������ۼ�")), 0, rs�ʻ�("�ʻ������ۼ�"))
        cur�ʻ�֧�� = IIf(IsNull(rs�ʻ�("�ʻ�֧���ۼ�")), 0, rs�ʻ�("�ʻ�֧���ۼ�"))
        cur����ͳ���ۼ� = IIf(IsNull(rs�ʻ�("����ͳ���ۼ�")), 0, rs�ʻ�("����ͳ���ۼ�"))
        curͳ�ﱨ���ۼ� = IIf(IsNull(rs�ʻ�("ͳ�ﱨ���ۼ�")), 0, rs�ʻ�("ͳ�ﱨ���ۼ�"))
    End If
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & TYPE_�������� & "," & rsTemp("���") & "," & _
        cur�ʻ����� & "," & cur�ʻ�֧�� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� & "," & _
        0 & "," & lngסԺ���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�������� & "," & rsTemp("����ID") & "," & _
        rsTemp("���") & "," & cur�ʻ����� & "," & cur�ʻ�֧�� - rsTemp("�����ʻ�֧��") & "," & rsTemp("�ۼƽ���ͳ��") * -1 & "," & _
        curͳ�ﱨ���ۼ� & ",NULL," & rsTemp("����") * -1 & "," & rsTemp("�ⶥ��") & "," & rsTemp("ʵ������") * -1 & "," & _
        rsTemp("�������ý��") * -1 & "," & rsTemp("ȫ�Ը����") * -1 & "," & rsTemp("�����Ը����") * -1 & "," & rsTemp("����ͳ����") * -1 & "," & _
        rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") * -1 & "," & rsTemp("�����ʻ�֧��") * -1 & ",NULL)" 'curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����")
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "select ����,����ͳ����,ͳ�ﱨ�����,���� from ���ս������ where ����ID=" & lng����ID
    Call OpenRecordset(rs�������, "����ҽ��")
    
    Do Until rs�������.EOF
        '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
        gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
            rs�������("����") & "," & rs�������("����ͳ����") * -1 & "," & rs�������("ͳ�ﱨ�����") * -1 & "," & rs�������("����") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        rs�������.MoveNext
    Loop

    ����������_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs��ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."�Ƿ������޸�:0-�������޸�;1-�����޸�
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs��ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����Ҫ��NO����š�����ID��ҽ����Ŀ���롢�շ�����շ����ơ��������š���񡢲��ء��������۸񡢽�ҽ��,�Ǽ�ʱ��(����ʱ��),Ӥ����,���մ���ID
    Dim rs�㷨 As New ADODB.Recordset          '����
    Dim rsTemp As New ADODB.Recordset
    Dim rs������� As New ADODB.Recordset
    
    Dim lng���� As Long
    Dim lng��ְ As Long, lng����� As Long, lng���� As Long
    Dim dblTemp As Double, lng���� As Long
    
    Dim dbl�����  As Double ''��һ����סԺ�ռ������Ŀ������ܵõ��Ľ��
    Dim dbl�ѱ������ As Double, dbl�ۼƽ��� As Double
    Dim dbl���� As Double, dbl���� As Double, dbl�ֶν��� As Double, dbl�ֶα��� As Double
    
    Dim clsҽ�� As New clsInsure
    Dim bln�����ʻ�֧��ȫ�Է� As Boolean, bln�����ʻ�֧�������Ը� As Boolean, bln�����ʻ�֧������ As Boolean
    Dim curȫ�Է� As Currency, cur�����Ը� As Currency
    Dim bln������ As Boolean, bln�޷ⶥ�� As Boolean
    Dim dbl�ʻ����
    Dim dbl������ߺ� As Double   '�����ָ�ò�����ǰ���ʵ��ۼ�
    Dim dbl�������� As Double     '���ε�����
    Dim blnExit As Boolean          '���ڸ����ʻ��������ߣ��ԷѶΣ����򱣴���ؼ�¼���˳�
    Dim bln������Ա As Boolean      '������Ա�����������ķ��ã�ȫ��ͳ��ҽ�ƻ���֧��
    Dim bln���ֲ����� As Boolean
    Dim bln���񲡴�Ⱦ������ As Boolean, bln�ƻ����� As Boolean, bln�������� As Boolean
    Dim str�������� As String, str��׼��Ŀ As String, dbl��׼��Ŀ As Currency, lng����ID As Long
    
    '������������������������������������������������������������������������������������
    '1����ʼ��һЩ����
    Set gcol������� = New Collection
    With g��������
        .����ID = rs��ϸ("����ID")
        .��ҳID = 0
        .��� = Int(Format(zlDatabase.Currentdate, "yyyy"))
    End With

    '1.1 �������ξ��﷢�����ۼƽ���������ۼƽ���ͳ������ԷѶΣ��򰴱�����ͬ�������������ʻ�������֧����
    '�ۼƽ���ͳ����Ϊÿ��֧�����ԷѶν�����ȡ����ȵ��ۼƽ���ͳ���
    gstrSQL = "select nvl(sum(A.�ۼƽ���ͳ��),0) as ���� " & _
              "  from ���ս����¼ A " & _
              "  Where A.����ID = " & g��������.����ID & " And A.���� = " & TYPE_�������� & " And A.���=" & g��������.���
    Call OpenRecordset(rsTemp, "�������")
    dbl������ߺ� = rsTemp("����")
    
    With g��������
        g��������.ͳ�ﱨ����� = 0
        g��������.�ۼƽ���ͳ�� = 0
        g��������.�ۼ�ͳ�ﱨ�� = 0
        g��������.ȫ�Էѽ�� = 0
        g��������.�����Ը���� = 0
        g��������.�����ʻ�֧�� = 0
        
        gstrSQL = "select A.����,A.��Ա���,A.��ְ,A.�����,Nvl(C.���,0) ����,Nvl(C.ID,0) ����ID,C.���� ��������," & _
                  "      B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ�" & _
                  " from �����ʻ� A,�ʻ������Ϣ B," & _
                  "         (Select * From ���ղ��� Where ���<>2" & _
                  "          Union " & _
                  "          Select * From ���ղ��� Where ���=2 And ���� In (" & str���ֲ� & ")) C" & _
                  " where A.����ID=B.����ID(+) and A.����=B.����(+) And A.����=C.����(+) ANd A.����ID=C.ID(+) " & _
                  "     and B.���(+)=" & .��� & " and A.����ID=" & .����ID & " and A.����=" & TYPE_��������
        Call OpenRecordset(rsTemp, "�������")
        
        '1-��ְ;2-����;3-����
        '���ݼ�������Ա�����������ߣ��ԷѶΣ�
        lng���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
        lng��ְ = IIf(IsNull(rsTemp("��ְ")), 1, rsTemp("��ְ"))
        lng���� = IIf(IsNull(rsTemp("�����")), 0, rsTemp("�����"))
        bln���ֲ����� = (rsTemp!���� = 2)
        bln������Ա = (lng��ְ = 3)
        lng����ID = rsTemp!����ID
        str�������� = IIf(IsNull(rsTemp!��������), "", rsTemp!��������)
        
        .סԺ���� = 1   '��ҽ����סԺ�����޹�
        .�ʻ��ۼ����� = IIf(IsNull(rsTemp("�ʻ������ۼ�")), 0, rsTemp("�ʻ������ۼ�"))
        .�ʻ��ۼ�֧�� = IIf(IsNull(rsTemp("�ʻ�֧���ۼ�")), 0, rsTemp("�ʻ�֧���ۼ�"))
        mCur�����Ը��� = IIf(IsNull(rsTemp("����ͳ���ۼ�")), 0, rsTemp("����ͳ���ۼ�"))
        .�ۼ�ͳ�ﱨ�� = IIf(IsNull(rsTemp("ͳ�ﱨ���ۼ�")), 0, rsTemp("ͳ�ﱨ���ۼ�"))
        
        gstrSQL = "select �����,nvl(������,0) as ������,nvl(�޷ⶥ��,0) as �޷ⶥ�� " & _
                " from ���������" & _
                " where ����=" & TYPE_�������� & " and nvl(����,0)=" & lng���� & _
                "       and ��ְ=" & lng��ְ & " and ����<=" & lng���� & " and (" & lng���� & "<=���� or ����=0)"
        Call OpenRecordset(rsTemp, "�������")
        If rsTemp.RecordCount = 0 Then
            MsgBox "���ڡ��������������������������õ���", vbInformation, gstrSysName
            Exit Function
        End If
        lng����� = rsTemp("�����")
        bln������ = (rsTemp("������") = 1) Or (lng��ְ <> 1)
        bln�޷ⶥ�� = (rsTemp("�޷ⶥ��") = 1)
    End With
    
    '������������������������������������������������������������������������������������
    '2����ȡʵ�ʷ�������
    If Not clsҽ��.GetCapability(support��������ҽ����Ŀ, 0, TYPE_��������) Then
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
    
        Do Until rs��ϸ.EOF
        'װ����д���¼��������������ʹ��
            If rs��ϸ("�Ƿ�ҽ��") = 1 Then
                If rs�������.RecordCount = 0 Then
                    rs�������.AddNew
                    rs�������("���մ���ID") = IIf(IsNull(rs��ϸ("����֧������ID")), 0, rs��ϸ("����֧������ID"))
                    rs�������("����") = rs��ϸ("����")
                    rs�������("���") = rs��ϸ("ʵ�ս��")
                Else
                    rs�������.MoveFirst
                    rs�������.Find "���մ���ID=" & IIf(IsNull(rs��ϸ("����֧������ID")), 0, rs��ϸ("����֧������ID"))
                    If rs�������.EOF Then
                        rs�������.AddNew
                        rs�������("���մ���ID") = IIf(IsNull(rs��ϸ("����֧������ID")), 0, rs��ϸ("����֧������ID"))
                        rs�������("����") = rs��ϸ("����")
                        rs�������("���") = rs��ϸ("ʵ�ս��")
                    Else
                        rs�������("����") = rs�������("����") + rs��ϸ("����")
                        rs�������("���") = rs�������("���") + rs��ϸ("ʵ�ս��")
                    End If
                End If
                rs�������.Update
            Else
                curȫ�Է� = curȫ�Է� + rs��ϸ("ʵ�ս��")
            End If
                
            dblTemp = dblTemp + rs��ϸ("ʵ�ս��")
            rs��ϸ.MoveNext
        Loop
        g��������.�������ý�� = dblTemp
        
        '2.2���������ͳ����
        gstrSQL = "select ID,�㷨,ͳ��ȶ�,��׼����,��׼����,�Ƿ�ҽ�� FROM ����֧������  where ����=" & TYPE_��������
        Call OpenRecordset(rs�㷨, "����ҽ��")
        
        dblTemp = 0
        If rs�������.RecordCount > 0 Then rs�������.MoveFirst
        Do Until rs�������.EOF
            
            rs�㷨.Filter = "ID=" & rs�������("���մ���ID")
            If rs�㷨.RecordCount > 0 Then
                If rs�㷨("�Ƿ�ҽ��") = 1 Then
                    '�㷨:1-�ܶ������Ŀ��2-סԺ�պ˶���Ŀ
                    If rs�㷨("�㷨") = 1 Then
                        If rs�㷨("ͳ��ȶ�") = 0 Then
                            curȫ�Է� = curȫ�Է� + rs�������("���")
                        Else
                            dblTemp = dblTemp + rs�������("���") * rs�㷨("ͳ��ȶ�") / 100
                        End If
                    Else
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
                    End If
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
        g��������.�����Ը���� = g��������.�������ý�� - curȫ�Է� - dblTemp
    Else
        '��������ܶ�
        If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
        Do Until rs��ϸ.EOF
            dblTemp = dblTemp + rs��ϸ("ʵ�ս��")
            rs��ϸ.MoveNext
        Loop
        g��������.�������ý�� = dblTemp
        g��������.����ͳ���� = dblTemp
    End If
    
    '�����������Ա�����﷢�������з��ã�ȫ����ͳ��ҽ�ƻ���֧�����ȿ۳������ʻ�
    If bln������Ա Then
        dblTemp = �������_����(g��������.����ID)
        dblTemp = IIf(dblTemp > 0, dblTemp, 0)
        dblTemp = IIf(g��������.����ͳ���� > dblTemp, dblTemp, g��������.����ͳ����)
        g��������.�����ʻ�֧�� = dblTemp
        g��������.ͳ�ﱨ����� = g��������.����ͳ���� - g��������.�����ʻ�֧��
        str���㷽ʽ = "ҽ������;" & g��������.ͳ�ﱨ����� & ";0|�����ʻ�;" & g��������.�����ʻ�֧�� & ";0"
        �����������_���� = True
        Exit Function
    End If
    
    '��������ֲ�����
    If bln���ֲ����� Then
        
        Dim rs���ֲ����� As New ADODB.Recordset
        str��׼��Ŀ = ""
        dbl��׼��Ŀ = 0
        bln���񲡴�Ⱦ������ = (InStr(1, ",����,��Ⱦ��,", "," & str�������� & ",") <> 0)
        bln�ƻ����� = (InStr(1, ",�ƻ�����,", "," & str�������� & ",") <> 0)
        bln�������� = (InStr(1, ",��������,", "," & str�������� & ",") <> 0)
        
        If bln�������� Then
            'ҩƷ���ã���ҽ������֧��50%
            g��������.����ͳ���� = 0
            With rs��ϸ
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    If InStr(1, ",5,6,7,", "," & !�շ���� & ",") <> 0 And !�Ƿ�ҽ�� = 1 Then
                        g��������.����ͳ���� = g��������.����ͳ���� + (!ʵ�ս�� * 0.5)
                    End If
                    .MoveNext
                Loop
            End With
            g��������.ͳ�ﱨ����� = g��������.����ͳ����
            str���㷽ʽ = "ҽ������;" & g��������.ͳ�ﱨ����� & ";0|�����ʻ�;0;0"
            �����������_���� = True
            Exit Function
        ElseIf bln�ƻ����� Then
            str���㷽ʽ = ���ֲ�����(str��������, lng��ְ, dbl��׼��Ŀ, False)
            �����������_���� = True
            Exit Function
        Else
            '������׼��Ŀ����ͳ����ܶ�
            With rsTemp
                gstrSQL = "Select �շ�ϸĿID From ������׼��Ŀ Where ����ID=" & lng����ID
                Call OpenRecordset(rsTemp, "�������")
                
                Do While Not .EOF
                    str��׼��Ŀ = str��׼��Ŀ & ";" & !�շ�ϸĿID
                    .MoveNext
                Loop
                str��׼��Ŀ = str��׼��Ŀ & ";"
            End With
            
            If Not clsҽ��.GetCapability(support��������ҽ����Ŀ, 0, TYPE_��������) Then
                Set rs���ֲ����� = New ADODB.Recordset
                With rs���ֲ�����
                    If .State = adStateOpen Then .Close
                    .Fields.Append "���մ���ID", adDouble, 18, adFldIsNullable
                    .Fields.Append "����", adDouble, 8, adFldIsNullable
                    .Fields.Append "���", adDouble, 18, adFldIsNullable
                    .Fields.Append "ͳ����", adDouble, 18, adFldIsNullable
                    
                    .CursorLocation = adUseClient
                    .Open , , adOpenStatic, adLockOptimistic
                End With
            
                If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
                Do Until rs��ϸ.EOF
                'װ����д���¼��������������ʹ��
                    If rs��ϸ("�Ƿ�ҽ��") = 1 And InStr(1, str��׼��Ŀ, ";" & rs��ϸ("�շ�ϸĿID") & ";") <> 0 Then
                        If rs���ֲ�����.RecordCount = 0 Then
                            rs���ֲ�����.AddNew
                            rs���ֲ�����("���մ���ID") = IIf(IsNull(rs��ϸ("����֧������ID")), 0, rs��ϸ("����֧������ID"))
                            rs���ֲ�����("����") = rs��ϸ("����")
                            rs���ֲ�����("���") = rs��ϸ("ʵ�ս��")
                        Else
                            rs���ֲ�����.MoveFirst
                            rs���ֲ�����.Find "���մ���ID=" & IIf(IsNull(rs��ϸ("����֧������ID")), 0, rs��ϸ("����֧������ID"))
                            If rs���ֲ�����.EOF Then
                                rs���ֲ�����.AddNew
                                rs���ֲ�����("���մ���ID") = IIf(IsNull(rs��ϸ("����֧������ID")), 0, rs��ϸ("����֧������ID"))
                                rs���ֲ�����("����") = rs��ϸ("����")
                                rs���ֲ�����("���") = rs��ϸ("ʵ�ս��")
                            Else
                                rs���ֲ�����("����") = rs���ֲ�����("����") + rs��ϸ("����")
                                rs���ֲ�����("���") = rs���ֲ�����("���") + rs��ϸ("ʵ�ս��")
                            End If
                        End If
                        rs���ֲ�����.Update
                    End If
                    rs��ϸ.MoveNext
                Loop
                
                '2.2���������ͳ����
                If rs�㷨.RecordCount <> 0 Then rs�㷨.MoveFirst
                If rs���ֲ�����.RecordCount > 0 Then rs���ֲ�����.MoveFirst
                Do Until rs���ֲ�����.EOF
                    
                    rs�㷨.Filter = "ID=" & rs���ֲ�����("���մ���ID")
                    If rs�㷨.RecordCount > 0 Then
                        If rs�㷨("�Ƿ�ҽ��") = 1 Then
                            '�㷨:1-�ܶ������Ŀ��2-סԺ�պ˶���Ŀ
                            If rs�㷨("�㷨") = 1 Then
                                If rs�㷨("ͳ��ȶ�") = 0 Then
                                    curȫ�Է� = curȫ�Է� + rs���ֲ�����("���")
                                Else
                                    dbl��׼��Ŀ = dbl��׼��Ŀ + rs���ֲ�����("���") * rs�㷨("ͳ��ȶ�") / 100
                                End If
                            Else
                                If Val(rs���ֲ�����("����")) > Val(rs�㷨("��׼����")) Then
                                    '���סԺ�ճ�����׼��������ô�������� ��׼����*��׼���� +  (����-��׼����)*ͳ��ȶ�
                                    '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                                    dbl����� = rs�㷨("��׼����") * rs�㷨("��׼����") + _
                                        (rs���ֲ�����("����") - IIf(rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0, 0, rs�㷨("��׼����"))) * rs�㷨("ͳ��ȶ�")
                                Else
                                    '���סԺ�յ�����׼��������ô�������� ����*��׼���� ���� ����*ͳ��ȶ�
                                    '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                                    If rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0 Then
                                        dbl����� = rs���ֲ�����("����") * rs�㷨("ͳ��ȶ�")
                                    Else
                                        dbl����� = rs���ֲ�����("����") * rs�㷨("��׼����")
                                    End If
                                End If
                                
                                '�ܽ��������С����ȡȫ��������ֻ�����
                                dbl��׼��Ŀ = dbl��׼��Ŀ + IIf(rs���ֲ�����("���") < dbl�����, rs���ֲ�����("���"), dbl�����)
                            End If
                        End If
                    End If
                    rs���ֲ�����.MoveNext
                Loop
            Else
                '��������ܶ�
                If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
                Do Until rs��ϸ.EOF
                    If InStr(1, str��׼��Ŀ, ";" & rs��ϸ("�շ�ϸĿID") & ";") <> 0 Then
                        dbl��׼��Ŀ = dbl��׼��Ŀ + rs��ϸ("ʵ�ս��")
                    End If
                    rs��ϸ.MoveNext
                Loop
            End If
            
            '�����У����񲡡���Ⱦ������׼��Ŀ��ҽ������֧�������µĽ���ͳ�ﲿ�֣��԰�ҽ���������
            If Not bln���񲡴�Ⱦ������ Then
                str���㷽ʽ = ���ֲ�����(str��������, lng��ְ, dbl��׼��Ŀ, False)
                �����������_���� = True
                Exit Function
            Else
                g��������.����ͳ���� = g��������.����ͳ���� - dbl��׼��Ŀ
                
                '�س�����ʻ�
                If bln���񲡴�Ⱦ������ Then
                    dbl�ʻ���� = �������_����(g��������.����ID)
                    If dbl�ʻ���� > 0 Then
                        If dbl��׼��Ŀ <= dbl�ʻ���� Then
                            dbl�ʻ���� = dbl��׼��Ŀ
                        End If
                    Else
                        dbl�ʻ���� = 0
                    End If
                    g��������.�����ʻ�֧�� = dbl�ʻ����
                    g��������.ͳ�ﱨ����� = dbl��׼��Ŀ - dbl�ʻ����
                Else
                    g��������.ͳ�ﱨ����� = dbl��׼��Ŀ
                End If
            End If
        End If
    End If
    
    '������������������������������������������������������������������������������������
    '3����ȥ�����ʻ����𸶶ν���ʣ�µļ��ǽ���ͳ��Ľ��
    '3.1��������ߡ��ⶥ��
    With g��������
        
        gstrSQL = "select max(decode(A.����,'A',A.���,0)) as ������ ,max(decode(A.����,'1',A.���,0)) as ���� " & _
                  "         ,max(decode(A.����,'" & (.סԺ���� + 1) & "',A.���,0)) as ʵ������,min(A.���) as ������� " & _
                  "  from ����֧���޶� A " & _
                  "  where A.����=" & TYPE_�������� & " and A.����=" & lng���� & " and A.���=" & .���
        Call OpenRecordset(rsTemp, "�������")
                
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
'            If mCur�����Ը��� < .���� Then .���� = mCur�����Ը���
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
    
        '3.2��������ǰ�۳������߽��ԷѶΣ����ó����ε�ʵ������
        If dbl������ߺ� > 0 Then
            '�����ò��˿϶��ж�ν���
            If .���� > dbl������ߺ� Then
                '���������ߣ�Ҫ����β�ֵ
                .���� = .���� - dbl������ߺ�
            Else
                '��ǰ�����߽���Ѿ�ȫ��棬���β����ٱ�����
                .���� = 0
            End If
                
            dbl�������� = .����
        Else
            dbl�������� = .����
        End If
    End With
    g��������.ʵ������ = dbl��������
    
    '3.3��ȡ��ʵ�ʽ���ͳ��Ľ��ȸ����ʻ�֧������֧���������ߣ����µĽ���ͳ���
    dbl�ʻ���� = �������_����(g��������.����ID) - g��������.�����ʻ�֧��
    '3.3.1.1��ʹ�ø����ʻ�֧��������֧����¼��������õ��ڸ����ʻ�������˳�
    blnExit = False
    dblTemp = 0
    If dbl�ʻ���� >= 0 Then
        If g��������.����ͳ���� <= dbl�ʻ���� Then
            dblTemp = g��������.����ͳ����
            blnExit = True
        Else
            dblTemp = dbl�ʻ����
        End If
        
        '3.3.1.2����������ʻ�֧����¼
        g��������.�����ʻ�֧�� = g��������.�����ʻ�֧�� + dblTemp
        str���㷽ʽ = "�����ʻ�;" & g��������.�����ʻ�֧�� & ";0"
        If blnExit Then
            '����ͳ�������ǳ����ԷѶβ��ֵĽ���ˣ�����Ҫ����
            If g��������.ͳ�ﱨ����� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ = "", "", "|") & "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
            g��������.����ͳ���� = 0
            �����������_���� = True
            Exit Function
        End If
    End If
    g��������.����ͳ���� = g��������.����ͳ���� - dblTemp
    
    '2004-11-25 ZYB
    '��ȥ�����Ը��ν�ʣ�µİ���������
    blnExit = False
    dblTemp = 0
    If g��������.����ͳ���� >= mCur�����Ը��� Then
        g��������.����ͳ���� = g��������.����ͳ���� - mCur�����Ը���
        mCur�����Ը���_֧�� = mCur�����Ը���
    Else
        mCur�����Ը���_֧�� = g��������.����ͳ����
        g��������.����ͳ���� = 0
        blnExit = True
    End If
    If blnExit Then
        '����ͳ�������ǳ����ԷѶβ��ֵĽ���ˣ�����Ҫ����
        If g��������.ͳ�ﱨ����� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ = "", "", "|") & "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
        g��������.����ͳ���� = 0
        �����������_���� = True
        Exit Function
    End If
    
    '3.3.2.1�����ԷѶ����ڵĽ������ۼƽ���ͳ�ﲢ�˳�
    blnExit = False
    dblTemp = 0
    If Not bln������ Then
        If dbl�������� > 0 Then
            If g��������.����ͳ���� <= dbl�������� Then
                dblTemp = g��������.����ͳ����
                blnExit = True
            Else
                dblTemp = dbl��������
            End If
            '3.3.2.2�������ۼƽ���ͳ���¼
            g��������.�ۼƽ���ͳ�� = dblTemp
            If blnExit Then
                '����ͳ�������ǳ����ԷѶβ��ֵĽ���ˣ�����Ҫ����
                If g��������.ͳ�ﱨ����� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ = "", "", "|") & "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
                g��������.����ͳ���� = 0
                �����������_���� = True
                Exit Function
            End If
        End If
    End If
    
    '----ʵ�ʽ���ͳ����=��ʵ�ʷ������ý��-�����ʻ�-�ԷѶ�ʣ���
    g��������.����ͳ���� = g��������.����ͳ���� - g��������.�ۼƽ���ͳ��
    
    '3.4��ȡ�÷��õ���
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "select B.����,B.����,B.����,A.���� " & _
              "  from ����֧������ A,���շ��õ� B " & _
              "  Where A.���� =" & TYPE_�������� & " And A.���� =" & lng���� & " And A.��� =" & g��������.��� & " And A.��ְ =" & lng��ְ & " And A.����� =" & lng����� & _
              "       and A.����=B.���� and A.����=b.���� and A.����=B.���� " & _
              "  order by B.����"
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp.RecordCount = 0 Then
        MsgBox "���ڡ���Ƚ�����������ñ���ȵ�ͳ��֧��������", vbInformation, gstrSysName
        Exit Function
    End If
    
    '������������������������������������������������������������������������������������
    '4������ôν���ɱ����Ľ��
    dbl�ۼƽ��� = 0   '����ֶ��ۼƽ���ͳ��
    dbl�ѱ������ = g��������.�ۼ�ͳ�ﱨ��
    
    '���޸� -- dbl��ν���ͳ���
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
            
            If g��������.����ͳ���� > dbl���� Then
                dblTemp = 0
                
                If g��������.����ͳ���� <= dbl���� Or dbl���� = 0 Then
                    '��ʵ��ֵ����
                    dbl�ֶν��� = g��������.����ͳ���� - dbl����
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
        
                '���Ρ�����ͳ���ͳ�ﱨ��������
                lng���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
                dblTemp = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
                dbl�ۼƽ��� = dbl�ֶν��� + dbl�ۼƽ���
                    
                gcol�������.Add Array(lng����, dbl�ֶν���, dbl�ֶα���, dblTemp)
            End If
        End If
        rsTemp.MoveNext
    Loop
    str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ = "", "", "|") & "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
    �����������_���� = True
End Function

Public Function סԺ�������_����(rs������ϸ As Recordset) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."�Ƿ������޸�:0-�������޸�;1-�����޸�
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����Ҫ��NO����š�����ID��ҽ����Ŀ���롢�շ�����շ����ơ��������š���񡢲��ء��������۸񡢽�ҽ��,�Ǽ�ʱ��(����ʱ��),Ӥ����,���մ���ID
    Dim rs�㷨 As New ADODB.Recordset          '����
    Dim rsTemp As New ADODB.Recordset
    Dim rs������� As New ADODB.Recordset
    
    Dim lng���� As Long
    Dim lng��ְ As Long, lng����� As Long, lng���� As Long
    Dim dblTemp As Double, lng���� As Long
    
    Dim dbl�����  As Double ''��һ����סԺ�ռ������Ŀ������ܵõ��Ľ��
    Dim dbl�ѱ������ As Double, dbl�ۼƽ��� As Double
    Dim dbl���� As Double, dbl���� As Double, dbl�ֶν��� As Double, dbl�ֶα��� As Double
    
    Dim clsҽ�� As New clsInsure
    Dim bln�����ʻ�֧��ȫ�Է� As Boolean, bln�����ʻ�֧�������Ը� As Boolean, bln�����ʻ�֧������ As Boolean
    Dim curȫ�Է� As Currency, cur�����Ը� As Currency
    Dim bln������ As Boolean, bln�޷ⶥ�� As Boolean
    Dim dbl�ʻ����
    Dim dbl������ߺ� As Double   '�����ָ�ò�����ǰ���ʵ��ۼ�
    Dim dbl�������� As Double     '���ε�����
    Dim blnExit As Boolean          '���ڸ����ʻ��������ߣ��ԷѶΣ����򱣴���ؼ�¼���˳�
    Dim bln���ֲ����� As Boolean    '���ֲ�����
    Dim bln������Ա As Boolean, bln�������� As Boolean
    Dim str�������� As String, str��׼��Ŀ As String, dbl��׼��Ŀ As Currency, lng����ID As Long
    
    '������������������������������������������������������������������������������������
    '1����ʼ��һЩ����
    Set gcol������� = New Collection
    With g��������
        .����ID = rs������ϸ("����ID")
        
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=" & rs������ϸ("����ID")
        Call OpenRecordset(rsTemp, "�������")
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        .��ҳID = rsTemp("��ҳID")
        .��� = Int(Format(zlDatabase.Currentdate, "yyyy"))
    End With
    
    '1.1 �������˵���Ժʱ��
    gstrSQL = "select ��Ժ����,nvl(��Ժ����,to_date('3000-01-01','yyyy-MM-dd')) as ��Ժ���� " & _
              "from ������ҳ where ����ID=" & g��������.����ID & " and ��ҳID=" & g��������.��ҳID
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp("��Ժ����") = CDate("3000-01-01") Then
        g��������.��;���� = 1
    Else
        '��ʾ�ò����Ѿ���Ժ
        g��������.��;���� = 0
    End If

    '1.2 ��������סԺ�ڼ��ۼƽ���������ۼƽ���ͳ������ԷѶΣ��򰴱�����ͬ�������������ʻ�������֧����
    '�ۼƽ���ͳ����Ϊÿ��֧�����ԷѶν��
    gstrSQL = "select nvl(sum(A.�ۼƽ���ͳ��),0) as ���� " & _
              "  from ���ս����¼ A " & _
              "  Where A.����ID = " & g��������.����ID & " And A.���� = " & TYPE_�������� & " And A.���= " & g��������.���
    Call OpenRecordset(rsTemp, "�������")
    dbl������ߺ� = rsTemp("����")
    
    With g��������
        g��������.ͳ�ﱨ����� = 0
        g��������.�ۼƽ���ͳ�� = 0
        g��������.�ۼ�ͳ�ﱨ�� = 0
        g��������.ȫ�Էѽ�� = 0
        g��������.�����Ը���� = 0
        g��������.�����ʻ�֧�� = 0
        
        gstrSQL = "select A.����,A.��Ա���,A.��ְ,A.�����,Nvl(C.���,0) ����,C.���� ��������,Nvl(C.ID,0) ����ID," & _
                  "      B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ�" & _
                  " from �����ʻ� A,�ʻ������Ϣ B," & _
                  "         (Select * From ���ղ��� Where ���<>2" & _
                  "          Union " & _
                  "          Select * From ���ղ��� Where ���=2 And ���� In (" & str���ֲ� & ")) C" & _
                  " where A.����ID=B.����ID(+) and A.����=B.����(+) And A.����=C.����(+) ANd A.����ID=C.ID(+) " & _
                  "     and B.���(+)=" & .��� & " and A.����ID=" & .����ID & " and A.����=" & TYPE_��������
        Call OpenRecordset(rsTemp, "�������")
        
        '1-��ְ;2-����;3-����
        '���ݼ�������Ա�����������ߣ��ԷѶΣ�
        lng���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
        lng��ְ = IIf(IsNull(rsTemp("��ְ")), 1, rsTemp("��ְ"))
        lng���� = IIf(IsNull(rsTemp("�����")), 0, rsTemp("�����"))
        bln���ֲ����� = (rsTemp!���� = 2)
        bln������Ա = (lng��ְ = 3)
        lng����ID = rsTemp!����ID
        str�������� = IIf(IsNull(rsTemp!��������), "", rsTemp!��������)
        
        .סԺ���� = 1   '��ҽ����סԺ�����޹�
        .�ʻ��ۼ����� = IIf(IsNull(rsTemp("�ʻ������ۼ�")), 0, rsTemp("�ʻ������ۼ�"))
        .�ʻ��ۼ�֧�� = IIf(IsNull(rsTemp("�ʻ�֧���ۼ�")), 0, rsTemp("�ʻ�֧���ۼ�"))
        '�����Ը��Σ����С�ڹ涨���ߣ�������ȡ�����Ը���
        mCur�����Ը��� = IIf(IsNull(rsTemp("����ͳ���ۼ�")), 0, rsTemp("����ͳ���ۼ�"))
        .�ۼ�ͳ�ﱨ�� = IIf(IsNull(rsTemp("ͳ�ﱨ���ۼ�")), 0, rsTemp("ͳ�ﱨ���ۼ�"))
        
        gstrSQL = "select �����,nvl(������,0) as ������,nvl(�޷ⶥ��,0) as �޷ⶥ�� " & _
                " from ���������" & _
                " where ����=" & TYPE_�������� & " and nvl(����,0)=" & lng���� & _
                "       and ��ְ=" & lng��ְ & " and ����<=" & lng���� & " and (" & lng���� & "<=���� or ����=0)"
        Call OpenRecordset(rsTemp, "�������")
        If rsTemp.RecordCount = 0 Then
            MsgBox "���ڡ��������������������������õ���", vbInformation, gstrSysName
            Exit Function
        End If
        lng����� = rsTemp("�����")
        bln������ = (rsTemp("������") = 1) Or (lng��ְ <> 1)
        bln�޷ⶥ�� = (rsTemp("�޷ⶥ��") = 1)
    End With
    
    '������������������������������������������������������������������������������������
    '2����ͳ��֧����Ŀ�ϼƷ�����������
    '2.1����ʼ����¼��
    If Not clsҽ��.GetCapability(support��������ҽ����Ŀ, 0, TYPE_��������) Then
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
                    rs�������("���մ���ID") = IIf(IsNull(rs������ϸ("���մ���ID")), 0, rs������ϸ("���մ���ID"))
                    rs�������("����") = rs������ϸ("����")
                    rs�������("���") = rs������ϸ("���")
                Else
                    rs�������.MoveFirst
                    rs�������.Find "���մ���ID=" & IIf(IsNull(rs������ϸ("���մ���ID")), 0, rs������ϸ("���մ���ID"))
                    If rs�������.EOF Then
                        rs�������.AddNew
                        rs�������("���մ���ID") = IIf(IsNull(rs������ϸ("���մ���ID")), 0, rs������ϸ("���մ���ID"))
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
        gstrSQL = "select ID,�㷨,ͳ��ȶ�,��׼����,��׼����,�Ƿ�ҽ�� FROM ����֧������  where ����=" & TYPE_��������
        Call OpenRecordset(rs�㷨, "����ҽ��")
        
        dblTemp = 0
        If rs�������.RecordCount > 0 Then rs�������.MoveFirst
        Do Until rs�������.EOF
            
            rs�㷨.Filter = "ID=" & rs�������("���մ���ID")
            If rs�㷨.RecordCount > 0 Then
                If rs�㷨("�Ƿ�ҽ��") = 1 Then
                    '�㷨:1-�ܶ������Ŀ��2-סԺ�պ˶���Ŀ
                    If rs�㷨("�㷨") = 1 Then
                        If rs�㷨("ͳ��ȶ�") = 0 Then
                            curȫ�Է� = curȫ�Է� + rs�������("���")
                        Else
                            dblTemp = dblTemp + rs�������("���") * rs�㷨("ͳ��ȶ�") / 100
                        End If
                    Else
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
                    End If
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
        g��������.�����Ը���� = g��������.�������ý�� - curȫ�Է� - dblTemp
    Else
        '��������ܶ�
        If rs������ϸ.RecordCount <> 0 Then rs������ϸ.MoveFirst
        Do Until rs������ϸ.EOF
            dblTemp = dblTemp + rs������ϸ("���")
            rs������ϸ.MoveNext
        Loop
        g��������.�������ý�� = dblTemp
        g��������.����ͳ���� = dblTemp
    End If
    
    '�����������Ա�����﷢�������з��ã��۳������ʻ��⣬ȫ����ͳ��ҽ�ƻ���֧��
    If bln������Ա Then
        dblTemp = �������_����(g��������.����ID)
        dblTemp = IIf(dblTemp > 0, dblTemp, 0)
        dblTemp = IIf(g��������.����ͳ���� > dblTemp, dblTemp, g��������.����ͳ����)
        g��������.�����ʻ�֧�� = dblTemp
        g��������.ͳ�ﱨ����� = g��������.����ͳ���� - dblTemp
        סԺ�������_���� = "ҽ������;" & g��������.ͳ�ﱨ����� & ";0|�����ʻ�;" & g��������.�����ʻ�֧�� & ";0"
        Exit Function
    End If
    
    '��������ֲ�����
    If bln���ֲ����� Then
        
        Dim rs���ֲ����� As New ADODB.Recordset
        str��׼��Ŀ = ""
        dbl��׼��Ŀ = 0
        bln�������� = (InStr(1, ",��������,", "," & str�������� & ",") <> 0)
        
        If Not bln�������� Then
            סԺ�������_���� = ���ֲ�����(str��������, lng��ְ, 0, True)
            Exit Function
        End If
        
        If bln�������� Then
            'ҩƷ�ѡ������ѡ�Ѫ����ͳ�����֧��50%
            g��������.����ͳ���� = 0
            With rs������ϸ
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    If InStr(1, ",5,6,7,F,K", "," & !�շ���� & ",") <> 0 And !������Ŀ�� = 1 Then
                        g��������.����ͳ���� = g��������.����ͳ���� + (!��� * 0.5)
                    End If
                    .MoveNext
                Loop
            End With
            g��������.ͳ�ﱨ����� = g��������.����ͳ����
            סԺ�������_���� = "ҽ������;" & g��������.ͳ�ﱨ����� & ";0|�����ʻ�;0;0"
            Exit Function
        Else
            '������׼��Ŀ����ͳ����ܶ�
            With rsTemp
                gstrSQL = "Select �շ�ϸĿID From ������׼��Ŀ Where ����ID=" & lng����ID
                Call OpenRecordset(rsTemp, "�������")
                
                Do While Not .EOF
                    str��׼��Ŀ = str��׼��Ŀ & ";" & !�շ�ϸĿID
                    .MoveNext
                Loop
                str��׼��Ŀ = str��׼��Ŀ & ";"
            End With
            
            If Not clsҽ��.GetCapability(support��������ҽ����Ŀ, 0, TYPE_��������) Then
                Set rs���ֲ����� = New ADODB.Recordset
                With rs���ֲ�����
                    If .State = adStateOpen Then .Close
                    .Fields.Append "���մ���ID", adDouble, 18, adFldIsNullable
                    .Fields.Append "����", adDouble, 8, adFldIsNullable
                    .Fields.Append "���", adDouble, 18, adFldIsNullable
                    .Fields.Append "ͳ����", adDouble, 18, adFldIsNullable
                    
                    .CursorLocation = adUseClient
                    .Open , , adOpenStatic, adLockOptimistic
                End With
            
                If rs������ϸ.RecordCount <> 0 Then rs������ϸ.MoveFirst
                Do Until rs������ϸ.EOF
                'װ����д���¼��������������ʹ��
                    If rs������ϸ("������Ŀ��") = 1 And InStr(1, str��׼��Ŀ, ";" & rs������ϸ("�շ�ϸĿID") & ";") <> 0 Then
                        If rs���ֲ�����.RecordCount = 0 Then
                            rs���ֲ�����.AddNew
                            rs���ֲ�����("���մ���ID") = IIf(IsNull(rs������ϸ("���մ���ID")), 0, rs������ϸ("���մ���ID"))
                            rs���ֲ�����("����") = rs������ϸ("����")
                            rs���ֲ�����("���") = rs������ϸ("���")
                        Else
                            rs���ֲ�����.MoveFirst
                            rs���ֲ�����.Find "���մ���ID=" & IIf(IsNull(rs������ϸ("���մ���ID")), 0, rs������ϸ("���մ���ID"))
                            If rs���ֲ�����.EOF Then
                                rs���ֲ�����.AddNew
                                rs���ֲ�����("���մ���ID") = IIf(IsNull(rs������ϸ("���մ���ID")), 0, rs������ϸ("���մ���ID"))
                                rs���ֲ�����("����") = rs������ϸ("����")
                                rs���ֲ�����("���") = rs������ϸ("���")
                            Else
                                rs���ֲ�����("����") = rs���ֲ�����("����") + rs������ϸ("����")
                                rs���ֲ�����("���") = rs���ֲ�����("���") + rs������ϸ("���")
                            End If
                        End If
                        rs���ֲ�����.Update
                    End If
                    rs������ϸ.MoveNext
                Loop
                
                '2.2���������ͳ����
                If rs�㷨.RecordCount <> 0 Then rs�㷨.MoveFirst
                If rs���ֲ�����.RecordCount > 0 Then rs���ֲ�����.MoveFirst
                Do Until rs���ֲ�����.EOF
                    
                    rs�㷨.Filter = "ID=" & rs���ֲ�����("���մ���ID")
                    If rs�㷨.RecordCount > 0 Then
                        If rs�㷨("�Ƿ�ҽ��") = 1 Then
                            '�㷨:1-�ܶ������Ŀ��2-סԺ�պ˶���Ŀ
                            If rs�㷨("�㷨") = 1 Then
                                If rs�㷨("ͳ��ȶ�") = 0 Then
                                    curȫ�Է� = curȫ�Է� + rs���ֲ�����("���")
                                Else
                                    dbl��׼��Ŀ = dbl��׼��Ŀ + rs���ֲ�����("���") * rs�㷨("ͳ��ȶ�") / 100
                                End If
                            Else
                                If Val(rs���ֲ�����("����")) > Val(rs�㷨("��׼����")) Then
                                    '���סԺ�ճ�����׼��������ô�������� ��׼����*��׼���� +  (����-��׼����)*ͳ��ȶ�
                                    '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                                    dbl����� = rs�㷨("��׼����") * rs�㷨("��׼����") + _
                                        (rs���ֲ�����("����") - IIf(rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0, 0, rs�㷨("��׼����"))) * rs�㷨("ͳ��ȶ�")
                                Else
                                    '���סԺ�յ�����׼��������ô�������� ����*��׼���� ���� ����*ͳ��ȶ�
                                    '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                                    If rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0 Then
                                        dbl����� = rs���ֲ�����("����") * rs�㷨("ͳ��ȶ�")
                                    Else
                                        dbl����� = rs���ֲ�����("����") * rs�㷨("��׼����")
                                    End If
                                End If
                                
                                '�ܽ��������С����ȡȫ��������ֻ�����
                                dbl��׼��Ŀ = dbl��׼��Ŀ + IIf(rs���ֲ�����("���") < dbl�����, rs���ֲ�����("���"), dbl�����)
                            End If
                        End If
                    End If
                    rs���ֲ�����.MoveNext
                Loop
            Else
                '��������ܶ�
                If rs������ϸ.RecordCount <> 0 Then rs������ϸ.MoveFirst
                Do Until rs������ϸ.EOF
                    If InStr(1, str��׼��Ŀ, ";" & rs������ϸ("�շ�ϸĿID") & ";") <> 0 Then
                        dbl��׼��Ŀ = dbl��׼��Ŀ + rs������ϸ("ʵ�ս��")
                    End If
                    rs������ϸ.MoveNext
                Loop
            End If
        End If
    End If
    
    '�ƻ���������׼��Ŀ��ҽ������֧�������µĽ���ͳ�ﲿ�֣��԰�ҽ���������
    g��������.����ͳ���� = g��������.����ͳ���� - dbl��׼��Ŀ
    g��������.ͳ�ﱨ����� = dbl��׼��Ŀ
    
    '������������������������������������������������������������������������������������
    '3����ȥ�����ʻ����𸶶ν���ʣ�µļ��ǽ���ͳ��Ľ��
    '3.1��������ߡ��ⶥ��
    With g��������
        
        gstrSQL = "select max(decode(A.����,'A',A.���,0)) as ������ ,max(decode(A.����,'1',A.���,0)) as ���� " & _
                  "         ,max(decode(A.����,'" & (.סԺ���� + 1) & "',A.���,0)) as ʵ������,min(A.���) as ������� " & _
                  "  from ����֧���޶� A " & _
                  "  where A.����=" & TYPE_�������� & " and A.����=" & lng���� & " and A.���=" & .���
        Call OpenRecordset(rsTemp, "�������")
                
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
'            If mCur�����Ը��� < .���� Then .���� = mCur�����Ը���
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
    
        '3.2��������ǰ�۳������߽��ԷѶΣ����ó����ε�ʵ������
        If dbl������ߺ� > 0 Then
            '�����ò��˿϶��ж�ν���
            If .���� > dbl������ߺ� Then
                '���������ߣ�Ҫ����β�ֵ
                .���� = .���� - dbl������ߺ�
            Else
                '��ǰ�����߽���Ѿ�ȫ��棬���β����ٱ�����
                .���� = 0
            End If
                
            dbl�������� = .����
        Else
            dbl�������� = .����
        End If
    End With
    g��������.ʵ������ = dbl��������
    
    '3.3��ȡ��ʵ�ʽ���ͳ��Ľ��ȸ����ʻ�֧������֧���������ߣ����µĽ���ͳ���
    dbl�ʻ���� = �������_����(g��������.����ID)
    '3.3.1.1��ʹ�ø����ʻ�֧��������֧����¼��������õ��ڸ����ʻ�������˳�
    blnExit = False
    dblTemp = 0
    If dbl�ʻ���� >= 0 Then
        If g��������.����ͳ���� <= dbl�ʻ���� Then
            dblTemp = g��������.����ͳ����
            blnExit = True
        Else
            dblTemp = dbl�ʻ����
        End If
        
        '3.3.1.2����������ʻ�֧����¼
        g��������.�����ʻ�֧�� = dblTemp
        סԺ�������_���� = "�����ʻ�;" & g��������.�����ʻ�֧�� & ";0"
        If blnExit Then
            '����ͳ�������ǳ����ԷѶβ��ֵĽ���ˣ�����Ҫ����
            If g��������.ͳ�ﱨ����� <> 0 Then סԺ�������_���� = סԺ�������_���� & IIf(סԺ�������_���� = "", "", "|") & "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
            g��������.����ͳ���� = 0
            Exit Function
        End If
    End If
    g��������.����ͳ���� = g��������.����ͳ���� - g��������.�����ʻ�֧��
    
    '2004-11-25 ZYB
    '��ȥ�����Ը��ν�ʣ�µİ���������
    blnExit = False
    dblTemp = 0
    If g��������.����ͳ���� >= mCur�����Ը��� Then
        g��������.����ͳ���� = g��������.����ͳ���� - mCur�����Ը���
        mCur�����Ը���_֧�� = mCur�����Ը���
    Else
        mCur�����Ը���_֧�� = g��������.����ͳ����
        g��������.����ͳ���� = 0
        blnExit = True
    End If
    If blnExit Then
        '����ͳ�������ǳ����ԷѶβ��ֵĽ���ˣ�����Ҫ����
        If g��������.ͳ�ﱨ����� <> 0 Then סԺ�������_���� = סԺ�������_���� & IIf(סԺ�������_���� = "", "", "|") & "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
        g��������.����ͳ���� = 0
        Exit Function
    End If
    
    '3.3.2.1�����ԷѶ����ڵĽ������ۼƽ���ͳ�ﲢ�˳�
    blnExit = False
    dblTemp = 0
    g��������.�ۼƽ���ͳ�� = 0
    If Not bln������ Then
        If dbl�������� > 0 Then
            If g��������.����ͳ���� <= dbl�������� Then
                dblTemp = g��������.����ͳ����
                blnExit = True
            Else
                dblTemp = dbl��������
            End If
            '3.3.2.2�������ۼƽ���ͳ���¼
            g��������.�ۼƽ���ͳ�� = dblTemp
            If blnExit Then
                '����ͳ�������ǳ����ԷѶβ��ֵĽ���ˣ�����Ҫ����
                If g��������.ͳ�ﱨ����� <> 0 Then סԺ�������_���� = סԺ�������_���� & IIf(סԺ�������_���� = "", "", "|") & "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
                g��������.����ͳ���� = 0
                Exit Function
            End If
        End If
    End If
    
    '----ʵ�ʽ���ͳ����=��ʵ�ʷ������ý��-�����ʻ�-�ԷѶ�ʣ���
    g��������.����ͳ���� = g��������.����ͳ���� - g��������.�ۼƽ���ͳ��
    
    '3.4��ȡ�÷��õ���
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "select B.����,B.����,B.����,A.���� " & _
              "  from ����֧������ A,���շ��õ� B " & _
              "  Where A.���� =" & TYPE_�������� & " And A.���� =" & lng���� & " And A.��� =" & g��������.��� & " And A.��ְ =" & lng��ְ & " And A.����� =" & lng����� & _
              "       and A.����=B.���� and A.����=b.���� and A.����=B.���� " & _
              "  order by B.����"
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp.RecordCount = 0 Then
        MsgBox "���ڡ���Ƚ�����������ñ���ȵ�ͳ��֧��������", vbInformation, gstrSysName
        Exit Function
    End If
    
    '������������������������������������������������������������������������������������
    '4������ôν���ɱ����Ľ��
    dbl�ۼƽ��� = 0   '����ֶ��ۼƽ���ͳ��
    dbl�ѱ������ = g��������.�ۼ�ͳ�ﱨ��
    
    '���޸� -- dbl��ν���ͳ���
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
            
            If g��������.����ͳ���� > dbl���� Then
                dblTemp = 0
                
                If g��������.����ͳ���� <= dbl���� Or dbl���� = 0 Then
                    '��ʵ��ֵ����
                    dbl�ֶν��� = g��������.����ͳ���� - dbl����
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
        
                '���Ρ�����ͳ���ͳ�ﱨ��������
                lng���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
                dblTemp = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
                dbl�ۼƽ��� = dbl�ֶν��� + dbl�ۼƽ���
                    
                gcol�������.Add Array(lng����, dbl�ֶν���, dbl�ֶα���, dblTemp)
            End If
        End If
        rsTemp.MoveNext
    Loop
    סԺ�������_���� = סԺ�������_���� & IIf(סԺ�������_���� = "", "", "|") & "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
End Function

Public Function סԺ����_����(lng����ID As Long) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    Dim cur�����ʻ� As Currency
    Dim var������� As Variant
On Error GoTo ErrH
    With g��������
        mCur�����Ը��� = mCur�����Ը��� - mCur�����Ը���_֧��
        cur�����ʻ� = .�����ʻ�֧��
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & .����ID & "," & TYPE_�������� & "," & .��� & "," & _
            .�ʻ��ۼ����� & "," & .�ʻ��ۼ�֧�� + cur�����ʻ� & "," & mCur�����Ը��� & "," & _
            .�ۼ�ͳ�ﱨ�� + .ͳ�ﱨ����� & "," & .��ҳID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�������� & "," & .����ID & "," & _
            .��� & "," & .�ʻ��ۼ����� & "," & .�ʻ��ۼ�֧�� & "," & .�ۼƽ���ͳ�� & "," & _
            .�ۼ�ͳ�ﱨ�� & "," & .��ҳID & "," & .���� & "," & .�ⶥ�� & "," & .ʵ������ & "," & _
            .�������ý�� & "," & .ȫ�Էѽ�� & "," & .�����Ը���� & "," & .����ͳ���� & "," & .ͳ�ﱨ����� & ",0," & _
            .�����Ը���� & "," & cur�����ʻ� & ",NULL," & .��ҳID & "," & .��;���� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        For Each var������� In gcol�������
            '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
            gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
                var�������(0) & "," & var�������(1) & "," & var�������(2) & "," & var�������(3) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        Next
    End With
    
    סԺ����_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
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
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "Select * " & _
              "  From ���ս����¼ Where ����=2 and ��¼ID='" & lng����ID & "'"
    Call OpenRecordset(rsTemp, "����ҽ��")
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "�ò��˵�ҽ���������ݶ�ʧ���������ϡ�"
        Exit Function
    End If
    
    gstrSQL = "select B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ� " & _
              " from �����ʻ� A,�ʻ������Ϣ B " & _
              " where A.����ID=B.����ID(+) and A.����=B.����(+) and B.���(+)=" & Year(zlDatabase.Currentdate) & " and A.����ID=" & rsTemp("����ID") & " and A.����=" & TYPE_��������
    Call OpenRecordset(rs�ʻ�, "����ҽ��")
    
    If rs�ʻ�.EOF = False Then
        lngסԺ���� = IIf(IsNull(rs�ʻ�("סԺ�����ۼ�")), 0, rs�ʻ�("סԺ�����ۼ�"))
        cur�ʻ����� = IIf(IsNull(rs�ʻ�("�ʻ������ۼ�")), 0, rs�ʻ�("�ʻ������ۼ�"))
        cur�ʻ�֧�� = IIf(IsNull(rs�ʻ�("�ʻ�֧���ۼ�")), 0, rs�ʻ�("�ʻ�֧���ۼ�"))
        cur�ۼƽ���ͳ�� = IIf(IsNull(rs�ʻ�("����ͳ���ۼ�")), 0, rs�ʻ�("����ͳ���ۼ�"))
        cur�ۼ�ͳ�ﱨ�� = IIf(IsNull(rs�ʻ�("ͳ�ﱨ���ۼ�")), 0, rs�ʻ�("ͳ�ﱨ���ۼ�"))
    End If
    
    '���˴������ݱ���������������ݱ������һ������
    '��˾Ͳ���Ҫ�������������
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & TYPE_�������� & "," & rsTemp("���") & "," & _
        cur�ʻ����� & "," & cur�ʻ�֧�� - rsTemp("�����ʻ�֧��") & "," & cur�ۼƽ���ͳ�� & "," & _
        0 & "," & lngסԺ���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '�������ݣ������˼����ۼ�
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�������� & "," & rsTemp("����ID") & "," & _
        rsTemp("���") & "," & cur�ʻ����� & "," & cur�ʻ�֧�� - rsTemp("�����ʻ�֧��") & "," & rsTemp("�ۼƽ���ͳ��") * -1 & "," & _
        cur�ۼ�ͳ�ﱨ�� & "," & lngסԺ���� & "," & rsTemp("����") * -1 & "," & rsTemp("�ⶥ��") & "," & rsTemp("ʵ������") * -1 & "," & _
        rsTemp("�������ý��") * -1 & "," & rsTemp("ȫ�Ը����") * -1 & "," & rsTemp("�����Ը����") * -1 & "," & rsTemp("����ͳ����") * -1 & "," & _
        rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") * -1 & "," & rsTemp("�����ʻ�֧��") * -1 & ",''," & _
        IIf(IsNull(rsTemp("��ҳID")), "null", rsTemp("��ҳID")) & "," & rsTemp("��;����") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "select ����,����ͳ����,ͳ�ﱨ�����,���� from ���ս������ where ����ID=" & lng����ID
    Call OpenRecordset(rs�������, "����ҽ��")
    
    Do Until rs�������.EOF
        '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
        gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
            rs�������("����") & "," & rs�������("����ͳ����") * -1 & "," & rs�������("ͳ�ﱨ�����") * -1 & "," & rs�������("����") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
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

Public Function BuildPatiInfo_����(ByVal bytType As Byte, ByVal strInfo As String, ByVal lng����ID As Long) As Long
'���ܣ����������ʻ���Ϣ
'������bytType=0-����,1-סԺ
'      strInfo='0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
'      8����;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(1,2,3);15����֤��;16�����;17�Ҷȼ�
'      18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23��������
'���أ�����ID
    Dim rsPati As ADODB.Recordset, str��λ���� As String, lng���� As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, curDate As Date
    Dim lng���� As Long, array��Ϣ As Variant
    
    On Error GoTo errHandle
    If Len(Trim(strInfo)) <> 0 Then
        curDate = zlDatabase.Currentdate
        array��Ϣ = Split(strInfo, ";")
        '�ӵ�7��������ȡ����λ����
        If array��Ϣ(7) Like "*(*" Then
            str��λ���� = Split(array��Ϣ(7), "(")(UBound(Split(array��Ϣ(7), "(")))
            str��λ���� = Mid(str��λ����, 1, Len(str��λ����) - 1)
        End If
        'ȡ����
        If IsDate(array��Ϣ(5)) Then
            lng���� = Int(curDate - CDate(array��Ϣ(5))) / 365
        End If
        
        lng���� = Val(array��Ϣ(8))
        #If gverControl < 6 Then
            '�ʻ�Ψһ������,����,ҽ����
            strSQL = "Select A.*,B.ҽ���� From ������Ϣ A," & _
                " (Select * From �����ʻ�" & _
                " Where ����=" & TYPE_�������� & _
                " And ҽ����='" & CStr(array��Ϣ(1)) & "'" & _
                " And ����=" & lng���� & ") B" & _
                " Where " & IIf(lng����ID = 0, "A.����ID=B.����ID", "A.����ID=B.����ID(+) and A.����ID=" & lng����ID) '���ܲ���ID�Ѿ�ȷ��
        #Else
            '�ʻ�Ψһ������,����,ҽ����
            strSQL = "Select A.����id, A.�����, A.סԺ��, A.���￨��, A.����֤��, A.�ѱ�, A.ҽ�Ƹ��ʽ, A.����, A.�Ա�, A.����, A.��������, A.�����ص�, A.���֤��, A.����֤��, A.���, A.ְҵ, A.����, A.����, A.����, A.ѧ��, A.����״��, A.��ͥ��ַ," & vbNewLine & _
                "      A.��ͥ�绰, A.��ͥ��ַ�ʱ� As �����ʱ�, A.�໤��, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ, A.��ϵ�˵绰, A.��ͬ��λid, A.������λ, A.��λ�绰, A.��λ�ʱ�, A.��λ������, A.��λ�ʺ�, A.������, A.������, A.��������, A.����ʱ��, A.����״̬," & vbNewLine & _
                "      A.��������, A.סԺ����, A.��ǰ����id, A.��ǰ����id, A.��ǰ����, A.��Ժʱ��, A.��Ժʱ��, A.��Ժ, A.Ic����, A.������, A.ҽ����, A.����, A.��ѯ����, A.�Ǽ�ʱ��, A.ͣ��ʱ��, A.����," & vbNewLine & _
                "      B.ҽ���� From ������Ϣ A," & _
                " (Select * From �����ʻ�" & _
                " Where ����=" & TYPE_�������� & _
                " And ҽ����='" & CStr(array��Ϣ(1)) & "'" & _
                " And ����=" & lng���� & ") B" & _
                " Where " & IIf(lng����ID = 0, "A.����ID=B.����ID", "A.����ID=B.����ID(+) and A.����ID=" & lng����ID) '���ܲ���ID�Ѿ�ȷ��
        #End If
        Set rsPati = New ADODB.Recordset
        rsPati.CursorLocation = adUseClient
        Call OpenRecordset(rsPati, "����ҽ��", strSQL)
        
        If rsPati.EOF Then
            '�ޱ����ʻ�����Ϊû�в�����Ϣ
            If lng����ID = 0 Then lng����ID = GetNextNO(1)
            strSQL = "zl_������Ϣ_Insert(" & lng����ID & ",NULL,NULL,NULL," & _
                "'" & array��Ϣ(3) & "','" & array��Ϣ(4) & "'," & IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
                "To_Date('" & Format(array��Ϣ(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "NULL,'" & array��Ϣ(6) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
                "NULL,NULL,NULL,NULL,NULL,NULL,'" & array��Ϣ(7) & "',NULL,NULL,NULL," & _
                "NULL,NULL,NULL," & TYPE_�������� & "," & _
                "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call SQLTest(App.ProductName, "����ҽ��", strSQL)
            gcnOracle.Execute strSQL, , adCmdStoredProc
            Call SQLTest
        Else
            '�в�����Ϣ�ͱ����ʻ���Ϣ
            If lng����ID = 0 Then lng����ID = rsPati!����ID
            strSQL = "zl_������Ϣ_Update(" & _
                lng����ID & "," & IIf(IsNull(rsPati!�����), "NULL", rsPati!�����) & "," & _
                IIf(IsNull(rsPati!סԺ��), "NULL", rsPati!סԺ��) & ",'" & IIf(IsNull(rsPati!�ѱ�), "", rsPati!�ѱ�) & "'," & _
                "'" & IIf(IsNull(rsPati!ҽ�Ƹ��ʽ), "", rsPati!ҽ�Ƹ��ʽ) & "'," & _
                "'" & array��Ϣ(3) & "','" & array��Ϣ(4) & "'," & IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
                "To_Date('" & Format(array��Ϣ(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "'" & IIf(IsNull(rsPati!�����ص�), "", rsPati!�����ص�) & "','" & array��Ϣ(6) & "'," & _
                "'" & IIf(IsNull(rsPati!���), "", rsPati!���) & "','" & IIf(IsNull(rsPati!ְҵ), "", rsPati!ְҵ) & "'," & _
                "'" & IIf(IsNull(rsPati!����), "", rsPati!����) & "','" & IIf(IsNull(rsPati!����), "", rsPati!����) & "'," & _
                "'" & IIf(IsNull(rsPati!ѧ��), "", rsPati!ѧ��) & "','" & IIf(IsNull(rsPati!����״��), "", rsPati!����״��) & "'," & _
                "'" & IIf(IsNull(rsPati!��ͥ��ַ), "", rsPati!��ͥ��ַ) & "','" & IIf(IsNull(rsPati!��ͥ�绰), "", rsPati!��ͥ�绰) & "'," & _
                "'" & IIf(IsNull(rsPati!�����ʱ�), "", rsPati!�����ʱ�) & "','" & IIf(IsNull(rsPati!��ϵ������), "", rsPati!��ϵ������) & "'," & _
                "'" & IIf(IsNull(rsPati!��ϵ�˹�ϵ), "", rsPati!��ϵ�˹�ϵ) & "','" & IIf(IsNull(rsPati!��ϵ�˵�ַ), "", rsPati!��ϵ�˵�ַ) & "'," & _
                "'" & IIf(IsNull(rsPati!��ϵ�˵绰), "", rsPati!��ϵ�˵绰) & "'," & IIf(IsNull(rsPati!��ͬ��λID), "NULL", rsPati!��ͬ��λID) & "," & _
                "'" & array��Ϣ(7) & "','" & IIf(IsNull(rsPati!��λ�绰), "", rsPati!��λ�绰) & "'," & _
                "'" & IIf(IsNull(rsPati!��λ�ʱ�), "", rsPati!��λ�ʱ�) & "','" & IIf(IsNull(rsPati!��λ������), "", rsPati!��λ������) & "'," & _
                "'" & IIf(IsNull(rsPati!��λ�ʺ�), "", rsPati!��λ�ʺ�) & "','" & IIf(IsNull(rsPati!������), "", rsPati!������) & "'," & _
                "" & IIf(IsNull(rsPati!������), "NULL", rsPati!������) & "," & TYPE_�������� & ")"
            Call SQLTest(App.ProductName, "����ҽ��", strSQL)
            gcnOracle.Execute strSQL, , adCmdStoredProc
            Call SQLTest
        End If
        
        '�������±����ʻ���Ϣ(�Զ�)
        strSQL = "zl_�����ʻ�_insert(" & lng����ID & "," & TYPE_�������� & "," & _
            lng���� & "," & _
            "'" & IIf(array��Ϣ(0) = "-1", array��Ϣ(1), array��Ϣ(0)) & "'," & _
            "'" & array��Ϣ(1) & "'," & _
            "'" & array��Ϣ(2) & "'," & _
            "'" & array��Ϣ(9) & "'," & _
            "'" & array��Ϣ(15) & "'," & _
            "'" & array��Ϣ(10) & "'," & _
            "'" & str��λ���� & "'," & _
            Val(array��Ϣ(11)) & "," & _
            Val(array��Ϣ(12)) & "," & _
            IIf(Val(array��Ϣ(13)) = 0, "NULL", Val(array��Ϣ(13))) & "," & _
            IIf(Val(array��Ϣ(14)) = 0, 1, Val(array��Ϣ(14))) & "," & _
            IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
            "'" & array��Ϣ(17) & "'," & _
            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call SQLTest(App.ProductName, "����ҽ��", strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
        
        '���������ʻ������Ϣ(�Զ�)
        strSQL = "zl_�ʻ������Ϣ_Insert(" & lng����ID & "," & TYPE_�������� & "," & Year(curDate) & "," & _
            Val(array��Ϣ(18)) & "," & Val(array��Ϣ(19)) & "," & _
            Val(array��Ϣ(20)) & "," & 0 & "," & Val(array��Ϣ(21)) & ")"
        Call SQLTest(App.ProductName, "����ҽ��", strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
    End If
    BuildPatiInfo_���� = lng����ID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ���ֲ�����(ByVal str�������� As String, ByVal lng��ְ As Long, ByVal dbl��׼��Ŀ As Currency, Optional ByVal blnסԺ As Boolean = False) As String
    Dim dbl�����ʻ� As Currency
    Dim dbl����֧�� As Currency, dbl�ʻ�֧�� As Currency, dbl��λ֧�� As Currency
    '���صĴ���ʽ��סԺ�������һ��
    '������Ҫ�޸Ľ���ͳ���ͳ�ﱨ�����
    
    dbl�����ʻ� = �������_����(g��������.����ID)
    dbl�����ʻ� = IIf(dbl�����ʻ� > 0, dbl�����ʻ�, 0)
    dbl�ʻ�֧�� = 0: dbl����֧�� = 0: dbl��λ֧�� = 0
    
    Select Case str��������
    Case "����", "��Ⱦ��"
        '��סԺ����ְ�����ݣ��۳������ʻ������ȫ��
        '�������ְ�����ݣ��۳������ʻ���������׼��Ŀ������׼��Ŀ��ģ��԰�ҽ���������
        If blnסԺ Then
            If dbl�����ʻ� > 0 Then
                If g��������.����ͳ���� <= dbl�����ʻ� Then
                    dbl�ʻ�֧�� = g��������.����ͳ����
                Else
                    dbl�ʻ�֧�� = dbl�����ʻ�
                End If
            Else
                dbl�ʻ�֧�� = dbl�����ʻ�
            End If
            dbl����֧�� = g��������.����ͳ���� - dbl�ʻ�֧��
        End If
    Case "ְҵ��", "��֢"
        '��סԺ������۳������ʻ������ȫ��
        If dbl�����ʻ� > 0 Then
            If g��������.����ͳ���� <= dbl�����ʻ� Then
                dbl�ʻ�֧�� = g��������.����ͳ����
            Else
                dbl�ʻ�֧�� = dbl�����ʻ�
            End If
        Else
            dbl�ʻ�֧�� = dbl�����ʻ�
        End If
        dbl����֧�� = g��������.����ͳ���� - dbl�ʻ�֧��
    Case "����", "�ƻ�����"
        '��סԺ�����ȫ�⣬���۳������ʻ�
        dbl����֧�� = g��������.����ͳ����
    Case "��λ֧��"
        dbl��λ֧�� = g��������.�������ý��
    Case "��������"
        '��סԺ����ҩ�ѡ������ѡ�Ѫ����ҽ������֧��50%�����µ��Է�
        '�������ҩ����ҽ������֧��50%�����µ��Է�
    End Select
    
    dbl����֧�� = Val(Format(dbl����֧��, "#####0.00;-#####0.00;0;"))
    dbl�ʻ�֧�� = Val(Format(dbl�ʻ�֧��, "#####0.00;-#####0.00;0;"))
    dbl��λ֧�� = Val(Format(dbl��λ֧��, "#####0.00;-#####0.00;0;"))
    g��������.ͳ�ﱨ����� = dbl����֧��
    g��������.�����ʻ�֧�� = dbl�ʻ�֧��
    
    '��Ϸ��ش�
    ���ֲ����� = "ҽ������;" & g��������.ͳ�ﱨ����� & ";0|�����ʻ�;" & g��������.�����ʻ�֧�� & ";0"
    If dbl��λ֧�� <> 0 Then ���ֲ����� = ���ֲ����� & "|��λ֧��;" & dbl��λ֧�� & ";0"
End Function


