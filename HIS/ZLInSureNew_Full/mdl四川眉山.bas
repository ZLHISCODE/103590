Attribute VB_Name = "mdl�Ĵ�üɽ"
Option Explicit
'����֧����ͳ�������סԺ֧����ͳ����ۼ�
'--------------------��ҵ��˾�������--------------------
'���
'   һ��-       �����ʻ�ʹ����󣬲����ٱ���
'   ���ز�-     �ȼ������ʻ������²���ͳ�����50%��������޶�Ϊ2000Ԫ��ʵ�ʱ�������޶�Ϊ1000Ԫ��
'   ����-       ͳ�����֧��
'סԺ��
'   һ��-       ��������������
'--------------------  �����������  --------------------
'�����һ���⣬�����������ʻ�����
'   һ��-       �����ʻ�ʹ����󣬲����ٱ���
'   ���ز�-     ͳ�����50%��������޶�Ϊ4000Ԫ��ʵ�ʱ�������޶�Ϊ2000Ԫ��
'   �˲о���-   ͳ�����80%
'   ��������-   ͳ�����֧��
'   ������Ա-   ͳ�����֧��
'�������������޶100%
'סԺ��
'   һ��-       ��������������

Private Type ComInfo_üɽ
    ����ID As Long
    ���� As Long
    ���� As String
    ҽ���� As String
    ��Ⱥ As String
    ����ID As Long
    �������� As String
    סԺ���� As Integer
    סԺ���� As Integer
    ���� As Currency
    �������� As Currency
    �����ܶ� As Currency
    �ʻ���� As Currency
    ����ͳ�� As Currency
    ʵ�ʱ��� As Currency            'ʵ�ʱ�������������100%��ͳ��֧���Ļ��ܽ��
    ����ʵ�ʱ��� As Currency        '����ʵ�ʱ������ֵ�ͳ����ⲿ�ֽ�����ֵ�����
    ȫ�Ը� As Currency
    �����Ը� As Currency
    ͳ��֧�� As Currency
    ͳ���Ը� As Currency
    �ʻ�֧�� As Currency
    ����޶� As Currency            '�޶ʹ���ڵ�ǰ��Ӧ�Ĳ��ֻ���Ⱥ
    �ѱ������ As Currency          '������Ѿ��������
    �������� As Single
End Type
Public gstrʵ�ʱ�������_���� As String                           '�����û�����ı������������ڼ���
Public gComInfo_üɽ As ComInfo_üɽ
Public rs����_���� As New ADODB.Recordset                   '������ܣ����ڼ���ͳ��
Public rs֧������_���� As New ADODB.Recordset               '����֧������
Public rs�ֵ�֧��_���� As New ADODB.Recordset               '�ֵ�֧����ϸ�����ڱ����������
Public Const gstrFormat_üɽ As String = "#####0.0;-#####0.0; ;"

Public Function ��ݱ�ʶ_üɽ(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    ��ݱ�ʶ_üɽ = frmIdentifyüɽ.ShowCard(bytType, lng����ID)
End Function

Public Function ҽ����ʼ��_üɽ() As Boolean
    ҽ����ʼ��_üɽ = True
End Function

Public Function �������_üɽ(ByVal strSelfNo As String) As Currency
    '����: ֱ�Ӷ������ڽ��
    '����: �Ƿ����
    '����: ���ظ����ʻ����
    Dim rsAccount As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = " Select Nvl(�ʻ����,0) �ʻ���� From �����ʻ� " & _
              " Where ����=[1] And ҽ����=[2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "���ظ����ʻ����", TYPE_�Ĵ�üɽ, strSelfNo)
    
    �������_üɽ = rsAccount!�ʻ����
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �����������_üɽ(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    Dim curTotal As Currency, cur�����ʻ� As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    
    '�ȼ��������ͳ���������ܣ�
    Call Calc_����ͳ��(rs��ϸ, True)
    
    '�������ġ���Ⱥ�����ּ���ͳ�ﱨ����ע���޶�Ĵ���
    Call Calc_���ﱨ������_����(True, True)
    
    '����������ʻ�֧����ͳ�ﱨ���赽ҽ�����Ĵ���
'    If gComInfo_üɽ.ͳ��֧�� <> 0 Then str���㷽ʽ = "ҽ������;" & gComInfo_üɽ.ͳ��֧�� & ";0"
    If gComInfo_üɽ.�ʻ�֧�� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ = "", "", "|") & "�����ʻ�;" & gComInfo_üɽ.�ʻ�֧�� & ";0"
    If str���㷽ʽ = "" Then str���㷽ʽ = "�����ʻ�;0;0"
    �����������_üɽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_üɽ(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, _
Optional ByVal bln��Ժ���� As Boolean = True) As Boolean
    Dim int���� As Integer
    Dim int��Ժ As Integer, int��Ժ As Integer
    Dim cur�ʻ���� As Currency, curͳ���ۼ� As Currency, intסԺ���� As Integer
    Dim rsTemp As New ADODB.Recordset
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    On Error GoTo errHand
    int���� = IIf(bln��Ժ����, 1, 2)
    
    '���¸����ʻ�
    If cur�����ʻ� <> 0 Then
        If Not �¸����ʻ�(gComInfo_üɽ.����ID, cur�����ʻ� * -1) Then Exit Function
    End If
    
    '��������Ϣ���浽���ս����¼��
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�Ĵ�üɽ & "," & gComInfo_üɽ.����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & ",0,0,0,0,NULL,0,0,0," & _
        gComInfo_üɽ.�����ܶ� & "," & gComInfo_üɽ.ȫ�Ը� & "," & gComInfo_üɽ.�����Ը� & "," & gComInfo_üɽ.����ͳ�� & "," & gComInfo_üɽ.ͳ��֧�� & ",0," & _
        0 & "," & cur�����ʻ� & ",null,null,null,null," & gComInfo_üɽ.����ID & ",'" & gstrUserName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���������շ�����")
    
    '���ո�����ı�����ϸ
    With rs����_����
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!�����ܶ�, 0) <> 0 Then
                gstrSQL = "ZL_���ձ�����¼_INSERT(" & int���� & "," & lng����ID & "," & _
                "'" & !������� & "','" & !�������� & "'," & !ͳ��ȶ� & "," & _
                "" & !��׼���� & "," & !��׼���� & "," & !�����ܶ� & "," & !�����ܶ� & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���������շ�����")
            End If
            .MoveNext
        Loop
    End With
    
    If cur�����ʻ� <> 0 Then
        cur�ʻ���� = 0: curͳ���ۼ� = 0: intסԺ���� = 0
        gstrSQL = " Select Nvl(�ʻ������ۼ�,0) �ʻ����,Nvl(����ͳ���ۼ�,0) ͳ���ۼ�,Nvl(סԺ�����ۼ�,0) ��Ժ,Nvl(��ԺסԺ����,0) ��Ժ  From �ʻ������Ϣ" & _
                  " Where ����ID=[1] ANd ���=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ����", gComInfo_üɽ.����ID, Format(zlDatabase.Currentdate, "yyyy"))
        '�����¸����ʻ�ʱ���Ѿ�������ʻ������Ա��β����Ӽ�����
        If Not rsTemp.EOF Then
            cur�ʻ���� = Nvl(rsTemp!�ʻ����, 0)
            curͳ���ۼ� = Nvl(rsTemp!ͳ���ۼ�, 0)
            int��Ժ = rsTemp!��Ժ
            int��Ժ = rsTemp!��Ժ
        End If
        gstrSQL = "zl_�ʻ������Ϣ_Insert(" & gComInfo_üɽ.����ID & ",25," & Format(zlDatabase.Currentdate, "yyyy") & _
                  "," & cur�ʻ���� & ",0," & curͳ���ۼ� & ",0," & int��Ժ & "," & int��Ժ & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ����")
    End If
    
    �������_üɽ = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_üɽ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim cur�ʻ���� As Currency, curͳ���ۼ� As Currency, intסԺ���� As Integer
    Dim lng��� As Long, lng��¼ID As Long
    Dim int��Ժ As Integer, int��Ժ As Integer
    Dim rsTemp As New ADODB.Recordset
On Error GoTo ErrH
    '�������������¼
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng��¼ID = rsTemp!����ID
    
    gstrSQL = "Select * From ���ս����¼ Where ����=25 And ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����¼", lng����ID)
    lng��� = Format(rsTemp!����ʱ��, "yyyy")
    If lng��� <> Format(zlDatabase.Currentdate, "yyyy") Then
        Err.Raise 9000, gstrSysName, "���ܳ���������ȵĵ��ݣ�"
        Exit Function
    End If
    With rsTemp
        gstrSQL = "zl_���ս����¼_insert(" & !���� & "," & lng��¼ID & ",25," & !����ID & "," & _
            lng��� & ",0,0,0,0," & Nvl(!סԺ����, 0) & "," & -1 * Nvl(!����, 0) & ",0," & -1 * Nvl(!ʵ������, 0) & "," & _
            -1 * Nvl(!�������ý��, 0) & "," & -1 * Nvl(!ȫ�Ը����, 0) & "," & -1 * Nvl(!�����Ը����, 0) & "," & -1 * Nvl(!����ͳ����, 0) & "," & -1 * Nvl(!ͳ�ﱨ�����, 0) & ",0," & _
            0 & "," & -1 * cur�����ʻ� & ",'" & lng����ID & "',null,null,null,null,'" & gstrUserName & "')" '֧��˳����������汻�����ļ�¼ID
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������������¼")
        
        gstrSQL = "Select * From ���ձ�����¼ Where ��¼ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���౨����¼", lng����ID)
        '���ո�����ı�����ϸ
        Do While Not .EOF
            If Nvl(!�����ܶ�, 0) <> 0 Then
                gstrSQL = "ZL_���ձ�����¼_INSERT(" & !���� & "," & lng��¼ID & "," & _
                "'" & !������� & "','" & !�������� & "'," & !ͳ��ȶ� & "," & _
                "" & !��׼���� & "," & !��׼���� & "," & -1 * !�����ܶ� & "," & -1 * !�����ܶ� & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���������շ�����")
            End If
            .MoveNext
        Loop
    End With
    
    '��ԭ�����ʻ�
    If cur�����ʻ� <> 0 Then
        If Not �¸����ʻ�(lng����ID, cur�����ʻ�) Then Exit Function
        
        cur�ʻ���� = 0: curͳ���ۼ� = 0: intסԺ���� = 0
        gstrSQL = " Select Nvl(�ʻ������ۼ�,0) �ʻ����,Nvl(����ͳ���ۼ�,0) ͳ���ۼ�,Nvl(סԺ�����ۼ�,0) ��Ժ,Nvl(��ԺסԺ����,0) ��Ժ  From �ʻ������Ϣ" & _
                  " Where ����ID=[1] ANd ���=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ����", lng����ID, lng���)
        '�����¸����ʻ�ʱ���Ѿ�������ʻ������Ա��β����Ӽ�����
        If Not rsTemp.EOF Then
            cur�ʻ���� = Nvl(rsTemp!�ʻ����, 0)
            curͳ���ۼ� = Nvl(rsTemp!ͳ���ۼ�, 0)
            int��Ժ = rsTemp!��Ժ
            int��Ժ = rsTemp!��Ժ
        End If
        gstrSQL = "zl_�ʻ������Ϣ_Insert(" & lng����ID & ",25," & Format(zlDatabase.Currentdate, "yyyy") & _
                  "," & cur�ʻ���� & ",0," & curͳ���ۼ� & ",0," & int��Ժ & "," & int��Ժ & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ����")
    End If
    
    ����������_üɽ = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ��Ժ�Ǽ�_üɽ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim str��Ժ����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    On Error GoTo errHand
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ĵ�üɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    ��Ժ�Ǽ�_üɽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_üɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ��
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
    On Error GoTo errHand
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ĵ�üɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    
    ��Ժ�Ǽ�_üɽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽǳ���_üɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    gstrSQL = " Select Count(*) Records From סԺ���ü�¼ " & _
              " Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ժ���", lng����ID, lng��ҳID)
    If rsTemp!Records <> 0 Then
        MsgBox "�Ѿ����ڷ��ü�¼���������������Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ĵ�üɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_üɽ = True
End Function

Public Function ��Ժ�Ǽǳ���_üɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�Ĵ�üɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_üɽ = True
End Function

Public Function סԺ�������_üɽ(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim curTotal As Currency
    Dim lng��ҳID As Long
    Dim cur�����Ը� As Currency, cur�����ʻ� As Long
    Dim str��Ժ��� As String, str������� As String
    Dim str����ʱ�� As String, str����ʱ�� As String
    Dim blnUpload As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    '�ӿڷ��صı������ȥ����סԺ�ڼ�����������Ļ��ܽ��󣬲��Ǳ��ε�ʵ�ʱ�����
    'rsExse��¼���е��ֶ��嵥
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    On Error GoTo errHand
    
    סԺ�������_üɽ = "�����ʻ�;0;0"
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function סԺ����_üɽ(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim cur�����ʻ� As Currency
    Dim lng��ҳID As Long
    Dim str��Ժ��� As String, str������� As String
    Dim str����ʱ�� As String, str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
        '������㣨���ص����ݼ�ȥ���ν������ݣ��͵��ڱ��ε���ʵ�������ݣ�
    On Error GoTo errHand
    
    סԺ����_üɽ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function סԺ�������_üɽ(lng����ID As Long) As Boolean
    Dim lng����ID As Long
    Dim str�˵���� As String
    Dim rsTemp As New ADODB.Recordset
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    MsgBox "��ҽ����֧�ֳ������뵽ҽ�������", vbInformation, gstrSysName
    סԺ�������_üɽ = False
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ҽ����ֹ_üɽ() As Boolean
    ҽ����ֹ_üɽ = True
End Function

Public Function �ҺŽ���_üɽ(ByVal lng����ID As Long, ByVal cur��� As Currency) As Boolean
    '��������ȡ�������
    
    �ҺŽ���_üɽ = True
End Function

Public Function �ҺŽ������_üɽ(ByVal lng����ID As Long) As Boolean
    �ҺŽ������_üɽ = True
End Function

Private Function CalcPrepare(Optional ByVal bln��Ժ���� As Boolean = True) As Boolean
    '���ڽ���ǰ����ȡ�ò��˵������Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim str��� As String
    
    '������Ϣ
    gstrSQL = " Select A.*,B.���� ��������,C.���� ��Ⱥ " & _
              " From �����ʻ� A,(Select * From ���ղ��� Where ����=" & TYPE_�Ĵ�üɽ & ") B, " & _
              " (Select * From ������Ⱥ Where ����=[1]) C" & _
              " Where A.����ID=B.ID(+) And A.����=[1] And A.����ID=[2]" & _
              " And A.��ְ=C.��� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�������ʻ���Ϣ", TYPE_�Ĵ�üɽ, gComInfo_üɽ.����ID)
    gComInfo_üɽ.����ID = Nvl(rsTemp!����ID, 0)
    gComInfo_üɽ.�������� = Nvl(rsTemp!��������)
    gComInfo_üɽ.���� = rsTemp!����
    gComInfo_üɽ.ҽ���� = rsTemp!ҽ����
    gComInfo_üɽ.���� = rsTemp!����
    gComInfo_üɽ.��Ⱥ = rsTemp!��Ⱥ
    gComInfo_üɽ.�ʻ���� = Nvl(rsTemp!�ʻ����, 0)
    
    '����Ƿ����ñ��ղ���
    gstrSQL = "Select count(*) Records From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡϵͳ����", TYPE_�Ĵ�üɽ)
    If rsTemp!Records = 0 Then
        MsgBox "�����ñ��ղ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�ȼ���Ƿ����úñ���ȵĽ������
    str��� = Format(zlDatabase.Currentdate(), "yyyy")
    gstrSQL = " Select Count(*) Records From ���ձ������� A,������Ⱥ B" & _
            " Where A.����=[1] And A.����=[2]" & _
            " And A.����=1 And A.��Ժ=[3] And A.���=[4]" & _
            " And A.��Ⱥ=B.��� And A.����=B.���� And B.����=[5]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "֧������", TYPE_�Ĵ�üɽ, gComInfo_üɽ.����, IIf(bln��Ժ����, 1, 2), str���, gComInfo_üɽ.��Ⱥ)
    If rsTemp!Records = 0 Then
        MsgBox "��δ���ñ���ȵı��ձ������ߣ�[��Ƚ������]", vbInformation, gstrSysName
        Exit Function
    End If
    gstrSQL = " Select Count(*) Records From ���ձ������� A,������Ⱥ B" & _
            " Where A.����=[1] And A.����=[2]" & _
            " And A.����=2 And A.��Ժ=[3] And A.���=[4]" & _
            " And A.��Ⱥ=B.��� And A.����=B.���� And B.����=[5]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "֧������", TYPE_�Ĵ�üɽ, gComInfo_üɽ.����, IIf(bln��Ժ����, 1, 2), str���, gComInfo_üɽ.��Ⱥ)
    If rsTemp!Records = 0 Then
        MsgBox "��δ���ñ���ȵı��ձ������ߣ�[��Ƚ������]", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�Ǳ�Ժ����ʱ������ģ���Ϊ����ı�����ֵ
    CalcPrepare = True
    If bln��Ժ���� = False Then Exit Function
    'ȡ�����סԺ����
    gstrSQL = "Select Nvl(סԺ�����ۼ�,0) סԺ���� From �ʻ������Ϣ Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����סԺ����ͳ��", TYPE_�Ĵ�üɽ, gComInfo_üɽ.����ID)
    If Not rsTemp.EOF Then
        gComInfo_üɽ.סԺ���� = rsTemp!סԺ���� + 1
    Else
        gComInfo_üɽ.סԺ���� = 1
    End If
End Function

Public Function ҽ������_üɽ() As Boolean
    ҽ������_üɽ = frmSetüɽ.ShowSet()
End Function

Public Sub Init_����_����()
    '��ʼ�������¼��
    Set rs����_���� = New ADODB.Recordset
    With rs����_����
        If .State = 1 Then .Close
        .Fields.Append "�������", adLongVarChar, 10  '0:��ʾ����
        .Fields.Append "��������", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "ͳ��ȶ�", adDouble, 18, adFldIsNullable
        .Fields.Append "��׼����", adDouble, 18, adFldIsNullable
        .Fields.Append "��׼����", adDouble, 5, adFldIsNullable
        .Fields.Append "�����ܶ�", adDouble, 18, adFldIsNullable
        .Fields.Append "�����ܶ�", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    gstrSQL = "Select * From ����֧������ Where ����=[1]"
    Set rs֧������_���� = zlDatabase.OpenSQLRecord(gstrSQL, "����֧������", TYPE_�Ĵ�üɽ)
    With rs֧������_����
        Do While Not .EOF
            rs����_����.AddNew
            rs����_����!������� = !����
            rs����_����!�������� = Nvl(!����)
            rs����_����!ͳ��ȶ� = Nvl(!ͳ��ȶ�, 0)
            rs����_����!��׼���� = Nvl(!��׼����, 0)
            rs����_����!��׼���� = Nvl(!��׼����, 0)
            rs����_����!�����ܶ� = 0
            rs����_����!�����ܶ� = 0
            rs����_����!���� = 0
            rs����_����.Update
            .MoveNext
        Loop
    End With
    
    Set rs�ֵ�֧��_���� = New ADODB.Recordset
    With rs�ֵ�֧��_����
        If .State = 1 Then .Close
        .Fields.Append "����", adDouble, 10  '0:��ʾ����
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "����ͳ��", adDouble, 18, adFldIsNullable
        .Fields.Append "ͳ�ﱨ��", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Sub Init_�ṹ��_����()
    With gComInfo_üɽ
        .סԺ���� = 0
        .סԺ���� = 0
        .�����ܶ� = 0
        .����ͳ�� = 0
        .ʵ�ʱ��� = 0
        .����ʵ�ʱ��� = 0
        .ȫ�Ը� = 0
        .�����Ը� = 0
        .ͳ��֧�� = 0
        .ͳ���Ը� = 0
        .�ʻ�֧�� = 0
        .����޶� = 0
        .�������� = 0
        .����ID = 0
        .��Ⱥ = 0
    End With
End Sub

Private Function Calc_����ͳ��(ByVal rsExse As ADODB.Recordset, Optional ByVal bln���� As Boolean = True) As Boolean
    Dim cur��� As Currency, curͳ�� As Currency, cur����ͳ�� As Currency
    Dim intסԺ���� As Integer
    Dim str������� As String
    Dim rsTemp As New ADODB.Recordset
    '�����մ�����㣺���η����ܶ�������ߡ�����ͳ���ȫ�Ը��������Ը��������½ṹ�壨�������ģ�
    'rsExse���ֶΣ�����ID,�շ����,�վݷ�Ŀ,���㵥λ,������,�շ�ϸĿID,����,����,
    '              ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Call Init_����_����
    Call Init_�ṹ��_����
    
    cur��� = 0: curͳ�� = 0:  cur����ͳ�� = 0
    gComInfo_üɽ.����ID = rsExse!����ID
    If Not CalcPrepare Then Exit Function        '���ﲡ�˽���ʱ������Ҫ��ȡ�����Ϣ����Ϊ������ˢ��ʱ���Ѿ��õ�
    
    '���ܷ������
    With rsExse
        Do While Not .EOF
            '���ж��Ƿ�������ҽ����Ӧ��Ŀ����
            gstrSQL = " Select �Ƿ�ҽ�� From ����֧����Ŀ" & _
                      " Where ����=[1] And �շ�ϸĿID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ������˶�Ӧ��ҽ������", TYPE_�Ĵ�üɽ, CLng(!�շ�ϸĿID))
            If rsTemp.EOF Then
                MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            
            If Nvl(!����֧������ID, 0) = 0 Then
                MsgBox "����Ŀδ���ö�Ӧ��ҽ�����࣬���ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            
            cur��� = cur��� + Nvl(!ʵ�ս��, 0)
            str������� = ""
            If rs֧������_����.RecordCount <> 0 Then
                rs֧������_����.MoveFirst
                rs֧������_����.Find "ID=" & !����֧������ID
                If .EOF Then
                    MsgBox "���մ��෢���ı䣬������Ԥ���㣡", vbInformation, gstrSysName
                    Exit Function
                End If
                str������� = rs֧������_����!����
            End If
            
            If str������� = "" Then
                MsgBox "���մ��෢���ı䣬������Ԥ���㣡", vbInformation, gstrSysName
                Exit Function
            End If
            
            '��Ϊrs����_����������rs֧������_���ײ����ģ��ߵ��ⲽ���϶�����ΪEOF
            rs����_����.MoveFirst
            rs����_����.Find "�������='" & str������� & "'"
            rs����_����!�����ܶ� = rs����_����!�����ܶ� + Nvl(!ʵ�ս��, 0)
            If rs֧������_����!�Ƿ�ҽ�� = 1 And Val(rs֧������_����!ͳ��ȶ�) <> 0 Then
                '���������ҽ����Ŀ��ͳ��ȶ������
                If rs֧������_����!������� = 3 Or rs֧������_����!������� = IIf(bln����, 1, 2) Then
                    '������������ȷ
                    If rsTemp!�Ƿ�ҽ�� = 1 Then
                        '�����ϸ��ĿҲ��ҽ����Ŀ
                        curͳ�� = curͳ�� + Nvl(!ʵ�ս��, 0)
                        rs����_����!�����ܶ� = rs����_����!�����ܶ� + Nvl(!ʵ�ս��, 0)
                        rs����_����!���� = rs����_����!���� + Nvl(!����, 0)
                    End If
                End If
            End If
            rs����_����.Update
            .MoveNext
        Loop
    End With
    
    gComInfo_üɽ.�����ܶ� = cur���
    gComInfo_üɽ.ȫ�Ը� = cur��� - curͳ��
    '�������ͳ����
    With rs����_����
        .MoveFirst
        Do While Not .EOF
            If !��׼���� = 0 And !��׼���� = 0 Then
                !�����ܶ� = !�����ܶ� * Nvl(!ͳ��ȶ�, 0) / 100
            Else
                If !���� > !��׼���� Then
                    '���סԺ�ճ�����׼��������ô������ ��׼����*��׼���� +  (����-��׼����)*ͳ��ȶ�
                    '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                    !�����ܶ� = !��׼���� * !��׼���� + _
                        (!���� - IIf(!��׼���� = 0 Or !��׼���� = 0, 0, !��׼����)) * !ͳ��ȶ�
                Else
                    If !��׼���� = 0 Or !��׼���� = 0 Then
                        '���סԺ�յ�����׼��������ô�������� ����*��׼���� ���� ����*ͳ��ȶ�
                        '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                        !�����ܶ� = !�����ܶ� * !ͳ��ȶ� / 100
                    Else
                        !�����ܶ� = !���� * !��׼����
                    End If
                End If
            End If
            cur����ͳ�� = cur����ͳ�� + !�����ܶ�
            .Update
            .MoveNext
        Loop
    End With
    gComInfo_üɽ.�����Ը� = curͳ�� - cur����ͳ��
    gComInfo_üɽ.����ͳ�� = cur����ͳ��
    
    Call Calc_ʵ�ʱ���_����
    Calc_����ͳ�� = True
End Function

Public Sub Calc_ʵ�ʱ���_����()
    Dim intBound As Integer, lngRow As Long
    Dim sin���� As Single
    Dim rsTemp As New ADODB.Recordset
    
    '����ʵ�ʱ����������㣬ֻҪ����������100%�Ĵ��࣬Ҫ����ֵ����㣻���򲻽���
    gstrSQL = " Select A.����,B.����ֵ ���� From ����֧������ A,(Select * From ���ղ��� Where ����=[1] And ����=1 And ���>10) B " & _
              " Where A.����=[2] And A.����=B.������(+) Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", 25)
    
    With rs����_����
        .MoveFirst
        Do While Not .EOF
            If !�����ܶ� <> 0 Then
                rsTemp.MoveFirst
                rsTemp.Find "����='" & !�������� & "'"
                sin���� = 100
                If Not rsTemp.EOF Then sin���� = Nvl(rsTemp!����, 0)
                '����û������˱��������û������Ϊ׼
                If InStr(1, gstrʵ�ʱ�������_����, "|" & !�������� & ";") <> 0 Then
                    intBound = UBound(Split(Mid(gstrʵ�ʱ�������_����, 2), "|"))
                    For lngRow = 0 To intBound
                        If Split(Split(Mid(gstrʵ�ʱ�������_����, 2), "|")(lngRow), ";")(0) = !�������� Then
                            sin���� = Val(Split(Split(Mid(gstrʵ�ʱ�������_����, 2), "|")(lngRow), ";")(1))
                            Exit For
                        End If
                    Next
                End If
                
                '���ʵ�ʱ���������Ϊ100%����ò��ֽ��ֱ�Ӱ���������������ٽ���ֵ�����
                Debug.Print sin����
                If sin���� <> 100 And !�������� <> "��������" Then
                    gComInfo_üɽ.����ʵ�ʱ��� = gComInfo_üɽ.����ʵ�ʱ��� + !�����ܶ�
                    gComInfo_üɽ.ʵ�ʱ��� = gComInfo_üɽ.ʵ�ʱ��� + (!�����ܶ� * sin���� / 100)
                End If
            End If
            .MoveNext
        Loop
    End With
End Sub

Public Sub Calc_���ﱨ������_����(Optional ByVal bln��Ժ As Boolean = True, Optional ByVal blnҽԺ���� As Boolean = False)
    Dim rsPara As New ADODB.Recordset
    Dim rsScale As New ADODB.Recordset
    Dim rsSum As New ADODB.Recordset
    Dim bln�ȿ۸����ʻ� As Boolean
    Dim sin���� As Single
    Dim sinʵ���������ͳ����� As Single
    '��ȡ���ղ���
    gstrSQL = "Select ����,���,����ֵ From ���ղ��� Where ����=[1] Order by ����,���"
    Set rsPara = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ղ���", TYPE_�Ĵ�üɽ)
    
    '����ʵ�ʽ���ͳ����
    gstrSQL = "Select ���� From ���ձ������� A,������Ⱥ B" & _
             " Where A.����=1 And A.����=[1] And A.��Ժ=[2]" & _
             " And B.����=A.���� And A.����=[3]" & _
             " And A.��Ⱥ=B.��� And B.����=[4] And A.����=0 And A.���=" & Format(zlDatabase.Currentdate, "yyyy")
    Set rsScale = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ﱨ������", gComInfo_üɽ.����, IIf(bln��Ժ, 1, 2), TYPE_�Ĵ�üɽ, gComInfo_üɽ.��Ⱥ)
    If rsScale.EOF Then
        sinʵ���������ͳ����� = 0
    Else
        sinʵ���������ͳ����� = Nvl(rsScale!����, 0)
    End If
    
    '��ʼ��
    gComInfo_üɽ.�ʻ�֧�� = 0
    gComInfo_üɽ.����޶� = 0
    gComInfo_üɽ.�������� = 0
    gComInfo_üɽ.�ѱ������ = 0
    rsPara.Filter = "����=" & gComInfo_üɽ.����
    
    Dim cur�۳� As Currency
    If (gComInfo_üɽ.�������� = "���ز�" Or gComInfo_üɽ.��Ⱥ = "����") And gComInfo_üɽ.�ʻ���� > 0 Then
        If sinʵ���������ͳ����� <> 0 Then
            cur�۳� = gComInfo_üɽ.�ʻ���� / sinʵ���������ͳ����� * 100
            If Calc_����ͳ�� > cur�۳� Then
                gComInfo_üɽ.����ͳ�� = Calc_����ͳ�� - cur�۳� + gComInfo_üɽ.����ʵ�ʱ���
                gComInfo_üɽ.�ʻ�֧�� = gComInfo_üɽ.�ʻ����
            Else
                cur�۳� = Calc_����ͳ�� * sinʵ���������ͳ����� / 100
                gComInfo_üɽ.�ʻ�֧�� = cur�۳�
                gComInfo_üɽ.����ͳ�� = gComInfo_üɽ.ʵ�ʱ���
            End If
        End If
    End If
    
    If gComInfo_üɽ.���� = 1 Then  '����
        If gComInfo_üɽ.�������� = "���ز�" Then
            bln�ȿ۸����ʻ� = (cur�۳� = 0)
            'ȡ��������
            rsPara.MoveFirst
            rsPara.Find "���=1"
            gComInfo_üɽ.�������� = Nvl(rsPara!����ֵ, 0)
            rsPara.MoveFirst
            rsPara.Find "���=2"
            gComInfo_üɽ.����޶� = Nvl(rsPara!����ֵ, 0)
            
            'ȡ�����Ѿ��������
            If gComInfo_üɽ.����޶� <> 0 Then
                gstrSQL = " Select Sum(����ͳ����) �ѱ������ From ���ս����¼ " & _
                          " Where ����=[2] And ����=1 And ����ID=" & gComInfo_üɽ.����ID & _
                          " And ���='" & Format(zlDatabase.Currentdate, "yyyy") & "' And ����ID=[1]"
                Set rsSum = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������ѱ������", CLng(gComInfo_üɽ.����ID), TYPE_�Ĵ�üɽ)
                gComInfo_üɽ.�ѱ������ = Nvl(rsSum!�ѱ������, 0)
            End If
        ElseIf gComInfo_üɽ.��Ⱥ = "����" Then
            bln�ȿ۸����ʻ� = (cur�۳� = 0)
            'ȡ��������
            rsPara.MoveFirst
            rsPara.Find "���=3"
            gComInfo_üɽ.�������� = sinʵ���������ͳ�����
        ElseIf gComInfo_üɽ.��Ⱥ = "�˲о���" Then
            'ȡ��������
            rsPara.MoveFirst
            rsPara.Find "���=4"
            gComInfo_üɽ.�������� = sinʵ���������ͳ�����
        ElseIf gComInfo_üɽ.�������� = "������Ⱥ" Then
            'ȡ��������
            rsPara.MoveFirst
            rsPara.Find "���=5"
            gComInfo_üɽ.�������� = 100
        ElseIf gComInfo_üɽ.�������� = "�ƻ�����" Then
            gComInfo_üɽ.�������� = 100
        Else
            If sinʵ���������ͳ����� <> 0 Then
                cur�۳� = gComInfo_üɽ.�ʻ���� / sinʵ���������ͳ����� * 100
                If Calc_����ͳ�� > cur�۳� Then
                    gComInfo_üɽ.����ͳ�� = Calc_����ͳ�� - cur�۳� + gComInfo_üɽ.����ʵ�ʱ���
                    gComInfo_üɽ.�ʻ�֧�� = gComInfo_üɽ.�ʻ����
                Else
                    cur�۳� = Calc_����ͳ�� * sinʵ���������ͳ����� / 100
                    gComInfo_üɽ.�ʻ�֧�� = cur�۳�
                    gComInfo_üɽ.����ͳ�� = 0
                End If
            End If
        End If
    Else                            '��ҵ��˾
        If gComInfo_üɽ.�������� = "���ز�" Then
            bln�ȿ۸����ʻ� = (cur�۳� = 0)
            'ȡ��������
            rsPara.MoveFirst
            rsPara.Find "���=1"
            gComInfo_üɽ.�������� = Nvl(rsPara!����ֵ, 0)
            rsPara.MoveFirst
            rsPara.Find "���=2"
            gComInfo_üɽ.����޶� = Nvl(rsPara!����ֵ, 0)
            
            'ȡ�����Ѿ��������
            If gComInfo_üɽ.����޶� <> 0 Then
                gstrSQL = " Select Sum(����ͳ����) �ѱ������ From ���ս����¼ " & _
                          " Where ����=[1] And ����=1 And ����ID=[2]" & _
                          " And ���='" & Format(zlDatabase.Currentdate, "yyyy") & "' And ����ID=[3]"
                Set rsSum = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������ѱ������", TYPE_�Ĵ�üɽ, gComInfo_üɽ.����ID, gComInfo_üɽ.����ID)
                gComInfo_üɽ.�ѱ������ = Nvl(rsSum!�ѱ������, 0)
            End If
        ElseIf gComInfo_üɽ.�������� = "����" Then
            rsPara.MoveFirst
            rsPara.Find "���=3"
            gComInfo_üɽ.�������� = 100
        Else
            If sinʵ���������ͳ����� <> 0 Then
                cur�۳� = gComInfo_üɽ.�ʻ���� / sinʵ���������ͳ����� * 100
                If Calc_����ͳ�� > cur�۳� Then
                    gComInfo_üɽ.����ͳ�� = Calc_����ͳ�� - cur�۳� + gComInfo_üɽ.����ʵ�ʱ���
                    gComInfo_üɽ.�ʻ�֧�� = gComInfo_üɽ.�ʻ����
                Else
                    cur�۳� = Calc_����ͳ�� * sinʵ���������ͳ����� / 100
                    gComInfo_üɽ.�ʻ�֧�� = cur�۳�
                    gComInfo_üɽ.����ͳ�� = gComInfo_üɽ.����ʵ�ʱ���
                End If
            End If
        End If
    End If
    
    '�����������Ϊ�㣬�����޶�ҲΪ�㣬����Ҫ��������ʻ�
    If Calc_����ͳ�� > 0 Then
        If gComInfo_üɽ.�������� <> 0 Or gComInfo_üɽ.����޶� <> 0 Then
            If gComInfo_üɽ.����޶� <> 0 Then
                gComInfo_üɽ.����޶� = gComInfo_üɽ.����޶� - gComInfo_üɽ.�ѱ������
            Else
                gComInfo_üɽ.����޶� = gComInfo_üɽ.�����ܶ�
            End If
            
            If bln�ȿ۸����ʻ� Then
                gComInfo_üɽ.�ʻ�֧�� = IIf(gComInfo_üɽ.�ʻ���� > Calc_����ͳ��, Calc_����ͳ��, gComInfo_üɽ.�ʻ����)
                gComInfo_üɽ.����ͳ�� = Calc_����ͳ�� - gComInfo_üɽ.�ʻ�֧�� + gComInfo_üɽ.����ʵ�ʱ���
            End If
            If gComInfo_üɽ.ʵ�ʱ��� <> 0 Then
                If gComInfo_üɽ.�ʻ���� - gComInfo_üɽ.�ʻ�֧�� > gComInfo_üɽ.ʵ�ʱ��� Then
                    gComInfo_üɽ.�ʻ�֧�� = gComInfo_üɽ.�ʻ�֧�� + gComInfo_üɽ.ʵ�ʱ���
                    gComInfo_üɽ.ʵ�ʱ��� = 0
                    gComInfo_üɽ.����ʵ�ʱ��� = 0
                Else
                    Dim sin�������� As Single
                    sin�������� = gComInfo_üɽ.ʵ�ʱ��� / gComInfo_üɽ.����ʵ�ʱ���
                    gComInfo_üɽ.ʵ�ʱ��� = gComInfo_üɽ.ʵ�ʱ��� - (gComInfo_üɽ.�ʻ���� - gComInfo_üɽ.�ʻ�֧��)
                    gComInfo_üɽ.����ʵ�ʱ��� = gComInfo_üɽ.ʵ�ʱ��� / sin��������
                    gComInfo_üɽ.�ʻ�֧�� = gComInfo_üɽ.�ʻ����
                End If
            End If
            gComInfo_üɽ.ͳ��֧�� = IIf(gComInfo_üɽ.����ͳ�� > gComInfo_üɽ.����޶�, _
                gComInfo_üɽ.����޶�, Calc_����ͳ��) * Val(gComInfo_üɽ.��������) / 100 + gComInfo_üɽ.ʵ�ʱ���
            gComInfo_üɽ.ͳ���Ը� = gComInfo_üɽ.����ͳ�� - gComInfo_üɽ.ͳ��֧��
        Else
            If gComInfo_üɽ.�ʻ�֧�� = 0 Then
                gComInfo_üɽ.�ʻ�֧�� = IIf(gComInfo_üɽ.�ʻ���� > gComInfo_üɽ.����ͳ��, gComInfo_üɽ.����ͳ��, gComInfo_üɽ.�ʻ����)
                gComInfo_üɽ.����ͳ�� = 0
                gComInfo_üɽ.ͳ��֧�� = 0
                gComInfo_üɽ.ͳ���Ը� = 0
            End If
        End If
    Else
        gComInfo_üɽ.����ͳ�� = gComInfo_üɽ.ʵ�ʱ���
        gComInfo_üɽ.�ʻ�֧�� = gComInfo_üɽ.�ʻ�֧�� + IIf(gComInfo_üɽ.�ʻ���� - gComInfo_üɽ.�ʻ�֧�� > gComInfo_üɽ.����ͳ��, gComInfo_üɽ.����ͳ��, gComInfo_üɽ.�ʻ���� - gComInfo_üɽ.�ʻ�֧��)
        gComInfo_üɽ.����ͳ�� = gComInfo_üɽ.����ͳ�� - IIf(gComInfo_üɽ.�ʻ���� - gComInfo_üɽ.�ʻ�֧�� > gComInfo_üɽ.����ͳ��, gComInfo_üɽ.����ͳ��, gComInfo_üɽ.�ʻ���� - gComInfo_üɽ.�ʻ�֧��)
    End If
    Call FormatData(blnҽԺ����)
    
    rsPara.Filter = 0
    rsPara.Close
    Set rsPara = Nothing
End Sub

Private Sub FormatData(Optional ByVal blnҽԺ���� As Boolean = False)
    '�����ݸ�ʽ��ΪһλС��
    With gComInfo_üɽ
        .ͳ��֧�� = Val(Format(.ͳ��֧��, gstrFormat_üɽ))
        If Not blnҽԺ���� Then .�ʻ�֧�� = Val(Format(.�ʻ�֧��, gstrFormat_üɽ))
    End With
End Sub

Private Function Calc_����ͳ��() As Currency
    '���ڼ��㷽�����⣬��Ҫȥ��ʵ�ʱ�����������100%����Ŀ�Ľ���ͳ������м��㣬������ͳ��Ľ���ֲ������仯
    '����������ҩ��100%����ͳ�ʵ�ʱ���������100%����������100
    '      ���  ��100%����ͳ�ʵ�ʱ���������50% ����������100
    '�������
    '      ���ν���ͳ���ܶ� ��200
    '      ���α����ܶ�     ��100(��)+50(��)
    Calc_����ͳ�� = gComInfo_üɽ.����ͳ�� - gComInfo_üɽ.����ʵ�ʱ���
End Function

Public Sub Calc_סԺ��������_����(Optional ByVal bln��Ժ As Boolean = True)
    Dim cur����ͳ�� As Currency, curʣ��ͳ�� As Currency, curͳ��֧�� As Currency, cur���α��� As Currency
    Dim cur����ͳ���ۼ� As Currency, cur������뵵�� As Currency '����Ƚ���ͳ���ۼƣ�������뵵�����ڼ�����ʼ���εı������
    Dim sin���� As Single
    Dim lng��� As Long, int��ʼ�� As Integer
    Dim blnExit As Boolean
    Dim rs���� As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    '����ͳ��֧��������½ṹ�壨��Ҫ�������ġ���Ⱥ��
    gstrSQL = "Select A.����,C.����,C.����,C.����,C.���� From ���ձ������� A,������Ⱥ B,���շ��õ� C " & _
             " Where A.����=1 And A.����=[1] And A.��Ժ=" & IIf(bln��Ժ, 1, 2) & _
             " And B.����=A.���� And A.����= [2]  And A.����<>0 " & _
             " And A.��Ⱥ=B.��� And B.����=[3] And A.���=" & Format(zlDatabase.Currentdate(), "yyyy") & _
             " And A.����=C.���� And C.����=A.���� And C.����=A.����" & _
             " Order by ����"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���շ��õ�", gComInfo_üɽ.����, TYPE_�Ĵ�üɽ, gComInfo_üɽ.��Ⱥ)
    
    cur����ͳ�� = 0: curͳ��֧�� = 0
    'Ӧ���ȼ�ȥʵ�ʱ������֣���Ϊ�ⲿ�ֱ�����һ�£�����ȼ���ͨͳ�ﲿ�֣�����Ҫ����ʵ�ʱ������ֽ�
    
    If gComInfo_üɽ.����ͳ�� >= gComInfo_üɽ.���� Then
        If Calc_����ͳ�� > gComInfo_üɽ.���� Then
            curʣ��ͳ�� = Calc_����ͳ�� - gComInfo_üɽ.����
'            gComInfo_üɽ.����ͳ�� = curʣ��ͳ�� + gComInfo_üɽ.����ʵ�ʱ���
        Else
            sin���� = gComInfo_üɽ.ʵ�ʱ��� / gComInfo_üɽ.����ʵ�ʱ���
            curʣ��ͳ�� = 0
            gComInfo_üɽ.����ʵ�ʱ��� = gComInfo_üɽ.����ʵ�ʱ��� - (gComInfo_üɽ.���� - Calc_����ͳ��)
            gComInfo_üɽ.ʵ�ʱ��� = gComInfo_üɽ.����ʵ�ʱ��� * sin����
'            gComInfo_üɽ.����ͳ�� = gComInfo_üɽ.����ʵ�ʱ���
        End If
    Else
        curʣ��ͳ�� = 0
        gComInfo_üɽ.����ʵ�ʱ��� = 0
        gComInfo_üɽ.ʵ�ʱ��� = 0
    End If
    gComInfo_üɽ.�������� = gComInfo_üɽ.����
    If curʣ��ͳ�� <= 0 Then
        curʣ��ͳ�� = 0
        gComInfo_üɽ.�������� = gComInfo_üɽ.����ͳ��
    End If
    
    '��ȡ����Ƚ���ͳ���ۼ�
    cur������뵵�� = 0
    cur����ͳ���ۼ� = 0
    lng��� = Format(zlDatabase.Currentdate, "yyyy")
    gstrSQL = " Select Nvl(����ͳ���ۼ�,0) ����ͳ���ۼ� From �ʻ������Ϣ " & _
              " Where ���=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����Ƚ���ͳ���ۼ�", lng���, gComInfo_üɽ.����ID)
    If Not rsTemp.EOF Then cur����ͳ���ۼ� = Nvl(rsTemp!����ͳ���ۼ�, 0)
    
    With rs����
        '����ÿ���ε�������ȫ�ּ�¼����
        .MoveFirst
        Do While Not .EOF
            If int��ʼ�� = 0 Then
                If (cur����ͳ���ۼ� >= Nvl(!����, 0) And cur����ͳ���ۼ� < Nvl(!����, 0)) Or Nvl(!����, 0) = 0 Then
                    int��ʼ�� = !����
                    If cur����ͳ���ۼ� <> 0 Then
                        'ֻ�н���ͳ���ۼƲ�Ϊ�㣬�Ž��м��㣬����cur������뵵�ζ�Ϊ�㣬��ʾȫ��
                        cur������뵵�� = IIf(Nvl(!����, 0) = 0, 0, (Nvl(!����, 0) - Nvl(!����, 0)) - (cur����ͳ���ۼ� - Nvl(!����, 0)))
                    End If
                End If
            End If
            
            rs�ֵ�֧��_����.AddNew
            rs�ֵ�֧��_����!���� = !����
            rs�ֵ�֧��_����!���� = !����
            rs�ֵ�֧��_����!���� = !����
            rs�ֵ�֧��_����.Update
            .MoveNext
        Loop
        
        '���������ʵ��ͳ����
        blnExit = False
        .MoveFirst
        .Find "����=" & int��ʼ��
        Do While Not .EOF
            If (blnExit Or curʣ��ͳ�� <= 0) Then Exit Do
            rs�ֵ�֧��_����.MoveFirst
            rs�ֵ�֧��_����.Find "����=" & !����
            
            If !���� <> int��ʼ�� Then cur������뵵�� = 0
            If !���� <> 0 Then
                If curʣ��ͳ�� + cur����ͳ���ۼ� > !���� Then
                    '(1)
                    cur����ͳ�� = !���� - !����
                Else
                    '(2)
                    cur����ͳ�� = (curʣ��ͳ�� + cur����ͳ���ۼ�) - Nvl(!����, 0)
                    blnExit = True
                End If
            Else
                'ȫ������(3)
                cur����ͳ�� = (curʣ��ͳ�� + cur����ͳ���ۼ�) - Nvl(!����, 0)
                blnExit = True
            End If
            '������벿�ִ��ڱ����ܵĽ���ͳ�����벿�ֵ��ڱ����ܵĽ���ͳ�ﲿ��-У��(2)
            If cur����ͳ�� > curʣ��ͳ�� Then cur����ͳ�� = curʣ��ͳ��
            '����������뵵�ν���ʾ���������-У��(2),(3)
            If cur����ͳ�� > cur������뵵�� And cur������뵵�� <> 0 Then
                cur����ͳ�� = cur������뵵��
            End If
            '�������ͳ��С���㣬���˳����ٽ��к���ļ���û������
            If cur����ͳ�� <= 0 Then
                cur����ͳ�� = 0
                blnExit = True
            End If
            cur���α��� = cur����ͳ�� * !���� / 100
            curͳ��֧�� = curͳ��֧�� + CCur(Format(cur���α���, "#####0.00;-#####0.00;0;"))
            rs�ֵ�֧��_����!����ͳ�� = cur����ͳ��
            rs�ֵ�֧��_����!ͳ�ﱨ�� = cur���α���
            rs�ֵ�֧��_����.Update

            .MoveNext
        Loop
    End With
    gComInfo_üɽ.ͳ���Ը� = (curʣ��ͳ�� - curͳ��֧��) + (gComInfo_üɽ.����ʵ�ʱ��� - gComInfo_üɽ.ʵ�ʱ���)
    gComInfo_üɽ.ͳ��֧�� = curͳ��֧�� + gComInfo_üɽ.ʵ�ʱ���
    If gComInfo_üɽ.ͳ��֧�� > 0 Then
        gComInfo_üɽ.����ͳ�� = curʣ��ͳ�� + gComInfo_üɽ.����ʵ�ʱ���
    Else
        gComInfo_üɽ.����ͳ�� = 0
    End If
    
    Call FormatData
End Sub

Private Function CopyNewRec(ByVal SourceRec As ADODB.Recordset) As ADODB.Recordset
    Dim RecTarget As New ADODB.Recordset
    Dim intFields As Integer, LngLocate As Long
    '������:����
    '��������:2000-11-02
    '�ü�¼����ƾ֤�ؼ���Ӧ
    'Ҳʹ���ڱ���
    
    LngLocate = -1
    Set RecTarget = New ADODB.Recordset
    With RecTarget
        If .State = 1 Then .Close
        If SourceRec.RecordCount <> 0 Then
            On Error Resume Next
            Err = 0
            LngLocate = SourceRec.AbsolutePosition
            If Err <> 0 Then LngLocate = -1
            SourceRec.MoveFirst
        End If
        For intFields = 0 To SourceRec.Fields.Count - 1
            .Fields.Append SourceRec.Fields(intFields).Name, SourceRec.Fields(intFields).Type, SourceRec.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
        Do While Not SourceRec.EOF
            .AddNew
            For intFields = 0 To SourceRec.Fields.Count - 1
                .Fields(intFields) = SourceRec.Fields(intFields).Value
            Next
            .Update
            SourceRec.MoveNext
        Loop
    End With
    
    If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
    If LngLocate > 0 Then SourceRec.Move LngLocate - 1
    Set CopyNewRec = RecTarget
End Function

Public Function Encrypt(ByVal strMoney As String, ByVal strCardNO As String) As String
    Dim intLen As Integer, LngProcess As Long
    Dim strTmp As String, strTmp_Source As String, strTmp_Target As String, strTmp_CardNO As String
    Dim strEncrypt As String
    '���ܽ���
    
    Encrypt = ""
    If Val(strMoney) = 0 Then Exit Function
    
    strEncrypt = "thisisajokebyzybzl"
    For intLen = 1 To Len(strMoney)
        strTmp_Source = Mid(strMoney, intLen, 1)
        strTmp_Target = Mid(strEncrypt, intLen, 1)
        If intLen Mod Len(strCardNO) = 0 Then
            strTmp_CardNO = Mid(strCardNO, intLen, 1)
        Else
            strTmp_CardNO = Mid(strCardNO, intLen Mod Len(strCardNO), 1)
        End If
        LngProcess = asc(strTmp_Source) Xor asc(strTmp_Target) Xor asc(strTmp_CardNO)
        
        If LngProcess < 32 Then
            LngProcess = LngProcess + 32
        ElseIf LngProcess > 127 Then
            LngProcess = LngProcess - (LngProcess - 107)
        End If
        
        If LngProcess = 34 Then
            Encrypt = Encrypt & """"
        ElseIf LngProcess = 39 Then
            Encrypt = Encrypt & "''"
        Else
            Encrypt = Encrypt & Chr(LngProcess)
        End If
    Next
End Function

Public Function ����ʻ���Ϣ_����(ByVal strCardNO As String, Optional ByVal blnUpdate As Boolean = False, Optional ByVal bln���� As Boolean = True) As Boolean
    Dim str�Ƚϴ� As String, lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select Nvl(�ʻ����,0)*100 ���,����,���ܴ�,����ID From �����ʻ� Where ����=[1]"
    If bln���� Then
        gstrSQL = gstrSQL & " And ҽ����=[2]"
    Else
        gstrSQL = gstrSQL & " And ����ID=[3]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ʻ���Ϣ", TYPE_�Ĵ�üɽ, strCardNO, CLng(Val(strCardNO)))
    If Not rsTemp.EOF Then
        strCardNO = rsTemp!����
        lng����ID = rsTemp!����ID
        str�Ƚϴ� = Encrypt(rsTemp!���, Nvl(rsTemp!����))
        If str�Ƚϴ� <> Nvl(rsTemp!���ܴ�) Then
            If Not blnUpdate Then
                MsgBox "�����ڷǷ��޸�ҽ�����˵ĸ����ʻ������飡", vbInformation, gstrSysName
                Exit Function
            Else
                gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & ",25,'���ܴ�','''" & str�Ƚϴ� & "''')"
                gcnOracle.Execute gstrSQL
            End If
        End If
    End If
    ����ʻ���Ϣ_���� = True
End Function

Public Function �¸����ʻ�(ByVal lng����ID As Long, ByVal cur��� As Currency) As Boolean
    Dim lngNextID As Long
    Dim rsBalance As New ADODB.Recordset
    
    On Error GoTo errHand
    If Not ����ʻ���Ϣ_����(lng����ID, False, False) Then Exit Function
    
    lngNextID = zlDatabase.GetNextID("�ʻ��䶯��¼")
    gstrSQL = "ZL_�ʻ��䶯��¼_INSERT(" & lngNextID & "," & TYPE_�Ĵ�üɽ & ",2," & lng����ID & "," & _
                cur��� & ",'" & gstrUserName & "','����" & IIf(cur��� > 0, "����", "֧��") & "',0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�¸����ʻ�")
    
    If Not ����ʻ���Ϣ_����(lng����ID, True, False) Then Exit Function
    �¸����ʻ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


