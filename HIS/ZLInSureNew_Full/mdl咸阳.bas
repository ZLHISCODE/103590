Attribute VB_Name = "mdl����"
Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'-------------��������
Public gcn���� As New ADODB.Connection        '���ӵ�ҽ��ǰ�÷�����

'-------------��������

Public Function ҽ����ʼ��_����() As Boolean
'���ܣ�����ҽ��ǰ�û�������
'���أ���ʼ���ɹ�������true�����򣬷���false
    
    ҽ����ʼ��_���� = ���ҽ��������_����
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������strSelfNO-���˱�ţ�ˢ���õ���strSelfPwd-�������룻bytType-ʶ�����ͣ�0-���1-סԺ
'���أ� �ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim strҽ���� As String, rsTemp As New ADODB.Recordset
    Dim STR���� As String, str�Ա� As String, str���֤���� As String, lng���� As Long
    Dim str�������� As String, str��λ���� As String
    Dim strIdentify As String, str���� As String, strComputer As String
    
    Dim cur�����ʻ� As Currency
   
    On Error GoTo errHandle
    
    '��ǰ�÷������ж����Ѿ���֤��ݵĲ�����Ϣ
    If bytType = 0 Then
        '������Ϣ
        strComputer = Get������
'        strComputer = "work1"
        gstrSQL = "SELECT CardNo AS ҽ������, UnitNo AS ��λ����, SelfSerial AS �������,PatiName As ���� " & _
                   "     ,PatiSex As �Ա�, 0 AS ����, '����' As ����, IdentityCard As ���֤��,balance as �����ʻ���� " & _
                   " From Outpatients " & _
                   "WHERE Terminal = '" & strComputer & "' AND AcceptTime IS NULL"
    ElseIf bytType = 1 Then
        'סԺ��Ϣ
        gstrSQL = "SELECT CardNo AS ҽ������, '' AS ��λ����, StaySerial As �������,PatiName AS ����" & _
                  "       ,PatiSex AS �Ա�, PatiYear AS ����, PatiFolk As ����, IdentityCard As ���֤�� " & _
                  "  From Inpatients " & _
                   "WHERE AcceptTime IS NULL"
    Else
        '��֧��
        Exit Function
    End If
    
    rsTemp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount = 0 Then
        MsgBox "δ��������֤������Ϣ��", vbInformation, gstrSysName
        Exit Function
    ElseIf rsTemp.RecordCount > 1 Then
        '����һ��ʱҪ����ѡ��
        If frmListSel.ShowSelect(TYPE_������, rsTemp, "ҽ������", "ҽ������ѡ��", "��ѡ������֤�Ĳ��ˣ�Ȼ����ȷ����") = False Then
            Exit Function
        End If
    End If
    
    strҽ���� = rsTemp("ҽ������")
    
    If bytType = 0 Then
        cur�����ʻ� = rsTemp("�����ʻ����")
    Else
        cur�����ʻ� = 0
    End If
    
    STR���� = rsTemp("����")
    str�Ա� = IIf(IsNull(rsTemp("�Ա�")), "", rsTemp("�Ա�"))
    str���֤���� = IIf(IsNull(rsTemp("���֤��")), "", rsTemp("���֤��"))
    str�������� = Get��������(str���֤����, 0)
    If IsDate(str��������) Then
        lng���� = DateDiff("yyyy", CDate(str��������), zlDatabase.Currentdate)
        str�������� = Format(CDate(str��������), "yyyy-MM-dd")
    Else
        lng���� = IIf(IsNull(rsTemp("����")), 0, Val(rsTemp("����")))
        str�������� = Format(DateAdd("yyyy", -1 * lng����, zlDatabase.Currentdate), "yyyy-MM-dd") '�����䵹������
    End If
    
    str��λ���� = IIf(IsNull(rsTemp("��λ����")), "", rsTemp("��λ����"))
    
    
    '����;ҽ����;����;����;�Ա�;��������;���֤;������λ
    strIdentify = strҽ���� & ";" & strҽ���� & ";;" & STR���� & ";" & str�Ա� & ";" & str�������� & ";" & str���֤���� & ";(" & str��λ���� & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    str���� = ";"                                       '8.���Ĵ���
    str���� = str���� & ";" & IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))             '9.˳���
    str���� = str���� & ";"                             '10��Ա���
    str���� = str���� & ";" & cur�����ʻ�               '11�ʻ����
    str���� = str���� & ";0"                            '12��ǰ״̬
    str���� = str���� & ";"                             '13����ID
    str���� = str���� & ";1"                            '14��ְ(1,2)
    str���� = str���� & ";"                             '15����֤��
    str���� = str���� & ";" & lng����                   '16�����
    str���� = str���� & ";"                             '17�Ҷȼ�
    str���� = str���� & ";" & cur�����ʻ�              '18�ʻ������ۼ�
    str���� = str���� & ";0"                            '19�ʻ�֧���ۼ�
    str���� = str���� & ";"                             '20����ͳ���ۼ�
    str���� = str���� & ";"                             '21ͳ�ﱨ���ۼ�
    str���� = str���� & ";"                             '22סԺ�����ۼ�
    
    lng����ID = BuildPatiInfo(bytType, strIdentify & str����, lng����ID, TYPE_������)
    
    '���ظ�ʽ:�м���벡��ID
    If lng����ID <> 0 Then
        ��ݱ�ʶ_���� = strIdentify & ";" & lng����ID & str����
        
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����:
'����: ���ظ����ʻ����Ľ��
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '�����ݿ��ж�ȡ����Ϊ�ղŲű����˵ģ�Ӧ����׼ȷ�ģ�
    gstrSQL = "Select �ʻ���� From �����ʻ� where ����=[1] and ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������, lng����ID)

    If rsTemp.EOF = False Then
        �������_���� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
    Else
        �������_���� = 0
    End If
    
    
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��

    Dim str������ As String, rsTemp As New ADODB.Recordset
    Dim strҽ���� As String, STR���� As String, cur�����ʻ� As Currency, dat����ʱ�� As Date
    
    Dim strmachine As String
    
    On Error GoTo errHandle
    
    strmachine = Get������()
    
    
    ''''''''
'    strmachine = "WORK1"
    
    ''''''''''''''''''''''
    
    '���ҽ����
    gstrSQL = "select B.ҽ����,C.���� " & _
              "  from �����ʻ� B,������Ϣ C " & _
              "  where B.����id=[1] and B.����=[2] and B.����ID=C.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", CLng(rs��ϸ("����ID")), TYPE_������)
    strҽ���� = rsTemp("ҽ����")
    STR���� = rsTemp("����")
   dat����ʱ�� = zlDatabase.Currentdate
    
    '�õ��ò��˵�ҽ��������
    str������ = InputBox("������ҽ������ר�ô����ţ�", "����Ԥ��")
    If str������ = "" Then
        Exit Function
    End If
            
    If zlCommFun.StrIsValid(str������, 7) = False Then
        MsgBox "�������к��зǷ��ַ���", vbInformation, gstrSysName
        Exit Function
    End If
    If Len(str������) < 7 Then
        MsgBox "�����ų��Ȳ���7λ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    
  
    
    '������ɷ�����ϸ�Ĵ���
    
    On Error Resume Next
    gcn����.BeginTrans
    
    '��������������ʱ�����ã������ò��˵õ��������ﲡ�˵������ǣ�ҽ������+����ʱ�䣩
    gstrSQL = "UPDATE Outpatients Set AcceptTime = GETDATE() WHERE Terminal='" & strmachine & _
                "' and cardno = '" & strҽ���� & "' and accepttime is null "
    gcn����.Execute gstrSQL
    
    gcn����.Execute "Delete from ClinicExses where CardNo='" & strҽ���� & "' and AcceptTime is null"
    gcn����.Execute "Delete ClinicSettles Where CardNo='" & strҽ���� & "' and AcceptTime is null"
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.����,A.����,A.���㵥λ,A.���,B.���� as ����,B.ͳ��ȶ�  " & _
                   " from �շ�ϸĿ A , " & _
                   " (select B. �շ�ϸĿID,C.����,C.ͳ��ȶ� from ����֧������ C,����֧����Ŀ B where B.����=" & TYPE_������ & " and B.����ID=C.ID) B " & _
                   " where A.ID=B.�շ�ϸĿID(+) and A.ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", CLng(rs��ϸ("�շ�ϸĿID")))
        If IsNull(rsTemp("����")) = True Then
            MsgBox "�շ���Ŀ��" & rsTemp("����") & "��û�����ö�Ӧ�ı��մ��ࡣ", vbInformation, gstrSysName
            gcn����.RollbackTrans
            Exit Function
        End If
        
        gstrSQL = "INSERT INTO ClinicExses(ID, RecipeNo, CardNo, PatiName, RecordTime, ItemCode, ItemName, ItemUnit, Price, Amount, Money, ExseKind, InsureKind, PayTax) " & _
                  "VALUES('" & str������ & "_" & rs��ϸ.AbsolutePosition & "','" & str������ & "','" & strҽ���� & "','" & STR���� & "','" & _
                  Format(dat����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','" & ToVarchar(rsTemp("����"), 10) & "','" & ToVarchar(rsTemp("����"), 90) & "','" & ToVarchar(rsTemp("���㵥λ"), 6) & "'," & _
                  Format(rs��ϸ("����"), "0.0000") & "," & Format(rs��ϸ("����"), "0.000") & "," & Format(rs��ϸ("ʵ�ս��"), "0.000") & ",'" & _
                  Switch(rsTemp("���") = "7", "0", rsTemp("���") = "5", "1", rsTemp("���") = "6", "2", rsTemp("���") = "J", "4", True, 3) & "','" & _
                  IIf(rsTemp("ͳ��ȶ�") = 100, "0", IIf(rsTemp("ͳ��ȶ�") = 0, "2", "1")) & "'," & Format(100 - rsTemp("ͳ��ȶ�"), "0;;\0") & ")"
        gcn����.Execute gstrSQL
        If Err <> 0 Then
            MsgBox "ҽ��������ϸ����ʧ�ܡ�", vbInformation, gstrSysName
            gcn����.RollbackTrans
            Exit Function
        End If
        
        rs��ϸ.MoveNext
    Loop
    gcn����.CommitTrans
    
    On Error GoTo errHandle
    '��ȡǰ�÷������Ƿ���ɽ���
    gstrSQL = "SELECT CardPay AS ����֧��, CashPay AS �ֽ�֧��, TotalExse AS ���úϼ� " & _
              " FROM ClinicSettles WHERE CardNo = '" & strҽ���� & "' AND AcceptTime IS NULL"
    If frm�ȴ�����.WaitForYB(rsTemp, gstrSQL) = False Then Exit Function
    If rsTemp.EOF = True Then
        cur�����ʻ� = 0
        Exit Function
    Else
        cur�����ʻ� = IIf(IsNull(rsTemp("����֧��")), 0, rsTemp("����֧��"))
    End If
    
    '�ı�ǰ�÷���������������¼����ʱ�䣨����������������ǣ�ҽ������+����ʱ�䣩
    gstrSQL = "UPDATE ClinicSettles Set AcceptTime = GETDATE() WHERE cardno = '" & strҽ���� & "' and accepttime is null"
    gcn����.Execute gstrSQL
    
    str���㷽ʽ = "�����ʻ�;" & cur�����ʻ� & ";0"  '�����޸ĸ����ʻ�
    �����������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim strҽ���� As String, StrInput As String, arrOutput  As Variant
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset
    Dim str����Ա As String, cur��������, datCurr As Date
    
    Dim rs�ʻ���� As New Recordset
    Dim cur�ʻ���� As Currency
    
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From ������ü�¼ Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    If rs��ϸ.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    str����Ա = ToVarchar(IIf(IsNull(rs��ϸ("����Ա����")), UserInfo.����, rs��ϸ("����Ա����")), 20)
    
    Do Until rs��ϸ.EOF
        cur�������� = cur�������� + rs��ϸ("���ʽ��")
        rs��ϸ.MoveNext
    Loop
    
    '���ý���
    
    '��������¼
    '---------------------------------------------------------------------------------------------
    datCurr = zlDatabase.Currentdate
    
    
    gstrSQL = "select �ʻ���� from �����ʻ� where ����=[1] and ����id=[2]"
    Set rs�ʻ���� = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������, lng����ID)
    
    If rs�ʻ����.EOF Then
        cur�ʻ���� = 0
    Else
        cur�ʻ���� = rs�ʻ����.Fields(0)
    End If
    
    
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ���� & ",0,0,0,0,0,0,0," & cur�������� & ",0,0," & _
        "0,0,0,0," & cur�����ʻ� & ",'')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    '---------------------------------------------------------------------------------------------

    �������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function


Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    
    ����������_���� = False
End Function

Public Function �����ʻ�תԤ��_����(lngԤ��ID As Long, curMoney As Currency) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
           
    '��������ҽ����֧�ָ�ҵ������ǿ�з���ʧ��
    
    �����ʻ�תԤ��_���� = False
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    '��ò���ҽ����
    gstrSQL = "select ҽ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, TYPE_������)
    strҽ���� = rsTemp("ҽ����")
    
    'ȷ���Ѿ���Ժ�Ǽǳɹ�
    gstrSQL = "UPDATE Inpatients Set AcceptTime = GETDATE() WHERE cardno = '" & strҽ���� & "' and accepttime is null "
    gcn����.Execute gstrSQL
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(ByVal lng����ID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'���أ����׳ɹ�����true�����򣬷���false
    
    On Error GoTo errHandle
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long, ByVal strҽ���� As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼����
    '��¼����,��¼״̬,NO����š��շ�����շ����ơ��������š���񡢲��ء��������۸񡢽�ҽ��,
    '�Ǽ�ʱ��,Ӥ����,ҽ����Ŀ���롢���մ���ID��������Ŀ���Ƿ��ϴ�
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    Dim rsTemp As New ADODB.Recordset, rs���� As New ADODB.Recordset
    Dim STR���� As String, dat����ʱ�� As Date
    Dim curͳ��֧�� As Currency, cur�����ʻ� As Currency, cur�������� As Currency
    Dim blnReturn As Boolean        '�Ƿ�ɹ���ȡҽ����������
    Dim str��Ŀ���� As String
    Dim str�������� As String
    
    On Error GoTo errHandle
    
    '���ҽ����
    gstrSQL = "select B.ҽ����,C.���� " & _
              "  from �����ʻ� B,������Ϣ C " & _
              "  where B.����id=[1] and B.����=[2] and B.����ID=C.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", lng����ID, TYPE_������)
    STR���� = rsTemp("����")
    dat����ʱ�� = zlDatabase.Currentdate
    
    '������ɷ�����ϸ�Ĵ���
    gstrSQL = "select ID,����,����,ͳ��ȶ�,��׼���� FROM ����֧������  where ����=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������)
    
    On Error Resume Next
    gcn����.BeginTrans
    gcn����.Execute "Delete from InpatiExses where CardNo='" & strҽ���� & "' and AcceptTime is null"
    Do Until rsExse.EOF
        cur�������� = cur�������� + rsExse("���")
        rs����.Filter = "ID=" & rsExse("���մ���ID")
        
        If rs����.EOF = True Then
            MsgBox "�շ���Ŀ��" & rsExse("�շ�����") & "��û�����ö�Ӧ�ı��մ��ࡣ", vbInformation, gstrSysName
            gcn����.RollbackTrans
            Exit Function
        End If
        
        '�жϲ��˷��ü�¼���Ƿ��ϴ���־�Ƿ���Ϊ0�����Ϊ0����ʾû���ϴ��������򣬱�ʾ���ϴ����������ϴ�
        If rsExse("�Ƿ��ϴ�") = 0 Then
            '���ڶԷ����յ���Ŀ����������ϵͳ���շ�ϸĿ��Ӧ�ı��룬�����ǶԷ�ҽ���ı��룬��һ��������ҽ��������ҽ��������֮һ�����ԣ�����ȥȡ����ϵͳ���շ�ϸĿ�ı���
            gstrSQL = "select distinct ���� from �շ�ϸĿ where ����=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", CStr(rsExse("�շ�����")))
            
            If rsTemp.EOF Then
                str��Ŀ���� = "001"
            Else
                str��Ŀ���� = ToVarchar(rsTemp.Fields("����"), 10)
            End If
            
            gstrSQL = "select b.���� " _
                    & " From " _
                    & " (select ������Ŀid from סԺ���ü�¼ " _
                    & "where no =[1] and ��¼����=[2]" & _
                     " and ���=[3]" & _
                     " and ��¼״̬=[4]" & _
                     " and ����id=[5])  a,������Ŀ b " & _
                     " Where a.������Ŀid = b.ID "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", CStr(rsExse("NO")), CLng(rsExse("��¼����")), CLng(rsExse("���")), CLng(rsExse("��¼״̬")), CLng(rsExse("����id")))
            
            If rsTemp.EOF Then
                str�������� = "001"
            Else
                str�������� = ToVarchar(rsTemp.Fields("����"), 10)
            End If
            
            gstrSQL = rsExse("NO") & "_" & rsExse("���") & "_" & rsExse("��¼����") & "_" & rsExse("��¼״̬")
            gstrSQL = "INSERT INTO InpatiExses(ID, CardNo, PatiName, RecordTime, ItemCode, ItemName, ItemUnit, Price, Amount, Money, ExseKind, InsureKind, PayTax) " & _
                      "VALUES('" & gstrSQL & "','" & strҽ���� & "','" & STR���� & "','" & _
                      Format(rsExse("����ʱ��"), "yyyy-MM-dd HH:mm:ss") & "','" & str��Ŀ���� & "','" & ToVarchar(rsExse("�շ�����"), 90) & "','��'," & _
                      Format(rsExse("�۸�"), "0.0000") & "," & Format(rsExse("����"), "0.000") & "," & Format(rsExse("���"), "0.000") & ",'" & _
                      str�������� & "','" & _
                      IIf(rs����("ͳ��ȶ�") = 100, "0", IIf(rs����("ͳ��ȶ�") = 0, "2", "1")) & "'," & Format(100 - rs����("ͳ��ȶ�"), "0;;\0") & ")"
            gcn����.Execute gstrSQL
        End If
        rsExse.MoveNext
        If Err <> 0 Then
            MsgBox "ҽ��������ϸ����ʧ�ܡ�", vbInformation, gstrSysName
            gcn����.RollbackTrans
            Exit Function
        End If
    Loop
    gcn����.CommitTrans
    
    On Error GoTo errHandle
    '��ȡǰ�÷������Ƿ���ɽ���
    gstrSQL = "SELECT AgentPay as ͳ��֧��, CardPay AS ����֧��, CashPay AS �ֽ�֧��, TotalExse AS ���úϼ�,FlowExse as ���޽�� " & _
              " FROM InpatiSettles WHERE CardNo = '" & strҽ���� & "' AND AcceptTime IS NULL"
    blnReturn = frm�ȴ�����.WaitForYB(rsTemp, gstrSQL)
    If blnReturn Then blnReturn = (rsTemp.RecordCount <> 0)
    
    If blnReturn = False Then
        cur�����ʻ� = 0
        curͳ��֧�� = 0
        With g��������
            .����ID = lng����ID
            .ͳ�ﱨ����� = curͳ��֧��
            .�����ʻ�֧�� = cur�����ʻ�
            .�������ý�� = cur��������
            
            .�����Ը���� = 0
        End With
        Exit Function
    Else
        cur�����ʻ� = IIf(IsNull(rsTemp("����֧��")), 0, rsTemp("����֧��"))
        curͳ��֧�� = IIf(IsNull(rsTemp("ͳ��֧��")), 0, rsTemp("ͳ��֧��"))
        With g��������
            .����ID = lng����ID
            .ͳ�ﱨ����� = curͳ��֧��
            .�����ʻ�֧�� = cur�����ʻ�
            .�������ý�� = cur��������
            
            .�����Ը���� = IIf(IsNull(rsTemp("���޽��")), 0, rsTemp("���޽��"))
        End With
    End If
    
    סԺ�������_���� = "ҽ������;" & curͳ��֧�� & ";0"
    If cur�����ʻ� > 0 Then
        סԺ�������_���� = סԺ�������_���� & "|�����ʻ�;" & cur�����ʻ� & ";0"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID     ���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset, strҽ���� As String
    Dim datCurr As Date
    
    If g��������.����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣"
        Exit Function
    End If
    
    On Error GoTo errHandle
    '���ҽ����
    gstrSQL = "select B.ҽ����,C.���� " & _
              "  from �����ʻ� B,������Ϣ C " & _
              "  where B.����id=[1] and B.����=[2] and B.����ID=C.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", lng����ID, TYPE_������)
    strҽ���� = rsTemp("ҽ����")
    
    
    '���ý���
    '�ı�ǰ�÷���������������¼����ʱ��
    gstrSQL = "UPDATE InpatiSettles Set AcceptTime = GETDATE() WHERE cardno = '" & strҽ���� & "' and accepttime is null"
    gcn����.Execute gstrSQL
    
    
    '���±��ؽ�������з�����ϸ���ϴ���־
    gstrSQL = "Update סԺ���ü�¼ Set �Ƿ��ϴ�=1 Where Nvl(�Ƿ��ϴ�,0)=0 And ����ID=" & lng����ID
    gcnOracle.Execute gstrSQL
    
    '��д�����
    datCurr = zlDatabase.Currentdate
    
    With g��������
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
            Year(datCurr) & ",0,0,0,0,0,0,NULL,0," & .�������ý�� & ",0,0," & _
            .ͳ�ﱨ����� & "," & .ͳ�ﱨ����� & ",0,0," & .�����ʻ�֧�� & ",'',0,0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        '���ս������
        gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & .ͳ�ﱨ����� & "," & .ͳ�ﱨ����� & ",NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End With
        
    סԺ����_���� = True
    Exit Function
errHandle:
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
    
    סԺ�������_���� = False
End Function

Public Function ���ҽ��������_����() As Boolean
'���ܣ������Ƿ�������ӵ�ǰ�÷�������
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String, strDatabase As String
    
    '��������Ѿ��򿪣��ǾͲ����ٲ���
    If gcn����.State = adStateOpen Then
        ���ҽ��������_���� = True
        Exit Function
    End If
    
    On Error GoTo ErrH
    
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ���ӿ�", TYPE_������)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽ��������"
                strServer = strTemp
            Case "ҽ�����ݿ�"
                strDatabase = strTemp
            Case "ҽ���û���"
                strUser = strTemp
            Case "ҽ���û�����"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    gcn����.Open "Provider=SQLOLEDB.1;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & _
                ";Initial Catalog=" & strDatabase & ";Data Source=" & strServer
    If Err <> 0 Then
        Err.Raise 9000, gstrSysName, "ҽ��ǰ�÷���������ʧ�ܡ�"
        Exit Function
    End If
    
    ���ҽ��������_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Private Function Get������() As String
'���ܣ���õ�ǰ�Ļ�����
    Dim STRNAME As String, l As Long
    
    STRNAME = Space(256): l = 256
    
    If GetComputerName(STRNAME, l) <> 0 Then
        Get������ = TrimStr(STRNAME)
    End If
End Function


