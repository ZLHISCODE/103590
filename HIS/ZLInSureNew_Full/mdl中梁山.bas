Attribute VB_Name = "mdl����ɽ"
Option Explicit

Public Function ҽ����ʼ��_����ɽ() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false

    
    'Ϊ�˱�����Ȩ�Ѷ����ӣ��˴����ٽ��жԸ���ҽ�������ݵļ��
    ҽ����ʼ��_����ɽ = True
End Function

Public Function ��ݱ�ʶ_����ɽ(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    
    Dim strTmpIden As String
    
    strTmpIden = frmIdentify����ɽ.ShowCard(bytType, lng����ID)
    ��ݱ�ʶ_����ɽ = strTmpIden
End Function

Public Function �������_����ɽ(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: bytYear-�������,0-�������,1-�������,2-�������
'����: ���ظ����ʻ����Ľ��
    
    '��ʹ�ø����ʻ�
    �������_����ɽ = 0
End Function

Public Function �������_����ɽ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, _
            ByVal curȫ�Է� As Currency, ByVal cur�����Ը� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
    
    �������_����ɽ = False
End Function


Public Function ����������_����ɽ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��

    ����������_����ɽ = False
End Function

Public Function �����ʻ�תԤ��_����ɽ(lngԤ��ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    
    �����ʻ�תԤ��_����ɽ = False
End Function


Public Function �����ʻ�תԤ������_����ɽ(lngԤ��ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false

    �����ʻ�תԤ������_����ɽ = False
End Function

Public Function ��Ժ�Ǽ�_����ɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false

    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��������ɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    ��Ժ�Ǽ�_����ɽ = True
End Function

Public Function ��Ժ�Ǽ�_����ɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��������ɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    ��Ժ�Ǽ�_����ɽ = True
End Function

Public Function ��Ժ�Ǽǳ���_����ɽ(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
            'ȡ��Ժ�Ǽ���֤�����ص�˳���
    On Error GoTo errHandle
    
        
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��������ɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    ��Ժ�Ǽǳ���_����ɽ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����ɽ(rs������ϸ As Recordset, ByVal lng����ID As Long) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����Ҫ��NO����š�����ID��ҽ����Ŀ���롢�շ�����շ����ơ��������š���񡢲��ء��������۸񡢽�ҽ��,�Ǽ�ʱ��(����ʱ��),Ӥ����,���մ���ID
    Dim rs��׼��Ŀ As New ADODB.Recordset
    Dim rs�㷨 As New ADODB.Recordset          '����
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng���� As Long
    Dim lng��ְ As Long, lng����� As Long, lng���� As Long
    
    Dim dbl�����  As Double ''��һ����סԺ�ռ������Ŀ������ܵõ��Ľ��
    
    Dim curȫ�Է� As Currency, cur�����Ը� As Currency, cur����ͳ�� As Currency, dblTemp As Double
    Dim blnȫ��ͳ�� As Boolean, bln�޷ⶥ�� As Boolean, blnĿ¼ As Boolean, bln��������Ŀ¼ As Boolean
    
    
    '������������������������������������������������������������������������������������
    '1����ʼ��һЩ����
    With g��������
        .�����Ը���� = 0
        .�������ý�� = 0
        .�ⶥ�� = 0
        .����ͳ���� = 0
        .�ۼƽ���ͳ�� = 0
        .�ۼ�ͳ�ﱨ�� = 0
        .ȫ�Էѽ�� = 0
        .�����Ը���� = 0
        .ͳ�ﱨ����� = 0
        .ʵ������ = 0    '���ڱ������ȫ��Ŀ¼�Ľ��
        
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
    
    '1.2 �������˵���Ժʱ��
    gstrSQL = "select ��Ժ����,nvl(��Ժ����,to_date('3000-01-01','yyyy-MM-dd')) as ��Ժ���� " & _
              "from ������ҳ where ����ID=" & g��������.����ID & " and ��ҳID=" & g��������.��ҳID
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp("��Ժ����") = CDate("3000-01-01") Then
        g��������.��;���� = 1
    Else
        '��ʾ�ò����Ѿ���Ժ
        g��������.��;���� = 0
    End If

    '1.3 ��������סԺ�ڼ��ۼƽ������
    With g��������
        gstrSQL = "select A.����,A.��Ա���,A.��ְ,A.�����," & _
                  "      B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ�" & _
                  " from �����ʻ� A,�ʻ������Ϣ B" & _
                  " where A.����ID=B.����ID(+) and A.����=B.����(+) " & _
                  "     and B.���(+)=" & .��� & " and A.����ID=" & .����ID & " and A.����=" & TYPE_��������ɽ
        Call OpenRecordset(rsTemp, "�������")
        
        lng���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
        lng��ְ = IIf(IsNull(rsTemp("��ְ")), 1, rsTemp("��ְ"))
        lng���� = IIf(IsNull(rsTemp("�����")), 0, rsTemp("�����"))
        .סԺ���� = IIf(IsNull(rsTemp("סԺ�����ۼ�")), 0, rsTemp("סԺ�����ۼ�"))
        .�ۼƽ���ͳ�� = IIf(IsNull(rsTemp("����ͳ���ۼ�")), 0, rsTemp("����ͳ���ۼ�"))
        .�ۼ�ͳ�ﱨ�� = IIf(IsNull(rsTemp("ͳ�ﱨ���ۼ�")), 0, rsTemp("ͳ�ﱨ���ۼ�"))
    
        gstrSQL = "select �����,nvl(ȫ��ͳ��,0) as ȫ��ͳ�� ,nvl(������,0) as ������ ,nvl(�޷ⶥ��,0) as �޷ⶥ�� " & _
                " from ���������" & _
                " where ����=" & TYPE_��������ɽ & " and nvl(����,0)=" & lng���� & _
                "       and ��ְ=" & lng��ְ & " and ����<=" & lng���� & " and (" & lng���� & "<=���� or ����=0)"
        Call OpenRecordset(rsTemp, "�������")
        If rsTemp.RecordCount = 0 Then
            MsgBox "���ڡ��������������������������õ���", vbInformation, gstrSysName
            Exit Function
        End If
        lng����� = rsTemp("�����")
        blnȫ��ͳ�� = (rsTemp("ȫ��ͳ��") = 1)
        bln�޷ⶥ�� = (rsTemp("�޷ⶥ��") = 1)
    End With
    
    '������������������������������������������������������������������������������������
    '2����ͳ��֧����Ŀ�ϼƷ�����������
    '2.1�����ݲ��˲��֣��ж��Ƿ���ȫ������Ŀ
    '2.2���������ͳ����
    If blnȫ��ͳ�� = False Then
        gstrSQL = "SELECT B.�շ�ϸĿID " & _
                 " FROM �����ʻ� A,������׼��Ŀ B,�շ�ϸĿ C " & _
                 " WHERE A.����ID=" & g��������.����ID & " AND A.����=" & TYPE_��������ɽ & " AND A.����ID=B.����ID And B.�շ�ϸĿID=C.ID and C.��� not in ('5','6','7') and Rownum<2"
        Call OpenRecordset(rsTemp, "סԺ����")
        If rsTemp.RecordCount > 0 Then bln��������Ŀ¼ = True '���ѡ����������Ŀ������Ϊ��׼��Ŀ�嵥��Ч
        
        gstrSQL = "SELECT B.�շ�ϸĿID " & _
                 " FROM �����ʻ� A,������׼��Ŀ B " & _
                 " WHERE A.����ID=" & g��������.����ID & " AND A.����=" & TYPE_��������ɽ & " AND A.����ID=B.����ID"
        Call OpenRecordset(rs��׼��Ŀ, "סԺ����")
        
        gstrSQL = "select ID,�㷨,ͳ��ȶ�,��׼����,��׼����,�Ƿ�ҽ�� FROM ����֧������  where ����=" & TYPE_��������ɽ
        Call OpenRecordset(rs�㷨, "סԺ����")
        
        dblTemp = 0
        If rs������ϸ.RecordCount > 0 Then rs������ϸ.MoveFirst
        Do Until rs������ϸ.EOF
            blnĿ¼ = False
            rs��׼��Ŀ.Filter = "�շ�ϸĿID=" & rs������ϸ("�շ�ϸĿID")
            
            If rs��׼��Ŀ.RecordCount = 0 Then
                'û�������ض���Ŀ
                If rs������ϸ("�շ����") = "5" Or rs������ϸ("�շ����") = "6" Or rs������ϸ("�շ����") = "7" Then
                    'ҩƷ
                    blnĿ¼ = False
                ElseIf bln��������Ŀ¼ = True Then
                    '����
                    blnĿ¼ = True
                End If
            Else
                If rs������ϸ("�շ����") = "5" Or rs������ϸ("�շ����") = "6" Or rs������ϸ("�շ����") = "7" Then
                    'ҩƷ
                    blnĿ¼ = True
                Else
                    '����
                    blnĿ¼ = False
                End If
            End If
            
            If blnĿ¼ = False Then
                '������׼��Ŀ�У�ֻ�а��������м���
                rs�㷨.Filter = "ID=" & rs������ϸ("���մ���ID")
                If rs�㷨.RecordCount > 0 Then
                    '�㷨:1-�ܶ������Ŀ��2-סԺ�պ˶���Ŀ
                    If rs�㷨("�㷨") = 1 Then
                        If rs�㷨("ͳ��ȶ�") = 0 Then
                            curȫ�Է� = curȫ�Է� + rs������ϸ("���")
                        Else
                            cur����ͳ�� = cur����ͳ�� + rs������ϸ("���") * rs�㷨("ͳ��ȶ�") / 100
                        End If
                    Else
                        If Val(rs������ϸ("����")) > Val(rs�㷨("��׼����")) Then
                            '���סԺ�ճ�����׼��������ô�������� ��׼����*��׼���� +  (����-��׼����)*ͳ��ȶ�
                            '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                            dbl����� = rs�㷨("��׼����") * rs�㷨("��׼����") + _
                                (rs������ϸ("����") - IIf(rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0, 0, rs�㷨("��׼����"))) * rs�㷨("ͳ��ȶ�")
                        Else
                            '���סԺ�յ�����׼��������ô�������� ����*��׼���� ���� ����*ͳ��ȶ�
                            '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                            If rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0 Then
                                dbl����� = rs������ϸ("����") * rs�㷨("ͳ��ȶ�")
                            Else
                                dbl����� = rs������ϸ("����") * rs�㷨("��׼����")
                            End If
                        End If
                        
                        '�ܽ��������С����ȡȫ��������ֻ�����
                        cur����ͳ�� = cur����ͳ�� + IIf(rs������ϸ("���") < dbl�����, rs������ϸ("���"), dbl�����)
                        
                        If rs������ϸ("���") > dbl����� Then
                            'ȫ������ȫ�Է�
                            curȫ�Է� = curȫ�Է� + rs������ϸ("���") - dbl�����
                        End If
                    End If
                Else
                    curȫ�Է� = curȫ�Է� + rs������ϸ("���")
                End If
            Else
                cur����ͳ�� = cur����ͳ�� + rs������ϸ("���")
                g��������.ʵ������ = g��������.ʵ������ + rs������ϸ("���")
            End If
            
            dblTemp = dblTemp + rs������ϸ("���")
            rs������ϸ.MoveNext
        Loop
        
        g��������.�������ý�� = dblTemp
        g��������.����ͳ���� = cur����ͳ��
        g��������.ȫ�Էѽ�� = curȫ�Է�
        g��������.�����Ը���� = g��������.�������ý�� - curȫ�Է� - cur����ͳ��
    Else
        Do Until rs������ϸ.EOF
                
            dblTemp = dblTemp + rs������ϸ("���")
            rs������ϸ.MoveNext
        Loop
        g��������.�������ý�� = dblTemp
        g��������.����ͳ���� = g��������.�������ý��
        g��������.ȫ�Էѽ�� = 0
        g��������.�����Ը���� = 0
    End If
        
    '������������������������������������������������������������������������������������
    '3��������ߡ��ⶥ�ߡ�֧������������
    '3.1��������ߡ��ⶥ��
    With g��������
        If bln�޷ⶥ�� = True Then
            .�ⶥ�� = 0
        Else
            '��鲡�˵Ĳ����Ƿ�������ⶥ��
            gstrSQL = "SELECT B.����ⶥ��,b.�ⶥ�߽�� " & _
                     " FROM �����ʻ� A,���ղ��� B " & _
                     " WHERE A.����ID=" & g��������.����ID & " AND A.����=" & TYPE_��������ɽ & " AND A.����ID=B.ID(+)"
            Call OpenRecordset(rsTemp, "�������")
            
            If Nvl(rsTemp("����ⶥ��"), 0) = 1 Then
                If IsNull(rsTemp("�ⶥ�߽��")) = True Then
                    bln�޷ⶥ�� = True '�������Ҳ�����޷ⶥ�ߣ��繤�ˡ�����
                Else
                    .�ⶥ�� = rsTemp("�ⶥ�߽��")
                End If
            Else
                gstrSQL = "select max(decode(A.����,'A',A.���,0)) as ������ " & _
                          "  from ����֧���޶� A " & _
                          "  where A.����=" & TYPE_��������ɽ & " and A.����=" & lng���� & " and A.����='A' and A.���=" & .���
                Call OpenRecordset(rsTemp, "�������")
                        
                .�ⶥ�� = IIf(IsNull(rsTemp("������")), 0, rsTemp("������"))
                If .�ⶥ�� = 0 Then
                    MsgBox "���ڡ���Ƚ�����������ñ���ȵķⶥ�ߡ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End With
    
    '3.3��ȡ�÷��õ���
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "select B.����,B.����,B.����,A.���� " & _
              "  from ����֧������ A,���շ��õ� B " & _
              "  Where A.���� =" & TYPE_��������ɽ & " And A.���� =" & lng���� & " And A.��� =" & g��������.��� & " And A.��ְ =" & lng��ְ & " And A.����� =" & lng����� & _
              "       and A.����=B.���� and A.����=b.���� and A.����=B.���� and B.����=1"
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp.RecordCount = 0 Then
        MsgBox "���ڡ���Ƚ�����������ñ���ȵ�ͳ��֧����������", vbInformation, gstrSysName
        Exit Function
    End If
    
    '������������������������������������������������������������������������������������
    '4������ôν���ɱ����Ľ��
    With g��������
        If bln�޷ⶥ�� = True Then
            '���ÿ����б����˵Ľ��
            .ͳ�ﱨ����� = .ʵ������ + (.����ͳ���� - .ʵ������) * rsTemp("����") / 100
        Else
            '��������ܱ����Ľ������ض�Ŀ¼�ģ�������Ƚ�
            dblTemp = .ʵ������ + (.����ͳ���� - .ʵ������) * rsTemp("����") / 100
            If dblTemp > .�ⶥ�� - .�ۼ�ͳ�ﱨ�� Then
                .ͳ�ﱨ����� = .�ⶥ�� - .�ۼ�ͳ�ﱨ��
                .�����Ը���� = .����ͳ���� - .ͳ�ﱨ����� / rsTemp("����") * 100
                If .�����Ը���� < 0 Then .�����Ը���� = 0                   '�������ͳ����������ߣ�Ϊ����
            Else
                .ͳ�ﱨ����� = dblTemp
            End If
        End If
    End With
    
    סԺ�������_����ɽ = "ҽ������;" & g��������.ͳ�ﱨ����� & ";0"
End Function

Public Function סԺ����_����ɽ(lng����ID As Long) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    Dim rsTemp As New ADODB.Recordset
    Dim var������� As Variant
        
    With g��������
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & .����ID & "," & TYPE_��������ɽ & "," & .��� & "," & _
            .�ʻ��ۼ����� & "," & .�ʻ��ۼ�֧�� & "," & .�ۼƽ���ͳ�� + .����ͳ���� & "," & _
            .�ۼ�ͳ�ﱨ�� + .ͳ�ﱨ����� & "," & .סԺ���� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
        
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��������ɽ & "," & .����ID & "," & _
            .��� & "," & .�ʻ��ۼ����� & "," & .�ʻ��ۼ�֧�� & "," & .�ۼƽ���ͳ�� & "," & _
            .�ۼ�ͳ�ﱨ�� & "," & .סԺ���� + 1 & "," & .���� & "," & .�ⶥ�� & "," & .ʵ������ & "," & _
            .�������ý�� & "," & .ȫ�Էѽ�� & "," & .�����Ը���� & "," & .����ͳ���� & "," & .ͳ�ﱨ����� & ",0," & _
            .�����Ը���� & ",0,NULL," & .��ҳID & "," & .��;���� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    End With
    
    סԺ����_����ɽ = True
End Function

Public Function סԺ�������_����ɽ(lng����ID As Long) As Boolean
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
    Call OpenRecordset(rsTemp, "ģ��ҽ��")
    
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "Select * " & _
              "  From ���ս����¼ Where ����=2 and ��¼ID='" & lng����ID & "'"
    Call OpenRecordset(rsTemp, "ģ��ҽ��")
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "�ò��˵�ҽ���������ݶ�ʧ���������ϡ�"
        Exit Function
    End If
    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
    gstrSQL = "select B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ� " & _
              " from �����ʻ� A,�ʻ������Ϣ B " & _
              " where A.����ID=B.����ID(+) and A.����=B.����(+) and B.���(+)=" & Year(zlDatabase.Currentdate) & " and A.����ID=" & rsTemp("����ID") & " and A.����=" & TYPE_��������ɽ
    Call OpenRecordset(rs�ʻ�, "ģ��ҽ��")
    
    If rs�ʻ�.EOF = False Then
        lngסԺ���� = IIf(IsNull(rs�ʻ�("סԺ�����ۼ�")), 0, rs�ʻ�("סԺ�����ۼ�"))
        cur�ʻ����� = IIf(IsNull(rs�ʻ�("�ʻ������ۼ�")), 0, rs�ʻ�("�ʻ������ۼ�"))
        cur�ʻ�֧�� = IIf(IsNull(rs�ʻ�("�ʻ�֧���ۼ�")), 0, rs�ʻ�("�ʻ�֧���ۼ�"))
        cur�ۼƽ���ͳ�� = IIf(IsNull(rs�ʻ�("����ͳ���ۼ�")), 0, rs�ʻ�("����ͳ���ۼ�"))
        cur�ۼ�ͳ�ﱨ�� = IIf(IsNull(rs�ʻ�("ͳ�ﱨ���ۼ�")), 0, rs�ʻ�("ͳ�ﱨ���ۼ�"))
    End If
    
    '���˴������ݱ���������������ݱ������һ������
    '��˾Ͳ���Ҫ�������������
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & TYPE_��������ɽ & "," & rsTemp("���") & "," & _
        cur�ʻ����� & "," & cur�ʻ�֧�� - rsTemp("�����ʻ�֧��") & "," & cur�ۼƽ���ͳ�� - rsTemp("����ͳ����") & "," & _
        cur�ۼ�ͳ�ﱨ�� - rsTemp("ͳ�ﱨ�����") & "," & lngסԺ���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    '�������ݣ������˼����ۼ�
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��������ɽ & "," & rsTemp("����ID") & "," & _
        rsTemp("���") & "," & cur�ʻ����� & "," & cur�ʻ�֧�� - rsTemp("�����ʻ�֧��") & "," & cur�ۼƽ���ͳ�� - rsTemp("����ͳ����") & "," & _
        cur�ۼ�ͳ�ﱨ�� - rsTemp("ͳ�ﱨ�����") & "," & lngסԺ���� & "," & rsTemp("����") * -1 & "," & rsTemp("�ⶥ��") & "," & rsTemp("ʵ������") * -1 & "," & _
        rsTemp("�������ý��") * -1 & "," & rsTemp("ȫ�Ը����") * -1 & "," & rsTemp("�����Ը����") * -1 & "," & rsTemp("����ͳ����") * -1 & "," & _
        rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") * -1 & "," & rsTemp("�����ʻ�֧��") * -1 & ",''," & _
        IIf(IsNull(rsTemp("��ҳID")), "null", rsTemp("��ҳID")) & "," & rsTemp("��;����") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
    
    
    gstrSQL = "select ����,����ͳ����,ͳ�ﱨ�����,���� from ���ս������ where ����ID=" & lng����ID
    Call OpenRecordset(rs�������, "ģ��ҽ��")
    
    Do Until rs�������.EOF
        '����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
        gstrSQL = "zl_���ս������_Insert(" & lng����ID & "," & _
            rs�������("����") & "," & rs�������("����ͳ����") * -1 & "," & rs�������("ͳ�ﱨ�����") * -1 & "," & rs�������("����") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ģ��ҽ��")
        
        rs�������.MoveNext
    Loop
    
    סԺ�������_����ɽ = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ���ʴ���_����ɽ(strNO As String, int���� As Integer, int״̬ As Integer, Optional lng����ID As Long) As Boolean
'���ܣ���סԺ���˵ļ��ʵ����ϴ���ҽ��ǰ�÷�����
'������lng����ID=�Ƿ�ֻ�ϴ�������ָ�����˵ķ���
    
    '��ʱ�����κδ���
    '����û�Ҫ���ӡ���ֵĲ���һ���嵥Ҫ�������ⲡ����Ŀ¼�е�ͳ����������ڴ˴����޸�
    ���ʴ���_����ɽ = True
End Function




