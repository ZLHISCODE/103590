Attribute VB_Name = "mdlPivaNormal"
Option Explicit






Public Function Piva_GetMedi(ByVal intStemp As Integer, ByVal strIDS As String, ByVal int��ʾ���� As Integer) As Recordset
    'ȡҽ����Ӧ��ҩƷ���������������Ǿ����ҩƷ����
    'intStemp��0-���ҽ����1-��ͨ�����ҽ����2-δͨ�����ҽ��
    'strIDS����ҽ��ID
    'int��ʾ���У���˸�ҩ�����������ݣ�0-�����ã�1-����
    '             ����ʱ������Һ��ҩ��¼���޼�¼
    Dim strTmp As String
    Dim strSqlTmp As String
    
    On Error GoTo errHandle
        
    gstrSQL = "Select Distinct a.Id, a.���id, a.����id, a.��ҳid, a.����ҽ��, a.�����, a.ҩʦ���ԭ��, g.���˲���id, g.���˿���id, b.���� ��������, f.��ǰ���� ����, p.��ҩ����," & vbNewLine & _
        "                Decode(a.ҽ����Ч, 0, '����', 1, '��ʱ') ҽ����Ч, m.���� ��ҩ;��, g.��ʶ�� As סԺ��, a.����, a.�Ա�, a.����, c.Id ҩƷid, c.���� ҩƷ����," & vbNewLine & _
        "                c.���, a.��������, i.���㵥λ, i.Id ҩ��id, a.ִ��Ƶ��, Nvl(a.ҩʦ��˱�־, 0) ��˱�־, a.ִ��ʱ�䷽��, a.Ƥ�Խ��, a.����ʱ��," & vbNewLine & _
        "                Nvl(t.�Ƿ�Ƥ��, 0) �Ƿ�Ƥ��, a.ִ������, a.ִ�б��" & vbNewLine & _
        " From ����ҽ����¼ A, ����ҽ����¼ L, סԺ���ü�¼ G, ���ű� B, �շ���ĿĿ¼ C, ������Ϣ F, ������ĿĿ¼ I, ҩƷ��� J, ҩƷ���� T, ��ҺҩƷ���� P, ������ĿĿ¼ M," & vbNewLine & _
        "     Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) K" & vbNewLine & _
        " Where a.���id = l.Id And a.Id = g.ҽ����� And a.����id = f.����id And a.���˿���id = b.Id And g.�շ�ϸĿid = c.Id And j.ҩƷid = c.Id And" & vbNewLine & _
        "      j.ҩƷid = p.ҩƷid And j.ҩ��id = i.Id And l.������Ŀid = m.Id And j.ҩ��id = t.ҩ��id And l.Id = k.Column_Value"
    
    If int��ʾ���� = 1 Then
        '�������ҩ���������ݣ�����Ҫ�������״̬
    Else
        '���������ҩ���������ݣ������״̬������
        If intStemp = 1 Then
            '���ҽ����־����ͨ�����
            gstrSQL = gstrSQL & " and (a.ҩʦ��˱�־=1 or a.ҩʦ��˱�־=3) "
        ElseIf intStemp = 2 Then
            '���ҽ����־: δͨ�����
            gstrSQL = gstrSQL & " and (a.ҩʦ��˱�־=2 or a.ҩʦ��˱�־=3) "
        Else
            '���ҽ����־: δ���
            gstrSQL = gstrSQL & " and (Nvl(a.ҩʦ��˱�־,0)=0 or a.ҩʦ��˱�־=3) "
        End If
    End If

    '�ų��Ա�ҩ����ȡҩ�����浥����ȡ
    gstrSQL = gstrSQL & " And Not Exists " & _
        " (Select 1 From ����ҽ����¼ Aa Where Aa.���id = a.���id And Aa.ִ������ = 5 And (Aa.ִ�б�� = 0 Or Aa.ִ�б�� = 2)) "
 
    If int��ʾ���� = 0 Then
        'ֻ��ѯ��Һ����
        gstrSQL = gstrSQL & " And Exists (Select 1 From ҩƷ�շ���¼ Aa, ��Һ��ҩ���� Bb Where Aa.Id = Bb.�շ�id And Aa.����id = g.Id) "
    End If
    
    '�ϲ������������
    strTmp = Replace(gstrSQL, "סԺ���ü�¼", "������ü�¼")
    gstrSQL = gstrSQL & " Union All " & strTmp
    
    '�ϲ��Ա�ҩ����ȡҩ����������������
    strTmp = "Select Distinct a.Id, a.���id, a.����id, a.��ҳid, a.����ҽ��, a.�����, a.ҩʦ���ԭ��, f.��ǰ����id As ���˲���id, f.��ǰ����id As ���˿���id," & vbNewLine & _
        "                b.���� ��������, f.��ǰ���� ����, p.��ҩ����, Decode(a.ҽ����Ч, 0, '����', 1, '��ʱ') ҽ����Ч, m.���� ��ҩ;��, f.סԺ��, a.����, a.�Ա�, a.����," & vbNewLine & _
        "                c.Id ҩƷid, c.���� ҩƷ����, c.���, a.��������, i.���㵥λ, i.Id ҩ��id, a.ִ��Ƶ��, Nvl(a.ҩʦ��˱�־, 0) ��˱�־, a.ִ��ʱ�䷽��, a.Ƥ�Խ��," & vbNewLine & _
        "                a.����ʱ��, Nvl(t.�Ƿ�Ƥ��, 0) �Ƿ�Ƥ��, a.ִ������, a.ִ�б��" & vbNewLine & _
        " From ����ҽ����¼ A, ����ҽ����¼ L, ���ű� B, �շ���ĿĿ¼ C, ������Ϣ F, ������ĿĿ¼ I, ҩƷ��� J, ҩƷ���� T, ��ҺҩƷ���� P, ������ĿĿ¼ M," & vbNewLine & _
        "     Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) K" & vbNewLine & _
        " Where a.���id = l.Id And a.����id = f.����id And a.���˿���id = b.Id And a.�շ�ϸĿid = c.Id And j.ҩƷid = c.Id And j.ҩƷid = p.ҩƷid And" & vbNewLine & _
        "      j.ҩ��id = i.Id And l.������Ŀid = m.Id And j.ҩ��id = t.ҩ��id And l.Id = k.Column_Value And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From ����ҽ����¼ Aa" & vbNewLine & _
        "       Where Aa.���id = a.���id And Aa.ִ������ = 5 And (Aa.ִ�б�� = 0 Or Aa.ִ�б�� = 2))"

    '�ϲ�����
    gstrSQL = gstrSQL & " Union All " & strTmp
    gstrSQL = gstrSQL & " Order By ��������,����id,���Id"
    
    Set Piva_GetMedi = zlDatabase.OpenSQLRecord(gstrSQL, "Piva_GetMedi", strIDS)

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function Piva_GetTrans(ByVal strIDS As String, ByVal lng����id As Long, ByVal strStep As String, ByVal intPack As Integer, ByVal blnShowOhters As Boolean) As ADODB.Recordset
        
    'ȡ��Һ��ҩ��¼
    'lngCenterID����Һ��������ID
    'str����ID������ID��
    'dateExeStart��dateExeEnd����Һ��ҩ���ݵ�ִ��ʱ�䷶Χ
    'strStep(��������)��01-��ҩӡǩ(1)��02-��ҩ�˲�(2)��03-���ͺ˲�(4)��04-�������(9)��10-�����ͨ��ҽ��(10)��11-���δͨ��ҽ��(10)��12-�ѷ��Ͳ鿴(5), 13-��ǩ�ղ鿴(6)��14-�ܾ�ǩ�ղ�(7)��15-�����ϲ鿴
    '�������ͣ�1�����ƣ�2����ҩ��3��У�ԣ�4����ҩ��5�����ͣ�6��ǩ�գ�7���ܾ�ǩ��  8��ȷ�Ͼ��գ�9���������룬10���������
    'intPack���룺0-���У�1-����ҩ��2-�����
    '�Ƿ�����0-�����(��Һ),1-�������,2-�������Ĵ��
    'intShowType:��ʾ��ʽ��0-��ͨ��ʾ��2-�����Ա�ҩ����ʾ
    
    Dim strOhterSQL As String       '������ȡҽ����[�Ա�ҩ]���������
    Dim strTmp As String
    
    On Error GoTo errHandle
    
    If strStep = "15" Then
        '�Ѱ�ҩ״̬
        '1.�������ͨ��
        gstrSQL = "Select A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,K.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��, A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' NO, F.���� As ҩƷ����,' ' ����ԭ��, " & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, e.����, e.����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, e.Ƶ��, '�������ͨ��' As ��������, " & _
            " 0 As ��ҩ����, (e.���ϵ��*e.ʵ������ / G.סԺ��װ) As ����,e.���ϵ��*e.ʵ������ As ʵ������, G.סԺ��λ As ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, e.�÷�, e.ҩƷid,0 as �������,0 As ����id,null As ����, A.��ҩ����,L.����ʱ�� As ҽ������ʱ��,nvl(T.������,'0') ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id, M.id As ��Ӧҽ��ID,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,x.��ҩ���� As ��ҩ����1, m.ִ������, m.ִ�б��  " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M, סԺ���ü�¼ D, ҩƷ�շ���¼ E, ����ҽ������ L ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� k,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V,��Һ��ҩ���� Z "

        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID And A.���˿���id = C.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And g.ҩƷid = e.ҩƷid And T.ҩ��id=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=K.����(+) And " & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID " & _
            " And m.Id = d.ҽ����� And d.Id = e.����id And a.ҽ��id = l.ҽ��id(+) And a.���ͺ� = l.���ͺ�(+) And a.id = z.��¼id And z.�շ�id = e.Id " & _
            " And a.����״̬=10 And A.id=V.Column_Value  And Exists (Select 1 From ��Һ��ҩ���� D, ҩƷ�շ���¼ E Where d.�շ�id = e.Id And d.��¼id = a.Id)"

        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
        
        '2.������˾ܾ�
        gstrSQL = gstrSQL & " Union All " & _
            "Select A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,K.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��, A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' NO, F.���� As ҩƷ����,' ' ����ԭ��, " & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, e.����, e.����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, e.Ƶ��, '������˾ܾ�' As ��������, " & _
            " 0 As ��ҩ����, (e.���ϵ��*e.ʵ������ / G.סԺ��װ) As ����,e.���ϵ��*e.ʵ������ As ʵ������, G.סԺ��λ As ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, e.�÷�, e.ҩƷid,0 as �������,0 As ����id,null As ����, A.��ҩ����,L.����ʱ�� As ҽ������ʱ��,nvl(T.������,'0') ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id, M.id As ��Ӧҽ��ID,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,x.��ҩ���� As ��ҩ����1, m.ִ������, m.ִ�б��  " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M, סԺ���ü�¼ D, ҩƷ�շ���¼ E, ����ҽ������ L ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� k,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V,��Һ��ҩ���� Z "

        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID And A.���˿���id = C.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And g.ҩƷid = e.ҩƷid And T.ҩ��id=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=K.����(+) And " & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID " & _
            " And m.Id = d.ҽ����� And d.Id = e.����id And a.ҽ��id = l.ҽ��id(+) And a.���ͺ� = l.���ͺ�(+)  and e.ʵ������>0 And a.id = z.��¼id And z.�շ�id = e.Id " & _
            " And a.����״̬=11 And A.id=V.Column_Value And Exists (Select 1 From ��Һ��ҩ���� D, ҩƷ�շ���¼ E Where d.�շ�id = e.Id And d.��¼id = a.Id)"
            
        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
                
        '�ϲ��������
        strTmp = Replace(gstrSQL, "סԺ���ü�¼", "������ü�¼")
        gstrSQL = gstrSQL & " Union All " & strTmp
        
        'δ��ҩ״̬
        '�����
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,K.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��,  A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' As NO, F.���� As ҩƷ����,' ' ����ԭ��, " & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, '' As ����, '' As ����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, '' As Ƶ��, 'δ��ҩ����' As ��������, " & _
            " 0 As ��ҩ����, (M.��������/ G.����ϵ�� / G.סԺ��װ) As ����,M.��������/ G.����ϵ�� As ʵ������, G.סԺ��λ As ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, '' As �÷�, M.�շ�ϸĿid As ҩƷid,0 as �������,0 As ����id,null As ����, " & _
            " A.��ҩ����,Null As ҽ������ʱ��,nvl(T.������,'0') ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,M.id As ��Ӧҽ��ID,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,x.��ҩ���� As ��ҩ����1, m.ִ������, m.ִ�б��  " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� k,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V "
        
        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID  And A.���˿���id = C.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And M.�շ�ϸĿid = F.ID And T.ҩ��id=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=K.����(+) And " & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID And a.����״̬=10  " & _
            " And  A.id=V.Column_Value  And Not Exists (Select 1 From ��Һ��ҩ���� D, ҩƷ�շ���¼ E Where d.�շ�id = e.Id And d.��¼id = a.Id) "
        
        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
                
        '��Ʒ��
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,K.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��,  A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' As NO, J.���� As ҩƷ����,' ' ����ԭ��, " & _
            " J.���� As ͨ����, '' As ��Ʒ��, I.���� As Ӣ����, '' as ���, '' As ����, '' As ����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, '' As Ƶ��, 'δ��ҩ����' As ��������, " & _
            " 0 As ��ҩ����, 0 As ����,0 As ʵ������, '' ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, '' As �÷�, Decode(Nvl(m.�շ�ϸĿid, 0), 0, j.Id, m.�շ�ϸĿid) As ҩƷid,0 as �������,0 As ����id,null As ����, " & _
            " A.��ҩ����,Null As ҽ������ʱ��,0 ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,M.id As ��Ӧҽ��ID,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,'' As ��ҩ����1, m.ִ������, m.ִ�б��  " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C,������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� k,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V "

        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID  And A.���˿���id = C.ID and M.�շ�ϸĿid is null and M.������Ŀid=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=K.����(+) And a.����״̬=10  " & _
            " And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And J.id = I.������Ŀid(+) And I.����(+) = 2 And j.Id = t.ҩ��id " & _
            " And  A.id=V.Column_Value And Not Exists (Select 1 From ��Һ��ҩ���� D, ҩƷ�շ���¼ E Where d.�շ�id = e.Id And d.��¼id = a.Id) "

        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
        
        '��ȡҽ����[�Ա�ҩ]���������
        strOhterSQL = "Union All" & vbNewLine & _
                    "Select Distinct a.Id As ��ҩid, a.���α��, a.���ȼ�, a.�Ƿ�ȷ�ϵ���, a.����id, a.���, a.��ҩ����, s.��ɫ, a.����, a.�Ա�, a.����, a.סԺ��, a.����," & vbNewLine & _
                    "                LPad(a.����, 10, ' ') ��������, k.����, m.��� ҽ�����, m.ҩʦ���ʱ��, m.ִ��Ƶ��, a.���˲���id, a.���˿���id, a.ִ��ʱ��, a.ƿǩ��, a.���ʱ��," & vbNewLine & _
                    "                m.����id, m.��ҳid, a.�Ƿ��������, a.�Ƿ�����, a.�ֹ���������, '' ����ԭ��, a.������Ա, a.����ʱ��, Nvl(a.��ӡ��־, 0) As ��ӡ��־, a.�Ƿ���," & vbNewLine & _
                    "                b.���� As ���˲���, c.���� As ���˿���, 0 As �շ�id, 0 As ����, '' As NO, f.���� As ҩƷ����, ' ' ����ԭ��, f.���� As ͨ����," & vbNewLine & _
                    "                h.���� As ��Ʒ��, i.���� As Ӣ����, f.���, f.����, '' As ����, m.�������� As ����, j.���㵥λ As ������λ, j.Id ҩ��id, m.ִ��Ƶ�� As Ƶ��," & vbNewLine & _
                    "                '' As ��������, 0 As ��ҩ����, (m.�������� / g.����ϵ�� / g.סԺ��װ) As ����, (m.�������� / g.����ϵ��) As ʵ������, g.סԺ��λ As ��λ," & vbNewLine & _
                    "                0 As ����, 0 As �������, Nvl(m.�����, -1) �����, Zc.ҽ������ As �÷�, m.�շ�ϸĿid As ҩƷid, 0 As �������, 0 As ����id, o.����," & vbNewLine & _
                    "                a.��ҩ����, r.����ʱ�� As ҽ������ʱ��, Nvl(t.������, '0') ��ҩ����, t.��ý, m.Ƥ�Խ��, m.����ʱ��, a.ҽ��id, m.Id As ��Ӧҽ��id, a.���ͺ�," & vbNewLine & _
                    "                Nvl(t.�Ƿ�Ƥ��, 0) �Ƿ�Ƥ��, x.��ҩ���� As ��ҩ����1, m.ִ������, m.ִ�б�� " & vbNewLine & _
                    "From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ F, ҩƷ��� G, ��ҺҩƷ���� X, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M, ������ҳ O, ��ҩ�������� S," & vbNewLine & _
                    "     ҩƷ���� T, ��λ״����¼ Q, ��λ���Ʒ��� K, ����ҽ����¼ Zc, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V," & vbNewLine & _
                    "     ����ҽ������ R"

        strOhterSQL = strOhterSQL & " Where a.ҽ��id = m.���id And (m.ִ������ = 5 And m.ִ�б�� = 0) And a.����id = s.��������id And" & vbNewLine & _
                    "      a.��ҩ���� = s.���� And a.���� = q.����(+) And a.���˲���id = q.����id(+) And q.��λ���� = k.����(+) And a.���˲���id = b.Id And" & vbNewLine & _
                    "      a.���˿���id = c.Id And m.�շ�ϸĿid = f.Id And f.Id = g.ҩƷid And g.ҩƷid = h.�շ�ϸĿid(+) And h.����(+) = 3 And" & vbNewLine & _
                    "      g.ҩƷid = x.ҩƷid(+) And g.ҩ��id = i.������Ŀid(+) And i.����(+) = 2 And g.ҩ��id = j.Id And j.Id = t.ҩ��id And" & vbNewLine & _
                    "      a.ҽ��id = Zc.Id And m.����id = o.����id(+) And m.��ҳid = o.��ҳid(+) And a.ҽ��id = r.ҽ��id And a.���ͺ� = r.���ͺ� And " & vbNewLine & _
                    "      a.Id = v.Column_Value "

        If intPack = 1 Then
            '�����
            strOhterSQL = strOhterSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            strOhterSQL = strOhterSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
    ElseIf strStep = "16" Then
        'ҽ������(�����)
        gstrSQL = " Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,K.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��,  A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' As NO, F.���� As ҩƷ����,' ' ����ԭ��, " & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, '' As ����, '' As ����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, '' As Ƶ��, 'ҽ������' As ��������, " & _
            " 0 As ��ҩ����, (M.��������/ G.����ϵ�� / G.סԺ��װ) As ����,M.��������/ G.����ϵ�� As ʵ������, G.סԺ��λ As ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, '' As �÷�, M.�շ�ϸĿid As ҩƷid,0 as �������,0 As ����id,null As ����, " & _
            " A.��ҩ����,Null As ҽ������ʱ��,nvl(T.������,'0') ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,M.id As ��Ӧҽ��ID,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,x.��ҩ���� As ��ҩ����1, m.ִ������, m.ִ�б�� " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� K,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V "
        
        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID  And A.���˿���id = C.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And M.�շ�ϸĿid = F.ID And T.ҩ��id=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=K.����(+) And " & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID And a.����״̬=12  " & _
            " And A.id=V.Column_Value "
        
        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
                
        '�ϲ�ҽ������(��Ʒ�ַ���)
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,K.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��,  A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' As NO, J.���� As ҩƷ����,' ' ����ԭ��, " & _
            " J.���� As ͨ����, '' As ��Ʒ��, I.���� As Ӣ����, '' as ���, '' As ����, '' As ����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, '' As Ƶ��, 'ҽ������' As ��������, " & _
            " 0 As ��ҩ����, 0 As ����,0 As ʵ������, '' ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, '' As �÷�, Decode(Nvl(m.�շ�ϸĿid, 0), 0, j.Id, m.�շ�ϸĿid) As ҩƷid,0 as �������,0 As ����id,null As ����, " & _
            " A.��ҩ����,Null As ҽ������ʱ��,0 ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,M.id As ��Ӧҽ��ID,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,'' As ��ҩ����1, m.ִ������, m.ִ�б�� " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C,������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� K,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V "

        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID  And A.���˿���id = C.ID and M.�շ�ϸĿid is null and M.������Ŀid=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=K.����(+) And a.����״̬=12  " & _
            " And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And J.id = I.������Ŀid(+) And I.����(+) = 2 And j.Id = t.ҩ��id " & _
            " And A.id=V.Column_Value "

        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
    Else
        '����
        gstrSQL = "Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����, S.��ɫ,A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,K.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��,  A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������," & IIf(strStep = "13", "W.����˵�� ����ԭ��,", "'' ����ԭ��,") & _
            "  A.������Ա,A.����ʱ��,Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, D.�շ�id, E.����, E.NO, F.���� As ҩƷ����, " & IIf(strStep = "04", "Y.����ԭ��, ", "' ' ����ԭ��,") & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, E.����, E.����, E.����, J.���㵥λ As ������λ,J.id ҩ��id, E.Ƶ��, '' As ��������, " & _
            " Case Nvl(E.�����, 'δ���') When 'δ���' Then E.ʵ������ * Nvl(E.����, 1) / G.סԺ��װ Else 0 End As ��ҩ����,M.����id,M.��ҳid,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,M.id As ��Ӧҽ��ID,A.���ͺ�, " & _
            " (D.���� / G.סԺ��װ)  As ����,D.���� As ʵ������, G.סԺ��λ As ��λ,Nvl(E.����,0) As ����, Nvl(L.ʵ������, 0)/ G.סԺ��װ As �������, Nvl(M.�����,-1) �����, E.�÷�, E.ҩƷid, n.��� As �������,E.����id, o.����, A.��ҩ����,r.����ʱ�� As ҽ������ʱ��,nvl(T.������,'0') ��ҩ����,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,x.��ҩ���� As ��ҩ����1, m.ִ������, m.ִ�б��  " & _
            " From  ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, ��Һ��ҩ���� D, ҩƷ�շ���¼ E, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X,  �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M, סԺ���ü�¼ N, ������ҳ O ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ Q,��λ���Ʒ��� K,Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V "
        
        '��ȡҽ����[�Ա�ҩ]���������
        strOhterSQL = " Union All "
        strOhterSQL = strOhterSQL & "Select Distinct a.Id As ��ҩid, a.���α��, a.���ȼ�, a.�Ƿ�ȷ�ϵ���, a.����id, a.���, a.��ҩ����, s.��ɫ, a.����, a.�Ա�, a.����, a.סԺ��, a.����," & vbNewLine & _
                    "                LPad(a.����, 10, ' ') ��������, k.����, m.��� ҽ�����, m.ҩʦ���ʱ��, m.ִ��Ƶ��, a.���˲���id, a.���˿���id, a.ִ��ʱ��, a.ƿǩ��, a.���ʱ��," & vbNewLine & _
                    "                a.�Ƿ��������, a.�Ƿ�����, a.�ֹ���������, '' ����ԭ��, a.������Ա, a.����ʱ��, Nvl(a.��ӡ��־, 0) As ��ӡ��־, a.�Ƿ���, b.���� As ���˲���," & vbNewLine & _
                    "                c.���� As ���˿���, 0 As �շ�id, 0 As ����, '' As NO, f.���� As ҩƷ����, ' ' ����ԭ��, f.���� As ͨ����, h.���� As ��Ʒ��," & vbNewLine & _
                    "                i.���� As Ӣ����, f.���, f.����, '' As ����, m.�������� As ����, j.���㵥λ As ������λ, j.Id ҩ��id, m.ִ��Ƶ�� As Ƶ��, '' As ��������," & vbNewLine & _
                    "                0 As ��ҩ����, m.����id, m.��ҳid, t.��ý, m.Ƥ�Խ��, m.����ʱ��, a.ҽ��id, m.Id As ��Ӧҽ��id, a.���ͺ�," & vbNewLine & _
                    "                (m.�������� / g.����ϵ�� / g.סԺ��װ) As ����, (m.�������� / g.����ϵ��) As ʵ������, g.סԺ��λ As ��λ, 0 As ����, 0 As �������," & vbNewLine & _
                    "                Nvl(m.�����, -1) �����, Zc.ҽ������ As �÷�, m.�շ�ϸĿid As ҩƷid, 0 As �������, 0 As ����id, o.����, a.��ҩ����," & vbNewLine & _
                    "                r.����ʱ�� As ҽ������ʱ��, Nvl(t.������, '0') ��ҩ����, Nvl(t.�Ƿ�Ƥ��, 0) �Ƿ�Ƥ��, x.��ҩ���� As ��ҩ����1, m.ִ������, m.ִ�б�� " & vbNewLine & _
                    " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ F, ҩƷ��� G, ��ҺҩƷ���� X, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M, ������ҳ O, ��ҩ�������� S," & vbNewLine & _
                    "     ҩƷ���� T, ��λ״����¼ Q, ��λ���Ʒ��� K, ����ҽ����¼ Zc, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) V," & vbNewLine & _
                    "     ����ҽ������ R"
        
        If strStep = "13" Then gstrSQL = gstrSQL & ",��Һ��ҩ״̬ W "
        
        If strStep = "04" Then gstrSQL = gstrSQL & ",���˷������� Y "
        
'        If strStep = "01" And bln������ Then
'            gstrSQL = gstrSQL & ",��������¼ Q,���������ϸ K "
'        End If
        
        gstrSQL = gstrSQL & ",(Select �ⷿid, ҩƷid, Nvl(����, 0) As ����, Nvl(ʵ������, 0) As ʵ������ " & _
            " From ҩƷ��� Where ���� = 1 And �ⷿid = [2]) L, ҩƷ�շ���¼ P, ����ҽ������ R " & IIf(strStep = "04", ", ��Һ��ҩ״̬ U ", "")
        
        gstrSQL = gstrSQL & " Where A.���˲���id = B.ID And A.���˿���id = C.ID And A.ID = D.��¼id And D.�շ�id = E.ID And E.ҩƷid = F.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And E.����id = N.ID And N.ҽ����� = M.ID And " & IIf(strStep = "13", "W.��ҩid=A.id And A.����״̬=W.�������� And A.����ʱ��=W.����ʱ�� And ", "") & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID And T.ҩ��id=J.ID And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) " & IIf(strStep = "04", " And Y.����id=N.id And y.����ʱ�� = u.����ʱ�� And u.��ҩid = v.Column_Value ", "") & _
            " And E.�ⷿid = L.�ⷿid(+) And E.ҩƷid = L.ҩƷid(+) And A.����=Q.����(+) And  A.���˲���id=Q.����id(+) And A.���˿���id=Q.����id(+) and Q.��λ����=K.����(+) And Nvl(E.����, 0) = L.����(+) " & _
            " And n.����id = o.����id(+) And n.��ҳid = o.��ҳid(+) And a.ҽ��id = r.ҽ��id And a.���ͺ� = r.���ͺ�  " & _
            " And e.���� = p.���� And e.No = p.No And e.�ⷿid + 0 = p.�ⷿid+0 And e.ҩƷid + 0 = p.ҩƷid+0 And e.��� = p.��� And (p.��¼״̬ = 1 Or Mod(p.��¼״̬, 3) = 0) And A.id=V.Column_Value  "
            
        strOhterSQL = strOhterSQL & " Where a.ҽ��id = m.���id And (m.ִ������ = 5 And m.ִ�б�� = 0) And a.����id = s.��������id And" & vbNewLine & _
                    "      a.��ҩ���� = s.���� And a.���� = q.����(+) And a.���˲���id = q.����id(+) And q.��λ���� = k.����(+) And a.���˲���id = b.Id And" & vbNewLine & _
                    "      a.���˿���id = c.Id And m.�շ�ϸĿid = f.Id And f.Id = g.ҩƷid And g.ҩƷid = h.�շ�ϸĿid(+) And h.����(+) = 3 And" & vbNewLine & _
                    "      g.ҩƷid = x.ҩƷid(+) And g.ҩ��id = i.������Ŀid(+) And i.����(+) = 2 And g.ҩ��id = j.Id And j.Id = t.ҩ��id And" & vbNewLine & _
                    "      a.ҽ��id = Zc.Id And m.����id = o.����id(+) And m.��ҳid = o.��ҳid(+) And a.ҽ��id = r.ҽ��id And a.���ͺ� = r.���ͺ� And " & vbNewLine & _
                    "      a.Id = v.Column_Value"

        If strStep = "01" Then
            '����ҩ
            gstrSQL = gstrSQL & " And A.����״̬=1 "
            strOhterSQL = strOhterSQL & " And A.����״̬=1 "
        ElseIf strStep = "02" Then
            '����ҩ
            gstrSQL = gstrSQL & " And A.����״̬=2 "
            strOhterSQL = strOhterSQL & " And A.����״̬=2 "
        ElseIf strStep = "03" Then
            '������
            gstrSQL = gstrSQL & " And A.����״̬=4 "
            strOhterSQL = strOhterSQL & " And A.����״̬=4 "
        ElseIf strStep = "11" Then
            '���������
            gstrSQL = gstrSQL & " And A.����״̬=10 "
            strOhterSQL = strOhterSQL & " And A.����״̬=10 "
        ElseIf strStep = "12" Then
            '�ѷ���
            gstrSQL = gstrSQL & " And A.����״̬=5 "
            strOhterSQL = strOhterSQL & " And A.����״̬=5 "
        ElseIf strStep = "13" Then
            '��ǩ��
            gstrSQL = gstrSQL & " And A.����״̬=6 "
            strOhterSQL = strOhterSQL & " And A.����״̬=6 "
        ElseIf strStep = "14" Then
            '�Ѿܾ�ǩ��
            gstrSQL = gstrSQL & " And A.����״̬=7 "
            strOhterSQL = strOhterSQL & " And A.����״̬=7 "
        ElseIf strStep = "04" Then
            gstrSQL = gstrSQL & " And A.����״̬=9 "
            strOhterSQL = strOhterSQL & " And A.����״̬=9 "
        End If
        
        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
            strOhterSQL = strOhterSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
            strOhterSQL = strOhterSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
        
        '�ϲ��������
        strTmp = Replace(gstrSQL, "סԺ���ü�¼", "������ü�¼")
        gstrSQL = gstrSQL & " Union All " & strTmp
    End If
    
    If blnShowOhters Then
        '�ϲ�SQL
        gstrSQL = gstrSQL & strOhterSQL
        '����
        gstrSQL = gstrSQL & " Order By ��ҩid "
    End If
    
    Set Piva_GetTrans = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Һ��ҩ��¼", strIDS, lng����id)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function PIVA_GetExcStatus(ByVal str��ҩids As String, ByVal intStatus As Integer) As ADODB.Recordset
    '��鲻���ϵ�ǰ״̬����Һ��
    'str��ҩids����Һ��ID��
    'intStatus����ǰӦ�õ�ҵ��״̬
    Dim i As Integer
    Dim arrExecute As Variant
    
    On Error GoTo errHandle
    arrExecute = GetArrayByStr(str��ҩids, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = " Select ID, ƿǩ��, ����״̬,�Ƿ��� " & _
            " From ��Һ��ҩ��¼ Where (����״̬ <> [2] " & IIf(intStatus = 2, " or �Ƿ���<>0", "") & ") And ID In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) "
        Set PIVA_GetExcStatus = zlDatabase.OpenSQLRecord(gstrSQL, "PIVA_GetStatus", CStr(arrExecute(i)), intStatus)
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function PIVA_GetTransCount(ByVal lngCenterID As Long, ByVal dateExeStart As Date, ByVal dateExeEnd As Date, _
    ByVal bln��� As Boolean, ByVal bln������ As Boolean, Optional ByVal intType As Integer, _
    Optional ByVal strMsg As String, Optional ByVal lngҩƷid As Long, Optional ByVal strƿǩ�� As String, _
    Optional ByVal lng����ID As Long, Optional intCheck As Integer, Optional ByVal strSourceDep As String) As ADODB.Recordset
    'ȡ������Һ����Ŀ
    'lngCenterID����Һ��������ID
    'dateExeStart��dateExeEnd����Һ��ҩ���ݵ�ִ��ʱ�䷶Χ
    On Error GoTo errHandle
    
    gstrSQL = "select ����, ����id, ����,  id,ҩʦ��˱�־,����,���� from " & _
        " (with W as (Select Distinct a.����״̬,c.ҩʦ��˱�־,a.���˲���id As ����id, '[' || b.���� || ']' || b.���� As ����,b.����,b.����, c.���id As ҽ��id,A.id,A.ƿǩ�� " & vbNewLine & _
        "       From ��Һ��ҩ��¼ A, ���ű� B, ����ҽ����¼ C" & IIf(bln������, ",��������¼ Q,���������ϸ K ", "") & vbNewLine & _
        "       Where a.���˲���id = b.Id And a.ҽ��id = c.���id And c.ִ������ <> 5 And a.����id = [1] And" & IIf(bln������, " c.id=k.ҽ��id and Q.id=K.��id and K.����ύ=1 and Q.�����=1 and", "") & vbNewLine & _
        "             a.ִ��ʱ�� Between [2] And [3] And "

    If strMsg <> "" Then
        gstrSQL = gstrSQL & IIf(intType = 1, "a.����=[4] and ", IIf(intType = 2, "a.����=[4] and ", "a.סԺ��=[4] and "))
    End If
    
    If lngҩƷid <> 0 Then
        gstrSQL = gstrSQL & "C.�շ�ϸĿid=[6] And "
    End If
    
    If lng����ID <> 0 Then
        gstrSQL = gstrSQL & "C.���˿���id=[7] And "
    End If
    
    gstrSQL = gstrSQL & " Exists" & vbNewLine & _
        "        (Select 1 From ��Һ��ҩ���� D Where d.��¼id = a.Id))," & vbNewLine & _
        "       R as (Select Distinct a.����״̬, a.���˲���id As ����id, '[' || b.���� || ']' || b.���� As ����,b.����,b.����," & vbNewLine & _
        "                                 a.Id" & vbNewLine & _
        "                 From ��Һ��ҩ��¼ A, ���ű� B,��Һ��ҩ���� C,ҩƷ�շ���¼ D" & vbNewLine & _
        "                 Where a.���˲���id = b.Id  And a.����id = [1] and A.id=C.��¼id and C.�շ�id=D.id And" & vbNewLine & _
        "                       a.ִ��ʱ�� Between [2]  and" & vbNewLine & _
        "                       [3] And "
        
    If strMsg <> "" Then
        gstrSQL = gstrSQL & IIf(intType = 1, "a.����=[4] and ", IIf(intType = 2, "a.����=[4] and ", "a.סԺ��=[4] and "))
    End If
    
    If lngҩƷid <> 0 Then
        gstrSQL = gstrSQL & "D.ҩƷid=[6] And "
    End If
    
    If lng����ID <> 0 Then
        gstrSQL = gstrSQL & "A.���˿���id=[7] And "
    End If
    
    If strƿǩ�� <> "" Then
        gstrSQL = gstrSQL & "A.ƿǩ��=[5] And "
    End If
    
    gstrSQL = gstrSQL & "Exists  (Select 1 From ��Һ��ҩ���� D Where d.��¼id = a.Id))"
    
    If bln��� = True And intCheck <> 0 Then
        gstrSQL = gstrSQL & ", T as (Select distinct D.���˲���id As ����id, '[' || B.���� || ']' || B.���� As ����, c.���id As id,c.ҩʦ��˱�־,B.����,b.����, a.�ⷿid, d.����," & _
            " d.����, d.��ʶ��, d.�շ�ϸĿid, d.���˿���id " & _
            " From ҩƷ�շ���¼ A, ���ű� B,����ҽ����¼ C,סԺ���ü�¼ D " & IIf(bln������, ",��������¼ Q,���������ϸ K ", "") & vbNewLine & _
            " Where D.���˲���id = B.ID And D.ҽ�����=C.id And A.����id=D.id And A.����=9 And C.ִ������<>5  And �ⷿid = [1] " & _
            " And A.�������� Between [2] And [3] " & IIf(bln������, " And c.id=k.ҽ��id and Q.id=K.��id and K.����ύ=1 and Q.�����=1", "") & _
            " Union all " & _
            " Select distinct D.���˲���id As ����id, '[' || B.���� || ']' || B.���� As ����, c.���id As id,c.ҩʦ��˱�־,B.����,b.����, a.�ⷿid, d.����," & _
            " '' as ����, d.��ʶ��, d.�շ�ϸĿid, d.���˿���id " & _
            " From ҩƷ�շ���¼ A, ���ű� B,����ҽ����¼ C,������ü�¼ D " & IIf(bln������, ",��������¼ Q,���������ϸ K ", "") & vbNewLine & _
            " Where D.���˲���id = B.ID And D.ҽ�����=C.id And A.����id=D.id And A.����=9 And C.ִ������<>5  And �ⷿid = [1] " & _
            " And A.�������� Between [2] And [3] " & IIf(bln������, " And c.id=k.ҽ��id and Q.id=K.��id and K.����ύ=1 and Q.�����=1", "") & ")"
    End If
    
    If bln��� = True Then
        '���ҽ��
        If intCheck = 0 Then
            gstrSQL = gstrSQL & " select Distinct '00' ����,����id,����,ҽ��id id,ҩʦ��˱�־,����,���� from  W where (Nvl(ҩʦ��˱�־, 0) = 0 or Nvl(ҩʦ��˱�־, 0)=3) and ����״̬=1" & vbNewLine & _
            "union all"
        Else
            gstrSQL = gstrSQL & " Select Distinct '00' ����, ����id, ����, ID, ҩʦ��˱�־, ����, ���� " & _
                " From t Where 1=1 "
            If strMsg <> "" Then
                gstrSQL = gstrSQL & IIf(intType = 1, " And t.����=[4] ", IIf(intType = 2, " And t.����=[4] ", " And t.��ʶ��=[4] "))
            End If

            If lngҩƷid <> 0 Then
                gstrSQL = gstrSQL & " And t.�շ�ϸĿid=[6]"
            End If

            If lng����ID <> 0 Then
                gstrSQL = gstrSQL & " And t.���˿���id=[7]"
            End If
            
            gstrSQL = gstrSQL & " union all"
        End If
            
         '��ҩ
        gstrSQL = gstrSQL & " select distinct '01' ����,����id,����,id,1 ҩʦ��˱�־,����,���� from  W where Nvl(ҩʦ��˱�־, 0) =1 and ����״̬=1 "
        
        If strƿǩ�� <> "" Then
            gstrSQL = gstrSQL & " and ƿǩ��=[5]"
        End If
    Else
        '��ҩ
        gstrSQL = gstrSQL & " select distinct '01' ����,����id,����,id,1 ҩʦ��˱�־,����,���� from  W where ����״̬=1"
        
        If strƿǩ�� <> "" Then
            gstrSQL = gstrSQL & " And ƿǩ��=[5]"
        End If
    End If
    '��ҩ
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '02' ����,����id,����,id,1 ҩʦ��˱�־,����,���� from  R where ����״̬=2"
    '����
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '03' ����,����id,���� ,id,1 ҩʦ��˱�־,����,���� from  R where ����״̬=4"

    '�������
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '04' ����,����id,����,id,1 ҩʦ��˱�־,����,���� from  R where ����״̬=9"
        
    If bln��� = True Then
        If intCheck = 0 Then
            '�����ͨ��ҽ���鿴
            gstrSQL = gstrSQL & " Union All " & _
                "select Distinct  '10' ����,����id,����,ҽ��id,1 ҩʦ��˱�־,����,���� from  W where  Nvl(ҩʦ��˱�־, 0) =1"
                
            'δ���ͨ��ҽ���鿴
            gstrSQL = gstrSQL & " Union All " & _
                "select Distinct  '11' ����,����id,����,ҽ��id,1 ҩʦ��˱�־,����,���� from  W where Nvl(ҩʦ��˱�־, 0) =2"
        Else
            gstrSQL = gstrSQL & " Union All " & _
                " Select Distinct '10' ����, ����id, ����, ID, ҩʦ��˱�־, ����, ���� " & _
                " From t Where t.ҩʦ��˱�־=1 "
            If strMsg <> "" Then
                gstrSQL = gstrSQL & IIf(intType = 1, " And t.����=[4] ", IIf(intType = 2, " And t.����=[4] ", " And t.��ʶ��=[4] "))
            End If

            If lngҩƷid <> 0 Then
                gstrSQL = gstrSQL & " And t.�շ�ϸĿid=[6]"
            End If

            If lng����ID <> 0 Then
                gstrSQL = gstrSQL & " And t.���˿���id=[7]"
            End If

            gstrSQL = gstrSQL & " Union All " & _
                " Select Distinct '11' ����, ����id, ����, ID, ҩʦ��˱�־, ����, ���� " & _
                " From t Where t.ҩʦ��˱�־=2 "
            If strMsg <> "" Then
                gstrSQL = gstrSQL & IIf(intType = 1, " And t.����=[4] ", IIf(intType = 2, " And t.����=[4] ", " And t.��ʶ��=[4] "))
            End If

            If lngҩƷid <> 0 Then
                gstrSQL = gstrSQL & " And t.�շ�ϸĿid=[6]"
            End If

            If lng����ID <> 0 Then
                gstrSQL = gstrSQL & " And t.���˿���id=[7]"
            End If
        End If

    End If
    '�ѷ��Ͳ鿴
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '12' ����,����id,����,id,1 ҩʦ��˱�־,����,���� from  R where ����״̬=5"

    '��ǩ�ղ鿴
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '13' ����,����id,����,id,1 ҩʦ��˱�־,����,���� from  R where ����״̬=6"

    '�ܾ�ǩ�ղ鿴
    gstrSQL = gstrSQL & " Union All " & _
        " select distinct '14' ����,����id,����,id,1 ҩʦ��˱�־,����,���� from  R where ����״̬=7"
    '��������˲鿴
    gstrSQL = gstrSQL & " Union All " & _
        "Select distinct '15' As ����,a.���˲���id As ����id, '[' || b.���� || ']' || b.���� As ����,a.Id,1 ҩʦ��˱�־,����,b.����" & vbNewLine & _
        "       From (Select a.ID, ���˲���id" & vbNewLine & _
        "              From ��Һ��ҩ��¼ A,����ҽ����¼ B" & vbNewLine & _
        "              Where A.ҽ��id=B.���id and a.����id = [1] And a.ִ��ʱ�� Between [2] And [3] And Nvl(a.����״̬, 0) In (10,11)"
    
    If strMsg <> "" Then
        gstrSQL = gstrSQL & IIf(intType = 1, " And a.����=[4] ", IIf(intType = 2, " And a.����=[4] ", " And a.סԺ��=[4] "))
    End If
    
    If lngҩƷid <> 0 Then
        gstrSQL = gstrSQL & " And b.�շ�ϸĿid=[6]"
    End If
    
    If lng����ID <> 0 Then
        gstrSQL = gstrSQL & " And a.���˿���id=[7]"
    End If
            
    If strƿǩ�� <> "" Then
        gstrSQL = gstrSQL & " And a.ƿǩ��=[5]"
    End If
        
       gstrSQL = gstrSQL & ") A, ���ű� B" & vbNewLine & _
        "       Where a.���˲���id = b.Id"
        
    'ҽ�����˲鿴
    gstrSQL = gstrSQL & " Union All " & _
        " Select distinct '16' As ����, A.���˲���id As ����id, '[' || B.���� || ']' || B.���� As ����, A.ID,1 ҩʦ��˱�־,b.����,b.���� " & _
        " From ��Һ��ҩ��¼ A, ���ű� B ,����ҽ����¼ C" & _
        " Where A.���˲���id = B.ID and A.ҽ��id=c.���id And A.����״̬=12 And A.����id = [1] And A.ִ��ʱ�� Between [2] And [3] "
        
    If strMsg <> "" Then
        gstrSQL = gstrSQL & IIf(intType = 1, " And a.����=[4] ", IIf(intType = 2, " And a.����=[4] ", " And a.סԺ��=[4] "))
    End If
    
    If lngҩƷid <> 0 Then
        gstrSQL = gstrSQL & " And c.�շ�ϸĿid=[6]"
    End If
    
    If lng����ID <> 0 Then
        gstrSQL = gstrSQL & " And a.���˿���id=[7]"
    End If
            
    If strƿǩ�� <> "" Then
        gstrSQL = gstrSQL & " And a.ƿǩ��=[5]"
    End If
    
    gstrSQL = gstrSQL & " Order By ����, ���� )" & IIf(strSourceDep = "", "", " N, Table(Cast(f_Num2list([8]) As Zltools.t_Numlist)) S Where n.����id = s.Column_Value ")
        
    Set PIVA_GetTransCount = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Һ����Ŀ", lngCenterID, dateExeStart, dateExeEnd, strMsg, strƿǩ��, lngҩƷid, lng����ID, strSourceDep)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function





Public Sub PIVA_AnalysisTrans(ByVal lngCenterID As Long, ByVal dateStart As String, ByVal dateEnd As String)
    'PIVA��̨�������ֽⷢҩ����������Һ��
    'lngCenterID����Һ��������ID
    'dateStart��dateEnd����ҩ���ݵ�����ʱ�䷶Χ
    On Error GoTo ErrHand
    gstrSQL = "Zl_��Һ��ҩ��¼_Insert("
    '��������ID
    gstrSQL = gstrSQL & lngCenterID
    '��ʼʱ��
    gstrSQL = gstrSQL & "," & dateStart
    '����ʱ��
    gstrSQL = gstrSQL & "," & dateEnd
    gstrSQL = gstrSQL & ")"

    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Һ��ҩ��¼")
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Function DeptSendWork_Get��������() As Recordset
'��ȡ���˿������ƣ�ȡ��������Ϊ�ٴ�����Ĳ���
    On Error GoTo ErrHand
    
    gstrSQL = "Select distinct a.Id, a.����, a.����,zlSpellCode(a.����) ����,zlWBCode(a.����) ��ʼ���, a.����ʱ��" & vbNewLine & _
            "From ���ű� A, ��������˵�� B" & vbNewLine & _
            "Where a.Id = b.����id And (b.�������� = '�ٴ�' Or b.�������� = '����') And" & vbNewLine & _
            "      (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000/1/1', 'yyyy/mm/dd'))"
    
    
    Set DeptSendWork_Get�������� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function DeptSendWork_Get��ҩ����() As Recordset
'��ȡҩƷ����ҩ����
    On Error GoTo ErrHand
    gstrSQL = "select ����,���� from ��Һ��ҩ����"
    
    Set DeptSendWork_Get��ҩ���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ����")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_GetƵ��() As Recordset
'��ȡҩƷ����ҩ����
    On Error GoTo ErrHand
    gstrSQL = "select ����,����,Ӣ������ from ����Ƶ����Ŀ where ���� not like '-%'"
    
    Set DeptSendWork_GetƵ�� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡƵ��")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function PIVA_�Ѱ�ҩ��Һ��(ByVal lngCenterID As Long, ByVal dateExeStart As Date, ByVal lng����id As Long, ByVal lng��ҳID As Long) As Recordset
    '��ȡ�ò��˵�����Ѿ���ҩ��δ��ҩ����Һ��
    Dim strTmp As String
    
    On Error GoTo errHandle

    gstrSQL = "Select Distinct a.Id As ��ҩid, a.����״̬, a.��ҩ����, a.ִ��ʱ��, a.ƿǩ��, a.������Ա, a.����ʱ��, a.�Ƿ���, a.��ҩ����, e.No, f.���� As ͨ����, f.���," & vbNewLine & _
        "                e.����, j.���㵥λ As ������λ, (d.���� / g.סԺ��װ) As ����, g.סԺ��λ As ��λ, r.����ʱ�� As ҽ������ʱ��" & vbNewLine & _
        " From ��Һ��ҩ��¼ A, ��Һ��ҩ���� D, ҩƷ�շ���¼ E, �շ���ĿĿ¼ F, ҩƷ��� G, ������ĿĿ¼ J," & vbNewLine & _
        "     (Select ID From ����ҽ����¼ Where ����id = [4] And ��ҳid = [5] And ������� = 'E') M, ����ҽ������ R" & vbNewLine & _
        " Where a.Id = d.��¼id And d.�շ�id = e.Id And e.ҩƷid = f.Id And f.Id = g.ҩƷid And g.ҩ��id = j.Id And a.����id = [1] And" & vbNewLine & _
        "      a.ҽ��id = r.ҽ��id And a.���ͺ� = r.���ͺ� And a.ִ��ʱ�� Between [2] And [3] And" & vbNewLine & _
        "      ((a.����״̬ > 1 And a.����״̬ < 6) Or (a.����״̬ = 1 And a.�Ƿ�ȷ�ϵ��� = 1)) And a.ҽ��id = m.Id "
    
    Set PIVA_�Ѱ�ҩ��Һ�� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Һ��ҩ��¼", lngCenterID, _
        CDate(Format(dateExeStart, "yyyy-mm-dd 00:00:00")), CDate(Format(dateExeStart, "yyyy-mm-dd 23:59:59")), _
        lng����id, lng��ҳID)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function






