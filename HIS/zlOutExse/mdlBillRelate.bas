Attribute VB_Name = "mdlBillRelate"
Option Explicit
Public Function zlGetBillChargeExistInsure(ByVal lng����ID As Long, _
    Optional lng����ID As Long, Optional bln���� As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շ�(���˷�)��¼�е�ָ��ҽ������
    '���:lng����ID-����ID
    '����:lng����ID-���ز���ID
    '     bln����-�Ƿ���
    '����:��������򷵻ص��ݵ�ʱ������
    '����:���˺�
    '����:2014-06-18 16:22:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errHandle
    lng����ID = 0:  bln���� = False
    strSQL = "" & _
        "    Select B.��¼ID,B.����,B.����ID,A.�Ƿ���  " & _
        "    From ���ս����¼ B,������ü�¼ A " & _
        "    Where B.����=1  And B.��¼ID=[1]    " & _
        "          And B.��¼ID=A.����ID And A.���=1 and Rownum <2"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng����ID)
    If Not rsTemp.EOF Then
        lng����ID = NVL(rsTemp!����ID, 0)
        lng����ID = NVL(rsTemp!��¼ID, 0)
        bln���� = NVL(rsTemp!�Ƿ���, 0) = 1
        zlGetBillChargeExistInsure = NVL(rsTemp!����, 0)
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlIsCheckExiseSingularity(ByVal lng������� As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ���շ��Ƿ�����쳣�����ϵ���
    '���:lng�������-ָ���Ľ������
    '����:
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2014-06-18 17:08:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select  1  " & _
    "   From  ������ü�¼ A,����Ԥ����¼ B, ������ü�¼ C  " & _
    "   Where c.��¼���� = a.��¼���� And c.No = a.No And A.����ID=B.����ID And Mod(a.��¼����,10)=1 And c.��¼״̬=2 " & _
    "        And b.�������=[1] And Rownum <2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ������һ���շ��Ƿ�����Ѿ����ϵĵ���", lng�������)
    zlIsCheckExiseSingularity = rsTemp.RecordCount <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsCheckExistErrBill(ByVal lng������� As Long, Optional ByVal bln������� As Boolean, _
    Optional ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ���շ��Ƿ�����쳣����
    '���:lng�������-ָ���Ľ������
    '����:
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2014-06-18 17:08:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If bln������� Then
        strSQL = "" & _
        "   Select  1  " & _
        "   From  ���ò����¼ A" & _
        "   Where a.��¼����=1 And Nvl(a.����״̬,0)=1 And a.�������=[1] And Rownum <2"
    Else
        If strNos <> "" Then
            strSQL = "" & _
            "   Select /*+cardinality(j,10) */ 1" & vbNewLine & _
            "   From ������ü�¼ A, Table(f_Str2list([2])) J" & vbNewLine & _
            "   Where Mod(A.��¼����, 10) = 1 And a.No = j.Column_Value And Nvl(a.����״̬,0)=1 And Rownum < 2"
        Else
            strSQL = "" & _
            "   Select  1  " & _
            "   From  ������ü�¼ A,����Ԥ����¼ B" & _
            "   Where Mod(a.��¼����,10)=1 And A.����ID=B.����ID And Nvl(a.����״̬,0)=1 And b.�������=[1] And Rownum <2"
        End If
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ������һ���շ��Ƿ�����Ѿ����ϵĵ���", lng�������, strNos)
    zlIsCheckExistErrBill = rsTemp.RecordCount <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlFromIDGetChargeBalance(ByVal bytType As Byte, _
    ByVal strValue As String, Optional blnHistory As Boolean, _
    Optional ByRef blnDel As Boolean, Optional ByVal bln���쳣 As Boolean, _
    Optional ByVal byt��¼���� As Byte = 1) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID��ȡ�շѽ�����Ϣ
    '���:bytType-��������:0-���ݽ���ID����;1-���ݽ�����Ų���,2-���ݵ��ݺ�����ȡ���㷽ʽ
    '     strValue-Ҫ���ҵ�ֵ(Ϊ0ʱ,����ID,Ϊ1ʱ,�������,2ʱΪһ���շ����漰�����е���)
    '     blnDel-�˷ѽ���:true-���˷ѽ���;false-���˷ѽ���
    '     bln���쳣-�Ƿ�����쳣���㣬���ݵ��ݺ�����ȡ��������ʱ��Ч
    '     byt��¼����-2-���ݵ��ݺ�����ȡ���㷽ʽʱ���룬���ֹҺ�/�շ�
    '����:�շѽ���������Ϣ��
    '       �ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '            �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '       ����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    '����:���˺�
    '����:2014-06-24 16:37:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable As String, strWhere As String
    Dim strTable1 As String
    
    On Error GoTo errHandle
    strTable = IIf(blnHistory, "H", "") & "����Ԥ����¼"
    Select Case bytType
    Case 0  '0-���ݽ���ID����
        strWhere = " And  A.����ID= [1]"
    Case 1  ';1-���ݽ�����Ų���
        strWhere = "  And A.�������= [1]"
    Case 2 '���ݵ��ݺ�����ȡ��������
        strTable1 = "Select distinct ����ID  " & _
            "    From ������ü�¼ M " & _
            "    Where M.NO in (Select Column_value From Table(f_str2List([2])))  " & _
            "          And Mod(M.��¼����,10)=[3]" & IIf(bln���쳣, "", " And Nvl(M.����״̬,0)<>1")
        strTable1 = ",(" & strTable1 & ") Q1"
        If blnHistory Then strTable1 = Replace(strTable1, "������ü�¼", "H������ü�¼")
        strWhere = " And A.����ID=Q1.����ID"
    End Select

    If blnDel Then
        '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        strSQL = "" & _
            "   Select  A.ID,decode(A.��¼״̬,2,A.����ID,NULL) as ����ID," & _
            "        Case when Mod(A.��¼����,10)=1 then 1  " & _
            "             when B.���� is not null then  2 " & _
            "             when nvl(A.�����ID,0)<>0  then  3 " & _
            "             when J.���㷽ʽ is not null   then  4 " & _
            "             else 0 end as ����, " & _
            "        Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,A.��Ԥ��," & _
            "        decode(A.��¼״̬,2,A.ժҪ,NULL) as ժҪ,decode(A.��¼״̬,2,1,0) as �˷�," & _
            "        A.�����ID,A.���㿨���, " & _
            "        decode(A.��¼״̬,2,A.�������,NULL) as �������,decode(A.��¼״̬,2,A.����,NULL) as ����, " & _
            "        decode(A.��¼״̬,2,A.������ˮ��,NULL) as ������ˮ��,nvl(C.�Ƿ�����,0) as ���ƿ�, " & _
            "        nvl(C.�Ƿ�����,0) as �Ƿ�����,nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
            "        Decode(C.��������,NULL,0,1) as  �Ƿ�����,Nvl(C.�Ƿ�ת�ʼ�����,0) as �Ƿ�ת�ʼ�����," & _
            "        Nvl(C.�Ƿ��˿��鿨,0) as �Ƿ��˿��鿨," & _
            "        C.���� as ���������,decode(A.��¼״̬,2,A.����˵��,NULL) as ����˵��,A.�������,decode(A.��¼״̬,2,A.У�Ա�־,0) as У�Ա�־, " & _
            "        decode(B.����,Null,0,1) as ҽ��,0 as ���ѿ�id,nvl(q.����,1) as ��������" & _
            "   From " & strTable & " A ,ҽ�ƿ���� C,һ��ͨĿ¼ J,���㷽ʽ q," & _
            "        (Select ���� From ���㷽ʽ where ���� in (3,4)) B " & strTable1 & _
            "   Where A.���㷽ʽ=J.���㷽ʽ(+) And A.�����ID=C.ID(+) " & _
            "         And A.���㷽ʽ=B.����(+) and A.���㷽ʽ=q.����(+) " & _
            "         And (a.��¼���� In (1, 11) Or Nvl(a.���㿨���, 0) = 0) " & strWhere
            
        strSQL = strSQL & " Union ALL " & _
            "   Select A.ID,decode(A.��¼״̬,2,A.����ID,NULL) as ����ID,5 as  ����,Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,-1*nvl(b.Ӧ�ս��,0) as ��Ԥ��,A.ժҪ," & _
            "        decode(A.��¼״̬,2,1,0) as �˷�,A.�����ID,A.���㿨���," & _
            "        decode(A.��¼״̬,2,A.�������,NULL) as �������,decode(A.��¼״̬,2,B.����,NULL) as ����, " & _
            "        decode(A.��¼״̬,2,B.������ˮ��,NULL) as ������ˮ��,nvl(M.���ƿ�,0) as ���ƿ�, " & _
            "        nvl( M.�Ƿ�����,0) as �Ƿ�����,nvl(M.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
            "        nvl(M.�Ƿ�����,0) as  �Ƿ�����, 0 as �Ƿ�ת�ʼ�����,0 as �Ƿ��˿��鿨," & _
            "        M.���� as ���������,A.����˵��,A.�������,A.У�Ա�־,0 as ҽ��,B.���ѿ�id,nvl(q.����,1) as ��������" & _
            "   From  " & strTable & " A ,���˿������¼ B, " & _
            "        ���ѿ����Ŀ¼ M ,���㷽ʽ q " & strTable1 & _
            "   Where  a.Id = b.����id And a.���㿨��� = m.���  " & _
            "         and Mod(A.��¼����,10)<>1 and A.���㷽ʽ=q.����(+) " & strWhere

        strSQL = "" & _
            "   Select /*+ Rule */ max(����id) as ����id,����,max(�˷�) as �˷�,��¼����,���㷽ʽ,Max(ժҪ) as ժҪ,�����ID,���������,max(���ƿ�) as ���ƿ�,���㿨���, " & _
            "         max(�������) as �������,max(����) as ����,max(������ˮ��) as ������ˮ��, max(����˵��) as ����˵��, " & _
            "         �������,max(У�Ա�־) as У�Ա�־,ҽ��,���ѿ�id,��������,max(�Ƿ�ת�ʼ�����) as �Ƿ�ת�ʼ�����," & _
            "         Max(�Ƿ��˿��鿨) as �Ƿ��˿��鿨," & _
            "         max(�Ƿ�����) as �Ƿ�����,max(�Ƿ�ȫ��) as �Ƿ�ȫ��,max(�Ƿ�����) as �Ƿ����� , nvl(sum(��Ԥ��),0) as ��Ԥ��" & _
            "   From (" & strSQL & ") " & _
            "   Group by ����, ��¼����,���㷽ʽ,�����ID,���������,���㿨���,�������,ҽ��,���ѿ�id,�������� having  sum(��Ԥ��) <>0"
        Set zlFromIDGetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շѽ��㷽ʽ", Val(strValue), strValue, byt��¼����)
        Exit Function
    End If
    
    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    strSQL = "" & _
        "   Select /*+ Rule */ A.ID,A.����ID," & _
        "        Case when Mod(A.��¼����,10)=1 then 1  " & _
        "             when B.���� is not null then  2 " & _
        "             when nvl(A.�����ID,0)<>0  then  3 " & _
        "             when J.���㷽ʽ is not null   then  4 " & _
        "             else 0 end as ����, " & _
        "        Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,A.��Ԥ��," & _
        "        A.ժҪ,decode(A.��¼״̬,2,1,0) as �˷�," & _
        "        A.�����ID,A.���㿨���, " & _
        "        A.�������,A.����,A.������ˮ��,nvl(C.�Ƿ�����,0) as ���ƿ�, " & _
        "        nvl(C.�Ƿ�����,0) as �Ƿ�����,nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
        "        Decode(C.��������,NULL,0,1) as  �Ƿ�����,Nvl(C.�Ƿ�ת�ʼ�����,0) as �Ƿ�ת�ʼ�����," & _
        "        Nvl(C.�Ƿ��˿��鿨,0) as �Ƿ��˿��鿨," & _
        "        C.���� as ���������,A.����˵��,A.�������,A.У�Ա�־, " & _
        "        decode(B.����,Null,0,1) as ҽ��,0 as ���ѿ�id,nvl(q.����,1) as ��������" & _
        "   From " & strTable & " A ,ҽ�ƿ���� C,һ��ͨĿ¼ J,���㷽ʽ q," & _
        "        (Select ���� From ���㷽ʽ where ���� in (3,4)) B " & strTable1 & _
        "   Where A.���㷽ʽ=J.���㷽ʽ(+) And A.�����ID=C.ID(+) " & _
        "         And A.���㷽ʽ=B.����(+) and A.���㷽ʽ=q.����(+) " & _
        "         And (a.��¼���� In (1, 11) Or Nvl(a.���㿨���, 0) = 0) " & strWhere

    strSQL = strSQL & " Union ALL " & _
        "   Select /*+ Rule */ A.ID,A.����ID,5 as  ����,Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,-1*nvl(b.Ӧ�ս��,0) as ��Ԥ��,A.ժҪ," & _
        "        decode(A.��¼״̬,2,1,0) as �˷�,A.�����ID,A.���㿨���," & _
        "        A.�������,B.����,B.������ˮ��,nvl( M.���ƿ�,0) as ���ƿ�, " & _
        "        nvl( M.�Ƿ�����,0) as �Ƿ�����,nvl(M.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
        "        nvl(M.�Ƿ�����,0) as  �Ƿ�����,0 as �Ƿ�ת�ʼ�����,0 as �Ƿ��˿��鿨," & _
        "        M.���� as ���������,A.����˵��,A.�������,A.У�Ա�־,0 as ҽ��,B.���ѿ�id,nvl(q.����,1) as ��������" & _
        "   From  " & strTable & " A ,���˿������¼ B, " & _
        "        ���ѿ����Ŀ¼ M ,���㷽ʽ q " & strTable1 & _
        "   Where  a.Id = b.����id And a.���㿨��� = m.���  " & _
        "         and Mod(A.��¼����,10)<>1 and A.���㷽ʽ=q.����(+) " & strWhere
    gstrSQL = "" & _
        "   Select  ����ID,����,�˷�,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���," & _
        "           �������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id,��������," & _
        "           max(�Ƿ�ת�ʼ�����) as �Ƿ�ת�ʼ�����,max(�Ƿ�����) as �Ƿ�����,max(�Ƿ�ȫ��) as �Ƿ�ȫ��," & _
        "           max(�Ƿ��˿��鿨) as �Ƿ��˿��鿨," & _
        "           max(�Ƿ�����) as �Ƿ����� , nvl(sum(��Ԥ��),0) as ��Ԥ��" & _
        "   From (" & gstrSQL & ") " & _
        "   Group by ����ID,����,�˷�,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id,��������"
    Set zlFromIDGetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շѽ��㷽ʽ", Val(strValue), strValue, byt��¼����)
    
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet����ID(ByVal lng������� As Long, _
    ByRef strNos As String, Optional ByRef intInusre As Integer, _
    Optional ByVal blnNoMove As Boolean, _
    Optional ByRef lng����ID As Long, _
    Optional ByVal bln������ As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շ�(���˷�)��¼�е�ָ����ID
    '���:lng�������-�������
    '     blnNoMove-�Ƿ�ת�Ƶ���ʷ����
    '     bln������-�Ƿ񲹳����
    '����:strNOs-�����漰�ĵ��ݺ�
    '     intInusre-ҽ�����
    '     lng����ID-����ID
    '����:����ָ���Ľ���ID
    '����:���˺�
    '����:2014-06-18 16:22:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, str����ID As String
    
    On Error GoTo errHandle
    strNos = ""
    If bln������ Then
        '79142,Ƚ����,2014-11-3,�������˷�Ԥ����ʧ�ܲ������˷��쳣�����޷���ȡ����
        strSQL = "" & _
            "   Select Distinct a.����id As ����id, a.No, b.����," & _
            "          Decode(a.��¼״̬, 2, a.����id, 0) As ����id" & _
            "   From ���ò����¼ A," & _
            "        (Select Distinct s.No, t.����" & _
            "          From ���ò����¼ S, ���ս����¼ T" & _
            "          Where s.����id = t.��¼id And t.����(+) = 1 And s.������� = [1]) B" & _
            "   Where a.No = b.No(+) And a.������� = [1]" & _
            "   Order By NO"
    Else
        strSQL = "Select Distinct a.����id, a.No, b.����, Decode(a.��¼״̬, 2, a.����id, 0) As ����id" & vbNewLine & _
                " From ������ü�¼ A," & vbNewLine & _
                "      (Select Distinct s.����id, t.����" & vbNewLine & _
                "        From ����Ԥ����¼ S, ���ս����¼ T" & vbNewLine & _
                "        Where s.����id = t.��¼id(+) And s.������� = [1] And t.����(+) = 1) B" & vbNewLine & _
                " Where a.����id = b.����id And Mod(a.��¼����, 10) = 1" & vbNewLine & _
                " Order By NO"
    End If
    If blnNoMove Then
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL = Replace(strSQL, "���ս����¼", "H���ս����¼")
        If bln������ Then
            strSQL = Replace(strSQL, "���ò����¼", "H���ò����¼")
        End If
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݽ�����Ż�ȡ��Ӧ�Ľ���ID", lng�������)
    If rsTemp.EOF Then Exit Function
    
    lng����ID = 0
    With rsTemp
        strNos = "": str����ID = ""
        Do While Not .EOF
            If InStr(str����ID & ",", "," & !����ID & ",") = 0 Then
                str����ID = str����ID & "," & Val(NVL(!����ID))
            End If
            If InStr(strNos & ",", "," & !NO & ",") = 0 Then
                strNos = strNos & "," & NVL(!NO)
            End If
            If Val(NVL(rsTemp!����ID)) <> 0 And lng����ID = 0 Then
                lng����ID = Val(NVL(rsTemp!����ID))
            End If
            If intInusre = 0 Then intInusre = Val(NVL(!����))
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If str����ID <> "" Then str����ID = Mid(str����ID, 2)
    zlGet����ID = str����ID
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetAdviceFromID(ByVal strҽ��IDs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ��ID,��ȡ��Ӧ��ҽ������
    '���:strҽ��IDs-ҽ��ID(����ö��ŷ���)
    '����:
    '����:�ɹ�,����ҽ�����ݼ�(ҽ��ID,ҽ������)
    '����:���˺�
    '����:2014-06-27 11:52:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select ID as ҽ��ID,ҽ������ " & _
    "   From ����ҽ����¼ " & _
    "   Where ID in (Select Column_value From Table(f_num2List([1])))"
    Set zlGetAdviceFromID = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ������", strҽ��IDs)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetFromNOToLastBalanceID(ByVal strNos As String, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal bln��ʷ��ͬ���� As Boolean = False, _
    Optional lng������� As Long, Optional bln������ As Boolean = False) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ���շѵ��ݵ�NO���������һ����Ч�Ľ��ʵ�ID
    '���:blnNoMoved�Ƿ��ں󱸱��У���ѯ����֮ǰ���ж���Ҫ���������
    '     bln��ʷ��ͬ����-�Ƿ�������ʷ��һ���ѯ
    '     bln������-�Ƿ񲹳����
    '����:lng�������-�������һ����Ч�Ľ������
    '����:����ID
    '����:���˺�
    '����:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
    Dim i As Long
    
    On Error GoTo errHandle:
    '87975
    strSQL = "With c_���� As (Select Column_Value As NO From Table(f_Str2list([1])))" & vbNewLine & _
            " Select Max(a.����id) As ����id" & vbNewLine & _
            " From ������ü�¼ A, c_���� M" & vbNewLine & _
            " Where a.No = m.No" & vbNewLine & _
            "       And a.�Ǽ�ʱ�� + 0 =" & vbNewLine & _
            "           (Select Max(m.�Ǽ�ʱ��)" & vbNewLine & _
            "            From ������ü�¼ M, c_���� J" & vbNewLine & _
            "            Where m.No = j.No And Mod(m.��¼����, 10) = 1 And m.��¼״̬ In (1, 3) And Nvl(m.����״̬, 0) <> 1)" & vbNewLine & _
            "            And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And Nvl(a.����״̬, 0) <> 1"

    If bln������ Then
        strSQL = Replace(strSQL, "������ü�¼", "���ò����¼")
        strSQL = Replace(strSQL, "Max(a.����id)", "Max(a.����id)")
    End If

    strSQL = "" & _
            "   Select /*+ Rule */ A.����ID,B.������� " & _
            "   From (" & strSQL & ") A,����Ԥ����¼ B " & _
            "   Where A.����ID=B.����ID(+) And Rownum<2"

    If Not blnNOMoved And bln��ʷ��ͬ���� Then
        strSQL1 = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL1 = Replace(strSQL, "���ò����¼", "H���ò����¼")
        strSQL1 = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
        strSQL = strSQL & " Union ALL " & strSQL1
    ElseIf blnNOMoved Then
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL1 = Replace(strSQL, "���ò����¼", "H���ò����¼")
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݵ��ݻ�ȡ���һ���������ʵĽ���ID", strNos)

    If rsTemp.EOF Then Exit Function

    lng������� = Val(NVL(rsTemp!�������))
    zlGetFromNOToLastBalanceID = Val(NVL(rsTemp!����ID))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlInvoiceGetNOs(ByVal strInvioceNo As String, Optional cllInvoiceNoInfor As Collection, _
    Optional blnNOMoved As Boolean, Optional bln������ As Boolean = False) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ�Ʊ��,��ȡ��Ӧ�ĵ��ݺ�
    '���:strInvioceNo-��Ʊ��
    '     blnNOMoved-�Ƿ�����ʷ��ռ�
    '     bln������-�Ƿ�ҽ���������
    '����:cllInvoiceNoInfor-array(No,���)
    '����:�ɹ����ش���ķ�Ʊ���漰�ĵ��ݺ�
    '����:���˺�
    '����:2013-04-12 15:59:32
    '����:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strNos As String
    Dim strSQL1 As String, strSQL As String

    On Error GoTo errHandle
    Set cllInvoiceNoInfor = New Collection
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 And bln������ = False Then
        strSQL = "" & _
            "   Select  /*+ RULE */  A.NO,Max(A.���) as ���,Max(C.�������) as �������" & _
            "   From Ʊ�ݴ�ӡ��ϸ A,������ü�¼ B,����Ԥ����¼ C" & _
            "   Where A.Ʊ��=[1] and Ʊ��=1 and A.�Ƿ����<>1" & _
            "         And A.No=B.NO And Mod(B.��¼����,10)=1  And nvl(B.��¼״̬,0)<>2 And B.����ID=C.����ID" & _
            "   Group by A.NO"
        If blnNOMoved Then
            strSQL = Replace(strSQL, "Ʊ�ݴ�ӡ��ϸ", "HƱ�ݴ�ӡ��ϸ")
            strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
            strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ӧ���ݵķ�Ʊ��", strInvioceNo)

        strNos = ""
        With rsTemp
            Do While Not .EOF
                strNos = strNos & "," & NVL(!NO)
                cllInvoiceNoInfor.Add Array(NVL(!NO), NVL(!���))
                .MoveNext
            Loop
            If strNos <> "" Then
                zlInvoiceGetNOs = Mid(strNos, 2)
                Exit Function
            End If
        End With
    End If
    
    strSQL = "" & _
        "   Select  Distinct NO  " & _
        "   From Ʊ�ݴ�ӡ���� A, " & _
        "           (   Select Max(M.��ӡID) as ��ӡID " & _
        "               From  Ʊ��ʹ����ϸ M   " & _
        "               Where M.Ʊ��=1 And M.����=1 And M.����=[1]  " & _
        "               Group by M.����" & _
        "               )  Q" & _
        "   Where A.��������=1  And ID=Q.��ӡID " & _
        "   Order by NO"
    If blnNOMoved Then
        strSQL = Replace(strSQL, "Ʊ�ݴ�ӡ��ϸ", "HƱ�ݴ�ӡ��ϸ")
        strSQL = Replace(strSQL, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Ӧ���ݵķ�Ʊ��", strInvioceNo)

    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & NVL(!NO)
            .MoveNext
        Loop
        If strNos <> "" Then
            zlInvoiceGetNOs = Mid(strNos, 2)
        End If
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetChargeInsure(ByVal lng����ID As Long, ByRef lng����ID As Long, _
    Optional ByVal blnNOMoved As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѵ�ҽ����
    '���:lng����ID-����ID
    '     blnNOMoved-�Ƿ�����ת��
    '����:lng����ID-����ID
    '����:����
    '����:���˺�
    '����:2014-07-02 14:30:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String

    On Error GoTo errHandle
    lng����ID = 0
    strSQL = "" & _
        "    Select B.��¼ID,B.����,B.����ID,A.�Ƿ���  " & _
        "    From ������ü�¼ A,���ս����¼ B " & _
        "    Where A.����ID=[1] And  mod(A.��¼����,10)=1 " & _
        "         And B.����=1 And A.����ID=B.��¼ID and Rownum<2 "
    If blnNOMoved Then
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL = Replace(strSQL, "���ս����¼", "H���ս����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݽ���ID��ȡָ����ҽ������", lng����ID)
    If rsTemp.EOF Then Exit Function
    lng����ID = NVL(rsTemp!����ID, 0)
    zlGetChargeInsure = NVL(rsTemp!����, 0)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlMakeClinicPreSwapData(ByVal strStartFact As String, _
    ByVal lng����ID As Long, ByRef strNos As String, Optional bln������ As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݶ������ݴ���һ����¼��Ϣ(���ۼ۵�λ)
    '���:strStartFact-��ʼ��Ʊ��
    '     lng����ID-�����շѽ���IDs
    '����:strNos-���ر��ν����Nos
    '����:ҽ��������ݵ����ݼ�(�������(1--n),����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��)
    '����:���˺�
    '����:2014-07-07 11:24:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String
    Dim i As Integer, j As Integer, intStartPage As Integer, intPages As Integer
    Dim p As Integer, strSQL As String
    Dim dbl���� As Double, curʵ�� As Currency, curͳ�� As Currency
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    Dim strTable  As String, strWhere As String
    
    Err = 0: On Error GoTo Errhand:
    With rsTmp.Fields
        .Append "�������", adBigInt, 50, adFldIsNullable
        .Append "�ѱ�", adVarChar, 50, adFldIsNullable
        .Append "NO", adVarChar, 8, adFldIsNullable
        .Append "���", adBigInt, , adFldIsNullable '����:42961
        .Append "ʵ��Ʊ��", adVarChar, 20, adFldIsNullable
        .Append "����ʱ��", adDBTimeStamp, , adFldIsNullable
        .Append "����ID", adBigInt, , adFldIsNullable
        .Append "�շ����", adVarChar, 2, adFldIsNullable
        .Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
        .Append "���㵥λ", adVarChar, 50, adFldIsNullable
        '79420,���ϴ�,2014/11/10:������¼���ֶδ�С
        .Append "������", adVarChar, 100, adFldIsNullable
        .Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
        .Append "����", adDouble, , adFldIsNullable
        .Append "����", adDouble, , adFldIsNullable
        .Append "ʵ�ս��", adCurrency, , adFldIsNullable
        .Append "ͳ����", adCurrency, , adFldIsNullable
        .Append "����֧������ID", adBigInt, , adFldIsNullable
        .Append "�Ƿ�ҽ��", adBigInt, , adFldIsNullable
        .Append "���ձ���", adVarChar, 50, adFldIsNullable
        .Append "ժҪ", adVarChar, 2000, adFldIsNullable
        .Append "�Ƿ���", adBigInt, , adFldIsNullable
        .Append "��������ID", adBigInt, , adFldIsNullable
        .Append "ִ�в���ID", adBigInt, , adFldIsNullable
    End With
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    strTable = ""
    strWhere = " And A.����ID=[1]"
    If bln������ Then
       strTable = ",(Select distinct �շѽ���ID From ���ò����¼ Where ����ID=[1]) B"
       strWhere = " And A.����ID=b.�շѽ���ID"
    End If

    strSQL = "Select A.NO,Nvl( A.�۸񸸺�, A.���) as ���,To_char(max(A.�Ǽ�ʱ��),'YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
            "       A.����ID,A.�ѱ�,A.�շ����,A.�վݷ�Ŀ,A.���㵥λ,A.������," & _
            "       A.�շ�ϸĿID,A.���մ���ID As ����֧������ID,Nvl(A.������Ŀ��,0) As �Ƿ�ҽ��,A.���ձ���," & _
            "       Avg(Nvl(A.����,0)*A.����) As ����,Avg(A.��׼����) As ����," & _
            "       Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.ͳ����) As ͳ����,max(A.ժҪ) as ժҪ," & _
            "       nvl(A.�Ӱ��־,0) as �Ƿ���,A.��������ID,A.ִ�в���ID,A.����ID" & _
            " From ������ü�¼ A" & strTable & _
            " Where Mod(A.��¼����,10)=1 " & strWhere & _
            " Group By A.NO, Nvl(A.�۸񸸺�, A.���),A.����id, A.�ѱ�, A.�շ����, A.�վݷ�Ŀ, A.���㵥λ, A.������, A.�շ�ϸĿid, A.���մ���id, Nvl(A.������Ŀ��, 0), A.���ձ���, A.ժҪ, Nvl(A.�Ӱ��־, 0)," & _
            "       A.��������id, A.ִ�в���id,A.����ID"
    
    strSQL = "Select '" & strStartFact & "' as ʵ��Ʊ��,A.NO,A.���,max(A.����ʱ��) as ����ʱ��," & _
            "       A.����ID,A.�ѱ�,A.�շ����,A.�վݷ�Ŀ,A.���㵥λ,A.������," & _
            "       A.�շ�ϸĿID,A.����֧������ID,A.�Ƿ�ҽ��,A.���ձ���," & _
            "       sum(A.����) as ����,max(A.����) As ����, Sum(A.ʵ�ս��) As ʵ�ս��, " & _
            "       Sum(A.ͳ����) As ͳ����,max(A.ժҪ) as ժҪ," & _
            "       Max(A.�Ƿ���) as �Ƿ���,max(A.��������ID) as ��������ID,max(A.ִ�в���ID ) as ִ�в���ID " & _
            " From (" & strSQL & ") A" & _
            " Group By A.NO,A.���,A.����id, A.�ѱ�, A.�շ����, A.�վݷ�Ŀ, A.���㵥λ, A.������, A.�շ�ϸĿid, A.����֧������ID, " & _
            "       A.�Ƿ�ҽ��, A.���ձ���" & _
            " Order by NO,��� "

    Set rsNo = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����շ�����-ҽ��", lng����ID)
    If rsNo.RecordCount <> 0 Then rsNo.MoveFirst
    With rsNo
        p = 1: strNos = ""
        Do While Not rsNo.EOF
            rsTmp.AddNew
            rsTmp!������� = p
            rsTmp!�ѱ� = !�ѱ�
            rsTmp!NO = NVL(!NO)    '����ȡ���۵�ʱ����ֵ
            rsTmp!��� = Val(NVL(!���))    '����ȡ���۵�ʱ����ֵ
            rsTmp!ʵ��Ʊ�� = NVL(!ʵ��Ʊ��)
            rsTmp!����ʱ�� = !����ʱ��
            rsTmp!����ID = Val(NVL(!����ID))
            rsTmp!�շ���� = NVL(!�շ����)
            rsTmp!�վݷ�Ŀ = NVL(!�վݷ�Ŀ)
            rsTmp!������ = NVL(!������)
            rsTmp!�շ�ϸĿID = Val(NVL(!�շ�ϸĿID))
            rsTmp!���㵥λ = NVL(!���㵥λ)
            rsTmp!���� = Val(NVL(!����))
            rsTmp!���� = Val(NVL(!����))
            rsTmp!ʵ�ս�� = Val(NVL(!ʵ�ս��))
            rsTmp!ͳ���� = Val(NVL(!ͳ����))
            rsTmp!����֧������ID = IIf(Val(NVL(!����֧������ID)) = 0, Null, Val(NVL(!����֧������ID)))
            rsTmp!�Ƿ�ҽ�� = Val(NVL(!�Ƿ�ҽ��))
            rsTmp!���ձ��� = NVL(!���ձ���)
            rsTmp!ժҪ = NVL(!ժҪ)
            rsTmp!�Ƿ��� = Val(NVL(!�Ƿ���))
            rsTmp!��������ID = Val(NVL(!��������ID))
            rsTmp!ִ�в���ID = Val(NVL(!ִ�в���ID))
            rsTmp.Update
            If InStr(strNos & ",", "," & !NO & ",") = 0 Then
                strNos = strNos & "," & !NO
                p = p + 1
            End If
            rsNo.MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set zlMakeClinicPreSwapData = rsTmp
    
    Exit Function
Errhand:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlInsureCheck(ByVal str���ս��� As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ��ҽ���Ƿ���Ҫ�϶�
    '���:str���ս���-���ս���
    '       strAdvance-ҽ�����صĽ���
    '����:
    '����:��Ҫ�϶�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo errHandle
    If Not (strAdvance <> "" And str���ս��� <> strAdvance) Then Exit Function
    '��ʽ����ǰ��,���㷽ʽ�ͽ�����δ�����仯ʱ��У��
    blnMedicareCheck = True
    varData = Split(str���ս���, "||"): varData1 = Split(strAdvance, "||")

    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsureCheck = blnMedicareCheck
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceNos(ByVal bytType As Byte, _
    ByVal strFindValue As String, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional bln������ As Boolean = False, _
    Optional int��¼���� As Integer = 1) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ���շѵ��ݵ�NO�����ID�������ţ�����ͬһ�ν����NOs
    '���:bytType-0-����NO������;1-���ݽ���ID������,2-���ݽ������������
    '    strFindValue-���ҵ�ֵ
    '    blnNOMoved-�Ƿ��ں󱸱��У���ѯ����֮ǰ���ж���Ҫ���������
    '    bln������-�Ƿ�ҽ��������
    '����:��ʽ��"AAA,BBB,CCC',..."
    '����:���˺�
    '����:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNos As String
    Dim i As Long, strHistory As String
    Dim strFeeTable As String, strSQL1 As String

    On Error GoTo errHandle:
    Select Case bytType
    Case 0 '0-����NO������
        If bln������ Then
            strSQL = "" & _
            "   Select distinct A.NO " & _
            "   From ������ü�¼ A,(Select distinct �շѽ���ID as ����ID From ���ò����¼ Where NO=[1] and ��¼����=1 ) B" & _
            "   Where A.����ID=B.����ID" & _
            "   Order by NO"
        Else
            strSQL = "" & _
            "   Select distinct B.NO " & _
            "   From ������ü�¼ A,������ü�¼ B" & _
            "   Where A.NO=[1] and Mod(A.��¼����,10)=1 And A.����ID=B.����ID And a.��¼״̬ In (1, 3)" & _
            "   Order by NO"
        End If
    Case 1  '1-���ݽ���ID������
        If bln������ Then
            strSQL = "" & _
            "    Select Distinct A.No " & _
            "    From ������ü�¼ A," & _
            "        (Select distinct C1.�շѽ���ID as ����ID " & _
            "         From ���ò����¼ A1,���ò����¼ B1,���ò����¼ C1  " & _
            "         Where A1.����ID=[2] and A1.��¼����=1  " & _
            "               And A1.NO=B1.NO and A1.��¼����=B1.��¼���� " & _
            "               And B1.�������=C1.������� and C1.��¼״̬ in (1,3) ) B " & _
            "    Where A.����ID=B.����ID    " & _
            "    Order By NO"
        Else
            strSQL = "" & _
            "    Select Distinct c.No " & _
            "    From ������ü�¼ A,������ü�¼ B,������ü�¼ C " & _
            "    Where A.����ID=[2] And Mod(a.��¼����, 10) = 1 And a.No = b.No And  b.��¼���� = 1 " & _
            "          and b.����ID=C.����ID    " & _
            "    Order By NO"
        End If
    Case 2  '2-���ݽ������������
        If bln������ Then
            strSQL = "" & _
            "    Select Distinct A.No " & _
            "    From ������ü�¼ A," & _
            "        (Select distinct C1.�շѽ���ID as ����ID " & _
            "         From ���ò����¼ A1,���ò����¼ B1,���ò����¼ C1  " & _
            "         Where A1.�������=[2] and A1.��¼����=1  " & _
            "               And A1.NO=B1.NO and A1.��¼����=B1.��¼���� " & _
            "               And B1.�������=C1.������� and C1.��¼״̬ in (1,3) ) B " & _
            "    Where A.����ID=B.����ID    " & _
            "    Order By NO"
        Else
            strSQL = "" & _
            "   Select Distinct c.No " & _
            "   From ������ü�¼ A,������ü�¼ B,������ü�¼ C," & _
            "        (Select ����ID From ����Ԥ����¼ Where �������=[2]) D" & _
            "   Where A.����ID=D.����ID And Mod(a.��¼����, 10) = 1 And a.No = b.No And  b.��¼���� = 1 " & _
            "         and b.����ID=C.����ID    " & _
            "   Order By NO"
        End If
    End Select
    If blnNOMoved Then
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
        If bln������ Then
            strSQL = Replace(strSQL, "���ò����¼", "H���ò����¼")
        End If
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݵ��ݻ�ȡһ�ν��ʵĵ���", strFindValue, Val(strFindValue))
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & !NO
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    zlGetBalanceNos = strNos
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlIsErrChargeCancel(ByVal strNo As String, Optional ByVal lng����ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ���ĵ����Ƿ��쳣���շ����ϲ�������
    '���:strNO-��ָ���ĵ����ж�
    '     lng����ID-������ID�����ж�
    '����:���쳣�շ�����,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2014-07-29 14:11:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    If strNo <> "" Then
        strSQL = "   Select 1 From ������ü�¼ Where NO=[1] and ��¼����=1 And ��¼״̬ IN (1,3) and nvl(����״̬,0)=1"
    Else
        strSQL = "" & _
        "   Select 1 From ������ü�¼ A,������ü�¼ B  " & _
        "   Where A.NO=B.NO and A.��¼����=1 And A.��¼״̬ IN (1,3) and nvl(A.����״̬,0)=1 And B.����ID=[2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�Ϊ�쳣�շ�����", strNo, lng����ID)
    zlIsErrChargeCancel = Not rsTemp.EOF
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetFirstBalanceID(ByVal strNos As String, Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal bln��ʷ��ͬ���� As Boolean = False, _
    Optional lng������� As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ݵ�һ�ν���ID
    '���:blnNoMoved�Ƿ��ں󱸱��У���ѯ����֮ǰ���ж���Ҫ���������
    '     bln��ʷ��ͬ����-�Ƿ�������ʷ��һ���ѯ
    '����:lng�������-�������һ����Ч�Ľ������
    '����:���ص�һ�ν���ID
    '����:���˺�
    '����:2014-07-30 13:57:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
    Dim i As Long
    
    On Error GoTo errHandle:
    strSQL = "" & _
        "   Select M.����ID,Q.�������" & _
        "   From ������ü�¼ M,����Ԥ����¼ Q" & _
        "   Where M.����ID=Q.����ID And  M.NO IN (select Column_value From Table(f_str2List([1])) ) " & _
        "         And  M.��¼���� =1 And M.��¼״̬ IN (1,3) And rownum <2 "

    If Not blnNOMoved And bln��ʷ��ͬ���� Then
        strSQL1 = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL1 = Replace(strSQL1, "����Ԥ����¼", "H����Ԥ����¼")
        strSQL = strSQL & " Union ALL " & strSQL1
    ElseIf blnNOMoved Then
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݵ��ݻ�ȡ���һ���������ʵĽ���ID", strNos)
    If rsTemp.EOF Then Exit Function
    lng������� = Val(NVL(rsTemp!�������))
    zlGetFirstBalanceID = Val(NVL(rsTemp!����ID))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlExistDelFeeChargeBill(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݺ�,�ж��Ƿ�����˷ѵ���
    '���:strNos-ָ�����շѵ�
    '����:�����˷ѵ���,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle:
    strSQL = "" & _
        "   Select 1 " & _
        "   From ������ü�¼ M" & _
        "   Where  M.NO in (Select Column_value From Table(f_str2List([1]))  )  " & _
        "       And Mod(M.��¼����,10)=1 And M.��¼״̬ =2 And Rownum <2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ�����˷ѵ���", strNos)
    If rsTemp.EOF Then Exit Function
    zlExistDelFeeChargeBill = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlIsMulitOneBalance(ByVal strNos As String, Optional ByRef lng����ID As Long, _
    Optional ByRef lng������� As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ�������Ƿ�൥��һ�ν�������
    '���:strNos-���ݺ�(����ö��ŷָ�)
    '����:lng����ID-����ID
    '     lng�������-����һ�ν���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-08-01 14:42:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle:
    strSQL = "" & _
        "   Select A.����ID,A.������� " & _
        "   From  ����Ԥ����¼ A,������ü�¼ M" & _
        "   Where  A.����ID=M.����ID And  M.NO in (Select Column_value From Table(f_str2List([1]))  )  " & _
        "       And  M.��¼���� =1 And M.��¼״̬ in (1,3) And Rownum <2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ�����˷ѵ���", strNos)
    If rsTemp.EOF Then Exit Function
    lng����ID = Val(NVL(rsTemp!����ID))
    lng������� = Val(NVL(rsTemp!�������))
    zlIsMulitOneBalance = lng������� < 0
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlRePrintReplenishTheBalanceBill(frmParent As Object, ByVal lngModule As Long, _
    ByVal bytType As Byte, strNo As String, ByVal intInsure As Integer, _
    ByVal objInvoice As zlPublicExpense.clsInvoice, _
    ByVal objFact As zlPublicExpense.clsFactProperty, _
    Optional blnDelOpt As Boolean, Optional DateDel As Date, Optional blnVirtualPrint As Boolean, _
    Optional ByVal blnDelRecord As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���´�ӡҽ���������Ʊ��
    '���:1-�ش�;2-����
    '       strNO -ָ��Ҫ�ش�ĵ��ݺ�
    '       intInsure-ҽ����
    '       objInvoice-��Ʊ����
    '       blnDelOpt-�˷��ش��������
    '       DateDel-�˷�ʱ��
    '       blnVirtualPrint-ҽ���ӿڴ�ӡƱ�ݣ�HIS������ӡֻ��Ʊ��
    '       blnDelRecord-�ش�ʱ���Ƿ��Ƕ��˷Ѽ�¼�����ش�(Ŀǰֻ�б���ҽ��(ҽ���ӿڴ�ӡƱ��)������)
    '����:
    '����:��ӡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-30 10:22:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strInvoice As String, strInfo As String
    Dim j As Integer, blnValid As Boolean, blnInput As Boolean
    Dim blnDo As Boolean, blnHaveInvoice As Boolean
    Dim lngLastUseIDTemp As Long
    Dim strRptName As String, strInvoiceNO As String
    Dim bytPrintType As Byte
    Dim lngUseId As Long

    blnHaveInvoice = objFact.LastUseID <> 0     '��Ҫ���˷���,����������õ�,�������շ�Ʊ,Ȼ���ش�Ʊ:30386
    If blnHaveInvoice = False And blnDelOpt Then
        blnHaveInvoice = objInvoice.zlCheckBillNOIsPrintInvoice(1, strNo)
    End If

    strRptName = "ZL" & glngSys \ 100 & "_BILL_1124"

    lngUseId = objFact.LastUseID
    '����ϸ����Ʊ��ʹ��
    If objFact.�ϸ���� Then
        '��ʱֻ�ж��Ƿ���,��ӡ֮ǰ�ٸ��������ж��Ƿ���
        If objInvoice.zlGetInvoiceGroupID(1124, UserInfo.����, EM_�շ��վ�, objFact.ʹ�����, lngUseId, objFact.��������ID, lngUseId, 1, strInvoiceNO) = False Then Exit Function
        Select Case lngUseId
            Case -1
                MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Case -2
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
        End Select
        If lngUseId <= 0 Then Exit Function
    End If

    If Not gobjTax Is Nothing And gblnTax Then
        blnDo = True
    Else
        If blnDelOpt Then
            blnDo = True
        Else
            If objFact.��ӡ��ʽ = 0 Then   '��ȱʡƱ�ݸ�ʽ��ʾ
                objFact.��ӡ��ʽ = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
            End If
            SetReportPrintSet gcnOracle, glngSys, strRptName, "Format", objFact.��ӡ��ʽ
            '����û�и�ʽ�Ĵ���,���,��Ҫǿ��ȱʡ��ָ����ʽ
            blnDo = ReportPrintSet(gcnOracle, glngSys, strRptName, frmParent)
            'ȡ��ѡ��ĸ�ʽ
            objFact.��ӡ��ʽ = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
        End If
    End If

    If blnDo Then
        'ȡ��һ��Ʊ�ݺ���
        If Not objFact.�ϸ���� Then

            '�п����ǵ�һ��ʹ��
            Do
                blnInput = False
                '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
                strInvoice = UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, 1124, ""))
                If strInvoice = "" Then
                    strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                    vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = zlstr.Increase(strInvoice)
                    strInvoice = UCase(InputBox("��ȷ��" & IIf(bytType = 1, "�ش�", "����") & "ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If

                '�û�ȡ������,�����ӡ
                If strInvoice = "" Then
                    If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    blnValid = True
                Else
                    '���������Ч��
                    If blnInput Then
                        If zlCommFun.ActualLen(strInvoice) <> objFact.Ʊ�ų��� Then
                            MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & objFact.Ʊ�ų��� & " λ��", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    Else
                        blnValid = True
                    End If
                End If
            Loop While Not blnValid
        Else
            Do
                '����Ʊ�����ö�ȡ
                blnInput = False
                If objInvoice.zlGetNextBill(1124, lngUseId, strInvoice) = False Then
                    strInvoice = ""
                End If

                If strInvoice = "" Then
                    '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
                    '30386:��ӡ�˷�Ʊ��,�����ش��ٷ���
                    If frmInputBox.InputBox(frmParent, "��ʼ��Ʊ��", "�޷�����Ʊ�����������ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                    vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoice, _
                                    blnHaveInvoice And blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                Else
                    '30386
                    If frmInputBox.InputBox(frmParent, "��ʼ��Ʊ��", "��ȷ��" & IIf(bytType = 1, "�ش�", "����") & "ʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoice, _
                                    blnHaveInvoice And blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                End If

                '�û�ȡ������,����ӡ
                If strInvoice = "" Then Exit Function

                '���������Ч��
                If blnInput Then
                    If objInvoice.zlGetInvoiceGroupID(1124, UserInfo.����, EM_�շ��վ�, objFact.ʹ�����, lngUseId, objFact.��������ID, lngLastUseIDTemp, 1, strInvoiceNO) = False Then Exit Function
                    If lngLastUseIDTemp = -3 Then
                        MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                    Else
                        lngUseId = lngLastUseIDTemp
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        End If

        bytPrintType = IIf(blnDelOpt, 3, 2)
        If blnDelOpt Then
            Call frmReplenishTheBalancePrint.ReportPrint(bytPrintType, strNo, intInsure, objFact, "", lngUseId, strInvoice, DateDel, blnVirtualPrint)
        Else
            Call frmReplenishTheBalancePrint.ReportPrint(bytPrintType, strNo, intInsure, objFact, "", lngUseId, strInvoice, zlDatabase.Currentdate, blnVirtualPrint, blnDelRecord)
        End If
        zlRePrintReplenishTheBalanceBill = True
    End If
End Function

Public Function zlPrintReplenishTheDelBalanceBill(frmParent As Object, ByVal lngModule As Long, _
    ByVal lng������� As Long, ByVal intInsure As Integer, _
    ByVal objInvoice As zlPublicExpense.clsInvoice, _
    ByVal objFact As zlPublicExpense.clsFactProperty, _
    Optional blnDelOpt As Boolean, Optional DateDel As Date, Optional blnVirtualPrint As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡҽ����������˷�Ʊ��(��Ʊ)
    '���:
    '       lng������� -ָ��Ҫ��ӡ���ݵĽ������
    '       intInsure-ҽ����
    '       objInvoice-��Ʊ����
    '       blnDelOpt-�˷��ش��������
    '       DateDel-�˷�ʱ��
    '       blnVirtualPrint-ҽ���ӿڴ�ӡƱ�ݣ�HIS������ӡֻ��Ʊ��
    '����:
    '����:��ӡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-30 10:22:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strInvoice As String, strInfo As String
    Dim j As Integer, blnValid As Boolean, blnInput As Boolean
    Dim blnDo As Boolean, blnHaveInvoice As Boolean
    Dim lngLastUseIDTemp As Long
    Dim strRptName As String, strInvoiceNO As String
    Dim bytPrintType As Byte
    Dim lngUseId As Long
    
    strRptName = "ZL" & glngSys \ 100 & "_BILL_1124_3"
    
    lngUseId = objFact.LastUseID
    '����ϸ����Ʊ��ʹ��
    If objFact.�ϸ���� Then
        '��ʱֻ�ж��Ƿ���,��ӡ֮ǰ�ٸ��������ж��Ƿ���
        If objInvoice.zlGetInvoiceGroupID(1124, UserInfo.����, EM_�շ��վ�, objFact.ʹ�����, lngUseId, objFact.��������ID, lngUseId, 1, strInvoiceNO) = False Then Exit Function
        Select Case lngUseId
            Case -1
                MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Case -2
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
        End Select
        If lngUseId <= 0 Then Exit Function
    End If


    If Not gobjTax Is Nothing And gblnTax Then
        blnDo = True
    Else
        If blnDelOpt Then
            blnDo = True
        Else
            If objFact.��ӡ��ʽ = 0 Then   '��ȱʡƱ�ݸ�ʽ��ʾ
                objFact.��ӡ��ʽ = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
            End If
            SetReportPrintSet gcnOracle, glngSys, strRptName, "Format", objFact.��ӡ��ʽ
            '����û�и�ʽ�Ĵ���,���,��Ҫǿ��ȱʡ��ָ����ʽ
            blnDo = ReportPrintSet(gcnOracle, glngSys, strRptName, frmParent)
            'ȡ��ѡ��ĸ�ʽ
            objFact.��ӡ��ʽ = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
        End If
    End If
    
    If blnDo Then
        'ȡ��һ��Ʊ�ݺ���
        If Not objFact.�ϸ���� Then
            
            '�п����ǵ�һ��ʹ��
            Do
                blnInput = False
                '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
                strInvoice = UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, 1124, ""))
                If strInvoice = "" Then
                    strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                    vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = zlCommFun.IncStr(strInvoice)
                    strInvoice = UCase(InputBox("��ȷ��ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If
                    
                '�û�ȡ������,�����ӡ
                If strInvoice = "" Then
                    If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    blnValid = True
                Else
                    '���������Ч��
                    If blnInput Then
                        If zlCommFun.ActualLen(strInvoice) <> objFact.Ʊ�ų��� Then
                            MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & objFact.Ʊ�ų��� & " λ��", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    Else
                        blnValid = True
                    End If
                End If
            Loop While Not blnValid
        Else
            Do
                '����Ʊ�����ö�ȡ
                blnInput = False
                If objInvoice.zlGetNextBill(1124, lngUseId, strInvoice) = False Then
                    strInvoice = ""
                End If
                
                If strInvoice = "" Then
                    '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
                    '30386:��ӡ�˷�Ʊ��,�����ش��ٷ���
                    If frmInputBox.InputBox(frmParent, "��ʼ��Ʊ��", "�޷�����Ʊ�����������ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                    vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoice, _
                                    False, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                Else
                    '30386
                    If frmInputBox.InputBox(frmParent, "��ʼ��Ʊ��", "��ȷ��ʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoice, _
                                    blnHaveInvoice And blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                End If
                
                '�û�ȡ������,����ӡ
                If strInvoice = "" Then Exit Function
                
                '���������Ч��
                If blnInput Then
                    If objInvoice.zlGetInvoiceGroupID(1124, UserInfo.����, EM_�շ��վ�, objFact.ʹ�����, lngUseId, objFact.��������ID, lngLastUseIDTemp, 1, strInvoiceNO) = False Then Exit Function
                    If lngLastUseIDTemp = -3 Then
                        MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                    Else
                        lngUseId = lngLastUseIDTemp
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        End If
        
        '1-�µ���ӡ,2-�ش�,3-�˷Ѵ�ӡ,4-����Ʊ��(ֻ��:2-��ϵͳԤ�������3-�û��Զ�����ʱ��ת��),6-�˷�Ʊ��(��Ʊ)��ӡ
        Call frmReplenishTheBalancePrint.ReportPrint(6, lng�������, intInsure, objFact, "", _
            lngUseId, strInvoice, DateDel, blnVirtualPrint)
        
        zlPrintReplenishTheDelBalanceBill = True
    End If
End Function

Public Function zlCheckRegBillIsExecuted(ByVal strNo As String, ByVal bln��������ҽ�� As Boolean, _
    ByRef blnExecuted_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ���ĹҺŵ����Ƿ��Ѿ���ִ��,����ҽ��������ҽ�������Ϻ�,ȡ������,Ҳ��ʾִ�й���
    '���:strNO-�Һŵ���
    '     bln��������ҽ��-�Ƿ񲻰��������ϵ�ҽ��
    '����:
    '����:True ��ʾ�ѱ�ִ��
    '����:���˺�
    '����:2014-10-10 11:16:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    blnExecuted_Out = False
    If bln��������ҽ�� Then strSQL = " And ҽ��״̬<>4"
    strSQL = _
        " Select count(ID) num From ���˹Һż�¼ Where NO=[1] And ִ��״̬>0 and ��¼����=1 and ��¼״̬ =1 " & _
        " Union All " & _
        " Select count(ID) num From ����ҽ����¼ Where �Һŵ�=[1] And (������Դ=1 or ������Դ=2)" & strSQL
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNo)
    Do While Not rsTmp.EOF
        If rsTmp!Num > 0 Then
            blnExecuted_Out = True
        End If
        rsTmp.MoveNext
    Loop
    zlCheckRegBillIsExecuted = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetInsureBalanceDetail(ByVal bytType As Byte, _
    ByVal strValue As String, Optional blnHistory As Boolean) As ADODB.Recordset
    '����:��ȡҽ��������ϸ����
    '���:bytType-��������:0-���ݽ���ID����;1-���ݽ�����Ų���,2-���ݵ��ݺ�����ȡ���㷽ʽ
    '     strValue-Ҫ���ҵ�ֵ(Ϊ0ʱ,����ID,Ϊ1ʱ,�������,2ʱΪһ���շ����漰�����е���)
    '����:����ҽ��������ϸ��¼
    '       �ֶ�:����id,NO,���㷽ʽ,���
    '����:Ƚ����
    '����:2015-07-13
    Dim strSQL As String, strWhere As String
    Dim strTable As String, strTable1 As String
    
    On Error GoTo errHandle
    strTable = IIf(blnHistory, "H", "") & "ҽ��������ϸ A"
    Select Case bytType
    Case 0  '0-���ݽ���ID����
        strWhere = " And  A.����ID= [1]"
    Case 1  '1-���ݽ�����Ų���
        strTable1 = "Select distinct ����ID  " & _
            "    From ����Ԥ����¼ Where �������= [1]"
        strTable1 = ",(" & strTable1 & ") B"
        If blnHistory Then strTable1 = Replace(strTable1, "����Ԥ����¼", "H����Ԥ����¼")
        strWhere = "  And A.����ID = B.����ID"
    Case 2 '2-���ݵ��ݺ�����ȡ��������
        strWhere = " And A.NO in (Select Column_value From Table(f_str2List([2])))"
    End Select
    
    strSQL = "Select a.����id, a.NO, a.���㷽ʽ, a.���" & _
            " From " & strTable & strTable1 & _
            " Where 1=1 " & strWhere
    Set zlGetInsureBalanceDetail = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ��������ϸ����", Val(strValue), strValue)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetYBBalanceNo(ByVal lng����ID As Long, Optional ByVal strNos As String, _
    Optional ByVal lng����ID As Long, Optional ByVal intInsure As Integer, _
    Optional ByVal blnDelCheck As Boolean, Optional ByVal blnHistory As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ݻ�ȡҽ��ԭ���㷽ʽ�ͽ�����
    '������
    '   strNOs - ���ݺ�,����ö��Ÿ�����A0001,A0002,...
    '   blnDelCheck - �Ƿ������������������
    '����:���ؽ�����Ϣ,��ʽ:���㷽ʽ|������||...
    '����:���˺�
    '����:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, varNos As Variant, strFilter As String
    Dim rsBalance As ADODB.Recordset, i As Integer, p As Integer
    Dim colBalance As Collection, strTemp As String
    
    On Error GoTo errHandle
    If blnDelCheck And intInsure = 0 Then Exit Function
    
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    Set rsBalance = zlGetInsureBalanceDetail(0, lng����ID, blnHistory)
    If strNos <> "" Then
        varNos = Split(strNos, ",")
        For i = 0 To UBound(varNos)
            If UBound(varNos) < 1 Then 'һ�ŵ���
                strFilter = " or No='" & varNos(i) & "'"
            Else '���ŵ���
                strFilter = strFilter & " or No='" & varNos(i) & "'"
            End If
        Next
        If strFilter <> "" Then strFilter = Mid(strFilter, 4)
        rsBalance.Filter = strFilter
    End If
    rsBalance.Sort = "No"
    If rsBalance.RecordCount = 0 Then Exit Function
    
    Set colBalance = New Collection
    p = 1: colBalance.Add Array()
    If rsBalance.RecordCount > 0 Then
        With rsBalance
            strTemp = NVL(!NO)
            Do While Not .EOF
                If strTemp <> NVL(!NO) Then
                    p = p + 1: colBalance.Add Array()
                    strTemp = NVL(!NO)
                End If
                If blnDelCheck Then
                    '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
                    If gclsInsure.GetCapability(support�����������, lng����ID, intInsure, NVL(!���㷽ʽ)) Then
                        str���㷽ʽ = NVL(!���㷽ʽ) & "|" & -1 * Val(NVL(!���))
                    End If
                Else
                    str���㷽ʽ = NVL(!���㷽ʽ) & "|" & Val(NVL(!���))
                End If
                
                Call SetBalanceVal(colBalance, p, str���㷽ʽ)
                .MoveNext
            Loop
        End With
    End If
    zlGetYBBalanceNo = GetMedicareStr(colBalance)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckOnlyUseTrans(ByVal str������� As String) As Boolean
    '���ܣ����ҽ����������Ƿ�����ܷ��ý��
    '��Σ�
    '   str������� - �������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    'ʣ����ý��
    strSQL = "Select Nvl(Sum(��Ԥ��), 0) As ���" & vbNewLine & _
            " From (Select Nvl(Sum(��Ԥ��), 0) As ��Ԥ��" & vbNewLine & _
            "       From ����Ԥ����¼" & vbNewLine & _
            "       Where ����id In (Select �շѽ���id From ���ò����¼ Where ��¼���� = 1 And ������� = [1])" & vbNewLine & _
            "       Union All"
    '���ζ�ʣ����õ�ҽ���������
    strSQL = strSQL & vbNewLine & _
            "       Select -1 * Nvl(Sum(a.��Ԥ��), 0) As ��Ԥ��" & vbNewLine & _
            "       From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
            "       Where a.��¼״̬ = 1 And a.���㷽ʽ = b.���� And b.���� In (3, 4) And a.�������= [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ҽ����������Ƿ�����ܷ��ý��", Val(str�������))
    zlCheckOnlyUseTrans = Val(NVL(rsTemp!���)) < 0
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetCanDelBalanceRecords(ByVal lng������� As Long, ByVal lng�����ID As Long) As ADODB.Recordset
    '���ܣ���ȡ���������˷ѵĽ��׽����¼
    '��Σ�
    '   lng������� - �������
    '   lng�����ID - ҽ�ƿ����
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    'ԭ�����ܽ�ע���ȥ������ǰ���˷ѵ�
    strSQL = _
        "Select Max(Ԥ��id) As ԭ����ID, Max(ԭ����id) As ����id, Max(����) As ����, Max(������ˮ��) As ������ˮ��," & vbNewLine & _
        "       Max(����˵��) As ����˵��,Max(��Ԥ��) As ��Ԥ��" & vbNewLine & _
        "From (Select Decode(a.��¼״̬, 2, f.Id, e.Id) As Ԥ��id, Decode(a.��¼״̬, 2, f.����id, e.����id) As ԭ����id," & vbNewLine & _
        "             Decode(a.��¼״̬, 2, f.����, e.����) As ����, Decode(a.��¼״̬, 2, f.������ˮ��, e.������ˮ��) As ������ˮ��," & vbNewLine & _
        "             Decode(a.��¼״̬, 2, f.����˵��, e.����˵��) As ����˵��, e.��Ԥ��, a.����id" & vbNewLine & _
        "      From ������ü�¼ A, ����Ԥ����¼ E, ������ü�¼ B, ����Ԥ����¼ F, ���ò����¼ C" & vbNewLine & _
        "      Where a.��¼���� = b.��¼���� And a.No = b.No And a.��� = b.��� And a.����id = e.����id And b.����id = f.����id" & vbNewLine & _
        "            And b.����id = c.�շѽ���id And b.��¼״̬ <> 2 And e.�����id = [2] And f.�����id = [2]" & vbNewLine & _
        "            And c.��¼���� = 1 And c.������� = [1]" & vbNewLine & _
        "            And Not Exists (Select 1" & vbNewLine & _
        "                   From ����Ԥ����¼" & vbNewLine & _
        "                   Where ������� In (Select m.�������" & vbNewLine & _
        "                       From ���ò����¼ M, ���ò����¼ N" & vbNewLine & _
        "                       Where m.��¼���� = n.��¼���� And m.No = n.No And n.��¼���� = 1 And n.������� = [1])" & vbNewLine & _
        "                             And ����id = a.����id))" & vbNewLine & _
        "Group By ����id"

    '����������˷ѽ��
    strSQL = strSQL & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select 0, ��¼id, '', '', '', -1 * ���" & vbNewLine & _
        "From �����˿���Ϣ" & vbNewLine & _
        "Where ��¼id In (Select �շѽ���id From ���ò����¼ Where ��¼���� = 1 And ������� = [1])"
    strSQL = _
        "Select Max(ԭ����ID) As ԭ����ID, ����id, Max(����) As ����," & vbNewLine & _
        "       Max(������ˮ��) As ������ˮ��, Max(����˵��) As ����˵��, Nvl(Sum(��Ԥ��), 0) As ���" & vbNewLine & _
        "From (" & strSQL & ")" & vbNewLine & _
        "Group By ����id" & vbNewLine & _
        "Having Nvl(Sum(��Ԥ��), 0) > 0" & vbNewLine & _
        "Order By ����id"
    Set zlGetCanDelBalanceRecords = zlDatabase.OpenSQLRecord(strSQL, "mdlBillRelate", lng�������, lng�����ID)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlCheckOtherSessionDoing(ByVal lng������� As Long, Optional ByVal strNos As String) As Boolean
    '����:��鵱ǰ�����Ƿ����ڱ������Ự����
    '���:lng�������-ָ���Ľ������
    '     strNos  - ���ݺ�
    '����:
    '����:�Ƿ���true,���򷵻�False
    '˵����"����Ԥ����¼.�Ự��"��ʽ��V$session.SID+'_'+V$session.SERIAL#
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If lng������� = 0 And strNos = "" Then zlCheckOtherSessionDoing = False: Exit Function
    If strNos <> "" Then
        strSQL = "Select /*+cardinality(j,10) */ 1" & vbNewLine & _
                " From ����Ԥ����¼ A, ������ü�¼ B, Table(f_Str2list([2])) J, V$session C" & vbNewLine & _
                " Where a.����id = b.����id And Mod(b.��¼����, 10) = 1 And b.No = j.Column_Value" & vbNewLine & _
                "       And a.�Ự�� = c.Sid || '_' || c.Serial# And c.Username Is Not Null" & vbNewLine & _
                "       And c.Audsid <> Userenv('sessionid') And Upper(c.Status) In ('ACTIVE', 'INACTIVE') And Rownum < 2"
    Else
        strSQL = "Select 1" & vbNewLine & _
                " From ����Ԥ����¼ A, V$session B" & vbNewLine & _
                " Where a.�Ự�� = b.Sid || '_' || b.Serial# And (a.������� = [1] Or a.����ID = [1])" & vbNewLine & _
                "       And b.Username Is Not Null And b.Audsid <> Userenv('sessionid')" & vbNewLine & _
                "       And Upper(b.Status) In ('ACTIVE', 'INACTIVE') And Rownum < 2"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ�����Ƿ����ڱ������Ự����", lng�������, strNos)
    zlCheckOtherSessionDoing = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetForceDelToCashNote(ByRef cllForceDelToCash As Collection) As String
    '��ȡǿ������ժҪ������"����˵��"�Զ��У���ʽ��XXXXǿ������:XXX��;XXX��
    '��Σ�
    '   cllForceDelToCash Array(����Ա,���������)
    Dim str����Ա As String
    Dim strTemp As String, i As Integer
    
    On Error GoTo errHandler
    If cllForceDelToCash Is Nothing Then Exit Function
    If cllForceDelToCash.Count = 0 Then Exit Function
    
    str����Ա = cllForceDelToCash(1)(0)
    For i = 1 To cllForceDelToCash.Count
        strTemp = strTemp & ";" & cllForceDelToCash(i)(1)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    GetForceDelToCashNote = str����Ա & "ǿ�����֣�" & strTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ThreeBalanceCheck(frmMain As Form, ByVal lngModule As Long, _
    ByVal objCard As Card, ByRef cllForceDelToCash As Collection, _
    Optional ByVal str��������� As String, Optional ByRef blnǿ������ As Boolean) As Boolean
    '������ǿ�����ּ��
    '��Σ�
    '   objCard ҽ�ƿ���Ϣ
    '   str��������� ���������
    '���Σ�
    '   cllForceDelToCash ǿ��������Ϣ��Array(����Ա,���������)
    '���أ�����ǿ�����֣�����True�����򣬷���False
    '105432
    Dim str����Ա As String
    
    On Error GoTo errHandler
    blnǿ������ = False
    If cllForceDelToCash Is Nothing Then Set cllForceDelToCash = New Collection
    
    If objCard Is Nothing Then
        If MsgBox("δ�ҵ�ָ����ҽ�ƿ����޷��жϸ�ҽ�ƿ��Ƿ�֧�����֣���ȷ��Ҫǿ����Ϊ�������㷽ʽ��", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If Not (objCard.�ӿ���� > 0 And Not objCard.���ѿ�) Then ThreeBalanceCheck = True: Exit Function
        If objCard.�Ƿ����� Then ThreeBalanceCheck = True: Exit Function
    End If
    
    If zlstr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "�����˿�ǿ������") Then
        If MsgBox("��" & str��������� & "����֧�����֣���ȷ��Ҫ����ǿ��������", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        'Array(����Ա,���������)
        cllForceDelToCash.Add Array(UserInfo.����, str���������)
    Else
        str����Ա = zlDatabase.UserIdentifyByUser(frmMain, "��" & str��������� & "��ǿ�����֣�Ȩ����֤��", _
            glngSys, lngModule, "�����˿�ǿ������", , True)
        If str����Ա = "" Then Exit Function
        'Array(����Ա,���������)
        cllForceDelToCash.Add Array(str����Ա, str���������)
    End If
    blnǿ������ = True
    ThreeBalanceCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelXMLExpend(ByVal lng������� As String, ByVal bln�쳣���� As Boolean) As String
    '��ȡ�����������˷ѽӿ�zlRetuenCheck��strXMLExpend����ֵ
    '��Ҫ�����շ��쳣��������
    '��Σ�
    '   lng������� - �������
    '   bln�쳣���� - �Ƿ��շ��쳣�����쳣������
    Dim i As Integer, strPriorNO As String
    Dim strSQL As String, rsRecord As ADODB.Recordset
    Dim strXMLExpend As String, strXMLSub As String
    'strXMLExpend˵��:
    '<TFDATA> //�˷�����
    '  <YCTF>1</YCTF> //�Ƿ��쳣����:1-�쳣����;0-�˷� �˽ڵ����û��
    '  <TFLIST> //�˷��б�
    '    <NO></NO> // �˷ѵ���
    '    <TFITEM> //�˷���
    '      <SerialNum></SerialNum> //���
    '      ��
    '    </TFITEM>
    '  </TFLIST>
    '  ....
    '</TFDATA >
    
    On Error GoTo errHandler
    If lng������� = 0 Then Exit Function
    strXMLExpend = "": strXMLSub = ""
    
    strSQL = "Select /*+cardinality(b,10)*/ Distinct a.NO, a.���" & vbNewLine & _
            " From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
            " Where a.��¼���� = 1 And a.����id = b.����id And b.������� = [1]" & vbNewLine & _
            " Order By a.NO, a.���"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���õ���", lng�������)
    If rsRecord.RecordCount = 0 Then Exit Function
    
    strXMLExpend = strXMLExpend & "<TFDATA>" & vbCrLf '�˷�����
    strXMLExpend = strXMLExpend & "  <YCTF>" & IIf(bln�쳣����, 1, 0) & "</YCTF>" & vbCrLf '�Ƿ��쳣����:1-�쳣����;0-�˷�
    Do While Not rsRecord.EOF
        If NVL(rsRecord!NO) <> strPriorNO Then
            If strPriorNO <> "" Then
                strXMLExpend = strXMLExpend & "    </TFITEM>" & vbCrLf
                strXMLExpend = strXMLExpend & "  </TFLIST>" & vbCrLf
            End If
            strXMLExpend = strXMLExpend & "  <TFLIST>" & vbNewLine '�˷��б�
            strXMLExpend = strXMLExpend & "    <NO>" & NVL(rsRecord!NO) & "</NO>" & vbCrLf '�˷ѵ���
            strXMLExpend = strXMLExpend & "    <TFITEM>" & vbCrLf '�˷���
        End If
        strXMLExpend = strXMLExpend & "      <SerialNum>" & Val(NVL(rsRecord!���)) & "</SerialNum>" & vbCrLf '���
        
        strPriorNO = NVL(rsRecord!NO)
        rsRecord.MoveNext
    Loop
    strXMLExpend = strXMLExpend & "    </TFITEM>" & vbCrLf
    strXMLExpend = strXMLExpend & "  </TFLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & "</TFDATA>"
    
    ZlGetDelXMLExpend = strXMLExpend
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetDelXMLExpendByGrid(ByVal vsfBill As VSFlexGrid) As String
    '�ӽ������л�ȡ�����������˷ѽӿ�zlRetuenCheck��strXMLExpend����ֵ
    Dim i As Integer
    Dim strXMLExpend As String, blnFindSelectItem As Boolean
    Dim strNo As String, strPriorNO As String
    'strXMLExpend˵��:
    '<TFDATA> //�˷�����
    '  <YCTF>1</YCTF> //�Ƿ��쳣����:1-�쳣����;0-�˷� �˽ڵ����û��
    '  <TFLIST> //�˷��б�
    '    <NO></NO> // �˷ѵ���
    '    <TFITEM> //�˷���
    '      <SerialNum></SerialNum> //���
    '      ��
    '    </TFITEM>
    '  </TFLIST>
    '  ....
    '</TFDATA >
    
    On Error GoTo errHandler
    strXMLExpend = "": blnFindSelectItem = False
    
    strXMLExpend = strXMLExpend & "<TFDATA>" & vbCrLf '�˷�����
    strXMLExpend = strXMLExpend & "  <YCTF>" & 0 & "</YCTF>" & vbCrLf '�Ƿ��쳣����:1-�쳣����;0-�˷�
    With vsfBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 Then
                blnFindSelectItem = True
                strNo = .TextMatrix(i, .ColIndex("���ݺ�"))
                If strNo <> strPriorNO Then
                    If strPriorNO <> "" Then
                        strXMLExpend = strXMLExpend & "    </TFITEM>" & vbCrLf
                        strXMLExpend = strXMLExpend & "  </TFLIST>" & vbCrLf
                    End If
                    strXMLExpend = strXMLExpend & "  <TFLIST>" & vbNewLine '�˷��б�
                    strXMLExpend = strXMLExpend & "    <NO>" & strNo & "</NO>" & vbCrLf '�˷ѵ���
                    strXMLExpend = strXMLExpend & "    <TFITEM>" & vbCrLf '�˷���
                End If
                strXMLExpend = strXMLExpend & "      <SerialNum>" & .RowData(i) & "</SerialNum>" & vbCrLf '���
                strPriorNO = strNo
            End If
        Next
    End With
    If blnFindSelectItem = False Then Exit Function
    strXMLExpend = strXMLExpend & "    </TFITEM>" & vbCrLf
    strXMLExpend = strXMLExpend & "  </TFLIST>" & vbCrLf
    strXMLExpend = strXMLExpend & "</TFDATA>"
    
    ZlGetDelXMLExpendByGrid = strXMLExpend
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
