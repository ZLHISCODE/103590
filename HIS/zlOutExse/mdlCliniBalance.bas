Attribute VB_Name = "mdlCliniBalance"
Option Explicit

Public Function GetӦ������㷽ʽ(Optional ByVal str���� As String = "�շ�") As String
    '��ȡӦ������㷽ʽ����֧Ʊ
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSql = _
        "Select b.����" & vbNewLine & _
        "From ���㷽ʽӦ�� A, ���㷽ʽ B" & vbNewLine & _
        "Where b.���� = a.���㷽ʽ And a.���ʽ Is Null And Nvl(b.Ӧ����, 0) = 1" & vbNewLine & _
        "      And a.Ӧ�ó��� = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "mdlCliniBalance", str����)
    If rsTemp.EOF Then Exit Function
    
    GetӦ������㷽ʽ = Nvl(rsTemp!����)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub DeleteBalanceRecord(ByVal lng����ID As Long, _
    ByVal lng��������ID As Long, Optional ByVal lng�����ID As Long, _
    Optional ByVal lng������� As Long, Optional ByVal blnMultiDel As Boolean, _
    Optional cllPro As Collection)
    '�������׵���ʧ��ɾ�������¼
    '��Σ�
    '   blnMultiDel �Ƿ����˿�
    Dim strSql As String
    
    On Error GoTo ErrHandler
    'Zl_���˽����¼_Delete(
    strSql = "Zl_���˽����¼_Delete("
    '  ����id_In     ����Ԥ����¼.����id%Type := Null,
    strSql = strSql & "" & ZVal(lng����ID) & ","
    '  ��������id_In ����Ԥ����¼.��������id%Type := Null,
    strSql = strSql & "" & ZVal(lng��������ID) & ","
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSql = strSql & "" & ZVal(lng�����ID) & ","
    '  �������_In   ����Ԥ����¼.�������%Type := Null,
    strSql = strSql & "" & ZVal(lng�������) & ","
    '  ����˿�_In   Number := 0
    strSql = strSql & "" & IIf(blnMultiDel, 1, 0) & ")"
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "ɾ�������¼"
    Else
        zlAddArray cllPro, strSql
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub CancelBillBalance(ByVal lng����ID As Long, Optional ByVal strNo As String, _
    Optional cllPro As Collection)
    'ȡ�����ݵĽ���
    Dim strSql As String
    
    On Error GoTo ErrHandler
    'Zl_�����շѽ���_Cancel_S(
    strSql = "Zl_�����շѽ���_Cancel_S("
    '  ����id_In   ������ü�¼.����id%Type,
    strSql = strSql & "" & lng����ID & ","
    '  No_In       ������ü�¼.No%Type := Null
    strSql = strSql & "'" & strNo & "')"
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "ȡ�����ݵĽ���"
    Else
        zlAddArray cllPro, strSql
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub CancelBillDelBalance(ByVal lng����ID As Long, Optional ByVal lng����ID As Long, _
    Optional cllPro As Collection)
    'ȡ�����ݵ��˷�
    Dim strSql As String
    
    On Error GoTo ErrHandler
    'Zl_�����˷ѽ���_Cancel_S(
    strSql = "Zl_�����˷ѽ���_Cancel_S("
    '  ����id_In   ������ü�¼.����id%Type,
    strSql = strSql & "" & lng����ID & ","
    '  �ؽ�id_In ������ü�¼.����id%Type := Null
    strSql = strSql & "" & ZVal(lng����ID) & ")"
    
    If cllPro Is Nothing Then
        zlDatabase.ExecuteProcedure strSql, "ȡ�����ݵ��˷�"
    Else
        zlAddArray cllPro, strSql
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function UpdateErrBillOperator(ByVal lng����ID As Long, ByVal lng������� As Long) As Boolean
    '�����쳣�շѣ����²���Ա
    Dim strSql As String
    
    On Error GoTo ErrHandler
    'Zl_�����쳣�շ�_���²���Ա
    strSql = "Zl_�����쳣�շ�_���²���Ա("
    '����id_In     ������ü�¼.����id%Type,
    strSql = strSql & "" & lng����ID & ","
    '����Ա���_In ������ü�¼.����Ա���%Type,
    strSql = strSql & "'" & UserInfo.��� & "',"
    '����Ա����_In ������ü�¼.����Ա����%Type,
    strSql = strSql & "'" & UserInfo.���� & "',"
    '�������_In   ����Ԥ����¼.�������%Type
    strSql = strSql & lng������� & ")"
    zlDatabase.ExecuteProcedure strSql, "�����쳣�շѸ��²���Ա"
    UpdateErrBillOperator = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Public Function Init֧����ʽ(objPayCards As Cards, Optional ByVal blnAddԤ���� As Boolean, _
    Optional ByVal str�ѻ�ҽ�� As String, Optional ByRef str��Ч�ѻ�ҽ�� As String) As Boolean
    '������Ч��֧����ʽ
    '˵����Ԥ����Ľ�������Ϊ-99
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim objCard As Card, objCards As Cards, lngKey As Long
    
    str��Ч�ѻ�ҽ�� = ""
    Set objPayCards = New Cards
    Set objCards = New Cards
    
    Set rsTemp = Get���㷽ʽ("�շ�")
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
        '���:bytType-  0-����ҽ�ƿ�;
        '               1-���õ�ҽ�ƿ�,
        '               2-���д��������˻���������
        '               3-���õ������˻���ҽ�ƿ�
       Set objCards = gobjSquare.objOneCardComLib.zlGetCards(3)
    End If
    
    With rsTemp
        lngKey = 1
        Do While Not .EOF
            blnFind = False
            For Each objCard In objCards
                If objCard.���㷽ʽ = Nvl(!����) Then blnFind = True: Exit For
            Next
            
            '����:1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
            If Not blnFind And InStr("3,4,8", Val(Nvl(!����))) = 0 And Val(Nvl(!Ӧ����)) = 0 Then
                If InStrEx(str�ѻ�ҽ��, Nvl(!����), "|") = False Then
                    Set objCard = New Card
                    objCard.���� = Mid(Nvl(!����), 1, 1)
                    objCard.�ӿڱ��� = Nvl(!����)
                    objCard.�ӿڳ����� = ""
                    objCard.�ӿ���� = -1 * lngKey
                    objCard.���㷽ʽ = Nvl(!����)
                    objCard.���� = Nvl(!����)
                    objCard.���� = True
                    objCard.ȱʡ��־ = Val(Nvl(!ȱʡ)) = 1
                    objCard.֧������ = True
                    objCard.�������� = Val(!����)
                    objPayCards.Add objCard, "K" & lngKey
                    
                    lngKey = lngKey + 1
                Else
                    str��Ч�ѻ�ҽ�� = str��Ч�ѻ�ҽ�� & "|" & Nvl(!����)
                End If
            End If
            .MoveNext
        Loop
    End With
    If str��Ч�ѻ�ҽ�� <> "" Then str��Ч�ѻ�ҽ�� = Mid(str��Ч�ѻ�ҽ��, 2)
    
    '��������,���㷽ʽҪ������"����"Ӧ�ó��ϲ���ʹ��
    For Each objCard In objCards
        rsTemp.Filter = "����='" & objCard.���㷽ʽ & "'"
        If Not rsTemp.EOF Then
            objCard.ȱʡ��־ = Val(Nvl(rsTemp!ȱʡ)) = 1
            objCard.֧������ = True
            objPayCards.Add objCard, "K" & lngKey
            lngKey = lngKey + 1
        End If
    Next
    
    If blnAddԤ���� Then
        '����Ԥ�������
        Set objCard = New Card
        objCard.���� = "Ԥ"
        objCard.�ӿڱ��� = ""
        objCard.�ӿڳ����� = ""
        objCard.�ӿ���� = -1 * lngKey
        objCard.���㷽ʽ = "Ԥ����"
        objCard.���� = "Ԥ����"
        objCard.���� = True
        objCard.ȱʡ��־ = False
        objCard.֧������ = True
        objCard.�������� = "-99"
        objPayCards.Add objCard, "K" & lngKey
    End If
    
    Init֧����ʽ = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteModifyPatiName(ByVal strNos As String, ByVal strName As String) As Boolean
    '�޸Ĳ�����Ϣ
    '��Σ�
    '   strNos ��������ö��ŷָ���A001,A002,...
    '   strName �����µ�����
    Dim cllPro As New Collection
    Dim strSql As String, arrNo As Variant, i As Long
    
    If strNos = "" Then ExecuteModifyPatiName = True: Exit Function
    arrNo = Split(strNos, ",")
    For i = 0 To UBound(arrNo)
        'Zl_���˷��ü�¼_Update_S(
        strSql = "Zl_���˷��ü�¼_Update_S( "
        '  No_In       ������ü�¼.No%Type,
        strSql = strSql & "'" & arrNo(i) & "',"
        '  ��¼����_In ������ü�¼.��¼����%Type,
        strSql = strSql & "" & 1 & ","
        '  ������_In   ������ü�¼.������%Type,
        strSql = strSql & "" & "NULL" & ","
        '  ����ʱ��_In ������ü�¼.����ʱ��%Type,
        strSql = strSql & "" & "NULL" & ","
        '  ����_In     ������ü�¼.����%Type := Null,
        strSql = strSql & "'" & strName & "')"
        '  ��Դ_In     Integer := 1,--����;2-סԺ
        '  ����_In     ������ü�¼.����%Type := Null,
        '  �Ա�_In     ������ü�¼.�Ա�%Type := Null,
        '  ��������_In ������Ϣ.��������%Type := Null
        zlAddArray cllPro, strSql
    Next

    On Error GoTo ErrHandler:
    zlExecuteProcedureArrAy cllPro, "�޸Ĳ�����Ϣ"
    ExecuteModifyPatiName = True
    Exit Function
ErrHandler:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
