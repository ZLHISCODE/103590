Attribute VB_Name = "mdlReadData"
Option Explicit

Public Function GetBaseCode(ByVal varInput As Variant) As ADODB.Recordset
'���ܣ���ȡ�����ֵ������
'������varInput=�����ֵ������б�Index,���ֵ����
'          strDefault=0-����ȡȱʡ���ֵ��<>0��ȡȱʡ���ֵ
'���أ����˺�Ļ����ֵ��¼��

    Dim strTables As String
    Dim strSql As String
    Dim arrTables As Variant
    Dim i As Long
    Dim strFilter As String
    Dim strSort As String
    If gclsPros.BaseCode Is Nothing Then
        '���б��롢���ơ����롢ȱʡ��־�ı�
        'ҽѧ��ʾ�Ƕ�ѡ��ʵ�ֶ�ѡ�Ĺ�������,����֧������ԴΪ��¼�������ͣ�ֻ֧��SQL,�Ժ���������¼���Ķ�ѡ
        '���ҽѧ��ʾ������
        If gclsPros.FuncType = f���ѡ�� Then
            strTables = "���ƽ��;�ֻ��̶�;����������;סԺ����ԭ��"
        ElseIf gclsPros.PatiType = PF_���� Then
            strTables = "ҽ�Ƹ��ʽ;�Ա�;����״��;ְҵ;����;����;Ѫ��;ѧ��;���֤δ¼ԭ��"
        Else
            strTables = "ҽ�Ƹ��ʽ;�Ա�;����״��;ְҵ;����;����;Ѫ��;����ϵ;����;��Ժ��ʽ;�ֻ��̶�;����������;��Ժ��ʽ;��Ⱦ��λ;���ƽ��;������������;סԺ����ԭ��;ѧ��;��Ժת��;���֤δ¼ԭ��;·���ϱ�����ԭ��"
        End If
        arrTables = Split(strTables, ";")
        For i = LBound(arrTables) To UBound(arrTables)
            strSql = strSql & " Union ALL " & vbNewLine & _
                    "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '" & arrTables(i) & "' ���� From " & arrTables(i)
        Next
        strSql = Mid(strSql, Len(" Union ALL " & vbNewLine))
        If gclsPros.FuncType <> f���ѡ�� Then
            '�������������ȫ���б��롢���ơ����롢ȱʡ��־�ı�
            If gclsPros.PatiType <> PF_���� Then
                '1�������б��룬����
                strTables = "�����¼�;��Ⱦ����;�����п�����"
                arrTables = Split(strTables, ";")
                For i = LBound(arrTables) To UBound(arrTables)
                    strSql = strSql & " Union ALL " & vbNewLine & _
                            "Select RowNum As ID, ����,���� ����, ����, 0 ȱʡ, '" & arrTables(i) & "' ���� From " & arrTables(i)
                Next
                
                '2�������б��룬���룬����
                strTables = "�ٴ���������;��ԭѧĿ¼;ҽԺ��ȾĿ¼;��е����Ŀ¼;ICU����;���Ȳ������"
                arrTables = Split(strTables, ";")
                For i = LBound(arrTables) To UBound(arrTables)
                    strSql = strSql & " Union ALL " & vbNewLine & _
                            "Select RowNum As ID, ����, ����, ����,0 ȱʡ, '" & arrTables(i) & "' ���� From " & arrTables(i)
                Next
            End If
            '������ر���Ҫ�жϹ�������ǲ���ϵͳ����
            If gclsPros.PatiType = PF_���� Then
                strTables = "����ȥ��"
                arrTables = Split(strTables, ";")
                For i = LBound(arrTables) To UBound(arrTables)
                    strSql = strSql & " Union ALL " & vbNewLine & _
                            "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '" & arrTables(i) & "' ���� From " & arrTables(i)
                Next
            Else
                strTables = "סԺ�����ڼ�"
                arrTables = Split(strTables, ";")
                For i = LBound(arrTables) To UBound(arrTables)
                    strSql = strSql & " Union ALL " & vbNewLine & _
                            "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '" & arrTables(i) & "' ���� From " & arrTables(i)
                Next
            End If
        Else
            '�������������ȫ���б��롢���ơ����롢ȱʡ��־�ı�
            '1�������б��룬����
            strTables = "�����¼�;��Ⱦ����"
            arrTables = Split(strTables, ";")
            For i = LBound(arrTables) To UBound(arrTables)
                strSql = strSql & " Union ALL " & vbNewLine & _
                        "Select RowNum As ID, ����,���� ����, ����, 0 ȱʡ, '" & arrTables(i) & "' ���� From " & arrTables(i)
            Next
        End If
        Set gclsPros.BaseCode = zlDatabase.OpenSQLRecord(strSql, "��ҳ��ȡ�����ֵ�")
    End If
    
    gclsPros.BaseCode.Filter = 0
    If gclsPros.BaseCode.RecordCount > 0 Then
        strSort = "����,ID" '�����Զ����α��ƶ�������
        If TypeName(varInput) <> "String" Then
            Select Case varInput
                Case BCC_���ʽ
                    strFilter = "����='ҽ�Ƹ��ʽ'"
                Case BCC_�Ա�
                    strFilter = "����='�Ա�'"
                Case BCC_����
                    strFilter = "����='����״��'"
                Case BCC_ְҵ
                    strFilter = "����='ְҵ'"
                Case BCC_����
                    strFilter = "����='����'"
                Case BCC_����
                    strFilter = "����='����'"
                Case BCC_Ѫ��
                    strFilter = "����='Ѫ��'"
                Case BCC_��ϵ
                    strFilter = "����='����ϵ'"
                Case BCC_��Ժ���
                    strFilter = "����='����'"
                Case BCC_��������
                    strFilter = "����='�ٴ���������'"
                Case BCC_��Ժ;��
                    strFilter = "����='��Ժ��ʽ'"
                Case BCC_�ֻ��̶�
                    strFilter = "����='�ֻ��̶�'"
                Case BCC_����������
                    strFilter = "����='����������'"
                Case BCC_��Ժ��ʽ
                    strFilter = "����='��Ժ��ʽ'"
                Case BCC_�����ڼ�
                    strFilter = "����='סԺ�����ڼ�'"
                Case BCC_ȥ��
                    strFilter = "����='����ȥ��'"
                Case BCC_�Ļ��̶�
                    strFilter = "����='ѧ��'"
                    strSort = "���� Desc,ID"
                Case BCC_���֤
                    strFilter = "����='���֤δ¼ԭ��'"
                Case BCC_����ԭ��
                    strFilter = "����='·���ϱ�����ԭ��'"
            End Select
        Else
            strFilter = "����='" & varInput & " '"
        End If
        gclsPros.BaseCode.Filter = strFilter
        gclsPros.BaseCode.Sort = strSort '�����Զ����α��ƶ�������
        Set GetBaseCode = gclsPros.BaseCode
    End If
    Exit Function
errH:
    If ErrCenter() <> 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetManData(Optional ByVal varManPros As Variant, Optional ByVal pfManFrom As PatiFrom = 0, Optional ByVal slSignLevel As SignLevel = -1) As ADODB.Recordset
'���ܣ���ȡ��Ա��Ϣ
'������varManPros=��Ա�����б����������Ա���ʣ���Ա���ʣ�ҽ��,��ʿ,��������Ա
'          pfManFrom=��Ա��Դ��0����������Դ��1-����=1��2-סԺ=1
'          slSignLevel=4��ǩ���ļ�����룬�ֱ��Ӧһ����רҵ����ְ�������ְ��
    Dim strManPros As String, strFilter As String
    Dim strSql As String, strSQLTmp As String
    Dim bln����ҽ�� As Boolean, blnAdd As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim arrFileds As Variant
    Dim int���� As Integer, int���� As Integer
    
    On Error GoTo errH
    
    If TypeName(varManPros) = "String" Then
        strManPros = varManPros
    Else
    '��Ա�����б�����
        pfManFrom = Decode(varManPros, MC_����ҽʦ, PF_����, PF_סԺ)
        strManPros = Decode(varManPros, MC_���λ�ʿ, "��ʿ", MC_�ʿػ�ʿ, "��ʿ", MC_��ĿԱ, "��������Ա", "ҽ��")
        slSignLevel = Decode(varManPros, MC_������, SL_������, MC_���λ�����, SL_����ҽʦ, MC_����ҽʦ, SL_����ҽʦ, MC_סԺҽʦ, SL_סԺҽʦ, -1)
    End If

    strFilter = strManPros & "=1"
    If pfManFrom <> 0 Then
        strFilter = strFilter & " And " & IIf(pfManFrom = PF_����, "����=1", "סԺ=1")
    End If
    If slSignLevel >= SL_������ Then
        strFilter = strFilter & IIf(slSignLevel = SL_������, " And ����", " And ����") & ">=" & Decode(slSignLevel, SL_������, 1, SL_����ҽʦ, 4, SL_����ҽʦ, 3, SL_סԺҽʦ, 1)
    End If
    '�ж��Ƿ��Ѿ�����
    If strManPros <> "ҽ��" Then
        blnAdd = InStr(gclsPros.LoadMans, strManPros) = 0
        If blnAdd Then
            gclsPros.LoadMans = gclsPros.LoadMans & "|" & strManPros
        End If
    Else
        If InStr(gclsPros.LoadMans, "ҽ��" & pfManFrom) = 0 Then
            blnAdd = True
            gclsPros.LoadMans = gclsPros.LoadMans & "|" & "ҽ��" & pfManFrom
        End If
    End If
    'û�й��˵����ݻ�û�������ݣ����ȡ���ݿ�
    If blnAdd Or gclsPros.ManInfo Is Nothing Then
        '��װSQL
        bln����ҽ�� = pfManFrom = PF_���� And strManPros = "ҽ��"
        If gclsPros.FuncType <> f������ҳ And Not bln����ҽ�� And pfManFrom <> 0 Then
            If strManPros <> "��ʿ" Then
                strSQLTmp = "And" & vbNewLine & _
                        "     C.����id In (Select ����id From ������Ա Where ��Աid = [2])"
            Else
                '��ʿ�����Ǻ�ҽ���������ң����Ұ������������µĲ���
                strSQLTmp = "And" & vbNewLine & _
                        "     C.����id In (Select ����id From ������Ա Where ��Աid = [2]" & vbNewLine & _
                                                "Union" & vbNewLine & _
                                                "Select Distinct B.����ID From ������Ա a, �������Ҷ�Ӧ b Where A.����id = B.����ID And A.��Աid = [2])"
            End If
        End If
        If gclsPros.FuncType = f������ҳ Then
            strSql = "Select Distinct A.ID, A.��� ����, A.���� ����, A.����,zlwbcode(����) ��ʼ���, A.����ְ��, A.רҵ����ְ��,0 ����,0 ����, 0 ҽ��, 0 ��ʿ, 0 ��������Ա, 0 ����,0 סԺ,0 ȱʡ" & vbNewLine & _
                        "From ��Ա�� A, ��Ա����˵�� B" & vbNewLine & _
                        "Where A.Id = B.��Աid And B.��Ա���� = [1]  And" & vbNewLine & _
                        "      (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & IIf(gstrNodeNo <> "-", " And (A.վ�� = '" & gstrNodeNo & "' Or A.վ�� Is Null)", "")
        Else
            strSql = "Select Distinct A.ID, A.��� ����, A.���� ����, A.����,zlwbcode(����) ��ʼ���, A.����ְ��, A.רҵ����ְ��,0 ����,0 ����, 0 ҽ��, 0 ��ʿ, 0 ��������Ա, 0 ����,0 סԺ,0 ȱʡ" & vbNewLine & _
                        "From ��Ա�� A, ��Ա����˵�� B ,������Ա C, ��������˵�� D" & vbNewLine & _
                        "Where A.Id = B.��Աid And B.��Ա���� = [1] And A.Id = C.��Աid And C.����id = D.����id And D.�������  In (" & IIf(pfManFrom = 0, "1,2", pfManFrom) & ",3) And" & vbNewLine & _
                        "      (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & IIf(gstrNodeNo <> "-", " And (A.վ�� = '" & gstrNodeNo & "' Or A.վ�� Is Null)", "") & strSQLTmp
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ҳ��ȡ��Ա��Ϣ", strManPros, UserInfo.ID)
        If gclsPros.ManInfo Is Nothing Then
            Set gclsPros.ManInfo = zlDatabase.CopyNewRec(rsTmp, True, "ID,����,����,����,��ʼ���,����,����,ҽ��,��ʿ,��������Ա,����,סԺ,ȱʡ")
        End If
        arrFileds = Array("ID", "����", "����", "����", "��ʼ���", "����", "����")
        With rsTmp
            .Sort = "Id": gclsPros.ManInfo.Filter = "": gclsPros.ManInfo.Sort = "ID"
            Do While Not .EOF
                gclsPros.ManInfo.Filter = "ID=" & !ID
                If gclsPros.ManInfo.EOF Then
                    int���� = Decode(!����ְ�� & "", "��������", 2, "���Ҹ�����", 1, 0)
                    int���� = Decode(!רҵ����ְ�� & "", "����ҽʦ", 5, "������ҽʦ", 4, "����ҽʦ", 3, "ҽʦ", 2, "ҽʿ", 1, 0)
                    gclsPros.ManInfo.AddNew arrFileds, Array(!ID, !����, !����, !����, !��ʼ���, int����, int����)
                End If
                Select Case strManPros
                    Case "��ʿ"
                        gclsPros.ManInfo!��ʿ = 1
                    Case "ҽ��"
                        gclsPros.ManInfo!ҽ�� = 1
                    Case "��������Ա"
                        gclsPros.ManInfo!��������Ա = 1
                End Select
                Select Case pfManFrom
                    Case PF_����
                        gclsPros.ManInfo!���� = 1
                    Case PF_סԺ
                        gclsPros.ManInfo!סԺ = 1
                End Select
                Call gclsPros.ManInfo.Update
                rsTmp.MoveNext
            Loop
        End With
    End If
    gclsPros.ManInfo.Filter = strFilter
    gclsPros.ManInfo.Sort = "����" '�����Զ����α��ƶ�������
    Set GetManData = gclsPros.ManInfo
    Exit Function
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SetCboFromName(ByVal strName As String, objCbo As Object, Optional ByVal strType As String, Optional ByVal blnAdd As Boolean) As Boolean
'���ܣ���ָ����������Ա���뵽��������
'blnAdd=ǿ������
    Static rsTmp As ADODB.Recordset
    Dim strSql As String, intIdx As Integer
    
    On Error GoTo errH
    If strType = "��Ա" Then
        If rsTmp Is Nothing Then
            strSql = "Select A.ID,A.��� ����,A.���� ����,Null As ����" & _
                " From ��Ա�� A,��Ա����˵�� B" & _
                " Where A.ID=B.��ԱID And B.��Ա���� IN(" & IIf(gclsPros.FuncType = f������ҳ, "'ҽ��','��ʿ','��������Ա'", "'ҽ��','��ʿ'") & ")" & _
                " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.����"
            Set rsTmp = New ADODB.Recordset
            Call zlDatabase.OpenRecordset(rsTmp, strSql, "SetCboFromName")
        End If
        
        rsTmp.Filter = "����='" & strName & "'"
        If Not rsTmp.EOF Then
            intIdx = objCbo.ListCount
            If objCbo.ListCount > 0 Then
                If objCbo.ItemData(objCbo.ListCount - 1) = -1 Then
                    intIdx = objCbo.ListCount - 1
                End If
            End If
            
            If IsNull(rsTmp!����) Then
                objCbo.AddItem rsTmp!����, intIdx
            Else
                objCbo.AddItem rsTmp!���� & "-" & Chr(13) & rsTmp!����
            End If
            objCbo.ItemData(objCbo.NewIndex) = Val(rsTmp!ID)
            Call zlControl.CboSetIndex(objCbo.hwnd, objCbo.NewIndex)
        ElseIf gclsPros.FuncType = f������ҳ And blnAdd Then '������ҳֱ����������Ա
            objCbo.AddItem strName
            objCbo.ListIndex = objCbo.NewIndex
            objCbo.ItemData(objCbo.NewIndex) = -999
        End If
    ElseIf blnAdd Then
        objCbo.AddItem strName
        objCbo.ListIndex = objCbo.NewIndex
    End If
    
    SetCboFromName = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDeptData() As ADODB.Recordset
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select Distinct A.Id, A.����, A.����, A.����, A.λ��,B.��������" & vbNewLine & _
            "From ���ű� A, ��������˵�� B" & vbNewLine & _
            "Where A.Id = B.����id And (B.������� In (2, 3) And B.�������� In ('�ٴ�', '����') OR B.��������='ICU') And" & vbNewLine & _
            "      (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & vbNewLine & _
             IIf(gstrNodeNo <> "-", " And (A.վ�� = '" & gstrNodeNo & "' Or A.վ�� Is Null)", "") & vbNewLine & _
            "Order By A.����"

    Set GetDeptData = zlDatabase.OpenSQLRecord(strSql, "��ȡ�ٴ�����")
    Exit Function
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetAllerData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal blnMoved As Boolean) As ADODB.Recordset
'���ܣ���ȡ��������
'������intType=0-������ҳ ��1-סԺ��ҳ�벡����ҳ
'���أ����˹�������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = " Select Distinct a.Id, a.��¼��Դ, a.����ʱ��, a.ҩ��id, a.ҩ����, a.������Ӧ, a.����Դ����, a.��¼ʱ��" & vbNewLine & _
             " From ���˹�����¼ A" & vbNewLine & _
             " Where a.��� = 1 And a.����id = [1] And a.��ҳid = [2] And a.��¼��Դ " & IIf(gclsPros.FuncType <> f������ҳ, " = 3 ", " in (3,4) ") & vbNewLine & _
             " Order By Nvl(a.����ʱ��, a.��¼ʱ��) Desc, a.ҩ����"

    If blnMoved Then
        strSql = Replace(strSql, "���˹�����¼", "H���˹�����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ҳ��ȡ������Ϣ", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        If gclsPros.FuncType = f������ҳ Then
            rsTmp.Filter = "��¼��Դ=4"
            If rsTmp.EOF Then
                rsTmp.Filter = "��¼��Դ=3"
            End If
        End If
    End If
    Set GetAllerData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiMainInfoData(ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal str�Һŵ� As String) As ADODB.Recordset
'���ܣ���ȡ������ҳ��Ϣ�Լ�������Ϣ����
'������lng����ID=����ID
'      lng��ҳID=סԺ���˲Ŵ�
'      str�Һŵ�=���ﲡ�˲Ŵ�
'       blnMove=�Ƿ��ת�������ж�ȡ��������ʷ���ж�ȡ
'���أ�������ҳ��Ϣ������Ϣ
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '������ҳ��ȡ���ξ�����Ϣ�Լ�������Ϣ,��Ⱦ���ϴ����ȸ����ȡ�����ж�
    If str�Һŵ� <> "" Then
        gclsPros.Moved = zlDatabase.NOMoved("���˹Һż�¼", str�Һŵ�)
        
        strSql = "Select B.Id As �Һ�id, A.����id, A.�����, A.ҽ�Ƹ��ʽ, A.��������, A.�����ص�, A.���֤��, A.����֤��, A.ְҵ, A.����, A.����, A.����, A.����, A.ѧ��, A.����״��," & vbNewLine & _
                    "       A.��ͥ��ַ, A.��ͥ�绰, A.��ͥ��ַ�ʱ�, A.�໤��, A.���ڵ�ַ, A.���ڵ�ַ�ʱ�, A.��ͬ��λid, A.������λ ��λ��ַ, A.��λ�绰, A.��λ�ʱ�, Nvl(A.����, 0) ����," & vbNewLine & _
                    "       Nvl(B.����, A.����) ����, Nvl(B.�Ա�, A.�Ա�) �Ա�, Nvl(B.����, A.����) ����, B.����ʱ��, B.������ַ, B.��Ⱦ���ϴ�, B.����," & vbNewLine & _
                    "       Nvl(Nvl(B.�������id, Decode(B.ת��״̬, 1, B.ת�����id, Null)), B.ִ�в���id) As ����id, B.ժҪ, B.����, C.������" & vbNewLine & _
                    "From ������Ϣ A, ���˹Һż�¼ B, ����������Ϣ C" & vbNewLine & _
                    "Where A.����id = B.����id And B.����id = C.����id(+) And B.���� = C.����(+) And " & IIf(str�Һŵ� = "NULL", "B.ID", "B.No") & "= [1] And B.��¼���� = 1 And B.��¼״̬ = 1"
         If gclsPros.Moved Then
            strSql = Replace(strSql, "���˹Һż�¼", "H���˹Һż�¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������Ϣ�Լ����ξ�����Ϣ", IIf(str�Һŵ� = "NULL", lng��ҳID, str�Һŵ�))
    '������ҳ��סԺ��ҳ��ȡ������Ϣ�Լ�������ҳ��Ϣ
    Else
        If gclsPros.FuncType <> f������ҳ Then
            strSql = "Select a.����id, a.��ҳid, Nvl(a.����, d.����) As ����, Nvl(a.�Ա�, d.�Ա�) As �Ա�, Nvl(a.����, d.����) As ����, a.���, a.����, a.Ѫ��, a.ְҵ, a.����," & vbNewLine & _
                            "       a.����, a.ѧ��, a.����״��, Nvl(a.��ϵ������, d.��ϵ������) ��ϵ������, Nvl(a.��ϵ�˹�ϵ, d.��ϵ�˹�ϵ) ��ϵ�˹�ϵ, Nvl(a.��ϵ�˵�ַ, d.��ϵ�˵�ַ) ��ϵ�˵�ַ," & vbNewLine & _
                            "       Nvl(a.��ϵ�˵绰, d.��ϵ�˵绰) ��ϵ�˵绰, Nvl(a.���ڵ�ַ, d.���ڵ�ַ) ���ڵ�ַ, Nvl(a.���ڵ�ַ�ʱ�, d.���ڵ�ַ�ʱ�) ���ڵ�ַ�ʱ�, Nvl(a.��ͥ��ַ, d.��ͥ��ַ) ��ͥ��ַ," & vbNewLine & _
                            "       Nvl(a.��ͥ�绰, d.��ͥ�绰) ��ͥ�绰, Nvl(a.��ͥ��ַ�ʱ�, d.��ͥ��ַ�ʱ�) ��ͥ��ַ�ʱ�, Nvl(a.��λ��ַ, d.������λ) ��λ��ַ, Nvl(a.��λ�绰, d.��λ�绰) ��λ�绰," & vbNewLine & _
                            "       Nvl(a.��λ�ʱ�, d.��λ�ʱ�) ��λ�ʱ�, a.סԺ��, a.��������, a.����Ժ, a.��Ժ����id, a.��Ժ����id, a.��Ժ����, a.��Ժ����, a.��Ժ��ʽ, a.��Ժ����, a.��Ժ����id," & vbNewLine & _
                            "       a.��Ժ����, a.��Ժ����, a.��Ժ��ʽ, a.�Ƿ�ȷ��, a.ȷ������, a.�·�����, a.���ȴ���, a.�ɹ�����, a.�����־, a.��������, a.ʬ���־, a.����ҽʦ, a.���λ�ʿ, a.סԺҽʦ," & vbNewLine & _
                            "       a.��ĿԱ���, a.��ĿԱ����, a.��Ŀ����, a.���ú�, a.��ҽ�������, a.������, a.�ѱ�, a.ҽ�Ƹ��ʽ, a.��ǰ����id, a.����, a.״̬, b.���� As ��Ժ����," & vbNewLine & _
                            "       c.���� As ��Ժ����, c.���� As ��Ժ���ұ���, d.��������, d.�����ص�, d.���, d.����, d.����, d.Email, d.Qq, d.��ͬ��λid, d.סԺ����, d.��ǰ����id," & vbNewLine & _
                            "       d.��Ժʱ��, d.��Ժʱ��, d.ҽ����, d.���֤��, d.����֤��, d.������,a.����ת�� " & vbNewLine & _
                            "From ������ҳ a, ���ű� b, ���ű� c, ������Ϣ d" & vbNewLine & _
                            "Where A.��Ժ����id = B.Id(+) And A.��Ժ����id = C.Id(+) And A.����id = D.����id And A.����id = [1] And A.��ҳid = [2]"
        Else
            strSql = "Select D.����id, [2] ��ҳid, Nvl(A.����, D.����) As ����, Nvl(A.�Ա�, D.�Ա�) As �Ա�, Nvl(A.����, D.����) As ����, a.��������, Nvl(A.ְҵ, D.ְҵ) ְҵ," & vbNewLine & _
                            "       Nvl(A.����, D.����) ����, Nvl(A.����, D.����) ����, Nvl(A.����״��, D.����״��) ����״��, Nvl(A.��ͥ��ַ, D.��ͥ��ַ) ��ͥ��ַ," & vbNewLine & _
                            "       Nvl(A.��ͥ�绰, D.��ͥ�绰) ��ͥ�绰, Nvl(A.��ͥ��ַ�ʱ�, D.��ͥ��ַ�ʱ�) ��ͥ��ַ�ʱ�, Nvl(A.��ϵ������, D.��ϵ������) ��ϵ������," & vbNewLine & _
                            "       Nvl(A.��ϵ�˹�ϵ, D.��ϵ�˹�ϵ) ��ϵ�˹�ϵ, Nvl(A.��ϵ�˵�ַ, D.��ϵ�˵�ַ) ��ϵ�˵�ַ, Nvl(A.��ϵ�˵绰, D.��ϵ�˵绰) ��ϵ�˵绰, Nvl(A.���ڵ�ַ, D.���ڵ�ַ) ���ڵ�ַ," & vbNewLine & _
                            "       Nvl(A.���ڵ�ַ�ʱ�, D.���ڵ�ַ�ʱ�) ���ڵ�ַ�ʱ�, Nvl(A.��λ�绰, D.��λ�绰) ��λ�绰, Nvl(A.��λ�ʱ�, D.��λ�ʱ�) ��λ�ʱ�, A.����Ժ, A.��Ժ����id, A.��Ժ����id," & vbNewLine & _
                            "       A.��Ժ����, A.��Ժ����, A.��Ժ��ʽ, A.��Ժ����, A.��Ժ����id, A.��Ժ����, A.��Ժ����, A.��Ժ��ʽ, A.�Ƿ�ȷ��, A.ȷ������, A.�·�����, A.Ѫ��, A.���ȴ���," & vbNewLine & _
                            "       A.�ɹ�����, A.�����־, A.��������, A.ʬ���־, A.����ҽʦ, A.���λ�ʿ, A.סԺҽʦ, A.��ĿԱ���, A.��ĿԱ����, A.��Ŀ����, A.���ú�, A.���, A.����," & vbNewLine & _
                            "       Nvl(A.��λ��ַ, D.������λ) ��λ��ַ, A.��ҽ�������, A.״̬, B.���� As ��Ժ����, C.���� As ��Ժ����, C.���� As ��Ժ���ұ���, D.��������, D.�����ص�, D.���֤��," & vbNewLine & _
                            "       D.����֤��, D.����, D.����, D.Email, D.Qq, D.��ͬ��λid, D.סԺ����, D.��ǰ����id, D.��Ժʱ��, D.��Ժʱ��, D.������, E.������," & vbNewLine & _
                            "       Nvl(E.������, A.������) As ������, Nvl(G.ҽ����, D.ҽ����) As ҽ����, Nvl(A.סԺ��, D.סԺ��) סԺ��, A.�ѱ�, A.ҽ�Ƹ��ʽ, A.��ǰ����id," & vbNewLine & _
                            "       Nvl(A.����, D.����) ����, F.������ As ��󲡰���, F.������ ��󵵰���, H.���� �����ұ��� ,a.����ת�� " & vbNewLine & _
                            "From ������ҳ A, ���ű� B, ���ű� C, ������Ϣ D, סԺ������¼ E," & vbNewLine & _
                            "     (Select N.����id, N.������, N.������, M.��Ժ����id" & vbNewLine & _
                            "       From ������ҳ M, סԺ������¼ N" & vbNewLine & _
                            "       Where N.����id = M.����id(+) And N.��ҳid = M.��ҳid(+) And N.����id = [1] And" & vbNewLine & _
                            "             N.��ҳid = (Select Max(��ҳid) As ��ҳid From סԺ������¼ Where ����id = [1])) F, �����ʻ� G, ���ű� H" & vbNewLine & _
                            "Where A.��Ժ����id = B.Id(+) And A.��Ժ����id = C.Id(+) And A.����id(+) = D.����id And A.����id = E.����id(+) And A.��ҳid = E.��ҳid(+) And" & vbNewLine & _
                            "      D.����id = F.����id(+) And A.����id = G.����id(+) And A.���� = G.����(+) And F.��Ժ����id = H.Id(+) And D.����id = [1] And" & vbNewLine & _
                            "      A.��ҳid(+) = [2]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������ҳ�벡����Ϣ", lng����ID, lng��ҳID)
        
        If rsTmp.RecordCount > 0 Then
            gclsPros.Moved = Val(NVL(rsTmp!����ת��)) <> 0
        End If
    End If
    
    Set GetPatiMainInfoData = rsTmp
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetPatiAuxiInfoData(ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal str�Һŵ� As String, Optional ByVal bytModel As Byte = 1) As ADODB.Recordset
'���ܣ���ȡ������ҳ��Ϣ�ӱ������Ϣ�ӱ�����
'������lng����ID=����ID
'      lng��ҳID=סԺ���˲Ŵ�
'      str�Һŵ�=���ﲡ�˲Ŵ�
'      bytModel =1 ������ҳ,=2������������������Ǽ�
'���أ�������ҳ��Ϣ�ӱ������Ϣ�ӱ�����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strDelicery As String
    On Error GoTo errH
    If bytModel = 1 Then
        If str�Һŵ� <> "" Then
            strSql = "Select Upper(��Ϣ��) ��Ϣ��, ��Ϣֵ,Null ����" & vbNewLine & _
                    "From ������Ϣ�ӱ�" & vbNewLine & _
                    "Where ����id = [1] And (����id = [2] Or ����id Is Null)" & vbNewLine & _
                    "Order By Nvl(����id, 999999999)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������Ϣ�ӱ�", lng����ID, lng��ҳID)
        Else
            strSql = "Select Decode(B.����,Null, Upper(A.��Ϣ��), A.��Ϣ��) ��Ϣ��, A.��Ϣֵ, B.����" & vbNewLine & _
                    "From ������ҳ�ӱ� A, ������Ŀ B" & vbNewLine & _
                    "Where A.��Ϣ�� = B.����(+) And A.����id = [1] And A.��ҳid = [2]" & vbNewLine & _
                    "Order By A.��Ϣ��"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������ҳ�ӱ�", lng����ID, lng��ҳID)
        End If
    Else
        strDelicery = "����ʱ��, �������, ̥��, ̥��, ����ʱ��1, ����ʱ��2, ����ʱ��3,�ܲ���ʱ��,�����Ѫ��,���Ʋ���֢,�����������"
        strSql = "Select A.��Ϣ��, A.��Ϣֵ, 0 as ���� " & vbNewLine & _
                "From ������ҳ�ӱ� A" & vbNewLine & _
                "Where A.����id = [1] And A.��ҳid = [2] And  Instr([3], ��Ϣ��) > 0" & vbNewLine & _
                "Order By A.��Ϣ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������ҳ�ӱ�", lng����ID, lng��ҳID, strDelicery)
    End If
    Set GetPatiAuxiInfoData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiDiagData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intType As Integer, Optional ByVal blnLast As Boolean, Optional ByVal bln��Ŀ As Boolean, Optional ByVal blnMoved As Boolean) As ADODB.Recordset
'���ܣ���ȡ���������Ϣ
'������intType=0-������ҳ ��1-סԺ��ҳ�벡����ҳ
'      blnLast=True-��ȡ���ξ������ϣ�False=��ȡ���һ�ξ�������(�ò���ֻ��������ҳ��Ч��
'���أ����������Ϣ
    Dim strSql As String, strSQLTmp As String, strDiagType As String
    Dim int��¼��Դ As Integer
    Dim strSQLJudge As String '�����ж�סԺ��ҳ�Ƿ������
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    'Ĭ��ֱ�Ӵ�����ҳID����
    strSQLTmp = "[2]"
    If intType = 0 Then
        If blnLast Then
        '���һ�εľ���ID
            strSQLTmp = "(Select Max(ID) As ��ҳid" & vbNewLine & _
                        "From ���˹Һż�¼" & vbNewLine & _
                        "Where ����id = [1] And ��¼���� = 1 And ��¼״̬ = 1 And" & vbNewLine & _
                        "      �Ǽ�ʱ�� =" & vbNewLine & _
                        "      (Select Max(A.�Ǽ�ʱ��)" & vbNewLine & _
                        "       From ���˹Һż�¼ A" & vbNewLine & _
                        "       Where A.����id = [1] And A.��¼���� = 1 And A.��¼״̬ = 1 And A.�Ǽ�ʱ�� < (Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID = [2])))"
        End If
        '���ö�ȡ��ϵ�����Լ������Դ
        If gclsPros.Have��ҽ Then
            strDiagType = " And A.��¼��Դ IN(1,3) And A.������� IN(1,11) "
        Else
            strDiagType = " And A.��¼��Դ IN(1,3) And A.�������=1 "
        End If
    Else
        '�ж��Ƿ�����ҳ��Դ�򲡰���Դ����ϡ�
        If gclsPros.FuncType <> f������ҳ Then
            int��¼��Դ = 3
            strSQLJudge = "Select 1 From ������ϼ�¼ Where ����id = [1] And ��ҳid =[2] And ��¼��Դ = [3] And Rownum < 2"
            If blnMoved Then
                 strSQLJudge = Replace(strSQLJudge, "������ϼ�¼", "H������ϼ�¼")
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQLJudge, "��ҳ��Դ����ж�", lng����ID, lng��ҳID, int��¼��Դ)
            If rsTmp.RecordCount > 0 Then
                strDiagType = " And A.��¼��Դ =[3] "
            Else
                strDiagType = " And A.��¼��Դ IN(1,2,4) "
            End If
        Else
            int��¼��Դ = 4
            If Not bln��Ŀ Then
                strSQLJudge = "Select Nvl(Max(Nvl(��¼��Դ, 0)), 0) ��¼��Դ" & vbNewLine & _
                                    "From ������ϼ�¼" & vbNewLine & _
                                    "Where ����id = [1] And ��ҳid = [2] And Nvl(��¼��Դ, 0) <= 4"
                If blnMoved Then
                    strSQLJudge = Replace(strSQLJudge, "������ϼ�¼", "H������ϼ�¼")
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQLJudge, "��ҳ��Դ����ж�", lng����ID, lng��ҳID, int��¼��Դ)
                If rsTmp.RecordCount > 0 Then
                    int��¼��Դ = Val(rsTmp!��¼��Դ & "")
                End If
            End If
            strDiagType = " And A.��¼��Դ =[3] "
        End If
        
        '���ö�ȡ��ϵ����
        If gclsPros.Have��ҽ Then
            strDiagType = strDiagType & " And A.������� IN(1,2,3,5,6,7,10,11,12,13,21) "
        Else
            strDiagType = strDiagType & " And A.������� IN(1,2,3,5,6,7,10,21) "
        End If
    End If
    If gclsPros.FuncType <> f������ҳ Then
        '��װSQL,���Ӳ������Ĳ��ò�ѯҽ����¼
        strSql = "Select A.��ע, A.Id, A.����id, A.��ҳid, A.ҽ��id, A.��¼��Դ, A.��ϴ���, Nvl(A.�������,1) �������, A.�������, A.��Ժ����, A.����id, A.���id, A.֤��id,B.���� ��������,C.���� �������,D.���� ֤������," & vbNewLine & _
                "       A.�������, A.��Ժ���, A.�Ƿ�δ��, A.�Ƿ�����, A.����ʱ��, B.���� As ��������,B.��� As �������, B.����, C.���� As ��ϱ���, D.���� As ֤�����," & vbNewLine & _
                IIf(gclsPros.FuncType = f���Ӳ���, " Null ҽ��id", " (Select F_List2str(Cast(Collect(C.ҽ��id || '') As T_Strlist)) ҽ��id" & vbNewLine & _
                "         From �������ҽ�� C,����ҽ����¼ F " & vbNewLine & _
                "         Where C.ҽ��ID = F.ID and C.���id = A.Id and nvl(F.�������,0) = 0) As ҽ��id") & ",B.�Ա�����, B.��Ч����, B.����, B.����, E.Id As ����, E.�Ƿ���,Null ����ID,A.��¼����,A.��¼�� " & vbNewLine & _
                "From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C, ��������Ŀ¼ D,����������� E" & vbNewLine & _
                "Where A.����id = B.Id(+) And A.���id = C.Id(+) And A.֤��id = D.Id(+)  And  B.����id = E.Id(+)" & strDiagType & "And A.ȡ��ʱ�� Is Null And A.������� Is Not Null And ����id = [1] And ��ҳid =" & strSQLTmp & vbNewLine & _
                "Order By A.�������, A.��¼��Դ Desc, A.��ϴ���, Nvl(A.�������,1), A.Id"
    Else
        If bln��Ŀ Then
            strSql = "Select A.��ע, A.Id, A.����id, A.��ҳid, A.ҽ��id, A.��¼��Դ, A.��ϴ���, Decode(Nvl(A.�������, 0), 0, 1, A.�������) �������, A.�������, A.��Ժ����, A.����id, A.���id, A.֤��id," & vbNewLine & _
                    "       B.���� ��������, Null �������, D.���� ֤������, A.�������, A.��Ժ���, A.�Ƿ�δ��, A.�Ƿ�����, A.����ʱ��," & vbNewLine & _
                    "       B.���� As ��������,B.��� As ������� ,  Null ��ϱ���, D.���� As ֤�����, Null ҽ��id, B.�Ա�����, B.��Ч����, B.����, B.����, C.Id As ����, C.�Ƿ���,NULL ����ID,A.��¼����,A.��¼�� " & vbNewLine & _
                    "From ������ϼ�¼ A, ��������Ŀ¼ B, ����������� C, ��������Ŀ¼ D " & vbNewLine & _
                    "Where A.����id = B.Id(+) And A.֤��id = D.Id(+) And A.����id = [1] And A.��ҳid = [2] " & strDiagType & " And B.����id = C.Id(+)  " & vbNewLine & _
                    "Order By A.�������, A.��ϴ���, Decode(Nvl(A.�������, 0), 0, 1, A.�������)"
        Else
            strSql = "Select a.��ע, a.Id, a.����id, a.��ҳid, a.��¼��Դ, Row_Number() Over(Partition By ������� Order By ��ϴ���) As ��ϴ���," & vbNewLine & _
                            "       Decode(Nvl(a.�������, 0), 0, 1, a.�������) �������, a.�������, a.��Ժ����, a.����id, a.���id, a.֤��id, a.��������, Null �������, a.֤������," & vbNewLine & _
                            "       a.�������, a.��Ժ���, a.�Ƿ�δ��, a.�Ƿ�����, a.����ʱ��, a.��������, a.�������, Null ��ϱ���, a.֤�����, Null ҽ��id, a.�Ա�����, a.��Ч����, a.����, a.����," & vbNewLine & _
                            "       a.����, a.�Ƿ���, a.����id, a.��¼����, a.��¼��" & vbNewLine & _
                            "From (Select Distinct a.Id, a.����id, a.��ҳid, a.��¼��Դ, Nvl(a.��ϴ���, 1) As ��ϴ���, Decode(Nvl(a.�������, 0), 0, 1, a.�������) �������," & vbNewLine & _
                            "                       a.�������, a.����id, a.���id, a.֤��id, '' || a.������� As �������, a.��Ժ����, a.��Ժ���, a.�Ƿ�δ��, a.�Ƿ�����, a.����ʱ��, a.��ע," & vbNewLine & _
                            "                       b.���� As ��������, b.���� ��������, b.��� As �������, b.�Ա�����, b.��Ч����, b.����, b.����, c.Id As ����, c.�Ƿ���, d.���� As ֤�����," & vbNewLine & _
                            "                       d.���� As ֤������, NULL ����id, a.��¼����, a.��¼��" & vbNewLine & _
                            "       From ������ϼ�¼ a, ��������Ŀ¼ b, ����������� c, ��������Ŀ¼ d " & vbNewLine & _
                            "       Where a.����id = b.Id(+) And a.֤��id = d.Id(+) And a.����id = [1] And a.��ҳid = [2]  " & strDiagType & " And a.��¼��Դ = [3] And b.����id = c.Id(+)  " & vbNewLine & _
                            "       Union All" & vbNewLine & _
                            "       Select Distinct a.Id, a.����id, a.��ҳid, a.��¼��Դ, Nvl(a.��ϴ���, 1) As ��ϴ���, Decode(Nvl(a.�������, 0), 0, 1, a.�������) �������," & vbNewLine & _
                            "                       a.�������, a.����id, a.���id, a.֤��id, '' || ������� As �������, a.��Ժ����, a.��Ժ���, �Ƿ�δ��, �Ƿ�����, a.����ʱ��, a.��ע," & vbNewLine & _
                            "                       '' || Null As ��������, '' || Null As ��������, '' || Null As �������, '' || Null As �Ա�����, '' || Null As ��Ч����," & vbNewLine & _
                            "                       '' || Null As ����, '' || Null As ����, 0 * Null As ����, 0 * Null As �Ƿ���, '' || Null As ֤�����," & vbNewLine & _
                            "                       '' || Null As ֤������, 0 * Null ����id, a.��¼����, a.��¼��" & vbNewLine & _
                            "       From ������ϼ�¼ a" & vbNewLine & _
                            "       Where a.����id = [1] And a.��ҳid = [2] " & strDiagType & " And a.��¼��Դ = 0 And a.����id Is Null And Not Exists" & vbNewLine & _
                            "        (Select 1" & vbNewLine & _
                            "              From ������ϼ�¼" & vbNewLine & _
                            "              Where a.����id = ����id And a.��ҳid = ��ҳid And a.������� = ������� And a.��ϴ��� = ��ϴ��� And ��¼��Դ = [3] And ����id Is Not Null)) a" & vbNewLine & _
                            "Order By a.�������, ��ϴ���, a.�������"
        End If
    End If
    If blnMoved Then
         strSql = Replace(strSql, "������ϼ�¼", "H������ϼ�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��ҳ���", lng����ID, lng��ҳID, int��¼��Դ)
    
    Set GetPatiDiagData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetOPSData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal bln��Ŀ As Boolean, Optional ByVal blnMoved As Boolean) As ADODB.Recordset
'���ܣ���ȡ���˵�������Ϣ
'���أ����˵�������Ϣ
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If gclsPros.FuncType <> f������ҳ Then
        strSql = "Select A.Id, A.����id, A.��ҳid, A.�������, A.��¼��Դ, A.��������, A.������ʼʱ��, A.��������ʱ��, Nvl(B.����, C.����) As ��������, A.�������� ��������," & vbNewLine & _
                "       Nvl(B.����, C.����) ����ԭ��, A.����ҽʦ, A.������ʿ, A.��һ����, A.�ڶ�����, A.����ҽʦ, A.׼������, A.������ҩʱ��, A.������ҩ����, A.����ʼʱ��, A.�ط�Ŀ��," & vbNewLine & _
                "       A.�пڲ�λ, A.��������, Decode(A.Asa�ּ�, 'I��', 'P1', 'II��', 'P2', 'III��', 'P3', 'IV��', 'P4', 'V��', 'P5', A.Asa�ּ�) Asa�ּ�, A.Nnis�ּ�, Decode(A.��������, 1, 'һ������', 2, '��������', 3, '��������', 4, '�ļ�����',9, '��', ' ') As ��������, A.�п�," & vbNewLine & _
                "       A.����, A.�ٴ�����, A.��ǰ������ҩ, A.��Ԥ�ڵĶ�������, A.������֢, A.������������, A.��������֢, A.�����Ѫ��Ѫ��, A.�����˿��ѿ�, A.�������Ѫ˨, A.���������л����," & vbNewLine & _
                "       A.�������˥��, A.�����˨��, A.�����Ѫ֢, A.�����Źؽڹ���, A.�ط��ƻ�, A.�пڸ�Ⱦ, A.����֢, A.��������id, A.������Ŀid, A.����ʽ ����id, D.���� ����ʽ, A.��¼����," & vbNewLine & _
                "       A.��¼��, A.ȡ��ʱ��, A.ȡ����, Decode(B.��������, '��', '�ļ�����', '��', '��������', '��', '��������', '��', 'һ������', '�ļ�', '�ļ�����', '����', '��������', '����', '��������', 'һ��', 'һ������', Null) ԭ�������� " & vbNewLine & _
                "From ���������¼ A, ��������Ŀ¼ B, ������ĿĿ¼ C, ������ĿĿ¼ D" & vbNewLine & _
                "Where C.Id(+) = A.������Ŀid And A.��������id = B.Id(+) And A.����ʽ = D.Id(+) And ����id = [1] And ��ҳid = [2] And" & vbNewLine & _
                "      (��¼��Դ <> 1 Or" & vbNewLine & _
                "       (��¼��Դ = 1 And ȡ��ʱ�� Is Null And" & vbNewLine & _
                "       ��¼���� =" & vbNewLine & _
                "       (Select Max(��¼����) From ���������¼ Where ����id = 1 And ��ҳid = 2 And ��¼��Դ = 1 And ȡ��ʱ�� Is Null)))" & vbNewLine & _
                "Order By Nvl(A.��������,999),A.ID"
    Else
        If bln��Ŀ Then
            strSql = "Select A.Id, A.�������, A.����id, A.��ҳid, A.��¼��Դ, A.��������, A.������ʼʱ��, A.��������ʱ��, B.���� As ��������, A.�������� ��������, B.���� ����ԭ��, A.����ҽʦ," & vbNewLine & _
                    "       A.������ʿ, A.��һ����, A.�ڶ�����, A.����ҽʦ, A.׼������, A.������ҩʱ��, A.������ҩ����, A.����ʼʱ��, A.�ط�Ŀ��, A.�пڲ�λ, A.��������," & vbNewLine & _
                    "       Decode(A.Asa�ּ�, 'I��', 'P1', 'II��', 'P2', 'III��', 'P3', 'IV��', 'P4', 'V��', 'P5', A.Asa�ּ�) Asa�ּ�, A.Nnis�ּ�," & vbNewLine & _
                    "       Decode(A.��������, 1, 'һ������', 2, '��������', 3, '��������', 4, '�ļ�����',9, '��', ' ') As ��������, A.�п�, A.����, A.�ٴ�����, A.��ǰ������ҩ, A.��Ԥ�ڵĶ�������," & vbNewLine & _
                    "       A.������֢, A.������������, A.��������֢, A.�����Ѫ��Ѫ��, A.�����˿��ѿ�, A.�������Ѫ˨, A.���������л����, A.�������˥��, A.�����˨��, A.�����Ѫ֢, A.�����Źؽڹ���," & vbNewLine & _
                    "       A.�ط��ƻ�, A.�пڸ�Ⱦ, A.����֢, A.��������id, A.������Ŀid, A.����ʽ ����id,Null ����ʽ,A.��¼����,A.��¼��, A.ȡ��ʱ��, A.ȡ����, " & vbNewLine & _
                    "      Decode(B.��������, '��', '�ļ�����', '��', '��������', '��', '��������', '��', 'һ������', '�ļ�', '�ļ�����', '����', '��������', '����', '��������', 'һ��', 'һ������', Null) ԭ�������� " & vbNewLine & _
                    "From ���������¼ A, ��������Ŀ¼ B" & vbNewLine & _
                    "Where A.��������id = B.Id(+) And A.����id = [1] And A.��ҳid = [2] And A.��¼��Դ = 4" & vbNewLine & _
                    "Order By Nvl(A.��������,999),A.Id"
        Else
            strSql = "Select A.Id, A.�������, A.����id, A.��ҳid, A.��¼��Դ,A.��������, A.������ʼʱ��, A.��������ʱ��, B.���� As ��������, A.�������� ��������, B.���� ����ԭ��, A.����ҽʦ," & vbNewLine & _
                    "       A.������ʿ, A.��һ����, A.�ڶ�����, A.����ҽʦ, A.׼������, A.������ҩʱ��, A.������ҩ����, A.����ʼʱ��, A.�ط�Ŀ��, A.�пڲ�λ, A.��������," & vbNewLine & _
                    "       Decode(A.Asa�ּ�, 'I��', 'P1', 'II��', 'P2', 'III��', 'P3', 'IV��', 'P4', 'V��', 'P5', A.Asa�ּ�) Asa�ּ�, A.Nnis�ּ�," & vbNewLine & _
                    "       Decode(A.��������, 1, 'һ������', 2, '��������', 3, '��������', 4, '�ļ�����',9, '��', ' ') As ��������, A.�п�, A.����, A.�ٴ�����, A.��ǰ������ҩ, A.��Ԥ�ڵĶ�������," & vbNewLine & _
                    "       A.������֢, A.������������, A.��������֢, A.�����Ѫ��Ѫ��, A.�����˿��ѿ�, A.�������Ѫ˨, A.���������л����, A.�������˥��, A.�����˨��, A.�����Ѫ֢, A.�����Źؽڹ���," & vbNewLine & _
                    "       A.�ط��ƻ�, A.�пڸ�Ⱦ, A.����֢, A.��������id, A.������Ŀid, A.����ʽ ����id,Null ����ʽ,A.��¼����,A.��¼��, A.ȡ��ʱ��, A.ȡ����, " & vbNewLine & _
                    "       Decode(B.��������, '��', '�ļ�����', '��', '��������', '��', '��������', '��', 'һ������', '�ļ�', '�ļ�����', '����', '��������', '����', '��������', 'һ��', 'һ������', Null) ԭ�������� " & vbNewLine & _
                    "From ���������¼ A, ��������Ŀ¼ B" & vbNewLine & _
                    "Where A.��������id = B.Id(+) And ����id = [1] And ��ҳid = [2]  " & vbNewLine & _
                    "Order By Nvl(A.��������,999),A.Id"
        End If
    End If
    
    If blnMoved Then
         strSql = Replace(strSql, "���������¼", "H���������¼")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", lng����ID, lng��ҳID)
    '����
    If Not bln��Ŀ Then
        rsTmp.Filter = "��¼��Դ=" & IIf(gclsPros.FuncType = f������ҳ, 4, 3)
        If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=" & IIf(gclsPros.FuncType = f������ҳ, 3, 1)
        If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=" & IIf(gclsPros.FuncType = f������ҳ, 1, 4)
    End If
    
    Set GetOPSData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDiagMatchData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˵���Ϸ������
'���أ����˵���Ϸ������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select ��������,������� From ��Ϸ������ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��Ϸ������", lng����ID, lng��ҳID)
    
    Set GetDiagMatchData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDiagExtraID(ByVal strCode As String) As ADODB.Recordset
'���ܣ���ȡ��������ID
'���أ������ڼ�������Ŀ¼�����Ӧ��ID
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH

    strSql = "Select ID from ��������Ŀ¼ where ���� = [1] and RowNum < 2 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��������ID", strCode)

    Set GetDiagExtraID = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetKSSData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˵Ŀ�����ʹ�������antibiotic)
'���أ����˵Ŀ�����ʹ�����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.ҩ��id, A.��ҩĿ��, A.ʹ�ý׶�, A.ʹ������, A.ҩƷ���� ����, һ���п�Ԥ����, Ddd��, ������ҩ" & vbNewLine & _
            "From ���˿����ؼ�¼ A" & vbNewLine & _
            "Where A.����id = [1] And A.��ҳid = [2]" & vbNewLine & _
            "Order By Ddd�� Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˿����ؼ�¼", lng����ID, lng��ҳID)
    
    Set GetKSSData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetChemothData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˻��Ƽ�¼
'���أ����˻��Ƽ�¼
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.����id, A.��ҳid, A.���, A.����id, A.��ʼ����, A.��������, A.�Ƴ���, A.����, A.���Ʒ���, A.����Ч��, B.���� || '-' || B.���� As ������Ϣ" & vbNewLine & _
            "From �������Ƽ�¼ A, ��������Ŀ¼ B" & vbNewLine & _
            "Where A.����id = B.Id And A.����id = [1] And A.��ҳid = [2]" & vbNewLine & _
            "Order By ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�������Ƽ�¼", lng����ID, lng��ҳID)
    
    Set GetChemothData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetRadiothData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˷������
'���أ����˷������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.����id, A.��ҳid, A.���, A.����id, A.��ʼ����, A.��������, A.��Ұ��λ, A.�������, A.�ۼ���, A.����Ч��, B.���� || '-' || B.���� As ������Ϣ" & vbNewLine & _
            "From �������Ƽ�¼ A, ��������Ŀ¼ B" & vbNewLine & _
            "Where A.����id = B.Id And A.����id = [1] And A.��ҳid =[2]" & vbNewLine & _
            "Order By ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˿����ؼ�¼", lng����ID, lng��ҳID)
    
    Set GetRadiothData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSpiritData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˾���ҩƷʹ�����
'���أ����˾���ҩƷʹ�����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select ���, ҩƷid, ҩ������, �Ƴ�, �������, ���ⷴӦ, ��Ч" & vbNewLine & _
            "From ������������" & vbNewLine & _
            "Where ����id = [1] And ��ҳid = [2]" & vbNewLine & _
            "Order By ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˿����ؼ�¼", lng����ID, lng��ҳID)
    
    Set GetSpiritData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetICUData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ������֢�໤ʹ�����
'���أ�������֢�໤ʹ�����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select ���, �໤������, To_Char(����ʱ��, 'yyyy-mm-dd HH24:mi') As ����ʱ��, To_Char(�˳�ʱ��, 'yyyy-mm-dd HH24:mi') As �˳�ʱ�� ,�˹������ѳ�,�ط���֢ҽѧ��,�ط����ʱ�� ,����ס�ƻ�,����סԭ��" & vbNewLine & _
            "From ������֢�໤���" & vbNewLine & _
            "Where ����id = [1] And ��ҳid = [2]" & vbNewLine & _
            "Order By ���"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������֢�໤���", lng����ID, lng��ҳID)
    
    Set GetICUData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetICUInstrumentsData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ������֢�໤ʹ�����
'���أ�������֢�໤ʹ�����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.���, A.��� ||'-' ||A.�໤������ �໤������, C.����||'.'||C.���� As ��е������, To_Char(��ʼʹ��ʱ��, 'yyyy-mm-dd HH24:mi') As ��ʼʹ��ʱ��," & vbNewLine & _
                "       To_Char(����ʹ��ʱ��, 'yyyy-mm-dd HH24:mi') As ����ʹ��ʱ��, ��Ⱦ�ۼ�ʱ��" & vbNewLine & _
                "From ��е����ʹ����� A, ��е����Ŀ¼ C" & vbNewLine & _
                "Where A.��е������ = C.����(+) And ����id = [1] And ��ҳid = [2]" & vbNewLine & _
                "Order By ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��е����ʹ�����", lng����ID, lng��ҳID)
    
    Set GetICUInstrumentsData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetInfectData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ������֢�໤ʹ�����
'���أ�������֢�໤ʹ�����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.���,To_Char(A.ȷ������, 'yyyy-mm-dd') As ȷ������, B.���� || '.' || A.��Ⱦ��λ ��Ⱦ��λ, A.��Ⱦ���� ҽԺ��Ⱦ����, C.���� ҽԺ��Ⱦ����" & vbNewLine & _
                    "From ���˸�Ⱦ��¼ A, ��Ⱦ��λ B, ҽԺ��ȾĿ¼ C" & vbNewLine & _
                    "Where A.��Ⱦ��λ = B.����(+) And A.��Ⱦ���� = C.����(+) And A.����id = [1] And A.��ҳid = [2]" & vbNewLine & _
                    "Order By A.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˸�Ⱦ��¼", lng����ID, lng��ҳID)
    
    Set GetInfectData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSampleData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ������֢�໤ʹ�����
'���أ�������֢�໤ʹ�����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select A.���,A.�걾, A.��ԭѧ���� || '-' || B.���� As ��ԭѧ����, To_Char(A.�ͼ�����, 'yyyy-mm-dd') As �ͼ�����" & vbNewLine & _
                    "From ���˲�ԭѧ��� A, ��ԭѧĿ¼ B" & vbNewLine & _
                    "Where A.��ԭѧ���� = B.����(+) And A.����id = [1] And A.��ҳid = [2]" & vbNewLine & _
                    "Order By A.���"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�걾��Դ", lng����ID, lng��ҳID)
    
    Set GetSampleData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckMergePath(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngDiagType As Long, ByVal lngDiag As Long) As Boolean
'���ܣ�����ٴ�·����Ӧ����ϲ����޸�
'������lngDiagType���������,lngDiag=����ID
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If lngDiag = 0 Or lngDiagType = 0 Then CheckMergePath = True: Exit Function
    strSql = " Select �������,����ID From �����ٴ�·�� Where ����ID=[1] And ��ҳID=[2]" & _
             " Union " & _
             " Select �������,����ID From ���˺ϲ�·�� Where ����ID=[1] And ��ҳID=[2]"
             
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gstrSysName, lng����ID, lng��ҳID)
    Do While Not rsTmp.EOF
        If lngDiagType = Val(rsTmp!������� & "") And lngDiag = Val(rsTmp!����id & "") Then
            Exit Function
        End If
        rsTmp.MoveNext
    Loop
    CheckMergePath = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckSign(ByVal intǩ������ As Long, ByVal lng��������ID As Long, Optional ByVal lngҽ������ID As Long, Optional ByVal lng���˿���ID As Long, _
    Optional ByVal int���˷�Χ As Integer = 2, Optional ByVal blnCheckCert As Boolean = True) As Boolean
'���ܣ��ж�һ�����Ż���һ�鲿�����Ƿ���������˵���ǩ�����Ƶ�
'������int���˷�Χ=1-����,2-סԺ(ȱʡ)
'     intǩ������:0-����ҽ���Ͳ�����1-סԺҽ��ҽ���Ͳ�����2-סԺ��ʿҽ����3-ҽ��ҽ���ͱ��棻4-�����¼�ͻ�������5-ҩƷ��ҩ��6-LIS;7-PACS;
'     lng��������ID=���lng��������ID=0������Ҫ���ݴ����ҽ�����ң����˿���ID���Ӧ��Ĭ�Ͽ�������
'                   ��ʿվУ�Ժ�ȷ��ֹͣʱ������Ĳ���ID�����жϲ����Ƿ������˵���ǩ��
'                   ����-1������ҩ�����ʱ������ж��Ƿ�ֿ������ã�
'     blnCheckCert=true ���֤���Ƿ�ͣ�ã�=false��ʾ�����
    Dim strSql As String, intTmp As Integer
    Dim rsTmp As Recordset
    
    '������϶�δ���ã��򷵻�false
    If intǩ������ = 0 Or intǩ������ = 1 Then
        intTmp = intǩ������ + 1
    ElseIf intǩ������ > 1 And intǩ������ <= 7 Then
        intTmp = intǩ������
    End If
    If Mid(gstrESign, intTmp, 1) <> "1" Then Exit Function
    If lng��������ID = 0 And (lng���˿���ID <> 0 Or lngҽ������ID <> 0) Then
        'ȡ��������
        lng��������ID = Get��������ID(UserInfo.ID, lngҽ������ID, lng���˿���ID, int���˷�Χ)
        If lng��������ID = 0 Then Exit Function
    End If
    grsSign.Filter = "����ID=" & lng��������ID & " and ����=" & intǩ������
    If grsSign.RecordCount = 0 Then
        strSql = "Select Zl_Fun_Getsignpar([1],[2]) as �Ƿ����� From dual"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlAdvice", intǩ������, lng��������ID)
        If rsTmp.RecordCount > 0 Then
            CheckSign = Val(rsTmp!�Ƿ����� & "") = 1
            grsSign.AddNew
            grsSign!����ID = lng��������ID
            grsSign!���� = intǩ������
            grsSign!�Ƿ����� = Val(rsTmp!�Ƿ����� & "")
        End If
    Else
        grsSign.MoveFirst
        CheckSign = Val(grsSign!�Ƿ����� & "") = 1
    End If
    If CheckSign = True And blnCheckCert Then
        If gobjESign Is Nothing Then
            On Error Resume Next
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            Err.Clear: On Error GoTo 0
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, gclsPros.SysNo)
            End If
        End If
        '���֤���Ƿ�ͣ��
        If gobjESign.CertificateStoped(UserInfo.����) Then CheckSign = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��������ID(ByVal lngҽ��ID As Long, ByVal lngҽ������ID As Long, ByVal lng���˿���ID As Long, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal lngִ�п���ID As Long) As Long
'���ܣ���ҽ��ȷ����������
'������int��Χ=1-����,2-סԺ(ȱʡ)
'˵������ҽ���������ҷ�Χ��,����˳�����£�
'      1��ҽ������(ҽ������)
'      2�����˿���
'      3������������/סԺ���˵�ĳЩ����ҽ����ִ�п���
'      4������������/סԺ���˵Ŀ�����ΪĬ�Ͽ���
'      5������������/סԺ���˵Ŀ���
'      6��Ĭ�Ͽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Integer
    Dim arr����ID(1 To 6) As Long
    
    '�������ű������ٴ���ҽ��
    strSql = "Select Distinct A.����ID,Nvl(A.ȱʡ,0) as ȱʡ" & _
        " From ������Ա A,��������˵�� B,���ű� C" & _
        " Where A.����ID=C.ID And A.����ID=B.����ID" & _
        " And B.������� IN([2],3) And A.��ԱID=[1]" & _
        " And B.�������� IN('�ٴ�','���','����','����','����','Ӫ��')" & _
        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lngҽ��ID, int��Χ)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!����ID = lngҽ������ID Then
            arr����ID(1) = rsTmp!����ID
        ElseIf rsTmp!����ID = lng���˿���ID Then
            arr����ID(2) = rsTmp!����ID
        ElseIf rsTmp!����ID = lngִ�п���ID Then
            arr����ID(3) = rsTmp!����ID
        ElseIf rsTmp!ȱʡ = 1 Then
            arr����ID(4) = rsTmp!����ID
        ElseIf arr����ID(4) = 0 Then
            arr����ID(5) = rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
    arr����ID(6) = UserInfo.DeptID
    
    For i = LBound(arr����ID) To UBound(arr����ID)
        If arr����ID(i) <> 0 Then
            Get��������ID = arr����ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckMecRed(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strfrmCation As String, Optional ByVal strOperateName As String) As Boolean
'���ܣ���鲡���Ƿ��Ѿ���Ŀ,�����Ƿ��ڴ������������(��ʱ��ҳ��������״̬���������޸�)
'       lng����ID:��ǰ����ID
'       lng��ҳID:��ǰ������ҳID
'       strfrmCation:���øú����Ĵ�������
'       strOperateName:���øú����Ĳ������ơ�strOperateNameΪ��ʱ����������ʾ
    Dim strSql As String, rsTmp As Recordset
    Dim int����״̬ As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    '��ȡ����״̬
    strSql = "Select Nvl(����״̬, 0) ����״̬,��Ŀ���� From ������ҳ Where ����id = [1] And ��ҳid = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, strfrmCation, lng����ID, lng��ҳID)
    rsTmp.MoveFirst
    int����״̬ = rsTmp!����״̬
    '��ҳ���������ж�
    Select Case int����״̬
        Case 1 '�ȴ����
            strMsg = "�ò����ȴ������,����"
        Case 3 '�������
            strMsg = "�ò������������,����"
        Case 5 '���鵵
            strMsg = "�ò����Ѿ����鵵,����"
        Case 10 '���մ���
            strMsg = "�ò����ڽ��մ�����,����"
        Case Else '2-�ܾ����4-��鷴��;6-�������;13-���ڳ��;14-��鷴��;16-�������
            strMsg = ""
    End Select
    
    If strMsg = "" Then
        If Not IsNull(rsTmp!��Ŀ����) Then
            strMsg = "�ò��˵Ĳ����Ѿ���Ŀ������"
        End If
    End If
    
    If strMsg <> "" Then  '������ҳ
        If strOperateName <> "" Then
            MsgBox strMsg & strOperateName & "��", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    CheckMecRed = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiPathInfo() As Boolean
'���ܣ���ȡ���˵��ٴ�·�������Ϣ
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If gclsPros.PathState <> PS_δ���� Then
        'ֻ������ҳ���������ϣ���ǰû��ģ�ȱʡ���������ڡ���ҽ��Ժ��ϡ���������������зǿգ����ȡ�������߼�
        strSql = "Select Nvl(�������, 2) As �������, Nvl(����id, 0) As ����id, Nvl(���id, 0) As ���id, ״̬" & vbNewLine & _
                "From �����ٴ�·��" & vbNewLine & _
                "Where ����id = [1] And ��ҳid = [2] And (�����Դ = 3 Or �����Դ Is Null)" & vbNewLine & _
                "Order By ����ʱ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, gclsPros.��ҳID)
        If rsTmp.RecordCount > 0 Then
            gclsPros.InPath = rsTmp!�������
            '����ж���·������ȡ��һ����״̬
            If rsTmp.RecordCount >= 2 Then gclsPros.PathState = Val(rsTmp!״̬ & "")
            rsTmp.MoveNext
            Do While Not rsTmp.EOF
                gclsPros.PathDiag = gclsPros.PathDiag & "," & rsTmp!������� & "|" & rsTmp!����id & "|" & rsTmp!���ID
                rsTmp.MoveNext
            Loop
            gclsPros.PathDiag = Mid(gclsPros.PathDiag, 2)
        Else
            gclsPros.InPath = 0
        End If
        '���·����ʱ���Ƿ�ȳ�Ժ��ϼ�¼ʱ���()ȡ��һ��·��
        If gclsPros.PathState = PS_�������� Then
            strSql = "Select Sign(Nvl(A.����ʱ��, Null) - Nvl(B.��¼����, Sysdate)) As �ж�" & vbNewLine & _
                    "From �����ٴ�·�� A, (Select ����id, ��ҳid, ��¼���� From ������ϼ�¼ Where ��¼��Դ = 3 And ��ϴ��� = 1 And ������� = [3]) B" & vbNewLine & _
                    "Where A.����id = B.����id(+) And A.��ҳid = B.��ҳid(+) And A.����id = [1] And A.��ҳid = [2] And" & vbNewLine & _
                    "      A.����ʱ�� = (Select Min(����ʱ��) From �����ٴ�·�� Where ����id = [1] And ��ҳid = [2])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, gclsPros.��ҳID, IIf(gclsPros.InPath > 10, DT_��Ժ���ZY, DT_��Ժ���XY))
            If rsTmp.RecordCount > 0 Then
                gclsPros.PathOutTime = Val(rsTmp!�ж� & "") = 1
            Else
                gclsPros.PathOutTime = False
            End If
        End If
    End If
    GetPatiPathInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStrucAddress(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strTypeName As String) As ADODB.Recordset
'���ܣ����ָ�����͵Ľṹ����ַ
'������strTypeName=��ַ���� �����ص�(������1),����(������2),��סַ(������3),���ڵ�ַ(������4)
    Dim strSql As String, rsTmp As Recordset
    Dim lngType As Long, blnNew As Boolean
    
    lngType = Decode(strTypeName, "�����ص�", 1, "����", 2, "��ͥ��ַ", 3, "���ڵ�ַ", 4, "��ϵ�˵�ַ", 5, "��λ��ַ", 6)
    
    blnNew = gclsPros.AdressInfo Is Nothing
    If blnNew Then
        strSql = "Select ����ID,��ҳID,��ַ���,ʡ,��,��,����,���� From ���˵�ַ��Ϣ Where ����ID=[1] And ��ҳID=[2]"
        On Error GoTo errH
        Set gclsPros.AdressInfo = zlDatabase.OpenSQLRecord(strSql, "��ѯ�ṹ����ַ", lng����ID, lng��ҳID)
    End If
    
    gclsPros.AdressInfo.Filter = "��ַ���=" & lngType
    Set GetStrucAddress = gclsPros.AdressInfo
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetKSSID(ByVal strName As String) As Long
'���ܣ��������ڽ�������ҳ�ӱ�Ŀ����� �Ƶ����±� ���˿����ؼ�¼�У���ǰû�м�¼ҩƷid�����ڸ������ƽ�id�����
'������strName=ҩƷ��
    Dim rsTmp As Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "Select Distinct A.Id From ������ĿĿ¼ A, ҩƷ���� C Where A.Id = C.ҩ��id And Nvl(C.������, 0) <> 0 And A.���� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strName)
    If rsTmp.RecordCount > 0 Then
        GetKSSID = Val(rsTmp!ID)
    Else
        GetKSSID = 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��ҳIDByCur(ByVal lng��ҳID As Long, Optional ByVal blnNext As Boolean = True) As Long
'���ܣ����ݵ�ǰ��ҳID��ȡָ������ҳID
'������lng��ҳID=Ҫ�����жϵ���ҳID
'      blnNext=True-��ȡ��lng��ҳID�����С��ҳID,False-��ȡ��lng��ҳIDС�������ҳID
'���أ�0-��������������ҳID,>0:������������ҳID
    Dim strSql As String, rsTmp As ADODB.Recordset
    If gclsPros.OpenMode = EM_�༭ Or gclsPros.OpenMode = EM_���� Then
        If blnNext Then
            strSql = "Select Min(A.��ҳid) As ��ҳid" & vbNewLine & _
                            "From ������ҳ A" & vbNewLine & _
                            "Where A.����id = [1] And Nvl(��������, 0) = 0 And ��Ŀ���� Is Not Null And ��ҳid > [2]"
        Else
           strSql = "Select Max(A.��ҳid) As ��ҳid" & vbNewLine & _
                            "From ������ҳ A" & vbNewLine & _
                            "Where A.����id = [1] And Nvl(��������, 0) = 0 And ��Ŀ���� Is Not Null And ��ҳid < [2]"
        End If
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��ҳID", gclsPros.����ID, lng��ҳID)
        If Not rsTmp.EOF Then
            Get��ҳIDByCur = IIf(IsNull(rsTmp!��ҳID), 0, Val(rsTmp!��ҳID & ""))
        Else
            Get��ҳIDByCur = 0
        End If
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetסԺ����Or��ҳid(ByVal lng����ID As Long, ByRef lng��ҳID As Long, ByRef lng���� As Long, ByVal bln��ȡ��ҳid As Boolean, Optional ByVal blnVali��ҳ As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ȡסԺ��������ҳID
    '����:lng����id-����id
    '     lng��ҳID-���˵���ҳID
    '     lng����=סԺ����(��ȥ���۲���)
    '     bln��ȡ��ҳid-true��ʾ��ȡ��ҳid,�����ȡסԺ����(��ȥ���۲���)
    '     blnVali��ҳ=�Ƿ���֤��ҳID��false,����֤��Ture-��֤����ʱbln��ȡ��ҳid �� False
    '����:lng����-����סԺ��������ҳid
    '����:��ȡ�Ĵ����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/5/10
    '-----------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    Dim int��ʶ As Integer
    
    Err = 0: On Error GoTo Errhand:
    ' Zl_��ȡסԺ��������ҳid
    '  ����id_In ������ҳ.����id%Type,
    '  ����_In   ������ҳ.��ҳid%Type,
    '  ��ʶ_In   Integer:1-����ָ����������ҳid,0-������ҳid����סԺ����(�ų������۲���)
    If Not blnVali��ҳ Then
        strSql = " Select Zl_��ȡסԺ��������ҳid([1],[2],[3]) As ���� From Dual"
    Else
        strSql = "Select Zl_��ȡסԺ��������ҳid(A.����ID, A.��ҳID,[3]) As ����" & vbNewLine & _
                "From ������ҳ A" & vbNewLine & _
                "Where A.����id = [1] And A.��Ŀ���� Is Not Null And A.��ҳid >[2]" & vbNewLine & _
                "Order By A.��ҳid Desc"
    End If
    If Not blnVali��ҳ Then int��ʶ = IIf(bln��ȡ��ҳid, 1, 0)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡסԺ��������ҳid", lng����ID, IIf(bln��ȡ��ҳid, lng����, lng��ҳID), int��ʶ)
    
    If Not rsTemp.EOF Then
        If Not bln��ȡ��ҳid Or blnVali��ҳ Then
            lng���� = Val(rsTemp!���� & "")
            If blnVali��ҳ Then
                MsgBox "��ѡ���" & lng���� & "����Ժ�Ժ�Ĳ�����Ϣ��", vbInformation, gstrSysName
                GetסԺ����Or��ҳid = False
                Exit Function
            End If
        Else
            lng��ҳID = Val(rsTemp!���� & "")
        End If
    ElseIf Not blnVali��ҳ Then
        If Not bln��ȡ��ҳid Or blnVali��ҳ Then
            lng���� = 0
        Else
            lng��ҳID = 0
        End If
    End If
    
    GetסԺ����Or��ҳid = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    GetסԺ����Or��ҳid = False
    If Not bln��ȡ��ҳid Or blnVali��ҳ Then
        lng���� = 0
    Else
        lng��ҳID = 0
    End If
End Function

Public Function GetNextNo(ByVal int��� As Integer, Optional ByVal int�ж� As Integer = 0, Optional ByVal strCode As String = "") As Variant
'����:�����ض���������µĺ���,����������ֻ����ZLHIS10������ҪOracle 8i(8.1.5)���ϰ汾֧��
'������
'int���=��Ŀ���:
'  1   ����ID ����
'  2   סԺ�� ����
'���أ�������
'˵����
'  ��Ź���0-����˳����,1-����˳����,2-��ִ�п��ҷ��±��(��Ҫ��ȡ���Һ����)
'            ������ţ�0-˳����,1-������(YYMMDD)+˳���(0000)
'            ��סԺ�ţ�0-˳����,1-����(YYMM)+˳���(0000),2-��(YYYY)+˳���(00000)
'  ���λȷ������1990Ϊ���������������������0��9/A��Z��˳����Ϊ��ȱ���
'  ������-10���������Ʊ�,���ڲ�������²�ȱ��(ȡ�˺�,��δʹ��)
'  For Update�ڲ��������������,����Waitѡ���Ա���������߷��ؿ�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    GetNextNo = Null
    
    On Error GoTo errH
    '���� 25779 ���ڵ���zl3_NextNO����,����int�ж� by lesfeng 2009-10-16 b
    If int�ж� = 0 Then
        strSql = "Select zl3_NextNO([1],[2],[3]) as NO From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetNextNo", int���, 0, strCode)
    Else
        strSql = "Select zl3_NextNO([1],[2],[3]) as NO From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetNextNo", int���, int�ж�, strCode)
    End If
    '���� 25779 ���ڵ���zl3_NextNO����,����int�ж� by lesfeng 2009-10-16 b
    If gcnOracle.Errors.Count > 0 Then 'Select�к�������ʱ,��VB�в��Զ���������
        Err.Raise gcnOracle.Errors(0).Number, , gcnOracle.Errors(0).Description
    End If
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!No) Then GetNextNo = rsTmp!No
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'����26071
Public Function GetBloodValue(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����Ѫ����Ѫ������Ѫ��ص���Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:Lesfeng
    '����:2009-11-18 12:11:40
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim strCombItem As String
        
    On Error GoTo Errhand
    
    With gclsPros.CurrentForm
        'Zl_GetѪ����Ѫ��Ϣ
        strSql = "select Zl_GetѪ����Ѫ��Ϣ([1],[2]) as Blood from dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption, lng����ID, lng��ҳID)
        If Not rsTmp.EOF Then
            strTmp = IIf(IsNull(rsTmp!Blood), "0", rsTmp!Blood)
            If strTmp <> "0" Then
                arrTmp = Split(strTmp, "|")
                strCombItem = ""
                If arrTmp(1) = "δ֪" Then arrTmp(1) = "����"
                Set rsTmp = GetBaseCode("Ѫ��")
                Do While Not rsTmp.EOF
                    strCombItem = strCombItem & "," & rsTmp!����
                Loop

                If Trim(arrTmp(1)) <> "" And strCombItem <> "" And InStr(1, strCombItem, Trim(arrTmp(1))) > 0 Then
                    .cboBaseInfo(BCC_Ѫ��).Text = Trim(arrTmp(1))
                Else
                    If Trim(arrTmp(1)) <> "" Then
                        .cboBaseInfo(BCC_Ѫ��).AddItem arrTmp(1)
                        .cboBaseInfo(BCC_Ѫ��).ListIndex = .cboBaseInfo(BCC_Ѫ��).NewIndex
                    End If
                End If
                strCombItem = "δ��,��,��,����"
                If arrTmp(2) = "δ��" Then arrTmp(2) = "δ��"
                If Trim(arrTmp(2)) <> "" And InStr(1, strCombItem, Trim(arrTmp(2))) > 0 Then
                    .cboBaseInfo(BCC_RH).Text = Trim(arrTmp(2))
                Else
                    If Trim(arrTmp(2)) <> "" Then
                        .cboBaseInfo(BCC_RH).AddItem arrTmp(2)
                        .cboBaseInfo(BCC_RH).ListIndex = .cboBaseInfo(BCC_RH).NewIndex
                    End If
                End If
                If Trim(arrTmp(3)) <> "0" And IsNumeric(arrTmp(3)) Then
                    .txtSpecificInfo(SLC_���ϸ��) = Trim(arrTmp(3))
                End If
                If Trim(arrTmp(4)) <> "0" And IsNumeric(arrTmp(4)) Then
                    .txtSpecificInfo(SLC_��ѪС��) = Trim(arrTmp(4))
                End If
                If Trim(arrTmp(5)) <> "0" And IsNumeric(arrTmp(5)) Then
                    .txtSpecificInfo(SLC_��Ѫ��) = Trim(arrTmp(5))
                End If
                If Trim(arrTmp(6)) <> "0" And IsNumeric(arrTmp(6)) Then
                    .txtSpecificInfo(SLC_��ȫѪ) = Trim(arrTmp(6))
                End If
                If Trim(arrTmp(7)) <> "0" And IsNumeric(arrTmp(7)) Then
                    .txtInfo(GC_������) = Trim(arrTmp(7))
                End If
                
                strCombItem = "��,��,δ��"
                If Trim(arrTmp(8)) <> "" And InStr(1, strCombItem, Trim(arrTmp(8))) > 0 Then
                    .cboBaseInfo(BCC_��Ѫ��Ӧ).Text = Trim(arrTmp(8))
                Else
                    If Trim(arrTmp(8)) <> "" Then
                        .cboBaseInfo(BCC_��Ѫ��Ӧ).AddItem arrTmp(8)
                        .cboBaseInfo(BCC_��Ѫ��Ӧ).ListIndex = .cboBaseInfo(BCC_��Ѫ��Ӧ).NewIndex
                    End If
                End If
            End If
        End If
        rsTmp.Close
    End With
    GetBloodValue = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCareValue(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݻ���ӿڼ��ػ�����ص���Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:��˶
    '����:2013-12-26 10:06:40
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim strItems As String
    
    On Error GoTo Errhand
    strItems = "�ؼ�����,һ������,��������,��������,ICU,CCU"
    'Zl3_Get��������
    strSql = "Select Zl3_Get��������([1], [2], [3]) As CareData From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng����ID, lng��ҳID, strItems)
    If rsTmp.EOF Then Exit Function
    If IsNull(rsTmp!CareData) Then Exit Function
    strTmp = rsTmp!CareData & ""
    arrTmp = Split(strTmp, "|")
    If UBound(arrTmp) <> UBound(Split(strItems, ",")) Then Exit Function
    With gclsPros.CurrentForm
        If arrTmp(0) <> "" Then
            .txtSpecificInfo(SLC_�ػ�).Text = Format(Val(arrTmp(0)), "###;-###;;")
        End If
        If arrTmp(1) <> "" Then
            .txtSpecificInfo(SLC_һ������).Text = Format(Val(arrTmp(1)), "###;-###;;")
        End If
        If arrTmp(2) <> "" Then
            .txtSpecificInfo(SLC_��������).Text = Format(Val(arrTmp(2)), "###;-###;;")
        End If
        If arrTmp(3) <> "" Then
            .txtSpecificInfo(SLC_��������).Text = Format(Val(arrTmp(3)), "###;-###;;")
        End If
        If arrTmp(4) <> "" Then
            .txtSpecificInfo(SLC_ICU).Text = Format(Val(arrTmp(4)), "###;-###;;")
        End If
        If arrTmp(5) <> "" Then
            .txtSpecificInfo(SLC_CCU).Text = Format(Val(arrTmp(5)), "###;-###;;")
        End If
    End With
    GetCareValue = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCareData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '���:lng����ID=����ID
    '     lng��ҳID=������ҳID
    '����:
    '����:���ػ����¼��
    '����:��˶
    '����:2013-12-26 10:32:02
    '----------------------------------------------------------------------------------------------
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select b.��Ŀ��λ as ��λ, b.��Ŀ���� as ��Ϣ��, b.��¼���� as ��Ϣֵ" & _
        " From ���˻����¼ A, ���˻������� B Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"
    Set GetCareData = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetFreeData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal bln��Ŀ As Boolean) As ADODB.Recordset
'-----------------------------------------------------------------------------------------------------------
'����:��ȡ��������
'���:lng����ID=����ID
'     lng��ҳID=������ҳID
'     bln��Ŀ=�Ƿ��ȡ��Ŀ������
'����:
'����:���ط��ü�¼��
'����:��˶
'����:2013-12-26 10:32:02
'----------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    If bln��Ŀ Then
        strSql = "Select ����, ���� ��Ŀ����, �ϼ�, �ϼ� || Decode(Nvl(�ϼ�, ''), '', '', '_') || ���� || '.' || ���� ����, ĩ��,���,0 Ӥ����" & vbNewLine & _
                "From (Select A.����, A.�ϼ�,A.ĩ��,A.����, B.���" & vbNewLine & _
                "       From ������Ŀ A," & vbNewLine & _
                "            (Select ������, Sum(���) ���" & vbNewLine & _
                "              From ���˷���" & vbNewLine & _
                "              Where ����id = [1] And ��ҳid = [2] And Nvl(����, 0) = 0" & vbNewLine & _
                "              Group By ������) B" & vbNewLine & _
                "       Where A.���� = B.������(+))" & vbNewLine & _
                "Start With �ϼ� Is Null" & vbNewLine & _
                "Connect By Prior ���� = �ϼ�" & vbNewLine & _
                "Order By �ϼ� || ����"
    Else
        strSql = "Select /*+ Rule*/" & vbNewLine & _
                " ����, ��Ŀ����,�ϼ�, ����, ĩ��, ���, Ӥ����" & vbNewLine & _
                "From (Select B.����, B.���� ��Ŀ����,B.�ϼ�, B.�ϼ� || Decode(Nvl(B.�ϼ�, ''), '', '', '_') || B.���� || '.' || B.���� ����, B.ĩ��," & vbNewLine & _
                "              Sum(Nvl(A.���, 0)) As ���, Nvl(A.Ӥ����, 0) Ӥ����" & vbNewLine & _
                "       From (Select ����, �ϼ�, ����, ĩ�� From ������Ŀ Start With �ϼ� Is Null Connect By Prior ���� = �ϼ�) B," & vbNewLine & _
                "            (Select B.����, A.���, A.Ӥ����" & vbNewLine & _
                "              From (Select B.������Ŀ, Nvl(A.Ӥ����, 0) Ӥ����, Sum(Nvl(A.ʵ�ս��, 0)) As ���" & vbNewLine & _
                "                     From סԺ���ü�¼ A, �շ���ĿĿ¼ B" & vbNewLine & _
                "                     Where A.�շ�ϸĿid = B.Id And A.��¼״̬ <> 0 And A.����id = [1] And A.��ҳid = [2]" & vbNewLine & _
                "                     Group By B.������Ŀ, Nvl(A.Ӥ����, 0)) A, ������Ŀ B" & vbNewLine & _
                "              Where A.������Ŀ = B.����) A" & vbNewLine & _
                "       Where B.���� = A.����(+)" & vbNewLine & _
                "       Group By B.����, B.����, B.ĩ��, B.�ϼ�, Nvl(A.Ӥ����, 0))" & vbNewLine & _
                "Start With �ϼ� Is Null" & vbNewLine & _
                "Connect By Prior ���� = �ϼ�" & vbNewLine & _
                "Order By �ϼ� || ����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng����ID, lng��ҳID)
    Set GetFreeData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBabyInfoData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ���˷�����Ϣ����������Ϣ��
    '���:lng����ID=����ID
    '     lng��ҳID=������ҳID
    '����:
    '����:������������Ϣ
    '����:��˶
    '����:2013-12-27 16:34:02
    '----------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "Select ����id, ��ҳid,����ʱ��, ̥������, ���䷽ʽ, ����̥λ, �������, ����ȱ��, Ӥ���Ա�, Ӥ������, Apgar����" & vbNewLine & _
            "From ���˷�����Ϣ" & vbNewLine & _
            "Where ����id = [1] And ��ҳid = [2]"

    Set GetBabyInfoData = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBabyDiagData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��������ϼ�¼��������������Ϣ��
    '���:lng����ID=����ID
    '     lng��ҳID=������ҳID
    '����:
    '����:����������������Ϣ
    '����:��˶
    '����:2013-12-27 16:34:02
    '----------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "Select A.����id, A.��ҳid, A.̥������, A.��ϴ���, A.����id, A.������Ϣ, B.����" & vbNewLine & _
            "From ��������ϼ�¼ A, ��������Ŀ¼ B" & vbNewLine & _
            "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.Id" & vbNewLine & _
            "Order By ��ϴ���"
            
    Set GetBabyDiagData = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiTransfer(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ����ת����Ϣ
    '���:lng����ID=����ID
    '     lng��ҳID=������ҳID
    '����:
    '����:���ز���ת����Ϣ
    '����:��˶
    '����:2013-1-2 10:20:11
    '----------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = _
                " Select A.����ID,B.���� AS ��������,A.��ʼʱ��" & _
                " From ���˱䶯��¼ A,���ű� B" & _
                " Where A.����ID=[1] And A.��ҳID=[2]" & _
                " And A.����ID=B.ID And A.��ʼԭ��=3 And A.��ʼʱ�� is Not NULL" & _
                " Order by A.��ʼʱ��"
                   
    Set GetPatiTransfer = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetAdvicePause(ByVal lngҽ��ID As Long) As String
'���ܣ���ȡָ��ҽ������ͣʱ��μ�¼
'���أ�"��ͣʱ��,��ʼʱ��;...."
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    strSql = "Select ��������,����ʱ�� From ����ҽ��״̬" & _
        " Where �������� IN(6,7) And ҽ��ID=[1]" & _
        " Order by ����ʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISWork", lngҽ��ID)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!�������� = 6 Then
            strTmp = strTmp & ";" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & ","
        ElseIf rsTmp!�������� = 7 Then
            '���õ���һ�벻����ͣ�ķ�Χ֮��
            strTmp = strTmp & Format(DateAdd("s", -1, rsTmp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
        End If
        rsTmp.MoveNext
    Next
    GetAdvicePause = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function AutoGetOPSInfo(ByVal bln���� As Boolean, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ����ϵͳ��ҽ���е�������Ϣ����Ҫ�������ж�ȡ���������ȡ��������û�а�װ���飬���ȡҽ��
'������bln���� �Ƿ��ȡ����ϵͳ
'      lng����ID ����ID
'      lng��ҳID ��ҳID
'���أ�������Ϣ��¼��
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, rsOther As ADODB.Recordset
    Dim rsReturn As ADODB.Recordset
    Dim strTmp As String
    Dim lng��������id As Long
    Dim blnDefault As Boolean
    
    Dim blnReadAdvice As Boolean
    
    blnReadAdvice = Not bln����
    
    On Error GoTo errH
    'Ϊ�˶�ȡ��ṹ����˲���ѯ����
    strSql = "Select a.ID,a.������ʿ, a.�п�, a.����, b.���� ��������, a.��������, a.������ʼʱ��, a.��������ʱ��, a.��������, a.��������id, a.������Ŀid, a.��������, a.����ҽʦ," & vbNewLine & _
            "       a.��һ����, a.�ڶ�����, a.������ʿ, a.����ʼʱ��, a.�������ʱ��,C.���� ����ԭ�� , C.���� ����ʽ , A.����ʽ ����id,  a.��������, a.��������, a.��Һ����, a.����ҽʦ, a.������ʼʱ��, a.��������ʱ��, a.�������, a.ASA�ּ�," & vbNewLine & _
            "       a.�ٴ�����, a.NNIS�ּ�, 'һ������' ��������, a.��ǰ������ҩ, a.������ҩ����, a.��Ԥ�ڵĶ�������, a.������֢, a.������������, a.��������֢, a.�����Ѫ��Ѫ��, a.�����˿��ѿ�," & vbNewLine & _
            "       a.�������Ѫ˨, a.���������л����, a.�������˥��, a.�����˨��, a.�����Ѫ֢, a.�����Źؽڹ���, a.׼������, a.������ҩʱ��, a.�пڲ�λ, a.�ط��ƻ�, a.�ط�Ŀ��, a.�пڸ�Ⱦ," & vbNewLine & _
            "       a.����֢" & vbNewLine & _
            "From ���������¼ A, ��������Ŀ¼ B, ������ĿĿ¼ C" & vbNewLine & _
            "Where c.Id = a.������Ŀid And a.��������id = b.Id And ����id = 0 And ��ҳid = 0 And ��¼��Դ = 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ")
    Set rsReturn = zlDatabase.CopyNewRec(rsTmp, True)
    With rsReturn
        If bln���� Then
            strSql = "Select a.Id, a.����ʱ�� ��������, a.��ʼʱ�� ������ʼʱ��, a.����ʱ�� ��������ʱ��, a.����ʼ ����ʼʱ��, a.������� �������ʱ��, a.�������� ��������," & vbNewLine & _
                    "       a.������ʼ ������ʼʱ��, a.�������� ��������ʱ��, a.������ģ ��������" & vbNewLine & _
                    "From ����������ҳ A" & vbNewLine & _
                    "Where ����id = [1] And ��ҳid = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", lng����ID, lng��ҳID)
            If rsTmp.EOF Then
                blnReadAdvice = True
            Else
                While Not rsTmp.EOF
                    blnDefault = True
                    .AddNew
                    !ID = rsTmp!ID
                    !�������� = rsTmp!��������
                    !������ʼʱ�� = rsTmp!������ʼʱ��
                    !��������ʱ�� = rsTmp!��������ʱ��
                    !����ʼʱ�� = rsTmp!����ʼʱ��
                    !�������ʱ�� = rsTmp!�������ʱ��
                    !�������� = rsTmp!��������
                    !������ʼʱ�� = rsTmp!������ʼʱ��
                    !��������ʱ�� = rsTmp!��������ʱ��
                    '�������ͣ�����ʽ
                    strSql = "Select b.���� As ����ʽ, b.Id, b.�������� ��������, a.��Ҫ����" & vbNewLine & _
                            "From ������������ A, ������ĿĿ¼ B" & vbNewLine & _
                            "Where a.������Ŀid = b.Id And a.������ҳid = [1]" & vbNewLine & _
                            "Order By a.���"
                    Set rsOther = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", Val(rsTmp!ID & ""))
                    If Not rsOther.EOF Then
                        !����ʽ = rsOther!����ʽ
                        !�������� = rsOther!��������
                        !����ID = Val(rsOther!ID & "")
                        rsOther.Filter = "��Ҫ����=1"
                        If Not rsOther.EOF Then !����ʽ = rsOther!����ʽ: !�������� = rsOther!��������
                    End If
                    
                    lng��������id = 0
                    '��������
                    strSql = "Select a.Id ��������id, a.��¼����, A.��������, b.���� ����ԭ��, Nvl(D.����,B.����) �������� , a.������Ŀid, d.Id ��������id, a.������, a.�пڷ��� �п�, a.������� ����," & vbNewLine & _
                        "Decode(d.��������, '��', '�ļ�����', '��', '��������', '��', '��������', '��', 'һ������', '�ļ�', '�ļ�����', '����', '��������', '����', '��������', 'һ��', 'һ������', Null) ��������" & vbNewLine & _
                        "From ������������ A, ������ĿĿ¼ B, ������϶��� C, ��������Ŀ¼ D" & vbNewLine & _
                        "Where a.������Ŀid = b.Id And b.Id = c.����id(+) And c.����id = d.Id(+) And a.��¼���� = 2 And a.������ҳid = [1]" & vbNewLine & _
                        "Order By ��������id"
                    Set rsOther = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", Val(rsTmp!ID & ""))
                    If Not rsOther.EOF Then
                        !�������� = rsOther!��������
                        !����ԭ�� = rsOther!����ԭ��
                        lng��������id = Val(rsOther!��������id & "")
                        !�������� = rsOther!��������
                        rsOther.Filter = "������=1"
                        If Not rsOther.EOF Then
                            !�������� = rsOther!��������
                            !����ԭ�� = rsOther!����ԭ��
                            lng��������id = Val(rsOther!��������id & "")
                            !�������� = rsOther!��������
                        End If
                        rsOther.Filter = "��������id=" & lng��������id
                        !��������ID = rsOther!��������ID
                        !�п� = rsOther!�п�
                        !���� = rsOther!����
                        !������Ŀid = rsOther!������Ŀid
                        !�������� = rsOther!��������
                        
                    End If
                    '����ҽ������Ա�Ķ�ȡ
                    If lng��������id <> 0 Then
                        strSql = "Select Distinct ��λ, ����, B.�Ƿ�Ψһ, B.����ҽ��, B.����ҽ��, B.������ʿ" & vbNewLine & _
                                "From ��������ֲ� A, �����λ B" & vbNewLine & _
                                "Where a.��λ = b.���� And ��������id = [1]"
                        Set rsOther = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", lng��������id)
                        If rsOther.EOF Then
                            strSql = "Select Distinct ��λ, ����, B.�Ƿ�Ψһ, B.����ҽ��, B.����ҽ��, B.������ʿ" & vbNewLine & _
                                    "From ����������Ա A, �����λ B" & vbNewLine & _
                                    "Where A.��λ = B.���� And ������ҳid = [1]"
                            Set rsOther = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", Val(rsTmp!ID & ""))
                        End If
                        If Not rsOther.EOF Then
                            rsOther.Filter = " �Ƿ�Ψһ=1  And ����ҽ��=1 "
                            If Not rsOther.EOF Then !����ҽʦ = rsOther!����
                            
                            rsOther.Filter = " ����ҽ��=1 And ��λ like '����%'"
                            If Not rsOther.EOF Then !����ҽʦ = rsOther!����
                            If Len(!����ҽʦ & "") = 0 Then
                                rsOther.Filter = "����ҽ��=1"
                                If Not rsOther.EOF Then !����ҽʦ = rsOther!����
                            End If
                            rsOther.Filter = "��λ='��һ����' OR ��λ='��1����' OR ��λ='�ڢ�����' OR ��λ='����ҽ��һ' OR ��λ='����ҽ��1' OR ��λ='����ҽ����' OR ��λ='����ҽʦһ' OR ��λ='����ҽʦ1' OR ��λ='����ҽʦ��' "
                            If Not rsOther.EOF Then
                                blnDefault = False '��ȡ����һ���֣���Ĭ�϶�ȡ
                                rsOther.Sort = "��λ,����"
                                !��һ���� = rsOther!����
                            End If
                            rsOther.Filter = "��λ='�ڶ�����' OR ��λ='��2����' OR ��λ='�ڢ�����' OR ��λ='����ҽ����' OR ��λ='����ҽ��2' OR ��λ='����ҽ����' OR ��λ='����ҽʦ��' OR ��λ='����ҽʦ2' OR ��λ='����ҽʦ��'"
                            If Not rsOther.EOF Then
                                blnDefault = False '��ȡ���ڶ����֣���Ĭ�϶�ȡ
                                rsOther.Sort = "��λ,����"
                                !�ڶ����� = rsOther!����
                            End If
                            If blnDefault Then
                                rsOther.Filter = " �Ƿ�Ψһ=0  And ����ҽ��=1  "
                                If Not rsOther.EOF Then
                                    rsOther.Sort = "��λ,����"
                                    !��һ���� = rsOther!����
                                    If rsOther.RecordCount <> 1 Then
                                        rsOther.MoveNext
                                        !�ڶ����� = rsOther!����
                                    End If
                                End If
                            End If
                        End If
                    End If
                    'ASA�ּ���NNIS�ּ�
                    strSql = "Select Upper(��������) ��Ŀ, �����ı�" & vbNewLine & _
                            "From ����Ҫ��Ӧ�� A, �������鸽�� B ,���������¼� C" & vbNewLine & _
                            "Where a.������Ŀid = b.������Ŀid And b.�����¼�id=c.ID And b.������ҳid = [1]"
                    Set rsOther = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", Val(!ID & ""))
                    If Not rsOther.EOF Then
                        strTmp = ""
                        rsOther.Filter = " ��Ŀ='ASA�ּ�' "
                        If Not rsOther.EOF Then strTmp = rsOther!�����ı� & ""
                        If Len(strTmp) <> 0 Then strTmp = MidB(strTmp, 1, 20)
                        !asa�ּ� = Decode(Trim(strTmp), "I��", "P1", "II��", "P2", "III��", "P3", "IV��", "P4", "V��", "P5", strTmp)
                        
                        strTmp = ""
                        rsOther.Filter = " ��Ŀ='NNIS�ּ�' "
                        If Not rsOther.EOF Then strTmp = rsOther!�����ı� & ""
                        If Len(strTmp) <> 0 Then strTmp = MidB(strTmp, 1, 20)
                        !NNIS�ּ� = strTmp
                    End If
                    .Update
                    rsTmp.MoveNext
                Wend
            End If
        End If
        
        If blnReadAdvice Then
            strSql = "Select a.Id, NVL(Trunc(E.����ʱ��),NVL(Trunc(a.����ʱ��),Trunc(a.��ʼִ��ʱ��))) ��������, Nvl(D.����,B.����) �������� , NVL(E.����ʱ��,NVL(A.����ʱ��,A.��ʼִ��ʱ��)) ������ʼʱ��, NVL(E.���ʱ��,a.ͣ��ʱ��) ��������ʱ��, a.������Ŀid, d.Id ��������id, b.���� ��������," & vbNewLine & _
                "Decode(d.��������, '��', '�ļ�����', '��', '��������', '��', '��������', '��', 'һ������', '�ļ�', '�ļ�����', '����', '��������', '����', '��������', 'һ��', 'һ������',Null) ��������" & vbNewLine & _
                "From ����ҽ����¼ A, ������ĿĿ¼ B, ������϶��� C, ��������Ŀ¼ D,����ҽ������ E" & vbNewLine & _
                "Where a.������Ŀid = b.Id And A.id = e.ҽ��id And b.Id = c.����id(+) And c.����id = d.Id(+) And a.������� = 'F' And a.����id = [1] And ��ҳid = [2] And" & vbNewLine & _
                "ҽ��״̬ = 8"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", lng����ID, lng��ҳID)
            While Not rsTmp.EOF
                blnDefault = True
                .AddNew
                !ID = rsTmp!ID
                !�������� = rsTmp!��������
                !�������� = rsTmp!��������
                !������ʼʱ�� = rsTmp!������ʼʱ��
                !��������ʱ�� = rsTmp!��������ʱ��
                !��������ID = rsTmp!��������ID
                !�������� = rsTmp!��������
                !�������� = rsTmp!��������
                '������Ϣ��ȡ
                strSql = "Select a.��ʼִ��ʱ�� ����ʼʱ��, a.ͣ��ʱ�� �������ʱ��, b.���� As ����ʽ, b.Id, b.�������� ��������" & vbNewLine & _
                        "From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
                        "Where a.������Ŀid = b.Id And a.������� = 'G' And a.���id = [1]"
                Set rsOther = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", Val(rsTmp!ID & ""))
                If Not rsOther.EOF Then
                    !����ʼʱ�� = rsOther!����ʼʱ��
                    !�������ʱ�� = rsOther!�������ʱ��
                    !����ʽ = rsOther!����ʽ
                    !�������� = rsOther!��������
                End If
                '����ҽ��������ҽ����ȡ
                strSql = "Select ��Ŀ, ���� From ����ҽ������ Where ҽ��id = [1]"
                Set rsOther = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", Val(rsTmp!ID & ""))
                If Not rsOther.EOF Then
                
                    strTmp = ""
                    rsOther.Filter = "��Ŀ='����ҽ��' OR ��Ŀ='����ҽʦ'"
                    If Not rsOther.EOF Then strTmp = rsOther!���� & ""
                    If Len(strTmp) <> 0 Then !����ҽʦ = MidB(strTmp, 1, 20)
                    rsOther.Filter = "��Ŀ='��һ����' OR ��Ŀ='��1����' OR ��Ŀ='�ڢ�����' OR ��Ŀ='����ҽ��һ' OR ��Ŀ='����ҽ��1' OR ��Ŀ='����ҽ����' OR ��Ŀ='����ҽʦһ' OR ��Ŀ='����ҽʦ1' OR ��Ŀ='����ҽʦ��' "
                    If Not rsOther.EOF Then
                        blnDefault = False '��ȡ����һ���֣���Ĭ�϶�ȡ
                        rsOther.Sort = "��Ŀ,����"
                        !��һ���� = MidB(rsOther!���� & "", 1, 20)
                    End If
                    rsOther.Filter = "��Ŀ='�ڶ�����' OR ��Ŀ='��2����' OR ��Ŀ='�ڢ�����' OR ��Ŀ='����ҽ����' OR ��Ŀ='����ҽ��2' OR ��Ŀ='����ҽ����' OR ��Ŀ='����ҽʦ��' OR ��Ŀ='����ҽʦ2' OR ��Ŀ='����ҽʦ��'"
                    If Not rsOther.EOF Then
                        blnDefault = False '��ȡ���ڶ����֣���Ĭ�϶�ȡ
                        rsOther.Sort = "��Ŀ,����"
                        !�ڶ����� = MidB(rsOther!���� & "", 1, 20)
                    End If
                    If blnDefault Then
                        rsOther.Filter = "��Ŀ Like '����*'"
                        If Not rsOther.EOF Then
                            rsOther.Sort = "��Ŀ,����"
                            !��һ���� = MidB(rsOther!���� & "", 1, 20)
                            rsOther.MoveNext
                            If Not rsOther.EOF Then
                                !�ڶ����� = MidB(rsOther!���� & "", 1, 20)
                            End If
                        End If
                    End If
                End If
                .Update
                rsTmp.MoveNext
            Wend
        End If
    End With
    
    Set AutoGetOPSInfo = rsReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetNameByCode(ByVal str��Ϣ�� As String, ByVal str��Ϣֵ As String) As String
'���ܣ�������Ϣֵ����Ϣ����ȡ����

    
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    GetNameByCode = str��Ϣֵ
    
    Select Case str��Ϣ��
        Case "��������"
            strSql = "Select ���� From �ٴ��������� where ����=[1]"
    End Select
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "����������", str��Ϣֵ)
    If rsTmp.RecordCount <> 0 Then
        GetNameByCode = rsTmp.Fields(0).Value
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean, Optional ByVal lngSys As Long) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
'      lngSys=ָ��ϵͳ���ڲ�ģ��Ȩ�ޣ���0�򲻴���Ĭ���ǵ�ǰϵͳ
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    If lngSys = 0 Then lngSys = gclsPros.SysNo
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear
        blnLoad = True
    End If
    On Error GoTo 0
    If blnLoad Then
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    strPrivs = GetPrivFunc(lngSys, lngProg)
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function IsPageNosCodeRule(ByVal ctCode As Code_Type) As Boolean
'����: ��鵵�����Ƿ���ݿ��ұ����Ż��߼�鲡�����Ƿ���˳����
'������intType=4-��鲡�����Ƿ���˳���ţ�5-��鵵�����Ƿ���ݿ��ұ�����
'53638:������,2013-05-10,�����ű������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTrue As Boolean
    On Error GoTo Errhand
    
    strSql = " Select ��Ź��� From ������Ʊ� Where ��Ŀ��� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "������Ʊ�", ctCode)
    If Not rsTmp.EOF Then
        If Val(rsTmp!��Ź��� & "") = IIf(ctCode = CT_������, 3, 0) Then
            blnTrue = True
        End If
    End If
    
    IsPageNosCodeRule = blnTrue
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ValidatePageNos(Optional ByVal blnSave As Boolean) As Boolean
'����: ��֤������ҳ�༭ʱ�Ĳ����ţ������ţ����Ƿ���Ч
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, rsPati As ADODB.Recordset
    Dim strTmp As String, strNo As String, strOutDeptCode As String, strTmpNo As String
    Dim bln˳�� As Boolean, blnDo As Boolean
    Dim strFilter As String
    Dim lngCount As Long
    Dim blnSamePageNo As Boolean

    On Error GoTo Errhand
    '����ID�Լ�סԺ�ŵ��ظ���飬��ǰ�ڽ������ݼ���У����ڱ����̵��ú���ã������������
    '(1)��������ʱ���ù�����֤�ɹ������ǲ���ID�ظ���Ӧ������֤����ID
    '#33282# ʹ��̨����ͬʱ¼�벡��ʱ�����ܻᵼ�²���ID �� סԺ�� �ظ����˴���������ʱ��� ����ID �� סԺ���Ƿ��ظ�������ظ����������µĲ���ID ��סԺ��
    If gclsPros.OpenMode = EM_�������� Then
        If Not gclsPros.IsExistPati Or gclsPros.OnlyPatiInfo Then
            If IsHavePageNos(CT_����ID, True, gclsPros.����ID) Then
                gclsPros.����ID = GetNextNo(CT_����ID)
            End If
            gclsPros.��ҳID = 1
            strNo = Trim(gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text)
            If strNo = "" Then
                strNo = GetNextNo(CT_סԺ��)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text = strNo
            End If
            
            If Not gclsPros.NewInNo And gclsPros.OutFile = "" Then
                If IsHavePageNos(CT_סԺ��, True, strNo, gclsPros.����ID) Then
                    strTmp = GetNextNo(CT_סԺ��)
                    If strNo <> "" Then
                        MsgBox "ԭ" & strNo & "סԺ���Ѿ�����,���������µ�" & strTmp & "סԺ�ţ�", vbInformation, gstrSysName
                        strNo = strTmp
                    End If
                End If
            End If
            gclsPros.InNo = strNo
            gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text = Trim(gclsPros.InNo)
        End If
    ElseIf gclsPros.OpenMode = EM_������ҳ Then
        strNo = Trim(gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text)
        If gclsPros.InNo = strNo Then
            If gclsPros.NewInNo And IsHavePageNos(CT_סԺ��, True, strNo, gclsPros.����ID) Then
                strTmp = GetNextNo(CT_סԺ��)
                If strNo <> "" Then
                    MsgBox "ԭ" & strNo & "סԺ���Ѿ�����,���������µ�" & strTmp & "סԺ�ţ�", vbInformation, gstrSysName
                    strNo = strTmp
                End If
            End If
            gclsPros.InNo = strNo
            gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text = Trim(gclsPros.InNo)
        Else
            If MsgBox("ԭ" & gclsPros.InNo & "סԺ���Ѹı��" & strNo & "סԺ�ţ����ܱ�����ҳ���Ƿ�ԭסԺ�ţ�", vbYesNo, gstrSysName) = vbYes Then
                gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��).Text = gclsPros.InNo
            End If
        End If
    End If
    If gclsPros.NewInNo Then
        If Not (gclsPros.EditPageNo And gclsPros.CurrentForm.txtInfo(GC_������).Text <> "" And blnSave) Then
            gclsPros.CurrentForm.txtInfo(GC_������).Text = gclsPros.InNo
        End If
    End If

    'סԺ�ż��
    If gclsPros.NewInNo Then
        If IsHavePageNos(CT_סԺ��, gclsPros.OpenMode = EM_�༭ Or gclsPros.Is��Ŀ, gclsPros.InNo, gclsPros.����ID) Then
            Call ShowMessage(gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��), "סԺ���Ѿ�����,������ȷ��סԺ��!")
        End If
    End If
    
    '��鲡���Ƿ��Ѿ�����
    If Not gclsPros.EditUnrecive And (gclsPros.OpenMode = EM_�������� Or gclsPros.OpenMode = EM_������ҳ) Then
        strSql = "Select ID From �������ռ�¼ Where ����id = [1] And ��ҳid = [2] And ����ʱ�� Is Not Null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, gclsPros.��ҳID)
        If rsTmp.RecordCount = 0 Then
            Call ShowMessage(gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��), "���˱���סԺ�����һ�û�н��գ����ܽ��б�Ŀ����!")
            Exit Function
        End If
    End If
    If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
    grsDeptInfo.Filter = "ID=" & gclsPros.��Ժ����ID: grsDeptInfo.Sort = "ID"
    If grsDeptInfo.RecordCount > 0 Then
        strOutDeptCode = grsDeptInfo!���� & ""
    End If
                
    '53638:������,2013-05-10,�����ż��
    strNo = Trim(gclsPros.CurrentForm.txtInfo(GC_������).Text)
    If strNo <> "" Then
        If IsHavePageNos(CT_������, True, strNo, gclsPros.����ID, gclsPros.��ҳID) Then
            If gclsPros.UseFileRules Then
                If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
                strTmp = NVL(GetNextNo(CT_������, , strOutDeptCode))
                gclsPros.CurrentForm.txtInfo(GC_������).Text = strTmp
                MsgBox "ԭ" & strNo & "�������Ѿ�����,����ʹ�������ɵ�" & strTmp & "�����ţ�", vbInformation, gstrSysName
            Else
                Call ShowMessage(gclsPros.CurrentForm.txtInfo(GC_������), "������ĵ������Ѿ�����������ʹ��,����������!")
                Exit Function
            End If
        End If
    Else
        If gclsPros.UseFileRules Then
            If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
            strTmp = NVL(GetNextNo(CT_������, , strOutDeptCode))
            gclsPros.CurrentForm.txtInfo(GC_������).Text = strTmp
        End If
    End If
    
    strNo = Trim(gclsPros.CurrentForm.txtInfo(GC_������).Text)
    
    '��ѯ��ͬ�����Ż���ID����Ϣ
    strSql = "Select Nvl(a.����id, [2]) ����id, Nvl(a.��ҳid, [3]) ��ҳid, Nvl(a.������,  [1]) ������, b.����, b.�Ա�, b.���֤��" & vbNewLine & _
                "From (Select ����id, ��ҳid, ������ From סԺ������¼ Where ������ =  [1] Or ����id = [2]) A" & vbNewLine & _
                "Full Join (Select ����id, ��ҳid, ����, �Ա�, ���֤�� From ������Ϣ Where ����id = [2]) B" & vbNewLine & _
                "On a.����id = b.����id"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strNo, gclsPros.����ID, gclsPros.��ҳID)
    '�����Ź���
    '1�����ò��������ݿ��в����ڣ���δ���벡���ţ�����ȡһ��������
    '2�����ò��������ݿ��в����ڣ������벡���ţ���������
    '3�����ò��������ݿ��д��ڣ���δ���벡���ţ����ж��Ƿ�ʹ�õ���������
    '   (1)ʹ�õ��������ţ�����ȡһ��������
    '   (2)δʹ�õ��������ţ���ȡ��һ�εĲ�����
    '4�����ò��������ݿ��д��ڣ������벡���ţ����ж��Ƿ�ʹ�õ���������
    '   (1)ʹ�õ��������ţ���������ͬ������(�������޸�ģʽ�����ų�����סԺ�Ĳ�����)������ȡһ��������
    '   (2)ʹ�õ��������ţ�����������ͬ������(�������޸�ģʽ�����ų�����סԺ�Ĳ�����)����������
    '   (3)δʹ�õ��������ţ���������ͬ�������Ҳ���ID����ͬ�ļ�¼, ���ж��������Ա����֤���Ƿ�һ��
    '         ����һ�£�����ʾ��Щ��Ϣ�в��죬��һ�£���������
    '   (4)δʹ�õ��������ţ������ڲ���ͬ�������Ҳ���ID��ͬ�ļ�¼����������
    '            ����������������ϲ����ڣ���Ϊ���������������д���ģ�
    rsTmp.Filter = "����id=" & gclsPros.����ID
    blnDo = True: strTmpNo = ""
    If rsTmp.EOF Then
        If strNo <> "" Then blnDo = False
    Else
        If strNo = "" Then
            If Not gclsPros.SinPageNo Then
                rsTmp.Filter = "����id=" & gclsPros.����ID & " And ��ҳID<>" & gclsPros.��ҳID
                rsTmp.Sort = "��ҳID"
                If Not rsTmp.EOF Then blnDo = False: strTmpNo = rsTmp!������ & ""
            End If
        Else
            If gclsPros.SinPageNo Then
                rsTmp.Filter = "������='" & strNo & "' "
                If rsTmp.EOF Then
                    blnDo = False
                ElseIf rsTmp.RecordCount = 1 Then
                    blnDo = Not (rsTmp!����ID = gclsPros.����ID And rsTmp!��ҳID = gclsPros.��ҳID)
                    blnSamePageNo = True
                Else
                    blnDo = True
                    blnSamePageNo = True
                End If
            Else
                rsTmp.Filter = "������='" & strNo & "' And ����id<> " & gclsPros.����ID
                rsTmp.Sort = "������,����id,��ҳID"
                If Not rsTmp.EOF Then
                    strTmp = zlStr.NeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_�Ա�).Text)
                    If rsTmp!���� & "" = gclsPros.CurrentForm.txtInfo(GC_����).Text And rsTmp!�Ա� & "" = strTmp And rsTmp!���֤�� & "" <> gclsPros.CurrentForm.cboBaseInfo(BCC_���֤).Text Then
                        MsgBox "���˵Ĳ������ظ�������ͬһ�����ŵ����������������Ա���ͬ�������֤�Ų�ͬ��" & vbCrLf & _
                            "¼��Ĳ��ˣ����֤��[" & gclsPros.CurrentForm.cboBaseInfo(BCC_���֤).Text & "]" & vbCrLf & _
                            "�������ظ��Ĳ��ˣ����֤��[" & rsTmp!���� & "]", vbInformation, gstrSysName
                        Exit Function
                    ElseIf rsTmp!���� & "" <> gclsPros.CurrentForm.txtInfo(GC_����).Text Or rsTmp!�Ա� & "" <> strTmp Then
                        MsgBox "���˵Ĳ������ظ�������ͬһ�����ŵ��������ˣ��������Ա�ͬ��" & vbCrLf & _
                            "¼��Ĳ��ˣ�����[" & gclsPros.CurrentForm.txtInfo(GC_����).Text & "],�Ա�[" & strTmp & "]" & vbCrLf & _
                            "�������ظ��Ĳ��ˣ�����[" & rsTmp!���� & "],�Ա�[" & rsTmp!�Ա� & "]", vbInformation, gstrSysName
                        Exit Function
                    End If
                Else
                    blnDo = False
                End If
            End If
        End If
    End If
    '��ȡ��һ�εĲ�����
    If strTmpNo <> "" And strTmpNo <> strNo Then
        MsgBox "������ͬ����ʹ��ͬһ�����Ų��Ҳ����Ѿ����ڲ�����,���Զ���ȡ���˵����������ţ�", vbInformation, gstrSysName
        blnDo = False
    End If
    If blnDo Then
        If blnSamePageNo Then
            MsgBox "��ǰ�������Ѿ���ʹ����,���Զ���ȡ�����ţ�", vbInformation, gstrSysName
        Else
            MsgBox "��ǰ�����Ų�����Ч������,���Զ���ȡ�����ţ�", vbInformation, gstrSysName
        End If
    End If
    
    bln˳�� = IsPageNosCodeRule(CT_������)
    Do While blnDo
        ' IIf(lngCount = 0, 0, 1)��ֹ����
        strTmpNo = GetNextNo(CT_������, IIf(bln˳�� = True, IIf(lngCount = 0, 0, 1), 0), strOutDeptCode) & ""
        If strTmpNo = "" Then Exit Function
        blnDo = IsHavePageNos(CT_������, True, strTmpNo) '���ڲ����������ѭ��ȥȡ
        If (lngCount >= 100 Or Not bln˳��) And blnDo Then  '���������ѭ�����˳�ѭ��
            strTmpNo = ""
            If blnSave Then
                MsgBox "�Զ���ȡ������ʧ��,�޷����б��棬���ֶ��޸Ĳ����Ż�����ϵ����Ա��", vbInformation, gstrSysName
                ValidatePageNos = False
                Exit Function
            Else
                MsgBox "�Զ���ȡ������ʧ�ܣ�", vbInformation, gstrSysName
                Exit Do
            End If
        End If
        lngCount = lngCount + 1
    Loop
    
    If strTmpNo <> "" Then
        gclsPros.CurrentForm.txtInfo(GC_������).Text = strTmpNo
        If strNo <> "" And strTmpNo <> strNo Then
            MsgBox "ԭ" & strNo & "�������Ѿ�����,����ʹ�������ɵ�" & strTmpNo & "�����ţ�", vbInformation, gstrSysName
        End If
    End If
    ValidatePageNos = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IsHavePageNos(ByVal intType As Integer, ByVal blnCurIn As Boolean, ParamArray arrInput() As Variant) As Boolean
'���ܣ��Ƿ���ں���
'������intType= 0-�Ƿ����סԺ�ţ�
'               1-���������Ƿ�ʹ���˸�סԺ��
'               2-�Ƿ���ڸò���ID,
'               3-����ǳ����˱�����Ժ�⻹�������ط�ʹ���˸ò�����
'      blnCurIn=�Ƿ��Ǳ���סԺ���˵�
'����=True-���ڸú���,False-�����ڸú���
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Select Case intType
        Case CT_סԺ��
            If Not blnCurIn Then
                strSql = "Select 1 From ������ҳ Where סԺ�� = [1] and ��ҳID > 0 "
            Else
                strSql = "Select 1 From ������ҳ Where סԺ�� = [1] And ����id <> [2] and ��ҳID > 0 "
            End If
        Case CT_סԺ��ex
            strSql = "Select ����ID  From ������Ϣ Where סԺ�� = [1]"
        Case CT_����ID
            strSql = "Select 1 From ������ҳ Where ����id = [1]"
        Case CT_������
            strSql = "Select  A.����ID,A.������" & vbNewLine & _
                "From סԺ������¼ A" & vbNewLine & _
                "Where A.������ = [1] And Not Exists" & vbNewLine & _
                " (Select 1 From סԺ������¼ Where ����id = [2] And ��ҳid = [3] And A.����id = ����id And A.��ҳid = ��ҳid) And A.����ID<>[2]"
        Case CT_������
            strSql = "Select 1 From סԺ������¼ Where ������ = [1]"
    
    End Select
    Select Case UBound(arrInput)
        Case 0
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�ж�סԺ������غ������", arrInput(0))
        Case 1
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�ж�סԺ������غ������", arrInput(0), arrInput(1))
        Case 2
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�ж�סԺ������غ������", arrInput(0), arrInput(1), arrInput(2))
    End Select
    
    If intType = CT_סԺ�� Then
        If gclsPros.OutFile = "" Then
            If rsTmp.EOF Then Exit Function
        Else
            If gclsPros.PatiOut.State = adStateOpen Then
                gclsPros.PatiOut.Filter = "סԺ��= " & IIf(Val(arrInput(0)) = 0, 0, arrInput(0))
                If gclsPros.PatiOut.EOF Then Exit Function
            End If
        End If
        IsHavePageNos = True
    Else
        IsHavePageNos = Not rsTmp.EOF
        If IsHavePageNos And intType = CT_סԺ��ex Then
            gclsPros.����ID = Val(rsTmp!����ID & "")
        End If
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiReSeeDoctor() As Boolean
'���ܣ��жϲ��˱����Ƿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String
    Dim strSql As String
    Dim vsTmp As VSFlexGrid
    
    On Error GoTo errH
    
    'ҽ�����������ϴ���ͬ��û��ת������
    strSQL1 = "Select ����ID,ִ���� as ҽ��,ִ�в���ID as ����ID From ���˹Һż�¼ Where ID=[2] And ת�����ID Is Null And �������ID Is Null"
    
    strSQL2 = "Select Max(ID) as ID From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
            " And �Ǽ�ʱ�� =(Select Max(a.�Ǽ�ʱ��) From ���˹Һż�¼ A Where a.����id=[1] And a.��¼����=1 And a.��¼״̬=1 And a.�Ǽ�ʱ��<(Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID=[2])) "
    strSQL2 = "Select ����ID,ִ���� as ҽ��,ִ�в���ID as ����ID From ���˹Һż�¼ Where ID=(" & strSQL2 & ") And ת�����ID Is Null And �������ID Is Null"
    
    strSql = "Select 1 From (" & strSQL1 & ") A,(" & strSQL2 & ") B Where A.����ID=B.����ID And A.ҽ��=B.ҽ�� And A.����ID=B.����ID"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "PatiReSeeDoctor", gclsPros.����ID, gclsPros.��ҳID)
    If rsTmp.EOF Then Exit Function
    
    '��Ҫ������ϴ���ͬ
    Set vsTmp = gclsPros.CurrentForm.vsDiagXY
    With vsTmp
        If .TextMatrix(.FixedRows, DI_�������) <> "" Then
            strSql = "Select Max(ID) as ��ҳID From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
                    " And �Ǽ�ʱ�� =(Select Max(a.�Ǽ�ʱ��) From ���˹Һż�¼ A Where a.����id=[1] And a.��¼����=1 And a.��¼״̬=1 And a.�Ǽ�ʱ��<(Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID=[2])) "
            strSql = "Select 1 From ������ϼ�¼" & _
                " Where ����ID=[1] And ��ҳID=(" & strSql & ")" & _
                " And �������=1 And ��¼��Դ IN(1,3) And ��ϴ���=1" & _
                " And (����ID=[3] And ����ID<>0 Or ���ID=[4] And ���ID<>0 Or �������=[5])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "PatiReSeeDoctor", gclsPros.����ID, gclsPros.��ҳID, _
                Val(.TextMatrix(.FixedRows, DI_����ID)), Val(.TextMatrix(.FixedRows, DI_���ID)), .TextMatrix(.FixedRows, DI_�������))
            If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
        End If
    End With
    
    If gclsPros.Have��ҽ Then
        Set vsTmp = gclsPros.CurrentForm.vsDiagZY
        With vsTmp
            If .TextMatrix(.FixedRows, DI_�������) <> "" Then
                strSql = "Select Max(ID) as ��ҳID From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
                       " And �Ǽ�ʱ�� =(Select Max(a.�Ǽ�ʱ��) From ���˹Һż�¼ A Where a.����id=[1] And a.��¼����=1 And a.��¼״̬=1 And a.�Ǽ�ʱ��<(Select �Ǽ�ʱ�� From ���˹Һż�¼ Where ID=[2])) "
                strSql = "Select 1 From ������ϼ�¼" & _
                    " Where ����ID=[1] And ��ҳID=(" & strSql & ")" & _
                    " And �������=11 And ��¼��Դ IN(1,3) And ��ϴ���=1" & _
                    " And (����ID=[3] And ����ID<>0 Or ���ID=[4] And ���ID<>0 Or �������=[5])"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "PatiReSeeDoctor", gclsPros.����ID, gclsPros.��ҳID, _
                    Val(.TextMatrix(.FixedRows, DI_����ID)), Val(.TextMatrix(.FixedRows, DI_���ID)), .TextMatrix(.FixedRows, DI_�������))
                If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
            End If
        End With
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceIDByDiag(ByVal strҽ��IDs As String, ByVal lng���ID As Long) As String
'���ܣ��������ID��ȡ������ҽ��ID
    Dim strTmp As String
    Dim lngPos As Long
    
    If strҽ��IDs <> "" And gclsPros.AdviceID <> 0 Then
        lngPos = InStr("," & strҽ��IDs & ",", "," & gclsPros.AdviceID & ",")
        If lngPos <= 0 Then
        '��ǰҽ��δ������ǰ���
            strTmp = strҽ��IDs
        Else
            strTmp = Replace("," & strҽ��IDs & ",", "," & gclsPros.AdviceID, "")
            If Len(strTmp) >= 2 Then
                strTmp = Mid(strTmp, 2, Len(strTmp) - 2)
            Else
                strTmp = ""
            End If
        End If
    Else
        strTmp = strҽ��IDs
    End If
    
    With gclsPros.DiagConn
        .Filter = "���ID=" & lng���ID & " And ��ʶID<>" & gclsPros.AplyMark
        .Sort = "��ʶID"
        Do While Not .EOF
            strTmp = strTmp & "," & !��ʶID
            .MoveNext
        Loop
    End With
    
    GetAdviceIDByDiag = strTmp
End Function

Public Function GetPatiRoom(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���Ժ����
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select B.����� as ��Ժ����,c.����� as ��Ժ����  " & vbNewLine & _
            "From ������ҳ A, ��λ״����¼ B,��λ״����¼ C " & vbNewLine & _
            "Where A.����id = [1] And A.��ҳid = [2] And A.��Ժ����id = B.����id(+) And A.��Ժ���� = B.����(+) And A.��ǰ����id = C.����id(+)  And" & vbNewLine & _
            "      A.��Ժ���� = C.����(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng����ID, lng��ҳID)
    Set GetPatiRoom = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInDeptTime(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal strDefault As String) As String
'��ȡ���ʱ��
'strDefault=��ֵʱ�ķ���ֵ
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select ��ʼʱ�� From ���˱䶯��¼" & _
            " Where ����ID=[1] And ��ҳID=[2] And ��ʼԭ�� IN(2,1) And ��ʼʱ�� is Not Null Order by ��ʼԭ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, lng����ID, lng��ҳID)
    If rsTmp.EOF Then
        GetInDeptTime = strDefault
    Else
        If IsNull(rsTmp!��ʼʱ��) Then
            GetInDeptTime = strDefault
        Else
            GetInDeptTime = Format(rsTmp!��ʼʱ��, "yyyy-MM-dd HH:mm")
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ValiAndGet��ҳID() As Boolean
'���ܣ���֤���ȡ��ҳID
'���أ��Ƿ�ɹ�

    Dim lngTmp As Long
    Dim lng���� As Long
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    If gclsPros.OpenMode <> EM_�������� Then
        If GetסԺ����Or��ҳid(gclsPros.����ID, gclsPros.��ҳID, lng����, False) = False Then
            MsgBox "��ȡָ����ҳ�Ĵ���ʧ��,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        If gclsPros.OpenMode = EM_������ҳ Then
            If gclsPros.OnLine Then
                '��ȡ��ǰ��������ҳID
                If GetסԺ����Or��ҳid(gclsPros.����ID, lngTmp, lng����, True) = False Then
                    MsgBox "��ȡָ����������ҳʧ��,���ܼ���!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Function
                End If
                If gclsPros.OnLineNew Then
                    gclsPros.��ҳID = lngTmp + 1
                Else
                    '��Ҫ������ҳ�Ƿ��Ѿ������˲���,���������,���ܽ����ٴν���
                    If lngTmp < gclsPros.��ҳID Then
                        ShowMsgbox "�˲������շ�ϵͳ�в�������Ժ��Ϣ,���ܴ���������"
                        Exit Function
                    End If
                End If
            End If
            strSql = "Select 1 from ������ҳ where ����id=[1] and ��ҳid=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������Ϣ", gclsPros.����ID, gclsPros.��ҳID)
            If rsTmp.EOF Then
                lng���� = lng���� + 1
            End If
        End If
        gclsPros.CurrentForm.txtSpecificInfo(SLC_��Ժ����).Text = lng����  '��ҳID
    End If
    ValiAndGet��ҳID = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ReadPatPricture(ByVal lng����ID As Long, ByRef imgPatient As Image, Optional ByRef strFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '������lng����ID=��ȡָ�����˵���Ƭ
    '           imgPatient=��Ƭ����λ��
    '           strFile=��Ƭ�ı���·��
    '74421,������,2014-07-04,��ȡ������Ƭ��Ϣ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Errhand
    imgPatient.Picture = Nothing
    strFile = ""
    strFile = sys.Readlob(gclsPros.SysNo, 27, lng����ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = LoadPicture(strFile)
        ReadPatPricture = True
        Kill strFile
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ExistInList(ByVal strסԺ�� As String, ByVal blnMessage As Boolean, Optional blnOnlyCheck As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '����:���סԺ���Ƿ������סԺ�����嵥
    '����:
    '     blnOnlyCheck-��Ϊ����ʱ���
    '����:���ڷ���true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strDate As String
    Dim strTmp As String
    
    If gclsPros.InputOutList = False Then ExistInList = True: Exit Function
    On Error GoTo errH
    
    strSql = "" & _
        "   Select A.����,A.����,B.����,B.ID " & _
        "   From ��Ժ�����嵥 A,���ű� B  " & _
        "   Where A.����ID=B.ID and A.סԺ��= [1] " & _
        "   Order by A.���� desc "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strסԺ��)
    
    If rsTmp.RecordCount = 0 Then
        If blnMessage = True Then
            '106826:�ڲ������ձ༭��ʱ�򣬲����Ǹò����Ƿ��Ѿ�����סԺ�ձ��嵥
            ExistInList = True
        End If
        zlControl.TxtSelAll gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
        Exit Function
    End If

    '���⣺22163��22577
    '��鷽ʽ��
    '1.�����HIS����ģʽ���򲻽��з�д����
    '2.����Ƕ�����װģʽ����ֻ������ʱ���ŷ�д���ݣ����򲻷�д����
    '3.����Ǳ�������ã��򲻼������
    strDate = Format(rsTmp!���� & "", "yyyy-mm-dd hh:mm")
    If strDate = "" Then strDate = "1989-01-01 " & Format(Now, "hh:mm")
    gclsPros.OutTime = strDate
    gclsPros.��Ժ����ID = Val(NVL(rsTmp!ID)) '���ó�Ժ���һ��Զ��жϿ�������
    strTmp = rsTmp!���� & ""
    If blnOnlyCheck Or gclsPros.ShareMedRec Or gclsPros.OpenMode = EM_�༭ Then
        ExistInList = True
        Exit Function
    End If
    gclsPros.CurrentForm.txtInfo(GC_����).Text = rsTmp!���� & ""
    gclsPros.CurrentForm.txtInfo(GC_��Ժ����).Text = rsTmp!���� & ""
     gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��).Text = strDate   '������Ժʱ����Զ�����סԺ����
    ExistInList = True
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Select�ⲿ��ҳid(strסԺ�� As String, Optional int���� As Integer = 0) As Integer
    Dim arrDate() As String
    Dim str��ȡ���� As String
    Dim vRect As RECT
    Dim rsTmp As New ADODB.Recordset
    Dim intTmp As Integer
    Dim strIfdate As String, blnALLPati As Boolean
    Dim objList As ListItem
    Dim lngTemp As Long
    Dim strKEY As String ''��¼ѡ���е�KEY
    Dim int��ҳid As Integer ''��¼��ҳid
    Dim objTxtסԺ�� As TextBox
    
    ReDim arrDate(2)
    arrDate(0) = zlDatabase.GetPara("��ʼ����", gclsPros.SysNo, gclsPros.Module)
    arrDate(1) = zlDatabase.GetPara("��������", gclsPros.SysNo, gclsPros.Module)
    If arrDate(0) = "" Or arrDate(1) = "" Then
        arrDate(0) = "": arrDate(1) = ""
        blnALLPati = True
    Else
        arrDate(0) = Format(arrDate(0), "yyyy-mm-dd")
        arrDate(1) = Format(arrDate(1), "yyyy-mm-dd")
    End If
    If Not blnALLPati Then blnALLPati = Val(zlDatabase.GetPara("��ȡ���г�Ժ����", gclsPros.SysNo, gclsPros.Module)) = 1
    
    If Val(zlDatabase.GetPara("��ȡ24Сʱ�ڳ�Ժ����", gclsPros.SysNo, gclsPros.Module)) <> 1 Then
        If gclsPros.OutFile = "" Then
            str��ȡ���� = " And (B.��Ժ����-B.��Ժ����)*24>=24"
        Else
            str��ȡ���� = "סԺʱ��>=24"
        End If
    End If
    
    If gclsPros.EditUnrecive = False And gclsPros.OutFile = "" Then
        str��ȡ���� = str��ȡ���� & " And E.����ʱ�� IS NOT NULL"
    End If
    If Not blnALLPati Then
        If gclsPros.OutFile = "" Then
            strIfdate = " And B.��Ժ���� Between Trunc(To_Date('" & arrDate(0) & "','yyyy-mm-dd')) And Trunc(To_Date('" & arrDate(1) & "','yyyy-mm-dd'))+1-1/24/60/60"
        Else
            strIfdate = " ��Ժ���� >= #" & Format(arrDate(0), "yyyy-mm-dd 00:00:00") & "# and ��Ժ���� <= #" & Format(arrDate(1), "yyyy-mm-dd 23:59:59") & "#"
        End If
    End If
    Set objTxtסԺ�� = gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
    vRect = zlControl.GetControlRect(objTxtסԺ��.hwnd)
    If str��ȡ���� <> "" Then
        strIfdate = IIf(strIfdate = "", "", strIfdate & " and ") & str��ȡ����
        If strIfdate <> "" Then
            If IsNumeric(objTxtסԺ��.Text) Then
                 strIfdate = strIfdate & " and " & "סԺ��=" & Trim(objTxtסԺ��.Text) & " and סԺ����>" & int����
            End If
        Else
            strIfdate = "סԺ��=" & Trim(objTxtסԺ��.Text) & " and סԺ����>" & int����
        End If
    End If
    With frmPageMedRecNOSel
        .Top = vRect.Top + 300
        .Left = vRect.Left
        strKEY = .ShowMe(gclsPros.CurrentForm, gclsPros.PatiOut, strIfdate)
        objTxtסԺ��.Text = Split(strKEY, "_")(0)
        If Val(objTxtסԺ��.Text) = 0 Then
            objTxtסԺ��.Text = ""
            Exit Function
        Else
            gclsPros.InNo = objTxtסԺ��.Text
        End If
        int��ҳid = Split(strKEY, "_")(1)
        gclsPros.PatiOut.Filter = " סԺ��=" & objTxtסԺ��.Text & " and סԺ����=" & int��ҳid
        gclsPros.��ҳID = int��ҳid
        LoadDataFromOutFile (objTxtסԺ��.Text)
        Select�ⲿ��ҳid = int��ҳid
    End With

End Function

Public Function Select��ҳID(lng����ID As Long, Optional int���� As Integer = 0) As Integer
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim objTxtסԺ�� As TextBox
    Dim strSql As String
    
    Set objTxtסԺ�� = gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
    vRect = zlControl.GetControlRect(objTxtסԺ��.hwnd)
    '���˺�:��ҳid��ʵ�ʵ�סԺ������һ��,��Ϊʵ�ʵ�סԺ�������������۲���
    '����26488 by lesfeng 2010-03-18 ����
    '����30659 by lesfeng 2010-06-04 ������� ����Ȩ��
    
    '56002:������,2012-11-23,���б�Ŀ�Ĳ���Ӧ��ֻ����������Ժ�Ĳ��ˣ�����������=0
    '39906:������,2013-05-07,��Ӳ������ձ�־
    If gclsPros.OnLine Then
        strSql = "" & _
            "   Select a.����ID||'-'||a.��ҳid as ID,a.סԺ��,b.����,B.�Ա�,TO_char(a.��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժ����, " & _
            "           TO_char(a.��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժ����,C.���� ��Ժ����," & _
            "         '��'||Zl_��ȡסԺ��������ҳid(a.����id,a.��ҳid,0)||'��' as סԺ����,decode(D.�������,null,'��',0,'��','��') As ����,Decode(E.����ʱ��, NULL, '��', '��') AS ����" & _
            "   from ������ҳ a,������Ϣ b,������� D,���ű� C,�������ռ�¼ E " & _
            "   Where a.����id=b.����id And A.����ID=E.����ID(+) ANd A.��ҳID=E.��ҳID(+) and a.��Ŀ���� is null and nvl(a.��������,0)=0  and a.��Ժ���� is not null " & _
            "           and a.����id=[1]  and a.��ҳid>[2] and A.����id = D.����id(+)  And D.����(+)=2 And A.��Ժ����ID=C.ID(+) " & _
            "   order by ��Ժ���� asc"
    Else
        strSql = "" & _
        "   Select a.����ID||'-'||a.��ҳid as ID,a.סԺ��,B.����,B.�Ա�,TO_char(a.��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժ����, " & _
        "           TO_char(a.��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժ����,C.���� ��Ժ����," & _
        "         '��'||Zl_��ȡסԺ��������ҳid(a.����id,a.��ҳid,0)||'��' as סԺ����,Decode(e.����ʱ��, NULL, '��', '��') AS ����" & _
        "   from ������ҳ a,������Ϣ b,���ű� C,�������ռ�¼ E   " & _
        "   Where a.����id=b.����id And A.����ID=E.����ID(+) and A.��ҳID=E.��ҳID(+) and a.��Ŀ���� is null and nvl(a.��������,0)=0  and a.��Ժ���� is not null " & _
        "           and a.����id=[1]  and a.��ҳid>[2] And A.��Ժ����ID=C.ID(+) " & _
        "   order by ��Ժ���� asc"
    End If
       
   '���˺�:���۲��˲��ܽ�������
   ' Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "����סԺ����", , , , , , True, lmx, lmy, 300, , , True)
    Set rsTemp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "����סԺ����", False, "", "", False, False, True, vRect.Left, vRect.Top, 300, blnCancel, False, True, lng����ID, int����)
    
    If blnCancel Then Select��ҳID = 0: Exit Function
    If rsTemp Is Nothing Then Select��ҳID = 0: Exit Function
    If rsTemp.State = 0 Then Select��ҳID = 0: Exit Function
    
    If rsTemp.RecordCount > 0 Then
        Select��ҳID = CInt(Mid(rsTemp!ID, InStr(rsTemp!ID, "-") + 1))
    Else
        Select��ҳID = 0
    End If
End Function

Public Function GetDeptCode(ByVal lngDeptID As Long) As String
'51446,������,2012-08-02
'���ܣ����ݿ���ID��ȡ���ұ���
    Dim strCode As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo Errhand
    If lngDeptID <= 0 Then Exit Function
    strSql = "select ���� From ���ű� where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ұ���", lngDeptID)
    If rsTemp.RecordCount > 0 Then strCode = NVL(rsTemp!����)
    
    GetDeptCode = strCode
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub OpenExtraData()
    Dim cnAccess As New ADODB.Connection
    
    If gclsPros.OutFile = "" Or gclsPros.OpenMode = EM_���� Or gclsPros.OpenMode = EM_�༭ Then Exit Sub
        
    On Error Resume Next
    cnAccess.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gclsPros.OutFile & ";Persist Security Info=False"
    If Err <> 0 Then
        MsgBox "�򲻿��ɻ�ȡ������Ϣ�����������ļ���", vbInformation, gstrSysName
        Exit Sub
    End If
    If gclsPros.PatiOut.State = 1 Then gclsPros.PatiOut.Close
    If gclsPros.FeesOut.State = 1 Then gclsPros.FeesOut.Close
    
    '�������������
    gclsPros.PatiOut.Open "select *,clng((��Ժ����-��Ժ����)*24) as סԺʱ�� from ������ҳ order by סԺ��,סԺ����", cnAccess, adOpenStatic, adLockReadOnly
    gclsPros.FeesOut.Open "select * from ���˷���", cnAccess, adOpenStatic, adLockReadOnly
    
End Sub

Public Sub GetDaysFromLast()
'���ܣ���ȡ���ϴ���Ժ������
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim blnGet As Boolean, lngDay As Long, strSec As String
    
    If gclsPros.��ҳID > 1 And gclsPros.InTime <> "" Then
        If gclsPros.FuncType = f������ҳ And gclsPros.MedPageSandard = ST_����ʡ��׼ Then
            blnGet = gclsPros.CurrentForm.cboBaseInfo(BCC_���ϴ�סԺʱ��).ListIndex = -1
        ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
            blnGet = gclsPros.CurrentForm.txtSpecificInfo(SLC_���ϴ�סԺʱ��).Text = ""
        End If
        If blnGet Then
            strSql = "select (To_Date('" & Format(gclsPros.InTime, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')-��Ժ����) ʱ��� from ������ҳ where ����ID=[1] And ��ҳid =[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "����һ����ס��Ժʱ��", gclsPros.����ID, gclsPros.��ҳID - 1)
            If Not rsTmp.EOF Then
                lngDay = Val(NVL(rsTmp!ʱ���))
                If lngDay >= 2 And lngDay <= 15 Then
                    strSec = "2-15��"
                ElseIf lngDay >= 16 And lngDay <= 31 Then
                    strSec = "16-31��"
                ElseIf lngDay > 31 Then
                    strSec = "��31��"
                Else
                    strSec = "����"
                End If
                lngDay = IIf(lngDay < 1, 1, lngDay)
                If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                    gclsPros.CurrentForm.txtSpecificInfo(SLC_���ϴ�סԺʱ��).Text = lngDay
                Else
                    Call Cbo.SeekIndex(gclsPros.CurrentForm.cboBaseInfo(BCC_���ϴ�סԺʱ��), strSec)
                End If
            End If
        End If
    End If
End Sub


