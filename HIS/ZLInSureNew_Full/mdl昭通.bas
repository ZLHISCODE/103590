Attribute VB_Name = "mdl��ͨ"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

Public blnConnIsOpen As Boolean, gcn��ͨ As New ADODB.Connection
Private mstr���� As String, mstr��Ժ״̬ As String, mcur���� As Currency, mcurͳ�� As Currency
'�����Ҫ���Է�ҩƷҲ����룬Ȼ���ϴ������ĵȴ���ˣ���¼�����ĿʱҲ�����Ƿ�ͨ���������������Է�ҩƷ�����Աѡ����ͨ/������Ҫ����Ϣ
'�޸ĺţ�5825
'����:
'1.�޸�I100-3��׼ҩƷĿ¼�ṹ�Է���ʡͳһ��׼(p4).
'2.I110-1�걨ҩƷ��,������ʡ��׼Ŀ¼�ļ�����,�걨������Ҫ����������������ʹ��.��Ŀ¼ҩƷ��Ȼ��ԭ���ķ�������(��Ҫ�걨����������)
'3.����I110-4����ǰ̨�޸�������¼

'����ʱʹ��
Dim gArrayTest() As String

Public Function ҽ����ʼ��_��ͨ(Optional ByVal bln����Ա��� As Boolean = True) As Boolean
    Dim rsTemp As New ADODB.Recordset, str����ֵ As String, lngPort As Long, strServer As String, _
        strSN As String, strDataSource As String
    On Error GoTo errHandle
    If blnConnIsOpen Then
        ҽ����ʼ��_��ͨ = True
        Exit Function
    End If
    
    strDataSource = Mid(gcnOracle.ConnectionString, InStr(UCase(gcnOracle.ConnectionString), "SERVER=") + 7)
    strDataSource = Left(strDataSource, InStr(strDataSource, """;") - 1)
    
    On Error Resume Next
    If gcn��ͨ.State = 1 Then gcn��ͨ.Close
    gcn��ͨ.ConnectionString = "Provider=MSDAORA.1;Password=his;User ID=ybuser;Data Source=" & strDataSource & ";Persist Security Info=True"
    gcn��ͨ.CursorLocation = adUseClient
    gcn��ͨ.Open
    If Err.Number <> 0 Then
        MsgBox "�����м����ݿ�ʧ��", vbInformation, "ҽ����ʼ��"
        Exit Function
    End If
    On Error GoTo errHandle
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��ͨ)
        
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "��ͨ���֤"
                strSN = str����ֵ
            Case "��ͨ������"
                strServer = str����ֵ
            Case "��ͨ�˿ں�"
                lngPort = CLng(str����ֵ)
        End Select
        rsTemp.MoveNext
    Loop
    
    If strSN = "" Or strServer = "" Or lngPort = 0 Then
        MsgBox "���ղ������ò��������������ӵ�ҽ��", vbInformation, "ҽ����ʼ��"
        Exit Function
    End If
    
    If frmConn��ͨ.ConnCenter(strServer, lngPort, strSN, IIf(bln����Ա���, UserInfo.ID, 0)) = False Then Exit Function
    
    gstrSQL = "Select * From ������� Where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ����ʼ��", TYPE_��ͨ)
    gstrҽԺ���� = Trim(rsTemp!ҽԺ����)
    
    blnConnIsOpen = True
    ҽ����ʼ��_��ͨ = True
    Exit Function
errHandle:
    WriteInfo "��ʼ����������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ֹ_��ͨ() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    On Error GoTo errHandle
    
    Call frmConn��ͨ.ConnClose
    ҽ����ֹ_��ͨ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_��ͨ(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, TYPE_��ͨ)
    
    If rsTemp.EOF Then
        �������_��ͨ = 0
    Else
        �������_��ͨ = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
    End If
End Function

Public Function ��ݱ�ʶ_��ͨ(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim strPatiInfo As String, cur��� As Currency
    Dim arr, datCurr As Date, str����� As String
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    If bytType = 1 Then
        strPatiInfo = frmIdentify��ͨסԺ.GetPatient(bytType, mstr��Ժ״̬)
    Else
        strPatiInfo = frmIdentify��ͨ.GetPatient(bytType)
    End If

    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
        lng����ID = BuildPatiInfo(bytType, strPatiInfo, lng����ID, TYPE_��ͨ)

        '���ظ�ʽ:�м���벡��ID
        If bytType = 1 Then
            strPatiInfo = frmIdentify��ͨסԺ.mstrPatient & lng����ID & ";" & frmIdentify��ͨסԺ.mstrOther
        Else
            strPatiInfo = frmIdentify��ͨ.mstrPatient & lng����ID & ";" & frmIdentify��ͨ.mstrOther
        End If
    Else
        ��ݱ�ʶ_��ͨ = ""
        MsgBox "ҽ��������Ϣ��ȡʧ��", vbInformation, gstrSysName
        Exit Function
    End If
    ��ݱ�ʶ_��ͨ = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_��ͨ = ""
End Function

Public Function �������_��ͨ(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency, Optional ByRef strAdvance As String) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
'        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ����
'        ����һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset, rsCheck As New ADODB.Recordset
    Dim str����Ա As String, datCurr As Date, str������ As String, str���㷽ʽ As String
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, curCount As Currency
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur���� As Currency, cur����ͳ���޶� As Currency, strTemp As String
    Dim cur���ͳ���޶� As Currency, cur�����Ը� As Currency, cur��� As Currency
    Dim cur�������� As Currency, cur���Ը� As Currency, strPara As String
    Dim blnOld As Boolean
    Dim blnBalance As Boolean
    
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "Select * From ������ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng����ID = rs��ϸ!����ID
    
    If rs��ϸ.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If
    curCount = 0
    
    WriteInfo vbCrLf & "��ʼ�������"
    '��֯������ϸ
    While Not rs��ϸ.EOF
        gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        strTemp = rsTemp!���
        
        gstrSQL = "Select * From ����֧����Ŀ Where ����=103 And �շ�ϸĿID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����֧����Ŀ", CLng(rs��ϸ!�շ�ϸĿID))
        If rsTemp.EOF Then             '���û�ж���������ʹ��
            gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�շ�ϸĿ", CLng(rs��ϸ!�շ�ϸĿID))
            Err.Raise 9000, gstrSysName, "��Ŀ[" & rsTemp!���� & "]û�ж�Ӧ��ҽ����Ŀ,���Ƚ��ж���"
            WriteInfo "��Ŀ[" & rsTemp!���� & "]û�ж�Ӧ��ҽ����Ŀ,�˳��������"
            Exit Function
        ElseIf Nvl(rsTemp!��ע) <> "������" And InStr(" 5 6 7 ", " " & strTemp & " ") > 0 Then
            strTemp = Nvl(rsTemp!��ע, "δ����")
            
            '�����ʡĿ¼�ļס�������Ŀ������Ҫ���������־
            gstrSQL = "Select lb From tab_syml where dm='" & rsTemp!��Ŀ���� & "'"
            Call OpenRecordset_OtherBase(rsCheck, "�����ʡĿ¼�ļס�������Ŀ������Ҫ���������־", gstrSQL, gcn��ͨ)
            If rsCheck.RecordCount <> 0 Then
                If rsCheck!lb = 15 Then
                    gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�շ�ϸĿ", CLng(rs��ϸ!�շ�ϸĿID))
                    Err.Raise 9000, gstrSysName, "��Ŀ[" & rsTemp!���� & "]" & strTemp & "������ʹ��"
                    WriteInfo "��Ŀ[" & rsTemp!���� & "]" & strTemp & "������ʹ��"
                    Exit Function
                End If
            End If
        End If
        
        '�ֺŷָ���,���ŷָ���
        If InStr(" 5 6 7 ", " " & strTemp & " ") > 0 Then
            strTemp = rs��ϸ!�շ�ϸĿID
'        ElseIf strTemp = "7" Then
'            '��ȡ�в�ҩ�ķ�������
'            If Trim(Nvl(rs��ϸ!ժҪ)) = "" Or InStr(1, ",UZY01,UZY02,UZY03,", "," & UCase(Nvl(rs��ϸ!ժҪ)) & ",") = 0 Then
'                strTemp = GetItemInfo_��ͨ(1, lng����ID, rs��ϸ!�շ�ϸĿID, Nvl(rs��ϸ!ժҪ), rs��ϸ!NO)
'            Else
'                strTemp = Nvl(rs��ϸ!ժҪ)
'            End If
        Else
            strTemp = rsTemp!��Ŀ����
        End If
        strPara = strPara & ";" & strTemp & "," & rs��ϸ!���� * rs��ϸ!���� & "," & _
            Round(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), 4) & "," & rs��ϸ!ʵ�ս��
        
        curCount = curCount + rs��ϸ!ʵ�ս��
        rs��ϸ.MoveNext
    Wend
    
    If strPara <> "" Then strPara = Mid(strPara, 2)             'ȥ����ͷ�ķֺ�
    
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҽ��������Ϣ", lng����ID)
    
    strPara = rsTemp!���� & vbTab & Nvl(rsTemp!����, " ") & vbTab & strPara
    WriteInfo "���״��ݲ���:" & strPara
    If frmConn��ͨ.Execute("I200", 1, strPara, "���ڽ���ҽ������,���Ժ�......") = False Then Exit Function
    If frmConn��ͨ.Query(0, 1) = False Then Exit Function
    strPara = frmConn��ͨ.strReturnInfo
    WriteInfo "���׷�������:" & strPara
    If strPara = "" Then
        Err.Raise 9000, gstrSysName, "�������ݸ�ʽ����", vbInformation, "�������"
        Exit Function
    End If
    
    blnBalance = True
    str������ = Split(strPara, vbTab)(0)
    cur�����ʻ� = Split(strPara, vbTab)(2)
        
    If cur�����ʻ� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
    End If
    
    WriteInfo "���׽��:" & Mid(str���㷽ʽ, 3)
    '�������
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        #If gverControl < 2 Then
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
        #Else
            strAdvance = str���㷽ʽ
            gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
    End If
    #If gverControl < 2 Then
        blnOld = True
        frm������Ϣ.ShowME lng����ID
    #End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��ͨ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
            "," & TYPE_��ͨ & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & _
            cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_��ͨ & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",NULL,NULL," & _
            cur�������� & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
            cur�����ʻ� & ",NULL,NULL,NULL,'" & str������ & "'" & IIf(blnOld, "", ",1") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    �������_��ͨ = True
    
    WriteInfo "�������﷢Ʊ����"
    Call SaveOutExse(str������)
    
    WriteInfo "������ｻ��"
    Exit Function
errHandle:
    If blnBalance Then
        ErrMsgBox "��ͨ��ҽ�����߽���ǰ��������շѵ��ݳ�����������Ϣ���¼��" & vbCrLf & _
            "�����ţ�" & str������ & "�������ʻ���" & cur�����ʻ�, vbInformation, "�������"
    Else
        ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
        Err.Clear
        Exit Function
    End If
    WriteInfo "��������:" & Err.Description
End Function

Public Function ����������_��ͨ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, strPara As String
    Dim curCount As Currency, str����Ա As String
    Dim datCurr As Date
    
    WriteInfo vbCrLf & "��ʼ�������"
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    str����Ա = UserInfo.���
    
    gstrSQL = "Select ����ID,���ʽ��,����Ա��� From ������ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "�Ҳ������ݵ���ϸ��¼,���ܽ��г���"
        Exit Function
    End If
    lng����ID = rsTemp!����ID
    If rsTemp!����Ա��� <> str����Ա Then
        Err.Raise 9000, gstrSysName, "ҽ���涨������ִ�б������������Ĳ���Ա���г���"
        Exit Function
    End If
    
    Do Until rsTemp.EOF
        curCount = curCount + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��ͨ, lng����ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    If IsNull(rsTemp!��ע) Then
        Err.Raise 9000, gstrSysName, "�õ��ݵľ����Ŷ�ʧ���������ϡ�"
        Exit Function
    End If
    str������ = rsTemp!��ע
    
    strPara = str������ & vbTab & rsTemp!�����ʻ�֧��
    WriteInfo "���״��ݲ���:" & strPara
    
    '���ýӿ�������
    If Not frmConn��ͨ.Execute("I220", 0, strPara, "���ڽ���ҽ������,���Ժ�......") Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��ͨ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_��ͨ & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_��ͨ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curCount * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        cur�����ʻ� * -1 & ",NULL,NULL,NULL,'" & str������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    ����������_��ͨ = True
    WriteInfo "����������"
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
    WriteInfo "��������:" & Err.Description
End Function

Public Function ��Ժ�Ǽ�_��ͨ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset, str���� As String, datCurr As Date, strInNote As String, _
        strPara As String, strסԺ��� As String

    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    WriteInfo vbCrLf & "��ʼ��Ժ�Ǽ�"
    
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,D.���� as ���ұ���,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽ�_��ͨ = False
        Exit Function
    End If
    
    '��ȡ��Ժ��ϣ����ֱ��룩
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID)  '��Ժ���
'    If strInNote <> "" Then
'        strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
'    End If
    WriteInfo "ȡ����Ժ��ϣ�" & strInNote
    
    strPara = rsTemp!���� & vbTab & rsTemp!���� & vbTab & mstr��Ժ״̬ & vbTab & _
              Nvl(rsTemp!סԺ����, " ") & vbTab & ToVarchar(Nvl(rsTemp!��Ժ����, "0"), 10) & vbTab & strInNote
    
'    gstrSQL = "Select sum(���) From ����Ԥ����¼ Where ����ID=" & lng����id & " And ��ҳID=" & lng��ҳID
'    Call OpenRecordset(rsTemp, gstrSysName)
    strPara = strPara & vbTab & "0"         ' Nvl(rsTemp(0), 0)
    WriteInfo "���״��ݲ���:" & strPara
    
    '���ýӿڽ��еǼ�
    If Not frmConn��ͨ.Execute("I300", 1, Replace(strPara, vbTab & vbTab, vbTab & " " & vbTab), "���ڽ���ҽ������,���Ժ�......") Then Exit Function
    If frmConn��ͨ.Query(0, 1) = False Then Exit Function
    strסԺ��� = Replace(frmConn��ͨ.strReturnInfo, " ", "")
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��ͨ & ",'˳���','''" & strסԺ��� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ���")
    
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��ͨ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_��ͨ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_��ͨ = False
End Function

Public Function ��Ժ�Ǽǳ���_��ͨ(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ�Ǽǳ�����Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim str˳��� As String
    Dim rsTemp As New ADODB.Recordset, str���� As String, datCurr As Date, strInNote As String

    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    WriteInfo vbCrLf & "��ʼ������Ժ"
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵�ҽ����Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽǳ���_��ͨ = False
        Exit Function
    End If
    If Nvl(Replace(rsTemp!˳���, " ", ""), "") = "" Then
        MsgBox "ȡҽ������סԺ��Ŵ���", vbInformation, gstrSysName
        Exit Function
    End If
    str˳��� = Nvl(Replace(rsTemp!˳���, " ", ""))
    
    'ֻҪ���ڲ��˷��ü�¼�������������Ժ,ֻ�ܳ�Ժ�Ǽ�
    gstrSQL = "SELECT 1 FROM סԺ���ü�¼ WHERE ����ID=[1] AND ��ҳID=[2] AND ROWNUM<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ֻҪ���ڲ��˷��ü�¼�������������Ժ,ֻ�ܳ�Ժ�Ǽ�", lng����ID, lng��ҳID)
    If rsTemp.RecordCount <> 0 Then
        MsgBox "�ò����ѷ�������,���ܰ�������Ժ,ֻ�ܰ����Ժ!", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "���״��ݲ���:" & str˳���
    '���ýӿڽ��еǼ�
    If Not frmConn��ͨ.Execute("I305", 0, str˳���, "���ڽ���ҽ������,���Ժ�......") Then Exit Function

     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��ͨ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽǳ���_��ͨ = True
    WriteInfo "������Ժ���"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽǳ���_��ͨ = False
End Function

Public Function ת��ת��_��ͨ(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ���ת��ת����Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strInNote As String, rsTemp As New ADODB.Recordset, strPara As String, strסԺ��� As String
    
    '������˵������Ϣ
    On Error GoTo errHandle
    WriteInfo vbCrLf & "��ʼ��Ժ��Ϣ�䶯"
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If Nvl(Replace(rsTemp!˳���, " ", ""), "") = "" Then
        MsgBox "ȡ������Ժ��Ŵ���", vbInformation, gstrSysName
        Exit Function
    End If
    strסԺ��� = Replace(rsTemp!˳���, " ", "")
    
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,D.���� as ���ұ���,A.��Ժ����,A.סԺҽʦ,C.˳���," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ת��ת��_��ͨ = False
        Exit Function
    End If
    
    '��ȡ��Ժ���
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, True, False) '��Ժ���
    If strInNote = "" Then strInNote = "��ͨ"
    WriteInfo "��ȡ��Ժ��ϣ�" & strInNote
    
    strPara = strסԺ��� & vbTab & Nvl(rsTemp!���ұ���, "0") & vbTab & ToVarchar(Nvl(rsTemp!��Ժ����, "0"), 10) & vbTab & strInNote
    WriteInfo "���״��ݲ���:" & strPara
    
    '���ýӿڽ��еǼ�
    If Not frmConn��ͨ.Execute("I309", 1, strPara, "���ڽ���ҽ������,���Ժ�......") Then Exit Function
     
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��ͨ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ת��ת��_��ͨ = True
    WriteInfo "��Ժ��Ϣ������"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ת��ת��_��ͨ = False
End Function

Public Function סԺ�������_��ͨ(rs��ϸ As Recordset, lng����ID As Long, strҽ���� As String) As String
'������rsDetail     ������ϸ(����)
'    ����ID         adBigInt, 19, adFldIsNullable
'    �շ����       adVarChar, 2, adFldIsNullable
'    �վݷ�Ŀ       adVarChar, 20, adFldIsNullable
'    ���㵥λ       adVarChar, 6, adFldIsNullable
'    ������         adVarChar, 20, adFldIsNullable
'    �շ�ϸĿID     adBigInt, 19, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ʵ�ս��       adSingle, 15, adFldIsNullable
'    ͳ����       adSingle, 15, adFldIsNullable
'    ����֧������ID adBigInt, 19, adFldIsNullable
'    �Ƿ�ҽ��       adBigInt, 19, adFldIsNullable
'    ժҪ           adVarChar, 200, adFldIsNullable
'    �Ƿ���       adBigInt, 19, adFldIsNullable
'    str���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    Dim rsTemp As New ADODB.Recordset, str���� As String, datCurr As Date, _
        strTemp As String, cur����֧�� As Currency, cur�ֽ� As Currency, curTemp As Currency, _
        cur����Ա���� As Currency, lng��ҳID As Long, cur��ͳ�� As Currency, cur����ͳ�� As Currency
    Dim cur�����ܶ� As Currency, strסԺ�� As String, str���㷽ʽ As String, strReturn() As String
    Dim strMessage As String
    
    On Error GoTo errHandle
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�з��ã����ܽ���Ԥ���㡣", vbInformation, gstrSysName
        Exit Function
    End If
    cur�����ܶ� = 0
    While Not rs��ϸ.EOF
        cur�����ܶ� = cur�����ܶ� + rs��ϸ!���
        rs��ϸ.MoveNext
    Wend
    WriteInfo "��ʼԤ����"
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ!����ID
    gstrSQL = "Select max(��ҳid) from ������ҳ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng��ҳID = rsTemp(0)
    
    If ���ʴ���_��ͨ("", 0, strMessage, lng����ID) = False Then
        MsgBox strMessage, vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select * From �����ʻ� Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    strסԺ�� = Nvl(Replace(rsTemp!˳���, " ", ""), "")
    If strסԺ�� = "" Then
        MsgBox "���ܻ�ȡ����סԺ˳��ţ����ܽ��н���", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "���״��ݲ�����" & strסԺ��
    If frmConn��ͨ.Execute("I361", 5, strסԺ��, "���ڶ�ȡ���˷�����Ϣ......") = False Then Exit Function
    If frmConn��ͨ.Query(0, 1) = False Then Exit Function
    WriteInfo "���أ�" & frmConn��ͨ.strReturnInfo
    strReturn = Split(frmConn��ͨ.strReturnInfo, vbTab)
    cur����ͳ�� = CCur(strReturn(0))
    cur��ͳ�� = CCur(strReturn(1))
    
    mcurͳ�� = cur����ͳ�� + cur��ͳ��
    
    curTemp = �������_��ͨ(lng����ID)
    gstrSQL = "Select nvl(sum(���),0) From ����Ԥ����¼ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If cur�����ܶ� - mcurͳ�� > rsTemp(0) Then
        cur����֧�� = cur�����ܶ� - mcurͳ�� - rsTemp(0)
    Else
        cur����֧�� = 0
    End If
    If cur����֧�� > curTemp Then cur����֧�� = curTemp
    If cur����֧�� < 0 Then cur����֧�� = 0
    mcur���� = cur����֧��
    str���㷽ʽ = "�����ʻ�;" & cur����֧�� & ";1"
    If cur����ͳ�� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ <> "", "|", "") & "����ͳ��;" & cur����ͳ�� & ";0"
    If cur��ͳ�� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ <> "", "|", "") & "��ͳ��;" & cur��ͳ�� & ";0"
    
    סԺ�������_��ͨ = str���㷽ʽ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_��ͨ(lng����ID As Long, lng����ID As Long) As Boolean
'���ܣ���סԺ���ý�����ϸ���ݲ��ҽ��н���
'���סԺ������ϸ����ʧ�ܣ���ֱ�ӽ������������غ���ʧ��
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim rs��ͨ As New ADODB.Recordset, strTemp As String, lng��ҳID As Long, strסԺ�� As String
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, datCurr As Date, strPara As String
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur����֧�� As Currency, cur��� As Currency, cur�������� As Currency, curͳ��֧�� As Currency
    Dim cur������֧�� As Currency, cur����Ա���� As Currency, str������ˮ�� As String, str���� As String
    Dim str��Ժ���� As String, str��Ժ��� As String
    On Error GoTo errHandle
    WriteInfo vbCrLf & "��ʼ��Ժ����"
    gstrSQL = "Select * from �����ʻ� Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    strסԺ�� = Replace(rsTemp!˳���, " ", "")
    
    gstrSQL = "Select max(��ҳID) From ������ҳ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng��ҳID = rsTemp(0)
    gstrSQL = "Select A.��Ժ����,B.���� From ������ҳ A,���ű� B Where A.��Ժ����ID=B.ID And ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    
    str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, True) '��Ժ���
    If str��Ժ��� = "" Then str��Ժ��� = " "
    
    WriteInfo "�޸ĳ�Ժ��Ϣ"
    strTemp = strסԺ�� & vbTab & rsTemp!���� & vbTab & ToVarchar(rsTemp!��Ժ����, 10) & vbTab & str��Ժ���
    WriteInfo "���״��ݲ�����" & strTemp
    frmConn��ͨ.Execute "I309", 1, strTemp, "���ڽ���ҽ������......"
    
    gstrSQL = "Select * From ����Ԥ����¼ Where ����id=[1] And ���㷽ʽ='�����ʻ�'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.EOF Then
        cur����֧�� = 0
    Else
        cur����֧�� = Nvl(rsTemp!��Ԥ��, 0)
    End If
    
    Screen.MousePointer = 0
    If cur����֧�� <> 0 Then
        If frmIdentify��ͨ.GetPatient(0) = "" Then Exit Function
        str���� = Split(frmIdentify��ͨ.mstrPatient, ";")(0)
        Unload frmIdentify��ͨ
    End If
    
    gstrSQL = "Select * from �����ʻ� Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    strסԺ�� = Replace(rsTemp!˳���, " ", "")
    
    If cur����֧�� <> 0 Then
        strPara = strסԺ�� & vbTab & str���� & vbTab & IIf(Nvl(rsTemp!����, "") = "", " ", rsTemp!����) & vbTab & cur����֧��
    Else
        strPara = strסԺ�� & vbTab & " " & vbTab & IIf(Nvl(rsTemp!����, "") = "", " ", rsTemp!����) & vbTab & cur����֧��
    End If
    WriteInfo "���״��ݲ�����" & strPara
    If frmConn��ͨ.Execute("I340", 1, strPara, "���ڽ��г�Ժ����......") = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��ͨ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
            "," & TYPE_��ͨ & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + mcurͳ�� & "," & intסԺ�����ۼ� & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ͨҽ��")
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��ͨ & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� + cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + mcurͳ�� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur�������� & ",0,0,NULL," & mcurͳ�� & ",NULL,NULL," & _
            cur����֧�� & ",NULL,NULL,NULL,'" & strסԺ�� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ͨҽ��")
    
    סԺ����_��ͨ = True
    
    Call SaveInExse(strסԺ��)
    
    WriteInfo "��Ժ����ɹ�"
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_��ͨ(lng����ID As Long) As Boolean
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim rs��ͨ As New ADODB.Recordset, strTemp As String, lng��ҳID As Long, strסԺ�� As String
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, datCurr As Date
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur����֧�� As Currency, cur��� As Currency, cur�������� As Currency, curͳ��֧�� As Currency
    Dim cur������֧�� As Currency, cur����Ա���� As Currency, str������ˮ�� As String, str���� As String
    Dim lng����ID As Long
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select * From ����Ԥ����¼ Where ����id=[1] And ���㷽ʽ='�����ʻ�'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.EOF Then
        cur����֧�� = 0
    Else
        cur����֧�� = Nvl(rsTemp!��Ԥ��, 0)
    End If
    
    gstrSQL = "Select * from ���ս����¼ Where ��¼id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    strסԺ�� = rsTemp!��ע
    lng����ID = rsTemp!����ID
    
    If frmConn��ͨ.Execute("I345", 0, strסԺ��, "����ȡ����Ժ����......") = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��ͨ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
            "," & TYPE_��ͨ & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� - cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� - rsTemp!ͳ�ﱨ����� & "," & intסԺ�����ۼ� & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ͨҽ��")
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��ͨ & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� - cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� - rsTemp!ͳ�ﱨ����� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur�������� & ",0,0,NULL," & 0 - rsTemp!ͳ�ﱨ����� & ",NULL,NULL," & _
            0 - cur����֧�� & ",NULL,NULL,NULL,'" & strסԺ�� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ͨҽ��")
    
    סԺ�������_��ͨ = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ���ʴ���_��ͨ(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset, lng��ҳID As Long, strסԺ��� As String, strPara As String, _
        rs��ϸ As New ADODB.Recordset, strID() As String, lngLoop As Long, strRetu() As String, _
        str���� As String, blnAll As Boolean, cur���� As Currency
    Dim strTemp As String, strMessage As String
    Dim int��ҩ��־ As Integer
    
    On Error GoTo errHandle
    blnAll = True
    If lng����ID <> 0 Then
        gstrSQL = "Select Max(��ҳID) From ������ҳ Where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
        lng��ҳID = rsTemp(0)
    End If
    
    If str���ݺ� <> "" Then
        gstrSQL = " Select A.* From סԺ���ü�¼ A,�����ʻ� B " & _
                  " Where A.�����־=2 And A.��¼״̬<>0 And A.��¼״̬<>3 And A.��¼״̬<>2 And nvl(A.���ӱ�־,0)<>9 " & _
                  " and nvl(A.ʵ�ս��,0)<>0 and A.��¼����=[1] and A.NO=[2]" & _
                  " and A.����ID=B.����ID and B.����=[2]" & _
                  " order by A.����ID,A.��ҳID,A.���"
        Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, int����, str���ݺ�, TYPE_��ͨ)
    Else
        gstrSQL = " Select * From סԺ���ü�¼ " & _
                  " Where �����־=2 And ��¼״̬<>0 And ��¼״̬<>3 And ��¼״̬<>2 And nvl(���ӱ�־,0)<>9 " & _
                  " and nvl(ʵ�ս��,0)<>0 and ����id=[1] And ��ҳid=[2]" & _
                  " order by ��ҳID,���"
        Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    End If
    
    If rs��ϸ.EOF Then
        MsgBox "û����Ҫ�ϴ��Ĳ��˷���", vbInformation, gstrSysName
        ���ʴ���_��ͨ = True
        Exit Function
    End If
    
    lng����ID = 0
    strPara = ""
    ReDim strID(rs��ϸ.RecordCount)
    lngLoop = 0
    While Not rs��ϸ.EOF
        If lng����ID <> rs��ϸ!����ID Then
            lng����ID = rs��ϸ!����ID
            gstrSQL = "Select * From �����ʻ� Where ����id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
            If rsTemp.EOF Then
                str��Ϣ = "ȡ����ҽ����Ϣʱʧ��"
                Exit Function
            End If
            If Nvl(Replace(rsTemp!˳���, " ", ""), "") = "" Then
                str��Ϣ = "ȡҽ������סԺ���ʧ��"
                Exit Function
            End If
            strסԺ��� = Replace(rsTemp!˳���, " ", "")
            '���ýӿڽ��з��õǼ�Ԥ��
'            If Not frmConn��ͨ.Execute("I320", 0, strסԺ���, "���ڽ���ҽ������,���Ժ�......") Then Exit Function
            
            If strPara <> "" Then   '����в�����ʾ�ò��������ݣ����ýӿ��ϴ�
                strPara = Left(strPara, Len(strPara) - 1)
                frmConn��ͨ.Execute "I320", 1, strPara, "���ڽ���ҽ������,���Ժ�......"
            
                If frmConn��ͨ.Query(0, 1, "���ڶ�ȡ��ϸ�ϴ����صĽ����") = False Then Exit Function
                frmConn��ͨ.strReturnInfo = Mid(frmConn��ͨ.strReturnInfo, InStr(1, frmConn��ͨ.strReturnInfo, vbTab) + 1)
                frmConn��ͨ.mlngRows = UBound(Split(frmConn��ͨ.strReturnInfo, ";"))
                For lngLoop = 1 To frmConn��ͨ.mlngRows
                    If Split(Split(frmConn��ͨ.strReturnInfo, ";")(lngLoop - 1), ",")(2) = "0" Then
                        gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & strID(lngLoop - 1) & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
                    Else
                        gstrSQL = "Select B.���� From סԺ���ü�¼ A,�շ�ϸĿ B Where A.�շ�ϸĿID=B.ID And A.ID=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(strID(lngLoop - 1)))
                        If rsTemp.RecordCount <> 0 Then
                            strMessage = strMessage & "��Ŀ[" & rsTemp(0) & "]�ϴ�ʧ�ܣ�������Ϣ��" & Split(Split(frmConn��ͨ.strReturnInfo, ";")(lngLoop - 1), ",")(3) & Chr(13) & Chr(10)
                        Else
                            MsgBox "û���ҵ���¼,IDΪ:" & strID(lngLoop - 1), vbInformation, gstrSysName
                            Exit Function
                        End If
                        blnAll = False
                    End If
                Next
            
            End If
            strPara = strסԺ��� & vbTab & "0" & vbTab
            lngLoop = 0
        End If
        
        gstrSQL = "Select * From ����֧����Ŀ Where �շ�ϸĿID=[1] And ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID), TYPE_��ͨ)
        If rsTemp.EOF Then
            gstrSQL = "Select ����||���� From �շ�ϸĿ Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
            str��Ϣ = "��¼����δ�������Ŀ[" & rsTemp.Fields(0).Value & "�������ϴ���ҽ�����ģ�"
            Exit Function
        End If
        
        '������۴����޼ۣ����޼�ִ��
        cur���� = Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), "0.##")
'        If Nvl(rsTemp!��ע, "0") <> "0" Then
'            If cur���� > CCur(rsTemp!��ע) Then
'                cur���� = CCur(rsTemp!��ע)
'            End If
'        End If
        int��ҩ��־ = 0
        strTemp = rsTemp!��Ŀ����
        gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        Select Case rsTemp!���
        Case "5", "6", "7"
            '��ȡ��Ŀ¼��ҩƷ��ҩ��־
            If Get�Է�ҩƷ(UCase(strTemp)) Then
                If Trim(Nvl(rs��ϸ!ժҪ)) = "" Or InStr(1, ",��ͨ,����,������Ҫ���󲡣�,", "," & UCase(Nvl(rs��ϸ!ժҪ)) & ",") = 0 Then
                    strTemp = GetItemInfo_��ͨ(2, lng����ID, rs��ϸ!�շ�ϸĿID, Nvl(rs��ϸ!ժҪ), rs��ϸ!NO)
                Else
                    strTemp = Nvl(rs��ϸ!ժҪ)
                End If
                int��ҩ��־ = IIf(strTemp = "����", 1, IIf(strTemp = "������Ҫ���󲡣�", 2, 0))
            End If
            strTemp = rsTemp!ID
'        Case "7"
'            '��ȡ�в�ҩ�ķ�������
'            If Trim(Nvl(rs��ϸ!ժҪ)) = "" Or InStr(1, ",UZY01,UZY02,UZY03,", "," & UCase(Nvl(rs��ϸ!ժҪ)) & ",") = 0 Then
'                strTemp = GetItemInfo_��ͨ(2, lng����ID, rs��ϸ!�շ�ϸĿID, Nvl(rs��ϸ!ժҪ), rs��ϸ!NO)
'            Else
'                strTemp = Nvl(rs��ϸ!ժҪ)
'            End If
        End Select
            
        strPara = strPara & Format(rs��ϸ!����ʱ��, "yyyymmdd") & "," & strTemp & "," & _
            Format(rs��ϸ!���� * rs��ϸ!����, "0.##") & "," & Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * _
            rs��ϸ!����), "0.##") & "," & int��ҩ��־ & ";"
        
        strID(lngLoop) = rs��ϸ!ID
        lngLoop = lngLoop + 1
        rs��ϸ.MoveNext
    Wend
    
    If strPara <> "" Then           '�������һ��ѭ���Ľ��
        gstrSQL = "Select * From �����ʻ� Where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
        If rsTemp.EOF Then
            str��Ϣ = "ȡ����ҽ����Ϣʱʧ��"
            Exit Function
        End If
        If Nvl(Replace(rsTemp!˳���, " ", ""), "") = "" Then
            str��Ϣ = "ȡҽ������סԺ���ʧ��"
            Exit Function
        End If
        strסԺ��� = Replace(rsTemp!˳���, " ", "")
        '���ýӿڽ��з��õǼ�Ԥ��
'        If Not frmConn��ͨ.Execute("I320", 0, strסԺ���, "���ڽ���ҽ������,���Ժ�......") Then Exit Function
        
        If strPara <> "" Then   '����в�����ʾ�ò��������ݣ����ýӿ��ϴ�
            strPara = Left(strPara, Len(strPara) - 1)
            frmConn��ͨ.Execute "I320", 1, strPara, "���ڽ���ҽ������,���Ժ�......"
            
            If frmConn��ͨ.Query(0, 1, "���ڶ�ȡ��ϸ�ϴ����صĽ����") = False Then Exit Function
            frmConn��ͨ.strReturnInfo = Mid(frmConn��ͨ.strReturnInfo, InStr(1, frmConn��ͨ.strReturnInfo, vbTab) + 1)
            frmConn��ͨ.mlngRows = UBound(Split(frmConn��ͨ.strReturnInfo, ";"))
            For lngLoop = 1 To frmConn��ͨ.mlngRows
                If Split(Split(frmConn��ͨ.strReturnInfo, ";")(lngLoop - 1), ",")(2) = "0" Then
                    gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & strID(lngLoop - 1) & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
                Else
                    gstrSQL = "Select B.���� From סԺ���ü�¼ A,�շ�ϸĿ B Where A.�շ�ϸĿID=B.ID And A.ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(strID(lngLoop - 1)))
                    If rsTemp.RecordCount <> 0 Then
                        strMessage = strMessage & "��Ŀ[" & rsTemp(0) & "]�ϴ�ʧ�ܣ�������Ϣ��" & Split(Split(frmConn��ͨ.strReturnInfo, ";")(lngLoop - 1), ",")(3) & Chr(13) & Chr(10)
                    Else
                        MsgBox "û���ҵ���¼,IDΪ:" & strID(lngLoop - 1), vbInformation, gstrSysName
                        Exit Function
                    End If
                    blnAll = False
                End If
            Next
        End If
        strPara = strסԺ��� & vbTab & "0" & vbTab
    End If
    Screen.MousePointer = vbDefault
    If strMessage <> "" Then
        str��Ϣ = strMessage
        Exit Function
    End If
    
    ���ʴ���_��ͨ = True
    Exit Function
errHandle:
    If ErrCenter() Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_��ͨ(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo errHandle
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��ͨ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽǳ���_��ͨ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ��Ժ�Ǽǳ���_��ͨ = False
End Function

Public Function ��Ժ�Ǽ�_��ͨ(lng����ID As Long, lng��ҳID As Long) As Boolean
    On Error GoTo errHandle
    Dim BLN�޷���Ժ As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    '����Ƿ����δ�����
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        '������δ�����,�ټ���Ƿ�����
        gstrSQL = "SELECT 1 FROM סԺ���ü�¼ WHERE ����ID IS NOT NULL AND  ����ID=[1] AND ��ҳID=[2] AND ROWNUM<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�����", lng����ID, lng��ҳID)
        If rsTemp.RecordCount = 0 Then
            '˵���ò������޷���Ժ,û�����Ҳû��δ�����
            gstrSQL = "Select * From �����ʻ� Where ����ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
            If Nvl(Replace(rsTemp!˳���, " ", "")) = "" Then
                MsgBox "ȡҽ������סԺ��Ŵ���", vbInformation, gstrSysName
                Exit Function
            End If
            
            WriteInfo "���״��ݲ���:" & Replace(rsTemp!˳���, " ", "")
            '���ýӿڽ��еǼ�
            If Not frmConn��ͨ.Execute("I305", 0, rsTemp!˳���, "���ڽ���ҽ������,���Ժ�......") Then Exit Function
        End If
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��ͨ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_��ͨ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ��Ժ�Ǽ�_��ͨ = False
End Function

Public Function GetItemInfo_��ͨ(ByVal intType As Integer, ByVal lng����ID As Long, ByVal lngϸĿID As Long, _
    ByVal strժҪ As String, Optional ByVal str��ע As String = "") As String
    '�в�ҩ���룬��ѡ���������
    '���в�ҩ�࣬�ҷ�ҽ��ҩƷ����ѡ���Ƿ�󲡻�������ҩ
    'intType-��������(0-ҽ��,1-�����շ�,2-סԺ����)
    Dim bln�в�ҩ As Boolean
    Dim rsTemp As New ADODB.Recordset
    '�봦���в�ҩ:���ࣺuzy01�����ࣺuzy02����ҽ����uzy03
    gstrSQL = "Select ��� From �շ�ϸĿ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ����в�ҩ", lngϸĿID)
    If rsTemp.RecordCount = 0 Then Exit Function
    If InStr(1, "5,6", rsTemp!���) <> 0 Then
        gstrSQL = "Select ��Ŀ����,��ע From ����֧����Ŀ Where �շ�ϸĿID=[1] And ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ����", lngϸĿID, TYPE_��ͨ)
        If rsTemp.RecordCount = 0 Then Exit Function
        If Not Get�Է�ҩƷ(UCase(rsTemp!��Ŀ����)) Then Exit Function
        If Nvl(rsTemp!��ע) <> "������" Then Exit Function
'    ElseIf rsTemp!��� = "7" Then
'        bln�в�ҩ = True
    Else
        Exit Function
    End If
    
    GetItemInfo_��ͨ = frm��ͨ_��Ŀ��Ϣ.ShowME(intType, lng����ID, lngϸĿID, strժҪ, str��ע, bln�в�ҩ)
End Function

Public Function CheckInsureItem_��ͨ(ByVal lng�շ�ϸĿID As Long) As Boolean
    '�����Ŀ�Ƿ�ͨ�����ĵ�����
    '�����ҩƷ��Ŀ������ʡĿ¼���Ǽס����࣬�򲻱ؽ���������飬��ֱ��ʹ��
    Dim str��ע As String, str��Ŀ���� As String
    Dim rsTemp As New ADODB.Recordset

    gstrSQL = "Select ��Ŀ����,��ע From ����֧����Ŀ Where ����=[1] And �շ�ϸĿID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����Ŀ�Ƿ�ͨ�����ĵ�����", TYPE_��ͨ, lng�շ�ϸĿID)
    If rsTemp.RecordCount <> 0 Then
        str��ע = Nvl(rsTemp!��ע)
        str��Ŀ���� = rsTemp!��Ŀ����

        If Not Get�Է�ҩƷ(UCase(str��Ŀ����)) Then Exit Function
        If str��ע <> "������" Then
            MsgBox "����Ŀ��δͨ����������������ʹ�ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
End Function

Public Sub SaveOutExse(ByVal str��ˮ�� As String)
    Dim str�������� As String, str������ϸ As String
    Dim arrData
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    'ɾ������ˮ�ŵ��������
    WriteInfo "ɾ������ˮ�ŵ��������"
    gcn��ͨ.Execute "Delete ������ϸ Where ��ˮ��='" & str��ˮ�� & "'"
    gcn��ͨ.Execute "Delete �������� Where ��ˮ��='" & str��ˮ�� & "'"
    
    '��ȡ������������
    str�������� = str��ˮ��
    WriteInfo "(��ȡ������������)���״��ݲ���:" & str��������
    If frmConn��ͨ.Execute("I280", 0, str��������, "���ڽ���ҽ������,���Ժ�......") = False Then Exit Sub
    If frmConn��ͨ.Query(0, 1) = False Then Exit Sub
    str�������� = frmConn��ͨ.strReturnInfo
    WriteInfo "���׷�������:" & str��������
    If str�������� = "" Then
        MsgBox "�������ݸ�ʽ����", vbInformation, "�������"
        Exit Sub
    End If
    '����������������,����֧�������,�ڲ���ֵ�м������շ�ʱ��(٦����)
    arrData = Split(str��������, vbTab)
    gstrSQL = " Insert Into ��������(��ˮ��,����֤��,����,֧��ǰ���,�ʻ�֧��,֧�������,�ֽ�֧��,����ʱ��,�շѵ�ַ,�绰,���׽��,����ʱ��1)" & _
        " Values ('" & str��ˮ�� & "','" & arrData(0) & "','" & arrData(1) & "'," & Val(arrData(2)) & "," & Val(arrData(3)) & "," & Val(arrData(4)) & "," & _
        "'" & arrData(5) & "','" & arrData(6) & "','" & arrData(7) & "','" & Val(arrData(8)) & "','" & arrData(9) & "','" & arrData(6) & "')"
    gcn��ͨ.Execute gstrSQL
    
    '��ȡ������ϸ����
    str������ϸ = str��ˮ��
    WriteInfo "(��ȡ������ϸ����)���״��ݲ���:" & str������ϸ
    If frmConn��ͨ.Execute("I280", 1, str������ϸ, "���ڽ���ҽ������,���Ժ�......") = False Then Exit Sub
    For lngLoop = 1 To frmConn��ͨ.mlngRows
        If frmConn��ͨ.Query(lngLoop - 1, 1, "���ڸ�������(" & lngLoop & "/" & (frmConn��ͨ.mlngRows) & ")......") = False Then Exit Sub
        str������ϸ = frmConn��ͨ.strReturnInfo
        arrData = Split(str������ϸ, vbTab)
        gstrSQL = " Insert Into ������ϸ(��ˮ��,����,��λ,���,����,����,����,���)" & _
            " Values ('" & str��ˮ�� & "','" & arrData(0) & "','" & arrData(1) & "','" & arrData(2) & "','" & arrData(3) & "'," & _
            Val(arrData(4)) & "," & Val(arrData(5)) & "," & Val(arrData(6)) & ")"
        gcn��ͨ.Execute gstrSQL
    Next
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub SaveInExse(ByVal str��ˮ�� As String)
    Dim strסԺ���� As String, strסԺ��ϸ As String
    Dim arrData
    Dim lngLoop As Long
    
    'ɾ������ˮ�ŵ��������
    WriteInfo "ɾ������ˮ�ŵ��������"
    On Error Resume Next
    gcn��ͨ.Execute "Delete סԺ��ϸ Where ��ˮ��='" & str��ˮ�� & "'"
    gcn��ͨ.Execute "Delete סԺ���� Where ��ˮ��='" & str��ˮ�� & "'"
    
    On Error GoTo errHand
    
    '��ȡ������������
    strסԺ���� = str��ˮ��
    WriteInfo "(��ȡ������������)���״��ݲ���:" & strסԺ����
    If frmConn��ͨ.Execute("I348", 0, strסԺ����, "���ڽ���ҽ������,���Ժ�......") = False Then Exit Sub
    If frmConn��ͨ.Query(0, 1) = False Then Exit Sub
    strסԺ���� = frmConn��ͨ.strReturnInfo
    WriteInfo "���׷�������:" & strסԺ����
    If strסԺ���� = "" Then
        MsgBox "�������ݸ�ʽ����", vbInformation, "סԺ����"
        Exit Sub
    End If
    '����������������
    arrData = Split(strסԺ����, vbTab)
    gstrSQL = " Insert Into סԺ����(��ˮ��,��λ����,���,��Ա״̬,����֤��,����,�Ա�,����,��Ժ����,��Ժ����,�Ʊ�,����,��������,Ѻ��,�ʻ�֧��,�ɷ�����)" & _
        " Values ('" & str��ˮ�� & "','" & arrData(0) & "','" & arrData(1) & "','" & arrData(2) & "','" & arrData(3) & "','" & arrData(4) & "'," & _
        "'" & arrData(5) & "'," & Val(arrData(6)) & ",'" & arrData(7) & "','" & arrData(8) & "','" & arrData(9) & "'," & _
        "'" & arrData(10) & "','" & arrData(11) & "'," & Val(arrData(12)) & "," & Val(arrData(13)) & ",'" & arrData(14) & "')"
    gcn��ͨ.Execute gstrSQL
    
    '��ȡ������ϸ����
    strסԺ��ϸ = str��ˮ�� & vbTab & "1"
    WriteInfo "(��ȡ������ϸ����)���״��ݲ���:" & strסԺ��ϸ
    If frmConn��ͨ.Execute("I348", 1, strסԺ��ϸ, "���ڽ���ҽ������,���Ժ�......") = False Then Exit Sub
    For lngLoop = 1 To frmConn��ͨ.mlngRows
        If frmConn��ͨ.Query(lngLoop - 1, 1, "���ڸ�������(" & lngLoop & "/" & (frmConn��ͨ.mlngRows) & ")......") = False Then Exit Sub
        strסԺ��ϸ = frmConn��ͨ.strReturnInfo
        arrData = Split(strסԺ��ϸ, vbTab)
        gstrSQL = " Insert Into סԺ��ϸ(��ˮ��,����,����,���˸����ȸ�����,���˸�����׼����,���˸������,ͳ����𸺵�)" & _
            " Values ('" & str��ˮ�� & "','" & arrData(0) & "'," & Val(arrData(1)) & "," & Val(arrData(2)) & "," & Val(arrData(3)) & "," & _
            Val(arrData(4)) & "," & Val(arrData(5)) & ")"
        gcn��ͨ.Execute gstrSQL
    Next
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function Get�Է�ҩƷ(ByVal strCode As String) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '���ơ����������ҩƷ�����ؼ�
    
    gstrSQL = "Select lb From tab_syml where upper(dm)='" & strCode & "'"
    Call OpenRecordset_OtherBase(rsCheck, "���ҩƷ����", gstrSQL, gcn��ͨ)
    If rsCheck.RecordCount = 0 Then Exit Function   '˵����������Ŀ
    If rsCheck!lb <> 15 Then Exit Function           '˵���Ǽ��ࡢ����ҩƷ
    Get�Է�ҩƷ = True
End Function
