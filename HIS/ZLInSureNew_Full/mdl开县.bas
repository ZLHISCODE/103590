Attribute VB_Name = "mdl����"
Option Explicit
Public gcur�ʻ���� As Currency                '����ר��,��������ʻ����

Public gcnOracle_���� As ADODB.Connection
Private mblnInit As Boolean

Public Function ҽ������_����() As Boolean
'���ܣ� �÷������ڹ����Ӧ�ò���������������ҽ�����ݷ����������Ӵ�
'���أ��ӿ����óɹ�������true�����򣬷���false
    
    Dim strConn As String
    
    If frmSet�ɶ�.ShowSet(TYPE_����) = False Then
        Exit Function
    End If
    
    strConn = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LHConnectionStrINg"), "dsn=lhyb;uid=sa;pwd=;")
    '���½�����ҽ���������Ĺ�������
    If gcnOracle_���� Is Nothing Then
                ҽ������_���� = True
    Else
            If gcnOracle_����.State = adStateClosed Then
                On Error Resume Next
                gcnOracle_����.Open strConn
                If Err = 0 Then
                    ҽ������_���� = True
                Else
                    Err.Clear
                End If
            Else
                ҽ������_���� = True
            End If
    End If
End Function


Public Function ҽ����ʼ��_����() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    If mblnInit Then
        ҽ����ʼ��_���� = True
        Exit Function
    End If
    '������ҽ���������Ĺ�������
    Dim strCnn As String
    strCnn = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("LHConnectionStrINg"), "")
    Err = 0
    On Error Resume Next
    If gcnOracle_���� Is Nothing Then
        Set gcnOracle_���� = New ADODB.Connection
    End If
    With gcnOracle_����
        If .State = adStateOpen Then .Close
        .ConnectionString = strCnn
        .Open
        If Err <> 0 Then
            MsgBox "���ܽ�����ҽ�������������ӣ��޷�ִ��ҽ������", vbExclamation, gstrSysName
            Exit Function
        End If
    End With
    '�������ҽ������ı��Ƿ���
    gstrSQL = "select * from RCPT_TAB,DIAG_REC "
    gcnOracle_����.Execute gstrSQL, 1
    If Err <> 0 Then
        MsgBox "RCPT_TAB���DIAG_REC��û�н������޷�ִ��ҽ������", vbExclamation, gstrSysName
        Exit Function
    End If
    mblnInit = True
    ҽ����ʼ��_���� = True
End Function
Public Function ҽ����ֹ_����() As Boolean
    mblnInit = False
    If gcnOracle_����.State = 1 Then
        gcnOracle_����.Close
    End If
    
    ҽ����ֹ_���� = True
End Function


Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������strSelfNO-���˱�ţ�ˢ���õ���strSelfPwd-�������룻bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmlhIDentified As frmIdentify����
     
    Set frmlhIDentified = New frmIdentify����
    
    With frmlhIDentified
        .Tag = bytType
        .Show 1
        'New:0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
        ��ݱ�ʶ_���� = .strPatiInfo
        
        If ��ݱ�ʶ_���� <> "" Then
            '�������˵�����Ϣ�������ʽ��
            '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
            '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
            '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)

            lng����ID = BuildPatiInfo(bytType, ��ݱ�ʶ_���� & ";;;;;;;;;;;;;;;;", lng����ID, TYPE_����)
            '���ظ�ʽ:�м���벡��ID
            ��ݱ�ʶ_���� = ��ݱ�ʶ_���� & ";" & lng����ID & ";;;;;;;;;;;;;;;;"
        End If
        
    End With
    Unload frmlhIDentified
    
    
End Function

Public Function �������_����() As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����Ľ��
    �������_���� = gcur�ʻ����
End Function

Public Function �������_����(lng����ID As Long) As Boolean
'�ù���Ŀǰδʹ�ã��������ʱͨ�����ô�����ϸ�ﵽ
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    
    Dim rsPay As New Recordset
    Dim strReptNo As String
    Dim strInterCode As String
    Dim rsList As New ADODB.Recordset
    Dim lngCount As Long, lng����ID As Long
    
    Dim cur����֧�� As Currency, cur�������� As Currency
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curDate As Date
    
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    �������_���� = False
    Err = 0
    On Error GoTo errHand:
    gstrSQL = _
        "Select a.ID,a.NO,a.�Ǽ�ʱ��,D.��Ŀ����,c.����,c.����,a.�շ�ϸĿid,a.������ as ҽ��,a.����,a.����ID," & _
        "       (nvl(a.����,1)*nvl(a.����,1)) as ����,a.ʵ�ս��/(nvl(a.����,1)*nvl(a.����,1)) as ����,a.ʵ�ս�� as ���ʽ�� " & _
        " From ������ü�¼ a,���ű� c, " & _
        "       (   Select * From ����֧����Ŀ Where ����=" & TYPE_���� & ") D" & _
        " Where  nvl(a.ʵ�ս��,0)<>0 and a.�շ�ϸĿID=d.�շ�ϸĿid(+) and  a.ִ�в���ID=c.id(+) and Nvl(a.���ӱ�־,0)<>9 And a.����ID=[1]" & _
        " order by a.ID"
    Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    With rsList
        If .RecordCount = 0 Then
            Err.Raise 9000, gstrSysName, "û����д�շѼ�¼", vbExclamation, gstrSysName
            Exit Function
        End If

        strReptNo = !NO
        strInterCode = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("intercode"), 713)
        strInterCode = IIf(IsNumeric(strInterCode), strInterCode, "0")
        lng����ID = !����ID

        lngCount = 0
        Do While Not .EOF
            lngCount = lngCount + 1
            '    hosp_id_IN      rcpt_tab.hosp_ID%type,
            '    rcpt_no_IN      rcpt_tab.rcpt_no%type,
            '    sno_IN          rcpt_tab.sno%type,
            '    p_name_IN       rcpt_tab.p_name%type,
            '    inter_id_IN     rcpt_tab.inter_id%type,
            '    class_IN        rcpt_tab.class%type,
            '    amount_IN       rcpt_tab.amount%type,
            '    doctor_id_IN        rcpt_tab.doctor_id%type,
            '    r_date_IN       rcpt_tab.r_date%type,
            '    dept_id_IN      rcpt_tab.dept_id%type,
            '    exe_id_IN       rcpt_tab.exe_id%type,
            '    fcywxh_IN       rcpt_tab.fcywxh%type,
            '    hosp_price_IN       rcpt_tab.hosp_price%type
            
            gstrSQL = "rcpt_tab_Insert("
            gstrSQL = gstrSQL & "'0',"
            gstrSQL = gstrSQL & "'" & Lpad(Nvl(!NO), 8, "0") & "',"
            gstrSQL = gstrSQL & "" & lngCount & ","
            gstrSQL = gstrSQL & "'" & Nvl(!����) & "',"
            gstrSQL = gstrSQL & "'" & Nvl(!��Ŀ����) & "',"
            gstrSQL = gstrSQL & "'" & "01" & "',"
            gstrSQL = gstrSQL & "" & Nvl(!����, 0) & ","
            gstrSQL = gstrSQL & "'" & Trim(Nvl(!ҽ��)) & "',"
'            gstrSQL = gstrSQL & "to_date('" & Format(!�Ǽ�ʱ��, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'),"
            gstrSQL = gstrSQL & "to_date('" & Format(!�Ǽ�ʱ��, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
            gstrSQL = gstrSQL & "'" & Trim(Nvl(!����)) & "',"
            gstrSQL = gstrSQL & "'" & Substr(Trim(Nvl(!����)), 1, 10) & "',"
            gstrSQL = gstrSQL & "'" & Substr(Nvl(!ID), 1, 20) & "',"
            gstrSQL = gstrSQL & "" & Nvl(!����, 0) & ")"
            
            ExecuteProcedure_���� "������ϸ���м��"
            .MoveNext
        Loop
        �������_���� = True
    End With
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Sub ExecuteProcedure_����(ByVal strCaption As String)
'���ܣ�ִ��SQL���
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_����.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Public Function �����ʻ�תԤ��_����(lngԤ��ID As Long, curMoney As Currency) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
End Function

Public Function סԺ�������_����(rsExse As Recordset, strSelfNo As String, strSelfPwd As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim strסԺ�� As String
    Dim STR���� As String
    Dim strReptNo As String
    Dim str������� As String
    Dim dbl�Ը���� As Double
    Dim dblͳ���ʽ� As Double
    Dim dblԭʼ��� As Double
    Dim dblAccount As Double
    Dim intWait As Integer
    Dim sngBegin As Single
    
    Dim rsTmp As New ADODB.Recordset
    Dim rsExpen As New ADODB.Recordset
    
    
    Err = 0: On Error GoTo errHand:
    
    gstrSQL = "select סԺ��,���� from ������Ϣ where ����id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", CLng(rsExse!����ID))
    
    strסԺ�� = IIf(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
    STR���� = rsTmp!����
    rsTmp.Close
    
    With rsExse
        dblԭʼ��� = 0
        .MoveFirst
        Do While Not .EOF
            dblԭʼ��� = dblԭʼ��� + !���
            .MoveNext
        Loop
        .MoveFirst
    End With
    
    gstrSQL = "" & _
        "   Select  a.id, A.NO,A.���,b.���� ,B.���� as ��������,C.��Ŀ���� as ҽ����Ŀ����,d.���� as ��Ŀ," & _
        "           A.����ʱ��,A.������ as ҽ��,decode(d.�Ƿ���,1,a.ʵ�ս��,Nvl(A.����,1)*A.����) as ����," & _
        "           decode(d.�Ƿ���,1,1,a.��׼����) ����" & _
        " from סԺ���ü�¼ A,���ű� B,�շ�ϸĿ d ," & _
        "       (   Select * From ����֧����Ŀ Where ����=[1]) C" & _
        " where nvl(a.ʵ�ս��,0)<>0 and  A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.�շ�ϸĿID(+) and a.�շ�ϸĿid=d.id " & _
        "       And A.����ID=[2] And A.���ʷ���=1 And Nvl(A.�Ƿ��ϴ�,0)=0 And A.��¼״̬<>0" & _
        " Order by A.��¼����,A.��¼״̬,A.NO,A.ID"
        
    Set rsExpen = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_����, CLng(rsExse!����ID))
    
    With rsExpen
        str������� = "02"
        Do While Not .EOF
            'ɾ����ǰδ���湦��
            If IsNull(!ҽ����Ŀ����) Then
                MsgBox "HIS�е���Ŀ��" & !��Ŀ & "��δ����ҽ����Ӧ�ı���," & vbCrLf & "���ܱ���ҽ������,���飡", vbExclamation + vbOKOnly, gstrSysName
                Exit Function
            End If
            
          '    hosp_id_IN      rcpt_tab.hosp_ID%type,
            '    rcpt_no_IN      rcpt_tab.rcpt_no%type,
            '    sno_IN          rcpt_tab.sno%type,
            '    p_name_IN       rcpt_tab.p_name%type,
            '    inter_id_IN     rcpt_tab.inter_id%type,
            '    class_IN        rcpt_tab.class%type,
            '    amount_IN       rcpt_tab.amount%type,
            '    doctor_id_IN        rcpt_tab.doctor_id%type,
            '    r_date_IN       rcpt_tab.r_date%type,
            '    dept_id_IN      rcpt_tab.dept_id%type,
            '    exe_id_IN       rcpt_tab.exe_id%type,
            '    fcywxh_IN       rcpt_tab.fcywxh%type,
            '    hosp_price_IN       rcpt_tab.hosp_price%type
            
            gstrSQL = "rcpt_tab_Insert("
            gstrSQL = gstrSQL & "'" & Format(strסԺ��, "00000000") & "',"
            gstrSQL = gstrSQL & "'" & Lpad(Nvl(!NO), 8, "0") & "',"
            gstrSQL = gstrSQL & "" & Nvl(!���, 0) & ","
            gstrSQL = gstrSQL & "'" & STR���� & "',"
            gstrSQL = gstrSQL & "'" & Nvl(!ҽ����Ŀ����, "") & "',"
            gstrSQL = gstrSQL & "'" & "02" & "',"
            gstrSQL = gstrSQL & "" & Nvl(!����, 0) & ","
            gstrSQL = gstrSQL & "'" & Trim(Nvl(!ҽ��)) & "',"
'            gstrSQL = gstrSQL & "to_date('" & Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'),"
            gstrSQL = gstrSQL & "to_date('" & Format(!����ʱ��, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
            gstrSQL = gstrSQL & "'" & Substr(Trim(Nvl(!����)), 1, 4) & "',"
            gstrSQL = gstrSQL & "'" & Substr(Trim(Nvl(!��������)), 1, 10) & "',"
            gstrSQL = gstrSQL & "'" & Substr(Nvl(!ID), 1, 20) & "',"
            gstrSQL = gstrSQL & "" & Nvl(!����, 0) & ")"
            ExecuteProcedure_���� "������ϸ���м��"
                            
'
'            gstrSQL = "insert into rcpt_tab(hosp_id,rcpt_no,sno,p_name,inter_id,class,amount,doctor_id,r_date,dept_id,exe_id,hosp_price)" _
'                & " values('" & Format(strסԺ��, "00000000") & "','" & !No & "'," & !��� & ",'" & str���� & "'," & !ҽ����Ŀ���� & ",'02'," & !���� & ",'" & !ҽ�� & _
'                "',to_date('" & Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'),'',''," & !���� & ")"
'            gcnOracle_����.Execute gstrSQL
            
            '�ϴ���Ͳ����ϴ�
            gstrSQL = "Update סԺ���ü�¼ set �Ƿ��ϴ�=1 where id=" & !ID
            gcnOracle.Execute gstrSQL
            
            .MoveNext
        Loop
        
        Do While True
            dbl�Ը���� = 0
            dblͳ���ʽ� = 0
            gstrSQL = "" & _
                "   Select acct_pay,self_pay " & _
                "   From diag_rec " & _
                "   where   LPAD(RTrim(hosp_id),8,'0')='" & Format(strסԺ��, "00000000") & "'" & _
                "           and rcpt_no is null and class='02'" & _
                "           and sno is null and p_name='" & STR���� & "'" & _
                "           and inter_id is null "
            If rsTmp.State = adStateOpen Then rsTmp.Close
            rsTmp.Open gstrSQL, gcnOracle_����
            If Not rsTmp.EOF Then
                dbl�Ը���� = dbl�Ը���� + IIf(IsNull(rsTmp!self_pay), 0, rsTmp!self_pay)
                dblͳ���ʽ� = dblͳ���ʽ� + IIf(IsNull(rsTmp!acct_pay), 0, rsTmp!acct_pay)
            End If
            
            If dbl�Ը���� + dblͳ���ʽ� > 0 Then '= dblԭʼ���
                סԺ�������_���� = "ҽ������;" & dblͳ���ʽ� & ";0"
                Exit Do
            End If
            
            '�޽��Ҳ������ͨ���˷�ʽ����
            If MsgBox("û�еõ�ҽ�������������ȴ���", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                סԺ�������_���� = "ҽ������;0;0"
                Exit Function
            End If
        Loop
    End With
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, rs�ʻ� As ADODB.Recordset) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    
    Dim strסԺ�� As String
    Dim STR���� As String
    Dim strReptNo As String
    Dim str������� As String
    Dim dbl�Ը���� As Double
    Dim dblͳ���ʽ� As Double
    Dim dblԭʼ��� As Double
    Dim lng����ID As Long
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curDate As Date
    
    Dim curסԺ���� As Currency, cur�������� As Currency
    Dim cur����ͳ�� As Currency, curͳ��֧�� As Currency
    Dim cur�����Ը� As Currency, curȫ�Ը� As Currency
    
    Dim rsTmp As New ADODB.Recordset
On Error GoTo ErrH
    סԺ����_���� = False
    
    gstrSQL = "select סԺ��,���� from ������Ϣ where ����id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", CLng(rs�ʻ�!����ID))
    
    strסԺ�� = rsTmp!סԺ��
    STR���� = rsTmp!����
    rsTmp.Close
    
    dbl�Ը���� = 0
    dblͳ���ʽ� = 0
    lng����ID = rs�ʻ�!����ID
    
    gstrSQL = "" & _
        "   Select acct_pay,self_pay " & _
        "   From diag_rec " & _
        "   Where LPAD(RTrim(hosp_id),8,'0')='" & Format(strסԺ��, "00000000") & "'" & _
        "           and rcpt_no is null and class='02'" & _
        "           and sno is null and p_name='" & STR���� & "'" & _
        "           and inter_id is null "
            
    If rsTmp.State = adStateOpen Then rsTmp.Close
    rsTmp.Open gstrSQL, gcnOracle_����
    If Not rsTmp.EOF Then
        dbl�Ը���� = dbl�Ը���� + IIf(IsNull(rsTmp!self_pay), 0, rsTmp!self_pay)
        dblͳ���ʽ� = dblͳ���ʽ� + IIf(IsNull(rsTmp!acct_pay), 0, rsTmp!acct_pay)
    End If

    gstrSQL = "" & _
        "   Select Sum(���ʽ��) as ���ʽ�� " & _
        "   From סԺ���ü�¼ " & _
        "   Where Nvl(���ӱ�־,0)<>9 And ����ID=" & lng����ID
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    dblԭʼ��� = rsTmp.Fields(0)
    
        
    '��д�����
    curDate = zlDatabase.Currentdate
        
    With rsTmp
        'סԺ����,�����ܶ�,����ͳ�ﲿ��,ͳ��֧������
        '���ڶԷ����ṩ�����Բ�����ȡסԺ�����ͽ���ͳ����
        
        curסԺ���� = 0
        cur�������� = dblԭʼ���
        cur����ͳ�� = 0
        curͳ��֧�� = dblͳ���ʽ�
        curȫ�Ը� = 0
        cur�����Ը� = cur�������� - curȫ�Ը� - cur����ͳ��
        
        '�ʻ������Ϣ
        Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
        If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
                
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(curDate) & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� + cur����ͳ�� & "," & _
            curͳ�ﱨ���ۼ� + curͳ��֧�� & "," & intסԺ�����ۼ� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
            Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & curסԺ���� & ",NULL," & curסԺ���� & "," & _
            cur�������� & "," & curȫ�Ը� & "," & cur�����Ը� & "," & cur����ͳ�� & "," & curͳ��֧�� & ",NULL,NULL,NULL,NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        '���ս������
        
        gstrSQL = "zl_���ս������_insert(" & lng����ID & ",1," & cur����ͳ�� & "," & curͳ��֧�� & ",NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End With
    '-------------------------------------------
    
    'ɾ���м����ݿ�Ľ�������
    gstrSQL = "delete from diag_rec where LPAD(RTrim(hosp_id),8,'0')='" & Format(strסԺ��, "00000000") & "' and rcpt_no is null and class='02'" _
        & " and sno is null and p_name='" & STR���� & "' and inter_id is null "
    gcnOracle_����.Execute gstrSQL
    
    
    סԺ����_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_����(lng����ID As Long, rs�ʻ� As ADODB.Recordset) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    Dim lng����ID As Long
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curDate As Date, lng��ID As Long
    
    
    Dim curסԺ���� As Currency, cur�������� As Currency
    Dim cur����ͳ�� As Currency, curͳ��֧�� As Currency
    Dim dbl�Ը����  As Currency, dblͳ���ʽ�  As Currency
    Dim cur�����Ը� As Currency, curȫ�Ը� As Currency
    
    Dim rsTmp As New ADODB.Recordset
    Dim strסԺ�� As String, STR���� As String
On Error GoTo ErrH
    סԺ�������_���� = False
    gstrSQL = "select סԺ��,���� from ������Ϣ where ����id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", CLng(rs�ʻ�!����ID))
    
    Err.Raise 9000, gstrSysName, "��ҽ���ӿڲ�֧�ֳ�Ժ�������!"
    Exit Function
    
    strסԺ�� = rsTmp!סԺ��
    STR���� = rsTmp!����
    rsTmp.Close
    
    dbl�Ը���� = 0
    dblͳ���ʽ� = 0
    lng����ID = rs�ʻ�!����ID
    
    gstrSQL = "delete from diag_rec where LPAD(RTrim(hosp_id),8,'0')='" & Format(strסԺ��, "00000000") & "' and class='02'" _
            & " and p_name='" & STR���� & "' "
    gcnOracle_����.Execute gstrSQL
    
    curDate = zlDatabase.Currentdate
    '��ȡ���Ϻ�Ľ���ID
    gstrSQL = "Select A.ID From ���˽��ʼ�¼ A,���˽��ʼ�¼ B" & _
        " Where A.NO=B.NO And A.��¼״̬=2 And B.��¼״̬=3" & _
        " And B.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    If rsTmp.EOF Then
        Err.Raise 9000, gstrSysName, "δ�������ϵĽ������ݣ�", vbInformation, gstrSysName
        Exit Function: סԺ�������_���� = False
    End If
    
    With rsTmp
        lng��ID = .Fields("ID").Value
        
        '�ʻ������Ϣ
        Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
        If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
    End With
    
    gstrSQL = "Select * From ���ս������ Where Nvl(����,0)=0 And ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    With rsTmp
        If Not .EOF Then
            cur����ͳ�� = IIf(IsNull(!����ͳ����), 0, !����ͳ����)
            curͳ��֧�� = IIf(IsNull(!ͳ�ﱨ�����), 0, !ͳ�ﱨ�����)
        End If
    End With
    
    gstrSQL = "Select * From ���ս����¼ Where ����=2 And ��¼ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
            
    With rsTmp
        If Not .EOF Then
            curסԺ���� = IIf(IsNull(!ʵ������), 0, !ʵ������)
            cur�������� = IIf(IsNull(!�������ý��), 0, !�������ý��)
            cur�����Ը� = IIf(IsNull(!�����Ը����), 0, !�����Ը����)
            If cur����ͳ�� = 0 Then cur����ͳ�� = IIf(IsNull(!����ͳ����), 0, !����ͳ����)
            If curͳ��֧�� = 0 Then curͳ��֧�� = IIf(IsNull(!ͳ�ﱨ�����), 0, !ͳ�ﱨ�����)
            curȫ�Ը� = IIf(IsNull(!ȫ�Ը����), 0, !ȫ�Ը����)
        End If
        
        '�����µ����ϼ�¼
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(curDate) & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� - cur����ͳ�� & "," & _
            curͳ�ﱨ���ۼ� - curͳ��֧�� & "," & intסԺ�����ۼ� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        '���ս������
        gstrSQL = "zl_���ս������_insert(" & lng��ID & ",1," & -1 * cur����ͳ�� & "," & -1 * curͳ��֧�� & ",NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
        
        '���ս����¼
        gstrSQL = "zl_���ս����¼_insert(2," & lng��ID & "," & TYPE_���� & "," & lng����ID & "," & Year(curDate) & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & curͳ�ﱨ���ۼ� & "," & _
            intסԺ�����ۼ� & "," & curסԺ���� & ",NULL," & curסԺ���� & "," & -1 * cur�������� & "," & _
             -1 * curȫ�Ը� & "," & -1 * cur�����Ը� & "," & _
            -1 * cur����ͳ�� & "," & -1 * curͳ��֧�� & ",NULL,NULL,NULL,NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End With
    סԺ�������_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ������Ϣ_����(ByVal lngErrCode As Long) As String
'���ܣ����ݴ���ŷ��ش�����Ϣ

End Function

Public Function ������ϸ_����(ByVal str���ݺ� As String, ByVal int���� As Integer, ByVal int״̬ As Integer, ByVal intClinic As Integer) As Boolean
'����: �������������ϸ(���۵�)��������ʹ�á�
'˵������ΪZLHIS9/10���շѻ��۵��ļ�¼��ʽ��ͬ�����Ա���ʹ�ü�¼���ʣ���¼״̬������
'------------------------------------------------------------------------------------------------------------------
'����ģ�飺1121-�����շ�
    On Error GoTo errHand
    Dim rsExse As New ADODB.Recordset
    Dim סԺ�� As String
    
    ������ϸ_���� = False
    
    'ɾ����ǰδ���湦��
    gcnOracle_����.BeginTrans
    gstrSQL = "" & _
        "   delete from rcpt_tab " & _
        "   where LPAD(RTrim(hosp_id),8,'0')='" & IIf(intClinic = 1, "00000000", Format(סԺ��, "00000000")) & "'" & _
        "           and rcpt_no='" & str���ݺ� & "' " & _
        "           and class='" & IIf(intClinic = 1, "01", "02") & "'"
        
    gcnOracle_����.Execute gstrSQL
    
    '�������id���������������ݺš���š�ҽ����Ŀ���롢���������ۡ�������ҽ�����������š�����Ա������ʱ��
    gstrSQL = "" & _
        "   Select A.ID,A.����ID,A.���� As ��������,A.No As ���ݺ�,Nvl(A.�۸񸸺�,A.���) As ���," & _
        "           C.��Ŀ���� As ҽ����Ŀ����,decode(d.�Ƿ���,1,Sum(A.ʵ�ս��),Avg(A.����*Nvl(A.����,1))) As ����,decode(d.�Ƿ���,1,1,Sum(A.��׼����)) As ����," & _
        "         Sum(A.ʵ�ս��) As ���,A.������ As ����ҽ��,b.����,B.���� As ��������,A.������ As ����Ա,a.����ʱ��,d.���� as ��Ŀ���� " & _
        "   From סԺ���ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=[4]) C,�շ�ϸĿ D " & _
        "   Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.�շ�ϸĿID(+) and A.�շ�ϸĿID=d.id " & _
        "           And A.��¼����=[1] And A.��¼״̬=[2] And A.NO=[3]" & _
        "   Group By A.ID,A.No,Nvl(A.�۸񸸺�,A.���),A.����ID,A.����,C.��Ŀ����,A.������,b.����,B.����,A.������,d.����,a.����ʱ��,d.�Ƿ��� "
    
    Set rsExse = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", int����, int״̬, str���ݺ�, TYPE_����)
    
    With rsExse
        If .EOF Then
            MsgBox "û��һ����������ϸ���ݣ�������û������ҽ�����룬���飡", vbInformation, gstrSysName
            gcnOracle_����.RollbackTrans
            Exit Function
        End If
        
        .MoveFirst
        Do While Not .EOF
            If IsNull(!ҽ����Ŀ����) = True Then
                MsgBox "�����а���δ���ñ���֧����Ŀ���շ���Ŀ��" & !��Ŀ���� & "��," & vbCrLf & "����ִ��ҽ�����ף�", vbInformation, gstrSysName
                gcnOracle_����.RollbackTrans
                Exit Function
            End If
            
            '    hosp_id_IN      rcpt_tab.hosp_ID%type,
            '    rcpt_no_IN      rcpt_tab.rcpt_no%type,
            
            '    sno_IN          rcpt_tab.sno%type,
            '    p_name_IN       rcpt_tab.p_name%type,
            '    inter_id_IN     rcpt_tab.inter_id%type,
           
            '    class_IN        rcpt_tab.class%type,
            '    amount_IN       rcpt_tab.amount%type,
            '    doctor_id_IN    rcpt_tab.doctor_id%type,
            '    r_date_IN       rcpt_tab.r_date%type,
            '    dept_id_IN      rcpt_tab.dept_id%type,
            '    exe_id_IN       rcpt_tab.exe_id%type,
            '    fcywxh_IN       rcpt_tab.fcywxh%type,
            '    hosp_price_IN   rcpt_tab.hosp_price%type
            
            gstrSQL = "rcpt_tab_Insert("
            gstrSQL = gstrSQL & "'" & IIf(intClinic = 1, "0", Format(סԺ��, "00000000")) & "',"
            gstrSQL = gstrSQL & "'" & Lpad(str���ݺ�, 8, "0") & "',"
            gstrSQL = gstrSQL & "" & Nvl(!���, 0) & ","
            gstrSQL = gstrSQL & "'" & Nvl(!��������) & "',"
            gstrSQL = gstrSQL & "'" & Nvl(!ҽ����Ŀ����, "") & "',"
            gstrSQL = gstrSQL & "'" & IIf(intClinic = 1, "01", "02") & "',"
            gstrSQL = gstrSQL & "" & Nvl(!����, 0) & ","
            gstrSQL = gstrSQL & "'" & Trim(Nvl(!����ҽ��)) & "',"
            '����Ҫ���ܴ�ʱ����
            'gstrSQL = gstrSQL & "to_date('" & Format(!����ʱ��, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'),"
            gstrSQL = gstrSQL & "to_date('" & Format(!����ʱ��, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
            gstrSQL = gstrSQL & "'" & Trim(Nvl(!����)) & "',"
            gstrSQL = gstrSQL & "'" & Substr(Trim(Nvl(!��������)), 1, 10) & "',"
            
            gstrSQL = gstrSQL & "'" & Substr(Nvl(!ID), 1, 20) & "',"
            gstrSQL = gstrSQL & "" & Nvl(!����, 0) & ")"
            ExecuteProcedure_���� "������ϸ���м��"
                                        
            
'            gstrSQL = "insert into rcpt_tab" _
                    & "(hosp_id,rcpt_no,sno,p_name,inter_id,class,amount,doctor_id,r_date,dept_id,exe_id,hosp_price)" _
             & " values('" & IIf(intClinic = 1, "0", Format(סԺ��, "00000000")) & "','" & ���ݺ� & "'," & ��� & ",'" & �������� & "'," & ҽ����Ŀ���� & ",'" & IIf(intClinic = 1, "01", "02") & "'," _
                      & ���� & ",'" & ����ҽ�� & "',to_date('" & ����ʱ�� & "','yyyy-mm-dd'),'','" & �������� & "'," & ���� & ")"
 '           gcnOracle_����.Execute gstrSQL
            .MoveNext
        Loop
    End With
    ������ϸ_���� = True
    gcnOracle_����.CommitTrans
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    gcnOracle_����.RollbackTrans
End Function
