Attribute VB_Name = "mdl��ɽ"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

Public gcn��ɽ As New ADODB.Connection
Private mstr����� As String

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTTOPMOST = -2

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetprivateprofileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, _
    ByVal lpDefault As String, ByVal lpRetrm_String As String, ByVal cbReturnString As Integer, ByVal FileName As String) As Integer

'���½ṹ��������¼��������������ڽ���ʱ�˶�
Private Type typBalance
    curҽ������ As Double
    cur�����ʻ� As Double
    cur�󲡻��� As Double
End Type
Private pre_Balance As typBalance

Public Function ҽ����ʼ��_��ɽ() As Boolean
'���ܣ������Ƿ�������ӵ�ǰ�÷�������
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String, rs��ɽ As New ADODB.Recordset
    '��������Ѿ��򿪣��ǾͲ����ٲ���
    If gcn��ɽ.State = adStateOpen Then
        ҽ����ʼ��_��ɽ = True
        Exit Function
    End If
     
    On Error GoTo ErrH
    
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", TYPE_�����ɽ)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "��ɽ������"
                strServer = strTemp
            Case "��ɽ�û���"
                strUser = strTemp
            Case "��ɽ�û�����"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        gcn��ɽ.Open "Provider=SQLOLEDB.1;Initial Catalog=hw_interface;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
    Else
        gcn��ɽ.Open "Driver={Microsoft ODBC for Oracle};Server=" & _
            strServer, strUser, strPass
    End If
    If Err <> 0 Then
        MsgBox "ҽ��ǰ�÷���������ʧ�ܡ�", vbInformation, gstrSysName
        ҽ����ʼ��_��ɽ = False
        Exit Function
    End If
    ҽ����ʼ��_��ɽ = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    ҽ����ʼ��_��ɽ = False
End Function

Public Function ��ݱ�ʶ_��ɽ(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmIDentified As New frmIdentify��ɽ
    Dim strPatiInfo As String, cur��� As Currency
    Dim arr, datCurr As Date, str����� As String
    Dim strSQL As String, str���ⲡ As String
    Dim strTemp As String, errLine As Integer
    
    '�ж��Ƿ񱣴���IC����֤��
    strTemp = Get���ղ���_��ɽ("����֤��")
    If strTemp = "" Then
        MsgBox "����ҽ�����������ñ���ҽ����IC����֤�롣", vbInformation, gstrSysName
        Exit Function
    End If
    
    frmIDentified.mstr��֤�� = strTemp
    frmIDentified.Tag = bytType
    frmIDentified.Show 1
    'New:0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����)
    On Error GoTo errHandle
    strPatiInfo = frmIDentified.mstrPatiInfo: errLine = 1
    cur��� = frmIDentified.mcur���: errLine = 2
    Unload frmIDentified: errLine = 3
    
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)

        lng����ID = BuildPatiInfo(bytType, strPatiInfo & ";;;;" & cur��� & ";;;;;;;" & cur��� & ";;;;;", lng����ID, TYPE_�����ɽ): errLine = 4
        '���ظ�ʽ:�м���벡��ID
        strPatiInfo = strPatiInfo & ";" & lng����ID & ";;;;" & cur��� & ";;;;;;;" & cur��� & ";;;;;": errLine = 5
    Else
        ��ݱ�ʶ_��ɽ = "": errLine = 6
        MsgBox "ҽ��������Ϣ��ȡʧ��", vbInformation, gstrSysName
        Exit Function
    End If
    arr = Split(strPatiInfo, ";"): errLine = 12
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        '����Ƿ����ⲡ
        str���ⲡ = frmIDentified.mstr���ⲡ: errLine = 7
        gstr���ⲡ�� = str���ⲡ: errLine = 8
    ElseIf Val(Get���ղ���_��ɽ("���õ���")) = 1 Then           '��¼�����Ƿ������ⲡ
        If gbln�������� = True Then
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�����ɽ & ",'�Ҷȼ�','''1''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            gbln�������� = False
            str���ⲡ = Get����ID(CStr(arr(1)), CStr(TYPE_�����ɽ))
        Else
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�����ɽ & ",'�Ҷȼ�','''0''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        End If
    Else
        str���ⲡ = Get����ID(CStr(arr(1)), CStr(TYPE_�����ɽ)): errLine = 9
    End If
    If bytType <> 0 Then
        ��ݱ�ʶ_��ɽ = strPatiInfo: errLine = 10
    End If
    '���Ϊ���ﲡ�ˣ��ͽ��Ž�������Ǽ�
    datCurr = zlDatabase.Currentdate: errLine = 11
    str����� = ToVarchar(lng����ID & Format(datCurr, "yyddhhmmss"), 16): errLine = 13
    mstr����� = str�����: errLine = 14
    '��������Ǽ�׼��
    If bytType <> 0 Then
        ��ݱ�ʶ_��ɽ = strPatiInfo
    Else
        strSQL = "insert into Check_doex_interface(Bill_no,App_code" & _
                ",Doct_flag,Doex_no,Ill_type,Ic_id,Is_bala,Regi_op_id) values('" & _
                str����� & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & IIf(bytType = 1, 1, 0) & "','" & _
                Left(str�����, 10) & "','" & str���ⲡ & _
                "','" & arr(2) & arr(0) & "','0','" & ToVarchar(UserInfo.����, 8) & "')": errLine = 15
        gcn��ɽ.Execute strSQL: errLine = 16
        '��������Ǽ�����
        strSQL = "insert into Check_bill_request(Bill_no,App_code," & _
                "Request_status) values('" & str����� & "','" & _
                Mid(gstrҽԺ����, 1, 4) & "','0')": errLine = 17
        gcn��ɽ.Execute strSQL: errLine = 18
        If Checkrequest(str�����) = False Then
            'ɾ��ʧ�ܵ�����Ǽǵ�
            strSQL = "delete from Check_bill_request where Bill_no = '" & str����� & _
                    "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": errLine = 19
            gcn��ɽ.Execute strSQL: errLine = 10
            strSQL = "delete from Check_doex_interface where Bill_no = '" & _
                    str����� & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": errLine = 21
            gcn��ɽ.Execute strSQL: errLine = 22
            ��ݱ�ʶ_��ɽ = ""
            Exit Function
        Else
            ��ݱ�ʶ_��ɽ = strPatiInfo
        End If
    End If
    Exit Function
errHandle:
    MsgBox "���������[�����֤]ģ���" & errLine & "��", vbInformation, "����"
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_��ɽ = ""
End Function

Public Function �������_��ɽ(lng����ID As Long, cur����֧�� As Currency, strҽ���� As String, curȫ�Ը� As Currency, cur���Ը� As Currency, curҽ������ As Currency) As Boolean
'���ܣ���������ý�����ϸ���ݲ��ҽ��н���
'������������ϸ����ʧ�ܣ���ֱ�ӽ������������غ���ʧ��
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim rs��ɽ As New ADODB.Recordset
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, curDate As Date
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur�����ʻ� As Currency, cur���� As Currency, cur����ͳ���޶� As Currency
    Dim cur���ͳ���޶� As Currency, cur�����Ը� As Currency, cur��� As Currency
    Dim cur�������� As Currency, cur�ز�ͳ�� As Currency, str�������� As String
    Dim strInipath As String, str�����ļ��� As String
    
    On Error GoTo errHandle
    '����������㣬�޷����н���
    cur��� = �������_��ɽ(Get����ID(CStr(strҽ����), CStr(TYPE_�����ɽ)))
    If cur����֧�� > cur��� Then
        Err.Raise 9000, gstrSysName, "��Ҫ�ķ����Ѿ�����ʣ�����", vbInformation, gstrSysName
        �������_��ɽ = False
        Exit Function
    End If
    If ������ϸ����(1, lng����ID) = False Then
        �������_��ɽ = False
        Exit Function
    End If
    
    WriteInfo vbCrLf & "��ʼ�������"
    '���н���׼��
    strSQL = "Update Check_doex_interface set Ps_account_pay = " & _
            CStr(cur����֧��) & ",Bala_op_id = '" & ToVarchar(UserInfo.����, 8) & _
            "' where Bill_no = '" & mstr����� & "' and " & _
            "App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSQL
    
    '�ύ��������
    strSQL = "update Check_bill_request set Request_status = '1',Request_Result=null where" & _
            " Bill_no ='" & mstr����� & "' and " & _
            " App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    WriteInfo "��������:" & strSQL
    gcn��ɽ.Execute strSQL
    
    str�����ļ��� = mstr����� & ".ini"
    'Modified By ���� ���� 06:10:58
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        '����д����������¿�ʧ�ܣ�������������з��ش�������һ���ͻ��������
        WriteInfo "�¿�:" & mstr�����
        Call Shell("D:\hw_ic_write\hw_ic_write.exe " & mstr�����, vbHide)
        '��ȡ�����ļ�
        strInipath = Trim(Get���ղ���_��ɽ("�����ļ�λ��"))
        If strInipath = "" Then strInipath = App.Path
        If Right(strInipath, 1) <> "\" Then strInipath = strInipath & "\"
        '�����ļ�δ��ֵ��ѯ��֮��ֵ����ע����һ��䣬����ҽ������ʱ����������·�������ļ�����
        strInipath = strInipath & str�����ļ���
        
        WriteInfo "�����ļ���:" & strInipath
        If Not frm�ȴ���Ӧ_ǭ��.ShowME(strInipath) Then
            WriteInfo "��ȡ�ļ������ж�"
            Err.Raise 9000, gstrSysName, "�ȴ��ļ����ز������ж�,���������Ƿ��ѽ���,�������ݽ��к˶�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If GetIniS("Sign", "Sign", "0", strInipath) = "0" Then
            WriteInfo "���ش���:" & GetIniS("Sign", "Error_txt", "", strInipath)
            Err.Raise 9000, gstrSysName, GetIniS("Sign", "Error_txt", "", strInipath), vbInformation, "ҽ������"
            Exit Function
        End If
        
        cur����֧�� = CCur(GetIniS("Recorde", "Ps_account_pay", "0", strInipath))
        cur��� = CCur(GetIniS("Recorde", "Ps_bala", "0", strInipath))
        WriteInfo "����:" & cur����֧�� & vbCrLf & "���:" & cur���
        cur��� = CCur(GetIniS("Recorde", "Ps_cost_pay", "0", strInipath))
        
    Else
        If Checkrequest(mstr�����) = False Then �������_��ɽ = False: Exit Function
        
        '���������
        curDate = zlDatabase.Currentdate
        '��ȡ�����ʻ�֧���͸����ֽ�֧��
        strSQL = "select Ps_account_pay,Ps_cost_pay,Ps_bala,Plan_pay,acc_cyc from Check_doex_interface" & _
                " where Bill_no ='" & mstr����� & "' and " & _
                " App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
        rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
        cur����֧�� = Nvl(rs��ɽ("Ps_account_pay"), 0)
        cur��� = Nvl(rs��ɽ("Ps_bala"), 0)
        curȫ�Ը� = Nvl(rs��ɽ("Ps_cost_pay"), 0)
        str�������� = Nvl(rs��ɽ("acc_cyc"), "")
    End If
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        cur�ز�ͳ�� = Nvl(rs��ɽ("Plan_pay"), 0)
    Else
        cur�ز�ͳ�� = 0
    End If
    curҽ������ = cur�ز�ͳ��
    cur�������� = curȫ�Ը� + cur����֧�� + cur�ز�ͳ��
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_�����ɽ, Get����ID(CStr(strҽ����), CStr(TYPE_�����ɽ)), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & Get����ID(CStr(strҽ����), CStr(TYPE_�����ɽ)) & _
            "," & TYPE_�����ɽ & "," & Year(curDate) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & _
            cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�����ɽ & "," & _
            Get����ID(CStr(strҽ����), CStr(TYPE_�����ɽ)) & "," & Year(curDate) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� + cur����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + cur�ز�ͳ�� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur�������� & "," & curȫ�Ը� & "," & cur���Ը� & ",NULL," & cur�ز�ͳ�� & ",NULL,NULL," & _
            cur����֧�� & ",NULL,NULL,NULL,'" & mstr����� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
'        gstrSQL = "zl_�������ڼ�¼_insert("
        gstrSQL = "Insert into zlhis.�������ڼ�¼ values (" & lng����ID & ",'" & str�������� & "'," & cur�������� & "," & cur����֧�� & "," & cur�ز�ͳ�� & ",'L',to_date('" & Format(curDate, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
        gcnOracle.Execute gstrSQL
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "������ҽ��")
    End If

    strSQL = "delete from Check_bill_request  where" & _
            " Bill_no ='" & mstr����� & "' and  App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSQL
    �������_��ɽ = True
    WriteInfo "�����������"
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    �������_��ɽ = False
End Function

Public Function ����������_��ɽ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, str�������� As String
    Dim curƱ���ܽ�� As Currency
    Dim curDate As Date
    
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ��  From ������ü�¼ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", lng����ID)
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", lng����ID)
    
    lng����ID = rsTemp("����ID")
    
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", TYPE_�����ɽ, lng����ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
'    str��ˮ�� = rsTemp("֧��˳���")
    
'    strInput = "99|" & str��ˮ�� & "|" & ToVarchar(UserInfo.����, 20)
'    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_�����ɽ, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�����ɽ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�����ɽ & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        cur�����ʻ� * -1 & ",Null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        gstrSQL = "Select * from �������ڼ�¼ where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
        If Not rsTemp.EOF Then
            str�������� = rsTemp!��������
    '        gstrSQL = "zl_�������ڼ�¼_insert("
            gstrSQL = "Insert into zlhis.�������ڼ�¼ values (" & lng����ID & ",'" & str�������� & "'," & curƱ���ܽ�� * -1 & "," & cur�����ʻ� * -1 & "," & Nvl(rsTemp("ͳ��"), 0) * -1 & ",'L',to_date('" & Format(curDate, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
            gcnOracle.Execute gstrSQL
        End If
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "������ҽ��")
    End If

    ����������_��ɽ = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function �����������_��ɽ(rs������ϸ As Recordset, str���㷽ʽ As String) As Boolean
    Dim cur����֧�� As Currency, cur�����ֽ�֧�� As Currency, cur�����ʻ�֧�� As Currency
    Dim curͳ��֧�� As Currency, cur���֧�� As Currency, lngCount As Long
    Dim strSQL As String, rs��ɽ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset, strBillNO As String
    Dim strMedi As String, strPageId As String
    Dim lng����ID As Long, cur�����ܶ� As Currency
    Dim i As Integer, frm�ȴ� As New frm�ȴ���Ӧ��ɽ
    Dim datCurr As Date, cur�����ʻ���� As Currency
    If Val(Get���ղ���_��ɽ("���õ���")) = 0 Then         '�������������,�����������
        �����������_��ɽ = False
        Exit Function
    End If
    '�ж��Ƿ��Ѿ���������
    If rs������ϸ.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(Get���ղ���_��ɽ("���õ���")) = 3 Or gbln�������� = False Or Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        cur�����ܶ� = 0
        While Not rs������ϸ.EOF
            cur�����ܶ� = cur�����ܶ� + rs������ϸ!ʵ�ս��
            rs������ϸ.MoveNext
        Wend
        str���㷽ʽ = "�����ʻ�;" & cur�����ܶ� & ";1"
        �����������_��ɽ = True
        Exit Function
    End If
    
    On Error GoTo errHandle
    '������˵Ĳ�����ҳ��Ҳͬʱ��������㵥��
    lng����ID = rs������ϸ!����ID
    strBillNO = mstr�����
    
    '�����ǰ��Ҫ�����
    strSQL = "select max(Charge_item_no) as charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
    If rs��ɽ.EOF Then
        i = 1
    Else
        i = Nvl(rs��ɽ("Charge_item_no"), 0) + 1
    End If
    rs������ϸ.MoveFirst
    lngCount = 0
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then Call ShowWindow(frm�ȴ�.hwnd, 9)
    SetPos frm�ȴ�.hwnd
    frm�ȴ�.Move (Screen.Width - frm�ȴ�.Width) / 2, (Screen.Height - frm�ȴ�.Height) / 2
    DoEvents
    Do While Not rs������ϸ.EOF
        '������еķ��ý��
        cur�����ʻ�֧�� = cur�����ʻ�֧�� + rs������ϸ("ʵ�ս��")
        gstrSQL = "Select * From �շ�ϸĿ where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "������ҽԺ", CLng(rs������ϸ("�շ�ϸĿID")))
        If rsTmp!��� = 5 Or rsTmp!��� = 6 Or rsTmp!��� = 7 Then
            strMedi = "1"
        Else
            strMedi = "2"
        End If
        
        '���������ύ׼��
        strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                " values('" & strBillNO & "','" & _
                Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & _
                rs������ϸ("����ID") & "','" & rs������ϸ("������") & _
                "',to_Date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & rs������ϸ("���ձ���") & _
                "','Ԥ����','" & strMedi & "','" & _
                rs������ϸ("���㵥λ") & "'," & rs������ϸ("����") & "," & _
                CStr(rs������ϸ("����")) & "," & CStr(rs������ϸ("ʵ�ս��")) & _
                ",to_date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & UserInfo.���� & "')"
        gcn��ɽ.Execute strSQL
        
        '�����ύ����
        strSQL = "Insert into Check_Item_Request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
        strBillNO & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','0')"
        gcn��ɽ.Execute strSQL
        lngCount = lngCount + 1
        '�����ѯ����(�����ڴ�������в��ȴ�����״̬)
'        If frm�ȴ�.Result(2, strBillNo, i) = False Then
'            �����������_��ɽ = False
'            MsgBox "�ڽ���Ĺ���֮�з����ж�", vbInformation, gstrSysName
'            GoTo ResetTrans
'        End If
'        '��ѯ�ύ���
'        strSql = "select Request_Result,Err_Code,Err_text from " & _
'                "check_item_request where Bill_no = '" & strBillNo & _
'                 "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
'                 "' and Charge_item_no = '" & CStr(i) & "'"
'        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
'        rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
'        If rs��ɽ.BOF Then
'            �����������_��ɽ = False
'            GoTo ResetTrans
'        Else
'            If rs��ɽ("Request_Result") = "0" Then
'                MsgBox "��������[" & rs��ɽ("Err_Code") & "]:" & vbCrLf & String(2, "��") & rs��ɽ("Err_text"), vbInformation, gstrSysName
'                �����������_��ɽ = False
'                GoTo ResetTrans
'            End If
'        End If

        '��HIS֮�еĻ������ݽ����޸�
        i = i + 1
        rs������ϸ.MoveNext
    Loop
    Do While True
        '��ѯ�ύ���
        strSQL = "select Request_Result,Err_Code,Err_text from " & _
                "check_item_request where Bill_no = '" & strBillNO & _
                 "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
                 "' and Request_result is Null"
        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
        rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
        If rs��ɽ.EOF Then Exit Do
        DoEvents
    Loop
    Unload frm�ȴ�
    cur�����ܶ� = cur�����ʻ�֧��
    '���н���׼��
    strSQL = "Update Check_doex_interface set Ps_account_pay = " & _
            CStr(cur����֧��) & ",Bala_op_id = '" & ToVarchar(UserInfo.����, 8) & _
            "' where Bill_no = '" & mstr����� & "' and " & _
            "App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSQL
    
    '�ύ��������
    strSQL = "update Check_bill_request set Request_status = '5',Request_Result=null where" & _
            " Bill_no ='" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSQL
    
    If Checkrequest(strBillNO) = False Then
        �����������_��ɽ = False
        GoTo ResetTrans
    End If
    
    '�ӶԷ������ݿ�֮����ȡ�����ʻ�֧�����ֽ�֧����ͳ��֧�������֧��
    strSQL = "select Ps_bala from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
    cur�����ʻ�֧�� = Nvl(rs��ɽ("Ps_bala"), 0)
    
    strSQL = "select Ps_account_pay,Ps_cost_pay,Plan_pay,Big_pay from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
    cur�����ʻ�֧�� = Nvl(rs��ɽ("Ps_account_pay"), 0)
    cur�����ֽ�֧�� = Nvl(rs��ɽ("Ps_cost_pay"), 0)
    curͳ��֧�� = Nvl(rs��ɽ("Plan_pay"), 0)
    cur���֧�� = Nvl(rs��ɽ("Big_pay"), 0)
    
'    '������������ʻ�֧��
'    cur�����ܶ� = cur�����ܶ� - curͳ��֧�� - cur���֧��
'    cur�����ʻ�֧�� = IIf(cur�����ʻ�֧�� > cur�����ܶ�, cur�����ܶ�, cur�����ʻ�֧��)
    
    str���㷽ʽ = "�����ʻ�;" & cur�����ʻ�֧�� & ";" & IIf(Val(Get���ղ���_��ɽ("���õ���")) = 3 Or Val(Get���ղ���_��ɽ("���õ���")) = 2, 1, 0) '�����޸ĸ����ʻ�
    If curͳ��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ = "", "", "|") & "ͳ��֧��;" & curͳ��֧�� & ";0" '�������޸�ͳ��֧��
    End If
    If cur���֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ = "", "", "|") & "���֧��;" & cur���֧�� & ";0" '�������޸Ĵ��֧��
    End If
    �����������_��ɽ = True
ResetTrans:             '�Ժ��ֵ��ݳ��ΪԤ������ϴ��ķ�����ϸ
    '�����ǰ��Ҫ�����
    rs������ϸ.MoveFirst
    strSQL = "select max(Charge_item_no) as charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
    i = Nvl(rs��ɽ("Charge_item_no"), 0) + 1
    rs������ϸ.MoveFirst
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then Call ShowWindow(frm�ȴ�.hwnd, 9)
    SetPos frm�ȴ�.hwnd
    frm�ȴ�.Move (Screen.Width - frm�ȴ�.Width) / 2, (Screen.Height - frm�ȴ�.Height) / 2
    DoEvents
    Do While Not rs������ϸ.EOF And lngCount > 0
        '������еķ��ý��
        cur�����ʻ�֧�� = cur�����ʻ�֧�� + rs������ϸ("ʵ�ս��")
        gstrSQL = "Select * From �շ�ϸĿ where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "������ҽԺ", CLng(rs������ϸ("�շ�ϸĿID")))
        If rsTmp!��� = 5 Or rsTmp!��� = 6 Or rsTmp!��� = 7 Then
            strMedi = "1"
        Else
            strMedi = "2"
        End If
        '���������ύ׼��
        strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                " values('" & strBillNO & "','" & _
                Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & _
                rs������ϸ("����ID") & "','" & rs������ϸ("������") & _
                "',to_Date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & rs������ϸ("���ձ���") & _
                "','Ԥ����','" & strMedi & "','" & _
                rs������ϸ("���㵥λ") & "'," & 0 - rs������ϸ("����") & "," & _
                CStr(rs������ϸ("����")) & "," & CStr(0 - rs������ϸ("ʵ�ս��")) & _
                ",to_date('" & Format(Date, "yyyy-MM-dd HH:MM:SS") & _
                "','yyyy-MM-dd HH24:MI:SS'),'" & UserInfo.���� & "')"
        gcn��ɽ.Execute strSQL
        
        '�����ύ����
        strSQL = "Insert into Check_Item_Request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
        strBillNO & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','0')"
        gcn��ɽ.Execute strSQL
        lngCount = lngCount - 1
        '�����ѯ����
'        If frm�ȴ�.Result(2, strBillNo, i) = False Then
'            �����������_��ɽ = False
'            MsgBox "�ڽ���Ĺ���֮�з����ж�", vbInformation, gstrSysName
'            Exit Function
'        End If
        '��ѯ�ύ���
        
        i = i + 1
        rs������ϸ.MoveNext
    Loop
    Do While True
        '��ѯ�ύ���
        strSQL = "select Request_Result,Err_Code,Err_text from " & _
                "check_item_request where Bill_no = '" & strBillNO & _
                 "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
                 "' and Request_result is Null"
        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
        rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
        If rs��ɽ.EOF Then Exit Do
        DoEvents
    Loop
    
    Unload frm�ȴ�
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    �����������_��ɽ = False
End Function

Public Function ������ϸ����(lng��� As Long, Optional lng����ID As Long, Optional strNO As String, Optional lng����ID As Long, Optional int���� As Integer, Optional int״̬ As Integer) As Boolean
'���ܣ�����ύ���������ϸ
'lng��� 1������  2��סԺ
'lng����ID�����������������
'strNo:���ݺ�
'int���ʣ�
'lng����ID  Ĭ��Ϊ0����ʾ�������ŵ��ݣ�����Ϊ������ָ�����˵ġ�����Ҫ����Ϊҽ���ڱ�����ʵ�ʱ���Ƿֲ������ύ���ݶ�����һ���ύ��
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim rs��ɽ As New ADODB.Recordset, strBillNO As String
    Dim strMedi As String, i As Integer, j As Integer, rsTemp As New ADODB.Recordset
    'i-ѭ���ۼӣ�������ţ��ڲ�������ϸʱ���ۼӣ�j-��ʱ��¼ԭ��ţ����ٴ�����ʱʹ��
    Dim blnInsert As Boolean
    Dim frm�ȴ� As New frm�ȴ���Ӧ��ɽ
     
    On Error GoTo errHandle
    If lng����ID = 0 Then
        If lng��� = 1 Then
            gstrSQL = "select ����ID from ������ü�¼ where ����ID = " & _
                    lng����ID & " and rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", lng����ID)
        Else
            gstrSQL = "select ����ID from סԺ���ü�¼ where NO = [1] and ��¼���� = [2] and ��¼״̬ = [3] and rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", strNO, int����, int״̬)
        End If
        
        lng����ID = rsTmp("����ID")
    End If
    If lng��� = 1 Then
       strBillNO = mstr�����
    Else
        gstrSQL = "select max(��ҳID) as ��ҳID from ������ҳ where ����ID =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", lng����ID)
        strBillNO = CStr(lng����ID) & "_" & CStr(rsTmp("��ҳID"))
    End If
    If lng��� = 1 Then
        '����ǰ���ݵļ�¼�ͼ���¼����ɾ��:ע�⣬�շ�ϸĿ��ν��д��ݻ���Ҫ�޸�
        strSQL = "delete from Check_item_list_interface where Bill_no = '" & _
                mstr����� & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
        strSQL = "delete from Check_item_request where Bill_no = '" & _
                mstr����� & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
        gstrSQL = "select A.ID,A.����ʱ��,A.���,A.NO,A.�շ����,A.������,A.�Ǽ�ʱ��," & _
                "A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��¼����,A.��¼״̬,D.��Ŀ���� as ϸĿ����,B.���� as ϸĿ����," & _
                "C.����  as ��Ŀ����,B.���㵥λ, (A.���� * A.����) as ����," & _
                "A.��׼����,A.ʵ�ս��,A.����Ա����,A.�Ƿ��ϴ� from  " & _
                "������ü�¼ A,�շ�ϸĿ B,������Ŀ C,����֧����Ŀ D" & _
                " where A.��¼״̬<>0 And Nvl(A.�Ƿ��ϴ�,0)=0 And A.�շ�ϸĿID = B.ID and A.������ĿID = C.ID and " & _
                "A.����ID =[3] and A.�շ�ϸĿID = D.�շ�ϸĿID and D.���� = [1] And a.����ID = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", TYPE_�����ɽ, lng����ID, lng����ID)
    Else
        gstrSQL = "select A.ID,A.����ʱ��,A.���,A.NO,A.�շ����,A.������,A.�Ǽ�ʱ��," & _
                "A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��¼����,A.��¼״̬,D.��Ŀ���� as ϸĿ����,B.���� as ϸĿ����,C.���� as " & _
                "��Ŀ����,B.���㵥λ, (A.���� * A.����) as ����,A.��׼����,A.ʵ�ս��," & _
                "A.����Ա����,A.�Ƿ��ϴ� from סԺ���ü�¼ A,�շ�ϸĿ B,������Ŀ C,����֧����Ŀ D,�����ʻ� E" & _
                " where A.��¼״̬<>0 And Nvl(A.�Ƿ��ϴ�,0)=0 And A.�շ�ϸĿID = B.ID and A.������ĿID = C.ID " & _
                " and A.NO =[2] and A.��¼״̬ = [3] and A.��¼���� = [4] and A.�շ�ϸĿID = D.�շ�ϸĿID " & _
                " and A.����ID=E.����ID And E.����=[1] And D.���� = [1] and A.����ID = [5]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", TYPE_�����ɽ, strNO, int״̬, int����, lng����ID)
    End If
    
    If rsTmp.BOF Then ������ϸ���� = False: Exit Function
    '�����ʼ���ݵĺ���
    strSQL = "select max(Charge_item_no) as Charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
    If rs��ɽ.EOF Then
        i = 1
    Else
        i = Nvl(rs��ɽ("Charge_item_no"), 0) + 1
    End If
    '�𲽽��з�����ϸ����
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then Call ShowWindow(frm�ȴ�.hwnd, 5)
    SetPos frm�ȴ�.hwnd
    frm�ȴ�.Move (Screen.Width - frm�ȴ�.Width) / 2, (Screen.Height - frm�ȴ�.Height) / 2
    DoEvents
    Do While Not rsTmp.EOF
        '�жϼ�¼�Ƿ��Ѿ��ϴ�
        '�������û�д��������ϴ����򲻱ع��������ݣ�������Ϊ���ϴ�
        blnInsert = True
        If Val(Get���ղ���_��ɽ("���õ���")) = 3 And lng��� = 2 Then
            strSQL = "Select Charge_item_no,Nvl(Flag,'0') AS Flag From Check_item_list_interface Where HIS_PK=" & rsTmp!ID
            If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
            rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
            If rs��ɽ.RecordCount = 0 Then
                '�����޴˼�¼���ɼ����ϴ�
            Else
                If rs��ɽ!flag = "1" Then
                    '��ID�ļ�¼�Ѿ��ɹ��ϴ�,��ת
                    GoTo nextRec
                Else
                    '�ϴ�ʧ�ܻ�ҽ���̳���δ��Ӧ�������ٲ�����ϸ��ֱ��ɾ���ϴε������¼�������ٴ��ϴ���Ȼ������µ������¼
                    blnInsert = False
                    j = i
                    i = rs��ɽ!Charge_item_no
                    
                    'ɾ�������
                    strSQL = "Delete check_item_request where Bill_no = '" & strBillNO & _
                             "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "' And charge_Item_NO='" & CStr(i) & "'"
                    gcn��ɽ.Execute strSQL
                End If
            End If
        End If
        
        If blnInsert Then
            '���ύ���ݵ�׼��,���Ϊ���ﲡ�˾ʹ��ݡ�����ID + ʱ�䡱�����ΪסԺ���ˣ��ʹ��ݲ���ID����ҳID
            If InStr(1, ",5,6,7,", "," & rsTmp!�շ���� & ",") <> 0 Then
                strMedi = "1"
            Else
                strMedi = "2"
            End If
            If Val(Get���ղ���_��ɽ("���õ���")) <> 1 Then
                strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                        "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code,App_item_name," & _
                        "Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id" & _
                        IIf(Val(Get���ղ���_��ɽ("���õ���")) = 3, ",HIS_PK)", ")") & _
                        " values('" & strBillNO & "','" & _
                        Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & rsTmp("NO") & "','" & _
                        rsTmp("������") & "',to_date('" & Format(rsTmp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:MM:SS") & "','yyyy-MM-dd HH24:MI:SS'),'" & _
                        rsTmp("ϸĿ����") & "','" & rsTmp("ϸĿ����") & "','" & strMedi & _
                        "','" & rsTmp("���㵥λ") & "'," & rsTmp("����") & "," & CStr(rsTmp("��׼����")) & "," & _
                        CStr(rsTmp("ʵ�ս��")) & ",to_date('" & Format(rsTmp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:MM:SS") & "','yyyy-MM-dd HH24:MI:SS'),'" & _
                        rsTmp("����Ա����") & "'" & IIf(Val(Get���ղ���_��ɽ("���õ���")) = 3, "," & rsTmp!ID & ")", ")")
            Else
                strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                        "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code,App_item_name," & _
                        "Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                        " values('" & strBillNO & "','" & _
                        Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & rsTmp("NO") & "','" & _
                        rsTmp("������") & "','" & rsTmp("�Ǽ�ʱ��") & "','" & _
                        rsTmp("ϸĿ����") & "','" & rsTmp("ϸĿ����") & "','" & strMedi & _
                        "','" & rsTmp("���㵥λ") & "'," & rsTmp("����") & "," & CStr(rsTmp("��׼����")) & "," & _
                        CStr(rsTmp("ʵ�ս��")) & ",'" & rsTmp("�Ǽ�ʱ��") & "','" & _
                        rsTmp("����Ա����") & "')"
            End If
            Call DebugTool(strSQL)
            gcn��ɽ.Execute strSQL
        End If
        
        '�����ύ����
        strSQL = "Insert into Check_item_request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
                strBillNO & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','0')"
        Call DebugTool(strSQL)
        gcn��ɽ.Execute strSQL
        '��ѯ�ύ���
        If Val(Get���ղ���_��ɽ("���õ���")) <> 2 Then
            If frm�ȴ�.Result(2, strBillNO, i) = False Then
                Call DebugTool("������ϸ���ݷ����ж�")
                ������ϸ���� = False
                MsgBox "������ϸ���ݷ����ж�", vbInformation, gstrSysName
                Exit Function
            End If
            strSQL = "select Request_Result,Err_Code,Err_text from check_item_request" & _
                    " where Bill_no = '" & strBillNO & _
                     "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "' and Charge_item_no = '" & _
                     CStr(i) & "'"
            If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
            rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
            If rs��ɽ.BOF Then
                Call DebugTool("Check_Item_Request��¼Ϊ��")
                ������ϸ���� = False
                Exit Function
            Else
                If rs��ɽ("Request_Result") = "1" Then
                    Call DebugTool("��������" & rs��ɽ("Err_Code") & ":" & vbCrLf & String(2, "��") & rs��ɽ("Err_text"))
                    MsgBox "��������" & rs��ɽ("Err_Code") & ":" & vbCrLf & String(2, "��") & rs��ɽ("Err_text"), vbInformation, gstrSysName
                    ������ϸ���� = False
                    Exit Function
                End If
            End If
        End If
'����ñʷ����Ѵ��ݣ�����ת
nextRec:
        '��HIS֮�еĻ������ݽ����޸�
        gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rsTmp("ID") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
        rsTmp.MoveNext
        
        If blnInsert Then
            i = i + 1       '�ۼ�
        Else
            i = j           'δ������ϸ����ԭ��ǰ��ϸ�ۼ�ֵ
        End If
    Loop
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        Do While True
            '��ѯ�ύ���
            strSQL = "select Request_Result,Err_Code,Err_text from " & _
                    "check_item_request where Bill_no = '" & strBillNO & _
                     "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
                     "' and Request_result is Null"
            If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
            rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
            If rs��ɽ.EOF Then Exit Do
            DoEvents
        Loop
        Unload frm�ȴ�
    End If
    '������������ϴ�����
    strSQL = "Delete check_item_request where Bill_no = '" & strBillNO & _
             "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSQL
    
    rs��ɽ.Close
    ������ϸ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ������ϸ���� = False
End Function

Private Function Get����ID(strҽ���� As String, strҽ�����ı��� As String) As String
'���ܣ�ͨ��ҽ�����ĺ����ҽ�����������ID
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select ����ID from �����ʻ� where ���� = [1] and ҽ���� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", CInt(strҽ�����ı���), strҽ����)
    If Not rsTmp.BOF Then
        Get����ID = CStr(rsTmp("����ID"))
    Else
        Get����ID = ""
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Get����ID = ""
End Function

Public Function �������_��ɽ(str����ID As String) As Currency
'���ܣ�ͨ�����˵���Ϣ����������
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strTime As String, rs��ɽ As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    'Modified By ���� ���� 06:06:13
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        '�����������ǭ��������ֱ�Ӵӱ����ʻ��ж�ȡ
        gstrSQL = "Select �ʻ���� ��� From �����ʻ� Where ����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ����", CLng(Val(str����ID)))
        �������_��ɽ = Nvl(rsTmp!���, 0)
    Else
        '���������㲻ͨ����ֱ�ӷ���
        gstrSQL = "select ����,���� from �����ʻ� where ����ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", CLng(Val(str����ID)))
        If rsTmp.BOF Then �������_��ɽ = 0: Exit Function
        '�����ݿ�֮�л�ȡ�ֿ����˵���֤��Ϣ
        strTime = CStr(Format(zlDatabase.Currentdate, "yyyymmddhhmmss")) & "00"
        strSQL = "insert into Check_doex_interface(Bill_no,App_code," & _
                "Ic_id,Doct_flag,Is_bala,Regi_op_id) values('" & strTime & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & _
                rsTmp("����") & rsTmp("����") & "','0','0','" & ToVarchar(UserInfo.����, 8) & "')"
        gcn��ɽ.Execute strSQL
        strSQL = "insert into Check_bill_request(Bill_no,App_code," & _
                "Request_status) values('" & strTime & "','" & Mid(gstrҽԺ����, 1, 4) & _
                "','2')"
        gcn��ɽ.Execute strSQL
        If Checkrequest(strTime) = False Then �������_��ɽ = 0: Exit Function
        '����Ϣ֮����ȡ���˵ĸ����ʻ����
        strSQL = "select Ps_Bala from Check_Doex_Interface where Bill_no = '" & strTime & "'" & _
                " and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
        rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
        If Not rs��ɽ.BOF Then
            �������_��ɽ = IIf(IsNull(rs��ɽ("Ps_Bala")), 0, rs��ɽ("Ps_Bala"))
        Else
            �������_��ɽ = 0
        End If
        strSQL = "delete from Check_bill_request where Bill_no = '" & _
                strTime & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
        strSQL = "delete from Check_doex_interface where Bill_no = '" & _
                strTime & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    �������_��ɽ = 0
End Function

Public Function ��Ժ�Ǽ�_��ɽ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strSQL As String, strInNote As String
    Dim intתԺ As Integer              '2-תԺ;1-��ͨ��Ժ
    Dim rsTmp As New ADODB.Recordset
    
    '������˵������Ϣ
    On Error GoTo errHandle
    gstrSQL = "select A.��Ժ����,A.��Ժ��ʽ,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.����ҽʦ,C.����," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", lng����ID, lng��ҳID)
    intתԺ = 1
    If (rsTmp!��Ժ��ʽ Like "*ת��*" Or rsTmp!��Ժ��ʽ Like "*תԺ*") Then intתԺ = 2
    
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID)   '��Ժ���
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 And gstr���ⲡ�� <> "" Then
        '����Ƿ����ⲡ
        strInNote = gstr���ⲡ��
    End If
    If rsTmp.BOF Then ��Ժ�Ǽ�_��ɽ = False: Exit Function
    '׼�������ύ
    strSQL = "Delete from Check_doex_interface where bill_no='" & lng����ID & "_" & lng��ҳID & "' and App_code='" & Mid(gstrҽԺ����, 1, 4) & "' and Doct_flag=1 and Hosp_No is null"
    gcn��ɽ.Execute strSQL
    strSQL = "Delete from Check_bill_request where bill_no='" & lng����ID & "_" & lng��ҳID & "' and App_code='" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSQL
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        strSQL = "Insert into Check_doex_interface(Bill_no,App_code,Doct_flag," & _
                "Doex_no,In_mode,Ill_type,Ic_id,Is_bala,Regi_op_id,Sec_off,The_bunk," & _
                "In_time,Tre_dr) values('" & lng����ID & "_" & lng��ҳID & _
                "','" & Mid(gstrҽԺ����, 1, 4) & "','1','" & Nvl(rsTmp("סԺ��")) & "_" & lng��ҳID & "','" & intתԺ & "','" & _
                strInNote & "','" & Nvl(rsTmp("����")) & Nvl(rsTmp("����")) & "','0','" & ToVarchar(UserInfo.����, 8) & _
                "','" & Nvl(rsTmp("סԺ����")) & "','" & ToVarchar(Nvl(rsTmp("��Ժ����"), "0"), 24) & "'," & _
                " '" & Nvl(rsTmp("��Ժ����")) & "'" & _
                ",'" & Nvl(rsTmp("����ҽʦ"), "δ֪") & "')"
    Else
        strSQL = "Insert into Check_doex_interface(Bill_no,App_code,Doct_flag," & _
                "Doex_no,In_mode,Ill_type,Ic_id,Is_bala,Regi_op_id,Sec_off,The_bunk," & _
                "In_time,Tre_dr) values('" & lng����ID & "_" & lng��ҳID & _
                "','" & Mid(gstrҽԺ����, 1, 4) & "','1','" & Nvl(rsTmp("סԺ��")) & "_" & lng��ҳID & "','" & intתԺ & "','" & _
                strInNote & "','" & Nvl(rsTmp("����")) & Nvl(rsTmp("����")) & "','0','" & ToVarchar(UserInfo.����, 8) & _
                "','" & Nvl(rsTmp("סԺ����")) & "','" & ToVarchar(Nvl(rsTmp("��Ժ����"), "0"), 24) & "'," & _
                " to_date('" & Format(rsTmp("��Ժ����"), "yyyy-MM-dd HH:MM:SS") & "','yyyy-MM-dd HH24:MI:SS')" & _
                ",'" & Nvl(rsTmp("����ҽʦ"), "δ֪") & "')"
    End If
    gcn��ɽ.Execute strSQL
    '������Ժ����
    strSQL = "Insert into Check_bill_request(Bill_no,App_code,Request_status)" & _
            "values('" & lng����ID & "_" & lng��ҳID & "','" & _
            Mid(gstrҽԺ����, 1, 4) & "','0')"
    gcn��ɽ.Execute strSQL
    '��ѯ����Ľ��
    If Checkrequest(lng����ID & "_" & lng��ҳID) = False Then
        ��Ժ�Ǽ�_��ɽ = False
        Exit Function
    End If
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�����ɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    ��Ժ�Ǽ�_��ɽ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_��ɽ = False
End Function

Public Function ���ʴ���_��ɽ(strNO As String, int���� As Integer, int״̬ As Integer, Optional lng����ID As Long) As Boolean
'��סԺ���˵ķ��ô��ݵ�ҽ������������ͬʱ�޸Ĳ��˷�����Ϣ֮�е�����
    If lng����ID = 0 Then
        ���ʴ���_��ɽ = ������ϸ����(2, , strNO, , int����, int״̬)
    Else
        ���ʴ���_��ɽ = ������ϸ����(2, , strNO, lng����ID, int����, int״̬)
    End If
End Function

Public Function סԺ�������_��ɽ(rs������ϸ As Recordset, lng����ID As Long, strҽ���� As String, str���� As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    Dim cur�����ʻ�֧�� As Currency, cur�����ֽ�֧�� As Currency
    Dim curͳ��֧�� As Currency, cur���֧�� As Currency, cur�����ܶ� As Currency
    Dim strSQL As String, rs��ɽ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset, strBillNO As String
    Dim strMedi As String, strPageId As String
    Dim i As Integer, j As Integer, frm�ȴ� As New frm�ȴ���Ӧ��ɽ
    Dim datCurr As Date, cur�����ʻ���� As Currency
    Dim blnInsert As Boolean
    
    '�ж��Ƿ��Ѿ���������
    If rs������ϸ.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    On Error GoTo errHandle
    '������˵Ĳ�����ҳ��Ҳͬʱ��������㵥��
    gstrSQL = "select max(��ҳID) as ��ҳID from ������ҳ where ����ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", lng����ID)
    strPageId = CStr(rsTmp("��ҳID"))
    strBillNO = CStr(lng����ID) & "_" & CStr(rsTmp("��ҳID"))
    rs������ϸ.Sort = "�Ƿ��ϴ� desc"
    
    '�����ǰ��Ҫ�����
    strSQL = "select max(Charge_item_no) as charge_item_no from " & _
            "Check_item_list_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
    If rs��ɽ.EOF Then
        i = 1
    Else
        i = Nvl(rs��ɽ("Charge_item_no"), 0) + 1
    End If
    rs������ϸ.MoveFirst
    If Val(Get���ղ���_��ɽ("���õ���")) = 3 Then
        '���Ԥ����
        strSQL = "Update Check_bill_request set Request_status = '6',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
        If Checkrequest(strBillNO) = False Then
            סԺ�������_��ɽ = ""
            Exit Function
        End If
    End If
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then Call ShowWindow(frm�ȴ�.hwnd, 5)
    SetPos frm�ȴ�.hwnd

    frm�ȴ�.Move (Screen.Width - frm�ȴ�.Width) / 2, (Screen.Height - frm�ȴ�.Height) / 2
    DoEvents
    
    Do While Not rs������ϸ.EOF
        '������еķ��ý��
        cur�����ʻ�֧�� = cur�����ʻ�֧�� + rs������ϸ("���")
        '������û�û���ϴ����ͽ����ϴ�:ע�⣬�շ�ϸĿ��ν��д��ݻ���Ҫ�޸�
        
        If IIf(IsNull(rs������ϸ("�Ƿ��ϴ�")), "0", rs������ϸ("�Ƿ��ϴ�")) = "0" Then
            gstrSQL = "select A.ID,A.����ʱ��,A.���,A.�շ����,A.NO,A.������,A.�Ǽ�ʱ��," & _
                    "A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��¼����,A.��¼״̬,D.��Ŀ���� as ϸĿ����,B.���� as ϸĿ����,C.����" & _
                    " as ��Ŀ����,B.���㵥λ, (A.���� * A.����) as ����," & _
                    "A.��׼����,A.ʵ�ս��,A.����Ա���� from סԺ���ü�¼ A," & _
                    "�շ�ϸĿ B,������Ŀ C,����֧����Ŀ D where A.�շ�ϸĿID = B.ID and " & _
                    "A.������ĿID = C.ID " & " And A.����ID=[3]" & _
                    " and A.NO =[4] and A.��¼״̬ = [5] and A.��¼���� = [6]" & _
                    " and (A.�۸񸸺� = [2] or A.�۸񸸺� Is Null And A.���=[2])" & _
                    " and (A.�Ƿ��ϴ� = 0 or A.�Ƿ��ϴ� is null) and " & _
                    "A.�շ�ϸĿID = D.�շ�ϸĿID and D.���� = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", TYPE_�����ɽ, CLng(rs������ϸ("���")), lng����ID, CStr(rs������ϸ("NO")), CInt(rs������ϸ("��¼״̬")), CInt(rs������ϸ("��¼����")))
            If Not rsTmp.BOF Then
                '�жϸñʷ����Ƿ��Ѿ��ϴ�
                blnInsert = True
                If Val(Get���ղ���_��ɽ("���õ���")) = 3 Then
                    strSQL = "Select Charge_item_no,Nvl(Flag,'0') AS Flag From Check_item_list_interface Where HIS_PK=" & rsTmp!ID
                    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
                    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
                    If rs��ɽ.RecordCount = 0 Then
                        '�����޴˼�¼���ɼ����ϴ�
                    Else
                        If rs��ɽ!flag = "1" Then
                            '��ID�ļ�¼�Ѿ��ɹ��ϴ�,��ת
                            GoTo nextRec
                        Else
                            '�ϴ�ʧ�ܻ�ҽ���̳���δ��Ӧ�������ٲ�����ϸ��ֱ��ɾ���ϴε������¼�������ٴ��ϴ���Ȼ������µ������¼
                            blnInsert = False
                            j = i
                            i = rs��ɽ!Charge_item_no
                            
                            'ɾ�������
                            strSQL = "Delete check_item_request where Bill_no = '" & strBillNO & _
                                     "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "' And charge_Item_NO='" & CStr(i) & "'"
                            gcn��ɽ.Execute strSQL
                        End If
                    End If
                End If
                
                If blnInsert Then
                    If InStr(1, ",5,6,7,", "," & rsTmp!�շ���� & ",") <> 0 Then
                        strMedi = "1"
                    Else
                        strMedi = "2"
                    End If
                    '���������ύ׼��
                    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
                        strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                                "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                                "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id)" & _
                                " values('" & strBillNO & "','" & _
                                Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & _
                                rsTmp("NO") & "','" & rsTmp("������") & _
                                "','" & rsTmp("�Ǽ�ʱ��") & _
                                "','" & rsTmp("ϸĿ����") & _
                                "','" & rsTmp("ϸĿ����") & "','" & strMedi & "','" & _
                                rsTmp("���㵥λ") & "'," & rsTmp("����") & "," & _
                                CStr(rsTmp("��׼����")) & "," & CStr(rsTmp("ʵ�ս��")) & _
                                ",'" & rsTmp("�Ǽ�ʱ��") & _
                                "','" & rsTmp("����Ա����") & "')"
                    Else
                        strSQL = "Insert into Check_item_list_interface(Bill_no,App_code," & _
                                "Charge_item_no,Reci_no,Dr_code,Reci_date,App_item_code," & _
                                "App_item_name,Dat_medi_flag,Com_unit,Sum_a,App_price,App_mon,Input_date,Input_op_id" & _
                                IIf(Val(Get���ղ���_��ɽ("���õ���")) = 3, ",HIS_PK)", ")") & _
                                " values('" & strBillNO & "','" & _
                                Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','" & _
                                rsTmp("NO") & "','" & rsTmp("������") & _
                                "',to_Date('" & Format(rsTmp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:MM:SS") & _
                                "','yyyy-MM-dd HH24:MI:SS'),'" & rsTmp("ϸĿ����") & _
                                "','" & rsTmp("ϸĿ����") & "','" & strMedi & "','" & _
                                rsTmp("���㵥λ") & "'," & rsTmp("����") & "," & _
                                CStr(rsTmp("��׼����")) & "," & CStr(rsTmp("ʵ�ս��")) & _
                                ",to_date('" & Format(rsTmp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:MM:SS") & _
                                "','yyyy-MM-dd HH24:MI:SS'),'" & rsTmp("����Ա����") & "'" & _
                                IIf(Val(Get���ղ���_��ɽ("���õ���")) = 3, "," & rsTmp!ID & ")", ")")
                    End If
                    Call DebugTool(strSQL)
                    gcn��ɽ.Execute strSQL
                End If
                
                '�����ύ����
                strSQL = "Insert into Check_Item_Request(Bill_no,App_code,Charge_item_no,Request_status) values('" & _
                strBillNO & "','" & Mid(gstrҽԺ����, 1, 4) & "','" & CStr(i) & "','0')"
                Call DebugTool(strSQL)
                gcn��ɽ.Execute strSQL
                '�����ѯ����
                If Val(Get���ղ���_��ɽ("���õ���")) <> 2 Then
                    If frm�ȴ�.Result(2, strBillNO, i) = False Then
                        Call DebugTool("�ڽ���Ĺ���֮�з����ж�")
                        סԺ�������_��ɽ = ""
                        MsgBox "�ڽ���Ĺ���֮�з����ж�", vbInformation, gstrSysName
                        Exit Function
                    End If
                    '��ѯ�ύ���
                    strSQL = "select Request_Result,Err_Code,Err_text from " & _
                            "check_item_request where Bill_no = '" & strBillNO & _
                             "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
                             "' and Charge_item_no = '" & CStr(i) & "'"
                    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
                    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
                    If rs��ɽ.BOF Then
                        Call DebugTool("Check_Item_Request��¼Ϊ��(סԺ�������_��ɽ)")
                        סԺ�������_��ɽ = ""
                        Exit Function
                    Else
                        If rs��ɽ("Request_Result") = "1" Then
                            Call DebugTool("��������[" & rs��ɽ("Err_Code") & "]:" & vbCrLf & String(2, "��") & rs��ɽ("Err_text"))
                            MsgBox "��������[" & rs��ɽ("Err_Code") & "]:" & vbCrLf & String(2, "��") & rs��ɽ("Err_text"), vbInformation, gstrSysName
                            סԺ�������_��ɽ = ""
                            Exit Function
                        End If
                    End If
                End If
'����ü�¼�Ѿ��ϴ�����ת��
nextRec:
                '��HIS֮�еĻ������ݽ����޸�
                gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rsTmp("ID") & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
                
                '�ϴ���ϸ����ۼӴ�����ϸ���
                If blnInsert Then
                    i = i + 1       '�ۼ�
                Else
                    i = j           'δ������ϸ����ԭ��ǰ��ϸ�ۼ�ֵ
                End If
            End If
        End If
        rs������ϸ.MoveNext
    Loop
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
'        Do While True
'            '��ѯ�ύ���
'            strSql = "select Request_Result,Err_Code,Err_text from " & _
'                    "check_item_request where Bill_no = '" & strBillNo & _
'                     "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & _
'                     "' and Request_result is Null"
'            If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
'            rs��ɽ.Open strSql, gcn��ɽ, adOpenStatic, adLockReadOnly
'            If rs��ɽ.EOF Then Exit Do
'            DoEvents
'        Loop
        Unload frm�ȴ�
    End If
    cur�����ܶ� = cur�����ʻ�֧��
    If Val(Get���ղ���_��ɽ("���õ���")) <> 1 Then
        '�����ύ׼��
        datCurr = zlDatabase.Currentdate
        'ȡ���������ʻ�֧��������Ϊ�ܶ����������Ϊ��ϣԤ����ʱ������ͳ��֧������֧����ֻ����ʽ����ʱ�Ÿ����ʻ�֧�����ֽ�֧��
        '�����ʻ�֧��=�ܷ���-ͳ��֧��-���֧��;����ʽ����ʱ����ʵ���ʻ�֧�������Ps_account_pay��Ȼ����ϣ���鲢�������ֶΣ�������У��
        'Ps_account_pay = " & _
                cur�����ʻ�֧�� & ",
        strSQL = "Update Check_doex_interface set Bala_op_id = '" & ToVarchar(UserInfo.����, 8) & _
                "',Out_time =to_date('" & Format(datCurr, "yyyy-MM-dd") & "','yyyy-MM-dd') " & _
                "where Bill_no = '" & strBillNO & "' and App_code = '" & _
                Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
        '���������������,Ŀǰ����֪������Ĳ���ֵ,�ڱ���֮����Ҫ�����޸�
        strSQL = "Update Check_bill_request set Request_status = '2',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
        If Checkrequest(strBillNO) = False Then
            סԺ�������_��ɽ = ""
            Exit Function
        End If
        strSQL = "select Ps_bala from" & _
                " Check_doex_interface where Bill_no = '" & strBillNO & _
                "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
        rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
        cur�����ʻ�֧�� = Nvl(rs��ɽ("Ps_bala"), 0) '�˴�ȡ�����ʻ����
        
        strSQL = "Update Check_bill_request set Request_status = '5',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
        If Checkrequest(strBillNO) = False Then
            סԺ�������_��ɽ = ""
            Exit Function
        End If
    Else
        MsgBox "������ֹ����㣬������ɺ�����ȷ��������......", vbInformation, "ҽҵ���"
    End If
    
    'ȡ��ǰ�ʻ����
    cur�����ʻ���� = �������_��ɽ(CStr(lng����ID))
    '�ӶԷ������ݿ�֮����ȡ�����ʻ�֧�����ֽ�֧����ͳ��֧�������֧������������Ҫ��ȡ�ʻ����ֽ�֧������ϣû�и����������ֶΣ�
    strSQL = "select Ps_account_pay,Ps_cost_pay,Plan_pay,Big_pay from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
    curͳ��֧�� = Nvl(rs��ɽ("Plan_pay"), 0)
    cur���֧�� = Nvl(rs��ɽ("Big_pay"), 0)
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Or Val(Get���ղ���_��ɽ("���õ���")) = 2 Then  '�����ｫ����ҽԺҲ�޸ĳ����ɽ��ͬ�Ĵ���ʽ.HXB
        cur�����ʻ�֧�� = Nvl(rs��ɽ("Ps_account_pay"), 0)            '��ɽ���ظ����ʻ�֧��
        cur�����ֽ�֧�� = Nvl(rs��ɽ("Ps_cost_pay"), 0)
    Else
        '��ͳ�������⣬���������ʻ�֧��
        cur�����ֽ�֧�� = cur�����ܶ� - curͳ��֧�� - cur���֧��
        '�ʻ�֧�����ܴ����ʻ����
        cur�����ʻ�֧�� = IIf(cur�����ʻ�֧�� > cur�����ʻ����, cur�����ʻ����, cur�����ʻ�֧��)
        '�����ʻ����֧������
        cur�����ʻ�֧�� = IIf(cur�����ʻ�֧�� >= cur�����ֽ�֧��, cur�����ֽ�֧��, cur�����ʻ�֧��)
        cur�����ֽ�֧�� = cur�����ֽ�֧�� - cur�����ʻ�֧��
'        '�����ʻ�֧����ڽ��㴦�Ѹ��£�
'        strSql = "Update Check_doex_interface set Ps_account_pay = '" & cur�����ʻ�֧�� & "'" & _
'                "where Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
'        gcn��ɽ.Execute strSql
    End If
    
'    '������������ʻ�֧�����
'    If Val(Get���ղ���_��ɽ("���õ���")) <> 1 Then
'        'cur�����ܶ� = cur�����ܶ� - curͳ��֧�� - cur���֧��
'        cur�����ʻ�֧�� = IIf(cur�����ʻ�֧�� > cur�����ܶ�, cur�����ܶ�, cur�����ʻ�֧��)
'    End If
'    gstrSQL = "Select Nvl(�ʻ����,0) ��� From �����ʻ� Where ����ID=" & lng����ID
'    Call OpenRecordset(rsTmp, "��ȡ�ʻ����")
'    cur�����ʻ���� = rsTmp!���
    
    If (Val(Get���ղ���_��ɽ("���õ���")) = 3 Or Val(Get���ղ���_��ɽ("���õ���")) = 0) Then
        סԺ�������_��ɽ = "�����ʻ�;" & cur�����ʻ�֧�� & ";1"
    Else
        סԺ�������_��ɽ = "�����ʻ�;" & cur�����ʻ�֧�� & ";0" '�������޸ĸ����ʻ�
    End If
    סԺ�������_��ɽ = סԺ�������_��ɽ & "|ͳ��֧��;" & curͳ��֧�� & ";0" '�������޸�ͳ��֧��
    סԺ�������_��ɽ = סԺ�������_��ɽ & "|���ͳ��;" & cur���֧�� & ";0" '�������޸Ĵ��֧��
    
    '������������ϴ�����
    strSQL = "Delete check_item_request where Bill_no = '" & strBillNO & _
             "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSQL
    
    With pre_Balance
        .cur�󲡻��� = cur���֧��
        .cur�����ʻ� = cur�����ʻ�֧��
        .curҽ������ = curͳ��֧��
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Resume
    סԺ�������_��ɽ = ""
End Function

Public Function סԺ����_��ɽ(lng����ID As Long, ByVal lng����ID As Long, Optional ByRef strAdvance As String = "") As Boolean
'�����˵ķ��ý��н��㣬���ڱ�ɽҽ������Ҫ���г�Ժ�Ǽǣ���˲����г�Ժ�Ǽ�
    Dim rsTmp As New ADODB.Recordset, cur������ As Currency
    Dim strBillNO As String, strSQL As String, datCurr As Date
    Dim rs��ɽ As New ADODB.Recordset, cur�����ʻ�֧�� As Currency
    Dim cur�����ֽ�֧�� As Currency, curͳ��֧�� As Currency
    Dim cur���֧�� As Currency, intסԺ�����ۼ� As Integer
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim curͳ���Ը� As Currency, cur�����Ը� As Currency
    Dim cur�����Ը� As Currency, cur��ͳ�� As Currency
    Dim cur���Ը� As Currency, cur���� As Currency
    Dim curȫ�Ը� As Currency, cur�ҹ��Ը� As Currency
    Dim cur�����ʻ� As Currency, str�������� As String
    Dim bln��;���� As Boolean
    Dim str���㷽ʽ As String
    Dim lng��ҳID As Long
    Dim blnRevise As Boolean, blnOld As Boolean
    
    On Error GoTo errHandle
    Call DebugTool("׼�����н��㣬��ǰ����ID��" & lng����ID)
    bln��;���� = Not IS��Ժ(lng����ID)
    Call DebugTool("�����ʻ��ĵ�ǰ״̬��ʾ�ò��˵�ǰ��" & IIf(bln��;����, "��Ժ���������н�", "��Ժ�������г�Ժ����"))
    
    '��ȡ���θ����ʻ�֧����
    gstrSQL = "Select Nvl(A.��Ԥ��,0) �����ʻ� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=" & TYPE_�����ɽ & _
        " And A.���㷽ʽ in ('�����ʻ�') And A.����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���θ����ʻ�֧����", lng����ID)
    cur�����ʻ�֧�� = 0
    If Not rsTmp.EOF Then
        cur�����ʻ�֧�� = rsTmp!�����ʻ�
        pre_Balance.cur�����ʻ� = cur�����ʻ�֧��
    End If
    
    gstrSQL = "select sum(ʵ�ս��) as ������,sum(���ʽ��) as �ѽ��� from סԺ���ü�¼ where " & _
            "����ID=[2] and ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", lng����ID, lng����ID)
    cur������ = Nvl(rsTmp("�ѽ���"), 0)
    gstrSQL = "select ��ҳID,��Ժ���� from ������ҳ where ��ҳID=(select max(��ҳID) from " & _
            "������ҳ where ����ID  = [1]) and ����ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", lng����ID)
    If rsTmp.BOF Then Exit Function
    strBillNO = lng����ID & "_" & rsTmp("��ҳID")
    lng��ҳID = rsTmp!��ҳID
    If Val(Get���ղ���_��ɽ("���õ���")) <> 1 Then
        '�����ύ׼��
        strSQL = "Update Check_doex_interface set Ps_account_pay = " & cur�����ʻ�֧�� & _
                ",Bala_op_id = '" & ToVarchar(UserInfo.����, 8) & "',Out_time = to_date('" & _
                Format(rsTmp("��Ժ����"), "yyyy-MM-dd") & "','yyyy-MM-dd') where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
        '���н�������
        If bln��;���� = False Then
            Call DebugTool("��Ժ����")
            strSQL = "Update Check_bill_request set Request_status = '1',Request_Result=null where " & _
                    "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        Else
            Call DebugTool("��;����")
            strSQL = "Update Check_Doex_interface Set Doct_Flag='1',Doex_Type='5' Where Bill_no='" & strBillNO & "' And App_code='" & Mid(gstrҽԺ����, 1, 4) & "'"
            gcn��ɽ.Execute strSQL
            strSQL = "Update Check_bill_request set Request_status = '6',Request_Result=null where " & _
                    "Bill_no = '" & strBillNO & "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
        End If
        gcn��ɽ.Execute strSQL
        If Checkrequest(strBillNO) = False Then סԺ����_��ɽ = False: Exit Function
    End If
    '�������
    'modify by ccy, add select field Ps_bala
    strSQL = "select Ps_bala,Ps_account_pay,Ps_cost_pay,Plan_pay,Big_pay,acc_cyc from" & _
            " Check_doex_interface where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'"
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close
    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly
    cur�����ʻ�֧�� = Nvl(rs��ɽ("Ps_account_pay"), 0)
    cur�����ֽ�֧�� = Nvl(rs��ɽ("Ps_cost_pay"), 0)
    curͳ��֧�� = Nvl(rs��ɽ("Plan_pay"), 0)
    cur���֧�� = Nvl(rs��ɽ("Big_pay"), 0)
    cur��ͳ�� = cur���֧��
    curȫ�Ը� = cur�����ʻ�֧��
    cur�����ʻ� = Nvl(rs��ɽ("Ps_bala"), 0)
    str�������� = Nvl(rs��ɽ("ACC_CYC"), "")
    
    '�Ƚ������������ʽ�������Ƿ�һ��
    If Not (cur�����ʻ�֧�� = pre_Balance.cur�����ʻ� And curͳ��֧�� = pre_Balance.curҽ������ And _
        cur���֧�� = pre_Balance.cur�󲡻���) Then
        blnRevise = True
        #If gverControl < 2 Then
            blnOld = True
        #End If
    End If
    
    If blnRevise Then
        str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�֧��
        str���㷽ʽ = str���㷽ʽ & "||ͳ��֧��|" & curͳ��֧��
        str���㷽ʽ = str���㷽ʽ & "||���ͳ��|" & cur��ͳ��
        If str���㷽ʽ <> "" Then
            str���㷽ʽ = Mid(str���㷽ʽ, 3)
            If blnOld Then
                gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
            Else
                strAdvance = str���㷽ʽ
                gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
        End If
    End If
    
    '��д�����
    datCurr = zlDatabase.Currentdate
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_�����ɽ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
            
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�����ɽ & "," & Year(datCurr) & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & _
            cur����ͳ���ۼ� + curͳ��֧�� + curͳ���Ը� + cur�����Ը� + cur�����Ը� + cur��ͳ�� + cur���Ը� & "," & _
            curͳ�ﱨ���ۼ� + curͳ��֧�� + cur��ͳ�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�����ɽ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & ",NULL," & cur�����Ը� & "," & _
        cur������ & "," & cur�����ֽ�֧�� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & cur�����ʻ�֧�� & _
        ",NULL," & lng��ҳID & "," & IIf(bln��;����, "1", "0") & ",'" & strBillNO & "'" & _
        IIf(blnOld, "", IIf(blnRevise, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
'        gstrSQL = "zl_�������ڼ�¼_insert(" & lng����ID & ",'" & str�������� & "'," & cur������ & "," & cur�����ʻ�֧�� & "," & curͳ��֧�� & ",'N',to_date('" & datCurr & "','yyyy-mm-dd HH:MI:SS'))"
        gstrSQL = "Insert into zlhis.�������ڼ�¼ values (" & lng����ID & ",'" & str�������� & "'," & cur������ & "," & cur�����ʻ�֧�� & "," & curͳ��֧�� + cur���֧�� & ",'N',to_date('" & Format(datCurr, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
        gcnOracle.Execute gstrSQL
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "������ҽ��")
    End If
    '�������סԺ����
    If bln��;���� = False Then
        strSQL = "Delete from Check_bill_request where bill_no='" & strBillNO & "' and App_code='" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
    End If
    
    סԺ����_��ɽ = True
    'modify by ccy
    If Val(Get���ղ���_��ɽ("���õ���")) = 1 Then
        Err.Raise 9000, gstrSysName, "���ĸ����ʻ����Ϊ[" & Format(cur�����ʻ�, "0.00") & "Ԫ]", vbInformation, "סԺ����"
    End If
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    סԺ����_��ɽ = False
End Function

Public Function סԺ�������_��ɽ(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, str�������� As String
    Dim curDate As Date, strBillNO As String, strSQL As String
        
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "Select * From סԺ���ü�¼ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵķ��ü�¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    strBillNO = rsTemp!����ID & "_" & rsTemp!��ҳID
    'ɾ�����ܴ��ڵ�����
    strSQL = "Delete from Check_bill_request where bill_no='" & strBillNO & "' and App_code='" & Mid(gstrҽԺ����, 1, 4) & "'"
    gcn��ɽ.Execute strSQL
    '���²��������Ա��Ժ����
    strSQL = "Insert into Check_bill_request(Bill_no,App_code,Request_status)" & _
            "values('" & strBillNO & "','" & _
            Mid(gstrҽԺ����, 1, 4) & "','9')"
    gcn��ɽ.Execute strSQL
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", TYPE_�����ɽ, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
'    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
'    str��ˮ�� = NVL(rsTemp("֧��˳���"), "0")
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_�����ɽ, rsTemp("����ID"), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & TYPE_�����ɽ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�����ɽ & "," & rsTemp("����ID") & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0,0," & _
        rsTemp("�����ʻ�֧��") * -1 & ",Null," & rsTemp("��ҳID") & "," & rsTemp("��;����") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    
    If Val(Get���ղ���_��ɽ("���õ���")) = 2 Then
        gstrSQL = "Select * from �������ڼ�¼ where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
        If Not rsTemp.EOF Then
            str�������� = rsTemp!��������
    '        gstrSQL = "zl_�������ڼ�¼_insert(" & lng����ID & ",'" & str�������� & "'," & NVL(rsTemp("�������ý��"), 0) * -1 & "," & NVL(rsTemp("�����ʻ�֧��"), 0) * -1 & "," & NVL(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",'N',to_date('" & curDate & "','yyyy-mm-dd HH:MI:SS'))"
    '        Call zlDatabase.ExecuteProcedure(gstrSQL, "������ҽ��")
            gstrSQL = "Insert into zlhis.�������ڼ�¼ values (" & lng����ID & ",'" & str�������� & "'," & Nvl(rsTemp("�ܶ�"), 0) * -1 & "," & Nvl(rsTemp("����"), 0) * -1 & "," & Nvl(rsTemp("ͳ��"), 0) * -1 & ",'N',to_date('" & Format(curDate, "yyyy-MM-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'))"
            gcnOracle.Execute gstrSQL
        End If
    End If

    סԺ�������_��ɽ = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽ�_��ɽ(lng����ID As Long, lng��ҳID As Long, Optional ByVal blnת��ͨ���� As Boolean = False) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    Dim strSQL As String, rs��ɽ As New ADODB.Recordset
    Dim strBillNO As String, rsTmp As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim bln����ó�Ժ As Boolean
    
    On Error GoTo errHandle
    '���ô�סԺ�Ƿ�û�з��÷���
    If blnת��ͨ���� = False Then
        gstrSQL = "Select sum(ʵ�ս��) as ���  from סԺ���ü�¼ where ����ID=[1] and ��ҳID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���˳�Ժ", lng����ID, lng��ҳID)
        If rsTemp.EOF = True Then
            bln����ó�Ժ = True
        Else
            bln����ó�Ժ = (Nvl(rsTemp("���"), 0) = 0)
        End If
    Else
        '�ɱ����ʻ��ĳ���ҽ����Ժ���ܵ����������Բ���Ҫ��HIS���з��ö��ѳ�����ֻ��ȡ��ҽ�����
        bln����ó�Ժ = True
    End If
    
    If bln����ó�Ժ = True Then
        '��������ó�Ժ���ͽ��䴦��Ϊ����Ժ�������ø�����סԺ��Ϣ
        gstrSQL = "select ��Ժ���� from ������ҳ where ����ID = " & lng����ID & _
                " and ��ҳID=" & lng��ҳID
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", lng����ID, lng��ҳID)
        strBillNO = lng����ID & "_" & lng��ҳID
        '���г�Ժ����
        strSQL = "Update Check_bill_request set Request_status= '3',Request_Result=null where " & _
                "Bill_no = '" & strBillNO & "' and App_code = '" & _
                Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
        '��ѯ������
        If Checkrequest(strBillNO) = False Then ��Ժ�Ǽ�_��ɽ = False: Exit Function
        
        'ɾ�����ε���Ժ�Ǽ���Ϣ
        strSQL = "Delete from Check_doex_interface where bill_no='" & lng����ID & "_" & lng��ҳID & "' and App_code='" & Mid(gstrҽԺ����, 1, 4) & "' and Doct_flag=1"
        gcn��ɽ.Execute strSQL
        strSQL = "Delete from Check_bill_request where bill_no='" & lng����ID & "_" & lng��ҳID & "' and App_code='" & Mid(gstrҽԺ����, 1, 4) & "'"
        gcn��ɽ.Execute strSQL
    End If
    '��HIS֮�еĻ������ݽ����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�����ɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    ��Ժ�Ǽ�_��ɽ = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    ��Ժ�Ǽ�_��ɽ = False
End Function

Public Function ��Ժ�Ǽǳ���_��ɽ(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�����ɽ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽҽ��")
    ��Ժ�Ǽǳ���_��ɽ = True
End Function

Public Function Checkrequest(strBillNO As String) As Boolean
'���ܣ��ж��Ƿ��ܹ������ȷ�Ĳ�����Ϣ
    Dim strSQL As String, rs��ɽ As New ADODB.Recordset
    Dim strResult As String '����Ľ��
    Dim strTmp As String, strError As String
    Dim frm�ȴ� As New frm�ȴ���Ӧ��ɽ, lngErrLine As Long
    
    On Error GoTo errHandle
    '�ύ���󣬽��в�ѯ
    If frm�ȴ���Ӧ��ɽ.Result(1, strBillNO) = False Then
        Checkrequest = False: lngErrLine = 1
        Unload frm�ȴ���Ӧ��ɽ
        DoEvents
        Exit Function
    End If
    Unload frm�ȴ���Ӧ��ɽ
    '���ݷ��صķ��ص�ֵ�жϽ��
    strSQL = "Select Request_Result,Err_text from " & _
            "Check_bill_request where Bill_no = '" & strBillNO & _
            "' and App_code = '" & Mid(gstrҽԺ����, 1, 4) & "'": lngErrLine = 2
    If rs��ɽ.State = adStateOpen Then rs��ɽ.Close: lngErrLine = 3
    rs��ɽ.Open strSQL, gcn��ɽ, adOpenStatic, adLockReadOnly: lngErrLine = 4
    If Not rs��ɽ.BOF Then
        strTmp = Nvl(rs��ɽ("Request_Result"), 0): lngErrLine = 5
        strError = Nvl(rs��ɽ("Err_text"), ""): lngErrLine = 6
    Else
        Exit Function
    End If
    Select Case strTmp
        Case "0"
            Err.Raise 9000, gstrSysName, "û�������������������", vbInformation, gstrSysName
            Checkrequest = False
            Exit Function
        Case "1"
            If strError <> "" Then
                Err.Raise 9000, gstrSysName, "ҽ���ӿڵ��ó�����������" & vbCrLf & vbCrLf & strError, vbInformation, gstrSysName
            Else
                Err.Raise 9000, gstrSysName, "ҽ���ӿڵ��ó��ִ���", vbInformation, gstrSysName
            End If
            Exit Function
        Case "9"
            Checkrequest = True
    End Select
    Checkrequest = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description & vbCrLf & "�ڹ���[CheckRequest]�еڡ�" & lngErrLine & "��", vbInformation, Err.Source
    Err.Clear
    Checkrequest = False
End Function

Public Function Get���ղ���_��ɽ(ByVal str������ As String) As String
'���ܣ���ñ��ղ���
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.������,A.����ֵ from ���ղ��� A " & _
              " where A.������=[1] and A.����=[2] and A.���� is null "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ɽҽ��", str������, TYPE_�����ɽ)
    
    If rsTemp.EOF = False Then
        Get���ղ���_��ɽ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
    End If
End Function

Public Sub SetPos(lHwnd As Long, Optional TopFlag As Boolean = True)
    If TopFlag Then
        SetWindowPos lHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    Else
        SetWindowPos lHwnd, HWND_NOTTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If
End Sub

Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String, ByVal strFileName As String) As String
    Dim ResultString As String * 144, Temp As Integer, s As String, i As Integer
    Temp% = GetprivateprofileString(SectionName, KeyWord, "", ResultString, 144, strFileName)
    '�����ؼ��ʵ�ֵ
    If Temp% > 0 Then '�ؼ��ʵ�ֵ��Ϊ��
        s = ""
        For i = 1 To 144
            If asc(Mid$(ResultString, i, 1)) = 0 Then
                Exit For
            Else
                s = s & Mid$(ResultString, i, 1)
            End If
        Next
    Else
        s = DefString
    End If
    GetIniS = Trim(s)
End Function

Private Function IS��Ժ(ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    '�жϲ����Ƿ��ѳ�Ժ
    gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
    If Split(rsTmp!�汾��, ".")(0) = 10 And Split(rsTmp!�汾��, ".")(1) >= 34 Then
        gstrSQL = " Select B.��Ժ���� From ������Ϣ A,������ҳ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.����ID=[1]"
    Else
        gstrSQL = " Select B.��Ժ���� From ������Ϣ A,������ҳ B" & _
            " Where A.����ID=B.����ID And A.סԺ����=B.��ҳID And A.����ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�жϲ����Ƿ��Ѿ���Ժ", lng����ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    If IsNull(rsTemp!��Ժ����) Then Exit Function
    IS��Ժ = True
End Function
