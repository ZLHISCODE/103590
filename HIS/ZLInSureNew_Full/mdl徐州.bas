Attribute VB_Name = "mdl����"
Option Explicit

Declare Function dysEncrypt Lib "ApiDll.DLL" (ByVal strPass As String) As String
Declare Function init_com% Lib "SURE32WC.DLL" (ByVal str As Long)
Declare Function close_com% Lib "SURE32WC.DLL" ()
Declare Function sele_card% Lib "SURE32WC.DLL" (ByVal crdno As Long)
Declare Function power_on% Lib "SURE32WC.DLL" ()
Declare Function power_off% Lib "SURE32WC.DLL" ()
Declare Function rd_str% Lib "SURE32WC.DLL" (ByVal apz As Long, ByVal Address As Long, ByVal Length As Long, ByVal Buffer$)
Declare Function wr_str% Lib "SURE32WC.DLL" (ByVal apz As Long, ByVal Address As Long, ByVal Length As Long, ByVal Buffer$)
Declare Function chk_sc% Lib "SURE32WC.DLL" (ByVal apz As Long, ByVal Length As Long, ByVal Buffer$)

Public gcn���� As New ADODB.Connection, intCOM���� As Integer
Private mcur����֧�� As Currency, mcurͳ��֧�� As Currency, mcur�ܶ� As Currency, mint����סԺ���� As Integer, _
        mcur����Ա As Currency, mcurȫ���� As Currency, mcur�������� As Currency, mcurҽ����֧�� As Currency, _
        mstr���ʷ��� As String

Public Function WriteCard(strState As String) As Boolean
    Dim lngReturn As Long, strReturn As String, strErrInfo As String, strInfo() As String
    lngReturn = init_com(intCOM����)
    If lngReturn <> 0 Then
        MsgBox "��ʼ���˿ڴ���", vbInformation, "����"
        Exit Function
    End If
    
    lngReturn = sele_card(43)
    If lngReturn <> 0 Then
        MsgBox "���忨���ʹ���", vbInformation, "����"
        GoTo powerOFF
    End If
    
    If power_on() <> 0 Then
        MsgBox "���ϵ����", vbInformation, "����"
        GoTo powerOFF
    End If
    
    strReturn = Space(129)
    lngReturn = rd_str(1, 0, 128, strReturn)
    If lngReturn <> 0 Then
        MsgBox "��ȡ����Ϣ����", vbInformation, "����"
        GoTo powerOFF
    End If
    strInfo = Split(Trim(strReturn), "@")
    strReturn = ""
    For lngReturn = 0 To 11
        strReturn = strReturn & IIf(strReturn <> "", "@", "") & strInfo(lngReturn)
    Next
    
    strReturn = "FFFF"
    lngReturn = chk_sc(0, 2, strReturn)
    If lngReturn <> 0 Then
        strErrInfo = "У�鿨ʧ��"
        Select Case lngReturn
            Case 2
                strErrInfo = strErrInfo & "-�޿�"
            Case 3
                strErrInfo = strErrInfo & "-δ�ϵ�"
            Case 4
                strErrInfo = strErrInfo & "-���ڴ���"
            Case 9
                strErrInfo = strErrInfo & "-���ݳ��ȴ���"
            Case 11
                strErrInfo = strErrInfo & "-�������"
            Case 14
                strErrInfo = strErrInfo & "-������"
        End Select
        MsgBox strErrInfo, vbInformation, "����"
        GoTo powerOFF
    End If
    
    strInfo(3) = strState
    If InStr(strInfo(11), Chr(0)) > 0 Then strInfo(11) = Left(strInfo(11), InStr(strInfo(11), Chr(0)) - 1)
    strInfo(11) = strInfo(11) & "@"
    strReturn = ""
    For lngReturn = 0 To 11
        strReturn = strReturn & IIf(strReturn <> "", "@", "") & strInfo(lngReturn)
    Next
    lngReturn = wr_str(1, 0, 200, strReturn)
    If lngReturn <> 0 Then
        MsgBox "д������ʧ��", vbInformation, "д��"
        GoTo powerOFF
    End If
    
    WriteCard = True
powerOFF:
    Call power_off
    Call close_com
End Function

Public Function MakeTransNO() As String
    Randomize
    MakeTransNO = Format(Date, "yymmdd") & Format(Time, "hhmmss") & Format(900099 * Rnd + 1, "0#####")
End Function

Public Function Getҽ����(lng����ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����", lng����ID)
    If rsTemp.EOF Then
        Getҽ���� = ""
    Else
        Getҽ���� = Nvl(rsTemp!ҽ����, "")
    End If
End Function

Public Function openConn����() As Boolean
    Dim rsTemp As New ADODB.Recordset, str����ֵ As String, strUser As String, strServer As String, _
        strPass As String, strDatabase As String
    On Error GoTo errHandle
    If gcn����.State <> adStateOpen Then
        gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����)
        
        Do Until rsTemp.EOF
            str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Select Case rsTemp("������")
                Case "�����û���"
                    strUser = str����ֵ
                Case "���ݷ�����"
                    strServer = str����ֵ
                Case "�����û�����"
                    strPass = str����ֵ
                Case "�������ݿ�"
                    strDatabase = str����ֵ
            End Select
            rsTemp.MoveNext
        Loop
        
        intCOM���� = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", 0)
        
        On Error Resume Next
        gcn����.ConnectionString = "Provider=MSDASQL.1;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
        gcn����.CursorLocation = adUseClient
        gcn����.Open
        
        If Err <> 0 Then
            MsgBox "ҽ��ǰ�÷���������ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    openConn���� = True
    Exit Function

errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ʼ��_����() As Boolean
    Dim rsTemp As New ADODB.Recordset, str����ֵ As String, strUser As String, strServer As String, _
        strPass As String, strDatabase As String
    
    If openConn����() = False Then Exit Function
    
    gstrSQL = "Select * From ������� Where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ����ʼ��")
    gstrҽԺ���� = Trim(rsTemp!ҽԺ����)
    
    ҽ����ʼ��_���� = True
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
'����:ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'����:bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'����:�ջ���Ϣ��
'ע��:1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmIDentified As New frmIdentify����, strPatiInfo As String
    
    WriteInfo vbCrLf & "��ʼ�����֤"
    
    strPatiInfo = frmIDentified.GetPatient(bytType)
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ
        lng����ID = BuildPatiInfo(bytType, strPatiInfo, lng����ID, TYPE_����)
        
        '���ظ�ʽ:�м���벡��ID
        strPatiInfo = frmIDentified.mstrPatient & lng����ID & ";" & frmIDentified.mstrOther
        mint����סԺ���� = frmIDentified.mintסԺ����
        Unload frmIDentified
    Else
        ��ݱ�ʶ_���� = ""
        MsgBox "ҽ��������Ϣ��ȡʧ��", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    
    WriteInfo "���������֤"
    
    ��ݱ�ʶ_���� = strPatiInfo
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_���� = ""
End Function

Public Function �������_����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, TYPE_����)
    
    If rsTemp.EOF Then
        �������_���� = 0
    Else
        �������_���� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'����:rsDetail     ������ϸ(����)
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
'�ֶ�:����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim curȫ���� As Currency, cur�������� As Currency, cur�ܶ� As Currency, strҽ���� As String, _
        strReturn As String, cur�ֽ� As String, rsTemp As New ADODB.Recordset, strTransNO As String, _
        strSQL As String, bln��ζ��ҩ As Boolean, cur��ҩ As Currency, cur��ζ��ҩ As Currency, _
        blnISҩƷ As Boolean, lng����ID As Long, strPara As String, i As Long
        
    On Error GoTo errHandle
    WriteInfo vbCrLf & "��ʼ����Ԥ����"
    
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�в��˷��ü�¼�����ܽ��н���", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = rs��ϸ!����ID
    
    While Not rs��ϸ.EOF
        gstrSQL = "Select A.���,B.��Ŀ����,B.��Ŀ���� From �շ�ϸĿ A,����֧����Ŀ B Where A.ID=B.�շ�ϸĿID " & _
            "And A.ID=[1] And B.�Ƿ�ҽ��=1 And B.����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID), TYPE_����)
        
        bln��ζ��ҩ = True
        If rsTemp.EOF Then
            curȫ���� = curȫ���� + rs��ϸ!ʵ�ս��
        Else
            If rsTemp!��� = "7" Then
                If cur��ҩ <> 0 Then bln��ζ��ҩ = False
                cur��ҩ = cur��ҩ + rs��ϸ!ʵ�ս��
            End If
            If rsTemp!��� = "5" Or rsTemp!��� = "6" Or rsTemp!��� = "7" Then
                blnISҩƷ = True
                strSQL = "Select * From mi_drug_trade_list Where trade_code='" & rsTemp!��Ŀ���� & "' And cancel_sign<>'1'"
            Else
                blnISҩƷ = False
                strSQL = "Select * From mi_dt_item Where item_code='" & rsTemp!��Ŀ���� & "' And cancal_sign<>'1'"
            End If
            Set rsTemp = gcn����.Execute(strSQL)
            If rsTemp.EOF Then
                curȫ���� = curȫ���� + rs��ϸ!ʵ�ս��
            Else
                If Trim(rsTemp!mi_class = 2) Then
                    If blnISҩƷ = True Then
                        cur�������� = cur�������� + rs��ϸ!ʵ�ս�� * 0.2
                    Else
                        cur�������� = cur�������� + rs��ϸ!ʵ�ս�� * Nvl(rsTemp!self_rate, 0) / 100
                    End If
                ElseIf Trim(rsTemp!mi_class) = 4 Then
                    cur��ζ��ҩ = cur��ζ��ҩ + rs��ϸ!ʵ�ս��
                ElseIf Trim(rsTemp!mi_class) <> 1 Then
                    curȫ���� = curȫ���� + rs��ϸ!ʵ�ս��
                End If
            End If
        End If
        cur�ܶ� = cur�ܶ� + rs��ϸ!ʵ�ս��
        rs��ϸ.MoveNext
    Wend
    If bln��ζ��ҩ = True Then curȫ���� = curȫ���� + cur��ζ��ҩ
    
    If cur�ܶ� = 0 Then
        MsgBox "û�в������˷��ã����ܽ��н���", vbInformation, gstrSysName
        Exit Function
    End If
    
    strҽ���� = Getҽ����(lng����ID)
    strTransNO = MakeTransNO()
    strPara = strҽ���� & "," & cur�ܶ� & "," & Format(cur��������, "0.00") & "," & Format(curȫ����, "0.00")
    
    WriteInfo "��������:��ˮ��---" & strTransNO
    WriteInfo "���������� ����---" & strPara
    
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','60','" & UserInfo.��� & "','" & strPara & "','9')"
    gcn����.Execute strSQL
    If frm�ȴ���Ӧ����.Result(strTransNO, strReturn) = False Then
        WriteInfo "������ֹ"
        MsgBox "������ֹ", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "���󷵻�:" & strReturn
    '0-�ɹ���־��1-ҽ����֧�����ã�2-�ʻ�֧�����ã�3-ͳ��֧�����ã�4-����Ա����֧������
    If InStr(strReturn, ",") > 0 Then
        If Split(strReturn, ",")(0) = "01" Then
            MsgBox "ҽ������ʧ��", vbInformation, gstrSysName
            Exit Function
        Else
            mcurҽ����֧�� = CCur(Split(strReturn, ",")(1))
            mcur����֧�� = CCur(Split(strReturn, ",")(2))
            mcurͳ��֧�� = CCur(Split(strReturn, ",")(3))
            mcur����Ա = CCur(Split(strReturn, ",")(4))
        End If
    Else
        MsgBox "ҽ������ʧ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    mcur�ܶ� = cur�ܶ�: mcur�������� = CCur(Format(cur��������, "0.00")): mcurȫ���� = curȫ����
'    mstr���ʷ��� = Right(Space(15) & mcur�ܶ�, 15)
'    For i = 1 To 9
'        mstr���ʷ��� = mstr���ʷ��� & Right(Space(15) & Split(strReturn, ",")(i), 15)
'    Next
    
    If mcur����֧�� <> 0 Then str���㷽ʽ = "�����ʻ�;" & mcur����֧�� & ";0"
    If mcurͳ��֧�� <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ <> "", "|", "") & "ͳ�����;" & mcurͳ��֧�� & ";0"
    If mcur����Ա <> 0 Then str���㷽ʽ = str���㷽ʽ & IIf(str���㷽ʽ <> "", "|", "") & "����Ա����;" & mcur����Ա & ";0"
    
    WriteInfo "��������Ԥ����:" & str���㷽ʽ
    �����������_���� = True
    Exit Function
    
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur����֧�� As Currency, strҽ���� As String, curȫ�Ը� As Currency, cur���Ը� As Currency, curҽ������ As Currency) As Boolean
    Dim str���� As String, lng����ID As String, rsTemp As New ADODB.Recordset, STR���� As String, _
        strSQL As String, strReturn As String, strPara As String, strTransNO As String, _
        intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, _
        cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, datCurr As Date, cur��� As Currency
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    WriteInfo vbCrLf & "��ʼ�������"
    '����ţ���Ʊ�ţ�ҽ���ţ������ܷ��ã�����������ã��������20%������ȫ������ã�����
    gstrSQL = "Select * From �����ʻ� Where ҽ����=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҽ����Ϣ", strҽ����, TYPE_����)
    
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "û���ҵ��ò��˵�ҽ����Ϣ"
        Exit Function
    End If
    cur��� = Nvl(rsTemp!�ʻ����, 0)
    str���� = Nvl(rsTemp!����, "666666")
    lng����ID = rsTemp!����ID
    
    gstrSQL = "Select * From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", lng����ID)
    STR���� = rsTemp!����
    
    strPara = lng����ID & "," & lng����ID & "," & strҽ���� & "," & mcur�ܶ� & "," & mcur�������� & "," & _
        mcurȫ���� & "," & dysEncrypt(str����)
    strTransNO = MakeTransNO()
    
    WriteInfo "��������:��ˮ��---" & strTransNO
    WriteInfo "���������� ����---" & strPara
    
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','61','" & UserInfo.��� & "','" & strPara & "','9')"
    gcn����.Execute strSQL
    If frm�ȴ���Ӧ����.Result(strTransNO, strReturn) = False Then
        WriteInfo "������ֹ"
        Err.Raise 9000, gstrSysName, "������ֹ"
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        Err.Raise 9000, gstrSysName, "ҽ������ʧ��"
        Exit Function
    End If
    
    WriteInfo "���󷵻�:" & strReturn

    'ҽԺ��,�����,��Ʊ��,ҽ����,����,�����ܷ���,��ȫ�������,�����������,ҽ����֧������,ͳ��֧��,����Ա֧��,�ʻ�֧��,ʱ��
    strSQL = "Insert Into hospital_clinic_payment (hospital_no,clinic_no,invoice_no,medical_card_no,name," & _
        "clinic_expense,full_self_expense,part_self_expense,mi_unpayment,social_payment,servant_payment," & _
        "account_payment,exectime) Values ('" & Trim(gstrҽԺ����) & "','" & lng����ID & "','" & lng����ID & "'," & _
        "'" & strҽ���� & "','" & STR���� & "'," & mcur�ܶ� & "," & mcurȫ���� & "," & mcur�������� & "," & _
        mcurҽ����֧�� & "," & mcurͳ��֧�� & "," & mcur����Ա & "," & mcur����֧�� & ",'" & Format(datCurr, "yyyy-mm-dd hh:mm:ss") & "')"
    WriteInfo "�����������:" & strSQL
    gcn����.Execute strSQL
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & cur����ͳ���ۼ� + mcurͳ��֧�� & _
        "," & curͳ�ﱨ���ۼ� + mcurͳ��֧�� & "," & intסԺ�����ۼ� & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur��� & "," & cur�ʻ�֧���ۼ� + mcur����֧�� & "," & _
        cur����ͳ���ۼ� + mcurͳ��֧�� + mcur����Ա & "," & curͳ�ﱨ���ۼ� + mcurͳ��֧�� + mcur����Ա & _
        "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & mcur�ܶ� & "," & mcurȫ���� & "," & mcur�������� & _
        ",NULL," & mcurͳ��֧�� + mcur����Ա & ",NULL,NULL," & mcur����֧�� & ",NULL,NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    WriteInfo "�������ɹ�"
    
    �������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    WriteInfo "��������:" & Err.Description
    Exit Function
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'����:�������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'����:lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, strҽ���� As String, _
        cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, _
        curͳ�ﱨ���ۼ� As Currency, intסԺ�����ۼ� As Integer, cur��� As Currency, strSQL As String, _
        strPara As String, strReturn As String, strTransNO As String, str���� As String, lng����ID As Long, _
        curȫ���� As Currency, cur�������� As Currency, STRNAME As String

    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng����ID = rsTemp("����ID")
    WriteInfo vbCrLf & "׼�������˷�"
    
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ʻ���Ϣ", lng����ID)
    cur��� = Nvl(rsTemp!�ʻ����, 0): str���� = Nvl(rsTemp!����, "666666")
    
    gstrSQL = "Select * From ������Ϣ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", lng����ID)
    STRNAME = rsTemp!����
    
    'ȡԭ���ݽ�������
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    
    mcur�ܶ� = rsTemp!�������ý��
    mcur����֧�� = rsTemp!�����ʻ�֧��
    mcurͳ��֧�� = rsTemp!ͳ�ﱨ�����
    curȫ���� = rsTemp!ȫ�Ը����
    cur�������� = rsTemp!�����Ը����
    
    '����ţ���Ʊ�ţ�ҽ���ţ�'0'������
    strҽ���� = Getҽ����(lng����ID)
    
'    strSql = "Select * From hospital_clinic_payment Where hospital_no='" & Trim(gstrҽԺ����) & "' And " & _
'        "clinic_no='" & lng����ID & "' And invoice_no='" & lng����ID & "' and medical_card_no='" & strҽ���� & "'"
'
'    WriteInfo "ȡԭҽ����¼:" & strSql
'
'    Set rsTemp = gcn����.Execute(strSql)
'
'    If rsTemp.EOF Then
'        WriteInfo "ȡԭҽ����¼ʧ��"
'        Err.Raise 9000,gstrSysName, "ǰ�û����ݿ���δ�ҵ�ԭ�н��׼�¼����������"
'        Exit Function
'    End If
    
    strPara = lng����ID & "," & lng����ID & "," & strҽ���� & ",0," & dysEncrypt(str����)
    strTransNO = MakeTransNO()

    WriteInfo "��������:��ˮ��---" & strTransNO
    WriteInfo "���������� ����---" & strPara

    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','62','" & UserInfo.��� & "','" & strPara & "','9')"
    WriteInfo "д���ױ�:" & strSQL
    gcn����.Execute strSQL
    If frm�ȴ���Ӧ����.Result(strTransNO, strReturn) = False Then
        WriteInfo "������ֹ"
        Err.Raise 9000, gstrSysName, "������ֹ"
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        Err.Raise 9000, gstrSysName, "ҽ������ʧ��"
        Exit Function
    End If

    WriteInfo "���󷵻�:" & strReturn
    
    strSQL = "Insert Into hospital_clinic_payment (hospital_no,clinic_no,invoice_no,medical_card_no,name," & _
        "clinic_expense,full_self_expense,part_self_expense,mi_unpayment,social_payment,servant_payment," & _
        "account_payment,exectime) Values ('" & Trim(gstrҽԺ����) & "','" & lng����ID & "','" & lng����ID & "'," & _
        "'" & strҽ���� & "','" & STRNAME & "',-" & mcur�ܶ� & ",-" & curȫ���� & _
        ",-" & cur�������� & ",-" & CStr(mcur�ܶ� - mcur����֧��) & ",0,0" & _
        ",-" & mcur����֧�� & ",'" & Format(datCurr, "yyyy-mm-dd hh:mm:ss") & "')"
    WriteInfo "�����˷�����:" & strSQL
    gcn����.Execute strSQL
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - mcur����֧�� & "," & cur����ͳ���ۼ� - mcurͳ��֧�� & "," & _
        curͳ�ﱨ���ۼ� - mcurͳ��֧�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� - mcur����֧�� & "," & cur����ͳ���ۼ� - mcurͳ��֧�� & "," & _
            curͳ�ﱨ���ۼ� - mcurͳ��֧�� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            0 - mcur�ܶ� & ",0,0,NULL," & 0 - mcurͳ��֧�� & ",NULL,NULL," & _
            0 - mcur����֧�� & ",NULL,NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
    WriteInfo "��������:" & Err.Description
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'����:����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'����:lng����ID-����ID��lng��ҳID-��ҳID
'����:���׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, strPara As String, strReturn As String, _
        strTransNO As String, strInNote As String, str���ֱ��� As String, str�������� As String, _
        strסԺ���� As String, str��Ժ���� As String, strSQL As String

    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    WriteInfo vbCrLf & "����ҽ����Ժ"
    
    gstrSQL = "Select * From �����ʻ� Where ����ID = [1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽ�_���� = False
        Exit Function
    End If
    strҽ���� = Nvl(rsTemp!ҽ����, "")
    strPara = lng����ID & "_" & lng��ҳID & "," & strҽ���� & ",0"
    strTransNO = MakeTransNO()
    
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,D.���� as ���ұ���,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽ�_���� = False
        Exit Function
    End If
    
    strסԺ���� = ToVarchar(Nvl(rsTemp!סԺ����), 20)
    str��Ժ���� = ToVarchar(Nvl(rsTemp!��Ժ����), 3)

    '��ȡ��Ժ��ϣ����ֱ��룩
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, True, True) '��Ժ���
    If strInNote <> "" Then
        str�������� = Left(strInNote, InStr(strInNote, "|") - 1)
        str���ֱ��� = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    strSQL = "Insert Into hos_hospital_daybook (hospital_no,inhospital_no,medical_card_no," & _
        "last_examine_date,examine_date,disease_code,disease_name,inhospital_circs,inhospital_route," & _
        "inhospital_times,inhospital_type,sickarea_section_name,sickbed_no,outhospital_circs,checkout_date," & _
        "hospital_expense,part_self_payment,full_self_payment,start_payment,social_payment,social_unpayment," & _
        "supplement_payment,supplement_unpayment,servant_self_payment,servant_social_payment,cancel_sign) " & _
        "values ('" & Trim(gstrҽԺ����) & "','" & lng����ID & "_" & lng��ҳID & "','" & strҽ���� & "',NULL,'" & _
        Format(datCurr, "yyyy-mm-dd") & "','" & str���ֱ��� & "','" & str�������� & "','1','1'," & _
        mint����סԺ���� & ",'1','" & strסԺ���� & "','" & str��Ժ���� & "'," & _
        "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'9')"
    gcn����.Execute strSQL
    
    WriteInfo "��������:��ˮ��---" & strTransNO
    WriteInfo "���������� ����---" & strPara
    
    '����ҽ����Ժ����
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','09','" & UserInfo.��� & "','" & strPara & "','9')"
    gcn����.Execute strSQL
    If frm�ȴ���Ӧ����.Result(strTransNO, strReturn) = False Then
        WriteInfo "������ֹ"
        '���ʧ�ܣ���ɾ�����뵽סԺ��¼���еļ�¼
        gcn����.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstrҽԺ����) & "' And " & _
            "inhospital_no='" & lng����ID & "_" & lng��ҳID & "' And cancel_sign='9'"
        MsgBox "������ֹ", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "ҽ������ʧ��", vbInformation, gstrSysName
        '���ʧ�ܣ���ɾ�����뵽סԺ��¼���еļ�¼
        gcn����.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstrҽԺ����) & "' And " & _
            "inhospital_no='" & lng����ID & "_" & lng��ҳID & "' And cancel_sign='9'"
        Exit Function
    End If
    
    WriteInfo "���󷵻�:" & strReturn
    
    '��дIC��״̬
    If WriteCard(1) = False Then
        MsgBox "д��״̬ʱʧ�ܣ���������Ժ��������Ӱ��", vbInformation, "д��"
    End If
     
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'˳���','''0''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����֤��','''0''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
    
    WriteInfo "�����Ժ�Ǽ�"
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_���� = False
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, strPara As String, strReturn As String, _
        strTransNO As String, strҽ���� As String, strSQL As String

    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    WriteInfo vbCrLf & "������Ժ�Ǽ�"
    
    gstrSQL = "Select * From �����ʻ� Where ����ID = [1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽǳ���_���� = False
        Exit Function
    End If
    strҽ���� = Nvl(rsTemp!ҽ����, "")
    strPara = lng����ID & "_" & lng��ҳID & "," & strҽ���� & ",9"
    strTransNO = MakeTransNO()
    
    WriteInfo "��������:��ˮ��---" & strTransNO
    WriteInfo "���������� ����---" & strPara
    
    '����ҽ����Ժ����
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','09','" & UserInfo.��� & "','" & strPara & "','9')"
    gcn����.Execute strSQL
    If frm�ȴ���Ӧ����.Result(strTransNO, strReturn) = False Then
        WriteInfo "������ֹ"
        MsgBox "������ֹ", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "ҽ������ʧ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "���󷵻�:" & strReturn
    
    '����ʱɾ����Ժʱ�������Ժ��¼
    gcn����.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstrҽԺ����) & "' And " & _
        "inhospital_no='" & lng����ID & "_" & lng��ҳID & "' And cancel_sign='9'"
    
    '��дIC��״̬
    If WriteCard(0) = False Then
        MsgBox "д��״̬ʱʧ�ܣ���������Ժ�Ĳ�������Ӱ��", vbInformation, "д��"
    End If
     
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    WriteInfo "��ɳ����Ǽǲ���"
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽǳ���_���� = False
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    ��Ժ�Ǽ�_���� = True
End Function

Public Function ҽ����Ժ(lng����ID As Long, lng��ҳID As Long, ByRef str���� As String) As Boolean
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, strPara As String, strReturn As String, _
        strTransNO As String, strҽ���� As String, strInNote As String, str���ֱ��� As String, str�������� As String, _
        strסԺ���� As String, str��Ժ���� As String, strSQL As String

    '��Ժ�������
    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    WriteInfo vbCrLf & "���˳�Ժ�Ǽ�"
    
    gstrSQL = "Select * From סԺ���ü�¼ Where ����ID=[1] And ��ҳID=[2] And �����־=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    strPara = Nvl(rsTemp!����ID, "")
    If strPara = "" Then
        MsgBox "���������ò��ܰ����Ժ", vbInformation, gstrSysName
        Exit Function
    End If
'    gstrSQL = "Select * From ���ս����¼ Where ����=2 And ����=" & TYPE_���� & " And ��¼ID=" & strPara
'    Call OpenRecordset(rsTemp, gstrSysName)
'    mstr���ʷ��� = Nvl(rsTemp!��ע, "")
    
    gstrSQL = "Select * From �����ʻ� Where ����ID = [1] And ����= [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ҽ����Ժ = False
        Exit Function
    End If
    strҽ���� = Nvl(rsTemp!ҽ����, "")
    strPara = lng����ID & "_" & lng��ҳID & "," & strҽ���� & ",1"
    strTransNO = MakeTransNO()
    
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,D.���� as ���ұ���,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ҽ����Ժ = False
        Exit Function
    End If
    strסԺ���� = ToVarchar(Nvl(rsTemp!סԺ����), 20)
    str��Ժ���� = ToVarchar(Nvl(rsTemp!��Ժ����), 3)

    '��ȡ��Ժ��ϣ����ֱ��룩
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, True, True)  '��Ժ���
    If strInNote <> "" Then
        str�������� = Left(strInNote, InStr(strInNote, "|") - 1)
        str���ֱ��� = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    strSQL = "Insert Into hos_hospital_daybook (hospital_no,inhospital_no,medical_card_no," & _
        "last_examine_date,examine_date,disease_code,disease_name,inhospital_circs,inhospital_route," & _
        "inhospital_times,inhospital_type,sickarea_section_name,sickbed_no,outhospital_circs,checkout_date," & _
        "hospital_expense,part_self_payment,full_self_payment,start_payment,social_payment,social_unpayment," & _
        "supplement_payment,supplement_unpayment,servant_self_payment,servant_social_payment,cancel_sign) " & _
        "values ('" & Trim(gstrҽԺ����) & "','" & lng����ID & "_" & lng��ҳID & "','" & strҽ���� & "',NULL,'" & _
        Format(datCurr, "yyyy-mm-dd") & "','" & str���ֱ��� & "','" & str�������� & "','1','1'," & _
        mint����סԺ���� & ",'1','" & strסԺ���� & "','" & str��Ժ���� & "'," & _
        "'1','" & Format(datCurr, "yyyy-mm-dd") & "'," & mstr���ʷ��� & ",'0')"
    gcn����.Execute strSQL
    
    WriteInfo "��������:��ˮ��---" & strTransNO
    WriteInfo "���������� ����---" & strPara
    
    '����ҽ����Ժ����
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','09','" & UserInfo.��� & "','" & strPara & "','9')"
    gcn����.Execute strSQL
    If frm�ȴ���Ӧ����.Result(strTransNO, strReturn) = False Then
        WriteInfo "������ֹ"
        MsgBox "������ֹ", vbInformation, gstrSysName
        '���ʧ�ܣ���ɾ�����뵽סԺ��¼���еļ�¼
        gcn����.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstrҽԺ����) & "' And " & _
            "inhospital_no='" & lng����ID & "_" & lng��ҳID & "' And cancel_sign='0'"
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "ҽ������ʧ��", vbInformation, gstrSysName
        '���ʧ�ܣ���ɾ�����뵽סԺ��¼���еļ�¼
        gcn����.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstrҽԺ����) & "' And " & _
            "inhospital_no='" & lng����ID & "_" & lng��ҳID & "' And cancel_sign='0'"
        Exit Function
    End If
    
    WriteInfo "���󷵻�:" & strReturn
    str���� = strSQL
    WriteInfo "��ɳ�Ժ�Ǽ�"
    ҽ����Ժ = True
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    ҽ����Ժ = False
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long) As Boolean
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    ��Ժ�Ǽǳ���_���� = True
End Function

Public Function ����ҽ����Ժ_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, strPara As String, strReturn As String, _
        strTransNO As String, strҽ���� As String, strSQL As String

    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    WriteInfo vbCrLf & "������Ժ�Ǽ�"
    gstrSQL = "Select * From �����ʻ� Where ����ID = [1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ����ҽ����Ժ_���� = False
        Exit Function
    End If
    strҽ���� = Nvl(rsTemp!ҽ����, "")
    strPara = lng����ID & "_" & lng��ҳID & "," & strҽ���� & ",8"
    strTransNO = MakeTransNO()
    
    WriteInfo "��������:��ˮ��---" & strTransNO
    WriteInfo "���������� ����---" & strPara
    
    '����ҽ����Ժ����
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','09','" & UserInfo.��� & "','" & strPara & "','9')"
    gcn����.Execute strSQL
    If frm�ȴ���Ӧ����.Result(strTransNO, strReturn) = False Then
        WriteInfo "������ֹ"
        MsgBox "������ֹ", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "ҽ������ʧ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "���󷵻�:" & strReturn
    
    '������ɾ����Ժʱ����ĳ�Ժ��¼
    gcn����.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstrҽԺ����) & "' And " & _
        "inhospital_no='" & lng����ID & "_" & lng��ҳID & "' And cancel_sign='0'"
        
    WriteInfo "��ɳ�����Ժ����"
    ����ҽ����Ժ_���� = True
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    ����ҽ����Ժ_���� = False
End Function

Public Function ���ʴ���_����(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset, lng��ҳID As Long, rs��ϸ As New ADODB.Recordset, str��� As String, _
        str��Ŀ���� As String, str��Ŀ���� As String, str��λ As String, strͨ�������� As String, _
        cur������ As Currency, lngTemp As Long, strReturn As String, strPara As String, cur�������� As Currency, _
        strTransNO As String, strҽ���� As String, strҽ��״̬ As String, datCurr As Date, strSQL As String, _
        cur��������Sum As Currency, curȫ����Sum As Currency
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    WriteInfo vbCrLf & "��ʼ��ϸ����"
    If lng����ID <> 0 Then
        gstrSQL = "Select Max(��ҳID) From ������ҳ Where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
        lng��ҳID = rsTemp(0)
    End If
    
    If str���ݺ� <> "" Then
        gstrSQL = " Select A.* From סԺ���ü�¼ A,�����ʻ� B" & _
                  " Where A.�����־=2 And A.��¼״̬<>0 And nvl(A.���ӱ�־,0)<>9 and nvl(A.ʵ�ս��,0)<>0 " & _
                  " and A.��¼����=[1] and A.NO=[2]" & _
                  " And A.����ID=B.����ID And B.����=[3]" & _
                  " order by A.��ҳID,A.���"
        Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "", int����, str���ݺ�, TYPE_����)
    Else
        gstrSQL = " Select A.* From סԺ���ü�¼ A,�����ʻ� B" & _
                  " Where A.�����־=2 And A.��¼״̬<>0 And nvl(A.���ӱ�־,0)<>9 and nvl(A.ʵ�ս��,0)<>0 " & _
                  " and A.����id=[1] And A.��ҳid=[2]" & _
                  " And A.����ID=B.����ID And B.����=[3]" & _
                  " order by A.��ҳID,A.���"
        Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "", lng����ID, lng��ҳID, TYPE_����)
    End If
    If rs��ϸ.EOF Then
'        MsgBox "û����Ҫ���ķ�����ϸ", vbInformation, gstrSysName
        WriteInfo "û����Ҫ���ݵ���ϸ���˳�"
        ���ʴ���_���� = True
        Exit Function
    End If
    
    lng����ID = rs��ϸ!����ID: lng��ҳID = rs��ϸ!��ҳID
    
    strҽ���� = Getҽ����(lng����ID)
    strPara = strҽ����
    strTransNO = MakeTransNO()
    
    WriteInfo "��������:��ˮ��---" & strTransNO
    WriteInfo "���������� ����---" & strPara
    
    'ȡҽ��״̬
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','04','" & UserInfo.��� & "','" & strPara & "','9')"
    gcn����.Execute strSQL
    If frm�ȴ���Ӧ����.Result(strTransNO, strReturn) = False Then
        WriteInfo "������ֹ"
        MsgBox "������ֹ", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "ҽ������ʧ��", vbInformation, gstrSysName
        Exit Function
    End If
    strҽ��״̬ = Trim(Split(strReturn, ",")(1))
    WriteInfo "���󷵻�:ҽ��״̬(" & strҽ��״̬ & ")"
    
    gcnOracle.Execute "Delete From ҽ������״̬ Where ����ID=" & lng����ID & " And to_char(����,'yyyy-mm-dd')='" & Format(datCurr, "yyyy-mm-dd") & "'"
    gcnOracle.Execute "Insert into ҽ������״̬ (����ID,����,ҽ��״̬) values (" & lng����ID & ",to_date('" & Format(datCurr, "yyyy-mm-dd") & _
        "','yyyy-mm-dd')," & IIf(strҽ��״̬ = "", "NULL", strҽ��״̬) & ")"
    
    lngTemp = 0
    While Not rs��ϸ.EOF
        gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        If rsTemp!��� = "7" Then
            lngTemp = lngTemp + 1
        End If
        rs��ϸ.MoveNext
    Wend
    cur��������Sum = 0: curȫ����Sum = 0
    rs��ϸ.MoveFirst
'    gcnOracle.Execute "Delete From ��ҽ����ϸ Where ����id=" & lng����id & " And ��ҳID=" & lng��ҳID
    
    While Not rs��ϸ.EOF
        gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        str��Ŀ���� = rsTemp!����
        str��Ŀ���� = rsTemp!����
        str��λ = Nvl(rsTemp!���㵥λ, "")
        str��� = rsTemp!���
        
        strSQL = "Insert Into hos_advice_carryout (hospital_no,advice_serial_no,inhospital_no," & _
            "item_drug_code,conv_price,quantity,norm_unit,all_expense,self_payment,carryout_date," & _
            "item_drug_name,general_code,cease_reason) values ('" & Trim(gstrҽԺ����) & "','" & IIf(rs��ϸ!ʵ�ս�� < 0, "-", "") & _
            rs��ϸ!NO & "_" & rs��ϸ!��� & "','" & rs��ϸ!����ID & "_" & rs��ϸ!��ҳID & "',"
        
        gstrSQL = "Select * From ����֧����Ŀ Where �շ�ϸĿID=[1] And ����=[2] And �Ƿ�ҽ��=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����֧����Ŀ", CLng(rs��ϸ!�շ�ϸĿID), TYPE_����)
        cur������ = 0: cur�������� = 0
        If rsTemp.EOF Then
            If str��� = "5" Or str��� = "6" Or str��� = "7" Then
'                str��Ŀ���� = "yz" & str��Ŀ����
                strͨ�������� = "8888"
            Else
'                str��Ŀ���� = "zz" & str��Ŀ����
                strͨ�������� = "9000"
            End If
            cur������ = rs��ϸ!ʵ�ս��
        Else
            str��Ŀ���� = rsTemp!��Ŀ����
            strͨ�������� = rsTemp!��Ŀ����
            If str��� = "5" Or str��� = "6" Or str��� = "7" Then
'                str��Ŀ���� = "@@" & rsTemp!��Ŀ����
                gstrSQL = "Select * From mi_drug_trade_list Where trade_code='" & rsTemp!��Ŀ���� & "'"
                Set rsTemp = gcn����.Execute(gstrSQL)
                If rsTemp.EOF Then
                    cur������ = rs��ϸ!ʵ�ս��
                ElseIf Trim(rsTemp!mi_class) = "1" Then
                    cur������ = 0
                ElseIf Trim(rsTemp!mi_class) = "2" Then
                    cur�������� = rs��ϸ!ʵ�ս�� * 0.2
                ElseIf Trim(rsTemp!mi_class) = "4" Then
                    If lngTemp < 2 Then
                        cur������ = rs��ϸ!ʵ�ս��
                    Else
                        cur������ = 0
                    End If
                Else
                    cur������ = rs��ϸ!ʵ�ս��
                End If
            ElseIf str��� = "J" Then
                If Val(rs��ϸ!ʵ�ս��) > 0 Then
                    cur�������� = IIf(rs��ϸ!ʵ�ս�� > 15, rs��ϸ!ʵ�ս�� - 15, 0)
                Else
                    cur�������� = IIf(Abs(rs��ϸ!ʵ�ս��) > 15, rs��ϸ!ʵ�ս�� + 15, 0)
                End If
            Else
'                str��Ŀ���� = "$$" & rsTemp!��Ŀ����
                gstrSQL = "Select * From mi_dt_item Where item_code='" & rsTemp!��Ŀ���� & "'"
                Set rsTemp = gcn����.Execute(gstrSQL)
                If rsTemp.EOF Then
                    cur������ = rs��ϸ!ʵ�ս��
                ElseIf Trim(rsTemp!mi_class) = "1" Then
                    cur������ = 0
                ElseIf Trim(rsTemp!mi_class) <> "2" Then
                    cur������ = rs��ϸ!ʵ�ս��
                Else
                    cur�������� = rs��ϸ!ʵ�ս�� * rsTemp!self_rate / 100
                End If
            End If
        End If
        
        gstrSQL = "Select * From ҽ������״̬ Where to_char(����,'yyyy-mm-dd')='" & Format(rs��ϸ!����ʱ��, "yyyy-mm-dd") & "' And ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ҽ��״̬", lng����ID)
        If rsTemp.EOF Then
            strҽ��״̬ = "0"
        Else
            strҽ��״̬ = Nvl(rsTemp!ҽ��״̬, "0")
        End If
        
        If strҽ��״̬ <> "0" Then
            strҽ��״̬ = "1"
            cur������ = rs��ϸ!ʵ�ս��
            cur�������� = 0
        End If
        
        strSQL = strSQL & "'" & str��Ŀ���� & "'," & Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), "0.0000") & _
            "," & rs��ϸ!���� * rs��ϸ!���� & ",'" & str��λ & "'," & rs��ϸ!ʵ�ս�� & "," & _
            Format(cur������ + cur��������, "0.0000") & ",'" & Format(datCurr, "yyyy-mm-dd") & "','" & _
            ToVarchar(str��Ŀ����, 60) & "','" & strͨ�������� & "','" & strҽ��״̬ & "')"
            
        If Nvl(rs��ϸ!�Ƿ��ϴ�, 0) = 0 Then
            WriteInfo "д��ǰ�û���ϸ:" & strSQL
            gcn����.Execute strSQL
        End If
        
        cur��������Sum = cur��������Sum + cur��������
        curȫ����Sum = curȫ����Sum + cur������
        
        gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rs��ϸ("ID") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        
        strSQL = "Insert Into ��ҽ����ϸ Values ('" & ToVarchar(str��Ŀ����, 50) & "','" & str��λ & "'," & _
            Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), "0.0000") & "," & rs��ϸ!���� * rs��ϸ!���� & _
            "," & rs��ϸ!ʵ�ս�� & "," & Format(cur������ + cur��������, "0.0000") & "," & lng����ID & "," & lng��ҳID & _
            ",to_date('" & Format(rs��ϸ!����ʱ��, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')," & rs��ϸ("ID") & "," & strҽ��״̬ & ")"
        If Nvl(rs��ϸ!�Ƿ��ϴ�, 0) = 0 Then
            WriteInfo "д����ҽ����ϸ:" & strSQL
            gcnOracle.Execute strSQL
        End If
        rs��ϸ.MoveNext
    Wend
    
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
'    cur������ = CCur(rsTemp!����֤��): cur�������� = CCur(rsTemp!˳���)
    cur������ = 0: cur�������� = 0
    cur��������Sum = cur��������Sum + cur��������
    curȫ����Sum = curȫ����Sum + cur������
    
    WriteInfo "���没���Ѵ������:�����������(" & cur��������Sum & ")  ȫ�������(" & curȫ����Sum & ")"
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'˳���','''" & cur��������Sum & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����֤��','''" & curȫ����Sum & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
    WriteInfo "�����ϸ����"
    ���ʴ���_���� = True
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rs��ϸ As Recordset, lng����ID As Long, strҽ���� As String) As String
    Dim rsTemp As New ADODB.Recordset, lng��ҳID As Long, curȫ���� As Currency, cur�������� As Currency, _
        cur�ܶ� As Currency, int�����־ As Integer, strReturn As String, strPara As String, _
        strTransNO As String, lng��ҩ As Long, datCurr As Date, cur��ζ��ҩ As Currency, _
        rs���� As New ADODB.Recordset, strTemp As String, str��� As String, strSQL As String, _
        rs������ϸ As New ADODB.Recordset, i As Long
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    If ���ʴ���_����("", 0, "", lng����ID) = False Then
        Exit Function
    End If
    
    WriteInfo vbCrLf & "��ʼסԺԤ����"
    
    gstrSQL = "Select max(��ҳID) From סԺ���ü�¼ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng��ҳID = rsTemp(0)
    
    gstrSQL = "Select NO From סԺ���ü�¼ Where �����־=2 And ����id=[1] And ��ҳID=[2] Group By NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    
    mcurȫ���� = 0: mcur�������� = 0: mcur�ܶ� = 0
    
    While Not rsTemp.EOF
        gstrSQL = "Select A.���,B.ʵ�ս��,A.ID,B.NO From �շ�ϸĿ A,סԺ���ü�¼ B Where A.ID=B.�շ�ϸĿID And " & _
            "B.�����־=2 And B.NO='" & rsTemp!NO & "' And B.����ID=[1] And B.��ҳID=[2] And nvl(ʵ�ս��,0)<>0"
        Set rs������ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
        
        lng��ҩ = 0
        '���㵥ζ��ҩ���
        While Not rs������ϸ.EOF
            If rs������ϸ!��� = "7" Then lng��ҩ = lng��ҩ + 1
            rs������ϸ.MoveNext
        Wend
        rs������ϸ.MoveFirst
        
        While Not rs������ϸ.EOF
            gstrSQL = "Select * From ����֧����Ŀ Where �շ�ϸĿID=[1]"
            Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs������ϸ!ID))
            If rs����.EOF Then
                mcurȫ���� = mcurȫ���� + rs������ϸ!ʵ�ս��
            Else
                str��� = rs������ϸ!���
                If str��� = "5" Or str��� = "6" Or str��� = "7" Then
                    gstrSQL = "Select * From mi_drug_trade_list Where trade_code='" & rs����!��Ŀ���� & "'"
                    Set rs���� = gcn����.Execute(gstrSQL)
                    If rs����.EOF Then
                        mcurȫ���� = mcurȫ���� + rs������ϸ!ʵ�ս��
                    ElseIf rs����!mi_class = "1" Then
                        
                    ElseIf rs����!mi_class = "2" Then
                        mcur�������� = mcur�������� + rs������ϸ!ʵ�ս�� * 0.2
                    ElseIf rs����!mi_class = "4" Then
                        If lng��ҩ < 2 Then
                            mcurȫ���� = mcurȫ���� + rs������ϸ!ʵ�ս��
                        End If
                    Else
                        mcurȫ���� = mcurȫ���� + rs������ϸ!ʵ�ս��
                    End If
                ElseIf str��� = "J" Then
                    mcur�������� = mcur�������� + IIf(rs������ϸ!ʵ�ս�� > 15, rs������ϸ!ʵ�ս�� - 15, 0)
                Else
                    gstrSQL = "Select * From mi_dt_item Where item_code='" & rs����!��Ŀ���� & "'"
                    Set rs���� = gcn����.Execute(gstrSQL)
                    If rs����.EOF Then
                        mcurȫ���� = mcurȫ���� + rs������ϸ!ʵ�ս��
                    Else
                        mcur�������� = mcur�������� + rs������ϸ!ʵ�ս�� * rs����!self_rate
                    End If
                End If
            End If
            mcur�ܶ� = mcur�ܶ� + rs������ϸ!ʵ�ս��
            rs������ϸ.MoveNext
        Wend
        
        rsTemp.MoveNext
    Wend
    gstrSQL = "Select * From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    mcurȫ���� = CCur(Nvl(rsTemp!����֤��, "0"))
    mcur�������� = CCur(Nvl(rsTemp!˳���, "0"))
    
    WriteInfo "�����������:�ܶ�(" & mcur�ܶ� & ")  ȫ����(" & mcurȫ���� & ")  ��������(" & mcur�������� & ")"
    
    gstrSQL = "Select ��Ժ���� From ������ҳ Where ����id=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If Format(rsTemp(0), "yyyy") <> Format(datCurr, "yyyy") Then
        int�����־ = 1
    Else
        int�����־ = 0
    End If
    
    If mcur�ܶ� = 0 Then
        mcur�������� = 0: mcurȫ���� = 0
    End If
    
    strTransNO = MakeTransNO()
    'ҽԺ�ţ�ҽ���ţ�סԺ�ܷ��ã����������ַ��ã���ȫ�����ַ��ã���Ժ�����
    strPara = Trim(gstrҽԺ����) & "," & strҽ���� & "," & mcur�ܶ� & "," & mcur�������� & "," & mcurȫ���� & "," & int�����־
    
    WriteInfo "��������:��ˮ��---" & strTransNO
    WriteInfo "���������� ����---" & strPara
    
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','22','" & UserInfo.��� & "','" & strPara & "','9')"
    gcn����.Execute strSQL
    If frm�ȴ���Ӧ����.Result(strTransNO, strReturn) = False Then
        WriteInfo "������ֹ"
        MsgBox "������ֹ", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "ҽ������ʧ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "���󷵻�:" & strReturn
    
    '0-�ɹ���־��1-���������ַ��ã�2-��ȫ�����ַ��ã�3-����֧����4-ͳ��֧����5-ͳ�ﲻ֧����6-��֧��
    '7-�󲡲�֧����8-����Ա������ã�9-����Ա����֧��
    mcur����֧�� = CCur(Split(strReturn, ",")(6))
    mcurͳ��֧�� = CCur(Split(strReturn, ",")(4))
    mcur����Ա = CCur(Split(strReturn, ",")(9))
'    mstr���ʷ��� = Right(Space(15) & mcur�ܶ�, 15)
    'mstr���ʷ��� = mcur�ܶ� & Mid(strReturn, InStr(strReturn, ","))
    mstr���ʷ��� = mcur�ܶ�
    strSQL = "Delete ��ҽ������ Where ����id=" & lng����ID & " And ��ҳid=" & lng��ҳID
    gcnOracle.Execute strSQL
    
    strSQL = "Insert Into ��ҽ������ Values (" & lng����ID & "," & lng��ҳID & "," & Format(mcur�ܶ�, "0.0000")
    For i = 1 To 9
'        mstr���ʷ��� = mstr���ʷ��� & Right(Space(15) & Split(strReturn, ",")(i), 15)
        strSQL = strSQL & "," & Format(Split(strReturn, ",")(i), "0.0000")
        mstr���ʷ��� = mstr���ʷ��� & "," & Format(Split(strReturn, ",")(i), "0.0000")
    Next
    strSQL = strSQL & ")"
    WriteInfo "д����ҽ����������:" & strSQL
    gcnOracle.Execute strSQL
    
    If mcur����֧�� <> 0 Then סԺ�������_���� = "��֧��;" & mcur����֧�� & ";0"
    סԺ�������_���� = סԺ�������_���� & IIf(סԺ�������_���� <> "", "|", "") & "ͳ�����;" & mcurͳ��֧�� & ";0"
    If mcur����Ա <> 0 Then סԺ�������_���� = סԺ�������_���� & IIf(סԺ�������_���� <> "", "|", "") & "����Ա����;" & mcur����Ա & ";0"
    
    WriteInfo "���Ԥ����:" & סԺ�������_����
    Exit Function
errHandle:
    WriteInfo "��������:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, lng����ID As Long) As Boolean
    Dim datCurr As Date, intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, _
        cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, cur��� As String, rsTemp As New ADODB.Recordset, _
        str��Ժ���� As String
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    cur��� = �������_����(lng����ID)
    
    gstrSQL = "Select max(��ҳid) from ������ҳ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If ҽ����Ժ(lng����ID, rsTemp(0), str��Ժ����) = False Then Exit Function
    str��Ժ���� = Replace(str��Ժ����, "'", "''")
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� + mcurͳ��֧�� + mcur����Ա + mcur����֧�� & _
        "," & curͳ�ﱨ���ۼ� + mcurͳ��֧�� + mcur����Ա + mcur����֧�� & "," & intסԺ�����ۼ� & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur��� & "," & cur�ʻ�֧���ۼ� & "," & _
        cur����ͳ���ۼ� + mcurͳ��֧�� + mcur����Ա + mcur����֧�� & "," & curͳ�ﱨ���ۼ� + mcurͳ��֧�� + mcur����Ա + mcur����֧�� & _
        "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & mcur�ܶ� & "," & mcurȫ���� & "," & mcur�������� & _
        ",NULL," & mcurͳ��֧�� + mcur����Ա + mcur����֧�� & ",NULL,NULL,NULL,NULL,NULL,NULL,'" & str��Ժ���� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    '��дIC��״̬
    If WriteCard(0) = False Then
        Err.Raise 9000, gstrSysName, "д��״̬ʱʧ�ܣ������˳�Ժ�Ĳ�������Ӱ��"
    End If
     
    סԺ����_���� = True
    
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    WriteInfo "��������:" & Err.Description
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
'����:�������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'����:lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, StrInput As String, sngArrInfo(20) As Single
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String, lng����ID As Long
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, strTemp As String
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, rstTemp As String
    Dim curƱ���ܽ�� As Currency, lng��ҳID As Long
    Dim datCurr As Date, cur�����ʻ� As Currency

        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        סԺ�������_���� = False
        Exit Function
    End If
    
    cur�����ʻ� = Nvl(rsTemp!�����ʻ�֧��, 0)
    strTemp = Nvl(rsTemp!��ע, "")
    
'    gstrSQL = "Select * From �����ʻ� Where ����id=" & rsTemp!����ID & " And ����=" & TYPE_����
'    Call OpenRecordset(rsTemp, gstrSysName)
'    If Nvl(rsTemp!��ǰ״̬, 0) = 0 Then
'        MsgBox "סԺ�������ǰ���ȳ������˳�Ժ"
'        Exit Function
'    End If
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)
    lng����ID = rsTemp!����ID
    
    gstrSQL = "Select max(��ҳid) from ������ҳ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng��ҳID = Nvl(rsTemp(0), 1)
    If ����ҽ����Ժ_����(lng����ID, lng��ҳID) = False Then
        Exit Function
    End If
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� - 1 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� - 1 & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        cur�����ʻ� * -1 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '��дIC��״̬
    If WriteCard(1) = False Then
        MsgBox "д��״̬ʱʧ�ܣ���������Ժ�Ĳ�������Ӱ��", vbInformation, "д��"
    End If

    סԺ�������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
    WriteInfo "��������:" & Err.Description
    סԺ�������_���� = False
End Function

