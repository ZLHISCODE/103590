Attribute VB_Name = "mdlPubServicePati"
Option Explicit

'*********************************************************************************************************************************************
'����:�����漰���÷��õ���ط���
'�ӿ�˵��:
'    1.Zl_���˽����쳣��¼_Modify-�����쳣��������,�޸ļ�ɾ��
'    2.zl_PatiSvr_NewPatiArchives-�½����˵���
'    3.zl_Patisvr_GetNextNo-��ȡ����ż�����ID��Ϣ
'����:���˺�
'����:2019*10*31 14:47:18
'*********************************************************************************************************************************************
Private mlngErrNum As Long, mstrSource As String, mstrErrMsg As String
Private Function GetServiceCall(ByRef objServiceCall_Out As Object, Optional blnShowErrMsg As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����������
    '����:objServiceCall_Out-���ع����������
    '����:��ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
     
    If Not gobjServiceCall Is Nothing Then Set objServiceCall_Out = gobjServiceCall: GetServiceCall = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
    If Err <> 0 Then
        mstrErrMsg = "������zlServiceCall����ʧ������ϵͳ����Ա��ϵ���ָ��ò�����"
        If blnShowErrMsg Then
            MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
            Err = 0: On Error GoTo 0
        Else
            Err.Raise Err.Number, Err.Source, mstrErrMsg: Exit Function
        End If
        Exit Function
    End If
    
    On Error GoTo errHandle
    If gobjServiceCall.InitService(gcnOracle, gstrDBUser, glngSys, glngModul) = False Then
        
        Set gobjServiceCall = Nothing: Exit Function
    End If
    Set objServiceCall_Out = gobjServiceCall
    GetServiceCall = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zl_Patisvr_GetNextNo(ByVal int��� As Integer, ByRef strNo_Out As String, Optional ByVal lng����ID As Long, _
    Optional blnShowErrMsg As Boolean = True, Optional ByRef strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���õ��ݺ�
    '���:
    '����:strErrMsg_Out-������Ϣ
    '     strNo_Out-���ص�һ���ŵ��ݺ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
  
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'input
    '    item_num    N   1   ��Ŀ���
    '    dept_id N       ����ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("item_num", int���, Json_num)
    strJson = strJson & "," & GetJsonNodeString("recv_id", lng����ID, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Patisvr_GetNextNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    output
    '        code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '        message C   1   "Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '        next_no C       ��һ������

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡ���ݺţ����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strNo_Out = objServiceCall.GetJsonNodeValue("output.next_no")
    zl_Patisvr_GetNextNo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function Zl_���˽����쳣��¼_Modify(ByVal int����״̬ As Integer, ByVal cllSaveData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�쳣������������
    '���:int����״̬-����״̬:0-������¼,1-����״̬�����½���˵����2-ɾ���쳣����
    '    cllSaveData-��ʽΪArray(����������,������ֵ)
    '          �������Ŀ���ư���: �쳣ID,��������,���ϱ�־,ҵ��id,�Ƿ�����,����id,��ҳid,����,�Ա�,����,�����,סԺ��,Ԥ������,Ԥ�����,ҽ�ƿ�����,����,�������id,�����������,��������,ͬ��״̬,������Ϣ)
    '          ���н�����ϢΪJson������ʽ����
    '           {"card_no":"00002","cardtype_id":23,"swapno":"J2223432","swapmoney":324,"otherswap_list":[{"swap_name":"POSM","swap_note":"A001"},{}]}
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '     str��ҳid-��ҳid=""ʱ����ʾ������ҳid����;0ʱ��ʾֻ����ҳidΪ���ҽ��,>0��ʾ��ѯָ����ҳ��ҽ��
    '����:strErrmsg_Out-���ش�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objExpense As Object
    If zlGetPubExpenseObject(objExpense) = False Then Exit Function
    Zl_���˽����쳣��¼_Modify = objExpense.Zl���˽����쳣��¼_Modify(int����״̬, cllSaveData)
End Function


Public Function Zl_ҽ�ƿ��䶯_Insert_Check(ByVal cllCheckData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ�ƿ��䶯ǰ������ݵĺϷ���
    '���:cllCheckData-��ʽΪArray(����������,������ֵ)
    '                   �������Ŀ���ư���:����״̬,����ID,�����ID,����,�¿���,�쳣״̬
    '                    ����״̬:1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ),7-��ֹʱ�����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int����״̬ As Integer, lng����id As Long, lng�����ID As Long, str���� As String, str�¿��� As String, int�쳣״̬ As Integer
    Dim varRetrun As Variant, strErrMsg As String
    Dim i As Long
    
    On Error GoTo errHandle
    If cllCheckData Is Nothing Then
        strErrMsg = "δ������Ҫ���ı�Ҫ����������!"
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    For i = 1 To cllCheckData.Count
        Select Case UCase(cllCheckData(i)(0))
        Case "����״̬"
            int����״̬ = Val(cllCheckData(i)(1))
        Case "����ID"
            lng����id = Val(cllCheckData(i)(1))
        Case "�����ID"
            lng�����ID = Val(cllCheckData(i)(1))
        Case "����"
            str���� = Trim(cllCheckData(i)(1))
        Case "�¿���"
            str�¿��� = Trim(cllCheckData(i)(1))
        Case "�쳣״̬"
            int�쳣״̬ = Val(cllCheckData(i)(1))
        End Select
    Next
    
    '  Zl_ҽ�ƿ��䶯_Insert_Check
    '����״̬_In  Integer,
    '�����id_In  ҽ�ƿ����.Id%Type,
    '����_In      ����ҽ�ƿ��䶯.����%Type,
    '�¿���_In    ����ҽ�ƿ��䶯.����%Type,
    '����id_In    ����ҽ�ƿ��䶯.����id%Type,
    'Ӧǩ��_Out   Out Integer,
    'Ӧ����Ϣ_Out Out Varchar2
    '
    varRetrun = zlDatabase.CallProcedure("Zl_ҽ�ƿ��䶯_Insert_Check", "ҽ�ƿ��䶯���", int����״̬, lng�����ID, str����, str�¿���, lng����id, int�쳣״̬, Empty, Empty)
    
    If varRetrun(0) <> 1 Then
        strErrMsg = varRetrun(1)
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Zl_ҽ�ƿ��䶯_Insert_Check = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function




Public Function zl_PatiSvr_NewPatiArchives(ByVal cllUpdBasePati As Collection, ByVal cllUpdContacts As Collection, _
    ByVal cllUpdCommunity As Collection, ByVal cllUpdVisit As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ�ƿ��䶯ǰ������ݵĺϷ���
    '���:cllUpdBasePati-�޸Ĳ��˻�����Ϣ:����ID,����,�Ա�,����,��������(yyyy-mm-dd hh24:mi:ss),���֤��,��������,�����,���￨��,����֤��,�ѱ�,ҽ�Ƹ��ʽ����,����,����,����,����״��,ѧ��,ְҵ,���,������λ,
    '            ��λ�ʱ�,��λ�绰,��λ������,��λ�ʺ�,��ͬ��λID,��ͥ��ַ,��ͥ�绰,��ͥ��ַ�ʱ�,����,�����ص�,���ڵ�ַ,���ڵ�ַ�ʱ�,�໤��,�ֻ���,
    '            ҽ����,Ic����,�Ǽ�ʱ��,����Ա����,���֤ǩԼ,ǩԼ����,����,����֤����
    '    cllUpdContacts-�޸Ĳ�����ϵ����Ϣ:(��ϵ������,��ϵ�����֤��,��ϵ�˵绰,��ϵ�˹�ϵ,��ϵ�˵�ַ)
    '    cllUpdCommunity-�޸�������Ϣ:�������,��������,������������
    '    cllUpdVisit-���¾�����Ϣ:����״̬,��������,����ʱ��
    '   ���ϼ��ϸ�ʽ:Array(����������,������ֵ)
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String, strJsonTemp As String
    Dim clldata As Variant, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    

    On Error GoTo errHandle
    If GetServiceCall(objServiceCall, True) = False Then Exit Function
    If Not cllUpdBasePati Is Nothing Then
        strJson = ""
        For i = 1 To cllUpdBasePati.Count
            varTemp = cllUpdBasePati(i)
            Select Case varTemp(0)
            Case "����ID"
                strJson = strJson & "," & GetJsonNodeString("pati_id", Val(varTemp(1)), Json_num, True)
            Case "����"
                strJson = strJson & "," & GetJsonNodeString("pati_name", Trim(varTemp(1)), Json_Text)
            Case "�Ա�"
                strJson = strJson & "," & GetJsonNodeString("pati_sex", Trim(varTemp(1)), Json_Text)
            Case "����"
                strJson = strJson & "," & GetJsonNodeString("pati_age", Trim(varTemp(1)), Json_Text)
            Case "��������" ':yyyy-mm-dd hh24:mi:ss
                strJson = strJson & "," & GetJsonNodeString("pati_birthdate", Trim(varTemp(1)), Json_Text)
            Case "���֤��"
                strJson = strJson & "," & GetJsonNodeString("pati_idcard", Trim(varTemp(1)), Json_Text)
            Case "��������"
                strJson = strJson & "," & GetJsonNodeString("pati_type", Trim(varTemp(1)), Json_Text)
            Case "�����"
                strJson = strJson & "," & GetJsonNodeString("outpatient_num", Val(varTemp(1)), Json_num, True)
            Case "���￨��"
                strJson = strJson & "," & GetJsonNodeString("vcard_no", Trim(varTemp(1)), Json_Text)
            Case "����֤��"
                strJson = strJson & "," & GetJsonNodeString("vcard_pwd", Trim(varTemp(1)), Json_Text)
            Case "�ѱ�"
                strJson = strJson & "," & GetJsonNodeString("fee_category", Trim(varTemp(1)), Json_Text)
            Case "ҽ�Ƹ��ʽ����"
                strJson = strJson & "," & GetJsonNodeString("mdlpay_mode_name", Trim(varTemp(1)), Json_Text)
            Case "����"
                strJson = strJson & "," & GetJsonNodeString("native_place", Trim(varTemp(1)), Json_Text)
            Case "����"
                strJson = strJson & "," & GetJsonNodeString("country_name", Trim(varTemp(1)), Json_Text)
            Case "����"
                strJson = strJson & "," & GetJsonNodeString("nation_name", Trim(varTemp(1)), Json_Text)
            Case "����״��"
                strJson = strJson & "," & GetJsonNodeString("mari_status", Trim(varTemp(1)), Json_Text)
            Case "ѧ��"
                strJson = strJson & "," & GetJsonNodeString("edu_name", Trim(varTemp(1)), Json_Text)
            Case "ְҵ"
                strJson = strJson & "," & GetJsonNodeString("ocpt_name", Trim(varTemp(1)), Json_Text)
            Case "���"
                strJson = strJson & "," & GetJsonNodeString("pati_identity", Trim(varTemp(1)), Json_Text)
            Case "������λ"
                strJson = strJson & "," & GetJsonNodeString("emp_name", Trim(varTemp(1)), Json_Text)
            Case "��λ�ʱ�"
                strJson = strJson & "," & GetJsonNodeString("emp_postcode", Trim(varTemp(1)), Json_Text)
            Case "��λ�绰"
                strJson = strJson & "," & GetJsonNodeString("emp_phno", Trim(varTemp(1)), Json_Text)
            Case "��λ������"
                strJson = strJson & "," & GetJsonNodeString("emp_bank_name", Trim(varTemp(1)), Json_Text)
            Case "��λ�ʺ�"
                strJson = strJson & "," & GetJsonNodeString("emp_bank_accnum", Trim(varTemp(1)), Json_Text)
            Case "��ͬ��λID"
                strJson = strJson & "," & GetJsonNodeString("ctt_unit_id", Val(varTemp(1)), Json_num, True)
            Case "��ͥ��ַ"
                strJson = strJson & "," & GetJsonNodeString("pat_home_addr", Trim(varTemp(1)), Json_Text)
            Case "��ͥ�绰"
                strJson = strJson & "," & GetJsonNodeString("pat_home_phno", Trim(varTemp(1)), Json_Text)
            Case "��ͥ��ַ�ʱ�"
                strJson = strJson & "," & GetJsonNodeString("pat_home_postcode", Trim(varTemp(1)), Json_Text)
            Case "����"
                strJson = strJson & "," & GetJsonNodeString("region", Trim(varTemp(1)), Json_Text)
            Case "�����ص�"
                strJson = strJson & "," & GetJsonNodeString("pat_baddr", Trim(varTemp(1)), Json_Text)
            Case "���ڵ�ַ"
                strJson = strJson & "," & GetJsonNodeString("pat_hous_addr", Trim(varTemp(1)), Json_Text)
            Case "���ڵ�ַ�ʱ�"
                strJson = strJson & "," & GetJsonNodeString("pat_hous_postcode", Trim(varTemp(1)), Json_Text)
            Case "�໤��"
                strJson = strJson & "," & GetJsonNodeString("pat_grdn_name", Trim(varTemp(1)), Json_Text)
            Case "�ֻ���"
                strJson = strJson & "," & GetJsonNodeString("phone_number", Trim(varTemp(1)), Json_Text)
            Case "ҽ����"
                strJson = strJson & "," & GetJsonNodeString("insurance_num", Trim(varTemp(1)), Json_Text)
            Case "IC����"
                strJson = strJson & "," & GetJsonNodeString("iccard_no", Trim(varTemp(1)), Json_Text)
            Case "�Ǽ�ʱ��" ':yyyy-mm-dd hh24:mi:ss
                strJson = strJson & "," & GetJsonNodeString("create_time", Trim(varTemp(1)), Json_Text)
            Case "����Ա���� "
                strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
            Case "���֤ǩԼ"
                strJson = strJson & "," & GetJsonNodeString("idcard_sign", Val(varTemp(1)), Json_num, True)
            Case "ǩԼ����"
                strJson = strJson & "," & GetJsonNodeString("idcard_sign_pwd", Trim(varTemp(1)), Json_Text)
            Case "����"
                strJson = strJson & "," & GetJsonNodeString("insurance_type", Trim(varTemp(1)), Json_Text)
            Case "����֤��"
                strJson = strJson & "," & GetJsonNodeString("cert_no_other", Trim(varTemp(1)), Json_Text)
            Case ""
            End Select
        Next
    End If
    
    If Not cllUpdContacts Is Nothing Then
        '������ϵ��
        strJsonTemp = ""
        For i = 1 To cllUpdContacts.Count
            varTemp = cllUpdContacts(i)
            Select Case UCase(varTemp(0))
            Case "��ϵ������"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("name", varTemp(1), Json_Text)
            Case "��ϵ�����֤��"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("idcard", varTemp(1), Json_Text)
            Case "��ϵ�˵绰"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("phone", varTemp(1), Json_Text)
            Case "��ϵ�˹�ϵ"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("relation", varTemp(1), Json_Text)
            Case "��ϵ�˵�ַ"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("address", varTemp(1), Json_Text)
            End Select
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonTemp = "," & GetNodeString("contacts") & ":{" & strJsonTemp & "}"
            strJson = strJson & strJsonTemp
        End If
    End If
    
    If Not cllUpdCommunity Is Nothing Then
        '��������
        strJsonTemp = ""
        For i = 1 To cllUpdCommunity.Count
            varTemp = cllUpdCommunity(i)
            Select Case UCase(varTemp(0))
            Case "�������"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("num", Val(varTemp(1)), Json_num, True)
            Case "��������"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("code", varTemp(1), Json_Text)
            Case "������������"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("oper_type", Val(varTemp(1)), Json_num, True)
            End Select
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonTemp = "," & GetNodeString("community_info") & ":{" & strJsonTemp & "}"
            strJson = strJson & strJsonTemp
        End If
    End If
    
    If Not cllUpdVisit Is Nothing Then
        '��������
        strJsonTemp = ""
        For i = 1 To cllUpdVisit.Count
            varTemp = cllUpdVisit(i)
            Select Case UCase(varTemp(0))
            Case "����״̬"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("statu", Val(varTemp(1)), Json_num, True)
            Case "��������"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("room", varTemp(1), Json_Text)
            Case "����ʱ��"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("time", varTemp(1), Json_Text)
            End Select
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonTemp = "," & GetNodeString("visit_info") & ":{" & strJsonTemp & "}"
            strJson = strJson & strJsonTemp
        End If
    End If
    
    If strJson = "" Then
        strErrMsg = "δ������Ҫ���ı�Ҫ����������!"
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_PatiSvr_NewPatiArchives
    'input
    '   pati_id N   1   ����id
    '   pati_name   C   1   ����
    '   pati_sex    C   1   �Ա�
    '   pati_age    C   1   ����
    '   pati_birthdate  C   1   ��������:yyyy-mm-dd hh24:mi:ss
    '   pati_idcard C   1   ���֤��
    '   pati_type   C   1   ��������
    '   outpatient_num  N   1   �����
    '   vcard_no    C   1   ���￨��
    '   vcard_pwd   C   1   ����֤��
    '   fee_category    C   1   �ѱ�
    '   mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����
    '   native_place    C   1   ����
    '   country_name    C   1   ����
    '   nation_name C   1   ����
    '   mari_status C   1   ����״��
    '   edu_name    C   1   ѧ��
    '   ocpt_name   C   1   ְҵ
    '   pati_identity   C   1   ���
    '   emp_name    C   1   ������λ
    '   emp_postcode    C   1   ��λ�ʱ�
    '   emp_phno    C   1   ��λ�绰
    '   emp_bank_name   C   1   ��λ������
    '   emp_bank_accnum C   1   ��λ�ʺ�
    '   ctt_unit_id N   1   ��ͬ��λid
    '   pat_home_addr   C   1   ��ͥ��ַ
    '   pat_home_phno   C   1   ��ͥ�绰
    '   pat_home_postcode   C   1   ��ͥ��ַ�ʱ�
    '   region  C   1   ����
    '   pat_baddr   C   1   �����ص�
    '   pat_hous_addr   C   1   ���ڵ�ַ
    '   pat_hous_postcode   C   1   ���ڵ�ַ�ʱ�
    '   pat_grdn_name   C   1   �໤��
    '   phone_number    C   1   �ֻ���
    '   insurance_num   C   1   ҽ����
    '   iccard_no   C   1   Ic����
    '   create_time C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
    '   operator_name   C   1   ����Ա����
    '   idcard_sign N       ���֤ǩԼ
    '   idcard_sign_pwd C       ǩԼ����
    '   insurance_type  C   1   ����
    '   cert_no_other   C   1   ����֤��
    '   contacts    C       ������ϵ����Ϣ�ڵ�
    '       name    C   1   ��ϵ������
    '       idcard  C   1   ��ϵ�����֤��
    '       phone   C   1   ��ϵ�˵绰
    '       relation    C   1   ��ϵ�˹�ϵ
    '       address C       ��ϵ�˵�ַ
    '   community_info  C       ������Ϣ�ڵ�
    '       num N   1   �������
    '       code    C   1   ��������
    '       oper_type   N   1   ������������
    '   visit_info  C       ������Ϣ�ڵ�
    '       statu   N       ���µľ���״̬
    '       room    C       ���µľ�������
    '       time    C       ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '   addr_list[] C       ��ַ��Ϣ�б�
    '       oper_fun    N   1   ��������:1-����,�޸�   2-ɾ��
    '       type    C   1   ��ַ���
    '       state   C   1   ��ַ_ʡ
    '       city    C   1   ��ַ_��
    '       county  C   1   ��ַ_��
    '       township    C   1   ��ַ_��
    '       other   C   1   ��ַ_����
    '       code    C   1   ��������
    '   ext_list[]  C       ������Ϣ�����б�
    '       info_name   C   1   ��Ϣ��
    '       upd_info_value  N   1   �޸ĵ���Ϣֵ
    '   cert_list[]         ֤���б�(��Ҫ�ǵ��ɰ󿨴���)
    '       cert_name   C   1   ֤������
    '       cert_no C   1   ֤�ź���
    '   allergic_drugs_list[]           ���˹���ҩ���б�:������ʱ������ɾ������ҩ�����ķ�ʽ
    '       pat_algc_cadn_id    N   1   ����ҩƷID
    '       pat_algc_cadn   C   1   ����ҩ������
    '       allergy_info    C   1   ��ÿҩ�ﷴӦ
    '       immune_list[]   C       ���������б�
    '       vaccinate_time  C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '       vaccinate_name  C   1   ��������
    '   card_property_list[]    C       ҽ�ƿ������б�
    '       cardtype_id N   1   ҽ�ƿ����ID
    '       card_no C   1   ����
    '       info_name   C   1   ��Ϣ��
    '       info_value  N   1   ��Ϣֵ
    '
    strJson = Mid(strJson, 2)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_NewPatiArchives"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, True, , , , True) = False Then Exit Function
    '    output
    '        code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '        message C   1   "Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    zl_PatiSvr_NewPatiArchives = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
