Attribute VB_Name = "mdlPubServiceCis"
Option Explicit


'*********************************************************************************************************************************************
'����:�����漰�����ٴ�����ط���
'�ӿ�˵��:
'    1.zl_CisSvr_ExistAdvice-�ж�ָ���ĹҺŵ������Ƿ����ҽ������
'    2.zl_CisSvr_UpdateOutMedRecord-�������ﲡ����¼
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

Public Function zl_CisSvr_ExistAdvice(ByVal str�Һŵ� As String, ByVal lng����ID As Long, ByRef bln����ҽ��_Out As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional lngModule As Long, Optional str��ҳid As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ���ĹҺŵ������Ƿ����ҽ������
    '���:str�Һŵ�-ָ���ĹҺŵ�
    '     lng����ID-����id
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '     str��ҳid-��ҳid=""ʱ����ʾ������ҳid����;0ʱ��ʾֻ����ҳidΪ���ҽ��,>0��ʾ��ѯָ����ҳ��ҽ��
    '����:bln����ҽ��_Out
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
    
    'zl_CisSvr_ExistAdvice
    '  --����:�ж��Ƿ����ҽ�����ݻ��ж�ָ���Һŵ��Ƿ��Ѿ���ҽ��
    '  --��Σ�Json_In:��ʽ
    '  -- input
    '  --   pati_id              N 1 ����ID
    '  --   pati_pageid          N   ��ҳId
    '  --   rgst_no              C 1 �Һŵ�������ö��ŷָ�
    '  --   only_valid           N   ֻ���û�����ϵ�ҽ��
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code               C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    exist              N 1 �Ƿ���ڣ�1-����;0-������
    strJson = strJson & "" & GetJsonNodeString("rgst_no", str�Һŵ�, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng����ID, Json_num)
    If str��ҳid <> "" Then
        strJson = strJson & "" & GetJsonNodeString("pati_pageid", Val(str��ҳid), Json_num)
    End If
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_CisSvr_ExistAdvice"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡָ�������µ�ҽ�����ݣ����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln����ҽ��_Out = Val(NVL(objServiceCall.GetJsonNodeValue("output.exist"))) = 1
    zl_CisSvr_ExistAdvice = True
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




Public Function zl_CisSvr_UpdateOutMedRecord(ByVal cllOutMedRec As Collection, Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ﲡ����¼
    '���:cllOutMedRec-���ﲡ�����ݼ�:array(����,ֵ)
    '                ���ư���������id,.������(�����),��������,�������,�洢״̬,���λ��)
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    
    If cllOutMedRec Is Nothing Then Exit Function
    If cllOutMedRec.Count = 0 Then Exit Function
    If blnShowErrMsg Then On Error GoTo errHandle:
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        Err.Raise -1001, strErrMsg, strErrMsg
        Exit Function
    End If
    
    For i = 1 To cllOutMedRec.Count
        varTemp = cllOutMedRec(i)
        Select Case UCase(varTemp(0))
        Case "����ID"
            strJson = strJson & "," & GetJsonNodeString("pati_id", Val(varTemp(1)), Json_num)
        Case "������", "�����"
            strJson = strJson & "," & GetJsonNodeString("mr_no", Trim(varTemp(1)), Json_Text)
        Case "��������", "��������", "�Ǽ�����", "�Ǽ�ʱ��"
            strJson = strJson & "," & GetJsonNodeString("create_date", Trim(varTemp(1)), Json_Text)
        Case "�������"
            strJson = strJson & "," & GetJsonNodeString("mr_type", Trim(varTemp(1)), Json_Text)
        Case "�洢״̬"
            strJson = strJson & "," & GetJsonNodeString("strgloc_status", Trim(varTemp(1)), Json_Text)
        Case "���λ��"
            strJson = strJson & "," & GetJsonNodeString("strgloc", Trim(varTemp(1)), Json_Text)
        Case Else
        End Select
    Next
    If strJson = "" Then Exit Function
    
    'zl_CisSvr_UpdateOutMedRecord
    '    input
    '       pati_id N   1   ����id
    '       mr_no   C   1   �����ţ�����ţ�
    '       create_date C   1   ��������
    '       mr_type C   1   �������
    '       strgloc_status  C   1   �洢״̬
    '       strgloc C   1   ���λ��
    
    strJson = Mid(strJson, 2)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_CisSvr_UpdateOutMedRecord"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, blnShowErrMsg, , , , blnShowErrMsg) = False Then Exit Function
    '    output
    '        code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '        message C   1   "Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    zl_CisSvr_UpdateOutMedRecord = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_CIsSvr_GetPatiPageInfo(ByVal int��ѯ��� As Integer, ByVal strPatiInfo As String, ByRef cllPatiPages_Out As Collection, _
    Optional ByVal bln��Ӥ����Ϣ As Boolean = False, Optional ByVal bln��ת����Ϣ As Boolean = False, Optional blnȡ���һ��סԺ As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID����ҳID����ѯ������Ϣ
    '���:int��ѯ���-:0-������Ϣ;1-������Ϣ��չ;2-��ȡ��ҳ
    '     strPatiInfo-������Ϣ,��ʽ����:һ����:����id:��ҳID,��;һ�֣�����id,��
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:cllPatiPages_Out-������ҳ��Ϣ
    '     strErrmsg_Out-���ش�����Ϣ
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
    
    'zl_CIsSvr_GetPatiPageInfo
    ' input
    '    query_type  C   1   ��ѯ����:0-������Ϣ;1-������Ϣ��չ;2-��ȡ��ҳ
    '    pati_pageids    C   1   ������Ϣ,��ʽ����:һ����:����id:��ҳID,��;һ�֣�����id,��
    '    is_babyinfo N   1   �Ƿ����Ӥ����Ϣ:1-����;0-������
    '    is_transdeptinfo    N   1   �Ƿ����ת����Ϣ:1-����;0-������
    '    is_lastpage N   1   �Ƿ�ȡ���һ��סԺ
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", int��ѯ���, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", strPatiInfo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("is_babyinfo", IIf(bln��Ӥ����Ϣ, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("is_transdeptinfo", IIf(bln��ת����Ϣ, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("is_lastpage", IIf(blnȡ���һ��סԺ, 1, 0), Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_CIsSvr_GetPatiPageInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, blnShowErrMsg) = False Then Exit Function
    '����            json    ����    ��չ    ֻȡ��ҳ
    'output
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�  ��  ��  ��
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ ��  ��  ��
    '    page_list[]     1   ������
    '        pati_id N   1   ����id  ��  ��  ��
    '        pati_pageid N   1   ��ҳid  ��  ��  ��
    '        pati_name   C   1   ����    ��  ��  ��
    '        pati_sex    C   1   �Ա�    ��  ��
    '        pati_age    C   1   ����    ��  ��
    '        fee_category    C   1   �ѱ�    ��  ��
    '        mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����        ��
    '        mdlpay_mode_code    C   1   ҽ�Ƹ��ʽ����        ��
    '        pati_bed    C   1   ��ǰ����        ��
    '        pati_type   C   1   ��������(��ͨ��ҽ��������)      ��
    '        pati_education  C   1   ѧ��        ��
    '        ocpt_name   C   1   ְҵ        ��
    '        country_name    C   1   ����        ��
    '        pati_marital_cstatus    C   1   ����״��        ��
    '        pati_nature N   1   ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���   ��  ��
    '        audit_sign  N   1   ��˱�־:������ҳ.��˱�־  ��  ��
    '        si_inp_status   N   1   סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)  ��  ��
    '        pati_wardarea_id    N   1   ��ǰ����id      ��
    '        pati_wardarea_name  C   1   ��ǰ��������        ��
    '        pati_dept_id    N   1   ��ǰ����id      ��
    '        pati_dept_name  C   1   ��ǰ��������        ��
    '        adta_time   C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
    '        adtd_time   C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
    '        insurance_type  N   1   ����        ��
    '        scheme_type C   1   ���ò���:Zl_Patiwarnscheme      ��
    '        garnt_money N   1   ������:Zl_Patientsurety     ��
    '        catalog_date    C   1   ��Ŀ����:yyyy-mm-dd hh24:mi:ss      ��
    '        baby_list[]     1   Ӥ����Ϣ��[����]    is_babyinfo=1
    '            pati_id N   1   ����id
    '            pati_pageid N   1   ��ҳid
    '            baby_num    N   1   Ӥ�����
    '            baby_name   C   1   Ӥ������
    '            baby_sex    C   1   Ӥ���Ա�
    '            baby_date   C   1   ����ʱ��
    '        trans_list[]    C       ת���б���Ϣ    is_transdeptinfo=1
    '            start_reason    C   1   ��ʼԭ��
    '            start_time  C   1   ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
    '            dept_name   C   1   ��������
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܸ���������ȡ������Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set cllPatiPages_Out = objServiceCall.GetJsonListValue("output.page_list")
    zl_CIsSvr_GetPatiPageInfo = True
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



