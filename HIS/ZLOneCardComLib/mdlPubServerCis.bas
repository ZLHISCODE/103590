Attribute VB_Name = "mdlPubServerCis"
Option Explicit

'*********************************************************************************************************************************************
'����:�����漰�����ٴ�����ط���
'�ӿ�˵��:
'    1.zl_CisSvr_GetPatPageInfByRange-������ȡ������Ϣ����
'    2.zl_CisSvr_GetPatiID:���ݴ��ŵȻ�ȡ����ID
'    3.zl_CIsSvr_GetPatiPageInfo-��ȡ���˲�����Ϣ
'    4.zl_CisSvr_UpdateOutMedRecord-�޸����ﲡ����¼��Ϣ
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
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zl_CisSvr_GetPatPageInfByRange(ByVal intQueryStatus As Integer, ByVal cllFilter As Collection, Optional ByVal str����Ids As String, Optional ByRef str����IDs As String, _
    Optional ByRef cllPatiPages_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ����Χ��������ȡ���˲�����Ϣ
    '���:intQueryStatus-��ѯ����(0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ )
    '     cllFilter-��������
    '     str����Ids-����ö���:����ID����ID:��ҳID
    '     rsPatiPage-��ҳ��Ϣ
    '     str����IDs-��ǰ����Ids
    '����:rsPatiPageInfo_Out-���صĲ�����Ϣ��
    '     strPatiIds_Out-���ص�ǰ���漰�Ĳ���IDs
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-30 21:23:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim lng����ID As Long, strErrMsg As String
    
    
    On Error GoTo errHandle
    
    Set cllPatiPages_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "�����ٴ������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'zl_CisSvr_GetPatPageInfByRange
    '    input
    '        query_type  N   1   ��ѯ����:0-����;1-������չ
    '        wararea_ids C       ����ids:����ö���
    '        pati_ids    C       ����ids:����ö��ŷ���
    '        pati_pageids    C       ��ҳIDs:����id:��ҳid,��
    '        adta_start_time C       ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
    '        adta_end_time   C       ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '        adtd_start_time C       ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
    '        adtd_end_time   C       ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '        fee_category    C       �ѱ�
    '        inp_status  N       סԺ״̬:0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ
    '        pati_natures    C       "�������ʣ�����ö��ŷ���
    '        0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
    '        NULL-��ʾ������"

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", intQueryStatus, Json_num)
    If InStr(str����Ids, ":") > 0 Then
        strJson = strJson & "," & GetJsonNodeString("pati_pageids", str����Ids, Json_Text)  '������id+��ҳid����
    Else
        strJson = strJson & "," & GetJsonNodeString("pati_ids", str����Ids, Json_Text)
    
    End If
    strJson = strJson & "," & GetJsonNodeString("wararea_ids", str����IDs, Json_Text)
    
    For i = 1 To cllFilter.count
    
        Select Case cllFilter(i)(0)
        Case "��Ժ����"
            strJson = strJson & "," & GetJsonNodeString("adta_start_time", cllFilter(i)(1), Json_Text)
            strJson = strJson & "," & GetJsonNodeString("adta_end_time", cllFilter(i)(2), Json_Text)
        Case "��Ժ����"
            strJson = strJson & "," & GetJsonNodeString("adtd_start_time", cllFilter(i)(1), Json_Text)
            strJson = strJson & "," & GetJsonNodeString("adtd_end_time", cllFilter(i)(2), Json_Text)
        Case "�ѱ�"
            strJson = strJson & "," & GetJsonNodeString("fee_category", cllFilter(i)(1), Json_Text)
        Case "��������"
            strJson = strJson & "," & GetJsonNodeString("pati_natures", cllFilter(i)(1), Json_Text)
        End Select
    Next
    
    strJson = "{""input"":{" & strJson & "}}"
    strServiceName = "zl_CisSvr_GetPatPageInfByRange"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    '����            json    ����    ��չ
    'output
    '    code    N       Ӧ���룺0-ʧ�ܣ�1-�ɹ�  ��  ��
    '    message C       Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ ��  ��
    '    page_list[]         ������  ��  ��
    '        pati_id N       ����id  ��  ��
    '        pati_pageid N       ��ҳid  ��  ��
    '        pati_name   C       ����    ��  ��
    '        pati_sex    C       �Ա�    ��  ��
    '        pati_age    C       ����    ��  ��
    '        inpatient_num   C       סԺ��  ��  ��
    '        pati_bed    C       ��ǰ����    ��  ��
    '        insurance_type  N       ����    ��  ��
    '        fee_category    C       �ѱ�    ��  ��
    '        pati_type   C       ��������(��ͨ��ҽ��������)  ��  ��
    '        adta_time   C       ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
    '        adtd_time   C       ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
    '        si_inp_status   N       סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)  ��  ��
    Set cllData = objServiceCall.GetJsonListValue("output.page_list")
    
    If cllData Is Nothing Then
        strErrMsg = "δ�ҵ����������Ĳ�����Ϣ�����飡"
         MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If cllData.count = 0 Then
          strErrMsg = "δ�ҵ����������Ĳ�����Ϣ�����飡"
          MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
          Exit Function
    End If
    Set cllPatiPages_Out = cllData
    zl_CisSvr_GetPatPageInfByRange = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zl_CisSvr_GetPatiID(cllFindCons As Collection, _
    Optional ByRef lng��ҳID_out As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݴ��ż�סԺ�Ż�ȡ����ID
    '���:
    '   cllFindCons-��������:array(��ѯ������,��ѯ������)
    '               ��ѯ������:��:סԺ��,���ۺ�,(����ID,����)��
    '����:
    '   lng��ҳID_out-���ص�ǰ���˵���ҳID
    '����:�ɹ����ز���ID,���򷵻�False
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHandle
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'zl_CisSvr_GetPatiID
    '  --���ݴ��š�סԺ�ʻ�ȡ����ID����ҳID
    '  --input
    '  --   wardarea_id          N 1 ��ǰ����id
    '  --   pati_bed             C 1 ��ǰ����
    '  --   inpatient_num        C 1 סԺ��
    '  --   obsv_no              C 1 ���ۺ�
    '  --output
    '  --    code                N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --    message             C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    pati_id             N 1 ����ID:δ�ҵ�ʱҲ�ɹ�������0
    '  --    pati_pageid         N   ��ҳID
    strJson = ""
    For i = 1 To cllFindCons.count
        Select Case cllFindCons(i)(0)
        Case "����ID"
            strJson = strJson & "," & GetJsonNodeString("wardarea_id", cllFindCons(i)(1), Json_num)
        Case "����"
            strJson = strJson & "," & GetJsonNodeString("pati_bed", cllFindCons(i)(1), Json_Text)
        Case "סԺ��"
            strJson = strJson & "," & GetJsonNodeString("inpatient_num", cllFindCons(i)(1), Json_Text)
        Case "���ۺ�"
            strJson = strJson & "," & GetJsonNodeString("obsv_no", cllFindCons(i)(1), Json_Text)
        End Select
    Next
    strJson = "{""input"":{" & Mid(strJson, 2) & "}}"
  
    strServiceName = "zl_CisSvr_GetPatiID"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    lng��ҳID_out = Val(NVL(objServiceCall.GetJsonNodeValue("output.pati_pageid")))
    zl_CisSvr_GetPatiID = Val(NVL(objServiceCall.GetJsonNodeValue("output.pati_id")))
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_CIsSvr_GetPatiPageInfo(ByVal int��ѯ���� As Integer, ByVal str������ҳIDs As String, ByRef cllPatiPage_Out As Variant, _
    Optional ByRef bln��ȡ���סԺ As Boolean, Optional bln��Ӥ����Ϣ As Boolean, Optional ByRef bln��ת����Ϣ As Boolean, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID����ҳid��Ϣ����ȡ������ҳ��Ϣ
    '���:int��ѯ����-0-ֻ��ȡ������Ϣ;1-��ȡ������Ϣ+��չ��Ϣ;2-����ȡȡ��ҳID�ֶ�
    '     str������ҳIDs-���ָ�ʽ:
    '           1.����id1:��ҳid1,����id2:��ҳid2...
    '           2.����id1,����id2,...����idn
    '      bln��ȡ���סԺ:����ȡ�������һ�εĲ���,(str������ҳIDs�ڶ��ָ�ʽ��Ч)
    '      bln��Ӥ����Ϣ:�Ƿ����Ӥ����Ϣ
    '      bln��ת����Ϣ:�Ƿ��ת����Ϣ
    '����:cllPatiPageInfo_Out-���صĲ�����Ϣ��
    '     strErrMsg_Out-���صĴ�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2012-09-19 15:50:18
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer

    On Error GoTo errHandle
    
    Set cllPatiPage_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    'input
    '    query_type  C   1   ��ѯ����:0-������Ϣ;1-������Ϣ��չ;2-��ȡ��ҳ
    '    pati_pageids    C   1   ������Ϣ,��ʽ����:һ����:����id:��ҳID,��;һ�֣�����id,��
    '    is_lastpage N   1   �Ƿ�ȡ���һ��סԺ
    '    is_babyinfo N   1   �Ƿ����Ӥ����Ϣ:1-����;0-������
    '    is_transdeptinfo    N   1   �Ƿ����ת����Ϣ:1-����;0-������

    strJson = strJson & "" & GetJsonNodeString("query_type", int��ѯ����, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", str������ҳIDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("is_lastpage", IIf(bln��ȡ���סԺ, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("is_babyinfo", IIf(bln��Ӥ����Ϣ, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("is_transdeptinfo", IIf(bln��ת����Ϣ, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    strServiceName = "zl_CIsSvr_GetPatiPageInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    '    ����            json    ����    ��չ    ֻȡ��ҳ
    '    output
    '        code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�  ��  ��  ��
    '        message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ ��  ��  ��
    '        page_list[]     1   ������
    '        pati_id N   1   ����id  ��  ��  ��
    '        pati_pageid N   1   ��ҳid  ��  ��  ��
    '        pati_name   C   1   ����    ��  ��  ��
    '        pati_sex    C   1   �Ա�    ��  ��
    '        pati_age    C   1   ����    ��  ��
    '        fee_category    C   1   �ѱ�    ��  ��
    '        mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����        ��
    '        mdlpay_mode_code    C   1   ҽ�Ƹ��ʽ����        ��
    '        pati_bed    C   1   ��ǰ����
    '        pati_type   C   1   ��������(��ͨ��ҽ��������)
    '        pati_education  C   1   ѧ��
    '        ocpt_name   C   1   ְҵ
    '        country_name    C   1   ����
    '        pati_marital_cstatus    C   1   ����״��
    '        pati_nature N   1   ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���   ��
    '        audit_sign  N   1   ��˱�־:������ҳ.��˱�־  ��  ��
    '        si_inp_status   N   1   סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)  ��  ��
    '        pati_wardarea_id    N   1   ��ǰ����id      ��
    '        pati_deptid N   1   ��ǰ����id      ��
    '        pati_wardarea_id    N   1   ��ǰ����id      ��
    '        pati_wardarea_name  C   1   ��ǰ��������        ��
    '        pati_dept_id    N   1   ��ǰ����id      ��
    '        pati_dept_name  C   1   ��ǰ��������        ��
    '        adta_time   C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
    '        adtd_time   C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
    '        insurance_type  N   1   ����        ��
    '        scheme_type C   1   ���ò���:Zl_Patiwarnscheme      ��
    '        garnt_money     1   ������:Zl_Patientsurety     ��
    '        catalog date    C   1   ��Ŀ����:yyyy-mm-dd hh24:mi:ss      ��
    '        baby_list[]     1   Ӥ����Ϣ��[����]    is_babyinfo=1
    '            pati_id N   1   ����id
    '            pati_pageid N   1   ��ҳid
    '            baby_num    N   1   Ӥ�����
    '            baby_name   C   1   Ӥ������
    '            baby_sex    C   1   Ӥ���Ա�
    '            baby_date   D   1   ����ʱ��
    '        trans_list[]    C       ת���б���Ϣ    is_transdeptinfo=1
    '            start_reason    C   1   ��ʼԭ��
    '            start_time  C   1   ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
    '            dept_name   C   1   ��������
    
        
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    Set cllData = objServiceCall.GetJsonListValue("output.page_list")
'    If clldata Is Nothing Then
'            strErrMsg_Out = "δ�ҵ����������Ĳ�����Ϣ�����飡"
'        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If
'
'    If clldata.count = 0 Then
'        strErrMsg_Out = "δ�ҵ����������Ĳ�����Ϣ�����飡"
'        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If
    Set cllPatiPage_Out = cllData
    zl_CIsSvr_GetPatiPageInfo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function zl_CisSvr_UpdateOutMedRecord(ByVal cllOutMedRec As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ﲡ����¼
    '���:cllOutMedRec-���ﲡ�����ݼ�:array(����,ֵ)
    '                ���ư���������id,.������(�����),��������,�������,�洢״̬,���λ��)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    
    If cllOutMedRec Is Nothing Then Exit Function
    If cllOutMedRec.count = 0 Then Exit Function
    
    On Error GoTo errHandle
    If GetServiceCall(objServiceCall, False) = False Then
        strErrMsg = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        Err.Raise -1001, strErrMsg, strErrMsg
        Exit Function
    End If
    
    For i = 1 To cllOutMedRec.count
        varTemp = cllOutMedRec(i)
        Select Case UCase(varTemp(0))
        Case "����ID"
            strJson = strJson & "," & GetJsonNodeString("pati_id", Val(varTemp(1)), Json_num)
        Case "������"
            strJson = strJson & "," & GetJsonNodeString("mr_no", varTemp(1), Json_Text, True)
        Case "�����"
            strJson = strJson & "," & GetJsonNodeString("outpatient_num", varTemp(1), Json_Text, True)
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
    '       mr_no   N   1   �����ţ�����ţ�
    '       create_date C   1   ��������
    '       mr_type C   1   �������
    '       strgloc_status  C   1   �洢״̬
    '       strgloc C   1   ���λ��
    
    strJson = Mid(strJson, 2)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_CisSvr_UpdateOutMedRecord"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    '    output
    '        code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '        message C   1   "Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    zl_CisSvr_UpdateOutMedRecord = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function Zl_CisSvr_PatiIsInhospital(ByVal lng����ID As Long, ByRef blnInhospital As Boolean, _
                Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��鲡���Ƿ���Ժ����
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/18 14:35
    '---------------------------------------------------------------------------------------
    Dim intReturn As Integer
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHandle
    blnInhospital = False
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_CisSvr_Patiisinhospital
    '    input
    '        pati_id            N 1 ����ID
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_CisSvr_PatiIsInhospital"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, Not blnNotShowErrMsg, strErrMsg) = False Then Exit Function
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
         
    blnInhospital = Val(objServiceCall.GetJsonNodeValue("output.inhouspital")) = 1
    Zl_CisSvr_PatiIsInhospital = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function


Public Function zl_Cissvr_Existadvice(ByVal lng����ID As Long, ByVal str�Һŵ� As String, ByRef blnHavAdvice As Boolean, _
                Optional ByVal lng��ҳId As Long, Optional ByVal blnOnlyValid As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ���Һŵ��Ƿ�����ҽ��
    ' ��� : str�Һŵ�-������ݺż��ö��ŷָ�
    '        blnOnlyValid-�Ƿ�ֻ�����Чҽ��
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2019/11/18 19:48
    '---------------------------------------------------------------------------------------
    
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'zl_CisSvr_GetPatiID
    '    input
    '    --   pati_id              N 1 ����ID
    '    --   pati_pageid          N   ��ҳId
    '    --   rgst_no              C 1 �Һŵ�������ö��ŷָ�
    '    --   only_valid           N   ֻ���û�����ϵ�ҽ��
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rgst_no", str�Һŵ�, Json_Text)
    If lng��ҳId <> 0 Then
        strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng��ҳId, Json_num)
    End If
    strJson = strJson & "," & GetJsonNodeString("only_valid", IIf(blnOnlyValid, 1, 0), Json_num)
    
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Cissvr_Existadvice"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    blnHavAdvice = Val(objServiceCall.GetJsonNodeValue("output.exist")) = 1
         
    zl_Cissvr_Existadvice = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Zl_Cissvr_GetPatiVitalSigns(ByVal lng����ID As Long, ByVal lng�Һ�ID As String, _
                ByRef cllVital As Collection, Optional ByVal blnOutPati As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��ȡ��������������Ϣ
    ' ��� : blnOutPati-���ﲡ��
    ' ���� : cllVital:������Ϣ(Collect)(��Ŀ,ֵ,��λ)
    ' ���� : ���ز��˵�������Ϣ��������Ŀ����ֵ����λ
    ' ���� : ���ϴ�
    ' ���� : 2019/11/18 19:48
    '---------------------------------------------------------------------------------------
    
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���Ӳ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Cissvr_GetPatiVitalSigns
    '    input
    '    --   pati_id              N 1 ����ID
    '    --   visit_id             N 1 �Һ�ID
    '    --   outpati_flag         N   �����־��1-���2-סԺ
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("visit_id", lng�Һ�ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("outpati_flag", IIf(blnOutPati, 1, 0), Json_num)
    
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Cissvr_GetPatiVitalSigns"
    If objServiceCall.CallService(strServiceName, strJson, strServiceName, "", glngModul) = False Then Exit Function
    Set cllVital = objServiceCall.GetJsonListValue("output.pativital_list")
    
    Zl_Cissvr_GetPatiVitalSigns = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Zl_Cissvr_Checkdepositno(ByVal lng����ID As Long, ByRef strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�Ԥ��NO�Ƿ����"���˽����쳣��¼"��
    '���:strNo-Ԥ�����ݺ�
    '����:�����NO�Ŵ���"���˽����쳣��¼"�з���true,���򷵻�False
    '����:����
    '����:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "�����ٴ������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Cissvr_Checkdepositerrorno
    '    input
    '    --   pati_id              N 1 ����ID
    '    --   bill_nos             C 1 ����Ԥ����¼.NO
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_nos", strNo, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Cissvr_Checkdepositerrorno"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    strNo = objServiceCall.GetJsonNodeValue("output.bill_nos")
    Zl_Cissvr_Checkdepositno = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Zl_Cissvr_GetPatiBaseInfo(ByVal lng����ID As Long, Optional ByVal lng��ҳId As Long = -1, _
                Optional ByRef cllPati_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����סԺ��Ϣ
    '���:lng����ID��lng��ҳID
    '����:cllPati_Out ������Ϣ
    '����:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHand
     
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "�����ٴ������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Cissvr_Checkdepositerrorno
'      input
'  --    query_type        N 1 ��ѯ��ʽ-- 1-ͨ������ID+��ҳID��ѯ������Ϣ,2-ͨ��ҽ��ID��ȡ���˻�����Ϣ ,3-ͨ���Һŵ���ȡ���˻�����Ϣ
'  --    pati_id           N   ����id--
'  --    page_id           N   ��ҳid--
'  --    advice_id         N   ҽ��ID--
'  --    pati_type         N   0-סԺ���� 1-���ﲡ��
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    If lng��ҳId <> -1 Then
        strJson = strJson & "," & GetJsonNodeString("page_id", lng��ҳId, Json_num)
    End If
    strJson = strJson & "," & GetJsonNodeString("query_type", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_type", 0, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
  
    strServiceName = "Zl_Cissvr_Getpatibaseinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    Set cllPati_Out = objServiceCall.GetJsonListValue("output.page_list")
    Zl_Cissvr_GetPatiBaseInfo = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zl_Cissvr_GetInpatiState(ByVal lng����ID As Long, ByVal lng��ҳId As Long, _
                Optional ByVal intPatiType As Integer, _
                Optional ByRef cllPati_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����סԺ״̬
    '���:lng����ID��lng��ҳID
    '     intPatiType:�������� 0-��ͨסԺ���� 1-�������۲��� 2-סԺ���۲���
    '����:cllPati_Out ������Ϣ
    '����:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object

    On Error GoTo errHand
    
    Set cllPati_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "�����ٴ������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Cissvr_Checkdepositerrorno
'      input
'    pati_id       N   1   ����ID
'    pati_pageid   N   1   ��ҳid
'    pati_type     N   1   �������� 0-��ͨסԺ���� 1-�������۲��� 2-סԺ���۲���

    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng��ҳId, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_type", intPatiType, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    strServiceName = "zl_Cissvr_GetInpatiState"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    If objServiceCall.GetJsonNodeValue("output.pati_type") = "" Then Exit Function
    If Val(objServiceCall.GetJsonNodeValue("output.pati_type")) <> intPatiType Then Exit Function
    cllPati_Out.Add Val(NVL(objServiceCall.GetJsonNodeValue("output.pati_state"))), "����״̬"
    cllPati_Out.Add NVL(objServiceCall.GetJsonNodeValue("output.out_time")), "��Ժ����"
    zl_Cissvr_GetInpatiState = True
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function
