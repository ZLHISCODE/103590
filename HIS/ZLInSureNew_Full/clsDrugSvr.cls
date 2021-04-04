VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************************************************************************************
'����:�����漰ҩƷ����������ýӿ�
'�ӿ�˵��:
'  һ����������
'     1.GetServiceCall-��ȡҩƷ���񲿼�
'  ��������ӿڲ���
'     1.ZlDrugsvr_Getadditional_Infor-��ȡҩƷ��һЩ��չ�򸽼ӵ���Ϣ���������÷���������Ƶ�Σ����͵�
'     2.zlDrugSvr_Check_IsOutTakeDrug-���ݲ���ID����ҳID,�жϸò����Ƿ��Ժ��ҩ
'     3.
'
'����:
'����:�ɹ�����true,���򷵻�False
'����:���˺�
'����:2019-08-29 19:23:18
'*********************************************************************************************************************************************

Private mobjServiceCall  As Object
Private mlngModule As Long
Private Function GetServiceCall(ByRef objServiceCall_Out As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����������
    '����:objServiceCall_Out-���ع����������
    '����:��ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    If Not mobjServiceCall Is Nothing Then Set objServiceCall_Out = mobjServiceCall: GetServiceCall = True: Exit Function
    Err = 0: On Error Resume Next
    Set mobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
    If Err <> 0 Then
        MsgBox "������zlServiceCall����ʧ������ϵͳ����Ա��ϵ���ָ��ò�����", vbInformation + vbOKOnly, gstrSysName
        Err = 0: On Error GoTo 0
        Exit Function
    End If
    If mobjServiceCall.InitService(gcnOracle, gstrDbUser, glngSys, glngModul) = False Then Set mobjServiceCall = Nothing: Exit Function
    
    Set objServiceCall_Out = mobjServiceCall
    GetServiceCall = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlDrugsvr_Getadditional_Infor(ByVal strNO As String, ByVal lng����Id As Long, _
    ByRef cllData_Out As Collection, Optional ByVal byt�������� As Byte = 2, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩƷ��һЩ��չ�򸽼ӵ���Ϣ���������÷���������Ƶ�Σ����͵�
    '���:strNo-���ݺ�
    '    str����Ids:����Id,����ö���
    '    byt��������:1-�շ��õ�;2-���ʵ�
    '
    '����::cllData_Out-(Ƶ��,�÷�,����,������λ,ִ������)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, cllData As Collection, cllTemp As Collection
    Dim objServiceCall As Object
    Dim i As Long, j As Long
 
    On Error GoTo errHandle
    
    Set cllData_Out = New Collection
    
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input
    '        billtype                    N   1   ��������:1 -�շѴ�����ҩ  ;2- ���ʵ�������ҩ
    '        rcp_no                  C   1   ���ݺ�
    '        rcpdtl_ids                  C       ������ϸids,Ŀǰ����ķ���ID
    '    ���� json
    '    output
    '        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
    '        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '        item_list[]                         ���������б�[����]
    '            rcp_no              C   1   NO
    '            rcpdtl_id               N   1   ������ϸid,Ŀǰ����ķ���ID
    '            frequency               C   1   Ƶ��
    '            usage               C   1   �÷�
    '            si_drug_form                C   1   ����
    '            loitem_detail_measunit              C   1   ������λ
    '            advice_exe_properties               N   1   ִ������:0~2-�Ƽ�����,3-��Ժ��ҩ,4-��ȡҩ
 
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("billtype", byt��������, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rcp_no", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("rcpdtl_ids", lng����Id, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"

    If objServiceCall.CallService("zl_DrugSvr_GetAdditional_Infor", strJson, , "", lngModule) = False Then Exit Function
    
    Set cllData = objServiceCall.GetJsonListValue("output.item_list")
    
    For i = 1 To cllData.Count
        Set cllTemp = cllData(i)
        cllData_Out.Add cllTemp("_frequency"), "Ƶ��"
        cllData_Out.Add cllTemp("_usage"), "�÷�"
        cllData_Out.Add cllTemp("_si_drug_form"), "����"
        cllData_Out.Add cllTemp("_loitem_detail_measunit"), "������λ"
        cllData_Out.Add cllTemp("_advice_exe_properties"), "ִ������"
        Exit For
    Next
    If cllData_Out.Count = 0 Then
        cllData_Out.Add "", "Ƶ��"
        cllData_Out.Add "", "�÷�"
        cllData_Out.Add "", "����"
        cllData_Out.Add "", "������λ"
        cllData_Out.Add 0, "ִ������"
    End If
    ZlDrugsvr_Getadditional_Infor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlDrugSvr_Check_IsOutTakeDrug(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID����ҳID,�жϸò����Ƿ��Ժ��ҩ
    '���:lng����ID-����ID
    '    lng��ҳId-��ҳID
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim i As Long, j As Long
 
    On Error GoTo errHandle
    
    
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input
    '        pati_id                 N   1   ����ID
    '        pati_pageid                 N   1   ��ҳID
    '    ���� json
    '    output
    '        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
    '        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '        isexist                 N   1   �Ƿ����: 1-����;0-������
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng��ҳID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"

    If objServiceCall.CallService("Zl_Cissvr_Isouttakedrug", strJson, , "", lngModule) = False Then Exit Function
    zlDrugSvr_Check_IsOutTakeDrug = objServiceCall.GetJsonNodeValue("output.isexist") = 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlPati_Merage(ByVal lng��������Id As Long, ByVal lng�ϲ�����ID As Long, ByVal str�ϲ�ԭ�� As String, ByVal str����Ա���� As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ϲ�����
    '���:lng��������Id-��������Id
    '    lng�ϲ�����ID-�ϲ�����ID
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, str������Ϣ As String, varData As Variant, strName As String, strBirthDate As String, strCardID As String, strSex As String
    Dim strAge As String, blnTrans As Boolean, strErrMsg As String, lng����Id As Long, lng�ϲ�ID As Long
    Dim i As Long
 
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    
 gcnOracle.BeginTrans: blnTrans = True
    str������Ϣ = zlDatabase.CallProcedure("zl_������Ϣ_MERGE", "ҽ��ִ�в�����Ϣ�ϲ�", lng�ϲ�����ID, lng��������Id, str�ϲ�ԭ��, str����Ա����, 0, Empty)
    
    If str������Ϣ <> "" Then   '����id ,�ϲ�id ,����,�Ա�,����,��������,����֤��|����ҳid :ԭ����id:ԭ��ҳid
        varData = Split(Split(str������Ϣ, "|")(0), ",")
        lng����Id = varData(0)
        lng�ϲ�ID = varData(1)
        strName = varData(2)
        strSex = varData(3)
        strAge = varData(4)
        strBirthDate = varData(5)
        strCardID = varData(6)
        varData = Split(Split(str������Ϣ, "|")(1), ",")
        For i = 0 To UBound(varData)
            varData = Split(varData(i), ":")
            strJson = strJson & "," & "{""page_id_new"":" & varData(0) & ",""pati_id_befor"":" & varData(1) & ",""page_id_befor"":" & varData(2) & "}"
        Next
        If strJson <> "" Then
            strJson = Mid(strJson, 2)
            strJson = "{""input"":{""retain_id"":" & lng����Id & ",""merge_id"":" & lng�ϲ�ID & ",""pati_name"":""" & strName & """,""pati_sex"":""" & strSex & """,""pati_age"":""" & strAge & """,""pati_borth_time"":""" & strBirthDate & """,""pati_identity"":""" & strCardID & """,""item_list"":[" & strJson & "]}}"
            If zlDrugSvr_Merage(strJson, strErrMsg, lngModule) = False Then
                 gcnOracle.RollbackTrans: blnTrans = False
                 If strErrMsg <> "" Then MsgBox "���˺ϲ�ʧ�ܣ�ԭ������:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                 Exit Function
            End If
             
            '��ʱ��������ԭ���������Э�̣����ֻ�Ǹ�����Ϣ������������,�Ժ��Ϊ����󣬿��ܻ���ɲ�һ�µ����
            If ZlStuffsvr_Merage(strJson, strErrMsg, lngModule) = False Then
                 gcnOracle.RollbackTrans: blnTrans = False
                 If strErrMsg <> "" Then MsgBox "���˺ϲ�ʧ�ܣ�ԭ������:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                 Exit Function
            End If
            If ZlPivassvr_merage(strJson, strErrMsg, lngModule) = False Then
                 gcnOracle.RollbackTrans: blnTrans = False
                 If strErrMsg <> "" Then MsgBox "���˺ϲ�ʧ�ܣ�ԭ������:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                 Exit Function
            End If
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    zlPati_Merage = True
    Exit Function
errHandle:
    If blnTrans Then gcnBJYB.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 

Public Function zlDrugSvr_Merage(strJsonIn As String, Optional ByRef strErrMsg_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ϲ�����(ҩƷ)
    '���:strJsonIn-Json��
    '����:strErrMsg_Out-���ش�����Ϣ
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim i As Long, j As Long
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input           ������Ϣ�ϲ����ù���Zl_������Ϣ_Mergeʱ�Ƚ�Ҫ�ϲ��Ĳ���id����ҳid�嵥��ѯ��������ε���֯
    '       retain_id   N   1   ����id
    '       merge_id     N   1   �ϲ�id
    '       pati_name    C   1   ����
    '       pati_sex C   1   �Ա�
    '       pati_age     C   1   ����
    '       pati_borth_time  C   1   ��������
    '       pati_identity    C   1   ����֤��
    '       item_list
    '          page_id_new   N   1   ����ҳid
    '          pati_id_befor     N   1   ԭ����id
    '          page_id_befor     N   1   ԭ��ҳid
    '
    '    ���� json
    '    output
    '    code    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    If objServiceCall.CallService("zl_DrugSvr_Merage", strJsonIn, , "ҽ��������Ϣ�ϲ�", lngModule, False) = False Then Exit Function
    If Val(Nvl(objServiceCall.GetJsonNodeValue("output.code"))) = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlDrugSvr_Merage = True
    Exit Function
errHandle:
    Err.Raise Err.Number, Err.Description
End Function
 
Public Function ZlStuffsvr_Merage(strJsonIn As String, Optional ByRef strErrMsg_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ϲ�����(����)
    '���:strJsonIn-Json��
    '����:strErrMsg_Out-���ش�����Ϣ
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim i As Long, j As Long
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input           ������Ϣ�ϲ����ù���Zl_������Ϣ_Mergeʱ�Ƚ�Ҫ�ϲ��Ĳ���id����ҳid�嵥��ѯ��������ε���֯
    '       retain_id   N   1   ����id
    '       merge_id     N   1   �ϲ�id
    '       pati_name    C   1   ����
    '       pati_sex C   1   �Ա�
    '       pati_age     C   1   ����
    '       pati_borth_time  C   1   ��������
    '       pati_identity    C   1   ����֤��
    '       item_list
    '          page_id_new   N   1   ����ҳid
    '          pati_id_befor     N   1   ԭ����id
    '          page_id_befor     N   1   ԭ��ҳid
    '
    '    ���� json
    '    output
    '    code    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    If objServiceCall.CallService("Zl_Stuffsvr_Merage", strJsonIn, , "ҽ��������Ϣ�ϲ�", lngModule, False) = False Then Exit Function
    If Val(Nvl(objServiceCall.GetJsonNodeValue("output.code"))) = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    ZlStuffsvr_Merage = True
    Exit Function
errHandle:
    Err.Raise Err.Number, Err.Description
End Function
 
 
Public Function ZlPivassvr_merage(strJsonIn As String, Optional ByRef strErrMsg_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ϲ�����(����)
    '���:strJsonIn-Json��
    '����:strErrMsg_Out-���ش�����Ϣ
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim i As Long, j As Long
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input           ������Ϣ�ϲ����ù���Zl_������Ϣ_Mergeʱ�Ƚ�Ҫ�ϲ��Ĳ���id����ҳid�嵥��ѯ��������ε���֯
    '       retain_id   N   1   ����id
    '       merge_id     N   1   �ϲ�id
    '       pati_name    C   1   ����
    '       pati_sex C   1   �Ա�
    '       pati_age     C   1   ����
    '       pati_borth_time  C   1   ��������
    '       pati_identity    C   1   ����֤��
    '       item_list
    '          page_id_new   N   1   ����ҳid
    '          pati_id_befor     N   1   ԭ����id
    '          page_id_befor     N   1   ԭ��ҳid
    '
    '    ���� json
    '    output
    '    code    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    If objServiceCall.CallService("Zl_Pivassvr_merage", strJsonIn, , "ҽ��������Ϣ�ϲ�", lngModule, False) = False Then Exit Function
    If Val(Nvl(objServiceCall.GetJsonNodeValue("output.code"))) = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    ZlPivassvr_merage = True
    Exit Function
errHandle:
    Err.Raise Err.Number, Err.Description
End Function
  
  
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mobjServiceCall = Nothing
End Sub
