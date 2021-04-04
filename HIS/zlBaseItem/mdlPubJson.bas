Attribute VB_Name = "mdlPubJson"
Option Explicit
Private mobjServiceCall As Object

'JSON�ڵ�����
Public Enum JSON_TYPE
    Json_Text = 0 '�ַ�
    Json_num = 1 '��ֵ
End Enum


Public Function zlGetNodeValueFromCollect(ByVal cllData As Collection, ByVal strKey As String, ByVal strType As String) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ڵ�����ݼ�
    '���:cllData-��ǰ������
    '     strKey-Key
    '     strType-"N"-����;"C"�ַ�
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-14 16:20:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    err = 0: On Error Resume Next
    varTemp = cllData(strKey)
    If err <> 0 Then
        err = 0: On Error GoTo 0
        If strType = "N" Then zlGetNodeValueFromCollect = Empty: Exit Function
        zlGetNodeValueFromCollect = "": Exit Function
    End If
    zlGetNodeValueFromCollect = varTemp
End Function

Public Function zlGetNodeObjectFromCollect(ByVal cllData As Collection, ByVal strKey As String) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ڵ�Ķ���
    '���:cllData-��ǰ������
    '     strKey-Key
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-14 16:20:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection
    err = 0: On Error Resume Next
    
    Set cllTemp = cllData(strKey)
    If err <> 0 Then
        err = 0: On Error GoTo 0
       Set zlGetNodeObjectFromCollect = cllTemp
       Exit Function
    End If
    Set zlGetNodeObjectFromCollect = cllTemp
End Function


Public Function GetJsonNodeString(ByVal strNodeName As String, ByVal strValue As String, _
    Optional ByVal intType As JSON_TYPE, Optional ByVal blnZeroToNull As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡJson�ӵ㴮
    '���:strNodeName-�ӵ���
    '     strValue-ֵ
    '     intType-����:0-�ַ�;1-����
    '     blnZeroToEmpty-�Ƿ���ֵ0ת��ΪNull��������Ϊ����ʱ��Ч
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-09 18:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    strJson = Chr(34) & strNodeName & Chr(34)
    If intType = Json_Text Then
        strJson = strJson & ":" & Chr(34) & strValue & Chr(34)
    Else
        If strValue = "" Or (blnZeroToNull And Val(strValue) = 0) Then
            strJson = strJson & ":null"
        Else
            strJson = strJson & ":" & IIF(Mid(strValue, 1, 1) = ".", "0", "") & strValue
        End If
    End If
    GetJsonNodeString = strJson
End Function
Public Function GetCollValue(ByVal colValue As Collection, ByVal varRow As Variant, Optional ByVal strElement As String) As Variant
    '���ܣ���ȡJson���鷵�صļ���������ָ���л�ָ��Ԫ�ص�ֵ
    '������
    '  varRow=���������йؼ���
    '  strElement=Ԫ����
    '���أ�
    '  ��δ����strElement����ʱ������ָ���еļ��϶��󣻵�����strElement����ʱ������ָ����ָ��Ԫ�ص�ֵ
    '  ʧ��ʱ����Nothing��Empty�������ᱨ��
    If strElement <> "" Then
        GetCollValue = Empty
    Else
        Set GetCollValue = Nothing
    End If
    
    If colValue Is Nothing Then Exit Function
    
    On Error Resume Next
    If strElement <> "" Then
        GetCollValue = colValue(varRow)(strElement)
    Else
        Set GetCollValue = colValue(varRow)
    End If
    err.Clear: On Error GoTo 0
End Function

Public Function CollectionExitsValue(ByVal coll As Collection, _
    ByVal strKey As String) As Boolean
    '���ݹؼ����ж�Ԫ���Ƿ�����ڼ�����
    Dim blnExits As Boolean

    If coll Is Nothing Then Exit Function
    CollectionExitsValue = True
    err = 0: On Error Resume Next
    blnExits = IsObject(coll(strKey))
    If err <> 0 Then err = 0: CollectionExitsValue = False
End Function


Public Function GetNodeString(ByVal strNodeName As String) As String
    GetNodeString = Chr(34) & strNodeName & Chr(34)
End Function


Private Function GetServiceCall(ByRef objServiceCall_Out As Object, Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����������
    '����:objServiceCall_Out-���ع����������
    '����:��ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String
    If Not mobjServiceCall Is Nothing Then Set objServiceCall_Out = mobjServiceCall: GetServiceCall = True: Exit Function
    
    err = 0: On Error Resume Next
    Set mobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
    If err <> 0 Then
        strErrMsg = "������zlServiceCall����ʧ������ϵͳ����Ա��ϵ���ָ��ò�����"
        If blnShowErrMsg Then
            MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
            err = 0: On Error GoTo 0
        Else
            err.Raise err.Number, err.Source, strErrMsg: Exit Function
        End If
        
        err = 0: On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo ErrHandle
    If mobjServiceCall.InitService(gcnOracle, gstrDbUser, glngSys, glngModul) = False Then Set mobjServiceCall = Nothing: Exit Function
    Set objServiceCall_Out = mobjServiceCall
    GetServiceCall = True
    Exit Function
ErrHandle:
    If blnShowErrMsg = False Then
        err.Raise err.Number, err.Source, err.Description: Exit Function
    End If
    
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlDrugSvr_GetPharmacyWindows(ByVal strҩ��IDs As String, ByRef rsData As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩ���ķ�ҩ����
    '���:
    '   strҩ��IDs ҩ��ID�������Ӣ�Ķ��ŷָ�
    '����:
    '   rsData �ֶΣ�ҩ��ID,��ҩ����,�Ƿ�ר��
    '����:��ȡ�ɹ�����True����ȡʧ�ܷ���False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim cllData As Collection, cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    
    Set rsData = New ADODB.Recordset
    With rsData.Fields
        .Append "ҩ��ID", adBigInt, 18, adFldIsNullable
        .Append "��ҩ����", adLongVarChar, 50, adFldIsNullable
        .Append "�Ƿ�ר��", adInteger, 2, adFldIsNullable
    End With
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open
    
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "����ҩƷ����ʧ�ܣ��޷���ȡҩ���ķ�ҩ���ڣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Drugsvr_Getpharmacywindows
    '  --���ܣ���ȡҩ�����漰�ķ�ҩ����
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --    pharmacy_ids            C   1  ҩ��ID1,ҩ��ID2��
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code                    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --    message                 C   1   ÿ��ҩ��id��Ӧ�ķ�ҩ����[����]
    '  --    window_list[]    ���������б�[����]
    '  --        pharmacy_id             N 1 ҩ��ID
    '  --        pharmacy_window         C 1 ��ҩ����
    '  --        expert_window           N 1 �Ƿ�ר�Ҵ��ڣ�1-�ǣ�0-����
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pharmacy_ids", strҩ��IDs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("zl_DrugSvr_Getpharmacywindows", strJson, , "", glngModul) = False Then Exit Function
    
    Set cllData = objServiceCall.GetJsonListValue("output.window_list")
    If cllData Is Nothing Then Exit Function
    
    For i = 1 To cllData.Count
        Set cllTemp = cllData(i)
        rsData.AddNew
        rsData!ҩ��Id = cllTemp("_pharmacy_id")
        rsData!��ҩ���� = cllTemp("_pharmacy_window")
        rsData!�Ƿ�ר�� = cllTemp("_expert_window")
        rsData.Update
    Next
    If rsData.RecordCount > 0 Then rsData.MoveFirst

    zlDrugSvr_GetPharmacyWindows = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlExseSvr_UpdRgstArrangeMent(ByVal int�������� As Integer, ByVal lngҽ��ID As Long, _
                Optional ByVal str����ʱ�� As String, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Դ����Ч�İ��š���Ч�ĳ����¼�е�ҽ��������
    '���:int��������-1-�޸�����,2-ͣ����Ա,3-������Ա
    '     str����ʱ��-ͣ�ú�����ʱ���룬����ʱ����ԭ����ʱ��
    '����:strErrMsg_Out
    '����:��ȡ�ɹ�����True����ȡʧ�ܷ���False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim intReturn As Integer
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���ӷ��÷���ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
'    Zl_ExseSvr_UpdRgstArrangement
'    --���ܣ�������Դ����Ч�İ��š���Ч�ĳ����¼�е�ҽ��������
'    --���
'    --input      ������Դ����Ч�İ��š���Ч�ĳ����¼�е�ҽ������
'    --  oper_type     N  1  ������ʽ��1-�޸�����,2-ͣ����Ա,3-������Ա
'    --  rgst_dr_id      N  1  ����id
'    --  revoke_time   C         ����ʱ��
'    --����
'    --output
'    --  code          C    1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'    --  message         C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_type", int��������, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rgst_dr_id", lngҽ��ID, Json_num)
    If str����ʱ�� <> "" Then
        strJson = strJson & "," & GetJsonNodeString("revoke_time", str����ʱ��, Json_Text)
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_ExseSvr_UpdRgstArrangement", strJson, , "", glngModul, False) = False Then Exit Function
    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    If intReturn <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out <> "" Then strErrMsg_Out = "���¹ҺŰ���ʧ�ܣ�"
        Exit Function
    End If
    
    zlExseSvr_UpdRgstArrangeMent = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
