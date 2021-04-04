Attribute VB_Name = "mdlPubServerExse"
Option Explicit

'*********************************************************************************************************************************************
'����:�����漰�����ٴ�����ط���
'�ӿ�˵��:
'    1.zl_ExseSvr_GetPatiSurplusInfo-��ȡ���������Ϣ
'    2.zl_ExseSvr_GetConsumerCardType-��ȡ���ѿ������Ϣ�ӿ�
'    3.Zl_Exsesvr_GetConsumerCardInfo���ݿ��źͽӿڱ�Ż�ȡ���ѿ���Ϣ
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
    Call gobjComLib.aveErrLog
End Function


Public Function zl_ExseSvr_GetPatiSurplusInfo(ByVal str����Ids As String, ByRef cllSurplusData_Out As Collection, _
    Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByVal strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���˷��������Ϣ
    '���:str����Ids-����ID,����ö��ŷ���
    '     blnNotShowErrMsg-�Ƿ���ʾ������Ϣ��,true-����ʾ;false-��ʾ
    '����:cllSurplusData_Out-���ز�����Ϣ��
    '     strErrMsg_out-����ʾʱ�����ش�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    On Error GoTo errHandle
    
    Set cllSurplusData_Out = New Collection
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "�����ٴ������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    

    'zl_ExseSvr_GetPatiSurplusInfo
    'input           ��ȡ���˵��շ����ķ����ܶ�
    '    pati_ids    C       ����ID,����ö��ŷ���
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_ids", str����Ids, Json_Text)
     strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetPatiSurplusInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    'output
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    surplus_list[]  C   1   ����б�
    '        pati_Id N       ����ID
    '        outdpst_surplus N   1   ����Ԥ�����
    '        indpst_surplus  N   1   סԺԤ�����
    '        outfee_surplus  N   1   ����������
    '        infee_surplus   N   1   סԺ�������
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set cllSurplusData_Out = objServiceCall.GetJsonListValue("output.surplus_list", "pati_id")
    zl_ExseSvr_GetPatiSurplusInfo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zl_ExseSvr_GetConsumerCardType(ByRef cllTypesData_out As Collection, Optional ByVal blnOnlyStart As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ѿ������Ϣ����ӿ�
    '���:blnOnlyStart-ֻ��ȡ���õĿ����
    '����:cllTypesData_out-���ؿ������Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:47:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim cllData As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    
    On Error GoTo errHandle
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�����ѿ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetConsumerCardType
    '  --����:��ȡ���ѿ����
    '  --��Σ�Json_In:��ʽ
    '  --    input
    '  --      enabled                N    �Ƿ�����:1-������;0-����
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("enabled", IIf(blnOnlyStart, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetConsumerCardType"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul) = False Then Exit Function
    
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code                      N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --    message                   C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    type_list[]               C  1  ֧�ֵĿ�����б�
    '  --          cardtype_id       N  1  id
    '  --          cardtype_num      N  1  ���
    '  --          cardtype_name     C  1  ����
    '  --          cardtype_stname   C  1  ����
    '  --          prefix_text         C  1  ǰ׺�ı�
    '  --          cardno_len          N  1  ���ų���
    '  --          default             N  1  ȱʡ��־
    '  --          fixed               N  1  �Ƿ�̶�:1-��ϵͳ�̶�;0-����ϵͳ�̶�
    '  --          strict              N  1  �Ƿ��ϸ����:1-���ϸ����;0-�����ϸ����
    '  --          self_make           N  1  �Ƿ�����:1-�ǵ�;0-����
    '  --          allow_return_cash   N  1  �Ƿ�����:1-����;0-������
    '  --          must_all_return     N   1   �Ƿ�ȫ��:1-����ȫ��;0-��������
    '  --          specpati            N   1   �ض�����
    '  --          component           C   1   ����
    '  --          memo                C   1   ��ע
    '  --          blnc_mode           C   1   ���㷽ʽ
    '  --          blnc_nature         N   1   ��������
    '  --          pwdtxt           N   1   �Ƿ�����
    '  --          enabled             N   1   �Ƿ�����:1-������;0-δ����
    '  --          pwd_len             N   1   ���볤��
    '  --          pwd_len_limit       N   1   ���볤������:0-��������;1-�̶����볤��;-n��ʾ�����������ö��λ��������,�����ܳ������볤��
    '  --          pwd_rule            N   1   �������:��-���ֺ��ַ����;1-��Ϊ�������
    '  --          readcard_nature     C   1   ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
    '  --          keyboard_mode       N   1   ���̿��Ʒ�ʽ:��0-��ֹʹ�������;1-ʹ����������� ,2-ʹ���ַ������
    '  --          def_delcash         N   1   �Ƿ�ȱʡ����:��������ʱ,Ĭ���Ƿ�����
    Set cllData = objServiceCall.GetJsonListValue("output.type_list")
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    Set cllTypesData_out = cllData
    zl_ExseSvr_GetConsumerCardType = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zl_Exsesvr_GetConsumerCardInfo(ByVal lngCardTypeID As String, ByVal strCardNo As String, ByRef dbl���_out As Double, _
    Optional lng���ѿ�ID_Out As Long, Optional strPwd_Out As String, Optional str�������_Out As String, Optional str����_Out As String, _
    Optional lng����ID_Out As Long, Optional bln�ض�����_Out As Boolean, Optional strErrMsg_Out As String, Optional bln��Ч�Լ�� As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ��źͽӿڱ�Ż�ȡ���ѿ���Ϣ
    '���:strCardNO-���ѿ���
    '       lngCardTypeID-���ѿ��ӿ����
    '����:lng���ѿ�ID_Out-���ѿ�id;strPwd_Out-���ѿ�����;dbl���_out-�������;str�������_Out-��������
    '       str����_Out-Ӧ�ó���;lng����ID_Out-����id;bln�ض�����_Out-�Ƿ����ض�����
    '       strErrMsg_out-����ʾʱ�����ش�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:����
    '����:2019-12-5 10:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall) = False Then
        strErrMsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ��"
        Exit Function
    End If
    
    'Zl_Exsesvr_Getconsumercardinfo
    '   input
    '       cardno                    C 1 ����
    '       cardtype_num         N 1 �ӿڱ��
    '       check_valid          N 1 ��鿨�ŵ���Ч��,1-���;0-�����
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("cardno", strCardNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("cardtype_num", lngCardTypeID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("check_valid", IIf(bln��Ч�Լ��, 1, 0), Json_num)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Getconsumercardinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
'    --  output
'    --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'    --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'    --    card_id               N 1 ���ѿ�id
'    --    card_pwd           C 1 ����
'    --    surplus               N 1 ���
'    --    limit_type           N 1 �������
'    --    occasion             N 1 Ӧ�ó���
'    --    pati_id                N 1 ����ID
'    --    specpati              N 1 �Ƿ��ض�����
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "δ�ҵ��������������ѿ���Ϣ�����飡"
        End If
        MsgBox strErrMsg_Out, vbOKOnly, gstrSysName
        
        Exit Function
    End If
    lng���ѿ�ID_Out = Val(NVL(objServiceCall.GetJsonNodeValue("output.card_id")))
    strPwd_Out = NVL(objServiceCall.GetJsonNodeValue("output.card_pwd"))
    dbl���_out = Val(objServiceCall.GetJsonNodeValue("output.surplus"))
    str�������_Out = NVL(objServiceCall.GetJsonNodeValue("output.limit_type"))
    str����_Out = NVL(objServiceCall.GetJsonNodeValue("output.occasion"))
    lng����ID_Out = Val(NVL(objServiceCall.GetJsonNodeValue("output.pati_id")))
    bln�ض�����_Out = Val(NVL(objServiceCall.GetJsonNodeValue("output.specpati"))) = 1
    zl_Exsesvr_GetConsumerCardInfo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

