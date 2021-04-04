Attribute VB_Name = "mdlService"
Option Explicit

Public Enum JSON_TYPE
    Json_Text = 0 '�ַ�
    Json_num = 1 '��ֵ
End Enum

Public Function OpenJson(ByVal strJson As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:����һ��Json��
'���:strJson-Json��
'����:
'����:���óɹ�,����true,���򷵻�False
'����:��ΰ��
'����:2019-08-11 19:36:34
'---------------------------------------------------------------------------------------------------------------------------------------------
    If InitSvr() Then
        OpenJson = gobjService.SetJsonString(strJson)
    End If
End Function

Public Function GetSvrOutInfo(ByVal strJson As String, Optional ByVal blnShowErrMsg As Boolean = True, Optional ByVal strErrMsg As String) As Boolean
'���ܣ���Ա����ٴ�����鷽�����õĳ��ν�����ʽ
'������strJsonǰһ��������ִ�к�ĵĳ���Json��ʽ
'      blnShowErrMsg �Ƿ��ڲ�������ʾ��Ϣ
'���أ�true/false  code=1ʱΪtrue,code=0ʱΪfalse
    If InitSvr() Then
        Call gobjService.SetJsonString(strJson)
        If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
            If blnShowErrMsg Then
                MsgBox gobjService.GetJsonNodeValue("output.message") & "", vbInformation, gstrSysName
            Else
                strErrMsg = gobjService.GetJsonNodeValue("output.message") & ""
            End If
        Else
            GetSvrOutInfo = True
        End If
    End If
End Function

Public Function GetNode(ByVal strKey As String, ByVal varValue As Variant, Optional ByVal blnFirst As Boolean, _
Optional ByVal bytType As Byte) As String
'����:��ȡ����JSONԪ�ؼ�ֵ��
'   strKey-   Json keyֵ
'   strValue- Json valueֵ ����\�ַ���
'   blnFirst  = T ��һ���ڵ�;F-�ǵ�һ�ڵ�
'   bytType   =1 ƴ���ַ�"{}";=2 ƴ���ַ�"[]"
    Dim strTemp As String
    
    Select Case TypeName(varValue)
    
    Case "String"
        If Left(varValue, 1) = Chr(91) Or Left(varValue, 1) = Chr(123) Then
            strTemp = varValue
        Else
            strTemp = Chr(34) & zlStr.ToJsonStr(varValue) & Chr(34)
        End If
    Case "Empty"
        strTemp = "null"
    Case Else
        strTemp = varValue
    End Select
    strTemp = IIf(Not blnFirst, ",", "") & Chr(34) & LCase(strKey) & Chr(34) & ":" & strTemp
    If bytType = 1 Then strTemp = Chr(123) & strTemp & Chr(125)
    If bytType = 2 Then strTemp = Chr(91) & strTemp & Chr(93)
    GetNode = strTemp
End Function

Public Function GetJsonListValue(ByVal strListPathNode As String, Optional ByVal strKeyNodes As String, Optional ByVal varNullValue As Variant) As Collection
'���ܣ���ȡJson�е��������ݻ��ӽ�����ݵ�������
'������
'  strList=Json������򸸽������·�����磺output��output.pati_list��output.pati_list[0].baby_list
'  strKeys=��������Ϊ�ؼ��ֵĽ���������Զ����","�ŷָ�����"pati_id,pati_pageid"��ע��ؼ��ֽ������ݲ���������ظ�
'  varNullValue=�������еĽ��ֵΪΪnullʱ�����ص�ת��ֵ
    If InitSvr() Then
        Set GetJsonListValue = gobjService.GetJsonListValue(strListPathNode, strKeyNodes, varNullValue)
    End If
End Function

Public Function GetJsonNodeValue(ByVal strPathNode As String, Optional ByVal varNullValue As Variant) As Variant
'���ܣ���ȡJsonָ������ֵ
'������
'  strElement=��㼰·�����磺output.message��output.pati_list[0].phone_number,output.num_list
'  varNullValue=�����ֵΪΪnullʱ�����ص�ת��ֵ
    If InitSvr() Then
        GetJsonNodeValue = gobjService.GetJsonNodeValue(strPathNode, varNullValue)
    End If
End Function

Public Function InitSvr() As Boolean
'���ܣ���ʼ������ӿڲ���
    If gobjService Is Nothing Then
        On Error Resume Next
        Set gobjService = CreateObject("zlServiceCall.clsServiceCall")
        If Not gobjService.InitService(gcnOracle, gstrDBUser, glngSys) Then
            Set gobjService = Nothing
        End If
        Err.Clear: On Error GoTo 0
    End If
    If gobjService Is Nothing Then
        MsgBox "zlServiceCall.clsServiceCall����ʧ��!", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not gobjService Is Nothing Then InitSvr = True
End Function

Public Function CallService(ByVal strServiceName As String, _
    ByVal strJson_In As String, Optional ByRef strJson_out As String, Optional ByVal strTitle As String, _
    Optional lngModule As Long, Optional blnShowErrMsg As Boolean = True, Optional ByVal strAskDate As String, _
    Optional varExpend As String, Optional lngSys As Long, Optional blnReadServiceErr As Boolean) As Boolean
'���ܣ����÷���
'���˵���� zlServiceCall.clsServiceCall.CallService �ӿ�
    If InitSvr() Then
        If lngModule = 0 Then lngModule = glngModul
        If Not gobjService.CallService(strServiceName, strJson_In, strJson_out, strTitle, lngModule, blnShowErrMsg, strAskDate, varExpend, lngSys, blnReadServiceErr) Then Exit Function
        If Not blnShowErrMsg Then
            If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
                varExpend = gobjService.GetJsonNodeValue("output.message")
                CallService = False: Exit Function
            End If
        End If
        CallService = True
    End If
End Function

Public Function GetJsonSurety(ByVal bytFunc As Byte, ByVal lngPatiID As Long, ByVal lngPageID As Long, Optional ByVal strGuarantor As String, _
   Optional ByVal dblAmount As Double, Optional ByVal bytType As Byte, Optional ByVal strReason As String, Optional ByVal strDueTime As String, _
   Optional ByVal strCreateTime As String) As String
'����:��ȡ����JSON��
'����:
'      bytFunc          ����ID 1-����;2-����;3-ɾ��
'      lngPatiId        ����id
'      lngPageId        ��ҳID
'      strGuarantor     ������
'      dblAmount        ������
'      bytType          ��������
'      strReason        ����ԭ��
'      strDueTime       ����ʱ��   ��ʽ�� "yyyy-MM-dd HH:mm:ss"
'      strCreateTime    �Ǽ�ʱ��   ���»�ɾ��ʱ����
    Dim strIn As String
    
    strIn = GetNode("func_id", bytFunc, True)
    strIn = strIn & GetNode("pati_id", lngPatiID)
    strIn = strIn & GetNode("pati_pageid", lngPageID)
    If bytFunc <> 3 Then
        strIn = strIn & GetNode("guarantor", strGuarantor)
        strIn = strIn & GetNode("garnt_amount", dblAmount)
        strIn = strIn & GetNode("garnt_prop", bytType)
        strIn = strIn & GetNode("garnt_reason", strReason)
        If strDueTime <> "" Then
            strIn = strIn & GetNode("due_time", strDueTime)
        End If
    End If

    strIn = strIn & GetNode("operator_code", UserInfo.���)
    strIn = strIn & GetNode("operator_name", UserInfo.����)

    If bytFunc <> 1 Then
        strIn = strIn & GetNode("create_time", strCreateTime)
    End If
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    GetJsonSurety = strIn
End Function


'---------------------------------------------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------------------------------------------
Public Function ZL_PatiSvr_GetPatiId(ByRef colPati As Collection, ByVal strFindName As String, ByVal strFindValue As String, Optional ByVal lngPatiID As Long) As Boolean
'����:��ȡ����ID
'����:
'       _pati_id ����ID
'       _pati_pageid ��ҳID
    Dim strIn As String
    Dim strOut As String
    
    strIn = GetNode("find_name", strFindName, True)
    strIn = strIn & GetNode("find_text", strFindValue)
    If lngPatiID > 0 Then strIn = strIn & GetNode("pati_id", lngPatiID)
    strIn = GetNode("other_cons_find", "{" & strIn & "}", True, 1)
    strIn = GetNode("input", strIn, True, 1)
    If Not CallService("Zl_Patisvr_Getpatiid", strIn, strOut, "��ȡ����ID", P������Ժ����) Then Exit Function
    Set colPati = GetJsonListValue("output.pati_list", "pati_id")
    If colPati.Count = 0 Then
        colPati.Add 0, "_pati_id"
    Else
        Set colPati = colPati(1)
    End If
    ZL_PatiSvr_GetPatiId = True
End Function

Public Function Zl_PatiSvr_GetNextId(ByVal strTableName As String, ByVal strColName As String) As Long
'  -----------------------------------------
'  --���ܣ���ȡָ��������Ӧ������(���淶������������Ϊ��������_id��)����һ��ֵ
'  --��Σ�Json_In:��ʽ
'  --input
'  --  table_name    C  1 ����
'  --  col_name      C  1 �ֶ���  �������Ʋ�һ����ID�������¼ID
'  -- ����:
'  --  output
'  --  next_id      N   1  ����
'  -------------------------------------------
    Dim strIn As String
    Dim strOut As String
    
    strIn = GetNode("table_name", strTableName, True)
    strIn = strIn & GetNode("col_name", strColName)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    
    If Not CallService("Zl_PatiSvr_GetNextId", strIn, strOut, "��ȡָ��������Ӧ������", P������Ժ����) Then Exit Function
    Zl_PatiSvr_GetNextId = GetJsonNodeValue("output.next_id")
End Function

Public Function zl_PatiSvr_GetNextNo(ByVal lngItemNum As Long, Optional ByVal lngDeptId As Long) As String
'  ---------------------------------------------------------------------------
'  --���ܣ����ܣ������ض���������µĺ���
'  lngItemNum:1  �������ID ;2  ����סԺ��;3  ���������
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    item_num            N   1   ��Ŀ���
'  --    dept_id             N   0   ����ID
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    next_no             C   1   ��һ������
'  ---------------------------------------------------------------------------
    Dim strIn As String
    Dim strOut As String
    
    strIn = GetNode("item_num", lngItemNum, True)
    strIn = strIn & GetNode("dept_id", lngDeptId)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    
    If Not CallService("zl_PatiSvr_GetNextNo", strIn, strOut, "�����ض���������µĺ���", P������Ժ����) Then Exit Function
    zl_PatiSvr_GetNextNo = GetJsonNodeValue("output.next_no")
End Function

Public Function zl_PatiSvr_GetPatiInfo(ByVal lngPatiID As Long, Optional ByVal bytQueryType As Byte, Optional ByVal bytCard As Byte, _
    Optional ByVal bytFamily As Byte, Optional ByVal bytDrug As Byte, Optional ByVal bytImmune As Byte, _
    Optional ByVal strPatiIDs As String, Optional ByVal strPatiName As String, Optional strOutNum As String, _
    Optional strIdCard As String, Optional strContactId As String, Optional dblCardTypeId As Double, _
    Optional strMedcCardName As String, Optional strCardNO As String, Optional strQRcode As String, _
    Optional strICCardNo As String, Optional strVisitCard As String, Optional strInsuranceNum As String, _
    Optional intStatu As Integer = -1, Optional strPhoneNumber As String, Optional strBed As String) As Boolean
'����:��ȡ������Ϣ
'����:
'     lngPatiID           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
'     bytQueryType        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;2-����
'     bytCard           N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
'     bytFamily         N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
'     bytDrug           N 1 �Ƿ��������ҩ��:1-������0-������
'     bytImmune         N 1 �Ƿ����������:1-����;0-������
'       strPatiIds          C   ����IDs:����ö���
'       strPatiName         C   ����:���Դ�%�ֺű������ƥ��
'       dblOutNum           N   �����
'       strIdCard           C   ���֤��
'       strContactId        C   ��ϵ�����֤��
'       dblCardTypeId       N   ҽ�ƿ����ID
'       strMedcCardName     C   ҽ�ƿ�����
'       strCardNo           C   ����
'       strQRcode           C   ��ά��
'       strICCardNo         C   Ic����
'       strVisitCard        C   ���￨��
'       strInsuranceNum     C   ҽ����
'       intStatu            C   ��ѯסԺ״̬:0-������;1-��Ժ ;2-���Ｐ��Ժ
'       strPhoneNumber      C   �ֻ���
'       strBed              C   ��ǰ����
    Dim strIn As String
    Dim strOut As String
    Dim strTemp As String

    On Error GoTo ErrH

    strIn = GetNode("pati_id", lngPatiID, True)
    If bytQueryType > 0 Then strIn = strIn & GetNode("query_type", bytQueryType)
    If bytCard > 0 Then strIn = strIn & GetNode("query_card", bytCard)
    If bytFamily > 0 Then strIn = strIn & GetNode("query_family", bytFamily)
    If bytDrug > 0 Then strIn = strIn & GetNode("query_drug", bytDrug)
    If bytImmune > 0 Then strIn = strIn & GetNode("query_immune", bytImmune)
    If bytFamily > 0 Then strIn = strIn & GetNode("query_family", bytFamily)
    If lngPatiID = 0 Then
        If strPatiIDs <> "" Then strTemp = strTemp & GetNode("pati_ids", strPatiIDs)
        If strPatiName <> "" Then strTemp = strTemp & GetNode("pati_name", strPatiName)
        If strOutNum <> "" Then strTemp = strTemp & GetNode("outpatient_num", strOutNum)
        If strIdCard <> "" Then strTemp = strTemp & GetNode("pati_idcard", strIdCard)
    
        If strContactId <> "" Then strTemp = strTemp & GetNode("contacts_idcard", strContactId)
        If dblCardTypeId > 0 Then strTemp = strTemp & GetNode("cardtype_id", dblCardTypeId)
        If strMedcCardName <> "" Then strTemp = strTemp & GetNode("medc_card_name", strMedcCardName)
        If strCardNO <> "" Then strTemp = strTemp & GetNode("card_no", strCardNO)
        If strQRcode <> "" Then strTemp = strTemp & GetNode("qrcode", strQRcode)
        If strICCardNo <> "" Then strTemp = strTemp & GetNode("iccard_no", strICCardNo)
    
        If strVisitCard <> "" Then strTemp = strTemp & GetNode("visit_card", strVisitCard)
        If strInsuranceNum <> "" Then strTemp = strTemp & GetNode("insurance_num", strInsuranceNum)
        If intStatu >= 0 Then strTemp = strTemp & GetNode("qrspt_statu", intStatu)
        If strPhoneNumber <> "" Then strTemp = strTemp & GetNode("phone_number", strPhoneNumber)
        If strBed <> "" Then strTemp = strTemp & GetNode("pati_bed", strBed)
    
        If strTemp <> "" Then
            strIn = strIn & GetNode("query_cons_list", "{" & Mid(strTemp, 2) & "}")
        End If
    End If
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    zl_PatiSvr_GetPatiInfo = CallService("Zl_Patisvr_Getpatiinfo", strIn, strOut, "��ȡ���˻�����Ϣ", P������Ժ����, False, , , , True)

    Exit Function

ErrH:
    MsgBox "��zl9InPatient.mdlService.zl_PatiSvr_GetPatiInfo�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, gstrSysName

End Function

Public Function Zl_Patisvr_Getpatiaddrssinfo(ByVal lngPatiID As Long, ByVal lngPageID As Long, Optional ByVal bytType As Byte) As Boolean
'  --���ܣ����ݲ���ID,��ȡ���˵ĵ�ַ��Ϣ
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    pati_id              N   1  ����ID
'  --    pati_pageid          N      ��ҳid
'  --    addr_type            N      ��ַ���:1-�����أ�2-����,3-��סַ,4-���ڵ�ַ,5-��ϵ�˵�ַ��6-��λ��ַ��Ϊ0ʱ��ʾ��ѯ�������͵ĵ�ַ��Ϣ
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
'  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    addr_list[]          C       ��ַ�б���Ϣ
'  --      pat_addr_type      C   1   ��ַ���
'  --      pat_addr_state     C   1   ��ַ_ʡ
'  --      pat_addr_city      C   1   ��ַ_��
'  --      pat_addr_county    C   1   ��ַ_��
'  --      pat_addr_township  C   1   ��ַ_��
'  --      pat_addr_other     C   1   ��ַ_����
'  --      pat_region_code    C   1   ��������
    Dim strIn As String
    Dim strOut As String
        
    strIn = GetNode("pati_id", lngPatiID, True)
    strIn = strIn & GetNode("pati_pageid", lngPageID)
    strIn = strIn & GetNode("addr_type", bytType)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    Zl_Patisvr_Getpatiaddrssinfo = CallService("Zl_Patisvr_Getpatiaddrssinfo", strIn, strOut, "��ȡ�ṹ����ַ", P������Ժ����)
End Function
 

  
'----------------------------------------------------------------------------------------------------------------------------
'------������ط���
'----------------------------------------------------------------------------------------------------------------------------
 
Public Function Zl_Exsesvr_CheckPatiChangeUndo(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal strUndoType As String, _
    ByVal strBeginTime As String, ByVal lngFeeItemID As Long) As Boolean
'����:�������˱䶯��¼ǰ���
'������
'   strUndoType-������ʽ
'   strBeginTime-��ʼʱ��
'   lngFeeItemID-������ĿID
    Dim strIn As String
    Dim strOut As String
    
    strIn = GetNode("pati_id", lngPatiID, True)
    strIn = strIn & GetNode("pati_pageid", lngPageID)
    strIn = strIn & GetNode("undo_type", strUndoType)
    strIn = strIn & GetNode("create_time", strBeginTime)
    strIn = strIn & GetNode("fee_item_id", lngFeeItemID)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    Zl_Exsesvr_CheckPatiChangeUndo = CallService("Zl_Exsesvr_Checkpatichangeundo", strIn, strOut, "�������˱䶯���", P������Ժ����)
End Function

Public Function zl_ExseSvr_GetInsuranceDisease(ByVal lng����ID As Long) As Boolean
'����:��ȡ���ղ�����
    Dim strIn As String
    Dim strOut As String
    
    strIn = GetNode("si_dz_id", lng����ID, True)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    zl_ExseSvr_GetInsuranceDisease = CallService("Zl_Exsesvr_Getinsurancedisease", strIn, strOut, "��ȡ���ղ���", P������Ժ����)
End Function

Public Function Zl_Exsesvr_Getpatisurplusinfo(ByVal varPatiId As Variant, ByVal bytFunc As Byte, _
    ByRef colFee As Collection, Optional ByVal lngPageID As Long, Optional ByVal lngModel As Long) As Boolean
'����:
'    bytFunc= 1-��ȡ���˷������;2-��ȡ����ĳһ��סԺ��δ�����;3-������ȡ���˷������
'���Σ�
'    colFee

    Dim strIn As String
    Dim strOut As String
    Dim colList As Collection
    '  output
    '    code              N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message           C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    infee_surplus     N 1 סԺ�������
    '    surplus_list[]    C 1 ����б�
    '      pati_Id         N   ����ID
    '      outdpst_surplus N 1 ����Ԥ�����
    '      indpst_surplus  N 1 סԺԤ�����
    '      outfee_surplus  N 1 ����������
    '      infee_surplus   N 1 סԺ�������
    ' ��ȡ�����Ƿ����
    On Error GoTo ErrH
    
    Set colFee = New Collection
    If bytFunc = 2 Then
        strIn = GetNode("pati_id", CLng(varPatiId), True)
        strIn = strIn & GetNode("pati_pageid", lngPageID)
    Else
        strIn = GetNode("pati_ids", CStr(varPatiId), True)
    End If
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    If Not CallService("Zl_Exsesvr_Getpatisurplusinfo", strIn, strOut, "��ȡ���˷���", lngModel) Then Exit Function
    If bytFunc = 1 Then
        Set colList = GetJsonListValue("output.surplus_list[0]")
        If Not colList Is Nothing Then
            If colList.Count > 0 Then
                colFee.Add colList("_outfee_surplus"), "����������"
                colFee.Add colList("_infee_surplus"), "סԺ�������"
                colFee.Add colList("_indpst_surplus"), "סԺԤ�����"
            End If
        End If
        If colFee.Count = 0 Then
            colFee.Add 0, "����������"
            colFee.Add 0, "סԺ�������"
            colFee.Add 0, "סԺԤ�����"
        End If
    ElseIf bytFunc = 2 Then
        colFee.Add Val(GetJsonNodeValue("output.infee_surplus") & ""), "סԺ�������"
    Else
        Set colFee = GetJsonListValue("output.surplus_list", "pati_id")
    End If
    Zl_Exsesvr_Getpatisurplusinfo = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Zl_Exsesvr_GetReceiveInvoice(ByVal bytKind As Long, ByRef colList As Collection, Optional ByVal bytFunc As Byte, _
    Optional ByVal strReceiver As String, Optional ByVal bytUseMode As Byte, Optional ByVal strUseType As String, _
    Optional ByVal strRecvIds As String) As Boolean
'����:
'   bytFunc= 0-��ȡƱ��������Ϣ 1-��ȡ��ȡָ��Ʊ�ֵĹ���Ʊ������
'   strReceiver-������
'   strRecvIds-����ids:Ʊ������id,����ö���
'  ---------------------------------------------------------------------------
'  --����:��ȡƱ��������Ϣ
'  --��Σ�Json_In:��ʽ
'  --    input
'  --      oper_fun  N 1 0-��ȡƱ��������Ϣ 1-��ȡ��ȡָ��Ʊ�ֵĹ���Ʊ������
'  --      recv_ids C 1 ����ids:Ʊ������id,����ö���
'  --      inv_type  N 1 Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
'  --      use_mode  N 1 ʹ�÷�ʽ:1-���ã���Ʊ�ݽ����������Լ�ʹ�ã�2-���ã���Ʊ���ɶ����Ա��ͬʹ��
'  --      use_type C 1 Ʊ��ʹ�����:1,4: Ʊ��ʹ�����.����;2Ԥ��:1-����Ԥ��;2-סԺԤ��;5:�洢����ҽ�ƿ����.ID
'  --      recvtr  C 1 ������
'  --      min_nums  N 1 ��Ʊ��������
'  --      nodeno  C  1  վ��
'  --����: Json_Out,��ʽ����
'  --    output
'  --    code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message C 1 "Ӧ����Ϣ�� �ɹ�ʱ���سɹ���Ϣ,ʧ��ʱ���ؾ���Ĵ�����Ϣ"
'  --    item_list C
'  --      recv_id N 1 ����ID
'  --      use_mode  N 1 ʹ�÷�ʽ:1-���ã���Ʊ�ݽ����������Լ�ʹ�ã�2-���ã���Ʊ���ɶ����Ա��ͬʹ��
'  --      use_type C 1 Ʊ��ʹ�����:1,4: Ʊ��ʹ�����.����;2Ԥ��:1-����Ԥ��;2-סԺԤ��;5:�洢����ҽ�ƿ����.ID
'  --      prefix_text C 1 ǰ׺�ı�
'  --      start_no  C 1 ��ʼ����
'  --      end_no  C 1 ��ֹ����
'  --      inv_no_cur  C 1 ��ǰ����
'  --      surplus_num C 1 ʣ������
'  --      create_time C 1 �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
'  --      use_time  C 1 ʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
'  --      recvtr  C 1 ������
'  --      use_typecode      C 1 ʹ��������
'  --      use_typeid        N 1 ʹ�����id
'
'  ---------------------------------------------------------------------------
    Dim strIn As String
    Dim strOut As String

    On Error GoTo ErrH

    strIn = GetNode("oper_fun", bytFunc, True)
    strIn = strIn & GetNode("inv_type", bytKind)
    If bytFunc = 0 Then
        If strRecvIds <> "" Then strIn = strIn & GetNode("recv_ids", strRecvIds)
        strIn = strIn & GetNode("recvtr", strReceiver)
        strIn = strIn & GetNode("use_mode", bytUseMode)
        strIn = strIn & GetNode("use_type", strUseType)
    End If
    strIn = strIn & GetNode("nodeno", gstrNodeNo)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    If Not CallService("Zl_Exsesvr_GetReceiveInvoice", strIn, strOut, "��ȡƱ��������Ϣ") Then Exit Function
    Set colList = GetJsonListValue("output.item_list")
    Zl_Exsesvr_GetReceiveInvoice = True
    Exit Function
ErrH:
   If ErrCenter() = 1 Then Resume
   Call SaveErrLog
End Function

Public Function zl_PatiSvr_GetPatiInfoEx(ByVal lng����ID As Long, _
    ByVal colOtherFindCons As Collection, ByRef colPatiDatas_Out As Collection, _
    Optional ByVal int��ѯ���� As Integer = 0, _
    Optional ByVal bln�������� As Boolean, _
    Optional ByVal bln��������ҩ�� As Boolean, _
    Optional ByVal bln����������Ϣ As Boolean, _
    Optional ByVal bln��������Ϣ As Boolean, Optional ByVal blnNotShowErrMsg As Boolean = True, _
    Optional ByVal bln�Ƿ�������ʾ As Boolean, _
    Optional ByRef strErrMsg As String, Optional ByVal strKeyNodes As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������ϸ��Ϣ����ӿ�
    '���:colOtherFindCons-������������(array(��ѯ����,��ѯֵ)
    '             ��ѯ����:����IDS,����,�Ա�,�������ڵ�,��query_cons_list[]�б��е���������
    '      int��ѯ����-0-����;1-����+��ϵ��;2-����
    '      strKeyNodes= ָ�����ؼ���������ʽ
    '����:colPatiDatas_Out-���ز�����Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:YWJ
    '����:2019-10-31 18:02:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim colData As Variant
    Dim intReturn As Integer, strJsonTemp As String
    
    
    On Error GoTo errHandle
    
    If Not colOtherFindCons Is Nothing Then
        For i = 1 To colOtherFindCons.Count
            Select Case UCase(colOtherFindCons(i)(0))
            Case "����IDS"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_ids", colOtherFindCons(i)(1), Json_Text)
            Case "����"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_name", colOtherFindCons(i)(1), Json_Text)
            Case "�����"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("outpatient_num", Val(colOtherFindCons(i)(1)), Json_Text, True)
            Case "סԺ��"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("inpatient_num", Val(colOtherFindCons(i)(1)), Json_Text, True)
            Case "���֤��", "�������֤", "���֤"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_idcard", colOtherFindCons(i)(1), Json_Text)
            Case "��ϵ�����֤"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("contacts_idcard", colOtherFindCons(i)(1), Json_Text)
            Case "ҽ����", "ҽ��֤��"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_num", colOtherFindCons(i)(1), Json_Text)
            Case "ҽ�ƿ����ID", "�����ID"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(colOtherFindCons(i)(1)), Json_num, True)
            Case "����"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("card_no", colOtherFindCons(i)(1), Json_Text)
            Case "��ά��"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("qrcode", colOtherFindCons(i)(1), Json_Text)
            Case "IC����", "IC", "IC��"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("iccard_no", colOtherFindCons(i)(1), Json_Text)
            Case "��ѯסԺ״̬", "סԺ״̬"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("qrspt_statu", Val(colOtherFindCons(i)(1)), Json_num, True)
            Case "�ֻ���", "�ֻ�"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("phone_number", colOtherFindCons(i)(1), Json_Text)
            Case "���￨��", "���￨"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("vcard_no", colOtherFindCons(i)(1), Json_Text)
            Case "��������"
                 strJsonTemp = strJsonTemp & "," & GetJsonNodeString("search_days", colOtherFindCons(i)(1), Json_num, True)
            Case Else
                strErrMsg = "Ŀǰ�ݲ���֧�ְ����Ϊ��" & UCase(colOtherFindCons(i)(0)) & "�������Ҳ��ˣ�"
                If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
            End Select
        Next
        If strJsonTemp <> "" Then strJsonTemp = Mid(strJsonTemp, 2)
    End If
    If lng����ID = 0 And strJsonTemp = "" Then
        strErrMsg = "����Ч�Ĳ�ѯ���������飡"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '��������
    'zl_PatiSvr_GetPatiInfo
    'input
    '    pati_id N   1   ����id������ID<>0ʱ����ѯ�б��е�������Ч
    '    query_type  N   1   ��ѯ����:�磺0-����;1-����+��ϵ��;2-����
    '    query_card  N   1   �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
    '    query_family    N   1   �Ƿ��������:1-����������Ϣ��0-������������Ϣ
    '    query_drug  N   1   �Ƿ��������ҩ��:1-������0-������
    '    query_immune    N   1   �Ƿ����������Ϣ:1-����;0-������
    '    query_cons_list[]   C   1   ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��
    '        pati_ids    C       ����IDs:����ö���
    '        pati_name   C       ����:���Դ�%�ֺű������ƥ��
    '        outpatient_num  N       �����
    '        pati_idcard C       ���֤��
    '        contacts_idcard C       ��ϵ�����֤��
    '        cardtype_id N       ҽ�ƿ����ID
    '        card_no C       ����
    '        qrcode  C       ��ά��
    '        iccard_no   C       Ic����
    '        insurance_num   C       ҽ����
    '        qrspt_statu C       ��ѯסԺ״̬:0-������;1-��Ժ ;2-���Ｐ��Ժ
    '        phone_number    C       �ֻ���
       
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_type", int��ѯ����, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_family", IIf(bln��������, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_drug", IIf(bln��������ҩ��, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_immune", IIf(bln����������Ϣ, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("query_card", IIf(bln��������Ϣ, 1, 0), Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("passshowcard", IIf(bln�Ƿ�������ʾ, 1, 0), Json_num, True)
    strJson = strJson & "," & GetNodeString("query_cons_list") & ":{" & strJsonTemp & "}"
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetPatiInfo"
    If CallService(strServiceName, strJson, , "", , False) = False Then Exit Function
    
    
    '����            json    ����    ����+��ϵ�� ����
    'output
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�  ��  ��  ��
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ ��  ��  ��
    '    pati_list[]         ������Ϣ�б�    ��  ��  ��
    '    pati_id N   1   ����id  ��  ��  ��
    '    pati_pageid N   1   ��ҳid��������Ϣ.��ҳID ��  ��  ��
    '    pati_name   C   1   ����    ��  ��  ��
    '    pati_sex    C   1   �Ա�    ��  ��  ��
    '    pati_age    C   1   ����    ��  ��  ��
    '    pati_birthdate  C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss ��  ��  ��
    '    fee_category    C   1   �ѱ�    ��  ��  ��
    '    outpatient_num  C   1   �����  ��  ��  ��
    '    inpatient_num   C   1   סԺ��
    '    mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����    ��  ��  ��
    '    mdlpay_mode_code    C   1   ҽ�Ƹ��ʽ����    ��  ��  ��
    '    pati_nation C   1   ����    ��  ��
    '    insurance_num   C   1   ҽ����  ��  ��  ��
    '    pati_idcard C   1   ���֤��    ��  ��  ��
    '    vcard_no    C   1   ���￨��            ��
    '    iccard_no   C   1   Ic����          ��
    '    health_num  C   1   ������          ��
    '    pati_education  C   1   ѧ��            ��
    '    ocpt_name   C   1   ְҵ            ��
    '    pati_identity   C   1   ���            ��
    '    ntvplc_name C   1   ����            ��
    '    country_name    C   1   ����            ��
    '    pati_marital_cstatus    C   1   ����״��            ��
    '    pat_home_addr   C   1   ��ͥ��ַ    ��  ��  ��
    '    pat_home_phno   C   1   ��ͥ�绰    ��  ��  ��
    '    pat_home_postcode   C   1   ��ͥ��ַ�ʱ�            ��
    '    pati_area   C   1   ����            ��
    '    pati_birthplace C   1   �����ص�    ��  ��  ��
    '    pat_hous_addr   C   1   ���ڵ�ַ            ��
    '    pat_hous_postcode   C   1   ���ڵ�ַ�ʱ�            ��
    '    emp_name    C   1   ������λ����            ��
    '    emp_phno    C   1   ��λ�绰            ��
    '    emp_postcode    C   1   ��λ�ʱ�            ��
    '    emp_bank_name   C   1   ��λ������          ��
    '    ctt_unit_id N   1   ��ͬ��λID          ��
    '    phone_number    C   1   �ֻ���  ��  ��  ��
    '    pati_bed    C   1   ��ǰ����    ��  ��  ��
    '    pati_type   C   1   ��������(��ͨ��ҽ��������)          ��
    '    balance_mode    N   1   ����ģʽ(0-�շѣ�1-����)            ��
    '    insurance_type  C   1   ����    ��  ��  ��
    '    pati_wardarea_id    N   1   ��ǰ����id          ��
    '    pati_wardarea_name  C   1   ��ǰ��������            ��
    '    pati_dept_id    N   1   ��ǰ����id          ��
    '    pati_dept_name  C   1   ��ǰ��������            ��
    '    adta_time   C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss          ��
    '    adtd_time   C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss          ��
    '    contacts_name   C   1   ��ϵ������      ��  ��
    '    contacts_relation   C   1   ��ϵ�˹�ϵ      ��  ��
    '    contacts_idcard C   1   ��ϵ�����֤��      ��  ��
    '    contacts_addr   C   1   ��ϵ�˵�ַ      ��  ��
    '    contacts_phno   C   1   ��ϵ�˵绰      ��  ��
    '    pat_grdn_name   C   1   �໤��          ��
    '    cert_no_other   C   1   ����֤��            ��
    '    is_inhspt   C   1   �Ƿ���Ժ:1-��Ժ ;0-����Ժ   ��  ��  ��
    '    pati_show_color N   1   ������ʾ��ɫ            ��
    '    visit_room  C   1   ��������            ��
    '    visit_statu N   1   ����״̬            ��
    '    visit_time  C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss          ��
    '    create_time C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss          ��
    '    family_list[]   C   1   ������Ա:���˼���() query_family=1����
    '        family_id   N   1   ����id  query_family=1
    '        family_relation C   1   ��ϵ
    '    drug_list[] C   1   ����ҩ���б�    query_drug=1ʱ����
    '        pat_algc_cadn_id    N   1   ����ҩƷID
    '        pat_algc_cadn   C   1   ����ҩ������
    '        allergy_info    C   1   ��ÿҩ�ﷴӦ
    '    immune_list[]   C   1   ���������б�    query_immune=1ʱ����
    '        vaccinate_time  C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '        vaccinate_name  C   1   ��������
    '    card_list[] C   1   ����ҽ�ƿ���Ϣ�б�(��������д����˿����ID�ģ��򷵻ظÿ����Ŀ���Ϣ)  query_card=1ʱ����
    '        cardtype_id N   1   ҽ�ƿ����ID
    '        card_no C   1   ����
    '        card_pwd    C   1   ����
    intReturn = gobjService.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg = gobjService.GetJsonNodeValue("output.message")
        If strErrMsg = "" Then
            strErrMsg = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set colData = gobjService.GetJsonListValue("output.pati_list", strKeyNodes)
    
    If colData Is Nothing Then
        strErrMsg = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If colData.Count = 0 Then
        strErrMsg = "δ�ҵ����������Ĳ�����Ϣ�����飡"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set colPatiDatas_Out = colData
    zl_PatiSvr_GetPatiInfoEx = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Patisvr_GetPatiExtendInfo(ByVal lngPatiID As Long, colPatiEx As Collection, Optional ByVal strInfoName As String, Optional ByVal lngVisitID As Long) As Boolean
'  ---------------------------------------------------------------------------
'  --���ܣ���ȡ������Ϣ�ӱ�
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    pati_id             N 1 ����id
'  --    info_names          C 1 ��Ϣ��������ö���
'  --    visit_id            N  ����id
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code                N 1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message             C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    slave_list[]        C     ������Ϣ�ӱ��б�
'  --     info_name          C 1   ��Ϣ��
'  --     info_value         N 1   ��Ϣֵ
'  --     visit_id           N 1   ����id
'  ---------------------------------------------------------------------------
    Dim strIn As String
    Dim strOut As String
    On Error GoTo ErrH

    strIn = GetNode("pati_id", lngPatiID, True)
    strIn = strIn & GetNode("info_names", strInfoName)
    strIn = strIn & GetNode("visit_id", lngVisitID)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    If Not CallService("Zl_Patisvr_GetPatiExtendInfo", strIn, strOut, "��ȡ���˴ӱ���Ϣ") Then Exit Function
    Set colPatiEx = GetJsonListValue("output.slave_list", "info_name")
    Zl_Patisvr_GetPatiExtendInfo = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Zl_Patisvr_GetpatAllergicDrugs(ByVal lngPatiID As Long, colDrug As Collection) As Boolean
'---------------------------------------------------------------------------
'--����:��ȡ������Ϣ�Ĺ���ҩ����Ϣ
'--��Σ�Json_In:��ʽ
'--  input
'--    pati_id             N   1 ����id
'--����: Json_Out,��ʽ����
'--  output
'--    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'--    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'--    drug_list[]         C       ����ҩ���б�
'--      medicinal_id      N   1   ����ҩƷID
'--      medicinal_name    C   1   ����ҩ������
'--      allergy_info      C   1   ��ÿҩ�ﷴӦ
'---------------------------------------------------------------------------
    Dim strIn As String
    Dim strOut As String
    On Error GoTo ErrH

    strIn = GetNode("pati_id", lngPatiID, True)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    If Not CallService("Zl_Patisvr_GetpatAllergicDrugs", strIn, strOut, "��ȡ���˹���ҩ��") Then Exit Function
    Set colDrug = GetJsonListValue("drug_list")
    Zl_Patisvr_GetpatAllergicDrugs = True
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Zl_Patisvr_GetPatImmuneInfo(ByVal lngPatiID As Long, colDrug As Collection) As Boolean
'  ---------------------------------------------------------------------------
'  --����:��ȡ������Ϣ��������Ϣ
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    pati_id           N   1 ����id
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    immune_list[]     C       ���������б�
'  --      vaccinate_time    C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --      vaccinate_name    C   1   ��������
'  ---------------------------------------------------------------------------
    Dim strIn As String
    Dim strOut As String
    On Error GoTo ErrH

    strIn = GetNode("pati_id", lngPatiID, True)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    If Not CallService("Zl_Patisvr_GetPatImmuneInfo", strIn, strOut, "��ȡ����������Ϣ") Then Exit Function
    Set colDrug = GetJsonListValue("immune_list")
    Zl_Patisvr_GetPatImmuneInfo = True
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetColVal(ByVal colData As Collection, ByVal strKey As String, Optional ByVal strType As String, Optional ByVal strDef As String, Optional ByRef lngExist As Long) As String
'����:ͨ���Ϲؼ��ֻ�ȡ���ϵ�ֵ,������������,���ֻ��ַ�
'��Σ�strType  N/n  ��ʾ�������ͣ�c��ʾ�ַ���
'      strDef  ȱʡֵ��������ʱ�����ֵΪȱʡֵ����
'����:lngExist �������Ƿ����������ֵ,0-����,-1������
    Dim strValue As String
    
    On Error GoTo ErrH
    
    If IsNull(colData(strKey)) Then
        strValue = ""
    Else
        strValue = colData(strKey)
    End If
     
    If UCase(strType) = "N" Then
        strValue = Val(strValue)
    End If
    
    GetColVal = strValue
    
    Exit Function
ErrH:
    Err.Clear
    lngExist = -1
    '���Ϸ��ʲ�������ʾ��������
    If strDef <> "" Then
        strValue = strDef
    Else
        If UCase(strType) = "N" Then
            strValue = 0
        Else
            strValue = ""
        End If
    End If
    GetColVal = strValue
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
            strJson = strJson & ":" & IIf(Mid(strValue, 1, 1) = ".", "0", "") & strValue
        End If
    End If
    GetJsonNodeString = strJson
End Function

Public Function GetNodeString(ByVal strNodeName As String) As String
    GetNodeString = Chr(34) & strNodeName & Chr(34)
End Function

Public Function GetMapPait(ByVal strName As String) As String
    Dim strValue As String
    '--    pati_id             N   1   ����id
    '--    pati_pageid         N   1   ��ҳid��������Ϣ.��ҳID
    '--    pati_name           C   1   ����
    '--    pati_sex            C   1   �Ա�
    '--    pati_age            C   1   ����
    '--    pati_birthdate      C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss
    '--    fee_category        C   1   �ѱ�
    '--    outpatient_num      C   1   �����
    '--    inpatient_num       C   1   סԺ��
    '--    mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����
    '--    mdlpay_mode_code    C   1   ҽ�Ƹ��ʽ����
    '--    pati_nation         C   1   ����
    '--    insurance_num       C   1   ҽ����
    '--    pati_idcard         C   1   ���֤��
    '--    vcard_no            C   1   ���￨��
    '--    iccard_no           C   1   Ic����
    '--    health_num          C   1   ������
    '--    inp_times           N   1   סԺ����
    '--    pati_education      C   1   ѧ��
    '--    ocpt_name           C   1   ְҵ
    '--    pati_identity       C   1   ���
    '--    ntvplc_name         C   1   ����
    '--    country_name        C   1   ����
    '--    pati_marital_cstatus    C   1   ����״��
    '--    pat_home_addr           C   1   ��ͥ��ַ
    '--    pat_home_phno           C   1   ��ͥ�绰
    '--    pat_home_postcode   C   1   ��ͥ��ַ�ʱ�
    '--    pati_area           C   1   ����
    '--    pati_birthplace     C   1   �����ص�
    '--    pat_hous_addr       C   1   ���ڵ�ַ
    '--    pat_hous_postcode   C   1   ���ڵ�ַ�ʱ�
    '--    emp_name            C   1   ������λ����
    '--    emp_phno            C   1   ��λ�绰
    '--    emp_postcode        C   1   ��λ�ʱ�
    '--    emp_bank_name       C   1   ��λ������
    '--    emp_bank_accnum     C   1   ��λ�ʺ�
    '--    emp_addr             C   1   ��λ��ַ
    '--    ctt_unit_id         N   1   ��ͬ��λID
    '--    phone_number        C   1   �ֻ���
    '--    pati_bed            C   1   ��ǰ����
    '--    pati_type           C   1   ��������(��ͨ��ҽ��������)
    '--    insurance_type      C   1   ����
    '--    insurance_name      C   1   ��������
    '--    pati_wardarea_id    N   1   ��ǰ����id
    '--    pati_wardarea_name  C   1   ��ǰ��������
    '--    pati_dept_id        N   1   ��ǰ����id
    '--    pati_dept_name      C   1   ��ǰ��������
    '--    adta_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
    '--    adtd_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
    '--    contacts_name       C   1   ��ϵ������
    '--    contacts_relation   C   1   ��ϵ�˹�ϵ
    '--    contacts_idcard     C   1   ��ϵ�����֤��
    '--    contacts_addr       C   1   ��ϵ�˵�ַ
    '--    contacts_phno       C   1   ��ϵ�˵绰
    '--    pat_grdn_name       C   1   �໤��
    '--    cert_no_other       C   1   ����֤��
    '--    is_inhspt            C   1   �Ƿ���Ժ:1-��Ժ ;0-����Ժ
    '--    pati_show_color      N   1   ������ʾ��ɫ
    '--    visit_room           C   1   ��������
    '--    visit_statu          N   1   ����״̬
    '--    visit_time           C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '--    create_time          C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
    '--    pati_email           C   1   email
    '--    pati_qq              C   1   qq
    '--    card_captcha         C   1  ����֤��
    '--    insurance_pwd        C       ҽ������
    Select Case UCase(strName)
    
    Case UCase("����ID")
        strValue = "_pati_id"
    Case UCase("��ҳID")
        strValue = "_pati_pageid"
    Case "����"
        strValue = "_pati_name"
    Case "�Ա�"
        strValue = "_pati_sex"
    Case "����"
        strValue = "_pati_age"
    Case "��������"
        strValue = "_pati_birthdate"
    Case "�ѱ�"
        strValue = "_fee_category"
    Case "�����"
        strValue = "_outpatient_num"
    Case "סԺ��"
        strValue = "_inpatient_num"
    Case "ҽ�Ƹ��ʽ����"
        strValue = "_mdlpay_mode_name"
    Case "ҽ�Ƹ��ʽ����"
        strValue = "_mdlpay_mode_code"
    Case "����"
        strValue = "_pati_nation"
    Case "ҽ����"
        strValue = "_insurance_num"
    Case "���֤��"
        strValue = "_pati_idcard"
    Case "���￨��", "���￨"
        strValue = "_vcard_no"
    Case UCase("IC����")
        strValue = "_iccard_no"
    Case "������"
        strValue = "_health_num"
    Case "סԺ����"
        strValue = "_inp_times"
    Case "ѧ��"
        strValue = "_pati_education"
    Case "ְҵ"
        strValue = "_ocpt_name"
    Case "���"
        strValue = "_pati_identity"
    Case "����"
        strValue = "_ntvplc_name"
    Case "����"
        strValue = "_country_name"
    Case "����״��"
        strValue = "_pati_marital_cstatus"
    Case "��ͥ��ַ"
        strValue = "_pat_home_addr"
    Case "��ͥ�绰"
        strValue = "_pat_home_phno"
    Case "��ͥ��ַ�ʱ�"
        strValue = "_pat_home_postcode"
    Case "����"
        strValue = "_pati_area"
    Case "�����ص�"
        strValue = "_pati_birthplace"
    Case "���ڵ�ַ"
        strValue = "_pat_hous_addr"
    Case "���ڵ�ַ�ʱ�"
        strValue = "_pat_hous_postcode"
    Case "������λ����"
        strValue = "_emp_name"
    Case "��λ�绰"
        strValue = "_emp_phno"
    Case "��λ�ʱ�"
        strValue = "_emp_postcode"
    Case "��λ������"
        strValue = "_emp_bank_name"
    Case "��λ�ʺ�"
        strValue = "_emp_bank_accnum"
    Case "��ͬ��λID"
        strValue = "_ctt_unit_id"
    Case "�ֻ���"
        strValue = "_phone_number"
    Case "��ǰ����"
        strValue = "_pati_bed"
    Case "��������"
        strValue = "_pati_type"
    Case "����"
        strValue = "_insurance_type"
    Case "��������"
        strValue = "_insurance_name"
    Case "��ǰ����id"
        strValue = "_pati_wardarea_id"
    Case "��ǰ��������"
        strValue = "_pati_wardarea_name"
    Case "��ǰ����id"
        strValue = "_pati_dept_id"
    Case "��ǰ��������"
        strValue = "_pati_dept_name"
    Case "��Ժʱ��"
        strValue = "_adta_time"
    Case "��Ժʱ��"
        strValue = "_adtd_time"
    End Select
    GetMapPait = strValue
End Function

Public Function GetJsonInPatiState(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngWardID As Long, ByVal lngDeptId As Long, _
    ByVal strPatiBed As String, ByVal strInTime As String, ByVal bytInpStatus As Byte, strPatiNum As String) As String
'����:
'���²��˾���״̬
'--input
'--    pati_list[]              ����
'--      pati_id              N 1   ����id
'--      pati_pageid          N 1   ��ҳid
'--      outpatient_num       C 1   �����
'--      inpatient_num        C 1   סԺ��
'--      in_time              C 1   ��Ժʱ��
'--      adtd_time            C 1   ��Ժʱ��
'--      pati_deptid          N 1   ��ǰ����id
'--      wardarea_id          N 1   ��ǰ����id
'--      pati_bed             C 1   ��ǰ����
'--      inp_status           N 1   �Ƿ���Ժ��0/1
'--      inp_times            N 1   סԺ����
'--      inp_times_increment  N 1   =1ʱ-סԺ��������;=-1סԺ�����Լ�
'������Ժ�Ǽ�ʱ������Ϣͬ��ʧ��,��סʱ�ٴθ��¡�����š�סԺ���� �ݲ����Ǹ���
    Dim strJsonIn As String
    
    strJsonIn = GetJsonNodeString("pati_id", lngPatiID, Json_num)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("pati_pageid", lngPageID, Json_num)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("wardarea_id", lngWardID, Json_num)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("pati_deptid", lngDeptId, Json_num)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("pati_bed", strPatiBed, Json_Text)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("in_time", strInTime, Json_Text)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("inp_status", bytInpStatus, Json_num)
    strJsonIn = strJsonIn & "," & GetJsonNodeString("inpatient_num", strPatiNum, Json_Text)
    strJsonIn = GetNode("pati_list", "[{" & strJsonIn & "}]", True, 1)
    GetJsonInPatiState = GetNode("input", strJsonIn, True, 1)
End Function


