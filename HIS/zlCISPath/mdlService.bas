Attribute VB_Name = "mdlService"
Option Explicit

Public gobjService As Object        'HIS��ַ������
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
    OpenJson = gobjService.SetJsonString(strJson)
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
    Set GetJsonListValue = gobjService.GetJsonListValue(strListPathNode, strKeyNodes, varNullValue)
End Function

Public Function GetJsonNodeValue(ByVal strPathNode As String, Optional ByVal varNullValue As Variant) As Variant
'���ܣ���ȡJsonָ������ֵ
'������
'  strElement=��㼰·�����磺output.message��output.pati_list[0].phone_number,output.num_list
'  varNullValue=�����ֵΪΪnullʱ�����ص�ת��ֵ
    GetJsonNodeValue = gobjService.GetJsonNodeValue(strPathNode, varNullValue)
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
    ByVal strJson_In As String, Optional ByRef strJson_out As String, Optional ByVal strTittle As String, _
    Optional lngModule As Long, Optional blnShowErrMsg As Boolean = True, Optional ByVal strAskDate As String, _
    Optional varExpend As String, Optional lngSys As Long, Optional blnReadServiceErr As Boolean) As Boolean
'���ܣ����÷���
'���˵���� zlServiceCall.clsServiceCall.CallService �ӿ�
    If InitSvr() Then
        If Not gobjService.CallService(strServiceName, strJson_In, strJson_out, strTittle, lngModule, blnShowErrMsg, strAskDate, varExpend, lngSys, blnReadServiceErr) Then Exit Function
        If Not blnShowErrMsg Then
            If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
                varExpend = gobjService.GetJsonNodeValue("output.message")
                CallService = False: Exit Function
            End If
        End If
        CallService = True
    End If
End Function

Public Function ZL_PatiSvr_GetPatiId(ByRef colPati As Collection, ByVal strFindName As String, ByVal strFindValue As String) As Boolean
'����:��ȡ����ID
'����:
'       _pati_id ����ID
'       _pati_pageid ��ҳID
    Dim strIn As String
    Dim strOut As String
    Dim strTemp As String
   
    
    strIn = GetNode("find_name", strFindName, True)
    strIn = strIn & GetNode("find_text", strFindValue)
    strIn = GetNode("other_cons_find", "{" & strIn & "}", True, 1)
    strIn = GetNode("input", strIn, True, 1)
    If Not CallService("Zl_Patisvr_Getpatiid", strIn, strOut, "��ȡ����ID", P�ٴ�·��Ӧ��) Then Exit Function
    Set colPati = GetJsonNodeValue("output.pati_list[0]")
End Function

Public Function ZL_PatiSvr_GetPatiInfo(ByVal lngPatiID As Long, Optional ByVal bytQueryType As Byte, Optional ByVal bytCard As Byte, _
    Optional ByVal bytFamily As Byte, Optional ByVal bytDrug As Byte, Optional ByVal bytImmune As Byte, _
    Optional ByVal strPatiIds As String, Optional ByVal strPatiName As String, Optional dblOutNum As Double, _
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
     
    strIn = GetNode("pati_id", lngPatiID, True)
    If bytQueryType > 0 Then strIn = strIn & GetNode("query_type", bytQueryType)
    If bytCard > 0 Then strIn = strIn & GetNode("query_card", bytCard)
    If bytFamily > 0 Then strIn = strIn & GetNode("query_family", bytFamily)
    If bytDrug > 0 Then strIn = strIn & GetNode("query_drug", bytDrug)
    If bytImmune > 0 Then strIn = strIn & GetNode("query_immune", bytImmune)
    If bytFamily > 0 Then strIn = strIn & GetNode("query_family", bytFamily)
    
    If strPatiIds <> "" Then strTemp = strTemp & GetNode("pati_ids", strPatiIds)
    If strPatiName <> "" Then strTemp = strTemp & GetNode("pati_name", strPatiName)
    If dblOutNum > 0 Then strTemp = strTemp & GetNode("outpatient_num", dblOutNum)
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
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    ZL_PatiSvr_GetPatiInfo = CallService("Zl_Patisvr_Getpatiinfo", strIn, strOut, "��ȡ���˻�����Ϣ", P�ٴ�·��Ӧ��)
End Function

Public Function GetMoneyInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByRef dblRemainMoney As Double, ByRef dblPrePayMoney As Double) As Boolean
'���ܣ���ȡָ�����˵�ʣ���
'������
'   dblRemainMoney-ʣ���
'   dblExpectedMoney-Ԥ�����
    Dim strIn As String
    Dim strOut As String
    Dim strTemp As String
    
    On Error GoTo errH
 
    strIn = GetNode("pati_id", lng����ID, True)
    strIn = strIn & GetNode("pati_pageid", lng��ҳID)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
              
    If Not CallService("Zl_Exsesvr_Getremainmoney", strIn, strOut, "�������", P�ٴ�·��Ӧ��) Then Exit Function
    
    dblRemainMoney = Val(GetJsonNodeValue("output.remain_money") & "")
    dblPrePayMoney = Val(GetJsonNodeValue("output.prepay_money") & "")
    dblRemainMoney = -1 * (dblRemainMoney - dblPrePayMoney)
    GetMoneyInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDrugStockBatch(ByVal strDrugIds As String, ByVal strPharmacyIDs As String, Optional colList As Collection) As Boolean
'����:��ȡ��治���ҩƷID
'����:  strDrugIds ҩƷIDs ���ҩƷID�ö��ŷָ�
'����:
'  --��Σ�Json_In:��ʽ
'  --  input
'  --   drug_ids       C   1   ҩƷID�������Ӣ�ĵĶ��ŷָ�
'  --   pharmacy_ids   C   0   �ⷿID�������Ӣ�ĵĶ��ŷָ�;���ַ���,��ѯ���пⷿ
'  --   return_price   N   0   �Ƿ񷵻��ۼۣ�1-���ؼ۸���Ϣ(�ۼ�);0-������
'  --   return_dept    N   0   �����ҷ��ؿ�棺1-�����ҷ��ؿ��;0-��ҩƷ���ؿ��;2-���ؿ�������ҩƷ�Ŀ��
'  --   query_type     N   1   ��ѯ����:�磺0-��ѯ��治����0,1-��ѯ���С�ڵ���0
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code                 N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
'  --    message              C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    item_list
'  --    drug_id              N   1   ҩƷID
'  --    pharmacy_id          N   1   �ⷿID(�����ҷ��ؿ����д���)
'  --    stock                N   1   ��������
'  --    price                N   1   ���ۼ�(���ؼ۸�ʱ���д���)
'  ---------------------------------------------------------------------------
    Dim strJson As String, strJsonOut As String
    Dim i As Long
    Dim strResult As String
    Dim strFilds As String
    Dim strKeyNodes As String
    
     
    Dim colItem As Collection
       
    strJson = GetJsonNodeString("drug_ids", strDrugIds, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_ids", strPharmacyIDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("return_dept", 0, Json_num)
    strJson = strJson & "," & GetJsonNodeString("query_type", 1, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    strKeyNodes = "drug_id"
    'drug_id,pharmacy_id,stock
    If Not CallService("zl_DrugSvr_GetStockBatch", strJson, strJsonOut, App.ProductName, p�ٴ�·������, True, , , , True) Then
        Exit Function
    End If
    Set colList = gobjService.GetJsonListValue("output.item_list", strKeyNodes)
    GetDrugStockBatch = True
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

Public Function GetStrByRS(ByVal rsTemp As ADODB.Recordset, Optional ByVal strFiledName As String = "ID") As String
'����:���ݴ����¼������ָ���ֶ��ö��ŷָ����ַ���
    Dim i As Long
    Dim strResult As String
    
    For i = 1 To rsTemp.RecordCount
        If InStr("," & strResult & ",", "," & rsTemp(strFiledName) & ",") = 0 Then
            strResult = strResult & "," & rsTemp(strFiledName)
        End If
        rsTemp.MoveNext
    Next
    If strResult <> "" Then strResult = Mid(strResult, 2)
    GetStrByRS = strResult
End Function


Public Function GetColVal(ByVal colData As Collection, ByVal strKey As String, Optional ByVal strType As String, Optional ByVal strDef As String, Optional ByRef lngExist As Long) As String
'����:ͨ���Ϲؼ��ֻ�ȡ���ϵ�ֵ,������������,���ֻ��ַ�
'��Σ�strType  N/n  ��ʾ�������ͣ�c��ʾ�ַ���
'      strDef  ȱʡֵ��������ʱ�����ֵΪȱʡֵ����
'����:lngExist �������Ƿ����������ֵ,0-����,-1������
    Dim strValue As String
    
    On Error GoTo errH
    
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
errH:
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

Public Function GetColObj(ByVal colData As Collection, ByVal strKey As String) As Collection
'����:ͨ���Ϲؼ��ֻ�ȡ�����еļ��϶���
    On Error GoTo errH
    Set GetColObj = colData(strKey)
    Exit Function
errH:
    Err.Clear
    Set GetColObj = New Collection
End Function
