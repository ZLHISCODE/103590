Attribute VB_Name = "mdlCISService"
Option Explicit
        
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
    err.Clear
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
    err.Clear
    Set GetColObj = New Collection
End Function

Public Function MergeStr(ByVal strIn1 As String, ByVal strIn2 As String, Optional ByVal strTag As String = ",") As String
'����:�������ַ�����ָ�����ӷ���ƴ�ӵ�һ��
'����:strTag ���ӷ���
    Dim i As Long
    Dim varTmp As Variant
    Dim strTmp As String
    
    On Error GoTo errH
    
    If strIn1 <> "" Or strIn2 <> "" Then
        If strIn1 = "" Then
            MergeStr = strIn2
        ElseIf strIn2 = "" Then
            MergeStr = strIn1
        Else
            '�ж��Ƿ������
            strTmp = strIn1 & strTag & strIn2
            varTmp = Split(strTmp, strTag)
            strTmp = ""
            For i = 0 To UBound(varTmp)
                If varTmp(i) & "" <> "" Then
                    If InStr(strTag & strTmp & strTag, strTag & varTmp(i) & strTag) = 0 Then
                        strTmp = strTmp & strTag & varTmp(i)
                    End If
                End If
            Next
            strTmp = Mid(strTmp, Len(strTag) + 1)
            MergeStr = strTmp
        
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetJsonNum(ByVal strNum As String) As String
'����:jsonƴ��ʱ��С�����Ҫ���⴦��
    Dim dblTmp As Double
    Dim strTmp As String
    dblTmp = Val(strNum)
    strTmp = dblTmp
    If Mid(strTmp, 1, 1) = "." Then strTmp = "0" & strTmp
    GetJsonNum = strTmp
End Function

Public Sub InitSQLSend(rsSQL As ADODB.Recordset)
'����:ҽ�����ʹ����ʼ����¼��
    'SQL��¼��
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "����", adInteger '1-���ü�¼,2-ҽ����¼,3-���ͼ�¼,4-���ϼ�¼
                                          '1-�Ƽ�,2-����,3-ǩ��,4-����,5-���� ����ҽ����
                                          '1-�Ƽ�,2-ǩ��,3-У��,4-����,5-����,6-����,,8-Ӥ��ת��
                                          '����ҽ������sql,����=2
                                          
    rsSQL.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsSQL.Fields.Append "��ĿID", adBigInt '�շ�ϸĿID
    rsSQL.Fields.Append "���", adBigInt '��������
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.Fields.Append "NO", adVarChar, 30, adFldIsNullable '����NO�滻����ʱ����
    rsSQL.Fields.Append "����ID", adVarChar, 300 '����ID ҩƷ����,���÷����ʱ��Ҫ������ǰ����,Ҫ���ַ���������sql�н����滻
    rsSQL.Fields.Append "����ID", adBigInt 'ҩ��ID ҩƷ����,���÷����ʱ��Ҫ������ǰ����,�������ɷ�ҩ����
    rsSQL.Fields.Append "�շ����", adVarChar, 30 '��������ҩƷ����,���÷����ʱ��Ҫ��,����='4' ҩƷ='5'
    rsSQL.Fields.Append "��ҽ��IDs", adVarChar, 50000 '������Һ��ҩ��¼������ʱʹ��
    rsSQL.Fields.Append "������Դ", adBigInt '������Դ,1-������ü�¼,2-סԺ���ü�¼
    rsSQL.Fields.Append "��������", adBigInt '��������,1-�շ�,2-����
    rsSQL.Fields.Append "SQLEX", adVarChar, 5000 '��������������Ҫ��Ϣƴ��
    rsSQL.Fields.Append "ҩƷ����", adVarChar, 5000 'ҩƷ������������Ҫ��Ϣƴ��,Ҫ���ǰ��ķ�������һ�����
    rsSQL.Fields.Append "ͬ�����", adVarChar, 5000  '1-ҩƷ,2-����,��ʽ:"ҽ��ID:���ͺ�:1,ҽ��ID:���ͺ�:2",һ��ҽ��ֻ������3�����,ҩƷ/����/ҩƷ+����
    rsSQL.Fields.Append "����", adBigInt '����
    rsSQL.Fields.Append "����", adDouble '��׼����
    rsSQL.Fields.Append "����", adBigInt '��������
    rsSQL.Fields.Append "��������ID", adBigInt '��������ID
    rsSQL.Fields.Append "ִ�����", adBigInt '�Զ�ִ�����,0-��ִ�����,1-Ҫִ�����
    
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
End Sub

Public Function GetTakeNos(ByVal lng����ID As Long, ByVal strTitle As String) As String
'����:��ȡ��ҩ��
    Dim strJsonIn As String
    On Error GoTo errH
    strJsonIn = "{""input"":{""dept_id"":" & lng����ID & "}}"
    Call CallService("Zl_Drugsvr_GetTakeNo", strJsonIn, , strTitle, , False, , , , True)
    GetTakeNos = gobjService.GetJsonNodeValue("output.data")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetStockCheck(ByVal bytType As Byte) As Collection
'���ܣ���ȡҩƷ�����ĳ�����ļ���
'������bytType:0-ҩƷ��1-����

    Dim colStock As Collection, i As Long
    Dim strJsonOut As String
    Dim colList As Collection
    Dim strName As String
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '�������
    If bytType = 0 Then
        strName = "zl_DrugSvr_Getstockcheck"
    Else
        strName = "Zl_Stuffsvr_Getstockcheck"
    End If
    
    If CallService(strName, "", strJsonOut, "mdlCisService.GetStockCheck") Then
       Set colList = gobjService.GetJsonListValue("output.item_list")
    End If
    If Not colList Is Nothing Then
        If colList.Count > 0 Then
            For i = 1 To colList.Count
                If bytType = 0 Then
                    colStock.Add Val("" & colList(i)("_check_type")), "_" & colList(i)("_pharmacy_id")
                Else
                    colStock.Add Val("" & colList(i)("_check_type")), "_" & colList(i)("_warehouse_id")
                End If
            Next
        End If
    End If
    Set GetStockCheck = colStock
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function GetPatiʣ���(ByVal strTitle As String, ByVal lngMod As Long, ByRef colIn As Collection) As Double
'���ܣ���ȡסԺ����ʣ���
'��Σ�colIn ������κͳ��� 4��Ԫ�أ�����id,��ҳid,�������ʣ�����,��ѯ��ʽ
'       query_type ��ѯ��ʽ
'               0-���ﲡ����Ϣ
'               1-����ҽ���´�,���ʱ�����,����ҽ���嵥��ʾҽ�����
'               2-סԺҽ���´�,���ʱ�����
'               3-סԺ�༭״̬����ʾ
'               4-��ȡ �������.����������Ȩ�޿��Ƽ����ʾ��
'               5-����ҽ������״̬����ʾ
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim strTmp As String
    
    On Error GoTo errH
    
    strJsonIn = "{""input"":{" & _
        """pati_id"":" & colIn("_pati_id") & "," & _
        """pati_pageid"":" & GetColVal(colIn, "_pati_pageid", "N", 0) & "," & _
        """pati_character"":" & GetColVal(colIn, "_pati_character", "N", 0) & "," & _
        """insurance_type"":" & GetColVal(colIn, "_insurance_type", "N", 0) & "," & _
        """query_type"":" & GetColVal(colIn, "_query_type", "N", 0) & _
        "}}"
        
    If CallService("Zl_Exsesvr_Getexcessfunds", strJsonIn, strJsonOut, strTitle, lngMod, True) Then
        Set colIn = New Collection: strTmp = "" & gobjService.GetJsonNodeValue("output.excess_funds")
        colIn.Add strTmp, "���": strTmp = "" & gobjService.GetJsonNodeValue("output.prepay_funds")
        colIn.Add strTmp, "Ԥ�����": strTmp = "" & gobjService.GetJsonNodeValue("output.prepaid_expenses")
        colIn.Add strTmp, "Ԥ�����": strTmp = "" & gobjService.GetJsonNodeValue("output.guarantee_amount")
        colIn.Add strTmp, "������": strTmp = ""
        GetPatiʣ��� = Val("" & gobjService.GetJsonNodeValue("output.excess_funds"))
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPati����(ByVal strTitle As String, ByVal lngMod As Long, ByVal lng����ID As Long) As ADODB.Recordset
'���ܣ���ȡסԺ����ʣ���
'��Σ�colIn 4��Ԫ�أ�����id,��ҳid,�������ʣ�����
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim rsTmp As New ADODB.Recordset
    Dim colList As New Collection
    Dim i As Long
    
    On Error GoTo errH
   
    
    rsTmp.Fields.Append "��ϵ", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "����", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "�Ա�", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "����", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "���￨��", adVarChar, 500, adFldIsNullable
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    
    strJsonIn = "{""input"":{""pati_id"":" & lng����ID & "}}"
    If CallService("zl_patisvr_getfamilymembers", strJsonIn, strJsonOut, strTitle, lngMod, True) Then
        Set colList = gobjService.GetJsonListValue("output.item_list")
        If colList.Count > 0 Then
            For i = 1 To colList.Count
                rsTmp.AddNew
                rsTmp!��ϵ = colList(i)("_relationship")
                rsTmp!���� = colList(i)("_pati_name")
                rsTmp!�Ա� = colList(i)("_pati_sex")
                rsTmp!���� = colList(i)("_pati_age")
                rsTmp!���￨�� = colList(i)("_visit_card_no")
                rsTmp.Update
            Next
            rsTmp.MoveFirst
        End If
    End If
    Set GetPati���� = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get���ò���(ByVal strTitle As String, ByVal lngMod As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ���ȡ ���ò���
'��Σ�colIn ������κͳ��� 4��Ԫ�أ�����id,��ҳid,�������ʣ�����
'
    Dim strJsonIn As String
    Dim strJsonOut As String
   
    
    On Error GoTo errH
    
    strJsonIn = "{""input"":{" & _
        """pati_id"":" & lng����ID & "," & _
        """pati_pageid"":" & lng��ҳID & _
        "}}"
        
    If CallService("Zl_Patisvr_Patiwarnscheme", strJsonIn, strJsonOut, strTitle, lngMod, True) Then
        Get���ò��� = "" & gobjService.GetJsonNodeValue("output.pati_scheme")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get������(ByVal strTitle As String, ByVal lngMod As Long, ByRef colIn As Collection) As ADODB.Recordset
'����:��ѯ���ʱ�����
    '��Σ�colIn ������κͳ��� 4��Ԫ�أ�����id,��ҳid,�������ʣ�����
'  --     query_type   N 1 ��ѯ��ʽ
'  --                     0-������ ����id / ���ò��� ���ң�����һ��ֵ,���ڼ��ʱ�����ʾ
'  --                     1-������id ���ң������б�,���˲����б��ſ�Ƿ�Ѳ���
    Dim strJsonIn As String
    Dim strJsonOut As String
 
    Dim rsTmp As New ADODB.Recordset
    Dim colList As New Collection
    Dim i As Long
    
    On Error GoTo errH
     
    rsTmp.Fields.Append "���ò���", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "��������", adInteger, , adFldIsNullable
    rsTmp.Fields.Append "����ֵ", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "������־1", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "������־2", adVarChar, 500, adFldIsNullable
    rsTmp.Fields.Append "������־3", adVarChar, 500, adFldIsNullable
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    
    strJsonIn = "{""input"":{" & _
        """pati_scheme"":""" & GetColVal(colIn, "_pati_scheme") & """," & _
        """wardarea_id"":" & GetColVal(colIn, "_wardarea_id", "N", 0) & "," & _
        """query_type"":" & GetColVal(colIn, "_query_type", "N", 0) & _
        "}}"
        

    If CallService("Zl_Exsesvr_Getwarnline", strJsonIn, strJsonOut, strTitle, lngMod, True) Then
    
        If Val("" & gobjService.GetJsonNodeValue("output.alarm_value")) > 0 Then
            rsTmp.AddNew
            rsTmp!����ֵ = Val("" & gobjService.GetJsonNodeValue("output.alarm_value"))
            rsTmp.Update
        End If
        
        
        Set colList = gobjService.GetJsonListValue("output.item_list")
        If colList.Count > 0 Then
            For i = 1 To colList.Count
                rsTmp.AddNew
                rsTmp!���ò��� = "" & colList(i)("_pati_scheme")
                rsTmp!�������� = Val("" & colList(i)("_alarm_way"))
                rsTmp!����ֵ = Val("" & colList(i)("_alarm_value"))
                rsTmp!������־1 = Val("" & colList(i)("_alarm_one"))
                rsTmp!������־2 = Val("" & colList(i)("_alarm_two"))
                rsTmp!������־3 = Val("" & colList(i)("_alarm_three"))
                rsTmp.Update
            Next
            rsTmp.MoveFirst
        End If
    End If
    Set Get������ = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetFeeJsonIn(ByVal colData As Collection) As String
'����:���ط����������,һ������һ����һ��
'Zl_Exsesvr_Newbill
    Dim i As Long
    Dim colItem As Collection
    Dim strPati As String
    Dim strItems As String
    Dim strOut As String
    
    On Error GoTo errH
    
    If colData.Count = 0 Then
        Exit Function
    End If
    
    For i = 1 To colData.Count
        Set colItem = colData(i)
        
        If strPati = "" Then
            strPati = GetJsonStrNode("billtype", "��������", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_source", "������Դ", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_id", "����ID", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_pageid", "��ҳID", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("baby_num", "Ӥ����", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("sgin_no", "��ʶ��", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("bed_num", "����", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_name", "����", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_sex", "�Ա�", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("pati_age", "����", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("fee_category", "�ѱ�", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("overtime_sign", "�Ӱ��־", "N", colItem, 1)
            strPati = strPati & "," & GetJsonStrNode("pati_deptid", "���˿���ID", "N", colItem, 1)
            strPati = strPati & "," & GetJsonStrNode("pati_wardarea_id", "���˲���ID", "N", colItem, 1)
            strPati = strPati & "," & GetJsonStrNode("operator_name", "����Ա����", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("operator_code", "����Ա���", "C", colItem)
            strPati = strPati & "," & GetJsonStrNode("outpati_tag", "�����־", "N", colItem)
            strPati = strPati & "," & GetJsonStrNode("rgst_id", "�Һ�ID", "N", colItem, 1)
            strPati = strPati & "," & GetJsonStrNode("emg_sign", "�Ƿ���", "N", colItem, 1)
        End If
        
        strItems = strItems & ",{" & GetJsonStrNode("fee_id", "����ID", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("fee_no", "NO", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("serial_num", "���", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("charge_tag", "����", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("placer", "������", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("plcdept_id", "��������ID", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("sub_serial_num", "��������", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("fitem_id", "�շ�ϸĿID", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("item_type", "�շ����", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("unit", "���㵥λ", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("pharmacy_window", "��ҩ����", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("packages_num", "����", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("send_num", "����", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("ext_mark", "���ӱ�־", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("exe_deptid", "ִ�в���ID", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("price_ftrnum", "�۸񸸺�", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("income_item_id", "������ĿID", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("receipt_name", "�վݷ�Ŀ", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("price", "��׼����", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("fee_amrcvb", "Ӧ�ս��", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("fee_ampaib", "ʵ�ս��", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("happen_time", "����ʱ��", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("create_time", "�Ǽ�ʱ��", "C", colItem)
        'strItems = strItems & "," & GetJsonStrNode("memo", "����ժҪ", "C", colItem)
        strItems = strItems & ",""memo"":""" & zlStr.ToJsonStr(GetColVal(colItem, "����ժҪ")) & """"
        strItems = strItems & "," & GetJsonStrNode("order_id", "ҽ�����", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("baby_num", "Ӥ����", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("exe_properties", "ִ������", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("decoction_method", "�巨", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("morphology", "��ҩ��̬", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("bakstuff_batch", "����", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("insurance", "������Ŀ��", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("insure_id", "���մ���ID", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("insure_code", "���ձ���", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("fee_type", "��������", "C", colItem)
        strItems = strItems & "," & GetJsonStrNode("si_manp_money", "ͳ����", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("synchro", "����ͬ����־", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("effective_time", "��Ч", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("receipt_issecret", "����", "N", colItem)
        strItems = strItems & "," & GetJsonStrNode("takedept_id", "��ҩ����ID", "N", colItem, 1)
        strItems = strItems & "," & GetJsonStrNode("group_id", "ҽ��С��ID", "N", colItem, 1)
        strItems = strItems & "}"
    Next
    
    strOut = "{""input"":{" & strPati & ",""item_list"":[" & Mid(strItems, 2) & "]}}"
    GetFeeJsonIn = strOut
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetFeeNodes(ByVal lng��Դ As Long, ByVal lng���� As Long, ByVal strPar As String, ByRef colOut As Collection)
'����:�������ý��
'����:lng��Դ ������Դ,1-������ü�¼,2-סԺ���ü�¼
'     lng���� ��������,1-�շ�,2-����
    Dim lngIdx As Long
    Dim varTmp As Variant
   
    
    On Error GoTo errH
    Set colOut = New Collection
    
    colOut.Add lng��Դ & "", "������Դ"
    colOut.Add lng���� & "", "��������"
    
    'ZL_������ʼ�¼_INSERT
    If lng��Դ = 1 And lng���� = 2 Then
        lngIdx = 0
        varTmp = Split(strPar, "<SPCHAR>")
        colOut.Add varTmp(lngIdx) & "", "FNO": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "���": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��ʶ��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "�Ա�": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "�ѱ�": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "�Ӱ��־": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "Ӥ����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���˿���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��������ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "������": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "��������": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�շ�ϸĿID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�շ����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "���㵥λ": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���ӱ�־": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "ִ�в���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�۸񸸺�": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "������ĿID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�վݷ�Ŀ": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "��׼����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "Ӧ�ս��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "ʵ�ս��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ʱ��": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "�Ǽ�ʱ��": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "F��ҩ����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����Ա���": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����Ա����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "F����ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���ʵ�ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ժҪ": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "ҽ�����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�����־": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��ҩ��̬": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "�巨": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "��ҳID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���˲���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ͬ����־": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��Ч": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�Һ�ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�Ƿ���": lngIdx = lngIdx + 1 '--NUMBER
        
        If lngIdx <= UBound(varTmp) Then
            colOut.Add varTmp(lngIdx) & "", "NO": lngIdx = lngIdx + 1 '--VARCHAR2
            colOut.Add varTmp(lngIdx) & "", "����ID": lngIdx = lngIdx + 1 '--NUMBER
            colOut.Add varTmp(lngIdx) & "", "��ҩ����": lngIdx = lngIdx + 1 '--VARCHAR2
        End If
    End If
    
    'ZL_���ﻮ�ۼ�¼_INSERT
    If lng��Դ = 1 And lng���� = 1 Then
        lngIdx = 0
        varTmp = Split(strPar, "<SPCHAR>")
        colOut.Add varTmp(lngIdx) & "", "FNO": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "���": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��ҳID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��ʶ��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���ʽ": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "�Ա�": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "�ѱ�": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "�Ӱ��־": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���˿���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��������ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "������": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "��������": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�շ�ϸĿID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�շ����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "���㵥λ": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "F��ҩ����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���ӱ�־": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "ִ�в���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�۸񸸺�": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "������ĿID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�վݷ�Ŀ": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "��׼����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "Ӧ�ս��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "ʵ�ս��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ʱ��": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "�Ǽ�ʱ��": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "����Ա����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "F����ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ժҪ": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "ҽ�����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�巨": lngIdx = lngIdx + 1 '--VARCHAR2
        'colOut.Add varTmp(lngIdx) & "", "������Դ": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���ձ���": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "��������": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "������Ŀ��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���մ���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��ҩ��̬": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "ִ����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "���˲���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ͬ����־": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��Ч": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�Һ�ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�Ƿ���": lngIdx = lngIdx + 1 '--NUMBER
        
        If lngIdx <= UBound(varTmp) Then
            colOut.Add varTmp(lngIdx) & "", "NO": lngIdx = lngIdx + 1 '--VARCHAR2
            colOut.Add varTmp(lngIdx) & "", "����ID": lngIdx = lngIdx + 1 '--NUMBER
            colOut.Add varTmp(lngIdx) & "", "��ҩ����": lngIdx = lngIdx + 1 '--VARCHAR2
        End If
    End If
    
    'ZL_סԺ���ʼ�¼_Insert
    If lng��Դ = 2 And lng���� = 2 Then
        lngIdx = 0
        varTmp = Split(strPar, "<SPCHAR>")
        colOut.Add varTmp(lngIdx) & "", "FNO": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "���": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��ҳID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��ʶ��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "�Ա�": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "�ѱ�": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "���˲���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���˿���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�Ӱ��־": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "Ӥ����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��������ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "������": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "��������": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�շ�ϸĿID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�շ����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "���㵥λ": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "������Ŀ��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���մ���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���ձ���": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���ӱ�־": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "ִ�в���ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�۸񸸺�": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "������ĿID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�վݷ�Ŀ": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "��׼����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "Ӧ�ս��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "ʵ�ս��": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "ͳ����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ʱ��": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "�Ǽ�ʱ��": lngIdx = lngIdx + 1 '--DATE
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����Ա���": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����Ա����": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "F����ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�ಡ�˵�": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "���ʵ�ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ժҪ": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "�Ƿ���": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "ҽ�����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�򵥼���": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��������": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "ҽ�����ٴ�����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��ҩ��̬": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "ҽ��С��ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "�巨": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "ִ������": lngIdx = lngIdx + 1 '--VARCHAR2
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��ҩ����ID": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����ͬ����־": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "��Ч": lngIdx = lngIdx + 1 '--NUMBER
        colOut.Add varTmp(lngIdx) & "", "����": lngIdx = lngIdx + 1 '--NUMBER
        
        If lngIdx <= UBound(varTmp) Then
            colOut.Add varTmp(lngIdx) & "", "NO": lngIdx = lngIdx + 1 '--VARCHAR2
            colOut.Add varTmp(lngIdx) & "", "����ID": lngIdx = lngIdx + 1 '--NUMBER
            colOut.Add varTmp(lngIdx) & "", "��ҩ����": lngIdx = lngIdx + 1 '--VARCHAR2
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetJsonStrNode(ByVal strNode As String, ByVal strKey As String, ByVal strType As String, ByVal colNode As Collection, Optional ByVal lngForce As Long) As String
'����:����JSON���Ľ��
'����:strNode �����,
'     strKey  ����colNode�еĹؼ���
'     strType �������,"N"����,"C" �ַ�
'     colNode Ŀ�꼯��
'     lngForce 1-�������������͵�,�����0ת��ΪNULL
    Dim strOut As String
    Dim strVal As String
    Dim lngExist As Long
    
    strOut = """" & strNode & """:"
    
    strVal = GetColVal(colNode, strKey, strType, lngExist)
    
    If lngExist = -1 Or UCase(strVal) = "NULL" Then
        If strType = "N" Then
            strVal = "null"
        Else
            strVal = """"""
        End If
    ElseIf strType = "C" Then
        strVal = """" & strVal & """"
    ElseIf strType = "N" Then
        If strVal = "" Then
            strVal = "null"
        ElseIf lngForce = 1 And strVal = "0" Then
            strVal = "null"
        Else
            strVal = GetJsonNum(strVal)
        End If
    End If
    
    strOut = strOut & strVal
    
    GetJsonStrNode = strOut
End Function

Public Function ExsesvrGetNextID(strTable As String, Optional strFild As String, Optional ByVal strTitle As String, Optional ByVal lngģ�� As Long, Optional ByVal lng���� As Long, Optional ByRef varIDs As Variant) As Double
'����:��ȡ�������еõ�IDֵ
    Dim strJsonIn As String
    Dim strTmp As String
    
    If lng���� <= 1 Then
        lng���� = 0
    End If
    
    strJsonIn = "{""input"":{""table_name"":""" & strTable & """,""col_name"":""" & strFild & """,""quantity"":" & lng���� & "}}"
    Call CallService("Zl_Exsesvr_Getnextid", strJsonIn, , strTitle, lngģ��, True)
    
    If lng���� > 1 Then
        strTmp = gobjService.GetJsonNodeValue("output.next_id")
    Else
        ExsesvrGetNextID = gobjService.GetJsonNodeValue("output.next_id")
        strTmp = ExsesvrGetNextID
    End If
    varIDs = Split(strTmp, ",")
        
End Function

Public Function GetFinishFeeJsonIn(ByVal colData As Collection) As String
'����:�Զ����ϵ����ķ��ø���Ϊִ�����״̬
 
    Dim i As Long
    Dim colItem As Collection
    Dim strItems As String
    Dim strOut As String
 
    Dim lngFee_origin As Long
    
    On Error GoTo errH
    
    If colData.Count = 0 Then
        Exit Function
    End If
    
    For i = 1 To colData.Count
        Set colItem = colData(i)
        If GetColVal(colItem, "�Զ�����") = "1" Then
            If Val(GetColVal(colItem, "����ID")) > 0 Then
                strItems = strItems & "," & GetColVal(colItem, "����ID")
            End If
        End If
    Next
    If strItems <> "" Then
        If GetColVal(colItem, "��������") = "2" And GetColVal(colItem, "������Դ") = "2" Then
            lngFee_origin = 2
        Else
            lngFee_origin = 1
        End If
    
    
        strOut = "{""input"":{""fee_ids"":""" & Mid(strItems, 2) & """,""oper_type"":1"
        strOut = strOut & "," & GetJsonStrNode("exe_people", "����Ա����", "C", colItem)
        strOut = strOut & "," & GetJsonStrNode("exe_time", "�Ǽ�ʱ��", "C", colItem)
        strOut = strOut & ",""fee_origin"":" & lngFee_origin
        strOut = strOut & ",""exe_status"":1"
        strOut = strOut & "}}"
    End If
    GetFinishFeeJsonIn = strOut
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetƤ�Խ����Ϣ(ByVal lng����ID As String, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, Optional ByRef blnHaveData As Boolean) As ADODB.Recordset
'���ܣ���ȡ����Ƥ�Խ����Ϣ
    Dim strSQL As String
    Dim rsƤ�� As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "select a.Ƥ�Խ��, a.ҩ��id, a.ҩƷid from" & vbNewLine & _
        "(Select a.Ƥ�Խ��, c.��Ŀid As ҩ��id,0 as ҩƷid,a.��ʼִ��ʱ��" & vbNewLine & _
        "From ����ҽ����¼ A, �����÷����� C" & vbNewLine & _
        "Where a.������Ŀid = c.�÷�id And Nvl(c.����, 0) = 0 And Nvl(a.ҽ��״̬, 0) = 8 and a.Ƥ�Խ�� is not null" & vbNewLine & _
        IIF(str�Һŵ� = "", " and a.����id=[1] and a.��ҳid=[2]", " and a.�Һŵ�=[3]") & vbNewLine & _
        "union all" & vbNewLine & _
        "Select a.Ƥ�Խ��, b.ҩ��id, b.ҩƷid,a.��ʼִ��ʱ��" & vbNewLine & _
        "From ����ҽ����¼ A, ҩƷ��� B, ҩƷ�÷����� C" & vbNewLine & _
        "Where a.������Ŀid = c.�÷�id And b.ҩƷid = c.ҩƷid And Nvl(c.����, 0) = 0 And Nvl(a.ҽ��״̬, 0) = 8 And  a.Ƥ�Խ��<>'����'" & vbNewLine & _
        IIF(str�Һŵ� = "", " and a.����id=[1] and a.��ҳid=[2]", " and a.�Һŵ�=[3]") & vbNewLine & _
        "union all" & vbNewLine & _
        "Select a.Ƥ�Խ��, c.��Ŀid As ҩ��id,0 as ҩƷid,a.��ʼִ��ʱ��" & vbNewLine & _
        "From ����ҽ����¼ A, �����÷����� C" & vbNewLine & _
        "Where a.������Ŀid = c.�÷�id And Nvl(c.����, 0) = 0 And a.Ƥ�Խ��='����'" & vbNewLine & _
        IIF(str�Һŵ� = "", " and a.����id=[1] and a.��ҳid=[2]", " and a.�Һŵ�=[3]") & vbNewLine & _
        " ) a" & vbNewLine & _
        "group by a.Ƥ�Խ��, a.ҩ��id, a.ҩƷid,a.��ʼִ��ʱ��" & vbNewLine & _
        "order by a.��ʼִ��ʱ�� desc"
    Set rsƤ�� = zlDatabase.OpenSQLRecord(strSQL, "mdlCISService", lng����ID, lng��ҳID, str�Һŵ�)
    
    If Not rsƤ��.EOF Then
        blnHaveData = True
    Else
        blnHaveData = False
    End If
    
    Set GetƤ�Խ����Ϣ = rsƤ��
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CISGetNextId(strTable As String, Optional strFild As String) As Long
    '------------------------------------------------------------------------------------
    '���ܣ���ȡָ��������Ӧ������(���淶������������Ϊ��������_id��)����һ��ֵ(�ٴ�)
    '������
    '   strTable��������;strFild�ֶ������������Ʋ�һ����ID�������¼ID
    '���أ�
    '------------------------------------------------------------------------------------
    '�����ô������,ԭ��������ʧЧ��û������ʱ,Ӧ�÷��ش���,��Ȼ������,��������!

    Dim strJsonIn As String
    Dim strJsonOut As String

    strJsonIn = "{""input"":{""table_name"":""" & Trim(strTable) & _
        """,""col_name"":""" & Trim(strFild) & """}}"
        
    If CallService("Zl_Cissvr_Getnextid", strJsonIn, strJsonOut, "zlCISkernel.mdlCISService.CISGetNextId", pסԺҽ���´�, True) Then
         CISGetNextId = Val(gobjService.GetJsonNodeValue("output.next_id"))
    End If
End Function

Public Function GetAdviceInfo(intType As Integer, ByVal strҽ��IDs As String) As Collection
'���ܣ���ȡ����ҽ��������Ϣ��չ����
'������intType ��ѯ���ͣ�0:��ѯ������Ϣ��1:��ѯ������Ϣ+��չ��Ϣ

    Dim strJsonIn As String
    Dim strJsonOut As String
 
    On Error GoTo errH

    strJsonIn = "{""input"":{""query_type"":" & intType & ",""advice_ids"":""" & strҽ��IDs & """}}"
        
    If CallService("Zl_Cissvr_Getadviceinfo", strJsonIn, strJsonOut, "zlCISkernel.mdlCISService.GetAdviceInfo", pסԺҽ���´�, True) Then
        Set GetAdviceInfo = gobjService.GetJsonListValue("output.advice_list")
    End If
     
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetRsΣ��ֵ��¼(ByVal intType As Integer, ByVal lngIdin As Long, Optional ByVal dt��ʼʱ�� As Date, Optional ByVal dt����ʱ�� As Date, _
                Optional ByVal int�������� As Integer, Optional ByVal lng�������id As Long, Optional ByVal lngȷ�Ͽ���id As Long, Optional ByVal intȷ��״̬ As Integer) As ADODB.Recordset
'���ܣ�ͨ��Σ��ֵid��ȡ����Σ��ֵ��Ϣ,�����ؼ�¼��
'intType 1-ͨ��id��ѯ,4-����Σ��ֵ�б�

    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim colList As Collection
    Dim rsReturn As Recordset
    Dim i As Long
    On Error GoTo errH

    If intType = 1 Then
        strJsonIn = "{""input"":{""use_type"":1" & _
            ",""cvalue_ids"":""" & lngIdin & """}}"
    ElseIf intType = 4 Then
        strJsonIn = "{""input"":{" & _
                        """use_type"":" & "4" & "," & _
                        """cvalue_time_begin"":""" & Format(dt��ʼʱ��, "yyyy-MM-dd HH:mm") & """," & _
                        """cvalue_time_end"":""" & Format(dt����ʱ��, "yyyy-MM-dd HH:mm") & """," & _
                        """pati_type"":" & int�������� & "," & _
                        """rpt_deptid"":" & lng�������id & "," & _
                        """cvalue_deptid"":" & lngȷ�Ͽ���id & "," & _
                        """cvalue_rec_status"":" & intȷ��״̬ & _
                    "}}"
    End If

    Set rsReturn = New ADODB.Recordset
    rsReturn.Fields.Append "ID", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "ҽ��ID", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "����", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "�Ա�", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "����", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "����ʱ��", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "״̬", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "�Ƿ�Σ��ֵ", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "Σ��ֵ����", adVarChar, 4000, adFldIsNullable
    
    
    rsReturn.Fields.Append "������Դ", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "����id", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "��ҳid", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "�Һŵ�", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "Ӥ��", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "�걾id", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "�������id", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "������", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "�������", adVarChar, 4000, adFldIsNullable
    rsReturn.Fields.Append "ȷ��ʱ��", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "ȷ����", adVarChar, 500, adFldIsNullable
    rsReturn.Fields.Append "ȷ�Ͽ���id", adBigInt, , adFldIsNullable
    rsReturn.Fields.Append "ȷ�Ͽ���", adVarChar, 500, adFldIsNullable
    
    rsReturn.CursorLocation = adUseClient
    rsReturn.LockType = adLockOptimistic
    rsReturn.CursorType = adOpenStatic
    rsReturn.Open

    If CallService("Zl_Cissvr_Getcriticalinfo", strJsonIn, strJsonOut, "GetRsΣ��ֵ��¼", pסԺҽ���´�, True) Then
        Set colList = gobjService.GetJsonListValue("output.cvalue_list")
        
        If Not colList Is Nothing Then
            If colList.Count > 0 Then

                For i = 1 To colList.Count
                    rsReturn.AddNew
                    rsReturn!ID = colList(i)("_cvalue_id")
                    rsReturn!ҽ��ID = colList(i)("_advice_id")
                    rsReturn!���� = colList(i)("_pat_name")
                    rsReturn!�Ա� = colList(i)("_pat_sex")
                    rsReturn!���� = colList(i)("_pat_age")
                    rsReturn!����ʱ�� = colList(i)("_cvalue_rec_create_time")
                    rsReturn!״̬ = colList(i)("_cvalue_rec_status")
                    rsReturn!�Ƿ�Σ��ֵ = colList(i)("_cvitem_result")
                    rsReturn!Σ��ֵ���� = colList(i)("_cvalue_rec_desc")
                    rsReturn!������Դ = colList(i)("_cvitem_source")
                    rsReturn!����ID = colList(i)("_pati_id")
                    rsReturn!��ҳID = colList(i)("_pati_pageid")
                    rsReturn!�Һŵ� = colList(i)("_rgst_no")
                    rsReturn!Ӥ�� = colList(i)("_baby_num")
                    rsReturn!�걾id = colList(i)("_lspcm_id")
                    rsReturn!�������id = colList(i)("_rpt_deptid")
                    rsReturn!������ = colList(i)("_rec_rptor")
                    rsReturn!������� = colList(i)("_proc_note")
                    rsReturn!ȷ��ʱ�� = colList(i)("_cvalue_cnfmtime")
                    rsReturn!ȷ���� = colList(i)("_cvalue_cnfmer")
                    rsReturn!ȷ�Ͽ���id = colList(i)("_cnfm_deptid")
                    rsReturn!ȷ�Ͽ��� = colList(i)("_cvalue_dept")
                    rsReturn.Update
                Next
                rsReturn.Filter = 0
            End If
        End If
        
    End If
    Set GetRsΣ��ֵ��¼ = rsReturn
     
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiBaseInfo(ByVal intType As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal lngҽ��ID As Long, Optional ByVal strNO As String) As Collection
'���ܣ���ȡ���˻�����Ϣ��չ����
'������intType 1-ͨ������id����ҳid��ѯ  2-��ȡҽ����¼�еĲ�����Ϣ,3-ͨ���Һŵ��Ų�ѯ

    Dim strJsonIn As String
    Dim strJsonOut As String
    On Error GoTo errH

    strJsonIn = "{""input"":{""query_type"":" & intType & _
        ",""pati_id"":" & lng����ID & _
        ",""page_id"":" & lng��ҳID & _
        ",""advice_id"":" & lngҽ��ID & ",""reg_no"":""" & strNO & """}}"
        
    If CallService("Zl_Cissvr_Getpatibaseinfo", strJsonIn, strJsonOut, "zlCISkernel.mdlCISService.GetPatiBaseInfo", pסԺҽ���´�, True) Then
        Set GetPatiBaseInfo = gobjService.GetJsonListValue("output.page_list")
    End If
     
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitSvr() As Boolean
'���ܣ���ʼ������ӿڲ���
    If gobjService Is Nothing Then
        On Error Resume Next
        Set gobjService = CreateObject("zlServiceCall.clsServiceCall")
        If Not gobjService.InitService(gcnOracle, gstrDBUser, glngSys) Then
            Set gobjService = Nothing
        End If
        err.Clear: On Error GoTo 0
    End If
    If gobjService Is Nothing Then
        MsgBox "zlServiceCall.clsServiceCall����ʧ��!", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not gobjService Is Nothing Then InitSvr = True
End Function

Public Function CallService(ByVal strServiceName As String, _
    ByVal strJson_In As String, Optional ByRef strJson_out As String, Optional ByVal strTitle As String, _
    Optional lngModule As Long, Optional blnShowErrMsg As Boolean = True, Optional ByVal strAskDate As String, Optional varExpend As String, _
    Optional lngSys As Long, Optional blnReadServiceErr As Boolean) As Boolean
'���ܣ����÷���
'���˵���� zlServiceCall.clsServiceCall.CallService �ӿ�
 
    If InitSvr() Then
        If Not gobjService.CallService(strServiceName, strJson_In, strJson_out, strTitle, lngModule, blnShowErrMsg, strAskDate, varExpend, lngSys, blnReadServiceErr) Then Exit Function
        If Not blnShowErrMsg Then
            If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
                CallService = False: Exit Function
            End If
        End If
        CallService = True
    End If
End Function

Public Function GetSvrOutInfo(ByVal strJson As String, Optional ByVal blnShowErrMsg As Boolean = True) As Boolean
'���ܣ���Ա����ٴ�����鷽�����õĳ��ν�����ʽ
'������strJsonǰһ��������ִ�к�ĵĳ���Json��ʽ
'      blnShowErrMsg �Ƿ��ڲ�������ʾ��Ϣ
'���أ�true/false  code=1ʱΪtrue,code=0ʱΪfalse
    If InitSvr() Then
        GetSvrOutInfo = gobjService.SetJsonString(strJson)
        
        If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
            GetSvrOutInfo = False
            If blnShowErrMsg Then
                MsgBox gobjService.GetJsonNodeValue("output.message") & "", vbInformation, gstrSysName
            End If
        End If
    End If
End Function

Public Function GetToDateStr(ByVal strIn As String) As String
'���ܣ���oracle����ת���Ĵ��н�������
    Dim strTmp As String
    strTmp = Split(strIn, ",")(0)
    strTmp = Split(strTmp, "'")(1)
    GetToDateStr = strTmp
End Function
  
Public Function Exeҽ��ȡ��ִ�����(colProSQL As Collection, ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal lng����ִ�� As Long, ByVal lng����ID As Long, ByVal strTitle As String, ByVal lngģ�� As Long) As Boolean
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim strFCOut As String '���ü��ִ�к�ĳ���
    Dim lng�������� As String
    Dim blnTran As Boolean
    Dim strStuff As String
    Dim strDrug As String
    Dim i As Long
    Dim strFeeState As String
    Dim clFst As Collection '����ִ��״̬�ĸ��µ��б�
    
    On Error GoTo errH
    
    strJsonOut = zlDatabase.CallProcedure("Zl_����ҽ��ִ��_Cancel_Check", strTitle & "_" & lngģ��, _
                lngҽ��ID, lng���ͺ�, lng����ִ��, Empty)
    If Not GetSvrOutInfo(strJsonOut) Then
        Exit Function
    End If
    
    lng�������� = Val("" & gobjService.GetJsonNodeValue("output.fee_origin"))
    
    strJsonIn = "{""input"":{""is_finish"":2,""fee_origin"":" & lng�������� & ",""fee_nos"":""" & gobjService.GetJsonNodeValue("output.fee_nos") _
        & """,""order_status"":" & gobjService.GetJsonNodeValue("output.order_status") & ",""order_ids"":""" & gobjService.GetJsonNodeValue("output.order_ids") & """,""exe_deptid"":" & lng����ID & "}}"
    
    If Not CallService("Zl_Exsesvr_GetOrderFeeExeInfo", strJsonIn, strJsonOut, strTitle, lngģ��, True) Then
        Exit Function
    End If
    
    strStuff = gobjService.GetJsonNodeValue("output.stuffdtl_ids") & ""
    strDrug = gobjService.GetJsonNodeValue("output.rcpdtl_ids") & ""
    strFCOut = strJsonOut
    Set clFst = gobjService.GetJsonListValue("output.cancel_list")
    If strStuff <> "" Then
        '���Ϸ���
        strFeeState = Getȡ��ִ�з���״̬(strStuff, lng��������)
        strJsonIn = "{""input"":{""audit_operator"":""" & UserInfo.���� & """,""stuffdtl_ids"":""" & Replace(strStuff & ":", ",", ":,") & """}}"
        Call Exe�Զ���ҩ����(strJsonIn, strFeeState, "", "", strTitle, lngģ��)
    End If
    
    If strDrug <> "" Then
        strFeeState = Getȡ��ִ�з���״̬(strDrug, lng��������)
        '��ҩ���� ������ƴ�� �ս�ȥ
        strJsonIn = Get�Զ���ҩ���(strDrug)
        Call Exe�Զ���ҩ����("", "", strJsonIn, strFeeState, strTitle, lngģ��)
    End If

    strJsonIn = ""
    strFeeState = ""
    If Not clFst Is Nothing Then
        If clFst.Count > 0 Then
            For i = 1 To clFst.Count
                strFeeState = strFeeState & ",{" & _
                    """fee_id"":" & clFst(i)("_fee_id") & _
                   ",""exe_nums"":0" & _
                   "}"
            Next
            strJsonIn = "{""input"":{""fee_origin"":" & lng�������� & ",""item_list"":[" & Mid(strFeeState, 2) & "]}}"
        End If
    End If
    
    gcnOracle.BeginTrans: blnTran = True
    For i = 1 To colProSQL.Count
        Call zlDatabase.ExecuteProcedure(colProSQL(i) & "", strTitle)
    Next
    If strJsonIn <> "" Then
        Call CallService("Zl_Exsesvr_UpdateExeInfo", strJsonIn, , strTitle, lngģ��, False, , , , True)
    End If
    gcnOracle.CommitTrans: blnTran = False
    
    '����ٸ���һ�η���ִ��״̬
    '��Ϊǰ���п��ܻ�ʧ��
    
    Call Upd����ִ��״̬(lngҽ��ID, lng���ͺ�)
    
    Exeҽ��ȡ��ִ����� = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get�Զ���ҩ���(ByVal strIDs As String) As String
'���ܣ��Զ���ҩʱ�����
    Dim i As Long, varTmp As Variant, strItem As String
    
    On Error GoTo errH
'    strJsonIn = "{""input"":{""audit_operator"":""" & UserInfo.���� & """,""rcpdtl_ids"":""" & Replace(strDrug & ":", ",", ":,") & """}}"
    varTmp = Split(strIDs, ",")
    For i = 0 To UBound(varTmp)
        strItem = strItem & ",{""rcpdtl_id"":" & varTmp(i)
        strItem = strItem & "}"
    Next
    Get�Զ���ҩ��� = "{""input"":{""audit_operator"":""" & UserInfo.���� & """,""rcpdtl_list"":[" & Mid(strItem, 2) & "]}}"
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Getȡ��ִ�з���״̬(ByVal strIDs As String, ByVal lng_fee_origin As Long) As String
'���ܣ����ݷ�����ϸ��ȡȡ��ִ�з��÷������
    Dim i As Long, strItem As String
    Dim varTmp As Variant, strTmp As String
    
    On Error GoTo errH
    
    varTmp = Split(strIDs, ",")
    For i = 0 To UBound(varTmp)
        strItem = strItem & ",{""fee_id"":" & varTmp(i)
        strItem = strItem & ",""exe_nums"":0"
        strItem = strItem & "}"
    Next
    
    strTmp = "{""input"":{""fee_origin"":" & lng_fee_origin & ",""item_list"":[" & Mid(strItem, 2) & "]}}"
    
    Getȡ��ִ�з���״̬ = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Exeҽ��ִ�����(colProSQL As Collection, ByVal lngҽ��ID As String, ByVal lng���ͺ� As Long, ByVal lng����ִ�� As Long, ByVal lng����ID As Long, _
   ByVal strTime As String, ByVal strTitle As String, ByVal lngģ�� As Long) As Boolean
'����:ҽ��ִ�����
'����:colProSQL- �ٴ����ִ�й��̼���,�����ж��
'     lngҽ��ID- ִ����ɵ�ҽ��ID
'     lng���ͺ�- ���ͺ�
'     lng����ִ��- �Ƿ񵥶�ִ��,0-������ִ��,1-����ִ��
'     lng����ID- ִ�п���ID �������ü�¼��ִ�п���ID,0-�����ֿ���,1-ָ������
'     strTime- ִ��ʱ��,�ַ������ڸ�ʽ:
'     strTitle- ����
'     lngģ��- ���õ�ģ���
'����:true �ɹ�,false ʧ��
    Dim i As Long
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim blnTran As Boolean
    Dim strSvrName As String
    Dim colPar As New Collection
    Dim clTag As Collection
     
    On Error GoTo errH
 
    strJsonOut = zlDatabase.CallProcedure("Zl_����ҽ��ִ��_Finish_Check", strTitle & "_" & lngģ��, _
                lngҽ��ID, lng���ͺ�, lng����ִ��, Empty)
    If Not GetSvrOutInfo(strJsonOut) Then
        Exit Function
    End If
    
    If Not Fҽ��ִ�����(colProSQL, strJsonOut, strTime, lng����ID, strTitle, lngģ��, colPar) Then
        Exit Function
    End If
    
    
'C "ҩƷ��ҩ"
'C "���ķ���"
'C "����ִ��"
 
'Col "ҩƷ���쳣"
'C "ҩƷ�շ�"

'Col "�������쳣"
'C "�����շ�"

'Col "ҽ�����쳣"
'C "�����շ�"


'Col "ҽ��ִ�����"

    strJsonIn = GetColVal(colPar, "�����շ�")
    Set clTag = GetColObj(colPar, "ҽ�����쳣")
    strSvrName = "Zl_Exsesvr_Billverify"
    If strJsonIn <> "" Then
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To clTag.Count
            Call zlDatabase.ExecuteProcedure(clTag(i) & "", strTitle)
        Next
        Call CallService(strSvrName, strJsonIn, , strTitle, lngģ��, False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    
    Call ExeҩƷ�շ�ȷ��(colPar, strTitle, lngģ��)
    
    Call Exe�����շ�ȷ��(colPar, strTitle, lngģ��)
     
    strJsonIn = GetColVal(colPar, "ҩƷ��ҩ")
    strSvrName = "Zl_Drugsvr_Autosenddrug"
    If strJsonIn <> "" Then
        Call CallService(strSvrName, strJsonIn, , strTitle, lngģ��, False, , , , True)
    End If
    
    strJsonIn = GetColVal(colPar, "���ķ���")
    strSvrName = "Zl_Stuffsvr_Autosendstuff"
    If strJsonIn <> "" Then
        Call CallService(strSvrName, strJsonIn, , strTitle, lngģ��, False, , , , True)
    End If
 
    strJsonIn = GetColVal(colPar, "����ִ��")
    Set clTag = GetColObj(colPar, "ҽ��ִ�����")
    strSvrName = "Zl_Exsesvr_Updateexeinfo"
    If strJsonIn <> "" Or clTag.Count > 0 Then
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To clTag.Count
            Call zlDatabase.ExecuteProcedure(clTag(i) & "", strTitle)
        Next
        If strJsonIn <> "" Then Call CallService(strSvrName, strJsonIn, , strTitle, lngģ��, False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    Exeҽ��ִ����� = True
    
    Call Upd����ִ��״̬(lngҽ��ID, lng���ͺ�)
    
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Exeҽ��ִ�еǼ�ɾ��(colData As Collection, ByVal strTitle As String, ByVal lngģ�� As Long) As Boolean
'����:ɾ��ҽ��ִ�еǼǼ�¼
    Dim strJsonOut As String
    Dim strItems As String
    Dim strOrderDel As String
    Dim strFeeSta As String
    Dim blnTran As Boolean
    Dim colProSQL As Collection
    Dim strTmp As String
    Dim strPJson As String '���ý��
    Dim colSQLlist As Collection
    Dim i As Long
    
    On Error GoTo errH
            
    '������Զ�ȡ��ִ��������ȵ���
    If Val(GetColVal(colData, "�Զ�ȡ��") & "") = 1 Then
        strJsonOut = zlDatabase.CallProcedure("Zl_����ҽ��ִ��_Delete_check", strTitle & "_" & lngģ��, _
                Val(GetColVal(colData, "ҽ��ID")), Val(GetColVal(colData, "���ͺ�")), _
                GetColVal(colData, "ִ��ʱ��") & "", Val(GetColVal(colData, "����ִ��")), 1, 1, Empty)
        If Not GetSvrOutInfo(strJsonOut) Then
            Exit Function
        End If
     
        If Val(gobjService.GetJsonNodeValue("output.auto_cancel") & "") = 1 Then
            Set colProSQL = New Collection
            strTmp = "Zl_����ҽ��ִ��_Cancel_S(" & colData("ҽ��ID") & "," & colData("���ͺ�") & ",'" & UserInfo.���� & "'," & colData("����ִ��") & ",null)"
            colProSQL.Add strTmp
            If Not Exeҽ��ȡ��ִ�����(colProSQL, colData("ҽ��ID"), colData("���ͺ�"), colData("����ִ��"), colData("ִ�в���ID"), strTitle, lngģ��) Then
                Exit Function
            End If
        End If
    End If
    
    'ɾ��ִ�еǼ�
    strJsonOut = zlDatabase.CallProcedure("Zl_����ҽ��ִ��_Delete_check", strTitle & "_" & lngģ��, _
                Val(GetColVal(colData, "ҽ��ID")), Val(GetColVal(colData, "���ͺ�")), _
                GetColVal(colData, "ִ��ʱ��") & "", Val(GetColVal(colData, "����ִ��")), 0, 0, Empty)
    If Not GetSvrOutInfo(strJsonOut) Then
        Exit Function
    End If
    
    strItems = strPJson
    strItems = strItems & ",""order_status"":" & Val(gobjService.GetJsonNodeValue("output.order_status") & "") 'ҽ��ִ��״̬
    strItems = strItems & ",""upd_price_sta"":" & Val(gobjService.GetJsonNodeValue("output.upd_price_sta") & "")
    strItems = strItems & ",""price_req_time"":""" & gobjService.GetJsonNodeValue("output.price_req_time") & """"
    strItems = strItems & ",""lis_upd"":" & Val(gobjService.GetJsonNodeValue("output.lis_upd") & "")
    'ҽ��ɾ��ִ�еǼǵ����JSONƴװ���
    strOrderDel = "{""input"":{" & Mid(strItems, 2) & "}}"
    
    strOrderDel = "Zl_����ҽ��ִ��_Delete_s(" & Val(GetColVal(colData, "ҽ��ID")) & "," & Val(GetColVal(colData, "���ͺ�")) & _
                ",to_date('" & GetColVal(colData, "ִ��ʱ��") & "','yyyy-mm-dd hh24:mi:ss')" & _
                "," & Val(GetColVal(colData, "����ִ��")) & _
                "," & Val(gobjService.GetJsonNodeValue("output.upd_price_sta") & "") & _
                ",to_date('" & gobjService.GetJsonNodeValue("output.price_req_time") & "','yyyy-mm-dd hh24:mi:ss')" & _
                "," & Val(gobjService.GetJsonNodeValue("output.lis_upd") & "") & _
                "," & Val(gobjService.GetJsonNodeValue("output.order_status") & "") & _
                ")"
    
    strFeeSta = strFeeSta & ",""oper_type"":1"
    strFeeSta = strFeeSta & ",""fee_origin"":" & Val(gobjService.GetJsonNodeValue("output.fee_origin") & "") '������Դ(Ĭ��=2��1-������ã�2-סԺ����)
    strFeeSta = strFeeSta & ",""exe_status"":" & Val(gobjService.GetJsonNodeValue("output.fee_status") & "")
    strFeeSta = strFeeSta & "," & GetJsonStrNode("exe_deptid", "ִ�в���ID", "N", colData)
    strFeeSta = strFeeSta & ",""order_ids"":""" & gobjService.GetJsonNodeValue("output.order_ids") & """"
    strFeeSta = strFeeSta & ",""fee_nos"":""" & gobjService.GetJsonNodeValue("output.fee_nos") & """"
    strFeeSta = "{""input"":{" & Mid(strFeeSta, 2) & "}}" '���·���ִ��״̬�����JSONƴװ���
    
    Set colSQLlist = GetColObj(colData, "sqllist")
    
    
    gcnOracle.BeginTrans: blnTran = True
 
        Call zlDatabase.ExecuteProcedure(strOrderDel, "Exeҽ��ִ�еǼ�ɾ��")
        If Not CallService("Zl_Exsesvr_Updateexeinfo", strFeeSta, strJsonOut, strTitle, lngģ��, False) Then
            gcnOracle.RollbackTrans: blnTran = False
            MsgBox "����[Zl_Exsesvr_Updateexeinfo]ʧ�ܣ�" & gobjService.GetJsonNodeValue("output.message"), vbInformation, gstrSysName
            Exit Function
        End If
        
        If Not colSQLlist Is Nothing Then
            For i = 1 To colSQLlist.Count
                If colSQLlist(i) <> "" Then
                    Call zlDatabase.ExecuteProcedure(colSQLlist(i), "Exeҽ��ִ�еǼ�ɾ��")
                End If
            Next
        End If
        
        
    gcnOracle.CommitTrans: blnTran = False
    
    Exeҽ��ִ�еǼ�ɾ�� = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Exeҽ��ִ�еǼ�(colData As Collection, ByVal strTitle As String, ByVal lngģ�� As Long) As Boolean
'����:ҽ��ִ�еǼ�
    Dim strJsonOut As String
    Dim lng�Զ����  As Long
    Dim strOrderIns As String
    Dim strFeeSta As String
    Dim blnTran As Boolean
    Dim colProSQL As Collection
    Dim strTmp As String
    Dim colSQLAfter As Collection, i As Long
  
    On Error GoTo errH
    
    strJsonOut = zlDatabase.CallProcedure("Zl_����ҽ��ִ��_Insert_check", strTitle & "_" & lngģ��, _
                Val(GetColVal(colData, "ҽ��ID")), Val(GetColVal(colData, "���ͺ�")), CDate(GetColVal(colData, "Ҫ��ʱ��")), _
                Val(GetColVal(colData, "��������")), _
                CDate(GetColVal(colData, "ִ��ʱ��")), Val(GetColVal(colData, "����ִ��")), Val(GetColVal(colData, "�Զ����")), Val(GetColVal(colData, "ִ�н��")), _
                Val(GetColVal(colData, "��Һ���")), Val(GetColVal(colData, "������Ŀ����")), Val(GetColVal(colData, "���ó���")), _
                Empty)
                 
    If Not GetSvrOutInfo(strJsonOut) Then
        Exit Function
    End If
    
    lng�Զ���� = Val(gobjService.GetJsonNodeValue("output.auto_finish") & "")
    
    
    
'ҽ��id_In       In Number,
'���ͺ�_In       In Number,
'Ҫ��ʱ��_In     In Date,
'��������_In     In Number,
'ִ��ժҪ_In     In Varchar2,
'ִ����_In       In Varchar2,
'ִ��ʱ��_In     In Date,
'ִ�н��_In     In Number,
'δִ��ԭ��_In   In Varchar2,
'����Ա����_In   In Varchar2,
'ִ�в���id_In   In Number,
'��Һͨ��_In     In Varchar2,
'��¼��Դ_In     In Number,
'ִ�з�ʽ_In     In Number,
'����id_In       In Number, --pati_id N 1 ����id
'��ҳid_In       In Number, --pati_pageid N 1 ��ҳid
'�Һŵ�_In       In Varchar2, --reg_no C 1 �Һŵ���
'ִ��״̬_In     In Number, --exe_status N 1 ִ��״̬
'����ɼ�����_In In Number, --lis_upd N 1 ���¼���ɼ�
'��id_In         In Number, --order_main_id N 1 ��id,��ҽ��id
'���¼Ƽ�״̬_In In Number, --upd_price_sta N 1 �Ƽ�״̬����
'�Ƽ�Ҫ��ʱ��_In In Date, --price_req_time C 1 �Ƽ�Ҫ��ʱ��
'ҽ��ids_In      In Varchar2 --order_ids C 1 Ҫ����״̬��ҽ��id��

    strOrderIns = "Zl_����ҽ��ִ��_Insert_S(" & Val(GetColVal(colData, "ҽ��ID")) & "," & Val(GetColVal(colData, "���ͺ�")) & ",to_date('" & GetColVal(colData, "Ҫ��ʱ��") & "','yyyy-mm-dd hh24:mi:ss')," & Val(GetColVal(colData, "��������"))
    strOrderIns = strOrderIns & ",'" & GetColVal(colData, "ִ��ժҪ") & "','" & GetColVal(colData, "ִ����") & "',to_date('" & GetColVal(colData, "ִ��ʱ��") & "','yyyy-mm-dd hh24:mi:ss')," & GetColVal(colData, "ִ�н��")
    strOrderIns = strOrderIns & ",'" & GetColVal(colData, "δִ��ԭ��") & "','" & GetColVal(colData, "����Ա����") & "'," & Val(GetColVal(colData, "ִ�в���ID"))
    strOrderIns = strOrderIns & ",'" & IIF("NULL" = UCase(GetColVal(colData, "��Һͨ��")), "", GetColVal(colData, "��Һͨ��")) & "'"
    strOrderIns = strOrderIns & "," & Val(GetColVal(colData, "��¼��Դ")) & "," & Val(GetColVal(colData, "ִ�з�ʽ"))
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.pati_id") & "")
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.pati_pageid") & "")
    strOrderIns = strOrderIns & ",'" & gobjService.GetJsonNodeValue("output.reg_no") & "'"
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.order_status") & "")
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.lis_upd") & "")
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.order_main_id") & "")
    strOrderIns = strOrderIns & "," & Val(gobjService.GetJsonNodeValue("output.upd_price_sta") & "")
    strOrderIns = strOrderIns & ",to_date('" & gobjService.GetJsonNodeValue("output.price_req_time") & "','yyyy-mm-dd hh24:mi:ss')"
    strOrderIns = strOrderIns & ",'" & gobjService.GetJsonNodeValue("output.order_ids") & "'"
    strOrderIns = strOrderIns & ")"
    
    strFeeSta = strFeeSta & ",""oper_type"":1,""exe_type"":1"
    strFeeSta = strFeeSta & ",""fee_origin"":" & Val(gobjService.GetJsonNodeValue("output.fee_origin") & "")
    strFeeSta = strFeeSta & ",""exe_status"":" & Val(gobjService.GetJsonNodeValue("output.fee_status") & "")
    strFeeSta = strFeeSta & "," & GetJsonStrNode("exe_deptid", "ִ�в���ID", "N", colData)
    strFeeSta = strFeeSta & "," & GetJsonStrNode("exe_people", "ִ����", "C", colData)
    strFeeSta = strFeeSta & ",""exe_time"":""" & gobjService.GetJsonNodeValue("output.fee_exe_time") & """"
    strFeeSta = strFeeSta & ",""order_ids"":""" & gobjService.GetJsonNodeValue("output.order_ids") & """"
    strFeeSta = strFeeSta & ",""fee_nos"":""" & gobjService.GetJsonNodeValue("output.fee_nos") & """"
    strFeeSta = "{""input"":{" & Mid(strFeeSta, 2) & "}}" '����ִ��״̬������η�װ���
    
    
    Set colSQLAfter = GetColObj(colData, "sqllistafter")
    
    gcnOracle.BeginTrans: blnTran = True
        Call zlDatabase.ExecuteProcedure(strOrderIns, "Exeҽ��ִ�еǼ�")
        If Not CallService("Zl_Exsesvr_Updateexeinfo", strFeeSta, strJsonOut, strTitle, lngģ��, False) Then
            gcnOracle.RollbackTrans: blnTran = False
            MsgBox "����[Zl_Exsesvr_Updateexeinfo]ʧ�ܣ�" & gobjService.GetJsonNodeValue("output.message"), vbInformation, gstrSysName
            Exit Function
        End If
        
        
        
        If Not colSQLAfter Is Nothing Then
            For i = 1 To colSQLAfter.Count
                If colSQLAfter(i) <> "" Then
                    Call zlDatabase.ExecuteProcedure(colSQLAfter(i), "Exeҽ��ִ�еǼ�")
                End If
            Next
        End If
        
    gcnOracle.CommitTrans: blnTran = False
    
    If lng�Զ���� = 1 Then
        Set colProSQL = New Collection
        strTmp = "Zl_����ҽ��ִ��_Finish_S(" & colData("ҽ��ID") & "," & colData("���ͺ�") & ",to_date('" & colData("ִ��ʱ��") & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.���� & "'," & colData("����ִ��") & ")"
        colProSQL.Add strTmp
        Call Exeҽ��ִ�����(colProSQL, colData("ҽ��ID"), colData("���ͺ�"), colData("����ִ��"), colData("ִ�в���ID"), colData("ִ��ʱ��"), strTitle, lngģ��)
    End If
    
    Exeҽ��ִ�еǼ� = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Exeҽ��ִ�еǼ��޸�(colData As Collection, ByVal strTitle As String, ByVal lngģ�� As Long, Optional colBefore As Collection, Optional colAfter As Collection) As Boolean
'���ܣ��޸�ҽ��ִ�еǼǼ�¼,ԭ����:Zl_����ҽ��ִ��_Update
'����:colData �����"ԭִ��ʱ��,ҽ��ID,���ͺ�,Ҫ��ʱ��,��������,ִ��ժҪ,ִ����,ִ��ʱ��,ִ�н��,δִ��ԭ��,����ִ��,����Ա���,����Ա����,ִ�в���ID"
'     strTitle- ����
'     lngģ��- ���õ�ģ���
'     colBefore  �ٴ����ִ�й���,�ɷŵ��޸�ִ�еǼǹ���ǰִ�еĹ���
'     colAfter   �ٴ����ִ�й���,�ɷŵ��޸�ִ�еǼǹ��̺�ִ�еĹ���
'˵��:�����ڲ��Ὺ������
'    zlAdviceRegRecUpd = AdviceRegRecUpd(colData, strTitle, lngģ��, colBefore, colAfter)
    Dim i As Long
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim blnTran As Boolean
    Dim strFeeSta As String
    Dim strCisOutPar As String
    
    On Error GoTo errH
    
    strJsonIn = "{""input"":{""old_exe_time"":""" & GetColVal(colData, "ԭִ��ʱ��") & """" & _
            ",""order_id"":" & GetColVal(colData, "ҽ��ID") & _
            ",""send_no"":" & GetColVal(colData, "���ͺ�") & _
            ",""require_time"":""" & GetColVal(colData, "Ҫ��ʱ��") & """" & _
            ",""reg_num"":" & GetJsonNum(GetColVal(colData, "��������")) & _
            ",""exe_memo"":""" & GetColVal(colData, "ִ��ժҪ") & """" & _
            ",""exe_people"":""" & GetColVal(colData, "ִ����") & """" & _
            ",""exe_time"":""" & GetColVal(colData, "ִ��ʱ��") & """" & _
            ",""exe_result"":" & GetColVal(colData, "ִ�н��") & _
            ",""no_exe_rea"":""" & GetColVal(colData, "δִ��ԭ��") & """" & _
            ",""exe_alone"":" & GetColVal(colData, "����ִ��") & _
            ",""operator_name"":""" & GetColVal(colData, "����Ա����") & """" & _
            "}}"
    
    
    
    gcnOracle.BeginTrans: blnTran = True
        If Not colBefore Is Nothing Then
            If colBefore.Count > 0 Then
                For i = 1 To colBefore.Count
                    Call zlDatabase.ExecuteProcedure(colBefore(i) & "", strTitle)
                Next
            End If
        End If
    
        strCisOutPar = zlDatabase.CallProcedure("Zl_����ҽ��ִ��_Update_S", strTitle & "_" & lngģ��, CDate(GetColVal(colData, "ԭִ��ʱ��", , "0")), _
            Val(GetColVal(colData, "ҽ��ID")), Val(GetColVal(colData, "���ͺ�")), _
            CDate(GetColVal(colData, "Ҫ��ʱ��", , "0")), Val(GetColVal(colData, "��������")), GetColVal(colData, "ִ��ժҪ"), GetColVal(colData, "ִ����"), _
            CDate(GetColVal(colData, "ִ��ʱ��", , "0")), Val(GetColVal(colData, "ִ�н��")), GetColVal(colData, "δִ��ԭ��"), Val(GetColVal(colData, "����ִ��")), _
            UserInfo.���, UserInfo.����, Val(GetColVal(colData, "ִ�в���ID", "N", "0")), _
            Empty)
        
        
        If Not colAfter Is Nothing Then
            If colAfter.Count > 0 Then
                For i = 1 To colAfter.Count
                    Call zlDatabase.ExecuteProcedure(colAfter(i) & "", strTitle)
                Next
            End If
        End If
        If strCisOutPar <> "" Then
            Call gobjService.SetJsonString(strCisOutPar)
            strFeeSta = strFeeSta & ",""oper_type"":1,""exe_type"":1"
            strFeeSta = strFeeSta & ",""fee_origin"":" & Val(gobjService.GetJsonNodeValue("output.fee_origin") & "")
            strFeeSta = strFeeSta & ",""exe_status"":" & Val(gobjService.GetJsonNodeValue("output.fee_status") & "")
            strFeeSta = strFeeSta & ",""exe_deptid"":" & GetColVal(colData, "ִ�в���ID", "N", "0")
            strFeeSta = strFeeSta & ",""exe_people"":""" & gobjService.GetJsonNodeValue("output.fee_exe_peo") & """"
            strFeeSta = strFeeSta & ",""exe_time"":""" & gobjService.GetJsonNodeValue("output.fee_exe_time") & """"
            strFeeSta = strFeeSta & ",""order_ids"":""" & gobjService.GetJsonNodeValue("output.order_ids") & """"
            strFeeSta = strFeeSta & ",""fee_nos"":""" & gobjService.GetJsonNodeValue("output.fee_nos") & """"
            
            strFeeSta = "{""input"":{" & Mid(strFeeSta, 2) & "}}" '����ִ��״̬������η�װ���
            Call CallService("Zl_Exsesvr_Updateexeinfo", strFeeSta, strJsonOut, strTitle, lngģ��, False, , , , True)
        End If
    gcnOracle.CommitTrans: blnTran = False
    
    Exeҽ��ִ�еǼ��޸� = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PatiSvrGetpatiinfo(intType As Integer, lng����ID As Long, Optional lngModel As Long, Optional int�������� As Integer) As Collection
'  ---------------------------------------------------------------------------
'  --����:��ȡ������Ϣ��Ŀǰֻ����ͨ������id����ȫ����Ϣ��
'  --��Σ�Json_In:��ʽ
'  --    input
'  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
'  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;3-����
'  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
'  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
'  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
'  --      query_immune      N 1 �Ƿ����������:1-����;0-������
'  --      query_cons_list[] C 1 ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��
'  --        pati_ids        C   ����IDs:����ö���
'  --        pati_name       C   ����:���Դ�%�ֺű������ƥ��
'  --        outpno  N   �����
'  --        pati_idcard     C   ���֤��
'  --        contacts_idcard C   ��ϵ�����֤��
'  --        cardtype_id     N   ҽ�ƿ����ID
'  --        card_no         C   ����
'  --        qrcode          C   ��ά��
'  --        iccard_no       C   Ic����
'  --        insurance_num   C   ҽ����
'  --        qrspt_statu     C   ��ѯסԺ״̬:0-������;1-��Ժ ;2-���Ｐ��Ժ
'  --        phone_number    C   �ֻ���
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""query_type"":" & intType & ",""pati_id"":" & lng����ID & IIF(int�������� = 0, "", ",""query_family"":1") & "}}"
        
    If CallService("Zl_Patisvr_Getpatiinfo", strJsonIn, strJsonOut, "PatiSvrGetpatiinfo", lngModel, False, , , , True) Then
        Set PatiSvrGetpatiinfo = gobjService.GetJsonListValue("output.pati_list")
    End If
End Function

Public Function PatiSvrGetVisitPatis(str����IDs As String, Optional str���￨�� As String, Optional lngModel As Long) As Collection
'  ---------------------------------------------------------------------------
'  --����:��ȡ������Ϣ
'  --��Σ�Json_In:��ʽ
'  --    input
'  --      pati_ids          C   ����IDs:����ö���
'  --      vcard_no          C   ���￨��
'  -------------------------------------------
    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""pati_ids"":""" & str����IDs & """,""vcard_no"":""" & str���￨�� & """}}"
    If CallService("Zl_Patisvr_Getvisitpatis", strJsonIn, strJsonOut, "PatiSvrGetVisitPatis", lngModel, False, , , , True) Then
        Set PatiSvrGetVisitPatis = gobjService.GetJsonListValue("output.pati_list", "pati_id")
    End If
End Function

Public Function ExseSvrGetPatisurplusinfo(str����IDs As String, Optional lngModel As Long) As Collection
'  ---------------------------------------------------------------------------
'  --����:��ȡ���˷�������Ԥ�����
'  --��Σ�Json_In:��ʽ
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    
    strJsonIn = "{""input"":{""pati_ids"":""" & str����IDs & """}}"
    If CallService("Zl_Exsesvr_Getpatisurplusinfo", strJsonIn, strJsonOut, "ExseSvrGetPatisurplusinfo", lngModel, False, , , , True) Then
        Set ExseSvrGetPatisurplusinfo = gobjService.GetJsonListValue("output.surplus_list")
    End If
End Function



Public Function ExseSvrGetremainmoney(lng����ID As Long, lng��ҳID As Long, strʣ��� As String, Optional str������ As String, Optional strԤ����� As String, Optional lngModel As Long) As Boolean
'  ---------------------------------------------------------------------------
'  --����:��ȡ����ʣ���͵�����
'  --��Σ�Json_In:��ʽ
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""pati_id"":" & lng����ID & ",""pati_pageid"":" & lng��ҳID & ",""insure_account_balance"":0}}"
    If CallService("Zl_Exsesvr_Getremainmoney", strJsonIn, strJsonOut, "ExseSvrGetremainmoney", lngModel, False, , , , True) Then
        strʣ��� = gobjService.GetJsonNodeValue("output.remain_money") & ""
        str������ = gobjService.GetJsonNodeValue("output.guarantee_money") & ""
        strԤ����� = gobjService.GetJsonNodeValue("output.expected_money") & ""
    End If
    ExseSvrGetremainmoney = True
End Function

Public Function PatiSvrGetPatiExtendInfo(lng����ID As Long, lng����ID As Long, str��Ϣֵ As String, Optional lngModel As Long) As Collection
'  ---------------------------------------------------------------------------
'  --����:��ȡ������Ϣ�ӱ�
'  --��Σ�Json_In:��ʽ
'  -------------------------------------------
    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""pati_id"":" & lng����ID & ",""visit_id"":" & lng����ID & ",""info_names"":""" & str��Ϣֵ & """}}"
    If CallService("Zl_Patisvr_Getpatiextendinfo", strJsonIn, strJsonOut, "PatiSvrGetPatiExtendInfo", lngModel, False, , , , True) Then
        Set PatiSvrGetPatiExtendInfo = gobjService.GetJsonListValue("output.slave_list")
    End If
End Function

Public Function ExseSvrGetinsureinfo(int���� As Integer, lng�շ�ϸĿID As Long, str���մ��� As String, bln������ As Boolean, Optional lngModel As Long) As Boolean
'  ---------------------------------------------------------------------------
'  --����:��ȡҽ�����������Ϣ
'  --��Σ�Json_In:��ʽ
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""insurance_type"":" & int���� & ",""fee_item_id"":" & lng�շ�ϸĿID & "}}"
    If CallService("Zl_Exsesvr_GetInsureIteminfo", strJsonIn, strJsonOut, "ExseSvrGetinsureinfo", lngModel, False, , , , True) Then
        str���մ��� = gobjService.GetJsonNodeValue("output.insure_name")
        bln������ = Val(gobjService.GetJsonNodeValue("output.isexist")) > 0
    End If
    ExseSvrGetinsureinfo = True
End Function


Public Function ExseSvrGetBillDetailInfo(str����ids As String, str����Nos As String, Optional lngModel As Long) As Collection
'  ---------------------------------------------------------------------------
'  --���ܣ���ȡ��ҩƷ��ҩҵ����صķ�����Ϣ����Ҫ���ڽ�����ʾ
'  --��Σ�json��ʽ
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    strJsonIn = "{""input"":{""fee_ids"":""" & str����ids & """,""bill_nos"":""" & str����Nos & """}}"
    If CallService("Zl_Exsesvr_GetBillDetailInfo", strJsonIn, strJsonOut, "ExseSvrGetBillDetailInfo", lngModel, False, , , , True) Then
        Set ExseSvrGetBillDetailInfo = gobjService.GetJsonListValue("output.fee_list")
    End If
End Function
 
Public Function Update���˸����־(lng�Һ�ID As Long, int�����־ As Integer) As Boolean
'  --���ܣ����ﲡ�˸����־

    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim strSQL As String
    Dim blnTrans As Boolean
    
    On Error GoTo errH
    strJsonIn = "{""input"":{""reg_id"":" & lng�Һ�ID & _
        ",""revst_sign"":" & int�����־ & "}}"
        
    strSQL = "Zl_���˾����¼_����(" & lng�Һ�ID & "," & int�����־ & ")"
    
   gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, "���Ϊ����")
    Call CallService("Zl_Exsesvr_Updateoutprevstsign", strJsonIn, strJsonOut, "Update���˸����־", p����ҽ���´�, False, , , , True)
    gcnOracle.CommitTrans: blnTrans = False
    Update���˸����־ = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Check����ȼ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long, dt�Ǽ�ʱ�� As Date, lngModel As Long) As String
'���ܣ�����ȼ����ǰ���

    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim strOut As String
    
    
    strJsonIn = "{""input"":{""pati_id"":" & lng����ID & ",""page_id"":" & lng��ҳID & ",""create_time"":""" & Format(dt�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & """}}"
    If CallService("Zl_Exsesvr_Chkpatichangenurse", strJsonIn, strJsonOut, "Check����ȼ����", lngModel, False) = False Then
        If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
            strOut = gobjService.GetJsonNodeValue("output.message")
        End If
    End If
    
    Check����ȼ���� = strOut
End Function


Public Function Checkҽ��ֹͣ(ByVal lngҽ��ID As Long, dt��ֹʱ�� As Date, ByVal int�ڲ����� As Integer, ByVal intҽʦ�ʸ� As Integer, ByVal intͣ����� As Integer, ByVal lngModel As Long) As Boolean
'���ܣ�ҽ��ֹͣǰ���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rs����ȼ� As ADODB.Recordset
    Dim strOut As String
    
    On Error GoTo errH
    
    If int�ڲ����� = 0 And (intҽʦ�ʸ� > 0 Or intͣ����� = 0) Then
        strSQL = "Select a.ҽ��״̬, a.ҽ������, a.����id, a.��ҳid, Nvl(a.Ӥ��, 0) as Ӥ��, Nvl(a.�������, '*') As �������, b.��������" & vbNewLine & _
                "From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
                "Where a.������Ŀid = b.Id(+) And a.Id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Checkҽ��ֹͣ", lngҽ��ID)
        
        If Not rsTmp.EOF Then
            If rsTmp!������� & "" = "H" And rsTmp!�������� & "" = "1" And Val(rsTmp!Ӥ�� & "") = 0 Then
                '��ȡ���˻���ȼ�id
                strSQL = "Select c.�շ�ϸĿid" & vbNewLine & _
                        "From ����ҽ����¼ A, ����ҽ���Ƽ� C, �շ���ĿĿ¼ D" & vbNewLine & _
                        "Where a.Id = c.ҽ��id And c.�շ�ϸĿid = d.Id And d.��� = 'H' And Nvl(d.��Ŀ����, 0) <> 0 And a.Id = Id_In And Rownum = 1 And" & vbNewLine & _
                        "      Exists (Select 1 From ������ҳ Where ����id =[1] And ��ҳid = [2] And ����ȼ�id = c.�շ�ϸĿid)"
                Set rs����ȼ� = zlDatabase.OpenSQLRecord(strSQL, "Checkҽ��ֹͣ", Val(rsTmp!����ID & ""), Val(rsTmp!��ҳID & ""))
                
                If Not rs����ȼ� Is Nothing Then
                    If Not rs����ȼ�.EOF Then
                        If Val(rs����ȼ�!�շ�ϸĿid & "") <> 0 Then
                             strOut = Check����ȼ����(Val(rsTmp!����ID & ""), Val(rsTmp!��ҳID & ""), dt��ֹʱ��, lngModel)
                             If strOut <> "" Then
                                MsgBox strOut, vbInformation, gstrSysName
                                Exit Function
                             End If
                        End If
                    End If
                End If
                
            End If
        End If
    End If
    Checkҽ��ֹͣ = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Checkҽ��У��(ByVal lngModel As Long, ByVal lngҽ��ID As Long, int״̬ As Integer, ByVal dtУ��ʱ�� As Date, Optional ByVal strУ��˵�� As String, Optional ByVal int�Զ�У�� As Integer _
                    , Optional ByVal str����Ա��� As String, Optional ByVal str����Ա���� As String, Optional strԤԼ��ԺSQL As String) As Boolean
'���ܣ�ҽ��У��ǰ���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rs����ȼ� As ADODB.Recordset
    Dim rsTmp1 As ADODB.Recordset
    Dim strOut As String
    Dim colPi As Collection
    Dim colPati As Collection
    
    On Error GoTo errH
    
    strԤԼ��ԺSQL = ""
    If int״̬ = 3 Then
        strSQL = "Select a.ҽ����Ч, a.ҽ��״̬, a.����ʱ��, a.����ҽ��, a.��ʼִ��ʱ��, a.����id, a.��ҳid, a.Ӥ��, a.ҽ������, a.�������, a.������Ŀid, a.ǰ��id," & vbNewLine & _
                    "       Nvl(b.��������, '0') as ��������, Nvl(a.ִ�б��, 0), a.ִ�п���id, a.�걾��λ, a.��������id, Nvl(a.������־, 0) As ������־, a.���˿���id" & vbNewLine & _
                    "From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
                    "Where a.������Ŀid = b.Id(+) And a.Id = [1]"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Checkҽ��ֹͣ", lngҽ��ID)
        
        If Not rsTmp.EOF Then
            If rsTmp!������� & "" = "H" And rsTmp!�������� & "" = "1" And Val(rsTmp!ҽ����Ч & "") = 0 Then
                '��ȡ���˻���ȼ�id
                strSQL = "Select b.id as �շ�ϸĿID" & vbNewLine & _
                        "From �����շѹ�ϵ A, �շ���ĿĿ¼ B" & vbNewLine & _
                        "Where a.�շ���Ŀid = b.Id And b.��� = 'H' And Nvl(b.��Ŀ����, 0) <> 0 And a.������Ŀid = [3] And Rownum = 1 And" & vbNewLine & _
                        "      Not Exists" & vbNewLine & _
                        " (Select 1 From ������ҳ Where ����id = [1] And ��ҳid = [2] And ����ȼ�id = a.�շ���Ŀid)"

                Set rs����ȼ� = zlDatabase.OpenSQLRecord(strSQL, "Checkҽ��ֹͣ", Val(rsTmp!����ID & ""), Val(rsTmp!��ҳID & ""), Val(rsTmp!������ĿID & ""))
                
                If Not rs����ȼ� Is Nothing Then
                    If Not rs����ȼ�.EOF Then
                        If Val(rs����ȼ�!�շ�ϸĿid & "") <> 0 Then
                             strOut = Check����ȼ����(Val(rsTmp!����ID & ""), Val(rsTmp!��ҳID & ""), CDate(rsTmp!��ʼִ��ʱ�� & ""), lngModel)
                             If strOut <> "" Then
                                MsgBox strOut, vbInformation, gstrSysName
                                Exit Function
                             End If
                        End If
                    End If
                End If
            ElseIf rsTmp!������� & "" = "Z" And rsTmp!�������� & "" = "2" Then
                '�����۲����´���Ժ֪ͨ;
                'ԤԼ�Ǽǵ�������1.��ǰ��ԤԼ,2.��ǰ���������۲��ˣ���ԺʱҲ������Ϊ��Ҫ��ԤԼ,��Ժ����ʱ����˱����Ժ����ܽ��գ�
                strSQL = "Select Count(*) as Count From ������ҳ Where ����id = [1] And Nvl(��ҳid, 0) = 0"
                Set rsTmp1 = zlDatabase.OpenSQLRecord(strSQL, "Checkҽ��ֹͣ", Val(rsTmp!����ID & ""))
                If Val(rsTmp!Count & "") = 0 Then
                    strSQL = "Select Count(*) Into v_Count From ������ҳ Where ����id = [1] And ��ҳid = [2] And �������� <> 1"
                    Set rsTmp1 = zlDatabase.OpenSQLRecord(strSQL, "Checkҽ��ֹͣ", Val(rsTmp!����ID & ""), Val(rsTmp!��ҳID & ""))
                    If Val(rsTmp!Count & "") = 0 Then
                        'Zl_��Ժ������ҳ_Insert
                        If Not CallService("Zl_Patisvr_Lockcheck", "{""input"":{""pati_id"":" & Val(rsTmp!����ID & "") & "}}", , "Checkҽ��ֹͣ", lngModel, True) Then
                            Exit Function
                        End If
                        
                        '��ȡסԺԤԼ��Ժ��SQL
                        Set colPi = New Collection
                        Set colPati = PatiSvrGetpatiinfo(3, Val(rsTmp!����ID & ""), lngModel)
                        colPi.Add 0, "_del_page"
                        colPi.Add 0, "_del_in"
                        colPi.Add str����Ա����, "_operator_name"
                        colPi.Add str����Ա���, "_operator_code"
                        colPi.Add IIF(Val(rsTmp!������־ & "") = 1, "����", ""), "_in_type"
                        colPi.Add Format(rsTmp!��ʼִ��ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), "_create_time"
                        colPi.Add rsTmp!����ҽ�� & "", "_out_doctor"
                        colPi.Add Format(rsTmp!��ʼִ��ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), "_in_date"
                        colPi.Add 0, "_rgst_id"
                        colPi.Add 0, "_inpatient_num"
                        colPi.Add 0, "_keep_num"
                        colPi.Add 1, "_status"
                        colPi.Add 0, "_pati_nature"
                        colPi.Add 0, "_again_in"
                        colPi.Add Val(rsTmp!��������ID & ""), "_pati_deptid"
                        
                        Call GetSQLסԺԤԼ�Ǽ�(colPati, colPi, strԤԼ��ԺSQL)
                    End If
                End If
            End If
        End If
    End If
    Checkҽ��У�� = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSQLסԺԤԼ�Ǽ�(colPati As Collection, colPi As Collection, strAPage As String) As Boolean
'����:סԺ����ҽ����ԤԼ��Ժ�Ǽ�
'����:colPati ���˻�����Ϣ
'     colPi ��Ժ������Ϣ
'����:strAPage ���˰�����[Zl_������ҳ_ԤԼ��Ժ�Ǽ�]���Json��ʽ
    Dim colP As Collection
    On Error GoTo errH
    
    strAPage = ""
                    
    Set colP = colPati(1)

    strAPage = strAPage & "(" & Val(GetColVal(colP, "_pati_id"))
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_del_page"))
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_del_in"))
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_rgst_id"))
    strAPage = strAPage & ",""" & GetColVal(colPi, "_inpatient_num") & """"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_name") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_sex") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_age") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_emp_phno") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_emp_postcode") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_fee_category") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_emp_addr") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_country_name") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pat_hous_addr") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pat_hous_postcode") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_marital_cstatus") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pat_home_addr") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pat_home_postcode") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pat_home_phno") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_contacts_addr") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_contacts_phno") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_contacts_relation") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_contacts_idcard") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_contacts_name") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_area") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_pati_education") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_mdlpay_mode_name") & "'"
    strAPage = strAPage & ",'" & GetColVal(colP, "_ocpt_name") & "'"
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_insurance_type"))
    strAPage = strAPage & ",To_Date(" & GetColVal(colPi, "_in_date") & ", 'yyyy-mm-dd hh24:mi:ss')"
    strAPage = strAPage & ",'" & GetColVal(colPi, "_out_doctor") & "'"
    strAPage = strAPage & ",To_Date(" & GetColVal(colPi, "_create_time") & ", 'yyyy-mm-dd hh24:mi:ss')"
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_pati_nature"))
    strAPage = strAPage & ",'" & GetColVal(colP, "_in_type") & "'"
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_again_in"))
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_status"))
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_keep_num"))
    strAPage = strAPage & ",'" & GetColVal(colPi, "_operator_code") & "'"
    strAPage = strAPage & "," & ZVal(GetColVal(colPi, "_pati_deptid"))
    strAPage = strAPage & ",'" & GetColVal(colPi, "_operator_name") & "')"
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Getҽ������(ByVal int���� As Integer, ByVal lng�շ�ϸĿID As Long, Optional ByRef colBat As Collection, Optional ByVal lngType As Long) As String
'���ܣ���ȡָ���еķ�������
'   int����+lng�շ�ϸĿID ������Ϊ0ʱ��ȡ��ǰ��Ŀ����������
'   colBat �������Ҳ�ǳ��Σ���δ�����Ҫ������ȡ�� �շ�ϸĿID����ƴ��[colBat(1)]������ �շ�ϸĿID,�������ƣ���ά����
'   lngType �������ͣ�0-������ȡҽ����������,���ؼ��ϣ�1-������ȡ�Ƿ�������֧����Ŀ�������ַ���

    Dim strJsonIn As String, str���� As String
    
    On Error GoTo errH
    
    If lng�շ�ϸĿID <> 0 And int���� <> 0 Then
        strJsonIn = "{""input"":{""insurance_type"":" & int���� & ",""fee_item_id"":" & lng�շ�ϸĿID & "}}"
        Call CallService("Zl_ExseSvr_GetInsureItemInfo", strJsonIn, , , , False, , , , True)
        str���� = gobjService.GetJsonNodeValue("output.insure_name") & ""
        Getҽ������ = IIF(str���� <> "", "ҽ������:" & str����, "")
    End If
    
    If Not colBat Is Nothing Then
        If colBat.Count > 0 Then
            strJsonIn = "{""input"":{""insurance_type"":" & int���� & ",""fee_item_ids"":""" & colBat(1) & """,""query_type"":" & lngType & "}}"
            Call CallService("Zl_ExseSvr_GetInsureItemInfo", strJsonIn, , , , False, , , , True)
            
            If lngType = 0 Then
                Set colBat = gobjService.GetJsonListValue("output.item_list", "fee_item_id")
            Else
                Getҽ������ = gobjService.GetJsonNodeValue("output.fee_item_ids") & ""
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillChareApp(colList As Collection, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str����ԭ�� As String, ByVal strTitle As String, ByVal lngģ�� As Long) As Boolean
'����:���������������룬סԺҽ�����˼��ҽ����ʱ���Զ�����ҩƷ�÷��õ���������
'����:colList ���������б�,colList������Դ:Zl_Exsesvr_CheckOrderRoll�е�charge_list[],ҽ�����˷��͵�ʱ��
'     lng����ID
'     lng��ҳID
'     str����ԭ��
'--ZL_ExseSvr_billCharge,���ò�����������
'--  input
'--    request_operator                C 1 ������
'--    request_time                    C 1 ����ʱ��
'--    request_type                    N 1 �������
'--    del_tag                         N 1 ɾ����־
'--    reason                          C 1 ����ԭ��
'--    item_list[]������������������б�
'--        fee_id                      N 1 ����ID
'--        request_dept_id             N 1 �����������ID
'--        fee_item_id                 N 1 �շ�ϸĿID
'--        quantity                    N 1 ����
'--        audit_dept_id               N 1 ��˲���ID
    Dim lng��Ժ����ID As Long, lng��˱�־ As Long, lng״̬ As Long
    Dim lng��˲���ID As Long, i As Long
    Dim strIt As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strJsonIn As String
    Dim colData As Collection
    Dim colRow As Collection
    Dim datCur As Date
    Dim strTime As String
    
    On Error GoTo errH
    
    strIt = ""
    strSQL = "select a.��Ժ����id,a.��˱�־,a.״̬ from ������ҳ a where a.����id=[1] and a.��ҳid=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, strTitle, lng����ID, lng��ҳID)
    lng��Ժ����ID = Val("" & rsTmp!��Ժ����ID)
    lng��˱�־ = Val("" & rsTmp!��˱�־)
    lng״̬ = Val("" & rsTmp!״̬)
    For i = 1 To colList.Count '��˲���id/ɾ����־����  del_tag / audit_dept_id
        strIt = strIt & ",{""fee_id"":" & colList(i)("_fee_id") & _
            ",""request_dept_id"":" & colList(i)("_request_dept_id") & _
            ",""item_id"":" & colList(i)("_fee_item_id") & ",""request_type"":1" & _
            ",""request_num"":" & GetJsonNum(colList(i)("_quantity") & "") & ",""sended_num"":0" & _
            ",""out_depti_id"":" & lng��Ժ����ID & ",""audit_sign"":" & lng��˱�־ & ",""inp_state"":" & lng״̬ & _
            "}"
    Next
    strJsonIn = "{""input"":{""item_list"":[" & Mid(strIt, 2) & "]}}"
    Call CallService("Zl_ExseSvr_CheckBillChargeOff", strJsonIn, , strTitle, lngģ��, False, , , , True)
    Set colData = gobjService.GetJsonListValue("output.item_list", "fee_id")
    
    strIt = ""
    For i = 1 To colList.Count '��˲���id/ɾ����־����  del_tag / audit_dept_id
        strIt = strIt & ",{""fee_id"":" & colList(i)("_fee_id") & _
            ",""request_dept_id"":" & colList(i)("_request_dept_id") & _
            ",""fee_item_id"":" & colList(i)("_fee_item_id") & _
            ",""quantity"":" & GetJsonNum(colList(i)("_quantity") & "")
        
        
        Set colRow = GetColObj(colData, "_" & colList(i)("_fee_id"))
        lng��˲���ID = GetColVal(colRow, "_audit_dept_id", "N")
        If lng��˲���ID <> 0 Then
            strIt = strIt & ",""audit_dept_id"":" & lng��˲���ID
        End If
        
        strIt = strIt & "}"
    Next
     
    datCur = zlDatabase.Currentdate
    strTime = Format(datCur, "yyyy-MM-dd HH:mm:ss")
    strJsonIn = "{""input"":{""request_operator"":""" & UserInfo.���� & """" & _
            ",""request_time"":""" & strTime & """" & _
            ",""request_type"":1" & _
            ",""reason"":""" & str����ԭ�� & """" & _
            ",""item_list"":[" & Mid(strIt, 2) & "]" & _
            "}}"
    Call CallService("Zl_ExseSvr_BillCharge", strJsonIn, , strTitle, lngģ��, False, , , , True)
    
    BillChareApp = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function BillChareAppPivas(ByVal strPivas As String, ByVal strExse As String, ByVal lng������� As Long, ByVal lng��ҺID As Long, ByVal lng�������ID As Long, ByVal blnAutoAduit As Boolean, ByVal str����ԭ�� As String, ByVal strTitle As String, ByVal lngģ�� As Long) As Boolean
'����:��ʿվ��Һ��������
'����:colList ���������б�,colList������Դ:Zl_Exsesvr_CheckOrderRoll�е�charge_list[],ҽ�����˷��͵�ʱ��
'       strPivas ��Һ�����б�
'                strJsonIn = "{""input"":{""query_type"":0,""pivas_id"":" & Val(.RowData(.Row)) & "}}"
'                Call CallService("Zl_PivasSvr_GetPivasContent", strJsonIn, strPivasOutPar, Me.Caption, , False, , , , True)
'       strExse  ���÷�����Σ�
'                strJsonIn = "{""input"":{""oper_type"":1,""fee_source"":" & IIF(lng������� = 1, 1, 2) & ",""fee_ids"":""" & str����ids & """}}"
'                Call CallService("Zl_ExseSvr_CheckBillChargeOff", strJsonIn, strExseOutPar, Me.Caption, , False, , , , True)
'       blnAutoAduit �Ƿ��Զ����,true �����Զ���ˣ�false�����Զ���ˣ���Ҫ��ȡ�ѷ�����
'       str����ԭ�� ��������ԭ��

'     lng����ID
'     lng��ҳID
'     str����ԭ��
'--ZL_ExseSvr_billCharge,���ò�����������
'--  input
'--    request_operator                C 1 ������
'--    request_time                    C 1 ����ʱ��
'--    request_type                    N 1 �������
'--    del_tag                         N 1 ɾ����־
'--    reason                          C 1 ����ԭ��
'--    item_list[]������������������б�
'--        fee_id                      N 1 ����ID
'--        request_dept_id             N 1 �����������ID
'--        fee_item_id                 N 1 �շ�ϸĿID
'--        quantity                    N 1 ����
'--        audit_dept_id               N 1 ��˲���ID
    
    Dim colPivas As Collection
    Dim colFee As Collection
    Dim strFees As String
    Dim lng����ID As Long, lng��ҳID As Long
    Dim strJsonIn As String
    Dim colDrug As Collection
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng��Ժ����ID As Long, lng��˱�־ As Long, lng״̬ As Long
    Dim lng��˲���ID As Long, i As Long
    Dim colTmp As New Collection
    Dim colRow As Collection
    Dim str�ѷ��� As String
    Dim colEO As Collection
    Dim colRowOther As Collection
    Dim strIt As String
    Dim strTime As String
    Dim strParExseCharge  As String
    Dim strParPivasCharge  As String
    Dim blnTran As Boolean
    Dim str��� As String
    Dim strDelDrug As String
    Dim strDelDrugPivas As String
    Dim lng�Զ���� As Long ' 1-���������Զ���ˣ�2-��ͨ��������ģʽ
    Dim strPatiList As String
    Dim lngҽ��ID As Long, lng���ͺ� As Long, strSQLAdd As String, strSQLCls As String
    On Error GoTo errH
    
    '�Ƚ��������չ���
    Call InitSvr
    Call gobjService.SetJsonString(strPivas)
    Set colPivas = gobjService.GetJsonListValue("output.item_list", "rcpdtl_id")
    strFees = gobjService.GetJsonNodeValue("output.fee_ids") & ""
    Call gobjService.SetJsonString(strExse)
    Set colFee = gobjService.GetJsonListValue("output.charge_list", "fee_id")
    lng����ID = colFee(1)("_pati_id")
    lng��ҳID = colFee(1)("_pati_pageid")
    
    strSQL = "select a.��Ժ����id,a.��˱�־,a.״̬ from ������ҳ a where a.����id=[1] and a.��ҳid=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, strTitle, lng����ID, lng��ҳID)
    lng��Ժ����ID = Val("" & rsTmp!��Ժ����ID)
    lng��˱�־ = Val("" & rsTmp!��˱�־)
    lng״̬ = Val("" & rsTmp!״̬)
    '������Ϣ�б�
    strPatiList = "{""pati_id"":" & lng����ID & ",""pati_dept_id"":" & lng��Ժ����ID & ",""fee_audit_status"":" & lng��˱�־ & ",""si_inp_status"":" & lng״̬ & "}"
    If Not blnAutoAduit Then
        '�����Զ���ˣ���Ҫ��ȡ�ѷ�����
        strJsonIn = "{""input"":{""rcpdtl_ids"":""" & strFees & """}}"
        Call CallService("Zl_Drugsvr_Getexecutednum", strJsonIn, , strTitle, lngģ��, False, , , , True)
        Set colDrug = gobjService.GetJsonListValue("output.data", "rcpdtl_id")
    End If
    
    For i = 1 To colPivas.Count
        Set colRow = New Collection
 
        colRow.Add lng��ҺID & "", "_pivas_id"
        colRow.Add colPivas(i)("_rcp_no") & "", "_fee_no"
        colRow.Add colPivas(i)("_rcpdtl_id") & "", "_fee_id"
        colRow.Add lng�������ID & "", "_request_dept_id"
        colRow.Add colPivas(i)("_drug_id") & "", "_item_id"
        colRow.Add IIF(blnAutoAduit, "0", "1"), "_request_type"
        colRow.Add colPivas(i)("_send_num") & "", "_request_num"
        
        If lngҽ��ID = 0 Then
            lngҽ��ID = colPivas(i)("_order_id")
            lng���ͺ� = colPivas(i)("_send_no")
        End If
        
        str��� = ""
        If Not blnAutoAduit Then
            Set colRowOther = GetColObj(colDrug, "_" & colPivas(i)("_rcpdtl_id"))
            str�ѷ��� = GetColVal(colRowOther, "_sended_num", "N")
        Else
            str�ѷ��� = 0
            
            Set colRowOther = GetColObj(colFee, "_" & colPivas(i)("_rcpdtl_id"))
            str��� = GetColVal(colRowOther, "_serial_num", "N")
            str��� = str��� & ":" & colPivas(i)("_send_num") & ":0" '�����Զ����,ִ��״̬������0
        End If
        
        colRow.Add str�ѷ���, "_sended_num"
        
        colRow.Add str���, "_serial_num" '�������

        colTmp.Add colRow
    Next
    strIt = ""
    For i = 1 To colTmp.Count '�����б�
        strIt = strIt & ",{"
        strIt = strIt & """fee_id"":" & colTmp(i)("_fee_id")
        strIt = strIt & ",""request_dept_id"":" & colTmp(i)("_request_dept_id")
        strIt = strIt & ",""item_id"":" & colTmp(i)("_item_id")
        strIt = strIt & ",""request_type"":" & colTmp(i)("_request_type")
        strIt = strIt & ",""request_num"":" & GetJsonNum("" & colTmp(i)("_request_num"))
        strIt = strIt & ",""sended_num"":" & GetJsonNum("" & colTmp(i)("_sended_num"))

        strIt = strIt & "}"
    Next
    
    strJsonIn = "{""input"":{""oper_type"":0,""item_list"":[" & Mid(strIt, 2) & "],""pati_list"":[" & strPatiList & "]}}"
    Call CallService("Zl_ExseSvr_CheckBillChargeOff", strJsonIn, , strTitle, lngģ��, False, , , , True)
    Set colEO = gobjService.GetJsonListValue("output.item_list", "fee_id")
    
    strIt = ""
    For i = 1 To colTmp.Count
        strIt = strIt & ",{"
        strIt = strIt & """fee_id"":" & colTmp(i)("_fee_id")
        strIt = strIt & ",""request_dept_id"":" & colTmp(i)("_request_dept_id")
        strIt = strIt & ",""fee_item_id"":" & colTmp(i)("_item_id")
        strIt = strIt & ",""request_type"":" & colTmp(i)("_request_type")
        strIt = strIt & ",""quantity"":" & GetJsonNum("" & colTmp(i)("_request_num"))
        Set colRowOther = GetColObj(colEO, "_" & colTmp(i)("_fee_id"))
        lng��˲���ID = GetColVal(colRowOther, "_audit_dept_id", "N")
        If lng��˲���ID <> 0 Then
            strIt = strIt & ",""audit_dept_id"":" & lng��˲���ID
        End If
        
        If blnAutoAduit Then
            strIt = strIt & ",""auto_aduit"":1"
            strIt = strIt & ",""outpati_account"":" & lng�������
            strIt = strIt & ",""fee_no"":""" & colTmp(i)("_fee_no") & """"
            strIt = strIt & ",""serial_num"":""" & colTmp(i)("_serial_num") & """"
        End If
        strIt = strIt & "}"
        
        '��������Զ���˻���Ҫɾ��ҩƷ����
        If blnAutoAduit Then
            strDelDrug = strDelDrug & ",{""rcpdtl_id"":" & colTmp(i)("_fee_id") & _
            ",""chargeoffs_num"":" & GetJsonNum("" & colTmp(i)("_request_num")) & _
            ",""dispensing_ids"":""" & colTmp(i)("_pivas_id") & """" & _
            "}"
        End If
    Next
    
    strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    strJsonIn = "{""input"":{""request_operator"":""" & UserInfo.���� & """,""request_code"":""" & UserInfo.��� & """,""request_time"":""" & strTime & """,""request_type"":" & IIF(blnAutoAduit, "0", "1") & _
        ",""del_tag"":2,""reason"":""" & str����ԭ�� & """,""item_list"":[" & Mid(strIt, 2) & "]}}"
    strParExseCharge = strJsonIn 'Zl_Exsesvr_BillChargeOff�����������ƴ�����
    
    strJsonIn = "{""input"":{""pivas_ids"":""" & lng��ҺID & """,""operator_status"":9,""operator_name"":""" & UserInfo.���� & """,""operator_time"":""" & strTime & """" & _
    ",""auto_aduit"":" & IIF(blnAutoAduit, 1, 0) & ",""operator_notes"":""" & str����ԭ�� & """}}"
    strParPivasCharge = strJsonIn '������������ Zl_Pivassvr_Statusupdate
    
    If strDelDrug <> "" Then
        strJsonIn = """pivas_list"":[{""pivas_ids"":""" & lng��ҺID & """,""operator_name"":""" & UserInfo.���� & """,""operator_time"":""" & strTime & """,""reason"":""" & str����ԭ�� & """}]"
        strParPivasCharge = ""
        strDelDrugPivas = "{""input"":{""item_list"":[" & Mid(strDelDrug, 2) & "]," & strJsonIn & "}}"
    End If
    lng�Զ���� = 2
    If strDelDrugPivas <> "" Then lng�Զ���� = 1
    
    strSQLAdd = "Zl_����ҽ������_����ͬ����־(9,'" & lng���ͺ� & "','" & lngҽ��ID & "','" & UserInfo.���� & "','" & OS.ComputerName & "'," & lng��ҺID & ")"
    strSQLCls = "Zl_����ҽ������_����ͬ����־(10,'" & lng���ͺ� & "','" & lngҽ��ID & "','" & UserInfo.���� & "','" & OS.ComputerName & "'," & lng��ҺID & ")"
    
    '��������Զ���˾Ϳ���ִ����,����,�Ⱦ�������,�ٷ�������,ʵ���Ͽ�������û��,�ȱ������ڲ���
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(strSQLCls, strTitle)
    Call zlDatabase.ExecuteProcedure(strSQLAdd, strTitle)
    If strDelDrugPivas <> "" Then
        Call CallService("Zl_Drugsvr_Delrecipebill", strDelDrugPivas, , strTitle, lngģ��, False, , , , True)
    End If
    If strParPivasCharge <> "" Then
        Call CallService("Zl_Pivassvr_Statusupdate", strParPivasCharge, , strTitle, lngģ��, False, , , , True)
    End If
    gcnOracle.CommitTrans: blnTran = False
    
    
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(strSQLCls, strTitle)
    Call CallService("Zl_Exsesvr_BillChargeOff", strParExseCharge, , strTitle, lngģ��, False, , , , True)
    gcnOracle.CommitTrans: blnTran = False
    
    BillChareAppPivas = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExseSvrAdviceisexist(strҽ��IDs As String, Optional lngModel As Long) As Collection
'  ---------------------------------------------------------------------------
'  --����:����ҽ��ID��ѯ�Ƿ��ڷ��ñ���ڼ�¼
'  --��Σ�Json_In:��ʽ
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String
    
    
    strJsonIn = "{""input"":{""advice_ids"":""" & strҽ��IDs & """}}"
    If CallService("Zl_Exsesvr_Adviceisexist", strJsonIn, strJsonOut, "ExseSvrAdviceisexist", lngModel, False, , , , True) Then
        Set ExseSvrAdviceisexist = gobjService.GetJsonListValue("output.advice_list")
    End If
End Function


Public Function ExseSvrGetnextno(int��� As Integer, Optional lng����ID As Long, Optional lngModel As Long, Optional ByVal lng���� As Long, Optional ByRef varNos As Variant) As String
'  ---------------------------------------------------------------------------
'  --����:�����ض���������µĺ���
'  --��Σ�Json_In:��ʽ
'  -------------------------------------------
'lng���� һ��������ȡ������ݺ�,������ֵ>1ʱ����һ������varNos
    Dim strTmp As String
    Dim strJsonIn As String
    Dim strJsonOut As String
    
    If lng���� <= 1 Then
        lng���� = 0
    End If
    
    strJsonIn = "{""input"":{""item_num"":" & int��� & ",""dept_id"":" & ZVal(lng����ID) & ",""quantity"":" & lng���� & "}}"
    If CallService("Zl_Exsesvr_Getnextno", strJsonIn, strJsonOut, "ExseSvrGetnextno", lngModel, False, , , , True) Then
        If lng���� > 1 Then
            strTmp = gobjService.GetJsonNodeValue("output.next_no")
        Else
            strTmp = gobjService.GetJsonNodeValue("output.next_no")
            ExseSvrGetnextno = strTmp
        End If
        varNos = Split(strTmp, ",")
    End If
End Function

Public Function ExseSvrGetSpecCalcFeeItem(Optional lngModel As Long) As Collection
'  -------------------------------------------
'  --���ܣ���ȡ���۷ѱ���ϸ�б�
'  --��Σ�Json_In:��ʽ
'  -------------------------------------------

    Dim strJsonIn As String
    Dim strJsonOut As String

    strJsonIn = ""
    If CallService("zl_ExseSvr_GetSpecCalcFeeItem", strJsonIn, strJsonOut, "ExseSvrGetSpecCalcFeeItem", lngModel, False, , , , True) Then
        Set ExseSvrGetSpecCalcFeeItem = gobjService.GetJsonListValue("output.feecategory_list")
    End If
End Function

Public Function ItemHaveCash(ByVal int������Դ As Integer, ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng���ID As Long, _
    ByVal lng���ͺ� As Long, ByVal str��� As String, ByVal str���ݺ� As String, ByVal int��¼���� As Integer, ByVal int������� As Integer, ByVal int��ʽ As Integer, _
    Optional ByVal blnMove As Boolean, Optional ByRef blnIsAbnormal As Boolean) As Boolean
'���ܣ��жϵ�ǰ��ִ��ҽ���Ƿ����շѻ���ʻ��۵��Ƿ������
'������int������Դ=1-����,2-סԺ
'      str���=����������ڴ�һ��ҽ�������ַֿ�ִ�е�����
'      int��ʽ=0-����Ƿ����δ�շѼ�¼
'              1-����Ƿ�������շѼ�¼
'      int�������=1=סԺ���͵��������
'      ���أ�strҽ��IDs=��ҽ������ص�ҽ��ID,NOs=ҽ�����͵ĵ��ݺźͲ��ĸ����еĵ��ݺ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strNoFilter As String
    Dim strOrders As String
    Dim lng��ID As Long
    Dim lngFO As Long
    Dim i As Long
    
    On Error GoTo errH
    
    If int������Դ = 2 And int��¼���� = 2 And int������� = 0 Then
        lngFO = 2
    Else
        lngFO = 1
    End If
    strOrders = lngҽ��ID
    strNoFilter = str���ݺ�
    
    If Not bln����ִ�� Then
        lng��ID = IIF(lng���ID <> 0, lng���ID, lngҽ��ID)
        strSQL = "Select a.ҽ��id, a.No" & vbNewLine & _
            "From ����ҽ����¼ C, ����ҽ������ A" & vbNewLine & _
            "Where a.ҽ��id In (Select ID From ����ҽ����¼ Where (ID =[1] Or ���id =[1]) And ������� =[2]) And a.���ͺ� =[3] And" & vbNewLine & _
            "      a.ҽ��id = c.Id  and c.�������=[2] " & vbNewLine & _
            "group by a.ҽ��id,a.no"
        If blnMove Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISService", lng��ID, str���, lng���ͺ�)
        
        For i = 1 To rsTmp.RecordCount
            If InStr("," & strOrders & ",", "," & rsTmp!ҽ��ID & ",") = 0 Then
                strOrders = IIF(strOrders = "", "", strOrders & ",") & rsTmp!ҽ��ID
            End If
            If InStr("," & strNoFilter & ",", "," & rsTmp!NO & ",") = 0 Then
                strNoFilter = IIF(strNoFilter = "", "", strNoFilter & ",") & rsTmp!NO
            End If
            rsTmp.MoveNext
        Next
    End If
     
    strSQL = "{""input"":{""fee_origin"":" & lngFO & ",""order_ids"":""" & strOrders & """,""fee_nos"":""" & strNoFilter & """,""oper_type"":" & int��ʽ & "}}"
    Call CallService("Zl_ExseSvr_GetOrderChargedInfo", strSQL, , "mdlCISService", , False, , , , True)
    
    '�Ƿ����쳣����
    If gobjService.GetJsonNodeValue("output.blance_sign") & "" = "1" Then
        blnIsAbnormal = True
    Else
        blnIsAbnormal = False
    End If
    
    If gobjService.GetJsonNodeValue("output.isexist") & "" = "1" Then
        ItemHaveCash = True
    End If
  
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetListNodeTxt(ByVal strJsonIn As String, ByVal strListName As String, Optional ByVal blnItem As Boolean) As String
'���ܣ���ָ���ķ��������Json���л�ȡĳ��Json������ֵ��������������Ŵ�ͷ��
'     ����˵����ZLHIS������̳��εı�׼���,Ҫ����ȡ������Ϊ��һ�� ��:output.item_list
'����: blnItem false-�������,true ֻҪԪ�ش�
'����:json���е�ĳ��Ƭ��,���������Ԫ�ؽ�,��ȡԪ��ʱ���ؿմ�
'����Ч������:
'        ,"charge_list":[{"fee_id":172780,"fee_item_id":10831},{"fee_id":172781,"fee_item_id":12391}]
'        ,{"fee_id":172780,"fee_item_id":10831},{"fee_id":172781,"fee_item_id":12391}

    Dim strTmp As String
    Dim objJson As New zl9ComLib.clsJson
    Dim strJsonPart As String
    
    On Error GoTo errH
    
    objJson.JSON = strJsonIn
    strJsonPart = objJson.PathItem("output." & strListName).ItemJSON
    
    If blnItem Then
        strJsonPart = Mid(strJsonPart, 2, Len(strJsonPart) - 2)
        If strJsonPart <> "" Then
            strTmp = "," & strJsonPart
        End If
    Else
        strTmp = ",""" & strListName & """:" & strJsonPart
    End If
    
    GetListNodeTxt = strTmp
    Exit Function
errH:
    err.Clear
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub Exe�Զ���ҩ����(ByVal strReStuff As String, ByVal strClsS As String, ByVal strReDrug As String, ByVal strClsD As String, ByVal strTitle As String, ByVal lngģ�� As Long)
'���ܣ��Զ���ҩ����
'˵�������������Բ������Ѿ����ˣ�����Ҫ���Դ���

    On Error GoTo errH
    
    If strReStuff <> "" Then  '�Զ��˲�
        Call CallService("Zl_Stuffsvr_Autoreturnstuff", strReStuff, , strTitle, lngģ��, False, , , , True)
        Call CallService("Zl_Exsesvr_Updateexeinfo", strClsS, , strTitle, lngģ��, False, , , , True)
    End If
    
    If strReDrug <> "" Then  '�Զ���ҩ
        Call CallService("Zl_Drugsvr_Autoreturndrug", strReDrug, , strTitle, lngģ��, False, , , , True)
        Call CallService("Zl_Exsesvr_Updateexeinfo", strClsD, , strTitle, lngģ��, False, , , , True)
    End If
    
    Exit Sub
errH:
    err.Clear
End Sub

Public Function Getҽ������ִ�й���SQL(ByVal strCisOutPar As String, ByVal strTime As String, ByRef strSQL As String) As Boolean
    Dim lng��ҽ��ID As Long
    Dim colPage As Collection, col_cis_data_list As Collection
    Dim colCISData As Collection
    Dim lng����ID As Long
    Dim lngɾ��Ժ As Long
    Dim lngɾ��ҳ As Long
    
    
    On Error GoTo errH
    
    strSQL = ""
    
'    CREATE OR REPLACE PROCEDURE ZL_����ҽ����¼_����_S
'(
'  ����id_In      In Number,
'  ��ҽ��id_In    In Number,
'  ɾ��Ժ_In      In Number,
'  ɾ��ҳ_In      In Number,
'  ����Ա����_In  In Varchar2,
'  ����Ա���_In  In Varchar2,
'  ����ʱ��_In    In Varchar2,
'  ����ȼ�ҽ��id In Number,
'  ���˱䶯_In    In Number,
'  ����ȼ�ͣ_In  In Number,
'  ���δ�ӡ_In    In Number,
'  Ӥ�����_In    In Number,
'  ��ҳid_In      In Number

    Call GetSvrOutInfo(strCisOutPar)
    Set colPage = gobjService.GetJsonListValue("output.page_list")
    Set col_cis_data_list = gobjService.GetJsonListValue("output.cis_data_list")
    Set colCISData = col_cis_data_list(1)
    If Not colPage Is Nothing Then
        If colPage.Count > 0 Then
            lngɾ��Ժ = Val(colPage(1)("_del_in"))
            lngɾ��ҳ = 1
        End If
    End If
    lng��ҽ��ID = Val(GetColVal(colCISData, "_main_order_id"))
    lng����ID = Val(gobjService.GetJsonNodeValue("output.order_info_list[0].pati_id") & "")
     
    strSQL = "ZL_����ҽ����¼_����_S("
    strSQL = strSQL & ZVal(Val(GetColVal(colCISData, "_pati_id")))  '  ����id_In      In Number,
    strSQL = strSQL & "," & lng��ҽ��ID '  ��ҽ��id_In    In Number,
    strSQL = strSQL & "," & ZVal(lngɾ��Ժ) '  ɾ��Ժ_In      In Number,
    strSQL = strSQL & "," & ZVal(lngɾ��ҳ) '  ɾ��ҳ_In      In Number,
    strSQL = strSQL & ",'" & UserInfo.���� & "'" '  ����Ա����_In  In Varchar2,
    strSQL = strSQL & ",'" & UserInfo.��� & "'" '  ����Ա���_In  In Varchar2,
    strSQL = strSQL & ",'" & strTime & "'" '  ����ʱ��_In    In Varchar2,
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_nurse_order_id"))) '  ����ȼ�ҽ��id In Number, v_Jtmp := v_Jtmp || ',"nurse_order_id":1';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_undo_nurse")))   '  ���˱䶯_In    In Number,v_Jtmp := v_Jtmp || ',"undo_nurse":1';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_nurse_stop")))  '  ����ȼ�ͣ_In  In Number, v_Jtmp := v_Jtmp || ',"nurse_stop":1';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_no_print")))   '  ���δ�ӡ_In    In Number, v_Jtmp := v_Jtmp || ',"no_print":1';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_baby_num")))   '  Ӥ�����_In    In Number,v_Jtmp := v_Jtmp || ',"baby_num":' || Nvl(r_Advice.Ӥ��, 0);
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_pati_pageid")))  '  ��ҳid_In      In Number ',"pati_pageid":' || Nvl(r_Advice.��ҳid, 0);
    strSQL = strSQL & ",'" & GetColVal(colCISData, "_del_allergy") & "'"     'ɾ����_In       In Number := Null,',"del_allergy":' || Nvl(n_ɾ����, 0);
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_oper_after_msg"))) '������Ϣ_In     In Number := Null, ',"oper_after_msg":' || Nvl(n_������Ϣ, 0);
    strSQL = strSQL & ",'" & GetColVal(colCISData, "_area_msg") & "'"   '����ִ����Ϣ_In In Varchar2 := Null, --ҽ��idƴ�����ŷָ� ',"area_msg":"' || v_����ִ��ids || '"';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_pacs_msg")))    '�����Ϣ_In
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_lis_msg")))     '������Ϣ_In
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_blood_msg")))      '��Ѫ��Ϣ_In     In Number := Null, ',"blood_msg":' || Nvl(n_��Ѫ��Ϣ, 0);
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_oper_msg"))) '������Ϣ_In     In Number := Null,',"oper_msg":' || Nvl(n_������Ϣ, 0);
    strSQL = strSQL & ",'" & GetColVal(colCISData, "_rgst_no") & "'" '�Һŵ�_In       In Varchar2 := Null,',"rgst_no":"' || r_Advice.�Һŵ� || '"';
    strSQL = strSQL & "," & ZVal(Val(GetColVal(colCISData, "_send_no")))     '���ͺ�_In       In Number := Null,',"send_no":"' || v_���ͺ� || '"';
    strSQL = strSQL & ",'" & GetColVal(colCISData, "_bill_no") & "'"   'No_In           In Varchar2 := Null',"bill_no":"' || v_No || '"';
    strSQL = strSQL & ")"
    
    Getҽ������ִ�й���SQL = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function F_����ҽ����¼_����(ByVal Id_In As String, ByVal intҽ��վģ��� As Integer) As Boolean
'���ܣ����ﲡ��ҽ����������ҽ��ǰ��ؼ��
'��������ҽ��
'
'
'����:
'  Id_In         In ����ҽ����¼.Id%Type,
'  ����Ա���_In In ��Ա��.���%Type := Null,
'  ����Ա����_In In ��Ա��.����%Type := Null,
'  ����ҽ��id_In In ����ҽ����¼.Id%Type := Null,
'  ����ʱ��_In   In ����ҽ��״̬.����ʱ��%Type := Null
    Dim blnRis As Boolean, blnPacs As Boolean, blnDo As Boolean, col_fee_no_list As Collection
    Dim strCisOutPar As String, bln��ԤԼ��Ժҽ�� As Boolean, strJsonIn As String
    Dim col_order_info_list As Collection, lng�Һſ���ID As Long, colSQL As New Collection
    Dim blnһ����ҩ As Boolean, bln��ǩ�� As Boolean, strSQL As String
    Dim strҽ������ As String, strAdvice��Ѫ  As String, strRisIds  As String, strLISIDs  As String
    Dim strSign As String, str���ͺŴ� As String, strTime As String
    Dim str_after_order_ids As String '�����Ϻ���ҩ��ҽ����ids
    Dim str_auto_item_ids As String '�������Զ���ɵ�ҽ����Ŀids,��ʽ��ҽ��ID:��ϸĿID,ҽ��ID:��ϸĿID,ҽ��ID:��ϸĿID,,,,
    Dim lngҽ��ID As Long, str_fee_nos As String, lng_outpati_account As Long
    Dim strIDs As String, lng_bill_prop As Long, str_rcp_nos As String, str_stuff_nos As String, colPar As New Collection
    Dim lng����ID  As Long, str�Һŵ� As String, str_all_order_ids As String
    Dim blnMoved As Boolean, colPage As Collection
    Dim strSource As String, str��ִ�з���IDs As String
    Dim intRule  As Integer, str�������ϵ���ǩ�� As String
    Dim lng֤��ID As Long, strTimeStamp As String, strTimeStampCode As String, lng�Һ�ID As Long, lngǩ��id As Long
    
    On Error GoTo errH
    
    Call HaveRIS
    blnRis = (Not gobjRis Is Nothing) And gbln����Ӱ����ϢϵͳԤԼ
    If blnRis = False Then
        Call CreateObjectPacs
        blnPacs = (Not gobjPACS Is Nothing) And gbln����PACSϵͳԤԼ
    End If
    strCisOutPar = zlDatabase.CallProcedure("Zl_����ҽ����¼_����_Check", "ҽ������", Val(Id_In), Empty)
    If Not GetSvrOutInfo(strCisOutPar) Then
        Exit Function
    End If
    Set col_order_info_list = gobjService.GetJsonListValue("output.order_info_list")
    Set col_fee_no_list = gobjService.GetJsonListValue("output.fee_no_list")
    bln��ԤԼ��Ժҽ�� = 1 = Val(GetColVal(col_order_info_list(1), "_is_order_appin"))
    bln��ǩ�� = 1 = Val(GetColVal(col_order_info_list(1), "_is_sign"))
    blnһ����ҩ = 1 = Val(GetColVal(col_order_info_list(1), "_is_merge"))
    strҽ������ = GetColVal(col_order_info_list(1), "_advice_note") & ""
    lngҽ��ID = Val(GetColVal(col_order_info_list(1), "_main_order_id"))
    str_all_order_ids = GetColVal(col_order_info_list(1), "_all_order_ids") & ""
    strAdvice��Ѫ = GetColVal(col_order_info_list(1), "_blood_order_ids") & ""
    strRisIds = GetColVal(col_order_info_list(1), "_ris_order_ids") & ""
    strLISIDs = GetColVal(col_order_info_list(1), "_lis_order_ids") & ""
    lng����ID = Val(GetColVal(col_order_info_list(1), "_pati_id"))
    lng�Һ�ID = Val(GetColVal(col_order_info_list(1), "_rgst_id"))
    str�Һŵ� = GetColVal(col_order_info_list(1), "_rgst_no") & ""
    lng�Һſ���ID = Val(GetColVal(col_order_info_list(1), "_rgst_deptid"))
    str_after_order_ids = GetColVal(col_order_info_list(1), "_after_order_ids") & ""
    str_auto_item_ids = GetColVal(col_order_info_list(1), "_auto_item_ids") & ""
    str���ͺŴ� = GetColVal(col_order_info_list(1), "_send_nos") & ""
    Set colPage = gobjService.GetJsonListValue("output.page_list")
    lng_bill_prop = Val(GetColVal(col_fee_no_list(1), "_bill_prop"))
    str_rcp_nos = GetColVal(col_fee_no_list(1), "_rcp_nos") & ""
    str_stuff_nos = GetColVal(col_fee_no_list(1), "_stuff_nos") & ""
    str_fee_nos = GetColVal(col_fee_no_list(1), "_fee_nos") & ""
    lng_outpati_account = Val(GetColVal(col_fee_no_list(1), "_outpati_account"))
    
    '����ǩ��������ʾ
    If bln��ǩ�� Then
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ��������ϡ�", vbInformation, gstrSysName
            Else
                MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ���������ϡ�", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        If gobjESign.CertificateStoped(UserInfo.����) = False Then strSign = vbCrLf & vbCrLf & "��ʾ����ҽ���Ѿ�ǩ��������ʱ����Ҫ�ٴ�ǩ����"
    End If
    
    If blnһ����ҩ Then
        If MsgBox("����һ����ҩ��ҽ������һ�����ϣ�ȷʵҪ������" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If MsgBox("ȷʵҪ����ҽ��""" & strҽ������ & """��" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    '����ʱ���е���ǩ��
    If bln��ǩ�� Then
        If gobjESign.CertificateStoped(UserInfo.����) = False Then
            '��ȡǩ��ҽ��Դ��
            strIDs = lngҽ��ID
            intRule = ReadAdviceSignSource(4, lng����ID, str�Һŵ�, strIDs, 0, blnMoved, strSource)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "���ܶ�ȡ��Ҫ���ϵ���ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode, , lng����ID, 0, lng�Һ�ID)
            If strSign <> "" Then
                If strTimeStamp <> "" Then
                    strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    strTimeStamp = "NULL"
                End If
                lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
                str�������ϵ���ǩ�� = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strIDs & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
            Else
                MsgBox "����ҽ������ǩ��ʧ�ܡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '��������ǰ��ҽӿ�
    blnDo = Check�������ҽ������ǰ(lng����ID, lng�Һ�ID, lngҽ��ID)
    If Not blnDo Then Exit Function
    
    strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    blnDo = Check�ѷ�ҩƷ����(str_all_order_ids, str_auto_item_ids, str_after_order_ids, str���ͺŴ�, strTime, lng_bill_prop, str_rcp_nos, str_stuff_nos, str��ִ�з���IDs, colPar)
    If Not blnDo Then Exit Function
    
    If str_fee_nos <> "" Then
        strJsonIn = "{""input"":{""bill_prop"":" & lng_bill_prop & _
        ",""outpati_account"":" & lng_outpati_account & ",""order_ids"":""" & str_all_order_ids & """" & _
         ",""fee_nos"":""" & str_fee_nos & """,""exe_fee_ids"":""" & str��ִ�з���IDs & """,""after_order_ids"":""" & str_after_order_ids & """}}"
        'colPar����  "���쳣���˴���","���˴���","���쳣��������","��������","���˷���"
        If Not Check���ҽ���������Ϸ���(strJsonIn, strTime, colPar) Then
            Exit Function
        End If
    End If
    
    '������������� ���ۻ���ԤԼ��Ժҽ������ж�
    If Not colPage Is Nothing Then
        If colPage.Count > 0 Then
            blnDo = Get��������Ժҽ��(colPage, colPar)
            If Not blnDo Then Exit Function
        End If
    End If
    
    '������������̣�����ҽ����ǩ��
    blnDo = Getҽ������ִ�й���SQL(strCisOutPar, strTime, strSQL)
    If Not blnDo Then Exit Function
    colSQL.Add strSQL
    If str�������ϵ���ǩ�� <> "" Then
        colSQL.Add str�������ϵ���ǩ��
    End If
    
    blnDo = Exeҽ����������(colPar, colSQL)
    If Not blnDo Then Exit Function
    '��Ѫ LIS RIS  PACS ����ҽ����ش���
    Call Get����ҽ�����������ӿ�(intҽ��վģ���, lngҽ��ID, strAdvice��Ѫ, strRisIds, strLISIDs, blnRis, blnPacs)
    
    '�������Ϻ���ҽӿ�
    Call Check�������ҽ�����Ϻ�(lng����ID, lng�Һ�ID, lngҽ��ID)
    
    '����ԤԼ���ķ���
    If bln��ԤԼ��Ժҽ�� Then
        If SvrԤԼ��Ժ��������(lng�Һſ���ID) Then
            Call SvrԤԼ��Ժȡ������(lng�Һ�ID)
        End If
    End If
    
    F_����ҽ����¼_���� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function F_����ҽ����¼_����_Bat(ByVal Ids_In As String, ByVal intҽ��վģ��� As Integer) As Boolean
                           
'���ܣ�����ҽ���������ϲ���ҽ������ǰ��ؼ��
'��������ҽ��
'
'
'����:
'  Id_In         In ����ҽ����¼.Id%Type,
'  ����Ա���_In In ��Ա��.���%Type := Null,
'  ����Ա����_In In ��Ա��.����%Type := Null,
'  ����ҽ��id_In In ����ҽ����¼.Id%Type := Null,
'  ����ʱ��_In   In ����ҽ��״̬.����ʱ��%Type := Null
    Dim blnRis As Boolean, blnPacs As Boolean, blnDo As Boolean, col_fee_no_list As Collection
    Dim strCisOutPar As String, bln��ԤԼ��Ժҽ�� As Boolean, strJsonIn As String
    Dim col_order_info_list As Collection, lng�Һſ���ID As Long, colSQL As New Collection
    Dim bln��ǩ�� As Boolean, strSQL As String, str��ִ�з���IDs As String
    Dim strAdvice��Ѫ    As String, strRisIds  As String, strLISIDs  As String
    Dim strSign As String, str���ͺŴ� As String, strTime As String, str_all_order_ids As String
    Dim str_after_order_ids As String '�����Ϻ���ҩ��ҽ����ids
    Dim str_auto_item_ids As String '�������Զ���ɵ�ҽ����Ŀids,��ʽ��ҽ��ID:��ϸĿID,ҽ��ID:��ϸĿID,ҽ��ID:��ϸĿID,,,,
    Dim lngҽ��ID As Long, str_fee_nos As String, lng_outpati_account As Long
    Dim strIDs As String, lng_bill_prop As Long, str_rcp_nos As String, str_stuff_nos As String, colPar As New Collection
    Dim lng����ID  As Long, str�Һŵ� As String, colCisOutPar As New Collection
    Dim blnMoved As Boolean, colPage As Collection, n As Long
    Dim strSource As String, Id_In As String, varBatҽ��ID As Variant
    Dim intRule  As Integer, str�������ϵ���ǩ�� As String
    Dim lng֤��ID As Long, strTimeStamp As String, strTimeStampCode As String, lng�Һ�ID As Long, lngǩ��id As Long
    
    On Error GoTo errH
    varBatҽ��ID = Split(Ids_In, ",")
    
    Call HaveRIS
    blnRis = (Not gobjRis Is Nothing) And gbln����Ӱ����ϢϵͳԤԼ
    If blnRis = False Then
        Call CreateObjectPacs
        blnPacs = (Not gobjPACS Is Nothing) And gbln����PACSϵͳԤԼ
    End If
    
    For n = 0 To UBound(varBatҽ��ID)
        Id_In = varBatҽ��ID(n)
        
        strCisOutPar = zlDatabase.CallProcedure("Zl_����ҽ����¼_����_Check", "ҽ������", Val(Id_In), Empty)
        If Not GetSvrOutInfo(strCisOutPar) Then
            Exit Function
        End If
        
        colCisOutPar.Add strCisOutPar
        
        Set col_order_info_list = gobjService.GetJsonListValue("output.order_info_list")
        Set col_fee_no_list = gobjService.GetJsonListValue("output.fee_no_list")
        
        If Not bln��ԤԼ��Ժҽ�� Then
            bln��ԤԼ��Ժҽ�� = 1 = Val(GetColVal(col_order_info_list(1), "_is_order_appin"))
        End If
        
        If Not bln��ǩ�� Then
            bln��ǩ�� = 1 = Val(GetColVal(col_order_info_list(1), "_is_sign"))
        End If
        
        strAdvice��Ѫ = MergeStr(strAdvice��Ѫ, GetColVal(col_order_info_list(1), "_blood_order_ids") & "")
        strRisIds = MergeStr(strRisIds, GetColVal(col_order_info_list(1), "_ris_order_ids") & "")
        strLISIDs = MergeStr(strLISIDs, GetColVal(col_order_info_list(1), "_lis_order_ids") & "")
        str_all_order_ids = MergeStr(str_all_order_ids, GetColVal(col_order_info_list(1), "_all_order_ids") & "")
        
        If lng����ID = 0 Then
            'һ�η�����ͬ��Ϣ
            lng����ID = Val(GetColVal(col_order_info_list(1), "_pati_id"))
            lng�Һ�ID = Val(GetColVal(col_order_info_list(1), "_rgst_id"))
            str�Һŵ� = GetColVal(col_order_info_list(1), "_rgst_no") & ""
            lng�Һſ���ID = Val(GetColVal(col_order_info_list(1), "_rgst_deptid"))
            str���ͺŴ� = GetColVal(col_order_info_list(1), "_send_nos") & ""
            lng_outpati_account = Val(GetColVal(col_fee_no_list(1), "_outpati_account"))
            lng_bill_prop = Val(GetColVal(col_fee_no_list(1), "_bill_prop"))
        End If
        
        str_after_order_ids = MergeStr(str_after_order_ids, GetColVal(col_order_info_list(1), "_after_order_ids") & "")
        str_auto_item_ids = MergeStr(str_auto_item_ids, GetColVal(col_order_info_list(1), "_auto_item_ids") & "")
        
        '��Ժ���ʵ�ҽ��һ��������ԭ����ֻ����һ��
        If colPage Is Nothing Then
            Set colPage = gobjService.GetJsonListValue("output.page_list")
        End If
        
        str_rcp_nos = MergeStr(str_rcp_nos, GetColVal(col_fee_no_list(1), "_rcp_nos") & "")
        str_stuff_nos = MergeStr(str_stuff_nos, GetColVal(col_fee_no_list(1), "_stuff_nos") & "")
        str_fee_nos = MergeStr(str_fee_nos, GetColVal(col_fee_no_list(1), "_fee_nos") & "")
        
    Next
    
    '����ǩ��������ʾ
    If bln��ǩ�� Then
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ��������ϡ�", vbInformation, gstrSysName
            Else
                MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ���������ϡ�", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        If gobjESign.CertificateStoped(UserInfo.����) = False Then strSign = vbCrLf & vbCrLf & "��ʾ����ҽ���Ѿ�ǩ��������ʱ����Ҫ�ٴ�ǩ����"
    End If
    
    '����ʱ���е���ǩ��
    If bln��ǩ�� Then
        If gobjESign.CertificateStoped(UserInfo.����) = False Then
            '��ȡǩ��ҽ��Դ��
            strIDs = Ids_In
            intRule = ReadAdviceSignSource(4, lng����ID, str�Һŵ�, strIDs, 0, blnMoved, strSource)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "���ܶ�ȡ��Ҫ���ϵ���ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode, , lng����ID, 0, lng�Һ�ID)
            If strSign <> "" Then
                If strTimeStamp <> "" Then
                    strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    strTimeStamp = "NULL"
                End If
                lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
                str�������ϵ���ǩ�� = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strIDs & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
            Else
                MsgBox "����ҽ������ǩ��ʧ�ܡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    For n = 0 To UBound(varBatҽ��ID)
        lngҽ��ID = varBatҽ��ID(n)
        '��������ǰ��ҽӿ�
        blnDo = Check�������ҽ������ǰ(lng����ID, lng�Һ�ID, lngҽ��ID)
        If Not blnDo Then Exit Function
    Next
    
    strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    blnDo = Check�ѷ�ҩƷ����(str_all_order_ids, str_auto_item_ids, str_after_order_ids, str���ͺŴ�, strTime, lng_bill_prop, str_rcp_nos, str_stuff_nos, str��ִ�з���IDs, colPar)
    If Not blnDo Then Exit Function
    
    If str_fee_nos <> "" Then
        strJsonIn = "{""input"":{""bill_prop"":" & lng_bill_prop & _
        ",""outpati_account"":" & lng_outpati_account & ",""order_ids"":""" & str_all_order_ids & """" & _
         ",""fee_nos"":""" & str_fee_nos & """,""exe_fee_ids"":""" & str��ִ�з���IDs & """,""after_order_ids"":""" & str_after_order_ids & """}}"
        'colPar����  "���쳣���˴���","���˴���","���쳣��������","��������","���˷���"
        If Not Check���ҽ���������Ϸ���(strJsonIn, strTime, colPar) Then
            Exit Function
        End If
    End If
    
    '������������� ���ۻ���ԤԼ��Ժҽ������ж�
    If Not colPage Is Nothing Then
        If colPage.Count > 0 Then
            blnDo = Get��������Ժҽ��(colPage, colPar)
            If Not blnDo Then Exit Function
        End If
    End If
    
    
    For n = 1 To colCisOutPar.Count
        strCisOutPar = colCisOutPar(n)
        '������������̣�����ҽ����ǩ��
        blnDo = Getҽ������ִ�й���SQL(strCisOutPar, strTime, strSQL)
        If Not blnDo Then Exit Function
        colSQL.Add strSQL
    
    Next
    
    If str�������ϵ���ǩ�� <> "" Then
        colSQL.Add str�������ϵ���ǩ��
    End If
    
    blnDo = Exeҽ����������(colPar, colSQL)
    If Not blnDo Then Exit Function
    
    '��Ѫ LIS RIS  PACS ����ҽ����ش���
    Call Get����ҽ�����������ӿ�(intҽ��վģ���, lngҽ��ID, strAdvice��Ѫ, strRisIds, strLISIDs, blnRis, blnPacs)
    
    For n = 0 To UBound(varBatҽ��ID)
        lngҽ��ID = varBatҽ��ID(n)
        '�������Ϻ���ҽӿ�
        Call Check�������ҽ�����Ϻ�(lng����ID, lng�Һ�ID, lngҽ��ID)
    Next
    
    '����ԤԼ���ķ���
    If bln��ԤԼ��Ժҽ�� Then
        If SvrԤԼ��Ժ��������(lng�Һſ���ID) Then
            Call SvrԤԼ��Ժȡ������(lng�Һ�ID)
        End If
    End If
    
    F_����ҽ����¼_����_Bat = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Function Check�������ҽ������ǰ(ByVal lng����ID As Long, ByVal lng�Һ�ID As Long, ByVal lngҽ��ID As Long) As Boolean
    Dim strErr As String
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    Call CreatePlugInOK(p����ҽ���´�, 0)
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        strErr = ""
        blnDo = gobjPlugIn.AdviceRevokedBefore(glngSys, p����ҽ���´�, lng����ID, lng�Һ�ID, lngҽ��ID, 0, strErr)
        Call zlPlugInErrH(err, "AdviceRevokedBefore")
        If 0 = err.Number Then '�ӿ�û�г������������жϽӿڵķ���ֵ
            If Not blnDo Then
                MsgBox strErr, vbInformation, gstrSysName
                Exit Function
            End If
        End If
        blnDo = False
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
    End If
    Check�������ҽ������ǰ = True
    Exit Function
errH:
    Check�������ҽ������ǰ = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check�������ҽ�����Ϻ�(ByVal lng����ID As Long, ByVal lng�Һ�ID As Long, ByVal lngҽ��ID As Long) As Boolean
    
    On Error GoTo errH
    
    Call CreatePlugInOK(p����ҽ���´�, 0)
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.AdviceRevoked(glngSys, p����ҽ���´�, lng����ID, lng�Һ�ID, lngҽ��ID, 0)
        Call zlPlugInErrH(err, "AdviceRevoked")
    End If
    Check�������ҽ�����Ϻ� = True
    Exit Function
errH:
    Check�������ҽ�����Ϻ� = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SvrԤԼ��Ժȡ������(ByVal lng�Һ�ID As Long)
'���ܣ�����ԤԼ���ķ���ȡ��סԺ����(�����ǰ����û������,ԤԼ����Ҳ�Ҳ���������,Ҳ���ᱨ��)
    Dim strJsIn As String
    Dim strJsOut As String
    Dim strErr As String
    Dim blnRet As Boolean
    
    strJsIn = "{""input_in"":{""rgst_id"": """ & lng�Һ�ID & """}}"
    zlWriteLog "ԤԼ���ķ��������־", "mdlCISKernel", "סԺ����ȡ��", LOGLEVEL_Trace, "����:" & strJsIn
    blnRet = sys.NewSystemSvr("ԤԼ����", "סԺ����ȡ��", strJsIn, strJsOut, strErr)
    zlWriteLog "ԤԼ���ķ��������־", "mdlCISKernel", "סԺ����ȡ��", LOGLEVEL_Trace, "���:" & "strJsOut=" & strJsOut & ";����ֵ=" & blnRet & ";strErr=" & strErr
    If strErr <> "" Then
        MsgBox "ԤԼ��Ժȡ������:" & strErr, vbInformation, gstrSysName
    End If
End Sub

Private Function Get����ҽ�����������ӿ�(ByVal intҽ��վģ��� As Long, ByVal lngҽ��ID As Long, _
    ByVal strAdvice��Ѫ As String, ByVal strRisIds As String, ByVal strLISIDs As String, _
    ByVal blnRis As Boolean, ByVal blnPacs As Boolean) As Boolean
'���ܣ�����ҽ�����������ӿڵ���
    Dim strErr As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If strAdvice��Ѫ <> "" Then
        If gobjPublicBlood.AdviceOperation(intҽ��վģ���, lngҽ��ID, 4, False, strErr) = False Then
            MsgBox "Ѫ��ϵͳ�ӿڵ���ʧ�ܣ�" & strErr, vbInformation, gstrSysName
        End If
    End If
    
    If strLISIDs <> "" Then
        Call InitObjLis(intҽ��վģ���)
        '����LIS�������뵥
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(CStr(lngҽ��ID), strErr) = False Then
                MsgBox "ɾ����������ʧ�ܣ�" & strErr, vbInformation, gstrSysName
            End If
        End If
    End If
    
    'RIS���ˣ�����ʧ�����˳�
    If strRisIds <> "" Then '��顢����������
        If HaveRIS(True) Then
            On Error Resume Next
            If gobjRis.HISRollAdvice(lngҽ��ID) <> 1 Then
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISRollAdvice)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
            err.Clear: On Error GoTo 0
            On Error GoTo errH
        End If
        
        'ɾ��ԤԼ��Ϣ����
        If blnRis Then
            Set rsTmp = GetDataRISԤԼ(lngҽ��ID & "")
            If Not rsTmp.EOF Then
                On Error Resume Next
                If 0 = gobjRis.HISSchedulingEx(Val(rsTmp!ID & ""), Val(rsTmp!ԤԼid & "")) Then
                    MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ����β���ɾ�����޸����Ѿ�ԤԼҽ����������Ӱ����Ϣϵͳ�ӿ�(HISSchedulingEx)ȡ��ϢԤԼδ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                End If
                err.Clear: On Error GoTo errH
            End If
        ElseIf blnPacs Then
            Set rsTmp = GetDataPACSԤԼ(lngҽ��ID & "")
            If Not rsTmp.EOF Then
                On Error Resume Next
                If False = gobjPACS.CancelSchedule(Val(rsTmp!ID & "")) Then
                    MsgBox "��ǰ������PACS��Ϣϵͳ�ӿڣ����β���ɾ�����޸����Ѿ�ԤԼҽ����������PACS��Ϣϵͳ�ӿ�(CancelSchedule)ȡ��ϢԤԼδ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                End If
                err.Clear: On Error GoTo errH
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check�ѷ�ҩƷ����(ByVal str_all_order_ids As String, ByVal strAutoitems As String, ByVal strOrderIds�ȷ� As String, ByVal strSendNos As String, ByVal strTime As String, _
    ByVal lng_billtype As Long, ByVal str_rcp_nos As String, ByVal str_stuff_nos As String, ByRef str��ִ�з���IDs As String, colPar As Collection) As Boolean
'���ܣ�����Ѿ���ҩƷ���ģ����ﲡ������ҽ��
'������strAutoitems ����ɵ���Ŀ����Ҫ�Զ��ˣ���ʽ��ҽ��ID:�շ�ϸĿID,,,,
'      strOrderIds�ȷ� �����Ϻ���ҩ��ҽ����ids�����ŷָ�
    Dim i As Long
    Dim strJsonIn As String
    Dim colDrugExe As Collection
    Dim colStuffExe As Collection
    Dim lngҽ��ID As Long
    Dim lng�շ�ϸĿID As Long
    Dim strDelDrug As String
    Dim strDelStuff As String
    Dim strRetDrug As String
    Dim strRetStuff As String
    Dim strDelDrugJson As String
    Dim strDelStuffJson As String
    
    Dim str�쳣��� As String
    Dim strSQL As String
    Dim strMsg As String
    Dim blnDo�ȷ� As Boolean
    Dim str��ִ��ҩƷ������ϸIDs As String '���������Ϻ���ҩ��ʱ��ҩƷ������ִ���ⲿ�ݷ��ò�����Ҫ���ţ���ʽ������id����ƴ��
    Dim bln�����Ϻ���ҩģʽ As Boolean
    Dim str���շѷ���IDs As String
    
    On Error GoTo errH
     
    
    Check�ѷ�ҩƷ���� = True
    bln�����Ϻ���ҩģʽ = strOrderIds�ȷ� <> ""
    
    If str_rcp_nos <> "" Then
        strJsonIn = "{""input"":{""billtype"":" & lng_billtype & _
            ",""rcp_nos"":""" & str_rcp_nos & """,""order_ids"":""" & str_all_order_ids & """}}"
        Call CallService("zl_DrugSvr_GetExecutedNum", strJsonIn, , , , False, , , , True)
        Set colDrugExe = gobjService.GetJsonListValue("output.data")
    End If
    
    If str_stuff_nos <> "" Then
        strJsonIn = "{""input"":{""billtype"":" & lng_billtype & ",""stuff_nos"":""" & str_stuff_nos & """,""order_ids"":""" & str_all_order_ids & """}}"
        Call CallService("zl_StuffSvr_GetExecutedNum", strJsonIn, , , , False, , , , True)
        Set colStuffExe = gobjService.GetJsonListValue("output.item_list")
    End If
    
    If bln�����Ϻ���ҩģʽ Then
        If Not Get���շѷ�����ϸ(lng_billtype, colDrugExe, colStuffExe, str���շѷ���IDs) Then
            Check�ѷ�ҩƷ���� = False
            Exit Function
        End If
    End If
    
    If Not colDrugExe Is Nothing Then
        If colDrugExe.Count > 0 Then
            For i = 1 To colDrugExe.Count
                blnDo�ȷ� = False
                lngҽ��ID = Val(colDrugExe(i)("_order_id") & "")
                lng�շ�ϸĿID = colDrugExe(i)("_drug_id")
                
                If Val(colDrugExe(i)("_sended_num") & "") <> 0 Then
                    '�����Ϻ����жϴ���
                    blnDo�ȷ� = InStr("," & strOrderIds�ȷ� & ",", "," & lngҽ��ID & ",") > 0
                    If blnDo�ȷ� Then
                        '��������Ϻ���ҩ���Ѿ�ִ�д�ʱ�ⲿ�ݷ��ò��ܽ��д���
                        str��ִ��ҩƷ������ϸIDs = str��ִ��ҩƷ������ϸIDs & "," & colDrugExe(i)("_rcpdtl_id")
                    Else
                        If InStr("," & strAutoitems & ",", "," & lngҽ��ID & ":" & lng�շ�ϸĿID & ",") = 0 Then
    
                            strMsg = "ҽ�����͵ķ��õ���""" & colDrugExe(i)("_rcp_no") & """�е������Ѿ������ֻ���ȫִ�У��������ϡ�"
                            Screen.MousePointer = 0
                            MsgBox strMsg, vbInformation, gstrSysName
                            Screen.MousePointer = 11
                            Check�ѷ�ҩƷ���� = False
                            Exit Function
                        End If
    
                        '�Զ���ҩ+ɾ������
                        'Zl_Drugsvr_Delrecipebill
                        strRetDrug = strRetDrug & "," & colDrugExe(i)("_rcpdtl_id") & ":" & Val(colDrugExe(i)("_sended_num") & "")
                    End If
                End If
                
                If Not blnDo�ȷ� Then
                    If str���շѷ���IDs = "" Or InStr("," & str���շѷ���IDs & ",", "," & colDrugExe(i)("_rcpdtl_id") & ",") = 0 Then
                        If InStr("," & str�쳣��� & ",", "," & lngҽ��ID & ",") = 0 Then
                            str�쳣��� = str�쳣��� & "," & lngҽ��ID
                        End If
                        
                        strDelDrug = strDelDrug & ",{""rcpdtl_id"":" & colDrugExe(i)("_rcpdtl_id")
                        strDelDrug = strDelDrug & ",""chargeoffs_num"":" & Val(colDrugExe(i)("_sended_num") & "")
                        strDelDrug = strDelDrug & "}"
                    End If
                End If
            Next
        End If
    End If
    
    If str�쳣��� <> "" Then
        strSQL = "Zl_����ҽ������_����ͬ����־(3,'" & strSendNos & "','" & Mid(str�쳣���, 2) & "','" & UserInfo.���� & "','" & OS.ComputerName & "')"
        colPar.Add strSQL, "���쳣���˴���"
    End If
    str�쳣��� = ""
    If Not colStuffExe Is Nothing Then
        If colStuffExe.Count > 0 Then
            For i = 1 To colStuffExe.Count
                blnDo�ȷ� = False
                lngҽ��ID = Val(colStuffExe(i)("_order_id") & "")
                lng�շ�ϸĿID = Val(colStuffExe(i)("_stuff_id") & "")
                
                If Val(colStuffExe(i)("_sended_num") & "") <> 0 Then
                    
                    blnDo�ȷ� = InStr("," & strOrderIds�ȷ� & ",", "," & lngҽ��ID & ",") > 0
                    
                    If InStr("," & strAutoitems & ",", "," & lngҽ��ID & ":" & lng�շ�ϸĿID & ",") = 0 Then
                    
                        strMsg = "ҽ�����͵ķ��õ���""" & colStuffExe(i)("_stuff_no") & """�е������Ѿ������ֻ���ȫִ�У��������ϡ�"
                        Screen.MousePointer = 0
                        MsgBox strMsg, vbInformation, gstrSysName
                        Screen.MousePointer = 11
                        
                        Check�ѷ�ҩƷ���� = False
                        Exit Function
                    End If
                    '�Զ�����+ɾ������
                    'zl_stuffsvr_autoreturnstuff
                    'Zl_Stuffsvr_Delbill
                    strRetStuff = strRetStuff & "," & colStuffExe(i)("_stuffdtl_id") & ":" & Val(colStuffExe(i)("_sended_num") & "")
                    '����Զ��������ǾͿ���ɾ���ĵ��ݣ�����������˵�������Ϻ���ҩ��û���κ�Ӱ���
                    blnDo�ȷ� = False
                End If
                
                '˵��:����������Ϻ���ҩ����ģʽ,������Ѿ�ִ�з��ϵ�����Ӧ��ֻ�������ϲ�ɾ������
                If Not blnDo�ȷ� Then
                    If str���շѷ���IDs = "" Or InStr("," & str���շѷ���IDs & ",", "," & colStuffExe(i)("_stuffdtl_id") & ",") = 0 Then
                    
                        If InStr("," & str�쳣��� & ",", "," & lngҽ��ID & ",") = 0 Then
                            str�쳣��� = str�쳣��� & "," & lngҽ��ID
                        End If
                        
                        strDelStuff = strDelStuff & ",{""stuffdtl_id"":" & colStuffExe(i)("_stuffdtl_id")
                        strDelStuff = strDelStuff & ",""return_num"":" & Val(colStuffExe(i)("_sended_num") & "")
                        strDelStuff = strDelStuff & "}"
                    End If
                End If
                
            Next
        End If
    End If
    
    If str�쳣��� <> "" Then
        strSQL = "Zl_����ҽ������_����ͬ����־(4,'" & strSendNos & "','" & Mid(str�쳣���, 2) & "','" & UserInfo.���� & "','" & OS.ComputerName & "')"
        colPar.Add strSQL, "���쳣��������"
    End If
    
    '���˴���,Zl_����ҽ������_����ͬ����־
    If strDelDrug <> "" Then
        strDelDrugJson = """item_list"":[" & Mid(strDelDrug, 2) & "]"
        If strRetDrug <> "" Then
            strDelDrugJson = strDelDrugJson & ",""return_list"":[{""audit_operator"":""" & zlStr.ToJsonStr(UserInfo.����) & """" & _
                ",""operator_time"":""" & strTime & """" & _
                ",""rcpdtl_ids"":""" & Mid(strRetDrug, 2) & """}]"
        End If
        strDelDrugJson = "{""input"":{" & strDelDrugJson & "}}"
        colPar.Add strDelDrugJson, "���˴���"
    End If
    
    '��������,Zl_����ҽ������_����ͬ����־,�������ĵ�ʱ���п���������û��ɾԺ���ĵ���
    If strDelStuff <> "" Or strRetStuff <> "" Then
        strDelStuffJson = ""
        If strDelStuff <> "" Then strDelStuffJson = ",""item_list"":[" & Mid(strDelStuff, 2) & "]"
        
        If strRetStuff <> "" Then
            strDelStuffJson = strDelStuffJson & ",""return_list"":[{""audit_operator"":""" & zlStr.ToJsonStr(UserInfo.����) & """" & _
            ",""operator_time"":""" & strTime & """" & _
            ",""stuffdtl_ids"":""" & Mid(strRetStuff, 2) & """}]"
        End If
        strDelStuffJson = "{""input"":{" & Mid(strDelStuffJson, 2) & "}}"
        colPar.Add strDelStuffJson, "��������"
    End If
    str��ִ�з���IDs = Mid(str��ִ��ҩƷ������ϸIDs, 2)
    Exit Function
errH:
    Check�ѷ�ҩƷ���� = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check���ҽ���������Ϸ���(ByVal strParIn As String, ByVal strTime As String, colPar As Collection) As Boolean
'����:����ҽ�����Ϸ�������ؼ���ж�
    Dim strJsonIn As String
    Dim strPar As String
    Dim strExseOut As String
    Dim lng_exist_balance  As Long
    Dim lng_exist_verify  As Long
    Dim blnHaveMsg As Boolean
    
    On Error GoTo errH
    
 
    Call CallService("Zl_Exsesvr_CheckOrderRevoke", strParIn, strExseOut, , , False, , , , True)
    lng_exist_balance = Val(gobjService.GetJsonNodeValue("output.exist_balance") & "")
    lng_exist_verify = Val(gobjService.GetJsonNodeValue("output.exist_verify") & "")
 
    '�������ҽ����Ӧ�ķ��ý������
    If lng_exist_balance = 1 Then
        If gbytBillOpt = 1 Then
            If MsgBox("Ҫ����ҽ���Ķ�Ӧ�����д����ѽ��ʵķ��ã�ȷʵҪ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
            blnHaveMsg = True
        ElseIf gbytBillOpt = 2 Then
            MsgBox "Ҫ����ҽ���Ķ�Ӧ�����д����ѽ��ʵķ��ã��������ϡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '����˼��ʷ��ü��
    If InStr(GetInsidePrivs(p����ҽ���´�), "��������˼���ҽ��") = 0 Then
        If lng_exist_verify = 1 Then
            MsgBox "Ҫ����ҽ���Ķ�Ӧ���ʻ��۷����Ѿ���ˣ��������ϡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '��Ϊǰ������н�����ʾ����ֹ���������µ���һ��
    If blnHaveMsg Then
        Call CallService("Zl_Exsesvr_CheckOrderRevoke", strParIn, strExseOut, , , False, , , , True)
    End If
    
    strPar = GetListNodeTxt(strExseOut, "del_list")
    If strPar <> "" Then
        strJsonIn = "{""input"":{""operator_name"":""" & UserInfo.���� & """,""operator_code"":""" & UserInfo.��� & """,""operator_time"":""" & strTime & """" & strPar & "}}"
        colPar.Add strJsonIn, "���˷���"
    End If
    
    Check���ҽ���������Ϸ��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Function Exeҽ����������(colPar As Collection, colSQL As Collection) As Boolean
'����:����ҽ������ִ�й���
'����:colPar   "���쳣���˴���","���˴���","���쳣��������","��������","���˷���",colSQL"����ҽ��"
    Dim blnTran As Boolean
    Dim strSQL As String
    Dim strSvrName As String
    Dim strSvrPar As String
    Dim i As Long
    Dim lngTestMode As Long '���Դ������ڵ���,0-����,1-����ģʽ
    
    On Error GoTo errH
    
    lngTestMode = 1
'    If lngTestMode = 1 Then gcnOracle.BeginTrans: blnTran = True
    
    'ҩƷ����Ϊ�о��䣬���Բ��ñ���쳣����Ϊ�п���û���շ���¼�ˣ���Ҫ���һ������������
    strSQL = GetColVal(colPar, "���쳣���˴���")
    strSvrPar = GetColVal(colPar, "���˴���")
    strSvrName = "Zl_DrugSvr_DelRecipeBill"
    
    If strSvrPar <> "" Then
        gcnOracle.BeginTrans: blnTran = True
        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "����ҽ������")
        Call CallService(strSvrName, strSvrPar, , "����ҽ������", , False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    
    '���� �����Ϻ���ҩģʽ��  �������ĵ�ʱ���п���������,û��ɾ�����ĵ���
    strSQL = GetColVal(colPar, "���쳣��������")
    strSvrPar = GetColVal(colPar, "��������")
    strSvrName = "Zl_StuffSvr_DelBill"
    If strSvrPar <> "" Then
        gcnOracle.BeginTrans: blnTran = True
        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "����ҽ������")
        Call CallService(strSvrName, strSvrPar, , "����ҽ������", , False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    
    
    'ҽ��
    'strSQL = GetColVal(colSQL, "����ҽ��")
    strSvrPar = GetColVal(colPar, "���˷���")
    strSvrName = "Zl_ExseSvr_DelBill"
    If colSQL.Count > 0 Then
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To colSQL.Count
            strSQL = colSQL(i)
'            Debug.Print strSQL
            Call zlDatabase.ExecuteProcedure(strSQL, "����ҽ������")
        Next
        If strSvrPar <> "" Then
            Call CallService(strSvrName, strSvrPar, , "����ҽ������", , False, , , , True)
        End If
         gcnOracle.CommitTrans: blnTran = False
    End If
    
    
    strSvrPar = GetColVal(colPar, "��������Ϣ����")
    strSvrName = "Zl_Patisvr_UpdateInpatiState"
    If strSvrPar <> "" Then
        Call CallService(strSvrName, strSvrPar, , "����ҽ������", , True)
    End If
    
    'If lngTestMode = 1 Then gcnOracle.RollbackTrans: blnTran = False
    'gcnOracle.CommitTrans: blnTran = False
    Exeҽ���������� = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get��������Ժҽ��(colPage As Collection, colPar As Collection) As Boolean
'���ܣ�������������ҽ��
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim colPati As Collection
    Dim colPd As Collection
    Dim colP As Collection
    Dim lng����ID As Long, lngD��Ժ As Long, lngD��ҳ As Long, strUpPait As String, strTitle As String, lngģ�� As Long
    
    On Error GoTo errH
    
    If Not colPage Is Nothing Then
        If colPage.Count > 0 Then
            lng����ID = colPage(1)("_pati_id")
            If Not CallService("Zl_Patisvr_Lockcheck", "{""input"":{""pati_id"":" & lng����ID & "}}", , strTitle, lngģ��, True) Then
                Exit Function
            End If
            lngD��Ժ = colPage(1)("_del_in")
            lngD��ҳ = 1
            strJsonIn = "{""input"":{""pati_id"":" & lng����ID & ",""query_type"":3}}"
            Call CallService("Zl_Patisvr_Getpatiinfo", strJsonIn, strJsonOut, strTitle, lngģ��, True)
            Set colPati = gobjService.GetJsonListValue("output.pati_list")
            Set colP = colPati(1)
            Set colPd = colPage(1)
            strUpPait = strUpPait & "," & GetJsonStrNode("pati_id", "_pati_id", "N", colP, 1)
            strUpPait = strUpPait & "," & GetJsonStrNode("pati_pageid", "_pati_pageid", "N", colPd, 1)
            strUpPait = strUpPait & "," & GetJsonStrNode("inpatient_num", "_inpatient_num", "C", colPd)
            strUpPait = strUpPait & "," & GetJsonStrNode("in_time", "_in_time", "C", colPd)
            strUpPait = strUpPait & "," & GetJsonStrNode("adtd_time", "_adtd_time", "C", colPd)
            strUpPait = strUpPait & ",""pati_deptid"":null"
            strUpPait = strUpPait & ",""wardarea_id"":null"
            strUpPait = strUpPait & ",""pati_bed"":"""""
            strUpPait = strUpPait & ",""inp_status"":null"
            If Val("" & colPd("_pati_pageid")) = 0 Then
                strUpPait = strUpPait & ",""inp_times"":null"
            Else
                strUpPait = strUpPait & "," & GetJsonStrNode("inp_times", "_inp_times", "N", colP, 1)
            End If
            strUpPait = "{""input"":{""pati_list"":[{" & Mid(strUpPait, 2) & "}]}}"
            
        End If
    End If
    
    If strUpPait <> "" Then
        colPar.Add strUpPait, "��������Ϣ����"
    End If
    
    Get��������Ժҽ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Fҽ��ִ�����(colProSQL As Collection, ByVal strCisPar As String, ByVal strTime As String, ByVal lng����ID As Long, ByVal strTitle As String, ByVal lngģ�� As Long, ByRef colPar As Collection) As Boolean
'���ܣ�ҽ��ִ�������ؼ��
'������strCisPar ҽ��ִ�������ؼ�飬strTime ����ʱ�䣬lng����ID ִ�п���ID
'      colProSQL ��ִ�й���
    Dim strJsonIn As String, i As Long
    Dim strExseOutPar As String, lng���ͺ� As Long
    Dim col_fee_list As Collection, strJsonF As String, strSQL As String
    Dim lng_is_affirm As Long, lng_fee_origin As Long, lng_is_verify As Long
    Dim lng����ID As Long '�����жϱ����Զ����Ϸ�ҩ
    Dim str�Զ�����IDs As String
    Dim str�Զ���ҩIDs As String
    Dim str����ids As String
    Dim colҩƷ�շ�ȷ�� As New Collection
    Dim col�����շ�ȷ�� As New Collection
    Dim col�����շ�ȷ�� As New Collection
    Dim col_order_tag_list As Collection, colPati As Collection, cl��Ǽ� As New Collection, cl��Ǽ� As New Collection
    Dim strPatiInfo As String, blnDo As Boolean, lng����ID As Long, str��� As String
    Dim strItemҩƷ������� As String
    Dim str����idsҩƷ As String
    Dim str����ids���� As String
    
    On Error GoTo errH
    
    Call GetSvrOutInfo(strCisPar)
    Set col_order_tag_list = gobjService.GetJsonListValue("output.order_tag_list")
    lng���ͺ� = Val(gobjService.GetJsonNodeValue("output.send_no") & "")
    lng����ID = Val(gobjService.GetJsonNodeValue("output.wardarea_id") & "")
    lng����ID = Val(gobjService.GetJsonNodeValue("output.pati_id") & "")
    strJsonIn = "{""input"":{""is_finish"":1"
    strJsonIn = strJsonIn & ",""fee_nos"":""" & gobjService.GetJsonNodeValue("output.fee_nos") & """"
    strJsonIn = strJsonIn & ",""fee_order_ids"":""" & gobjService.GetJsonNodeValue("output.fee_order_ids") & """"
    strJsonIn = strJsonIn & ",""exe_deptid"":" & lng����ID
    strJsonIn = strJsonIn & "}}"
    
    If Not CallService("Zl_ExseSvr_GetOrderFeeExeInfo", strJsonIn, strExseOutPar, strTitle, lngģ��, True) Then
        Exit Function
    End If
    
    lng_is_affirm = Val(gobjService.GetJsonNodeValue("output.is_affirm") & "")
    
    '�����������һ����Ҫ��˷��ã�һ���ǲ���Ҫ��˷���
    Set col_fee_list = gobjService.GetJsonListValue("output.fee_list")
    If Not col_fee_list Is Nothing Then
        If col_fee_list.Count > 0 Then
            For i = 1 To col_fee_list.Count
                lng_is_verify = Val(col_fee_list(i)("_is_verify") & "")
                If InStr(",5,6,7,", "," & col_fee_list(i)("_fee_type") & ",") > 0 Then
                    blnDo = Fҽ��ִ���쳣�ж�(col_order_tag_list, col_fee_list(i))
                    If lng����ID <> 0 And Not blnDo Then
                        If lng����ID = Val(col_fee_list(i)("_exe_dept_id") & "") Then
                            If Val(col_fee_list(i)("_rec_state") & "") = 1 Or lng_is_verify = 1 Then
                                If Val(col_fee_list(i)("_fee_origin") & "") = 2 And Val(col_fee_list(i)("_bill_prop") & "") = 2 Then
                                    str�Զ���ҩIDs = str�Զ���ҩIDs & "," & col_fee_list(i)("_fee_id")
                                    str����idsҩƷ = str����idsҩƷ & "," & col_fee_list(i)("_fee_id")
                                End If
                            End If
                        End If
                        If lng_is_verify = 1 Then
                            colҩƷ�շ�ȷ��.Add col_fee_list(i)
                            col�����շ�ȷ��.Add col_fee_list(i)
                        End If
                    End If
                ElseIf Val(col_fee_list(i)("_stuff_used") & "") = 1 Then
                    blnDo = Fҽ��ִ���쳣�ж�(col_order_tag_list, col_fee_list(i))
                    If lng����ID <> 0 And Not blnDo Then
                        '�շѵ��ͼ��ʵ��ſ��Է���
                        If Val(col_fee_list(i)("_rec_state") & "") = 1 Or lng_is_verify = 1 Then
                            If Val(col_fee_list(i)("_fee_origin") & "") = 1 And (Val(col_fee_list(i)("_bill_prop") & "") = 1 Or Val(col_fee_list(i)("_bill_prop") & "") = 11) Or Val(col_fee_list(i)("_bill_prop") & "") = 2 Then
                                str�Զ�����IDs = str�Զ�����IDs & "," & col_fee_list(i)("_fee_id")
                                str����ids���� = str����ids���� & "," & col_fee_list(i)("_fee_id")
                            End If
                        End If
                        
                        If lng_is_verify = 1 Then
                            col�����շ�ȷ��.Add col_fee_list(i)
                            col�����շ�ȷ��.Add col_fee_list(i)
                        End If
                    End If
                Else
                    str����ids = str����ids & "," & col_fee_list(i)("_fee_id")
                    If lng_is_verify = 1 Then
                        col�����շ�ȷ��.Add col_fee_list(i)
                    End If
                End If
                
                lng_fee_origin = Val(col_fee_list(i)("_fee_origin") & "")
            Next
            
            
            
            If str����idsҩƷ <> "" Then
                strJsonIn = "{""input"":{""fee_ids"":""" & Mid(str����idsҩƷ, 2) & """,""oper_type"":1,""exe_people"":""" & UserInfo.���� & """,""exe_time"":""" & strTime & """,""fee_origin"":" & lng_fee_origin & ",""exe_status"":1}}"
                colPar.Add strJsonIn, "����ִ��ҩƷ"
            End If
            
            If str����ids���� <> "" Then
                strJsonIn = "{""input"":{""fee_ids"":""" & Mid(str����ids����, 2) & """,""oper_type"":1,""exe_people"":""" & UserInfo.���� & """,""exe_time"":""" & strTime & """,""fee_origin"":" & lng_fee_origin & ",""exe_status"":1}}"
                colPar.Add strJsonIn, "����ִ������"
            End If
            
            
            If str����ids <> "" Then
                strJsonIn = "{""input"":{""fee_ids"":""" & Mid(str����ids, 2) & """,""oper_type"":1,""exe_people"":""" & UserInfo.���� & """,""exe_time"":""" & strTime & """,""fee_origin"":" & lng_fee_origin & ",""exe_status"":1}}"
                colPar.Add strJsonIn, "����ִ��"
            End If
           
        End If
    End If
     
    '������շ�ȷ����ִ���շ�ȷ�ϣ���Ҫ��ָ���ʻ��۵�ִ��������
    If colҩƷ�շ�ȷ��.Count > 0 Or col�����շ�ȷ��.Count > 0 Then
        '--��ȡ���˻�����Ϣ
        strJsonIn = "{""input"":{""pati_id"":" & lng����ID & ",""query_type"":1}}"
        Call CallService("Zl_Patisvr_Getpatiinfo", strJsonIn, , strTitle, lngģ��, False, , , , True)
        Set colPati = gobjService.GetJsonListValue("output.pati_list")
        Set colPati = colPati(1)
        
        strPatiInfo = """pati_id"":" & lng����ID & _
            "," & GetJsonStrNode("pati_name", "_pati_name", "C", colPati) & _
            "," & GetJsonStrNode("pati_sex", "_pati_sex", "C", colPati) & _
            "," & GetJsonStrNode("pati_age", "_pati_age", "C", colPati) & _
            "," & GetJsonStrNode("pati_outpno", "_outpatient_num", "C", colPati) & _
            ",""auditor"":""" & UserInfo.���� & """,""auditor_code"":""" & UserInfo.��� & """,""audit_time"":""" & strTime & """"
        
        If colҩƷ�շ�ȷ��.Count > 0 Then
            strItemҩƷ������� = ""
            
            For i = 1 To colҩƷ�շ�ȷ��.Count
                str��� = "Zl_����ҽ������_����ͬ����־(5,'" & lng���ͺ� & "','" & colҩƷ�շ�ȷ��(i)("_order_id") & "','" & UserInfo.���� & "','" & OS.ComputerName & "')"
                cl��Ǽ�.Add str���
                
                str��� = "Zl_����ҽ������_����ͬ����־(7,'" & lng���ͺ� & "','" & colҩƷ�շ�ȷ��(i)("_order_id") & "')"
                cl��Ǽ�.Add str���
                
                strItemҩƷ������� = strItemҩƷ������� & ",{" & GetJsonStrNode("billtype", "_bill_prop", "N", colҩƷ�շ�ȷ��(i))
                strItemҩƷ������� = strItemҩƷ������� & "," & GetJsonStrNode("rcp_no", "_fee_no", "C", colҩƷ�շ�ȷ��(i))
                strItemҩƷ������� = strItemҩƷ������� & "," & GetJsonStrNode("rcpdtl_ids", "_fee_id", "C", colҩƷ�շ�ȷ��(i))
                
                If i = colҩƷ�շ�ȷ��.Count Then
                    If str�Զ���ҩIDs <> "" Then
                        strItemҩƷ������� = strItemҩƷ������� & ",""drug_auto_send"":1"
                        strItemҩƷ������� = strItemҩƷ������� & ",""auto_send_ids"":""" & Mid(str�Զ���ҩIDs, 2) & """"
                        str�Զ���ҩIDs = ""
                    End If
                End If
                
                strItemҩƷ������� = strItemҩƷ������� & "}"
                
            Next
            
            strJsonIn = "{""input"":{" & strPatiInfo & ",""item_list"":[" & Mid(strItemҩƷ�������, 2) & "]}}"
              
            colPar.Add cl��Ǽ�, "ҩƷ���쳣"
            colPar.Add strJsonIn, "ҩƷ�շ�"
            Set cl��Ǽ� = New Collection
            
        End If
        
        If col�����շ�ȷ��.Count > 0 Then
            strItemҩƷ������� = ""
            For i = 1 To col�����շ�ȷ��.Count
                str��� = "Zl_����ҽ������_����ͬ����־(6,'" & lng���ͺ� & "','" & col�����շ�ȷ��(i)("_order_id") & "','" & UserInfo.���� & "','" & OS.ComputerName & "')"
                cl��Ǽ�.Add str���
                
                str��� = "Zl_����ҽ������_����ͬ����־(8,'" & lng���ͺ� & "','" & col�����շ�ȷ��(i)("_order_id") & "')"
                cl��Ǽ�.Add str���
                
                strItemҩƷ������� = strItemҩƷ������� & ",{" & GetJsonStrNode("billtype", "_bill_prop", "N", col�����շ�ȷ��(i))
                strItemҩƷ������� = strItemҩƷ������� & "," & GetJsonStrNode("stuff_no", "_fee_no", "C", col�����շ�ȷ��(i))
                strItemҩƷ������� = strItemҩƷ������� & "," & GetJsonStrNode("stuffdtl_ids", "_fee_id", "C", col�����շ�ȷ��(i))
                
                If i = col�����շ�ȷ��.Count Then
                    If str�Զ�����IDs <> "" Then
                        strItemҩƷ������� = strItemҩƷ������� & ",""stuff_auto_send"":1"
                        strItemҩƷ������� = strItemҩƷ������� & ",""auto_send_ids"":""" & Mid(str�Զ�����IDs, 2) & """"
                        str�Զ�����IDs = ""
                    End If
                End If
                
                strItemҩƷ������� = strItemҩƷ������� & "}"
                
            Next
            
            strJsonIn = "{""input"":{" & strPatiInfo & ",""item_list"":[" & Mid(strItemҩƷ�������, 2) & "]}}"
            colPar.Add cl��Ǽ�, "�������쳣"
            colPar.Add strJsonIn, "�����շ�"
        End If
    End If
    
    '�������
    If col�����շ�ȷ��.Count > 0 Then
        strJsonF = ""
        For i = 1 To col�����շ�ȷ��.Count
            strJsonF = strJsonF & ",{ " & GetJsonStrNode("fee_source", "_fee_origin", "N", col�����շ�ȷ��(i))
            strJsonF = strJsonF & "," & GetJsonStrNode("fee_no", "_fee_no", "C", col�����շ�ȷ��(i))
            strJsonF = strJsonF & "," & GetJsonStrNode("serial_nums", "_serial_num", "C", col�����շ�ȷ��(i))
            strJsonF = strJsonF & ",""pati_id"":" & lng����ID
            strJsonF = strJsonF & "}"
        Next
        strJsonIn = "{""input"":{""operator_name"":""" & UserInfo.���� & """,""operator_code"":""" & UserInfo.��� & """,""operator_time"":""" & strTime & """,""item_list"":[" & Mid(strJsonF, 2) & "]}}"
        colPar.Add cl��Ǽ�, "ҽ�����쳣"
        colPar.Add strJsonIn, "�����շ�"
    End If
    
    If str�Զ���ҩIDs <> "" Then
        strJsonIn = "{""input"":{""rcpdtl_ids"":""" & Mid(str�Զ���ҩIDs, 2) & """,""send_type"":1,""operator_name"":""" & UserInfo.���� & """,""operator_code"":""" & UserInfo.��� & """}}"
        colPar.Add strJsonIn, "ҩƷ��ҩ"
    End If
    
    If str�Զ�����IDs <> "" Then
        strJsonIn = "{""input"":{""stuffdtl_ids"":""" & Mid(str�Զ�����IDs, 2) & """,""send_type"":1,""operator_name"":""" & UserInfo.���� & """,""operator_code"":""" & UserInfo.��� & """}}"
        colPar.Add strJsonIn, "���ķ���"
    End If
    
    '���� ҽ��ִ����ɹ���
    Set colPati = New Collection
    For i = 1 To colProSQL.Count
        If InStr("ZLHIS_" & UCase(colProSQL(i) & ""), UCase("ZLHIS_Zl_����ҽ��ִ��_Finish_S(")) > 0 Then
            strSQL = ""
            If Not Getҽ��ִ�����SQL(strCisPar, strTime, strSQL) Then
                Exit Function
            End If
'            colProSQL(i) = strSQL
        Else
            strSQL = colProSQL(i)
        End If
        colPati.Add strSQL
    Next
    
    colPar.Add colPati, "ҽ��ִ�����"
    
    
    Fҽ��ִ����� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Fҽ��ִ���쳣�ж�(col_order_tag As Collection, col_item As Collection) As Boolean
'����:ҽ��ִ��ʱ�ж��Ƿ����쳣��ҩƷ����
'���أ�true �Ƿ��쳣��false ����
    Dim i As Long
    
    On Error GoTo errH
     
    If Not col_order_tag Is Nothing Then
        For i = 1 To col_order_tag.Count
            If col_order_tag(i)("_drug_tag") <> 0 And col_order_tag(i)("_order_id") = col_item("_order_id") And col_item("_stuff_used") = 0 Or _
                col_order_tag(i)("_stuff_tag") <> 0 And col_order_tag(i)("_order_id") = col_item("_order_id") And col_item("_stuff_used") = 1 Then
                  
                Fҽ��ִ���쳣�ж� = True
                Exit Function
            End If
        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Getҽ��ִ�����SQL(ByVal strCisPar As String, ByVal strTime As String, ByRef strSQL As String) As Boolean
    On Error GoTo errH
    Call GetSvrOutInfo(strCisPar)
'    Set col_order_tag_list = gobjService.GetJsonListValue("output.order_tag_list")
'    lng���ͺ� = Val(gobjService.GetJsonNodeValue("output.send_no") & "")
'    lng����ID = Val(gobjService.GetJsonNodeValue("output.wardarea_id") & "")
    
'Create Or Replace Procedure Zl_����ҽ��ִ��_Finish_s
'(
'���ҽ��ids_In     Varchar2,
'�ɼ����ҽ��ids_In Varchar2,
'ҽ��id_In          Number,
'���ͺ�_In          Number,
'���ʱ��_In        Date,
'����Ա����_In      Varchar2,
'����id_In          Number := Null,
'��ҳid_In          Number := Null,
'�Һŵ�_In          Varchar2 := Null,
'������ʼʱ��_In    Varchar2 := Null,
'����ִ�п���id_In Number:=Null

    strSQL = "Zl_����ҽ��ִ��_Finish_S("
    strSQL = strSQL & "'" & gobjService.GetJsonNodeValue("output.finish_order_ids") & "'" '���ҽ��ids_In     Varchar2,
    strSQL = strSQL & ",'" & gobjService.GetJsonNodeValue("output.lis_order_ids") & "'" '�ɼ����ҽ��ids_In Varchar2,
    strSQL = strSQL & "," & gobjService.GetJsonNodeValue("output.order_id") 'ҽ��id_In          Number,
    strSQL = strSQL & "," & gobjService.GetJsonNodeValue("output.send_no") '���ͺ�_In          Number,
    strSQL = strSQL & ",to_date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')" '���ʱ��_In        Date,
    strSQL = strSQL & ",'" & UserInfo.���� & "'"  '����Ա����_In      Varchar2,
    strSQL = strSQL & "," & gobjService.GetJsonNodeValue("output.pati_id") '����id_In          Number := Null,
    strSQL = strSQL & "," & ZVal(Val("" & gobjService.GetJsonNodeValue("output.pati_pageid"))) '��ҳid_In          Number := Null,
    strSQL = strSQL & ",'" & gobjService.GetJsonNodeValue("output.rgst_no") & "'" '�Һŵ�_In          Varchar2 := Null,
    strSQL = strSQL & ",'" & gobjService.GetJsonNodeValue("output.operate_time") & "'" '������ʼʱ��_In    Varchar2 := Null,
    strSQL = strSQL & "," & ZVal(Val("" & gobjService.GetJsonNodeValue("output.operate_deptid"))) '����ִ�п���id_In Number:=Null
    strSQL = strSQL & ")"
    
    Getҽ��ִ�����SQL = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckסԺҽ������(ByVal strPars As String, ByRef colOut As Collection, ByVal strTitle As String, ByVal lngģ�� As Long) As Boolean
'���ܣ�סԺҽ�����ϼ��,ͬʱ����ҽ�����ϵ����
'��Σ�strPars Ҫ���ϵ�ҽ����Ϣ��ҽ��ID:����ȼ�ҽ��ID,ҽ��ID:����ȼ�ҽ��ID,....
'���Σ�colOut��Ӧ����μ���
    Dim i As Long
    Dim varTmp As Variant
    Dim strJsonIn As String
    Dim strTime As String
    Dim strCHK���� As String
    
    On Error GoTo errH
    Set colOut = New Collection
    varTmp = Split(strPars, ",")
    strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    For i = 0 To UBound(varTmp)
        strCHK���� = zlDatabase.CallProcedure("Zl_����ҽ����¼_����_Check", strTitle & "_" & lngģ��, Val("" & varTmp(i)), Empty)
        If Not GetSvrOutInfo(strCHK����) Then
            Exit Function
        End If
        If Not Getҽ������ִ�й���SQL(strCHK����, strTime, strJsonIn) Then
            Exit Function
        End If
        colOut.Add strJsonIn
        strJsonIn = ""
    Next
    
    CheckסԺҽ������ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get���շѷ�����ϸ(ByVal lng_billtype As Long, colDrugExe As Collection, colStuffExe As Collection, ByRef str���շѷ���IDs As String) As Boolean
'���ܣ������Ϻ���ҩģʽʱ�ж��Ƿ����շ�
    Dim i As Long
    Dim strJsonIn As String
    Dim strFeeids As String
    Dim strFid As String
    On Error GoTo errH
    
    If Not colDrugExe Is Nothing Then
        If colDrugExe.Count > 0 Then
            For i = 1 To colDrugExe.Count
                strFid = colDrugExe(i)("_rcpdtl_id") & ""
                If InStr("," & strFeeids & ",", "," & strFid & ",") = 0 Then
                    strFeeids = strFeeids & "," & strFid
                End If
            Next
        End If
    End If
    
    If Not colStuffExe Is Nothing Then
        If colStuffExe.Count > 0 Then
            For i = 1 To colStuffExe.Count
                strFid = colStuffExe(i)("_stuffdtl_id") & ""
                If InStr("," & strFeeids & ",", "," & strFid & ",") = 0 Then
                    strFeeids = strFeeids & "," & strFid
                End If
            Next
        End If
    End If
    
    If strFeeids <> "" Then
        strJsonIn = "{""input"":{""fee_origin"":" & lng_billtype & ",""fee_ids"":""" & Mid(strFeeids, 2) & """,""query_type"":7}}"
        Call CallService("Zl_Exsesvr_Getorderfeeinfo", strJsonIn, , , , False, , , , True)
        str���շѷ���IDs = gobjService.GetJsonNodeValue("output.fee_ids") & ""
    End If
    
    Get���շѷ�����ϸ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ExeҩƷ�շ�ȷ��(colPar As Collection, ByVal strTitle As String, ByVal lngģ�� As Long) As Boolean
'���ܣ�ҽ��ִ����ɣ����ʻ��۵��Զ���˷���ʱ����ҩƷ���ý����շ�ȷ��
'˵����
    Dim strJsonIn As String, clTag As Collection, i As Long, strSvrName As String, blnTran As Boolean
    On Error GoTo errH
    strJsonIn = GetColVal(colPar, "ҩƷ�շ�")
    Set clTag = GetColObj(colPar, "ҩƷ���쳣")
    strSvrName = "Zl_DrugSvr_RecipeAffirm"
    If strJsonIn <> "" Then
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To clTag.Count
            Call zlDatabase.ExecuteProcedure(clTag(i) & "", strTitle)
        Next
        Call CallService(strSvrName, strJsonIn, , strTitle, lngģ��, False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    strJsonIn = GetColVal(colPar, "����ִ��ҩƷ")
    If strJsonIn <> "" Then
'        Call CallService("Zl_ExseSvr_BillInforUpdate", strJsonIn, , strTitle, lngģ��, False, , , , True)
    End If
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Exe�����շ�ȷ��(colPar As Collection, ByVal strTitle As String, ByVal lngģ�� As Long) As Boolean
'���ܣ�ҽ��ִ����ɣ����ʻ��۵��Զ���˷���ʱ�������ķ��ý����շ�ȷ��
'˵����
    Dim strJsonIn As String, clTag As Collection, i As Long, strSvrName As String, blnTran As Boolean
    On Error GoTo errH
    strJsonIn = GetColVal(colPar, "�����շ�")
    Set clTag = GetColObj(colPar, "�������쳣")
    strSvrName = "Zl_StuffSvr_BillAffirm"
    If strJsonIn <> "" Then
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To clTag.Count
            Call zlDatabase.ExecuteProcedure(clTag(i) & "", strTitle)
        Next
        Call CallService(strSvrName, strJsonIn, , strTitle, lngģ��, False, , , , True)
        gcnOracle.CommitTrans: blnTran = False
    End If
    strJsonIn = GetColVal(colPar, "����ִ������")
    If strJsonIn <> "" Then
'        Call CallService("Zl_ExseSvr_BillInforUpdate", strJsonIn, , strTitle, lngģ��, False, , , , True)
    End If
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Upd����ִ��״̬(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long) As Boolean
'���ܣ�����ִ��״̬ͳһ����,��������ķ�ʽ
    Dim rsExc As ADODB.Recordset 'ҽ��ִ�мƼ�
    Dim rsSend As ADODB.Recordset
    Dim lng������� As Long, lng��¼���� As Long, lng��ҳID As Long
    Dim colStu As Collection, colRcp As Collection, strJsonIn As String
    Dim strHead As String, strTime As String, strItem As String, strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select a.ҽ��id,a.���id,a.�������,a.NO,a.�շ�ϸĿid,a.�������,a.��¼����,Sum(a.��ִ����) As ��ִ���� from (" & vbNewLine & _
        "Select b.ҽ��id,c.���id,c.�������,b.No,a.�շ�ϸĿid,Decode(a.ִ��״̬,1,1,Decode(a.ִ��״̬,1,1,0))*a.���� As ��ִ����,b.�������,b.��¼����" & vbNewLine & _
        "From ҽ��ִ�мƼ� A,����ҽ������ B,����ҽ����¼ C" & vbNewLine & _
        "Where b.ҽ��id=c.id and a.����<>0 and a.ҽ��id =[1] And a.���ͺ� =[2] And a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ�) a" & vbNewLine & _
        "Group By a.ҽ��id, a.NO, a.�շ�ϸĿid,a.�������,a.��¼����,a.���id,a.�������"

    Set rsExc = zlDatabase.OpenSQLRecord(strSQL, "Upd����ִ��״̬", lngҽ��ID, lng���ͺ�)
    If Not rsExc.EOF Then
        lng������� = Val(rsExc!������� & "")
        lng��¼���� = Val(rsExc!��¼���� & "")
'        lng��ҳID = Val(rsExc!��ҳID & "")
'        If Val(rsExc!���ID & "") <> 0 And InStr(",5,6,7,", "," & rsExc!������� & ",") > 0 Then
            strSQL = "select a.ҽ��id,a.NO,a.�������,a.��¼���� from ����ҽ������ a, ����ҽ����¼ b where a.���ͺ�+0=[2] and a.ҽ��id=b.id and b.���id=[1]"
            Set rsSend = zlDatabase.OpenSQLRecord(strSQL, "Upd����ִ��״̬", lngҽ��ID, lng���ͺ�)
'        End If
    End If
    
    Call Getִ��״̬������(rsExc, rsSend, colRcp, colStu)
    
    'סԺû���շ��ʵ���һ˵
    If lng������� = 1 Then
        strHead = strHead & ",""fee_origin"":1"
    ElseIf lng��¼���� = 2 Then
        strHead = strHead & ",""fee_origin"":2"
    Else
        strHead = strHead & ",""fee_origin"":1"
    End If
    strTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    strHead = strHead & ",""operator_name"":""" & UserInfo.���� & """"
    strHead = strHead & ",""operator_code"":""" & UserInfo.��� & """"
    strHead = strHead & ",""operator_time"":""" & strTime & """"
    
    strItem = Get״̬��ϸ(rsExc, colRcp, colStu, strTime)
    
    If strItem <> "" Then
        strJsonIn = "{""input"":{" & Mid(strHead, 2) & ",""item_list"":" & strItem & "}}"
        Call In���·���ִ��״̬(strJsonIn)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function In���·���ִ��״̬(ByVal strJsonIn As String)
    On Error GoTo errH
    Call CallService("Zl_Exsesvr_UpdateExeInfo", strJsonIn, , "In���·���ִ��״̬", , False, , , , True)
    Exit Function
errH:
    err.Clear
End Function

Private Function Get״̬��ϸ(rsA As ADODB.Recordset, colStu As Collection, colRcp As Collection, ByVal strTime As String) As String
'���ܣ���ȡ״̬��ϸ���
    Dim i As Long
    Dim strAll As String
    Dim strItem As String, strList As String
    
    On Error GoTo errH
    
    If Not colStu Is Nothing Then
        If colStu.Count > 0 Then
        For i = 1 To colStu.Count
            strItem = "{""fee_no"":""" & colStu(i)("_stuff_no") & """"
            strItem = strItem & ",""advice_id"":" & colStu(i)("_order_id")
            strItem = strItem & ",""fee_item_id"":" & colStu(i)("_stuff_id"): strAll = strAll & "," & strItem '����һ�Σ����ú���ȥ��
            strItem = strItem & ",""exe_nums"":" & GetJsonNum("" & colStu(i)("_sended_num"))
            strItem = strItem & ",""exe_people"":""" & IIF(Val("" & colStu(i)("_sended_num")) = 0, "", UserInfo.����) & """"
            strItem = strItem & ",""exe_time"":""" & IIF(Val("" & colStu(i)("_sended_num")) = 0, "", strTime) & """"
            strItem = strItem & "}"
            strList = strList & "," & strItem
        Next
        End If
    End If
    
    
    If Not colRcp Is Nothing Then
        If colRcp.Count > 0 Then
        For i = 1 To colRcp.Count
            strItem = "{""fee_no"":""" & colRcp(i)("_rcp_no") & """"
            strItem = strItem & ",""advice_id"":" & colRcp(i)("_order_id")
            strItem = strItem & ",""fee_item_id"":" & colRcp(i)("_drug_id"): strAll = strAll & "," & strItem '����һ�Σ����ú���ȥ��
            strItem = strItem & ",""exe_nums"":" & GetJsonNum("" & colRcp(i)("_sended_num"))
            strItem = strItem & ",""exe_people"":""" & IIF(Val("" & colRcp(i)("_sended_num")) = 0, "", UserInfo.����) & """"
            strItem = strItem & ",""exe_time"":""" & IIF(Val("" & colRcp(i)("_sended_num")) = 0, "", strTime) & """"
            strItem = strItem & "}"
            strList = strList & "," & strItem
        Next
        End If
    End If
    
    If Not rsA.EOF Then
        For i = 1 To rsA.RecordCount
        
            strItem = "{""fee_no"":""" & rsA!NO & """"
            strItem = strItem & ",""advice_id"":" & rsA!ҽ��ID
            strItem = strItem & ",""fee_item_id"":" & rsA!�շ�ϸĿid
            
            If InStr("," & strAll & ",", "," & strItem & ",") = 0 Then
                strItem = strItem & ",""exe_nums"":" & GetJsonNum("" & rsA!��ִ����)
                
                strItem = strItem & ",""exe_people"":""" & IIF(Val("" & rsA!��ִ����) = 0, "", UserInfo.����) & """"
                strItem = strItem & ",""exe_time"":""" & IIF(Val("" & rsA!��ִ����) = 0, "", strTime) & """"
                
                strItem = strItem & "}"
                strList = strList & "," & strItem
            End If
            
            rsA.MoveNext
        Next
    End If
    If strList <> "" Then
    Get״̬��ϸ = "[" & Mid(strList, 2) & "]"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Getִ��״̬������(rsA As ADODB.Recordset, rsB As ADODB.Recordset, ByRef colStu As Collection, ByRef colRcp As Collection) As Boolean
'���ܣ���ȡ����ִ��״̬ʱ��Ӧ��ҩƷ���ĵ����ж�Ӧ��ִ����
    Dim i As Long
    Dim strNos As String, strOrderIds As String
    
    
    On Error GoTo errH
    
    If rsA.EOF Then Exit Function
    
    For i = 1 To rsA.RecordCount
        
        If InStr("," & strNos & ",", "," & rsA!NO & ",") = 0 Then strNos = strNos & "," & rsA!NO
        
        If InStr("," & strOrderIds & ",", "," & rsA!ҽ��ID & ",") = 0 Then strOrderIds = strOrderIds & "," & rsA!ҽ��ID
    
        
        rsA.MoveNext
    Next
    rsA.MoveFirst
    
    If Not rsB Is Nothing Then
        If Not rsB.EOF Then
            For i = 1 To rsB.RecordCount
                
                If InStr("," & strNos & ",", "," & rsB!NO & ",") = 0 Then strNos = strNos & "," & rsB!NO
        
                If InStr("," & strOrderIds & ",", "," & rsB!ҽ��ID & ",") = 0 Then strOrderIds = strNos & "," & rsB!ҽ��ID
                
                rsB.MoveNext
            Next
            rsB.MoveFirst
        End If
    End If
    
    
    Call Get����ִ����(Mid(strNos, 2), Mid(strOrderIds, 2), colStu)
    
    Call GetҩƷִ����(Mid(strNos, 2), Mid(strOrderIds, 2), colRcp)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetҩƷִ����(ByVal strN As String, ByVal strOd As String, ByRef colOut As Collection) As Boolean
'���ܣ����ݵ��ݺ�+ҽ��ID����ȡҩƷ��ִ������
'������strN ���ݺŶ���ƴ����strOd ҽ��ID����ƴ��
    Dim strJson As String
    
    On Error GoTo errH
    
    If strN = "" Then Exit Function
    
    strJson = "{""input"":{""billtype"":3,""rcp_nos"":""" & strN & """,""order_ids"":""" & strOd & """}}"
    Call CallService("Zl_Drugsvr_Getexecutednum", strJson, , "���·���ִ��״̬", , False, , , , True)
    Set colOut = gobjService.GetJsonListValue("output.data")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get����ִ����(ByVal strN As String, ByVal strOd As String, ByRef colOut As Collection) As Boolean
'���ܣ����ݵ��ݺ�+ҽ��ID����ȡ������ִ������
'������strN ���ݺŶ���ƴ����strOd ҽ��ID����ƴ��
    Dim strJson As String
    
    On Error GoTo errH
    
    If strN = "" Then Exit Function
    
    strJson = "{""input"":{""billtype"":3,""stuff_nos"":""" & strN & """,""order_ids"":""" & strOd & """}}"
    Call CallService("Zl_Stuffsvr_Getexecutednum", strJson, , "���·���ִ��״̬", , False, , , , True)
    Set colOut = gobjService.GetJsonListValue("output.item_list")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get��ҩ����(ByVal bln���� As Boolean, ByVal lng����ID As Long, ByVal strҩ����ϸ As String, ByRef colOut As Collection)
'���ܣ���ȡ��ҩ����
'������bln���� �Ƿ��Ǽ��ʵ� true �Ǽ��ʵ���false-�շѵ�
    Dim i As Long
    Dim varTmp As Variant
    Dim strItem As String
    Dim strJson As String
    Dim lng_billtype As Long
    
    On Error GoTo errH
    lng_billtype = IIF(bln����, 2, 1)
    varTmp = Split(strҩ����ϸ, ",")
    For i = 0 To UBound(varTmp)
        strItem = strItem & ",{""billtype"":" & lng_billtype
        strItem = strItem & ",""pharmacy_id"":" & varTmp(i)
        strItem = strItem & ",""pati_id"":" & lng����ID
        strItem = strItem & ",""valid_days"":null"
        strItem = strItem & ",""defaultwindow"":null"
        strItem = strItem & "}"
    Next
    strJson = "{""input"":{""item_list"":[" & Mid(strItem, 2) & "]}}"
    
    Call CallService("Zl_Drugsvr_Getsendwindows", strJson, , "Get��ҩ����", , False, , , , True)
    Set colOut = gobjService.GetJsonListValue("output.item_list")


    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
