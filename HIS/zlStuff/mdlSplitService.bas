Attribute VB_Name = "mdlSplitService"
Option Explicit

Public Function zlSplitService_CheckErrData(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strType As String, ByVal strInputById As String, ByVal strInputByNo As String, _
    ByRef colOutlist As Collection, ByRef colOutExpenseList As Collection, ByRef strErrMsg_Out As String) As Boolean
    '��ҩ/��ҩ���÷��񣺼������쳣״̬�������շѡ����״̬
    'strInputByid��1.������id���м�飬����id,����id...
    'strInputByNO: 2.������NO���м�飬��������:no1,no2|...
    '���Σ�1.���update_drug_status=2 then ��Ҫ���´��������շ�/���״̬���ٵ��÷��÷�����¶Է��ļǷ�ͬ��״̬
    '      2.���rcp_no_new<>"" then ��Ҫ���´�����NO������ID��
    '      3.���fee_status=2 then ���ܽ��з�ҩ����ҩ�Ȳ���
    Dim arrInput As Variant, i As Integer
    Dim strJson As String, StrJson_In As String, strJson_List As String
    
    On Error GoTo ErrHandler
    strErrMsg_Out = ""
    'Zl_Exsesvr_Checkerrordata
    '  --���ܣ����ݷ���NO�����ID����շѼ����쳣��Ϣ
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --      fee_type              C   1 �������'4'-���ģ�'5,6,7'-ҩƷ
    '  --      rcpdtl_ids            C   1 ������ϸids,����ö��ŷָ�
    '  --      bill_list[]                  ���飬����NO��Ϣ
    '  --         billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
    '  --         rcp_nos                C   1 ����Nos,����ö��ŷָ�
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --     code                   N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --     message                C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --     billid_list[]                   ������ID����ʱ����id�б�
    '  --        rcpdtl_id           N   1 ������ϸid
    '  --        fee_status          N   1 ����״̬�� 0-����,1-����
    '  --        cancel_status       N   1 ����״̬:0-����״̬,1-����ͬ����־�쳣
    '  --        update_status       N   1 �Ƿ�ͬ��״̬:0-����״̬,1-δ����ҩƷ/���ļ���״̬
    '  --     billno_list[]                 ��NO����ʱ����NO�б�
    '  --         billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
    '  --         rcp_no                 C   1 ����no
    '  --         fee_status             N   1 ����״̬������շ�ʱ,0-δ�շ�,1-���շ�,2-�쳣�շ�;��Լ���ʱ,0-����,1-����
    '  --         cancel_status          N   1 ����״̬:0-����״̬,1-����ͬ����־�쳣
    '  --         update_drug_status     N   1 �Ƿ�ͬ��״̬:0-����״̬,2-δ����ҩƷ/�����շ�״̬
    '  --     expense_list[]               ��ҩƷ����
    '  --         billtype               N   1 (ԭʼ)��������:1-�շѴ���;2-���ʴ���
    '  --         rcp_no                 C   1 (ԭʼ)����no
    '  --         rcpdtl_id              N   1 (ԭʼ)������ϸid
    '  --         rcp_no_new             C   1 �����ɵĴ���NO
    '  --         rcpdtl_id_new          N   1 �����ɴ�����ϸid
    '  --         pati_pageid            N   1  ��ҳID
    If strInputById = "" And strInputByNo = "" Then zlSplitService_CheckErrData = True: Exit Function
    
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_type", strType, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("rcpdtl_ids", strInputById, 0)
    If strInputByNo <> "" Then
        arrInput = Split(strInputByNo, "|")
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("billtype", Split(arrInput(i), ":")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("rcp_nos", Split(arrInput(i), ":")(1), 0)
            strJson_List = IIf(strJson_List = "", "", strJson_List & ",") & "{" & strJson & "}"
        Next
        strJson_List = ",""bill_list"":[" & strJson_List & "]"
    End If
    
    '����
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"

    If objServiceCall.CallService("Zl_Exsesvr_CheckErrorData", StrJson_In, , "", lngMode, False, , , , True) = False Then Exit Function
    If strInputById <> "" Then
        Set colOutlist = objServiceCall.GetJsonListValue("output.billid_list", "rcpdtl_id")
    Else
       Set colOutlist = objServiceCall.GetJsonListValue("output.billno_list", "billtype,rcp_no")
    End If
    Set colOutExpenseList = objServiceCall.GetJsonListValue("output.expense_list")
    zlSplitService_CheckErrData = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

Public Function zlSplitService_CisUpdateSyncState(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal intType As Integer, ByVal strInput As String, ByRef strErrMsg_Out As String) As Boolean
    '����ҽ��ͬ�����¼����
    'intType��1-���䣬2-ҩƷ ��3-����
    'strInput��ҽ��id,���ͺ�|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '�������list���ؼ���
    Dim varList As Variant   '����Ԫ��
    
    On Error GoTo ErrHandle
    strErrMsg_Out = ""
    '����Zl_CisSvr_UpdateSyncState
    '  ---------------------------------------------------------------------------
    '  --���ܣ�ͬ�����¼����
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --      order_list[]
    '  --          order_id          N 1 ҽ��id
    '  --          send_no           N 1 ���ͺ�
    '  --          sign_type         N 1 ���ñ��¼�����ͣ�
    '  --                                  ˵����1-���������¼
    '  --                                        2-��� ����ҩƷͬ�����
    '  --                                        3-��� ��������ͬ�����
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  ---------------------------------------------------------------------------
    
    If strInput = "" Then zlSplitService_CisUpdateSyncState = True: Exit Function
    
    If strInput <> "" Then
        arrInput = Split(strInput, "|")
        
        For i = 0 To UBound(arrInput)
            
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("order_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("send_no", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("sign_type", intType, 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
    End If
    StrJson_In = IIf(StrJson_In = "", "", StrJson_In & ",") & """order_list"":[" & strJson_List & "]"
    
    '����
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_CisSvr_UpdateSyncState"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, , "", lngMode, False, , , , True) = False Then Exit Function
    
    zlSplitService_CisUpdateSyncState = True
    Exit Function
ErrHandle:
    strErrMsg_Out = err.Description
End Function


Public Function zlSplitService_CheckAdviceaffirm(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOutput As String, ByRef strErrMsg_Out As String) As Boolean
    '����ָ������ID����ҳID���Һ�ID��ҽ��������Ϣ
    'strInput������id,��ҳID,�Һ�ID,�Һŵ���|...
    'strOutPut��ҽ��id,���ͺ�|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_List As String
    Dim strService As String
    Dim colOutlist As Collection, colOrderList As Collection
    Dim arrOrder As Variant
    
    On Error GoTo ErrHandle
    strErrMsg_Out = ""
    'Zl_CISSvr_GetAffirmErrorData
    ' ---------------------------------------------------------------------------
    '  --���ܣ��ٴ�ҽ��ִ�����ʱ�Զ���˷��ã��쳣��δ��ҩƷ�����Ľ����շ�ȷ�ϣ���Դ����쳣���ݻ�ȡ����
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --      pati_list[]���˹ؼ���Ϣ�����ڻ�ȡҽ��
    '  --           pati_id                    N 1 ����id
    '  --           pati_pageid                N 1 ��ҳid��סԺ���˴��룬���ﴫ0
    '  --           rgst_id                    N 1 �Һ�id�����ﲡ�˴��룬סԺ���˴���
    '  --           rgst_no                    C 1 �Һŵ���
    '  --����: Json_Out,��ʽ����
    '  --   output:
    '  --    code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --     pati_bill_list[]
    '  --         pati_id                      N 1 ����id
    '  --         pati_pageid                  N 0 ��ҳid��סԺ���˴��룬���ﴫ0
    '  --         rgst_id                      N 0 �Һ�id�����ﲡ�˴��룬סԺ���˴���
    '  --         rgst_no                      C 0 �Һŵ���
    '  --         order_ids                    C 1 ���쳣������ҽ��idƴ��
    '  --         fee_nos                      C 1 ���쳣�����е��ݺ�ƴ��
    '  --         order_list[]ҽ��������Ϣ�б�
    '  --             send_no                  N 1 ���ͺ�
    '  --             advice_id                N 1 ҽ��id
    '  --             fee_no                   C 1 ���ݺ�
    '  --             bill_prop                N 1 ��¼����
    '  --             outpati_account          N 1 �Ƿ�������� 0-����������ʣ�1-���������
    '  --             pati_source              N 1 ������Դ 1-����ҽ����2-סԺҽ��
   
    
    If strInput <> "" Then
        arrInput = Split(strInput, "|")
        
        For i = 0 To UBound(arrInput)
            StrJson_In = ""
            StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(arrInput(i), ",")(0), 1)
            StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageid", Split(arrInput(i), ",")(1), 1)
            StrJson_In = StrJson_In & "," & GetJsonNodeString("rgst_id", Split(arrInput(i), ",")(2), 1)
            StrJson_In = StrJson_In & "," & GetJsonNodeString("rgst_no", Split(arrInput(i), ",")(3), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & StrJson_In & "}"
            Else
                strJson_List = strJson_List & "," & "{" & StrJson_In & "}"
            End If
        Next
    End If
    StrJson_In = IIf(StrJson_In = "", "", StrJson_In & ",") & """pati_list"":[" & strJson_List & "]"
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_CISSvr_GetAffirmErrorData"

    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False, , , , True) = False Then Exit Function
    
    '����
    Set colOutlist = objServiceCall.GetJsonListValue("output.pati_bill_list")
    If colOutlist Is Nothing Then Exit Function
    
    'ֻ��Ҫ�������ݣ�ҽ��id,���ͺ�|...
    For i = 0 To colOutlist.Count - 1
        'ѭ��ȡ�ӽڵ�����
        Set colOrderList = objServiceCall.GetJsonListValue("output.pati_bill_list[" & i & "].order_list")
        
        For Each arrOrder In colOrderList
            strOutput = IIf(strOutput = "", "", strOutput & "|") & arrOrder("_advice_id") & "," & arrOrder("_send_no")
        Next
    Next

    zlSplitService_CheckAdviceaffirm = True
    Exit Function
ErrHandle:
    strErrMsg_Out = err.Description
End Function

Public Function GetJsonNodeString(ByVal strNodeName As String, ByVal strValue As String, Optional ByVal intType As Integer, _
    Optional ByVal blnZeroToNull As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡJson�ӵ㴮
    '���:strNodeName-�ӵ���
    '     strValue-ֵ
    '     intType-����:0-�ַ�;1-����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-09 18:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    
    strJson = Chr(34) & strNodeName & Chr(34)
    
    If intType = 0 Then
        strJson = strJson & ":" & Chr(34) & zlStr.ToJsonStr(strValue) & Chr(34)
    Else
        If strValue = "" Or (blnZeroToNull And Val(strValue) = 0) Then
            strJson = strJson & ":null"
        Else
            strJson = strJson & ":" & IIf(Mid(strValue, 1, 1) = ".", "0", "") & strValue
        End If
    End If
        
    GetJsonNodeString = strJson
End Function

Public Function zlSplitService_CallAccountVerify(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strMain As String, ByVal strInput As String) As Boolean
    '���Ϻ������˻��ۼ��˵�
    '��Σ�
    '   strMain������ʱ��,����Ա����,����Ա���
    '   strInput��������Դ|���ݺ�|����ID|���s||������Դ|���ݺ�|����ID|���s||...
    Dim arrInput As Variant, arrItem As Variant, arrMain As Variant
    Dim i As Integer, StrJson_In As String
    Dim strJson_bill As String, strJson_item As String, strJson_List As String
    
    'Zl_Exsesvr_Billverify
    '  --���ܣ����õ������
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --    operator_time         C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '  --    operator_name         C 1 ����Ա����
    '  --    operator_code         C 1 ����Ա���
    '  --    item_list
    '  --        fee_source        N 1 ������Դ:1-����;2-סԺ
    '  --        fee_no            C 1 ���õ��ݺ�
    '  --        serial_nums       C 0 ��Ŵ���������ʾ���ŵ���
    '  --        pharmacy_window   C 0 ��ҩ���ڣ�������ԴΪ����ʱ���룬��ʽ���ⷿID1:��ҩ����1,�ⷿID2:��ҩ����2,....
    '  --        pati_id           N 0 ����id��������ԴΪסԺ�Ұ��������ʱ����(��Ҫ��Լ��ʱ�)
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    If strInput = "" Then Exit Function
    
    '������Դ|���ݺ�|����ID|���s||������Դ|���ݺ�|����ID|���s||...
    arrInput = Split(strInput, "||")
    strJson_List = ""
    
    For i = 0 To UBound(arrInput)
        arrItem = Split(arrInput(i), "|")
        
        strJson_item = ""
        strJson_item = strJson_item & "" & GetJsonNodeString("fee_source", arrItem(0), 1)
        strJson_item = strJson_item & "," & GetJsonNodeString("fee_no", arrItem(1), 0)
        strJson_item = strJson_item & "," & GetJsonNodeString("pati_id", arrItem(2), 1)
        strJson_item = strJson_item & "," & GetJsonNodeString("serial_nums", arrItem(3), 0)
        
        strJson_List = IIf(strJson_List = "", "", strJson_List & ",") & "{" & strJson_item & "}"
    Next
    
    arrMain = Split(strMain, ",")
    strJson_bill = ""
    strJson_bill = strJson_bill & "" & GetJsonNodeString("operator_time", arrMain(0), 0)
    strJson_bill = strJson_bill & "," & GetJsonNodeString("operator_name", arrMain(1), 0)
    strJson_bill = strJson_bill & "," & GetJsonNodeString("operator_code", arrMain(2), 0)
    
    StrJson_In = "{""input"":{" & strJson_bill & ",""item_list"":[" & strJson_List & "]" & "}}"
        
    '���÷���
    If objServiceCall.CallService("Zl_Exsesvr_Billverify", StrJson_In, , "", lngMode, False, , , , True) = False Then
        MsgBox "���á�Zl_Exsesvr_Billverify��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    zlSplitService_CallAccountVerify = True
End Function

Public Function zlSplitService_CallCheckDrugById_bak(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, colOutlist As Collection) As Boolean
    '��ҩ/��ҩ���÷��񣺼������쳣״̬�������շѡ����״̬
    '������id���м��
    'strInput�� ����id,����id...
    '���Σ�1.���update_stuff_status=2 then ��Ҫ���´��������շ�/���״̬���ٵ��÷��÷�����¶Է��ļǷ�ͬ��״̬
    '      2.���fee_status=2 then ���ܽ��з�ҩ����ҩ�Ȳ���
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '�������list���ؼ���
    Dim varList As Variant   '����Ԫ��
    
    On Error GoTo ErrHandle
    
    '����Zl_Exsesvr_In_Checkstuff
    
'  ---------------------------------------------------------------------------
'  --���ܣ����ݷ���ID��鲡��סԺ�쳣������Ϣ
'  --��Σ�Json_In:��ʽ
'  --  input
'  --      stuff_ids             C   1 ������ϸids,����ö��ŷָ�
'  --����: Json_Out,��ʽ����
'  --  output
'  --     code                   N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
'  --     message                C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --     bill_list[]
'  --        stuff_id            N   1 ���ϴ�����ϸid
'  --        fee_status          N   1 ����״̬�� 0-����,1-����
'  --        cancel_status       N   1 ����״̬:0-����״̬,1-����ͬ����־�쳣
'  --        update_stuff_status N   1 �Ƿ�ͬ��״̬:0-����״̬,1-δ����ҩƷ/���ļ���״̬
'  ------------------------------------------------------------------------------------------------------------
    
    If strInput = "" Then zlSplitService_CallCheckDrugById_bak = True: Exit Function
    
    strJson_List = ""
    If strInput <> "" Then
        strJson = ""
        strJson = strJson & "" & GetJsonNodeString("stuff_ids", strInput, 0)
    End If
    
    '����
    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "Zl_Exsesvr_In_Checkstuff"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_In_Checkstuff��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallCheckDrugById_bak = False
        Exit Function
    End If
    
    Set colOutlist = objServiceCall.GetJsonListValue("output.bill_list", "stuff_id")
    
    zlSplitService_CallCheckDrugById_bak = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function zlSplitService_CallCheckDrugByNo_bak(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, colOutlist As Collection, colOutExpenseList As Collection) As Boolean
    '��ҩ/��ҩ���÷��񣺼������쳣״̬�������շѡ����״̬
    '������NO���м��
    'strInput����������:no1,no2|...
    '���Σ�1.���update_stuff_status=2 then ��Ҫ���´��������շ�/���״̬���ٵ��÷��÷�����¶Է��ļǷ�ͬ��״̬
    '      2.���rcp_no_new<>"" then ��Ҫ���´�����NO������ID��
    '      3.���fee_status=2 then ���ܽ��з�ҩ����ҩ�Ȳ���
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '�������list���ؼ���
    Dim varList As Variant   '����Ԫ��
    
    On Error GoTo ErrHandle
    
    '����Zl_Exsesvr_Out_Checkdrug
    
'  --���ܣ����ݵ������ͺ�NO�ż�鲡�������쳣������Ϣ
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    bill_list[]
'  --         billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
'  --         rcp_nos                C   1 ����Nos,����ö��ŷָ�
'  --����: Json_Out,��ʽ����
'  --  output
'  --     code                       N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
'  --     message                    C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --     bill_list[]
'  --         billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
'  --         rcp_no                 C   1 ����no
'  --         fee_status             N   1 ����״̬������շ�ʱ,0-δ�շ�,1-���շ�,2-�쳣�շ�;��Լ���ʱ,0-����,1-����
'  --         cancel_status          N   1 ����״̬:0-����״̬,1-����ͬ����־�쳣
'  --         update_stuff_status     N   1 �Ƿ�ͬ��״̬:0-����״̬,2-δ����ҩƷ/�����շ�״̬
'  --     expense_list[]
'  --         rcp_no                 C   1 (ԭʼ)����no
'  --         rcpdtl_id              N   1 (ԭʼ)����id
'  --         rcp_no_new             C   1 �����ɵĴ���NO
'  --         rcpdtl_id_new          N   1 �����ɴ���id
    
    strJson_List = ""
    If strInput <> "" Then
        arrInput = Split(strInput, "|")
        
        For i = 0 To UBound(arrInput)
            
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("billtype", Split(arrInput(i), ":")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("rcp_nos", Split(arrInput(i), ":")(1), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""bill_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = strJson_List
    
    '����
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Out_Checkstuff"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Out_Checkstuff��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallCheckDrugByNo_bak = False
        Exit Function
    End If
    
    Set colOutlist = objServiceCall.GetJsonListValue("output.bill_list", "billtype,stuff_no")
    Set colOutExpenseList = objServiceCall.GetJsonListValue("output.expense_list")
    
    zlSplitService_CallCheckDrugByNo_bak = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function zlSplitService_UpdateExseInfo(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strBase As String, ByVal strItemList As String, ByVal strDeptchangeList As String, ByVal strDelNoList As String) As Boolean
    '��ҩ/��ҩ���÷��񣺸��·���ִ�в��š�ִ���ˡ���ҩ���ڼ�ִ��״̬����Ϣ
    'strBase��������Ϣ����ʽ ������Դ,����Ա����,����Ա����,����ʱ��
    'strItemList�����·���ID�б���ʽ ����ID,��ִ������(��ҩ����),ִ����,ִ��ʱ��,��ҩ����|...
    '             ��ʱ�⼸��ֵ����ҩƷ�����ж���ѯ�������루�򰴹��ܴӽ�������֯������ı䷢ҩ���ڵȣ�
    'strDeptchangeList������ִ�в��� ��ʽ ����id,ԭִ�в���id,��ִ�в���id|...
    'strDelNoList����ҩ�Զ����� ��ʽ NO;���(���1:����:ִ��״̬1,���2:����2:ִ��״̬2,...);����״̬|...
    Dim arrInput As Variant, strPart As String
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim strִ���� As String, strִ��ʱ�� As String
    
    On Error GoTo ErrHandle
    
    'Zl_Exsesvr_Updateexeinfo
'    ---------------------------------------------------------------------------
'  --���ܣ����·���ִ�в��š�ִ���ˡ���ҩ���ڼ�ִ��״̬����Ϣ
'  --��Σ�Json_In:��ʽ
'  --input
'  --  fee_origin            N  1  ������Դ(Ĭ��=2��1-������ã�2-סԺ����)
'  --  operator_code         C     ����Ա����
'  --  operator_name         C     ����Ա����
'  --  operator_time         C     ����ʱ��
'  --  item_list                   ���б����ִ�������Ϣ�������б�ʱͬʱ��Ҫ����fee_origin
'  --    fee_id              C  1  ����id
'  --    exe_nums            N  1  ��ִ������:Ϊ0��ʾ��δִ��
'  --    exe_people          C     ִ����:����ִ�л���ȫִ��ʱ����Ҫ���룬������ʱ����operator_nameΪ׼
'  --    exe_time            D     ִ��ʱ��:yyyy-mm-dd hh24:mi:ss,:����ִ�л���ȫִ��ʱ����Ҫ���룬������ʱ����"create_time"Ϊ׼
'  --    pharmacy_window     C     ��ҩ����:ҩƷ��������Ч,�޴˽ӵ㣬������·�ҩ����
'  --  deptchange_List       C  1  ִ�п��ұ����Ϣ�б�
'  --    fee_id              C  1  ����id
'  --    exe_old_deptid      N     ԭִ�п���ID
'  --    exe_deptid          N  1  ִ�в���id
'  --  delrcp_list           C     [����]ȡ����Һʱ����Ҫͬ������
'  --    rcp_no              C  1  ����no
'  --    serial_nums         C  1  ��ʽ: ���1:����:ִ��״̬1,���2:����2:ִ��״̬2,...
'  --    operator_status     N     ����״̬��0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
'  --����: Json_Out,��ʽ����
'  --output
'  --   code                 C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --   message              C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  ---------------------------------------------------------------------------

    
    '1.������Ϣ
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_origin", Split(strBase, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_code", Split(strBase, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_name", Split(strBase, ",")(2), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_time", Split(strBase, ",")(3), 0)
        
    '2.���·���id��Ӧ�ķ�����Ϣ
    strJson_List = ""
    If strItemList <> "" Then
        arrInput = Split(strItemList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("fee_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_nums", Split(arrInput(i), ",")(1), 1)
            
            'ִ�������Ϊ���򲻴���ýڵ�
            If Split(arrInput(i), ",")(2) <> "" Then
                If Split(arrInput(i), ",")(2) = "null" Then
                    '����ģ�����������null������մ�������ȡ����ҩʱ���ִ����
                    strJson = strJson & "," & GetJsonNodeString("exe_people", "", 0)
                Else
                    strJson = strJson & "," & GetJsonNodeString("exe_people", Split(arrInput(i), ",")(2), 0)
                End If
            End If
            
            'ִ��ʱ�����Ϊ���򲻴���ýڵ�
            If Split(arrInput(i), ",")(3) <> "" Then
                strJson = strJson & "," & GetJsonNodeString("exe_time", Split(arrInput(i), ",")(3), 0)
            End If
            
            '��ҩ�������Ϊ���򲻴��ýڵ�
            If Split(arrInput(i), ",")(4) <> "" Then
                strJson = strJson & "," & GetJsonNodeString("pharmacy_window", Split(arrInput(i), ",")(4), 0)
            End If
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""item_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '3.�������·���id��Ӧ��ִ�в���id
    strJson_List = ""
    If strDeptchangeList <> "" Then
        arrInput = Split(strDeptchangeList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("fee_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_old_deptid", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("exe_deptid", Split(arrInput(i), ",")(2), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""deptchange_List"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '4.������������
    strJson_List = ""
    If strDelNoList <> "" Then
        arrInput = Split(strDelNoList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcp_no", Split(arrInput(i), ";")(0), 0)
            strJson = strJson & "," & GetJsonNodeString("serial_nums", Split(arrInput(i), ";")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("operator_status", Split(arrInput(i), ";")(2), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""delrcp_list"":[" & strJson_List & "]"
    End If
    StrJson_In = StrJson_In & strJson_List
    
    '����
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Updateexeinfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Updateexeinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_UpdateExseInfo = False
        Exit Function
    End If

    zlSplitService_UpdateExseInfo = True

    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_WriteOff(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strBase As String, ByVal strAccAudit As String) As Boolean
    '��ҩͬʱ����ʱ���õ��������񣺺ϲ�����������ˣ�����������ˣ�����/סԺ��¼ɾ�������·��ü�¼״̬�����ʼ��ȹ���
    '�ϲ�Ϊһ�����������Ч���Ʒ�ҩ/��ҩ����
    'strBase��������Ϣ����ʽ ������Դ,����Ա����,����Ա����,����ʱ��
    'strAccAudit�������б���ʽ ����ID,����ʱ��,��������,�������,��������,�ѷ�����|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim arrList As Variant, arrPart As Variant
    Dim strPart1 As String, strPart2 As String, strPart3 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String
    Dim n As Integer
    Dim strJson_Part As String, strCheckJson_In As String
    
    On Error GoTo ErrHandle
    
'    'Zl_Exsesvr_Drugwriteoff
'      ---------------------------------------------------------------------------
'  --���ܣ�ҩƷ�����ķ�������(�������ͨ��������ܾ���ȡ���ܾ�)
'  --��Σ�Json_In:��ʽ
'  --input
'  --  fee_origin            N  1  ������Դ��1-���2-סԺ��
'  --  operator_code         C  1  ����Ա����
'  --  operator_name         C  1  ����Ա����
'  --  operator_time         C     ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --  rcpdtl_list                 [����]ÿ��������ϸ��Ϣ
'  --    rcpdtl_id           N  1  ������ϸid(����id)
'  --    request_time        D  1  ����ʱ��
'  --    oper_type           N  1  ��������:0-���ͨ��;1-��˲�ͨ�� 2-��˾ܾ� 3-ȡ���ܾ�;
'  --    request_type        N  1  �������Ĭ�ϴ�1��
'  --    quantity            N  1  ��������
'  --    sended_num          N  1  �ѷ�����
'
'  --����: Json_Out,��ʽ����
'  --output
'  --   code                          C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --   message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  ---------------------------------------------------------------------------
    
    '1.������Ϣ
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_origin", Split(strBase, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_code", Split(strBase, ",")(1), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_name", Split(strBase, ",")(2), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("operator_time", Split(strBase, ",")(3), 0)
    
    '2.������Ϣ
    strJson_List = ""
    If strAccAudit <> "" Then
        arrInput = Split(strAccAudit, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("rcpdtl_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("request_time", Split(arrInput(i), ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("oper_type", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("request_type", Split(arrInput(i), ",")(3), 1)
            strJson = strJson & "," & GetJsonNodeString("quantity", Split(arrInput(i), ",")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("sended_num", Split(arrInput(i), ",")(5), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""rcpdtl_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
     
    '����
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Drugwriteoff"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Drugwriteoff��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_WriteOff = False
        Exit Function
    End If
    
    zlSplitService_WriteOff = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_WriteOffCheck(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strBase As String, ByVal strAccAudit As String, ByVal strPatiList As String, ByRef strCheckMsg As String) As Boolean
    '��ҩͬʱ����ʱ���õ������������ʼ��
    '�ϲ�Ϊһ�����������Ч���Ʒ�ҩ/��ҩ����
    'strBase��������Ϣ����ʽ ��ֹ��������,������Դ
    'strAccAudit�������б���ʽ  ����ID,����ʱ��,��������,�������,��������,�ѷ�����|...
    'strPatiList�������б���ʽ  ����id,������˱�־,סԺ״̬,������Ŀ����|...
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim arrList As Variant, arrPart As Variant
    Dim strPart1 As String, strPart2 As String, strPart3 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String
    Dim n As Integer
    Dim strJson_Part As String, strCheckJson_In As String
    Dim colOutlist As New Collection
    
    On Error GoTo ErrHandle
''    'Zl_Exsesvr_Drugwriteoff_Check
'  ---------------------------------------------------------------------------
'  --���ܣ�ҩƷ�����ķ���������˼��
'  --��Σ�Json_In:��ʽ
'  --input      ҩƷ��������ǰ���
'  --  part_ban_writeoffs    N  1  ��ֹ��������:0-����;1-������������(�����ŵ��ݵĲ��ֻ�ĳ�ʵĲ���)
'  --  fee_origin            N  1  ������Դ:1-���2-סԺ
'  --  rcpdtl_list[]               ���������б�
'  --    oper_type           N  1  ��������:0-���ͨ�� 1-��˲�ͨ�� 2-��˾ܾ� 3-ȡ���ܾ�;
'  --    rcpdtl_id           N  1  ������ϸID(����ID)
'  --    request_time        D     ����ʱ��
'  --    request_type        N     �������ȱʡΪ1
'  --    quantity            N  1  ����������Ϊ���nullʱ,������ID��������ֱ������
'  --    sended_num          N  1  �ѷ�����
'  --  pati_list[]                 ������Ϣ
'  --    pati_id             N     ����ID,ΪNULL��0ʱ����ʾ���ŵ���
'  --    fee_audit_status    N     ������˱�־:0���-δ���;1-����˻�ʼ���;2-������,��Ͻ���Ȩ��
'  --    si_inp_status       N     סԺ״̬:0-����סԺ;1-��δ���;2-����ת�ƻ�����ת����;3-��Ԥ��Ժ
'  --    catalog_date        C     ������Ŀ���ڣ�yyyy-mm-dd hh24:mi:ss
'  --����: Json_Out,��ʽ����
'  --output
'  --   code                          C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --   message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --  tip_list[]  C  1  ��ʾ�б�:��Ҫ�ǿ��ܴ��ڶ����ʾѯ�ʷ�ʽ���������б�,��ֹʱ������һ����Ϣ
'  --    tip_mode  C  1  ���Ʒ�ʽ:1-��ʾѯ��;2-��ֹ
'  --    tip_message  C  1  ��ʾ��Ϣ
'  ---------------------------------------------------------------------------
    
    '1.������Ϣ
    StrJson_In = StrJson_In & "" & GetJsonNodeString("part_ban_writeoffs", Split(strBase, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("fee_origin", Split(strBase, ",")(1), 1)
    
    '2.������Ϣ
    strJson_List = ""
    If strAccAudit <> "" Then
        arrInput = Split(strAccAudit, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""            '�����б�����ID,����ʱ��,��������,�������,��������,�ѷ�����|...
            strJson = strJson & "" & GetJsonNodeString("rcpdtl_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("request_time", Split(arrInput(i), ",")(1), 0)
            strJson = strJson & "," & GetJsonNodeString("oper_type", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("request_type", Split(arrInput(i), ",")(3), 1)
            strJson = strJson & "," & GetJsonNodeString("quantity", Split(arrInput(i), ",")(4), 1)
            strJson = strJson & "," & GetJsonNodeString("sended_num", Split(arrInput(i), ",")(5), 1)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""rcpdtl_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
    
    '3.������Ϣ
    strJson_List = ""
    If strPatiList <> "" Then
        arrInput = Split(strPatiList, "|")
        
        For i = 0 To UBound(arrInput)
            strJson = ""            '����id,������˱�־,סԺ״̬,������Ŀ����|...
            strJson = strJson & "" & GetJsonNodeString("pati_id", Split(arrInput(i), ",")(0), 1)
            strJson = strJson & "," & GetJsonNodeString("fee_audit_status", Split(arrInput(i), ",")(1), 1)
            strJson = strJson & "," & GetJsonNodeString("si_inp_status", Split(arrInput(i), ",")(2), 1)
            strJson = strJson & "," & GetJsonNodeString("catalog_date", Split(arrInput(i), ",")(3), 0)
            
            If strJson_List = "" Then
                strJson_List = "{" & strJson & "}"
            Else
                strJson_List = strJson_List & "," & "{" & strJson & "}"
            End If
        Next
        
        strJson_List = ",""pati_list"":[" & strJson_List & "]"
    End If
    
    StrJson_In = StrJson_In & strJson_List
     
    '����
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Drugwriteoff_Check"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Drugwriteoff_Check��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        '�������ʧ��ʱ
        strCheckMsg = objServiceCall.GetJsonNodeValue("output.message")
        zlSplitService_WriteOffCheck = False
        Exit Function
    Else
        '������óɹ�������ҵ������
        Set colOutlist = objServiceCall.GetJsonListValue("output.tip_list")
        If colOutlist.Count > 0 Then
            '���Ʒ�ʽ,��ʾ��Ϣ
            strCheckMsg = colOutlist(1)(1) & "," & colOutlist(1)(2)
        End If
    End If
    
    zlSplitService_WriteOffCheck = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function zlSplitService_CallUpdateSynchrosign(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal intType As Integer, strInputById As String, ByRef strErrMsg_Out As String) As Boolean
    '���·��üǷ�ͬ�����
    '������NO���и���
    'intType��0-���¼Ƿ�ͬ����־��1-����ת��ͬ����־
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '�������list���ؼ���
    Dim varList As Variant   '����Ԫ��
    
    On Error GoTo ErrHandle
    strErrMsg_Out = ""
    If strInputById = "" Then zlSplitService_CallUpdateSynchrosign = True: Exit Function
    '����Zl_Exsesvr_Sync_Update
    '  --���ܣ������շ�ͬ����־
    '  --��Σ�Json_In:��ʽ
    '  --  input
    '  --    sign_type           N 1 ��־���ͣ�0-�Ƿ�ͬ����־,1-ת��ͬ����־
    '  --    detail_ids  C  1  ������ϸid��(����id��),֧�ֶ��id���á�,���ָ�
    '  --    bill_list[]
    '  --      billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
    '  --      rcp_no                 C   1 ����No
    '  --����: Json_Out,��ʽ����
    '  --  output
    '  --     code                   N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    '  --     message                C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("sign_type", intType, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("detail_ids", strInputById, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Sync_Update"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False, , , , True) = False Then Exit Function
    
    zlSplitService_CallUpdateSynchrosign = True
    Exit Function
ErrHandle:
    strErrMsg_Out = err.Description
End Function

Public Function zlSplitService_GetAdvice(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strAdvice_ids As String, _
    ByRef colAdvice As Collection) As Boolean
    'ȡҽ����Ϣ
    'strRcpdtl_ids��ҽ��id,ҽ��id...
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("advice_ids", strAdvice_ids, 0)
'    strJson_In = strJson_In & "," & GetJsonNodeString("advice_ids", strPatis, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "zl_CisSvr_GetAdviceInfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�zl_CisSvr_GetAdviceInfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '��������
    Set colAdvice = objServiceCall.GetJsonListValue("output.advice_list")
    
    If colAdvice Is Nothing Then Exit Function
    
    zlSplitService_GetAdvice = True
End Function







Public Function zlSplitService_GetCloseAccount(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal colInput As Collection, _
    ByRef colCloseAccount As Collection) As Boolean
    'ȡ���ʼ�¼
    'colInput����ѯ������ϣ�Json��input���ڵ���ΪԪ�ص�KEYֵ������ĳԪ��Ϊ�ձ�ʾ�ýڵ�ֵΪ��
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    
    On Error GoTo ErrHandle
    
'    input           ����������ѯ����������Ϣ
'        audit_dept_id   N       ��˲���ID(ҩ��)
'        request_begin_time  D       ���뿪ʼʱ��
'        request_end_time    D       �������ʱ��
'        audit_begin_time    D       ��˿�ʼʱ��
'        audit_end_time  D       ��˽���ʱ��

'        cancel_status    N   1   ״̬
'        request_dept_id N       ���벿��ID
'        request_operator    C       ������
'        pati_id  N       ����ID
'        cancel_condition    C       ��������

'        cancel_check    N       �˲飨ѡ�����������������Ҫ�˲顿ʱ���룬0-δ�˲� 1-�Ѻ˲飩
'        rcpdtl_id   C       ������ϸid,[����]��[1,2,3]
'        request_dept_ids   C     ���벿��id��������������ѯ
'        item_ids           C     �շ�ϸĿid��,����������ѯ
'    output
'        code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'        message C   1   "Ӧ����Ϣ��
'        fee_cancel_list         [����]����������ÿ���������ʼ�¼
'        rcpdtl_id N       ������ϸid(����id)
'        request_type  N       �������
'        item_id   N       �շ�ϸĿid
'        request_dept_id   N       ���벿��id
'        request_dept  C       ���벿��
'        audit_dept_id N       ��˲���id
'        quantity  N       ����
'        request_operator  C       ������
'        request_time  D       ����ʱ��
'        auditor   C       �����
'        audit_time    D       ���ʱ��
'        cancel_status N       ״̬
'        cancel_reason C       ����ԭ��
'        checker   C       �˲���
'        price_retail  N       ���ۼ�
'        advice_id N       ҽ��id
'        pati_id   N       ����ID
'        pati_name C       ��������
'        inpatient_num C       סԺ��

    '���
    StrJson_In = ""
    If ExistsColObject(colInput, "audit_dept_id") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("audit_dept_id", colInput("audit_dept_id"), 1)
    If ExistsColObject(colInput, "request_begin_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_begin_time", colInput("request_begin_time"), 0)
    If ExistsColObject(colInput, "request_end_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_end_time", colInput("request_end_time"), 0)
    If ExistsColObject(colInput, "audit_begin_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("audit_begin_time", colInput("audit_begin_time"), 0)
    If ExistsColObject(colInput, "audit_end_time") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("audit_end_time", colInput("audit_end_time"), 0)
    
    If ExistsColObject(colInput, "cancel_status") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_status", colInput("cancel_status"), 1)
    If ExistsColObject(colInput, "request_dept_id") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_dept_id", colInput("request_dept_id"), 1)
    If ExistsColObject(colInput, "request_operator") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_operator", colInput("request_operator"), 0)
    If ExistsColObject(colInput, "pati_id") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_id", colInput("pati_id"), 1)
    If ExistsColObject(colInput, "cancel_condition") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_condition", colInput("cancel_condition"), 0)
    
    If ExistsColObject(colInput, "cancel_check") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_check", colInput("cancel_check"), 1)
    If ExistsColObject(colInput, "rcpdtl_id") Then StrJson_In = StrJson_In & "," & """rcpdtl_id"":" & "[" & colInput("rcpdtl_id") & "]"
    If ExistsColObject(colInput, "request_dept_ids") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("request_dept_ids", colInput("request_dept_ids"), 0)
    If ExistsColObject(colInput, "item_ids") Then StrJson_In = StrJson_In & "," & GetJsonNodeString("item_ids", colInput("item_ids"), 0)
    
    If StrJson_In = "" Then Exit Function
    
    StrJson_In = Mid(StrJson_In, 2)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_ExseSvr_GetChargeOffInfo"
    
    '����
'    StrJson_In = "{""input"":{""audit_dept_id"":305,""cancel_status"":0,""rcpdtl_id"":[4528923,4528923]}}"
'    StrJson_In = "{""input"":{""audit_dept_id"":305,""cancel_status"":0,""rcpdtl_id"":[1]}}"
'    strService = "Zl_ExseSvr_GetChargeOffInfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_ExseSvr_GetChargeOffInfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colCloseAccount = objServiceCall.GetJsonListValue("output.fee_cancel_list")
    
    If colCloseAccount Is Nothing Then Exit Function
    
    zlSplitService_GetCloseAccount = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_GetFee(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strRcpdtl_ids As String, _
    ByRef colPati As Collection, Optional ByVal strKeyNodes As String) As Boolean
    'ȡ������Ϣ
    'strRcpdtl_ids������id,����id...
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    
    'Zl_Exsesvr_GetBillDetailInfo
'  -------------------------------------------------------------------------------------------------
'  --���ܣ���ȡ��ҩƷ��ҩҵ����صķ�����Ϣ����Ҫ���ڽ�����ʾ
'  --��Σ�json��ʽ
'  --Input
'  --   fee_ids    C     ����id��֧�ֶ��id����ʽ�� ����id,����id,��
'  --   bill_nos   C     ����no,��¼���ʣ���ʽ: no,��¼����|,...
'  --���Σ�json��ʽ
'  --Json_Out
'  --fee_list      [����]ÿ������ID��Ϣ
'  --  bill_prop           N    ��¼����:1-�շѵ�;2-���ʵ�;3-�Զ����ʵ�;4-�Һŵ�;5-���￨;6-Ԥ����
'  --  bill_no             C    ���ݺ�
'  --  fee_id              N    ������ϸid(����id)
'  --  fee_num             N    ���
'  --  iden_id             N    ��ʶ��
'  --  pati_bed            C    ����
'  --  fee_ampaid          N    ʵ�ս��
'  --  packages_num        N    ����
'  --  quantity            N    ����
'  --  placer              C    ������
'  --  operator_code       C    ����Ա���
'  --  operator_name       C    ����Ա����
'  --  create_time         D    �Ǽ�ʱ��
'  --  happen_time         D    ����ʱ��
'  --  rcp_type            N    �������(������NO��˵��1-��ҩ��2-��ҩ��3-���)
'  --  fee_type            C    �ѱ�
'  --  rec_status          N    ��¼״̬
'  --  register_id         N    �Һ�id
'  --  register_no         C    �Һ�NO
'  --  register_time       D    �ҺŵǼ�ʱ��
'  --  income_item_id      N    ������Ŀid
'  --  fee_origin          N    ������Դ(1-������ã�2-סԺ����)
'  --  bill_deptid         N    ��������id
'  --  order_id            N    ҽ��ID
'  --  fee_item_id         N    �շ�ϸĿid
'  --  fee_status         N    ����״̬
'  -------------------------------------------------------------------------------------------------
  
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("fee_ids", strRcpdtl_ids, 0)
'    strJson_In = strJson_In & "," & GetJsonNodeString("fee_ids", strPatis, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Exsesvr_GetBillDetailInfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Exsesvr_GetBillDetailInfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '��������
    Set colPati = objServiceCall.GetJsonListValue("output.fee_list", strKeyNodes)
    
    If colPati Is Nothing Then Exit Function
    
    zlSplitService_GetFee = True
End Function

Public Function zlSplitService_GetNOByInvoice(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
     ByRef colOutlist As Collection) As Boolean
    'ͨ��Ʊ�ݺ�ȡ����NO
    'strInput��Ʊ�ݺ�
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim strOutNos As String
    Dim n As Integer
    
    'zl_ExseSvr_GetNoByInvoice
'  -------------------------------------------------------------------------------------------------
'  --���ܣ���Ʊ�ݺŷ�ҩ����ҩ��ͨ��¼�뷢Ʊ�Ż�ȡ��Ӧ��ҩƷ����NO
'  --��Σ�json��ʽ
'  --Input
'  --   invc_no  C  1  Ʊ�ݺ�
'  --���Σ�json��ʽ
'  --Json_Out
'  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message  C  1  "Ӧ����Ϣ�� �ɹ�ʱ���ش���No��[����] ʧ��ʱ���ؾ���Ĵ�����Ϣ"
'  --  rcp_nos  C  1 �������ݺţ�����ö��ŷָ�
'  -------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandle
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("invc_no", strInput, 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
        
    strService = "zl_ExseSvr_GetNoByInvoice"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�zl_ExseSvr_GetNoByInvoice��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    strOutNos = objServiceCall.GetJsonNodeValue("output.rcp_nos")
    
    If strOutNos <> "" Then
        For n = 0 To UBound(Split(strOutNos, ","))
            colOutlist.Add strOutNos
        Next
    End If

    zlSplitService_GetNOByInvoice = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetExseSpec(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
     ByRef strOutput As String) As Boolean
    '��ȡ���������õĹ�񣨲��ϣ�
    'strInput������id
    'strOutPut������id�����û���򷵻�0
    Dim StrJson_In As String
    
    On Error GoTo ErrHandle
    'Zl_Exsesvr_Getexsespec
    '  --���ܣ����ù���Ƿ���������ü�¼
    '  --input   ���ݲ���id����Ƿ���������ü�¼
    '  --  item_id       N   1   �շ�ϸĿid
    '  --output
    '  --  code          C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --  message       C   1   Ӧ����Ϣ��
    '  --  item_id       N   1   �շ�ϸĿid
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("item_id", Val(strInput), 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Getexsespec", StrJson_In, "", "", lngMode) = False Then
        MsgBox "���á�Zl_Exsesvr_Getexsespec��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    strOutput = Nvl(objServiceCall.GetJsonNodeValue("output.item_id"))
    zlSplitService_GetExseSpec = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_StuffAdjustPriceType(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    Optional ByVal strKey As String) As Boolean
    '���ļ۸����Ե���ʱ�����ĵ���ӯ���Ϳ��仯���ݴ���
    'strInput:����ID��ԭ�۸����ͣ��¼۸�����|...
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    Dim arrPart As Variant
    Dim strJson_Part As String, strJson_List As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_AdjustPriceType
'  ---------------------------------------------------------------------------
'  --input      ���ļ۸����Ե���ʱ�����ĵ���ӯ���Ϳ��仯���ݴ���
'  --    item_list[]         �����б�
'  --       stuff_id      N    ҩƷid
'  --       price_type_old    N    ԭ�۸����ͣ�0-���ۣ�1-ʱ��
'  --       price_type_new    N    �¼۸����ͣ�0-���ۣ�1-ʱ��
'  --output
'  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message  C  1  Ӧ����Ϣ��
'  ---------------------------------------------------------------------------
  
    If strInput = "" Then Exit Function
    
    '���
    strJson_List = ""
    arrPart = Split(strInput, "|")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("stuff_id", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("price_type_old", Split(arrPart(i), ",")(1), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("price_type_new", Split(arrPart(i), ",")(2), 1)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    StrJson_In = """item_list"":[" & strJson_List & "]"
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_AdjustPriceType"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_StuffSvr_AdjustPriceType��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
   
    zlSplitService_StuffAdjustPriceType = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_StuffExistRec(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal lng����ID As Long, _
    ByRef intExist As Integer, Optional ByVal strKey As String) As Boolean
    '��ȡָ���������Ƿ�����շ���¼
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_CheckStuffExistRec
'  ---------------------------------------------------------------------------
'  --input      �ж������Ƿ�����շ���¼
'  --  stuff_id      N    ����id
'  --output
'  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message  C  1  Ӧ����Ϣ��
'  --  isexist  N 1 �Ƿ����: 1-����;0-������
'  ---------------------------------------------------------------------------
  
    If lng����ID = 0 Then Exit Function
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("stuff_id", lng����ID, 1)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_CheckStuffExistRec"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_StuffSvr_CheckStuffExistRec��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '����ֵ
    intExist = Nvl(objServiceCall.GetJsonNodeValue("output.isexist"), 0)
    
    zlSplitService_StuffExistRec = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_HightCostStuffExistRec(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal lng����ID As Long, _
    ByRef intExist As Integer, Optional ByVal strKey As String) As Boolean
    '�жϸ�ֵ�����Ƿ����ʹ�ü�¼
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_CheckHCostExistRec
'  ---------------------------------------------------------------------------
'  --input      �жϸ�ֵ�����Ƿ����ʹ�ü�¼
'  --  stuff_id      N    ����id
'  --output
'  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message  C  1  Ӧ����Ϣ��
'  --  isexist  N 1 �Ƿ����: 1-����;0-������
'  ---------------------------------------------------------------------------
  
    If lng����ID = 0 Then Exit Function
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("stuff_id", lng����ID, 1)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_CheckHCostExistRec"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_StuffSvr_CheckHCostExistRec��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '����ֵ
    intExist = Nvl(objServiceCall.GetJsonNodeValue("output.isexist"), 0)
    
    zlSplitService_HightCostStuffExistRec = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_StuffExistStock(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef intExist As Integer, Optional ByVal strKey As String) As Boolean
    '�ж������Ƿ���ڿ�����
    'strInput:����/����ID����Ʒ��/���0-�����1-��Ʒ�֣�
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_CheckExistStock
'  ---------------------------------------------------------------------------
'  --input      �ж������Ƿ���ڿ�����
'  --  stuff_id      N  1  ����id
'  --  is_item      N  1  �Ƿ�Ʒ�ֲ�ѯ��0-������ѯ��1-��Ʒ�ֲ�ѯ
'  --output
'  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message  C  1  Ӧ����Ϣ��
'  --  isexist  N 1 �Ƿ����: 1-����;0-������
'  ---------------------------------------------------------------------------
  
    If strInput = "" Then Exit Function
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("stuff_id", Split(strInput, ",")(0), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_item", Split(strInput, ",")(1), 1)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_CheckExistStock"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_StuffSvr_CheckExistStock��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '����ֵ
    intExist = Nvl(objServiceCall.GetJsonNodeValue("output.isexist"), 0)
    
    zlSplitService_StuffExistStock = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_StuffExecutePrice(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal lng����ID As Long, _
Optional ByVal strKey As String) As Boolean
    '��������ۼۣ��ɱ����Ƿ��������Ч��δִ�еļ۸����������ִ�е���
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_ExecutePrice
'  ---------------------------------------------------------------------------
'  --input      ��������ۼۣ��ɱ����Ƿ��������Ч��δִ�еļ۸����������ִ�е���
'  --  stuff_id      N    ����id
'  --output
'  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message  C  1  Ӧ����Ϣ��
'  ---------------------------------------------------------------------------
  
    If lng����ID = 0 Then Exit Function
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("stuff_id", lng����ID, 1)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_ExecutePrice"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_StuffSvr_ExecutePrice��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    
    zlSplitService_StuffExecutePrice = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_StuffGetCostPriceAdjust(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '��ȡ���ĳɱ��۵��ۼ�¼
    'strInput:Ʒ��ID����λ��0-ɢװ��λ��1-��װ��λ��
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_GetCostPriceAdjust
'  ---------------------------------------------------------------------------
'  --���ܣ���ȡ���ĳɱ��۵��ۼ�¼
'  --input
'  --  stuff_id      N   1 ����id
'  --  show_unit    N   1   ��ʾ��λ:0-ɢװ��λ;1-�ⷿ��λ
'  --output
'  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --  message  C  1  Ӧ����Ϣ��
'  --  price_list[]  ���ĳɱ��۵��ۼ�¼
'  --     stuff_id   N 1  ����ID
'  --     stuff_name   C 1  ������Ϣ
'  --     stock_name   C 1  �ⷿ
'  --     batch_number   C 1  ����
'  --     effective_time   C 1  Ч��
'  --     place_name   C 1  ����
'  --     unit_name   C 1  ��λ
'  --     cost_old   N 1  ԭ�ɱ���
'  --     cost_new    N 1  �ֳɱ���
'  --     adjust_time   C 1  ����ʱ��
'  --     adjust_reson   C 1  ����˵��
'  --     adjust_no   C 1  ���۵��ݺ�
'  --     drug_revoke_time  C 1 ����ʱ��
'  --     node_no      C    0  վ�����
'  --     is_stock    N   1 �Ƿ��п������  0-��1-��
'  ---------------------------------------------------------------------------
  
    If strInput = "" Then Exit Function
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("stuff_id", Val(Split(strInput, ",")(0)), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("show_unit", Val(Split(strInput, ",")(1)), 1)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_GetCostPriceAdjust"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_StuffSvr_GetCostPriceAdjust��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '����ֵ
    Set colOutlist = objServiceCall.GetJsonListValue("output.price_list")
    
    zlSplitService_StuffGetCostPriceAdjust = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetStockShow(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '��ȡָ���ⷿ�Ŀ�����ݣ�������ʾ
    'strInput:�ⷿid���ⷿid...
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    
    On Error GoTo ErrHandle
    
    'Zl_StuffSvr_GetStockShow
'  ---------------------------------------------------------------------------
'  --���ܣ���ȡָ���ⷿ�Ŀ�����ݣ�������ʾ
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    warehouse_ids        C   1   �ⷿID��
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
'  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    item_list
'  --      stuff_id              N   1   ����ID
'  --      warehouse_id          N   1   �ⷿID
'  --      stock                N   1   ��������
'  --      real_stock          N  1 ʵ�ʿ��
'  --      avg_price           N  1 ƽ���ۼ�
'  --      avg_cost            N  1 ƽ���ɱ���
'  ---------------------------------------------------------------------------
  
    If strInput = "" Then Exit Function
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("warehouse_ids", strInput, 0)
        
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_StuffSvr_GetStockShow"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_StuffSvr_GetStockShow��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
    
    '����ֵ
    Set colOutlist = objServiceCall.GetJsonListValue("output.item_list", strKey)
    
    zlSplitService_GetStockShow = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function





Public Function zlSplitService_GetRequestCancel(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection) As Boolean
    'ȡ���������¼
    'strInput������id,����id...
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    
'  --��ѯ�Ƿ�������������¼
'  --���      json
'  --  input      ��ѯ�Ƿ�������������¼
'  --    rcpdtl_id          C     ������ϸid,[����]��[1,2,3]
'  --    request_type       N    �������
'  --    cancel_status       N  1 ״̬
'  --����      json
'  -- output
'  --   code     C  1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --   message  C  1   Ӧ����Ϣ��
'  --   fee_cancel_list      [����]����������ÿ���������ʼ�¼
'  --     rcpdtl_id          N    ������ϸid(����id)

    '���
    StrJson_In = ""
'    StrJson_In = StrJson_In & GetJsonNodeString("rcpdtl_id", "[" & strInput & "]", 0)
'    If Not IsNull(colInput("rcpdtl_id")) Then StrJson_In = StrJson_In & "," & """rcpdtl_id"":" & "[" & colInput("rcpdtl_id") & "]"
    
    StrJson_In = StrJson_In & """rcpdtl_id"":" & "[" & strInput & "]"
    StrJson_In = StrJson_In & "," & GetJsonNodeString("request_type", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("cancel_status", 0, 1)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    strService = "Zl_Exsesvr_Getrequestcancel"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Exsesvr_Getrequestcancel��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.fee_cancel_list")
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetRequestCancel = True
End Function

Public Function zlSplitService_CallAccountDel_Check(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef strMsg As String, Optional ByRef intִ��״̬ As Integer = 1) As Boolean
    '����/סԺ��¼���˼��
    'strInput:������Ҫ����θ�ʽ��no,�ѽ��ֹ����(1),ҽ����ֹ��������(0),����״̬(1),������Դ|���,��������;���,��������...|����id,�ѷ�����;����id,�ѷ�����...
    Dim arrPart As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String, strJson_Part As String
    Dim strService As String
    Dim colJsonOutList As Collection    '�������list���ؼ���
    Dim varList As Variant   '����Ԫ��
    Dim strPart1 As String, strPart2 As String, strPart3 As String
    Dim strJson_In_Part1 As String, strJson_In_Part2 As String, strJson_In_Part3 As String
    
'    on error goto errHandle
    
    '����Zl_ExseSvr_DelBill_Check ����/סԺ��¼ͨ�÷���
    
'  ---------------------------------------------------------------------------
'  --���ܣ����ָ������ָ�����н�������
'  --��Σ�Json_In:��ʽ
'  --input
'  --        fee_no                  C   1   ���õ��ݺ�
'  --        fee_bill_type           N   1   ��������:2-������ʵ�,3-�Զ����ʵ�
'  --        balance_ban_writeoffs                   N   1   �ѽ��ֹ����:����ѽ��ʵ��ݽ�ֹ����,����ҽ�����ʵĵ��ݡ�����ԭʼ��������ֻȡδ���ʲ���
'  --        part_ban_writeoffs                  N   1   ��ֹ��������:1-������0-����
'  --        fee_origin        N 1            ������Դ��1-������ʣ�2-סԺ���ʣ�
'  --        item_list[]                         ���������б�
'  --            serial_num              N   1   ���
'  --            quantity                N   1   ��������(Ϊ��ʱ�������ֱ������)
'  --        excute_list[]                           ҩƷ����������Ӧ��ִ���б�
'  --            fee_id              N   1   ����ID
'  --            sended_num              N   1   �ѷ�����
'  --����: Json_Out,��ʽ����
'  --    output
'  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
'  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --        item_list[]                         ���������б�
'  --            serial_num              N   1   ���
'  --            quantity                N   1   ��������
'  --            execute_tag             N   1   ִ��״̬��0-δִ��;1-��ִ��;2-����ִ��
'
'  ---------------------------------------------------------------------------

    If strInput = "" Then zlSplitService_CallAccountDel_Check = True: Exit Function
    
    strJson_List = ""
    
    arrPart = Split(strInput, "|")
    
    strPart1 = arrPart(0)
    strPart2 = arrPart(1)
    strPart3 = arrPart(2)
        
    strJson_In_Part1 = ""
    strJson_In_Part1 = strJson_In_Part1 & "" & GetJsonNodeString("fee_no", Split(strPart1, ",")(0), 0)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("fee_bill_type", 2, 1)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("balance_ban_writeoffs", Split(strPart1, ",")(1), 1)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("part_ban_writeoffs", Split(strPart1, ",")(2), 1)
'    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("oper_type", Split(strPart1, ",")(3), 1)
    strJson_In_Part1 = strJson_In_Part1 & "," & GetJsonNodeString("fee_origin", Split(strPart1, ",")(4), 1)

    strJson_List = ""
    arrPart = Split(strPart2, ";")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("serial_num", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("quantity", Split(arrPart(i), ",")(1), 1)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    strJson_In_Part2 = ",""item_list"":[" & strJson_List & "]"
    
    strJson_List = ""
    arrPart = Split(strPart3, ";")
    For i = 0 To UBound(arrPart)
        strJson_Part = ""
        strJson_Part = strJson_Part & "" & GetJsonNodeString("fee_id", Split(arrPart(i), ",")(0), 1)
        strJson_Part = strJson_Part & "," & GetJsonNodeString("sended_num", Split(arrPart(i), ",")(1), 1)
        
        If strJson_List = "" Then
            strJson_List = "{" & strJson_Part & "}"
        Else
            strJson_List = strJson_List & "," & "{" & strJson_Part & "}"
        End If
    Next
    strJson_In_Part3 = ",""excute_list"":[" & strJson_List & "]"
    
    '����
    StrJson_In = "{""input"":{" & strJson_In_Part1 & strJson_In_Part2 & strJson_In_Part3 & "}}"
    
    strService = "Zl_ExseSvr_DelBill_Check"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_ExseSvr_DelBill_Check��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    intִ��״̬ = objServiceCall.GetJsonNodeValue("output.item_list[0].execute_tag")
    
    If strJson_Out = "0" Then
        strMsg = strJson_Out = objServiceCall.GetJsonNodeValue("output.message")
        zlSplitService_CallAccountDel_Check = False
        Exit Function
    End If

    zlSplitService_CallAccountDel_Check = True
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function zlSplitService_GetPati(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    'ȡ������Ϣ
    'strInput����Ŀ����;��Ŀ����
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String, strJson_List As String
    Dim strService As String
    Dim arrItem As Variant
    Dim i As Integer
        
'  ---------------------------------------------------------------------------
'Zl_Patisvr_Getpatiinfo
'  --����:��ȡ������Ϣ
'  --��Σ�Json_In:��ʽ
'  --    input
'  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
'  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;3-����
'  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
'  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
'  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
'  --      query_immune      N 1 �Ƿ����������:1-����;0-������
'  --      query_cons_list   C 1 ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��
'  --        pati_ids        C   ����IDs:����ö���
'  --        pati_name       C   ����:���Դ�%�ֺű������ƥ��
'  --        outpatient_num  C   �����
'  --        pati_idcard     C   ���֤��
'  --        contacts_idcard C   ��ϵ�����֤��
'  --        cardtype_id     N   ҽ�ƿ����ID
'  --        medc_card_name  N   ҽ�ƿ�����
'  --        card_no         C   ����
'  --        qrcode          C   ��ά��
'  --        iccard_no       C   Ic����
'  --        visit_card      C   ���￨��
'  --        insurance_num   C   ҽ����
'  --        qrspt_statu     C   ��ѯסԺ״̬:0-������;1-��Ժ ;2-���Ｐ��Ժ
'  --        phone_number    C   �ֻ���
'  --        pati_bed        C   ����
'  --����      json
'  --output
'  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message             C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    pati_list[]                 ������Ϣ�б�
'  --    pati_id             N   1   ����id
'  --    pati_pageid         N   1   ��ҳid��������Ϣ.��ҳID
'  --    pati_name           C   1   ����
'  --    pati_sex            C   1   �Ա�
'  --    pati_age            C   1   ����
'  --    pati_birthdate      C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   1   �ѱ�
'  --    outpatient_num      C   1   �����
'  --    inpatient_num       C   1   סԺ��
'  --    mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����
'  --    mdlpay_mode_code    C   1   ҽ�Ƹ��ʽ����
'  --    pati_nation         C   1   ����
'  --    insurance_num       C   1   ҽ����
'  --    pati_idcard         C   1   ���֤��
'  --    vcard_no            C   1   ���￨��
'  --    iccard_no           C   1   Ic����
'  --    health_num          C   1   ������
'  --    inp_times           N   1   סԺ����
'  --    pati_education      C   1   ѧ��
'  --    ocpt_name           C   1   ְҵ
'  --    pati_identity       C   1   ���
'  --    ntvplc_name         C   1   ����
'  --    country_name        C   1   ����
'  --    pati_marital_cstatus    C   1   ����״��
'  --    pat_home_addr           C   1   ��ͥ��ַ
'  --    pat_home_phno           C   1   ��ͥ�绰
'  --    pat_home_postcode   C   1   ��ͥ��ַ�ʱ�
'  --    pati_area           C   1   ����
'  --    pati_birthplace     C   1   �����ص�
'  --    pat_hous_addr       C   1   ���ڵ�ַ
'  --    pat_hous_postcode   C   1   ���ڵ�ַ�ʱ�
'  --    emp_name            C   1   ������λ����
'  --    emp_phno            C   1   ��λ�绰
'  --    emp_postcode        C   1   ��λ�ʱ�
'  --    emp_bank_name       C   1   ��λ������
'  --    emp_bank_accnum     C   1   ��λ�ʺ�
'  --    emp_addr             C   1   ��λ��ַ
'  --    ctt_unit_id         N   1   ��ͬ��λID
'  --    phone_number        C   1   �ֻ���
'  --    pati_bed            C   1   ��ǰ����
'  --    pati_type           C   1   ��������(��ͨ��ҽ��������)
'  --    insurance_type      C   1   ����
'  --    insurance_name      C   1   ��������
'  --    pati_wardarea_id    N   1   ��ǰ����id
'  --    pati_wardarea_name  C   1   ��ǰ��������
'  --    pati_dept_id        N   1   ��ǰ����id
'  --    pati_dept_name      C   1   ��ǰ��������
'  --    adta_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --    contacts_name       C   1   ��ϵ������
'  --    contacts_relation   C   1   ��ϵ�˹�ϵ
'  --    contacts_idcard     C   1   ��ϵ�����֤��
'  --    contacts_addr       C   1   ��ϵ�˵�ַ
'  --    contacts_phno       C   1   ��ϵ�˵绰
'  --    pat_grdn_name       C   1   �໤��
'  --    cert_no_other       C   1   ����֤��
'  --    is_inhspt            C   1   �Ƿ���Ժ:1-��Ժ ;0-����Ժ
'  --    pati_show_color      N   1   ������ʾ��ɫ
'  --    visit_room           C   1   ��������
'  --    visit_statu          N   1   ����״̬
'  --    visit_time           C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    create_time          C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    pati_email           C   1   email
'  --    pati_qq              C   1   qq
'  --    card_captcha         C   1  ����֤��
'  --    family_list[]        C   1   ������Ա:���˼���() query_family=1����
'  --        family_id        N   1   ����id  query_family=1
'  --        family_relation  C   1   ��ϵ
'  --    drug_list[]          C   1   ����ҩ���б�    query_drug=1ʱ����
'  --        pat_algc_cadn_id N   1   ����ҩƷID
'  --        pat_algc_cadn    C   1   ����ҩ������
'  --        allergy_info     C   1   ��ÿҩ�ﷴӦ
'  --    immune_list[]        C   1   ���������б�    query_immune=1ʱ����
'  --        vaccinate_time   C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --        vaccinate_name   C   1   ��������
'  --    card_list[]          C   1   ����ҽ�ƿ���Ϣ�б�(��������д����˿����ID�ģ��򷵻ظÿ����Ŀ���Ϣ)  query_card=1ʱ����
'  --        cardtype_id      N   1   ҽ�ƿ����ID
'  --        card_no          C   1   ����
'  --        card_pwd         C   1   ����
'  ---------------------------------------------------------------------------
    
    On Error GoTo ErrHandle
    
'  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
'  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;3-����
'  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
'  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
'  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
'  --      query_immune      N 1 �Ƿ����������:1-����;0-������
'  --      query_cons_list   C 1 ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��
'  --        pati_ids        C   ����IDs:����ö���
    
    '���
    StrJson_In = ""
    
    If Split(strInput, ";")(0) = "����id" Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", Split(strInput, ";")(1), 1)
    Else
        StrJson_In = StrJson_In & "" & GetJsonNodeString("pati_id", 0, 1)
    End If
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", 3, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_card", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_family", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_drug", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_immune", 0, 1)
    
    Select Case Split(strInput, ";")(0)
        Case "����ids"
            strJson_List = GetJsonNodeString("pati_ids", Split(strInput, ";")(1), 0)
        Case "ҽ����"
            strJson_List = GetJsonNodeString("insurance_num", Split(strInput, ";")(1), 0)
    End Select
    
    If strJson_List <> "" Then
        strJson_List = strJson_List & "," & GetJsonNodeString("qrspt_statu", 2, 1)
    End If
    
    strJson_List = ",""query_cons_list"":{" & strJson_List & "}"
    
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"
    strService = "Zl_Patisvr_Getpatiinfo"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Patisvr_Getpatiinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.pati_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetPati = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_CallPatiIsOut(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strPaitId As Long, ByVal strPageId As Long, ByRef intOutSign As Integer) As Boolean
    '��ҩ/��ҩ���÷��񣺲�ѯ�����Ƿ��ѳ�Ժ
    '������Ϣ������id����ҳid
    'intOutSign��0-δ��Ժ��1-��Ժ
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '�������list���ؼ���
    Dim varList As Variant   '����Ԫ��
    
    On Error GoTo ErrHandle
    
    '����zl_CisSvr_PatiIsOut
'    input           ��ѯ�����Ƿ��Ѿ���Ժ
'       pati_id N   1   ����id
'       pati_pageid  N   1   ��ҳid
'
'    output
'      code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'      message C   1   "Ӧ����Ϣ��
'      pati_outsign    N       ��Ժ��ǣ�0-δ��Ժ��1-��Ժ

    
    If strPaitId = 0 Then Exit Function

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", strPaitId, 1)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", strPageId, 1)

    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "zl_CisSvr_PatiIsOut"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�zl_CisSvr_PatiIsOut��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallPatiIsOut = False
        Exit Function
    Else
        intOutSign = Val(objServiceCall.GetJsonNodeValue("output.pati_outsign"))
    End If

    zlSplitService_CallPatiIsOut = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiByRange(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String) As Boolean
    '����Χ���Ҳ�����Ϣ
    'strInput�� ��ѯ������Ŀǰ��������ID����
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    
    'Zl_Cissvr_Getpatpageinfbyrange
'  ---------------------------------------------------------------------------
'  --����:��ȡ������Ϣ
'  --��Σ�Json_In:��ʽ
'  --  input
'  --    query_type          N 1 ��ѯ����:0-����;1-������չ
'  --    wararea_ids         C   ����ids:����ö���
'  --    dept_ids            C   ����IDs:���ö���
'  --    pati_ids            C   ����ids:����ö��ŷ���
'  --    pati_pageIds        C   ��ҳIDs:����id:��ҳid,��
'  --    adta_start_time     C   ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adta_end_time       C   ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_start_time     C   ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_end_time       C   ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   �ѱ�
'  --    inp_status          N   סԺ״̬:0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ
'  --    pati_natures        C   �������ʣ�����ö��ŷ�0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�NULL-��ʾ������
'  --    pati_name           C   ����:���Դ�%�ֺű������ƥ��
'  --    nodeno              C   վ����
'  --    change_dept_pati    N   �Ƿ��ѯת�Ʋ���
'  --����      json
'  --output
'  -- code                   N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  -- message                C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --   page_list[]          ������  ��  ��
'  --    pati_id             N    ����id  ��  ��
'  --    pati_pageid         N    ��ҳid  ��  ��
'  --    pati_name           C    ����  ��  ��
'  --    pati_sex            C    �Ա�  ��  ��
'  --    pati_age            C    ����  ��  ��
'  --    inpatient_num       C    סԺ��  ��  ��
'  --    pati_bed            C    ��Ժ����  ��  ��
'  --    insurance_type      N    ����  ��  ��
'  --    fee_category        C    �ѱ�  ��  ��
'  --    pati_type           C    ��������(��ͨ,ҽ��,����)  ��  ��
'  --    adta_time           C    ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
'  --    adtd_time           C    ��Ժʱ��:yyyy-mm-dd hh24:mi:ss  ��  ��
'  --    si_inp_status       N    סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)  ��  ��
'  --    pati_nature         N    ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
'  --    pati_wardarea_id    N    ��ǰ����id
'  --    pati_wardarea_name  C    ��ǰ��������
'  --    pati_dept_id        N    ��ǰ����id
'  --    pati_dept_name      C    ��ǰ��������
'  --    mdlpay_mode_name    C    ҽ�Ƹ��ʽ����
'  --    mdlpay_mode_code    C    ҽ�Ƹ��ʽ����
'  --    pat_rsdpscn         C    סԺҽʦ
'  --    pati_desc           C    ���˱�ע
'  --    catalog_date        C    ��Ŀ����:yyyy-mm-dd hh24:mi:ss
'  --    create_pati         C    �Ǽ���
'  --    in_objective     C    סԺĿ��
'  ---------------------------------------------------------------------------
    
    On Error GoTo ErrHandle
    
'  --    query_type          N 1 ��ѯ����:0-����;1-������չ
'  --    wararea_ids         C   ����ids:����ö���
'  --    dept_ids            C   ����IDs:���ö���
'  --    pati_ids            C   ����ids:����ö��ŷ���
'  --    pati_pageIds        C   ��ҳIDs:����id:��ҳid,��
'  --    adta_start_time     C   ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adta_end_time       C   ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_start_time     C   ��Ժ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_end_time       C   ��Ժ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   �ѱ�
'  --    inp_status          N   סԺ״̬:0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ
'  --    pati_natures        C   �������ʣ�����ö��ŷ�0-��ͨסԺ����,1-�������۲���,2-סԺ���۲��ˣ�NULL-��ʾ������
'  --    pati_name           C   ����:���Դ�%�ֺű������ƥ��
'  --    nodeno              C   վ����
'  --    change_dept_pati    N   �Ƿ��ѯת�Ʋ���
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("wararea_ids", Val(strInput), 0)

    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatpageinfbyrange"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Cissvr_Getpatpageinfbyrange��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.page_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    zlSplitService_GetPatiByRange = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetPatiId(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef strOutput As String, Optional ByVal strKey As String) As Boolean
    'ȡ������Ϣ������סԺ�ţ�����+����
    'strInput��סԺ�� �� ����id|���ţ������Ƿ��С�|��������
    'strOutPut��������Ϣ������ID
    Dim StrJson_In As String, strJson_Out As String
    Dim strService As String
        
'  ---------------------------------------------------------------------------
'Zl_Cissvr_Getpatiid
'  --����:��ȡ������Ϣ
'  --��Σ�Json_In:��ʽ
'  --input
'  --   wardarea_id          N 1 ��ǰ����id
'  --   pati_bed             C 1 ��ǰ����
'  --   inpatient_num        C 1 סԺ��
'  --output
'  --    code                N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message             C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    pati_id             N 1 ����ID:δ�ҵ�ʱҲ�ɹ�������0
'  --    pati_pageid         N   ��ҳID
'  ---------------------------------------------------------------------------
    
    On Error GoTo ErrHandle
    '���
    StrJson_In = ""
    
    If InStr(strInput, "|") > 0 Then
        StrJson_In = StrJson_In & "" & GetJsonNodeString("wardarea_id", Split(strInput, "|")(0), 1)
        StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_bed", Split(strInput, "|")(1), 0)
    Else
        StrJson_In = StrJson_In & "" & GetJsonNodeString("inpatient_num", strInput, 0)
    End If
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatiid"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Patisvr_Getpatiinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��������
    strOutput = objServiceCall.GetJsonNodeValue("output.pati_id")
    
    zlSplitService_GetPatiId = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_GetPatiName(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal colInput As Collection, _
    ByRef colPati As Collection, Optional ByVal strKey As String, Optional ByVal bytQueryType As Integer = 3) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ڰ�һ������������ѯ������Ϣ
    '���:
    '   colInput ��ѯ������ϣ�Json��input���ڵ���ΪԪ�ص�KEYֵ������ĳԪ��Ϊ�ձ�ʾ�ýڵ�ֵΪ��
    '   bytQueryType ������Ϣ��ѯ����:�磺0-����;1-����+��ϵ��;3-����
    '����:
    '����:
    '˵��:Ŀǰ֧�ֵĲ�ѯ������һ�㶼�ǰ�����һ�ֲ�ѯ������ID������ţ����������￨�ţ�ҽ���ţ�����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    
    On Error GoTo ErrHandle
    
    'Zl_Patisvr_Getpatiinfo
'  ---------------------------------------------------------------------------
'  --����:��ȡ������Ϣ
'  --��Σ�Json_In:��ʽ
'  --    input
'  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
'  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;3-����
'  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
'  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
'  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
'  --      query_immune      N 1 �Ƿ����������:1-����;0-������
'  --      query_cons_list   C 1 ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��
'  --        pati_ids        C   ����IDs:����ö���
'  --        pati_name       C   ����:���Դ�%�ֺű������ƥ��
'  --        outpatient_num  C   �����
'  --        pati_idcard     C   ���֤��
'  --        contacts_idcard C   ��ϵ�����֤��
'  --        cardtype_id     N   ҽ�ƿ����ID
'  --        medc_card_name  N   ҽ�ƿ�����
'  --        card_no         C   ����
'  --        qrcode          C   ��ά��
'  --        iccard_no       C   Ic����
'  --        visit_card      C   ���￨��
'  --        insurance_num   C   ҽ����
'  --        qrspt_statu     C   ��ѯסԺ״̬:0-������;1-��Ժ ;2-���Ｐ��Ժ
'  --        phone_number    C   �ֻ���
'  --        pati_bed        C   ����
'  --        dept_id         N   ��ǰ����ID
'  --����      json
'  --output
'  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message             C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    pati_list[]                 ������Ϣ�б�
'  --    pati_id             N   1   ����id
'  --    pati_pageid         N   1   ��ҳid��������Ϣ.��ҳID
'  --    pati_name           C   1   ����
'  --    pati_sex            C   1   �Ա�
'  --    pati_age            C   1   ����
'  --    pati_birthdate      C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss
'  --    fee_category        C   1   �ѱ�
'  --    outpatient_num      C   1   �����
'  --    inpatient_num       C   1   סԺ��
'  --    mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����
'  --    mdlpay_mode_code    C   1   ҽ�Ƹ��ʽ����
'  --    pati_nation         C   1   ����
'  --    insurance_num       C   1   ҽ����
'  --    pati_idcard         C   1   ���֤��
'  --    vcard_no            C   1   ���￨��
'  --    iccard_no           C   1   Ic����
'  --    health_num          C   1   ������
'  --    inp_times           N   1   סԺ����
'  --    pati_education      C   1   ѧ��
'  --    ocpt_name           C   1   ְҵ
'  --    pati_identity       C   1   ���
'  --    ntvplc_name         C   1   ����
'  --    country_name        C   1   ����
'  --    pati_marital_cstatus    C   1   ����״��
'  --    pat_home_addr           C   1   ��ͥ��ַ
'  --    pat_home_phno           C   1   ��ͥ�绰
'  --    pat_home_postcode   C   1   ��ͥ��ַ�ʱ�
'  --    pati_area           C   1   ����
'  --    pati_birthplace     C   1   �����ص�
'  --    pat_hous_addr       C   1   ���ڵ�ַ
'  --    pat_hous_postcode   C   1   ���ڵ�ַ�ʱ�
'  --    emp_name            C   1   ������λ����
'  --    emp_phno            C   1   ��λ�绰
'  --    emp_postcode        C   1   ��λ�ʱ�
'  --    emp_bank_name       C   1   ��λ������
'  --    emp_bank_accnum     C   1   ��λ�ʺ�
'  --    emp_addr             C   1   ��λ��ַ
'  --    ctt_unit_id         N   1   ��ͬ��λID
'  --    phone_number        C   1   �ֻ���
'  --    pati_bed            C   1   ��ǰ����
'  --    pati_type           C   1   ��������(��ͨ��ҽ��������)
'  --    insurance_type      C   1   ����
'  --    insurance_name      C   1   ��������
'  --    pati_wardarea_id    N   1   ��ǰ����id
'  --    pati_wardarea_name  C   1   ��ǰ��������
'  --    pati_dept_id        N   1   ��ǰ����id
'  --    pati_dept_name      C   1   ��ǰ��������
'  --    adta_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --    adtd_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --    contacts_name       C   1   ��ϵ������
'  --    contacts_relation   C   1   ��ϵ�˹�ϵ
'  --    contacts_idcard     C   1   ��ϵ�����֤��
'  --    contacts_addr       C   1   ��ϵ�˵�ַ
'  --    contacts_phno       C   1   ��ϵ�˵绰
'  --    pat_grdn_name       C   1   �໤��
'  --    cert_no_other       C   1   ����֤��
'  --    is_inhspt            C   1   �Ƿ���Ժ:1-��Ժ ;0-����Ժ
'  --    pati_show_color      N   1   ������ʾ��ɫ
'  --    visit_room           C   1   ��������
'  --    visit_statu          N   1   ����״̬
'  --    visit_time           C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    create_time          C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
'  --    pati_email           C   1   email
'  --    pati_qq              C   1   qq
'  --    card_captcha         C   1  ����֤��
'  --    family_list[]        C   1   ������Ա:���˼���() query_family=1����
'  --        family_id        N   1   ����id  query_family=1
'  --        family_relation  C   1   ��ϵ
'  --    drug_list[]          C   1   ����ҩ���б�    query_drug=1ʱ����
'  --        pat_algc_cadn_id N   1   ����ҩƷID
'  --        pat_algc_cadn    C   1   ����ҩ������
'  --        allergy_info     C   1   ��ÿҩ�ﷴӦ
'  --    immune_list[]        C   1   ���������б�    query_immune=1ʱ����
'  --        vaccinate_time   C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
'  --        vaccinate_name   C   1   ��������
'  --    card_list[]          C   1   ����ҽ�ƿ���Ϣ�б�(��������д����˿����ID�ģ��򷵻ظÿ����Ŀ���Ϣ)  query_card=1ʱ����
'  --        cardtype_id      N   1   ҽ�ƿ����ID
'  --        card_no          C   1   ����
'  --        card_pwd         C   1   ����
'  ---------------------------------------------------------------------------
'  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
'  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;3-����
'  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
'  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
'  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
'  --      query_immune      N 1 �Ƿ����������:1-����;0-������
'  --      query_cons_list   C 1 ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��

    '���
    StrJson_In = ""
    If Not IsNull(colInput("pati_id")) Then
        StrJson_In = GetJsonNodeString("pati_id", colInput("pati_id"), 1)
    Else
        StrJson_In = GetJsonNodeString("pati_id", 0, 1)
    End If
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_type", bytQueryType, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_card", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_family", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_drug", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("query_immune", 0, 1)
    
    If IsNull(colInput("pati_id")) Then
        '���ܰ�����һ�ַ�ʽ��ѯ������ţ����������￨�ţ�ҽ���ţ�����
        strJson_List = ""
        If Not IsNull(colInput("outpatient_num")) Then strJson_List = strJson_List & "," & GetJsonNodeString("outpatient_num", colInput("outpatient_num"), 0)
        If Not IsNull(colInput("pati_name")) Then strJson_List = strJson_List & "," & GetJsonNodeString("pati_name", colInput("pati_name"), 0)
        If Not IsNull(colInput("pati_vcard_no")) Then strJson_List = strJson_List & "," & GetJsonNodeString("visit_card", colInput("pati_vcard_no"), 0)
        'If Not IsNull(colInput("insurance_num")) Then StrJson_In = StrJson_In & "," & GetJsonNodeString("insurance_num", colInput("insurance_num"), 0)
        If Not IsNull(colInput("pati_bed")) Then strJson_List = strJson_List & "," & GetJsonNodeString("pati_bed", colInput("pati_bed"), 0)
        If strJson_List <> "" Then
            strJson_List = GetJsonNodeString("qrspt_statu", 2, 1) & strJson_List '2-���Ｐ��Ժ
            strJson_List = ",""query_cons_list"":{" & strJson_List & "}"
        ElseIf ExistsColObject(colInput, "pati_deptid") Then '�����Ҳ�ѯ
            strJson_List = GetJsonNodeString("qrspt_statu", 1, 1)   '1-��Ժ
            strJson_List = strJson_List & "," & GetJsonNodeString("dept_id", colInput("pati_deptid"), 1)
            strJson_List = ",""query_cons_list"":{" & strJson_List & "}"
        End If
    End If
    StrJson_In = "{""input"":{" & StrJson_In & strJson_List & "}}"
    
    strService = "Zl_Patisvr_Getpatiinfo"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Patisvr_Getpatiinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colPati = objServiceCall.GetJsonListValue("output.pati_list", strKey)
    
    If colPati Is Nothing Then Exit Function
    
    zlSplitService_GetPatiName = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function zlSplitService_GetPatiPage(ByVal objServiceCall As Object, ByVal lngMode As Long, ByVal strInput As String, _
    ByRef colOutlist As Collection, Optional ByVal strKey As String, Optional ByRef colOutListBaby As Collection) As Boolean
    'ȡ������ҳ��Ϣ
    'strInput������id:��ҳid,...
    Dim StrJson_In As String, strJson_Out As String, strJson_OutList As String, strJson_OutListKey As String
    Dim strService As String
    Dim varList As Variant  '����Ԫ��
    Dim colTmp As New Collection, colbaby As New Collection
    Dim i As Integer, n As Integer
    
'---------------------------------------------------------------------------
'Zl_Cissvr_Getpatipageinfo
'  --����:��ȡ������ҳ�����Ϣ
'  --��Σ�Json_In:��ʽ
'  --    input
'  --      query_type          C 1 ��ѯ����:0-������Ϣ;1-������Ϣ��չ;2-��ȡ��ҳ
'  --      pati_pageids        C 1 ������Ϣ,��ʽ����:һ����:����id:��ҳID,��;һ�֣�����id,��
'  --      is_babyinfo         N 1 �Ƿ����Ӥ����Ϣ:1-����;0-������
'  --      is_transdeptinfo    N 1 �Ƿ����ת����Ϣ:1-����;0-������
'  --      is_lastpage         N 1 �Ƿ�ȡ���һ��סԺ
'  --����: Json_Out,��ʽ����
'  --  output
'  --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    pati_count            N 1 ��ѯ�Ĳ�����Ϣ����
'  --    page_list[]             1 ������
'  --      pati_id             N 1 ����id
'  --      pati_pageid         N 1 ��ҳid
'  --      pati_name           C 1 ����
'  --      pati_sex            C 1 �Ա�
'  --      pati_age            C 1 ����
'  --      inpatient_num       C 1 סԺ��
'  --      fee_category        C 1 �ѱ�
'  --      mdlpay_mode_name    C 1 ҽ�Ƹ��ʽ����
'  --      mdlpay_mode_code    C 1 ҽ�Ƹ��ʽ����
'  --      pati_bed            C 1 ��ǰ����
'  --      pati_type           C 1 ��������(��ͨ��ҽ��������)
'  --      pati_show_color     N    ����������ɫ
'  --      pati_education      C 1 ѧ��
'  --      ocpt_name           C 1 ְҵ
'  --      country_name        C 1 ����
'  --      pati_marital_cstatus  C 1 ����״��
'  --      pati_nature         N 1 ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
'  --      audit_sign          N 1 ��˱�־:������ҳ.��˱�־
'  --      si_inp_status       N 1 סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)
'  --      pati_wardarea_id    N 1 ��ǰ����id
'  --      pati_wardarea_name  C 1 ��ǰ��������
'  --      pati_dept_id        N 1 ��ǰ����id
'  --      pati_dept_name      C 1 ��ǰ��������
'  --      adta_time           C 1 ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --      adtd_time           C 1 ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
'  --      insurance_type      N 1 ����
'  --      rgst_id             N 1 �Һ�id
'  --      catalog_date        C 1 ��Ŀ����:yyyy-mm-dd hh24:mi:ss
'  --      in_objective        C 1 סԺĿ��
'  --      reg_name            C 1 �Ǽ���
'  --      reg_date            C 1 סԺ�Ǽ�ʱ��
'  --      pat_rsdpscn         C 1 סԺҽʦ
'  --      pati_desc           C 1 ���˱�ע
'  --      baby_list[]           1 Ӥ����Ϣ��[����]
'  --        pati_id           N 1 ����id
'  --        pati_pageid       N 1 ��ҳid
'  --        baby_num          N 1 Ӥ�����
'  --        baby_name         C 1 Ӥ������
'  --        baby_sex          C 1 Ӥ���Ա�
'  --        baby_date         C 1 ����ʱ��
'  --      trans_list[]        C   ת���б���Ϣ
'  --        start_reason      C 1 ��ʼԭ��
'  --        start_time        C 1 ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
'  --        dept_name         C 1 ��������
'  ---------------------------------------------------------------------------
    
    On Error GoTo ErrHandle
    
'  --    input
'  --      query_type          C 1 ��ѯ����:0-������Ϣ;1-������Ϣ��չ;2-��ȡ��ҳ
'  --      pati_pageids        C 1 ������Ϣ,��ʽ����:һ����:����id:��ҳID,��;һ�֣�����id,��
'  --      is_babyinfo         N 1 �Ƿ����Ӥ����Ϣ:1-����;0-������
'  --      is_transdeptinfo    N 1 �Ƿ����ת����Ϣ:1-����;0-������
'  --      is_lastpage         N 1 �Ƿ�ȡ���һ��סԺ
    
    '���
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("pati_pageids", strInput, 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_babyinfo", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_transdeptinfo", 0, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("is_lastpage", 0, 1)
    
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    strService = "Zl_Cissvr_Getpatipageinfo"
    
'    StrJson_In = "{""input"":{""pati_list"":""100692,null|53942,null""}}"
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode) = False Then
        MsgBox "���á�Zl_Cissvr_Getpatipageinfo��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    '��������
    Set colOutlist = objServiceCall.GetJsonListValue("output.page_list", strKey)
    
    If colOutlist Is Nothing Then Exit Function
    
    Set colOutListBaby = New Collection
    
    'Ӥ�����ݵĲ���ID,��ҳid��ͬ��Ӥ����Ų�ͬ��Ҫô����key��Ҫô�� ����id+��ҳid+Ӥ����� ��Ϊkey
    For n = 1 To colOutlist.Count
        Set colbaby = objServiceCall.GetJsonListValue("output.page_list[" & n - 1 & "].baby_list")
        If colbaby.Count > 0 Then
            For i = 1 To colbaby.Count
                Set colTmp = Nothing
                
                colTmp.Add colbaby(i)("_pati_id"), "_pati_id"
                colTmp.Add colbaby(i)("_pati_pageid"), "_pati_pageid"
                colTmp.Add colbaby(i)("_baby_num"), "_baby_num"
                colTmp.Add colbaby(i)("_baby_name"), "_baby_name"
                colTmp.Add colbaby(i)("_baby_sex"), "_baby_sex"
                colTmp.Add colbaby(i)("_baby_date"), "_baby_date"
                
                colOutListBaby.Add colTmp, "_" & colbaby(i)("_pati_id") & "_" & colbaby(i)("_pati_pageid") & "_" & colbaby(i)("_baby_num")
            Next
        End If
    Next
    
    zlSplitService_GetPatiPage = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlSplitService_CallIsCloseAcc(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal strInput As String, ByRef inState As Integer) As Boolean
    '��ҩ/��ҩ���÷��񣺲�ѯ�Ƿ��ѽ���
    'strInput��������Դ|No
    'inState�����ؽ���״̬
    Dim arrInput As Variant
    Dim i As Integer
    Dim strJson As String, StrJson_In As String, strJson_Out As String, strJson_List As String
    Dim strService As String
    Dim colJsonOutList As Collection    '�������list���ؼ���
    Dim varList As Variant   '����Ԫ��
    
    On Error GoTo ErrHandle
    'Zl_Exsesvr_Getfeebalancestate
'  ---------------------------------------------------------------------------
'  --����:���ݵ��ݺ���Ϣ����ȡ���ݶ�Ӧ�Ľ���״̬
'  --��Σ�Json_In:��ʽ
'  --input
'  --    query_mode  N 1 ��ѯ��ʽ:0-�������;1-סԺ����
'  --    bill_nos  C 1 ���ݺ�
'  --����: Json_Out,��ʽ����
'  -- output
'  --    code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
'  --    message C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
'  --    state N 1 ״̬:-1-�����ڼ��ʵ���;0-δ����;1-���ֽ���;2-ȫ������
'
'  ---------------------------------------------------------------------------
    
    If strInput = "" Then Exit Function

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_mode", Split(strInput, "|")(0), 1)
    strJson = strJson & "," & GetJsonNodeString("bill_nos", Split(strInput, "|")(1), 0)

    StrJson_In = "{""input"":{" & strJson & "}}"
    
    strService = "Zl_Exsesvr_Getfeebalancestate"
    
    '���÷���
    If objServiceCall.CallService(strService, StrJson_In, strJson_Out, "", lngMode, False) = False Then
        MsgBox "���á�Zl_Exsesvr_Getfeebalancestate��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���س���
    strJson_Out = objServiceCall.GetJsonNodeValue("output.code")
    
    If strJson_Out = "0" Then
        zlSplitService_CallIsCloseAcc = False
        Exit Function
    Else
        inState = Val(objServiceCall.GetJsonNodeValue("output.state"))
    End If

    zlSplitService_CallIsCloseAcc = True
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSplitService_GetExseBillByTime(ByVal objServiceCall As Object, ByVal lngMode As Long, _
    ByVal colQueryCons As Collection, ByRef strBill_Out As String, ByRef strErrMsg_Out As String) As Boolean
    '��ʱ�䷶Χ��ȡ���õ���
    '��Σ�
    '   colQueryCons = ��ѯ��������Ա(Key)��������Դ,��ʼʱ��,����ʱ��,ִ�в���IDS,����ִ�в���IDS
    '                           ���У�������Դ��0-������;1-����;2-סԺ
    '����:
    '   strBill_Out = ������Ϣ����ʽ������1:NO,����1:NO,...�����У����ݣ�24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ�������
    Dim StrJson_In As String
    Dim colOutlist As Collection, colTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    strBill_Out = "": strErrMsg_Out = ""
    'Zl_Exsesvr_Getbillbytime
    '  --���ܣ���ʱ�䷶Χ��ȡ���õ���
    '  --��Σ�json��ʽ
    '  --  input
    '  --    query_type          N 0 ��ѯ��ʽ:0-��ȡҩƷҽ�����õ��ݣ�1-��ȡ����ҽ�����õ���
    '  --    fee_source          N 1 ������Դ:0-������;1-����;2-סԺ
    '  --    start_time          C 1 ��ʼʱ�䣬��ʽ��yyyy-mm-dd hh24:mi:ss
    '  --    end_time            C 1 ����ʱ�䣬��ʽ��yyyy-mm-dd hh24:mi:ss
    '  --    exe_deptids         C 0 ִ�в���ID������ö�Ӣ�ĺŷָ�
    '  --    excp_exe_deptids    C 0 ��������ִ�в���ID������ö�Ӣ�ĺŷָ�
    '  --���Σ�json��ʽ
    '  --  output
    '  --    code                C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --    message             C 1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    bill_nos            C 1 ������Ϣ:��ʽ����������1:NO1,��������2:NO2,...
    '  --                            ���У���������: 1-�շѴ���;2-���ʵ�����;3-���ʱ���
     
    StrJson_In = ""
    StrJson_In = StrJson_In & "" & GetJsonNodeString("query_type", 1, 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("fee_source", colQueryCons("������Դ"), 1)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("start_time", Format(colQueryCons("��ʼʱ��"), "yyyy-MM-dd HH:mm:ss"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("end_time", Format(colQueryCons("����ʱ��"), "yyyy-MM-dd HH:mm:ss"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("exe_deptids", colQueryCons("ִ�в���IDS"), 0)
    StrJson_In = StrJson_In & "," & GetJsonNodeString("excp_exe_deptids", colQueryCons("����ִ�в���IDS"), 0)
    StrJson_In = "{""input"":{" & StrJson_In & "}}"
    
    If objServiceCall.CallService("Zl_Exsesvr_Getbillbytime", StrJson_In, "", "", lngMode, False, , , , True) = False Then Exit Function
    
    strBill_Out = objServiceCall.GetJsonNodeValue("output.bill_nos")
    
    If strBill_Out <> "" Then
        '��������ת����24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ�������
        strBill_Out = "," & strBill_Out
        strBill_Out = Replace(strBill_Out, ",1:", ",24:")
        strBill_Out = Replace(strBill_Out, ",2:", ",25:")
        strBill_Out = Replace(strBill_Out, ",3:", ",26:")
        strBill_Out = zlStr.TrimEx(strBill_Out, ",")
    End If
    
    zlSplitService_GetExseBillByTime = True
    Exit Function
ErrHandler:
    strErrMsg_Out = err.Description
End Function

