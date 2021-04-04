Attribute VB_Name = "mdlPubServiceExse"

Option Explicit

'*********************************************************************************************************************************************
'����:�����漰���÷��õ���ط���
'�ӿ�˵��:
'    0.zl_ExseSvr_GetNextNo-��ȡ��һ�ŵ��ݺ�
'    1.Zl_Exsesvr_Updatecardinvoice-���Ѱ�����ҽ��Ʊ��ʹ��ʱ��������ҽ��Ʊ�ݵ���ظ��²�����
'    2.zl_ExseSvr_GetReceiveInvoice-��ȡƱ��������Ϣ
'    3.zl_ExseSvr_GetNextInvoice-���ݵ�ǰ��Ʊ�ż�Ʊ��ʹ����ϸ����ȡһ����Ч�ķ�Ʊ��
'    4.zl_ExseSvr_GetFullNo-�Զ����뵥�ݺ�(����)
'    5.Zl_Exsesvr_GetBillOperControls-��ȡ���ݲ�����������
'    6.zl_ExseSvr_GetBillTotalMoney-��ȡ���ݵ�Ӧ�ջ�ʵ���ܶ�
'    7.zl_ExseSvr_GetBillInfoByNo-��ȡ���ݵ�Ӧ�ջ�ʵ���ܶ�
'    8.zl_ExseSvr_GetPatiInvoiceClass-��ȡƱ��ʹ�����
'    9.zl_ExseSvr_InvoiceClassUsed-���ָ��Ʊ���Ƿ�������ʹ�����
'    10.zl_ExseSvr_Actualmoney-���ݷѱ������۽��
'    11.Zl_Exsesvr_GetDepositDetail-��ѯָ�����˵�Ԥ����֧��ϸ���
'    12.zl_ExseSvr_BillInHistory-���ָ�����õ����Ƿ����󱸱�ռ�
'    13.zl_ExseSvr_GetCardFeeInfoByNo- ����ָ���Ŀ��ѵ��ݣ���ȡ���ü����㼰Ԥ����Ϣ
'    14.Zl_Exsesvr_CheckCardnoIsUsed-���ָ�������Ƿ���Ʊ��ʹ����ϸ�д��ڣ�����ʱ����������ID
'    15.Zl_Exsesvr_Addcardfeeinfo-���ӿ��Ѽ�Ԥ������
'    16.Zl_Exsesvr_UpdCardFeeBlncInfo:��ɿ����շѽ���
'    17.zl_ExseSvr_GetBillStatuByNo-��ȡָ���շѵ��ݵ��쳣���շѼ����ʵ�״̬.
'    18.zl_ExseSvr_DelCardFeeCheck-�˿����˲����ѺϷ��Լ��
'    19.Zl_Exsesvr_Delcardfeeinfo:�˿��Ѽ�Ԥ����������
'    20.zl_ExseSvr_GetRelatedTransInfo:���ݹ�������ID,��ȡ������Ϣ
'    21.Zl_Exsesvr_Getbalanceinfo-���ݵ��ݵ�������ȡ������Ϣ
' �����Һ���ط���
'    1.Zl_Exsesvr_Getapptregisterinfo-���ݹҺŵ���ȡԤԼ�Һ���Ϣ

'����:���˺�
'����:2019*10*31 14:47:18
'*********************************************************************************************************************************************
Public gobjServiceCall As Object
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


Public Function Zl_Exsesvr_Updatecardinvoice(ByVal int�������� As Integer, ByVal str���õ��� As String, _
    ByVal lng����ID As Long, ByVal str��Ʊ�� As String, ByVal strʹ���� As String, Optional ByVal strʹ��ʱ�� As String, _
    Optional ByVal intʹ������ As Integer = 1, Optional ByRef strUseFact As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ѱ�����ҽ��Ʊ��ʹ��ʱ��������ҽ��Ʊ�ݵ���ظ��²���
    '���:int��������-1-������2-�˿���3-�ش�4-����5-����
    '     str���õ���-���õ��ݺ�
    '     strʹ��ʱ��-��ʽ:yyyy-mm-dd hh24:mi:ss
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:strUseFact-���ر����ʹ�õķ�Ʊ��,����ö���
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
    
    '    input           ȡ���˷������
    '        fun_oper    N   1   ��������:1-������2-�˿���3-�ش�4-����5-����
    '        fee_no  C   1   ���õ���
    '        recv_id N   1   ����id
    '        inv_no  C   1   ��ǰ��Ʊ�Ż�ʼʹ�÷�Ʊ��
    '        inv_usenums N   1   ��Ʊʹ������
    '        use_time    C   1   Ʊ��ʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
    '        inv_user    C   1   ��Ʊʹ����
    If strʹ��ʱ�� = "" Then strʹ��ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fun_oper", int��������, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_no", str���õ���, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("recv_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_no", str��Ʊ��, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("inv_usenums", intʹ������, Json_num)
    strJson = strJson & "," & GetJsonNodeString("use_time", strʹ��ʱ��, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("inv_user", strʹ����, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Updatecardinvoice"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    'output
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   "Ӧ����Ϣ��
    '    �ɹ�ʱ���سɹ���Ϣ
    '    ʧ��ʱ���ؾ���Ĵ�����Ϣ ""
    '    inv_outnos  C   1   ����ҽ��Ʊ��:ʹ�õ�����ҽ��Ʊ��,����ö��ŷ���
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܸ��·�Ʊʹ����Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strUseFact = objServiceCall.GetJsonNodeValue("output.inv_outnos")
    Zl_Exsesvr_Updatecardinvoice = True
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

 


Public Function zl_ExseSvr_GetReceiveInvoice(ByVal intƱ�� As Integer, ByVal str����ids As String, ByRef cllBillInfo_Out As Collection, _
    Optional ByVal bln�Ƿ����� As Boolean = True, Optional ByVal strʹ����� As String, Optional ByVal str������ As String, _
    Optional ByVal int���ٷ�Ʊ�� As Integer = 0, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long, Optional ByVal bytOperFun As Byte = 0, Optional ByVal blnResolveToRecord As Boolean, _
    Optional ByRef rsInvoice_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ѱ�����ҽ��Ʊ��ʹ��ʱ��������ҽ��Ʊ�ݵ���ظ��²���
    '���:intƱ��-1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '     bln�Ƿ�����-true-���ã���Ʊ�ݽ����������Լ�ʹ�ã�false-���ã���Ʊ���ɶ����Ա��ͬʹ��
    '     str����ids-����ö���(����:1,2...)
    '     strʹ�����-Ʊ��ʹ�����:1,4: Ʊ��ʹ�����.����;2Ԥ��:1-����Ԥ��;2-סԺԤ��;5:�洢����ҽ�ƿ����.ID
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '     bytOperFun-0-��ȡƱ��������Ϣ 1-��ȡ��ȡָ��Ʊ�ֵĹ���Ʊ������
    '     blnResolveToRecord-�Ƿ�ת��Ϊ��¼��
    '����:cllBillInfo_Out-����Ʊ��������Ϣ��
    '     rsInvoice_Out-���ط�Ʊ��Ϣ��(blnResolveToRecord=trueʱ)
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
   
    'zl_ExseSvr_GetReceiveInvoice
    '   input
    '        oper_fun    N   1   0-��ȡƱ��������Ϣ 1-��ȡ��ȡָ��Ʊ�ֵĹ���Ʊ������
    '        recv_ids N   1   ����ids:Ʊ������id
    '        inv_type    N   1   Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '        use_mode    N   1   ʹ�÷�ʽ:1-���ã���Ʊ�ݽ����������Լ�ʹ�ã�2-���ã���Ʊ���ɶ����Ա��ͬʹ��
    '        use_type    C   1   Ʊ��ʹ�����:1,4: Ʊ��ʹ�����.����;2Ԥ��:1-����Ԥ��;2-סԺԤ��;5:�洢����ҽ�ƿ����.ID
    '        recvtr  C   1   ������
    '        min_nums  N 1 ��Ʊ��������
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_fun", bytOperFun, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_type", intƱ��, Json_num)
    strJson = strJson & "," & GetJsonNodeString("recv_ids", str����ids, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("use_mode", IIf(bln�Ƿ�����, 1, 2), Json_num)
    strJson = strJson & "," & GetJsonNodeString("use_type", strʹ�����, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("min_nums", int���ٷ�Ʊ��, Json_num)
    strJson = strJson & "," & GetJsonNodeString("recvtr", str������, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetReceiveInvoice"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '        code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '        message C   1   "Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣʧ��ʱ���ؾ���Ĵ�����Ϣ ""
    '        item_list C
    '            recv_id N   1   ����ID
    '            use_mode    N   1   ʹ�÷�ʽ:1-���ã���Ʊ�ݽ����������Լ�ʹ�ã�2-���ã���Ʊ���ɶ����Ա��ͬʹ��
    '            use_type    C   1   Ʊ��ʹ�����:1,4: Ʊ��ʹ�����.����;2Ԥ��:1-����Ԥ��;2-סԺԤ��;5:�洢����ҽ�ƿ����.ID
    '            prefix_text C   1   ǰ׺�ı�
    '            start_no    C   1   ��ʼ����
    '            end_no  C   1   ��ֹ����
    '            inv_no_cur  C   1   ��ǰ����
    '            surplus_num C   1   ʣ������
    '            create_time C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
    '            use_time    C   1   ʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
    '            recvtr  C   1   ������

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "δ�ҵ�����������Ʊ��������Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set cllBillInfo_Out = objServiceCall.GetJsonListValue("output.item_list", "recv_id")
    If blnResolveToRecord Then Call zlGetReceiveInvoiceRecFromCollect(cllBillInfo_Out, rsInvoice_Out)
    zl_ExseSvr_GetReceiveInvoice = True
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
 

Public Function zl_ExseSvr_GetNextInvoice(ByVal lng����ID As Integer, ByVal str��Ʊ�� As String, ByRef str��һ�ŷ�Ʊ_Out As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ѱ�����ҽ��Ʊ��ʹ��ʱ��������ҽ��Ʊ�ݵ���ظ��²���
    '���:str��Ʊ��-��һ�ŷ�Ʊ��
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:str��һ�ŷ�Ʊ_Out-������Ч��Ʊ�ݺ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    str��һ�ŷ�Ʊ_Out = ""
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetNextInvoice
    '    input
    '       recv_id N   1   ����id:Ʊ������id
    '       inv_no  C   1   ��Ʊ��

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("recv_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_no", str��Ʊ��, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetNextInvoice"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '       code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '       message C   1   "Ӧ����Ϣ��
    '       inv_no  C   1   ��һ����Ʊ��
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "δ�ҵ�����������Ʊ����Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    str��һ�ŷ�Ʊ_Out = objServiceCall.GetJsonNodeValue("output.inv_no")
    zl_ExseSvr_GetNextInvoice = True
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


Public Function zl_ExseSvr_GetFullNo(ByVal int��� As Integer, ByVal strInputNo As String, ByVal lng����ID As Long, ByRef strFullNo_Out As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Զ����뵥�ݺ�(����)
    '���:int���-������Ʊ��е����
    '     strInputNO-����ĵ��ݺ�
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:strFullNo_Out-���������ĵ��ݺ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    strFullNo_Out = strInputNo
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetFullNo
    '    input
    '        item_num    N   1   ��Ŀ���
    '        input_no    C   1   ����ĵ��ݺ�
    '        dept_id N       ����ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("item_num", int���, Json_num)
    strJson = strJson & "," & GetJsonNodeString("input_no", strInputNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("dept_id", lng����ID, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetFullNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '        code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '        message C   1   "Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '        full_no C       �����ĵ��ݺ�
    '
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "�����Զ����뵥����Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strFullNo_Out = objServiceCall.GetJsonNodeValue("output.full_no")
    zl_ExseSvr_GetFullNo = True
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

Public Function Zl_Exsesvr_GetBillOperControls(ByVal int���� As Integer, ByVal lng��ԱID As Long, ByRef bln���ڿ���_Out As Boolean, ByRef intʱ������_Out As Integer, _
    ByRef int���˵���_Out As Integer, ByRef dbl�������_Out As Double, Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ݲ�����������
    '���:int����-����:1-�Һŵ���,2-�շѵ�,3-���۵�,4-�������,5-סԺ����,6-Ԥ����,7-���ʵ���,8-���￨
    '     lng��ԱID-��ԱID
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:intʱ������_out-����ʱ������(0(NULL)-������,n-n����)
    '     bln���ڿ���_out-�Ƿ���ڵ��ݿ���:true-���ڣ�false-������
    '     dbl�������_Out
    '     int���˵���_out
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
    
    'zl_ExseSvr_GetBillOperControls
    '    input           ����ת�������õ�ת�룬ת������
    '        bill_type   N   1   ��������:1-�Һŵ���,2-�շѵ�,3-���۵�,4-�������,5-סԺ����,6-Ԥ����,7-���ʵ���,8-���￨
    '        operator_id N   1   ��ԱID

    strJson = strJson & "" & GetJsonNodeString("bill_type", int����, Json_num)
    strJson = strJson & "," & GetJsonNodeString("operator_id", lng��ԱID, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Exsesvr_GetBillOperControls"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    'output
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    is_exist    N   1   ���ڿ�������:1-����;0-������
    '    time_limit  N   1   0(NULL)-������,n-n����
    '    other_bill  N   1   �Ƿ�������������ݽ��в���
    '    uplimit_money   N   1   �������

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡ��Ч�ĵ��ݿ�����Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    bln���ڿ���_Out = Val(objServiceCall.GetJsonNodeValue("output.is_exist")) = 1
    intʱ������_Out = Val(objServiceCall.GetJsonNodeValue("output.time_limit"))
    int���˵���_Out = Val(objServiceCall.GetJsonNodeValue("output.other_bill"))
    dbl�������_Out = Val(objServiceCall.GetJsonNodeValue("output.uplimit_money"))
    Zl_Exsesvr_GetBillOperControls = True
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


Public Function zl_ExseSvr_GetBillTotalMoney(ByVal bln���� As Boolean, ByVal int��¼���� As Integer, ByVal str���ݺ� As String, _
    ByRef dblӦ�ս��_Out As Double, ByRef dblʵ�ս��_Out As Double, Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional lngModule As Long, Optional ByVal lng����ID As Long, Optional ByVal str��¼״̬s As String = "0,1") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ݵ�Ӧ�ջ�ʵ���ܶ�
    '���:bln����-�Ƿ�����
    '     int��¼����-:1-�շѵ�;3-�Զ����ʵ���2 -���ʼ�¼��4-�Һż�¼ ;5-���￨
    '     lng����id-�ಡ�˵�ʱ������ID��Ч
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:dblӦ�ս��_Out-Ӧ�պϼ�
    '     dblʵ�ս��_Out-ʵ�պϼ�
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
    
    'zl_ExseSvr_GetBillTotalMoney
    '    input
    '        fee_origin  N   1   ������Դ:1-����;2-סԺ
    '        bill_type   N   1   ��������:1-�շѵ�;3-�Զ����ʵ���2 -���ʼ�¼��4-�Һż�¼ ;5-���￨
    '        fee_no  C   1   ���ݺ�:���õ��ݺ�
    '        pati_id N   1   ����id
    '        rec_status  C       ��¼״̬s:���Զ��״̬,����:0,1


    strJson = strJson & "" & GetJsonNodeString("fee_origin", IIf(bln����, 1, 2), Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_type", int��¼����, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_no", str���ݺ�, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rec_status", str��¼״̬s, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetBillTotalMoney"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '        code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '        message C   1   "Ӧ����Ϣ�� �ɹ�ʱ���سɹ���Ϣ ,ʧ��ʱ���ؾ���Ĵ�����Ϣ ""
    '        fee_amrcvb  N   1   Ӧ�ս��
    '        fee_ampaib  N   1   ʵ�ս��
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡ���õ���Ϊ" & str���ݺ� & "��Ӧ�ջ�ʵ���ܶ���飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
     
    dblӦ�ս��_Out = Val(objServiceCall.GetJsonNodeValue("output.fee_amrcvb"))
    dblʵ�ս��_Out = Val(objServiceCall.GetJsonNodeValue("output.fee_ampaib"))
    zl_ExseSvr_GetBillTotalMoney = True
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






Public Function zl_ExseSvr_GetBillInfoByNo(ByVal bln���� As Boolean, ByVal int�������� As Integer, ByVal str���ݺ� As String, _
    ByRef str����Ա����_Out As String, ByRef str�Ǽ�ʱ��_Out As String, ByRef lng����id_out As Long, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ݵ�Ӧ�ջ�ʵ���ܶ�
    '���:bln����-�Ƿ�����
    '     int��������-:1-�շѵ�;3-�Զ����ʵ���2 -���ʼ�¼��4-�Һż�¼ ;5-���￨;-1-���ʵ�;-2-Ԥ����;-3-�������
    '     lng����id-�ಡ�˵�ʱ������ID��Ч
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:str����Ա����_Out
    '     str�Ǽ�ʱ��_Out
    '     lng����id_Out
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
    
    'zl_ExseSvr_GetBillInfoByNo
    ' input
    '    fee_origin  N   1   ������Դ:1-����;2-סԺ ;bill_type>0ʱ��Ч
    '    bill_type   N   1   ��������:1-�շѵ�;3-�Զ����ʵ���2 -���ʼ�¼��4-�Һż�¼ ;5-���￨;-1-���ʵ�;-2-Ԥ����;-3-�������
    '    bill_no C   1   ���ݺ�
    
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_origin", IIf(bln����, 1, 2), Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_type", int��������, Json_num)
    strJson = strJson & "," & GetJsonNodeString("bill_no", str���ݺ�, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetBillInfoByNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   "Ӧ����Ϣ��
    '    operator_name   C   1   ����Ա����
    '    create_time C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
    '    pati_id N   1   ����id

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡ���õ���Ϊ" & str���ݺ� & "�ĵǼ�ʱ�估����Ա�����Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    str����Ա����_Out = Nvl(objServiceCall.GetJsonNodeValue("output.operator_name"))
    str�Ǽ�ʱ��_Out = Nvl(objServiceCall.GetJsonNodeValue("output.create_time"))
    lng����id_out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.pati_id")))
    zl_ExseSvr_GetBillInfoByNo = True
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
Public Function zl_ExseSvr_GetFeeBillByCardNo(ByVal lng����ID As Long, ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
    ByRef strCardFeeNo_out As String, ByRef strPriceBillNo_out As String, ByRef int�շѱ�־ As Integer, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID������,��ȡ���õĵ�����Ϣ
    '���:lng����ID
    '     lngCardTypeID-�����ID
    '     strCardNo-����
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:str����Ա����_Out
    '     str�Ǽ�ʱ��_Out
    '     lng����id_Out
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
    
    'zl_ExseSvr_GetFeeBillByCardNo
    'input
    '    pati_id N   1   ����id
    '    cardtype_id N   1   �����id
    '    cardno  C   1   ����
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardtype_id", lngCardTypeID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardno", strCardNo, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetFeeBillByCardNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   "Ӧ����Ϣ��
    '    operator_name   C   1   ����Ա����
    '    create_time C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
    '    pati_id N   1   ����id

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡ����Ϊ" & strCardNo & "�ķ��õ�����Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strCardFeeNo_out = Nvl(objServiceCall.GetJsonNodeValue("output.feeno"))
    strPriceBillNo_out = Nvl(objServiceCall.GetJsonNodeValue("output.priceno"))
    int�շѱ�־ = Val(Nvl(objServiceCall.GetJsonNodeValue("output.charge_sign")))
    zl_ExseSvr_GetFeeBillByCardNo = True
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


Public Function zl_ExseSvr_GetPatiInvoiceClass(ByVal lng����ID As Long, ByVal lng��ҳid As Long, ByVal int���� As Integer, ByRef strʹ�����_Out As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ��ʹ�����
    '���:lng����id
    '     lng��ҳid
    '    int����
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:strʹ�����_Out
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
    
    'zl_ExseSvr_GetPatiInvoiceClass
    ' input
    '    pati_id N   1   ����id
    '    pati_pageid N   1   ��ҳid
    '    insure_type N   1   ����

    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng��ҳid, Json_num)
    strJson = strJson & "," & GetJsonNodeString("insure_type", int����, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetPatiInvoiceClass"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '       code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '       message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '       use_type    C   1   Ʊ��ʹ�����
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡƱ�ݵ�ʹ��������飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    strʹ�����_Out = Trim(Nvl(objServiceCall.GetJsonNodeValue("output.use_type")))
    zl_ExseSvr_GetPatiInvoiceClass = True
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
Public Function zl_ExseSvr_InvoiceClassUsed(ByVal intƱ�� As Integer, bln�Ƿ�����_Out As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ��Ʊ���Ƿ�������ʹ�����
    '���:intƱ��-1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:bln�Ƿ�����_Out
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
    
    'zl_ExseSvr_InvoiceClassUsed
    ' input
    '    inv_type    N   1   Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
 

    strJson = strJson & "" & GetJsonNodeString("inv_type", intƱ��, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_InvoiceClassUsed"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    '    output
    '       code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '       message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '       is_start    N   1   �Ƿ�����:1-�����˵ģ�0-δ����

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡƱ�ݵ�ʹ��������飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln�Ƿ�����_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.is_start"))) = 1
    zl_ExseSvr_InvoiceClassUsed = True
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

Public Function zl_ExseSvr_Actualmoney(ByVal str�ѱ� As String, ByVal lng�շ�ϸĿid As Long, _
    ByVal lng������Ŀid As Long, ByVal dblӦ�ս�� As Double, ByRef dblʵ�ս��_Out As Double, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, _
    Optional ByVal dbl���� As Double, Optional ByVal dbl�ɱ��� As Long, Optional ByVal lngҽ��id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷѱ������۽��
    '���:str�ѱ�- �ѱ�����
    '     lng�շ�ϸĿid-ϸĿid
    '     lng������Ŀid
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '     dbl����\dbl�ɱ���\lngҽ��id:ҩƷ�����ĲŴ���
    '����:dblʵ�ս��_Out
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
    
    'zl_ExseSvr_Actualmoney
    '    input           ����ʵ�ս��
    '        fee_category    C   1   �ѱ�
    '        fee_item_id N   1   �շ�ϸĿid
    '        income_item_id  N   1   ������Ŀid
    '        fee_amrcvb  N   1   Ӧ�ս��
    '        quantity    N   1   ����
    '        price_cost  N   1   �ɱ���
    '        order_id    N   1   ҽ��id
    strJson = strJson & "" & GetJsonNodeString("fee_category", str�ѱ�, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("fee_item_id", lng�շ�ϸĿid, Json_num)
    strJson = strJson & "," & GetJsonNodeString("income_item_id", lng������Ŀid, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_amrcvb", dblӦ�ս��, Json_num)
    strJson = strJson & "," & GetJsonNodeString("quantity", dbl����, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("price_cost", dbl�ɱ���, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("order_id", lngҽ��id, Json_num, True)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_Actualmoney"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    fee_category    C   1   �ѱ�
    '    net_receipts_fee    N   1   ʵ�ս��

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡָ�������µ�ҽ�����ݣ����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    dblʵ�ս��_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.fee_ampaib")))
    zl_ExseSvr_Actualmoney = True
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




Public Function Zl_Exsesvr_GetDepositDetail(ByVal lng����ID As Long, ByVal str��ʼʱ�� As String, ByVal str��ֹʱ�� As String, _
    ByVal int��ѯ���� As Integer, ByRef cllDpstDetail_Out As Collection, Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ѯָ�����˵�Ԥ����֧��ϸ���
    '���:lng����id-  ����id
    '     str��ʼʱ��- ��ʽ:yyyy-mm-dd hh24:mi:ss
    '     str��ֹʱ��-��ʽ:yyyy-mm-dd hh24:mi:ss
    '     int��ѯ����-0-����,1-����;2-סԺ

    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '     dbl����\dbl�ɱ���\lngҽ��id:ҩƷ�����ĲŴ���
    '����:cllDpstDetail_Out-Ԥ����֧��ϸ����
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
    
    Err = 0: On Error GoTo errHandle
    'Zl_Exsesvr_GetDepositDetail
    '    input
    '        pati_id N   1   ����id
    '        begin_time  C   1   ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
    '        end_time    C   1   ��ֹʱ��:yyyy-mm-dd hh24:mi:ss
    '        type_sign   N   1   ���ͱ�־:0-����,1-����;2-סԺ
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("begin_time", str��ʼʱ��, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("end_time", str��ֹʱ��, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("type_sign", int��ѯ����, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Exsesvr_GetDepositDetail"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '
    '    output
    '        code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '        message C   1   "Ӧ����Ϣ��
    '        item_list C
    '            business_type   C   1   ҵ������:�ڳ�����ֵ���շ��á����ʵ�
    '            happen_time C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '            earlystage  N   1   �ڳ����
    '            recharge    N   1   ���ڳ�ֵ
    '            consume N   1   ��������

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡָ�������µ�Ԥ����֧��ϸ���ݣ����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllDpstDetail_Out = objServiceCall.GetJsonListValue("output.item_list")
    Zl_Exsesvr_GetDepositDetail = True
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



Public Function zl_ExseSvr_BillInHistory(ByVal strNO As String, ByVal int�������� As Integer, _
    ByVal bln���� As Boolean, ByRef bln���ں󱸱�_Out As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ�����õ����Ƿ����󱸱�ռ�
    '���:strNO- ���ݺ�
    '     int��������-1-�շѵ�,2-Ԥ����,3-���ʵ�,4-�Һŵ�,5-���￨����,6-���ʵ���;7-�Զ����ʵ�
    '     bln����-�Ƿ�����
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:bln���ں󱸱�_Out
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
    
    'zl_ExseSvr_BillInHistory
    '    input
    '            bill_no C   1   ���ݺ�
    '            bill_type   C   1   ��������:1-�շѵ�,2-Ԥ����,3-���ʵ�,4-�Һŵ�,5-���￨����,6-���ʵ���;7-�Զ����ʵ�
    '            outpati_flag    N       �����־��1-���2-סԺ
    '

    strJson = strJson & "" & GetJsonNodeString("bill_no", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("bill_type", int��������, Json_num)
    strJson = strJson & "," & GetJsonNodeString("outpati_flag", IIf(bln����, 1, 2), Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_BillInHistory"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   "Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    exits_history   C   1   ������ʷ�󱸱�:1-����;1-������


    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "�����ж��Ƿ���Ч����ʷת�����ݣ����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln���ں󱸱�_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.exits_history"))) = 1
    zl_ExseSvr_BillInHistory = True
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

Public Function zl_Exsesvr_BillIsPrintInvoice(ByVal strNO As String, ByVal int�������� As Integer, _
    ByVal intƱ�� As Integer, ByRef bln�Ƿ��ӡ_Out As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ���ĵ����Ƿ��Ѿ���ӡ��Ʊ��
    '���:strNO- ���ݺ�
    '     int��������-1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
    '     intƱ��-1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:bln�Ƿ��ӡ_Out
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
    
    'Zl_Exsesvr_Billisprintinvoice
    '    input
    '        fee_no  C   1   ���ݺ�
    '        bill_type   N   1   ��������:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
    '        inv_type    N   1   Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨


    strJson = strJson & "" & GetJsonNodeString("bill_no", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("bill_type", int��������, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_type", intƱ��, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Exsesvr_Billisprintinvoice"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    printed N   1   �Ƿ��ӡ:1-�Ѵ�ӡ;0-δ��ӡ

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "�����ж��Ƿ��ӡ��Ʊ�ݣ����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln�Ƿ��ӡ_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.printed"))) = 1
    zl_Exsesvr_BillIsPrintInvoice = True
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

Public Function zl_ExseSvr_GetCardFeeInfoByNo(ByVal strNO As String, ByVal int��ѯ���� As Integer, _
    ByRef cllFeeData_out As Collection, ByRef cllPriceBill_out As Collection, ByRef cllBalance_Out As Collection, ByRef cllDeposit_Out As Collection, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String, Optional bln�Ƿ����Ԥ�� As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���Ŀ��ѵ��ݣ���ȡ���ü����㼰Ԥ����Ϣ
    '���:strNO- ���ݺ�
    '     int��ѯ���ͣ�0-��ȡ��������:1-��ȡ���ϵ���;2-ʣ����õ���
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:cllFeeData_Out-�������õ���Ϣ
    '     cllPriceBill_out-���۵���Ϣ
    '     cllBalance_out-����������Ϣ
    '     cllDeposit_Out-����ͬʱ��Ԥ����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
     
     Set cllFeeData_out = New Collection: Set cllBalance_Out = New Collection: Set cllDeposit_Out = New Collection
     Set cllPriceBill_out = New Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetCardFeeInfoByNo
    '    input
    '       fee_no  C   1   ���ݺ�:���õ��ݺ�
    '       query_type  N   ��ѯ���ͣ�0-��ȡ��������:1-��ȡ���ϵ���;2-ʣ����õ���
    '       query_deposit N 1 �Ƿ����Ԥ��:1-����Ԥ����Ϣ��0-������Ԥ����Ϣ
    strJson = strJson & "" & GetJsonNodeString("fee_no", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("query_type", int��ѯ����, Json_num)
    strJson = strJson & "," & GetJsonNodeString("query_deposit", IIf(bln�Ƿ����Ԥ��, 1, 0), Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetCardFeeInfoByNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   "Ӧ����Ϣ��
    '    fee_list            [����]ÿ������ID��Ϣ
    '        fee_id  N   1   ����id
    '        fee_num N   1   ���
    '        pati_id N   1   ����id
    '        pati_name   C   1   ����
    '        pati_sex    C   1   �Ա�
    '        pati_age    C   1   ����
    '        fee_category    C   1   �ѱ�
    '        item_id N   1   �շ���Ŀid
    '        income_item_id  N   1   ������Ŀid
    '        quantity    N   1   ����
    '        fee_amrcvb  N   1   Ӧ�ս��
    '        fee_ampaid  N   1   ʵ�ս��
    '        placer  C   1   ������
    '        operator_code   C   1   ����Ա���
    '        operator_name   C   1   ����Ա����
    '        create_time D   1   �Ǽ�ʱ��
    '        happen_time D   1   ����ʱ��
    '        rec_status  N   1   ��¼״̬
    '        mrbkfee_sign N   1   �Ƿ�����:1-�ǲ�����;0-���ǲ�����
    '        invoice_no  N   1   ��Ʊ��
    '        kpbooks_sign N   1   �Ƿ����:1-�Ǽ���;0-����
    '        fee_status   N   1   ����״̬:1-�쳣״̬;0-��������
    '        cardtype_id N   1   �����ID
    '        card_no C   1   ����
    '    pricebill_info  C       �������ɻ��۷�����Ϣ
    '        fee_no  C       ���۵���
    '        fee_amrcvb  N   1   Ӧ�ս��
    '        fee_ampaid  N   1   ʵ�ս��
    '        charged_statu   N   1   �շ�״̬:0-δ�շ�;1-���շ�;2-��ȫ��
    '    balance_list[]  C       ������Ϣ�б�
    '        blnc_mode   C   1   ���㷽ʽ����
    '        balance_id  N   1   ����ID
    '        blnc_money  N   1   ���ʽ��
    '        pay_cardno  N   1   ֧������
    '        pay_swapno  C   1   ������ˮ��
    '        pay_swapmemo    C   1   ����˵��
    '        relation_id N   1   ��������id
    '        cardtype_id N   1   �����id
    '        consume_card    N   1   �Ƿ����ѿ�:1-��;0-����
    '        blnc_nature N   1   ��������:1-�ֽ���㷽ʽ,2-������ҽ������ , 8-���㿨���� ,9-����
    '        blnc_statu  N   1   ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
    '        consume_card_id N   1   ���ѿ�id
    '        original_money N,1 ԭʼ���,��ʣ�����ʱ����
    '        original_id N 1 ԭ����ID:����ʱ����
    '    deposit_info    C       Ԥ����Ϣ
    '        deposit_id  N   1   Ԥ��id
    '        deposit_no  C   1   Ԥ�����ݺ�
    '        deposit_money   N   1   Ԥ�����
    '        blnc_mode   C   1   ���㷽ʽ
    '        pay_cardno  N   1   ֧������
    '        pay_swapno  C   1   ������ˮ��
    '        pay_swapmemo    C   1   ����˵��
    '        relation_id N   1   ��������id
    '        cardtype_id N   1   �����id
    '        consume_card    N   1   �Ƿ����ѿ�:1-��;0-����
    '        blnc_nature N   1   ��������
    '        blnc_statu  N   1   ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
    '        consume_card_id N   1   ���ѿ�id
    '        blnc_no C   1   �������
    '        blnc_memo   C   1   ժҪ

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡ���ݺ�Ϊ��" & strNO & "���ĵ�����Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '��ȡ������Ϣ��
    Set cllFeeData_out = objServiceCall.GetJsonListValue("output.fee_list")
    Set cllBalance_Out = objServiceCall.GetJsonListValue("output.balance_list")
    
    Set cllDeposit_Out = objServiceCall.GetJsonListValue("output.deposit_info")
    '���۵�
   Set cllPriceBill_out = objServiceCall.GetJsonListValue("output.pricebill_info")
    
    If cllFeeData_out Is Nothing Then Set cllFeeData_out = New Collection
    If cllBalance_Out Is Nothing Then Set cllBalance_Out = New Collection
    If cllDeposit_Out Is Nothing Then Set cllDeposit_Out = New Collection
    If cllPriceBill_out Is Nothing Then Set cllPriceBill_out = New Collection
    
    zl_ExseSvr_GetCardFeeInfoByNo = True
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

Public Function zl_ExseSvr_GetMrbkFeeInfo(ByVal lng����ID As Long, ByVal str���ݺ� As String, int��¼״̬ As Integer, _
     ByRef cllFeeData_out As Collection, Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���id��ȡ�����������Ϣ
    '���:lng����ID-����ID
    '     str���ݺ�
    '     int��¼״̬-(1,3)-ԭʼ��¼,2-���ʼ�¼
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:cllFeeData_out-���ط������ݼ�
    '
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Collection
    Dim objServiceCall As Object
    Dim intReturn As Integer
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'zl_ExseSvr_GetMrbkFeeInfo
    '    input
    '        pati_id N   1   ����id
    '        fee_no  C       ���ݺ�:���������漰�ĵ��ݺ�
    '        rec_status  N   1   ��¼״̬:1-ԭʼ��¼;2-��������
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rec_status", int��¼״̬, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_no", str���ݺ�, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_ExseSvr_GetMrbkFeeInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    
    ' output
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   "Ӧ����Ϣ
    '    fee_list[]  C       ������ϸ����
    '        fee_no  C   1   ���ݺ�
    '        fee_num N   1   ���
    '        pati_id N   1   ����id
    '        pati_name   C   1   ����
    '        pati_sex    C   1   �Ա�
    '        pati_age    C   1   ����
    '        fee_category    C   1   �ѱ�
    '        fee_status  N   1   ����״̬:1-�쳣״̬;0-��������
    '        rec_status  N   1   ��¼״̬:1-������¼;2-���ʼ�¼;3-�����ʵļ�¼
    '        charge_sign        N       1       �շѱ�־:0-����;1-����;2-���۵�
    '        fee_ampaid  N   1   ʵ�ս��
    '        happen_time C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
    '        operator_name   C   1   ����Ա����
    '        memo    C   1   ժҪ
    '        pricebill_no    C   1   ���۵���
    '        price_charged   N   1   �������շ�:1-���۵��Ѿ����շѴ����շ�;0-δ�շ�
    '        balance_info  C       ������Ϣ�б�
    '            blnc_mode   C   1   ���㷽ʽ����
    '            balance_id  N   1   ����ID����ѯ���ϵĵ���ʱΪ����ID
    '            blnc_money  N   1   ���ʽ��
    '            pay_cardno  N   1   ֧������
    '            pay_swapno  C   1   ������ˮ��
    '            pay_swapmemo    C   1   ����˵��
    '            relation_id N   1   ��������id
    '            cardtype_id N   1   �����id
    '            consume_card    N   1   �Ƿ����ѿ�:1-��;0-����
    '            blnc_nature N   1   ��������:1-�ֽ���㷽ʽ,2-������ҽ������ , 8-���㿨���� ,9-����
    '            blnc_statu  N   1   ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
    '            consume_card_id N   1   ���ѿ�id
    '            blnc_no C   1   �������
    '            blnc_memo   C   1   ժҪ
    '            original_id N   1   ԭ����ID:����ʱ����
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡָ�����˵Ĳ�����,���飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set cllFeeData_out = objServiceCall.GetJsonListValue("output.fee_list")
    
    For i = 1 To cllFeeData_out.Count
        Set cllTemp = objServiceCall.GetJsonListValue("output.fee_list[" & i - 1 & "].balance_info")
        If Not cllTemp Is Nothing Then
            Call RemoveCollectionItemFromKey(cllFeeData_out(i), "_balance_info")
            cllFeeData_out(i).Add cllTemp, "_balance_info"
         End If
    Next
    
    
    zl_ExseSvr_GetMrbkFeeInfo = True
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

Public Function zlGetCardFeeDataFromColl(ByVal strNO As String, ByVal cllCardFee As Collection, _
    ByRef rsCardFee_Out As Recordset, Optional ByRef objBalanceItems_Out As clsBalanceItems, Optional ByRef dblMoney_Out As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��񷵻صļ���,����¼����ʽ������Ϣ
    '���:cllCardFee-��ǰ����
    '
    '����:rsCardFee_Out-���صĿ����ü���
    '     objBalanceItems_out-������Ϣ�б���Ҫ�ǿ��ܴ��ڼ��ʣ���Ҫ��objBalanceItems_out
    '     dblMoney_Out:ʵ�ս��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection
    Dim i As Long, bln���� As Boolean
    
    On Error GoTo errHandle
    
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    dblMoney_Out = 0
    Set rsCardFee_Out = New ADODB.Recordset
    With rsCardFee_Out
        If .State = adStateOpen Then .Close
        
        .Fields.Append "���ݺ�", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "����id", adBigInt, , adFldIsNullable
        .Fields.Append "���", adBigInt, , adFldIsNullable
        .Fields.Append "����id", adBigInt, , adFldIsNullable
        .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�ѱ�", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�շ���Ŀid", adBigInt, , adFldIsNullable
        .Fields.Append "������Ŀid", adBigInt, , adFldIsNullable
        .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "Ӧ�ս��", adDouble, , adFldIsNullable
        .Fields.Append "ʵ�ս��", adDouble, , adFldIsNullable
        
        .Fields.Append "������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����Ա���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����Ա����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "�Ǽ�ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��¼״̬", adBigInt, , adFldIsNullable
        
        .Fields.Append "�Ƿ�����", adBigInt, , adFldIsNullable
        .Fields.Append "��Ʊ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "�Ƿ����", adBigInt, , adFldIsNullable
        .Fields.Append "����״̬", adBigInt, , adFldIsNullable
        .Fields.Append "�����ID", adBigInt, , adFldIsNullable
        .Fields.Append "����", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "�Ƿ�Һŷ���", adBigInt, , adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    If cllCardFee Is Nothing Then Exit Function
    '    fee_id  N   1   ����id
    '    fee_num N   1   ���
    '    pati_id N   1   ����id
    '    pati_name   C   1   ����
    '    pati_sex    C   1   �Ա�
    '    pati_age    C   1   ����
    '    fee_category    C   1   �ѱ�
    '    item_id N   1   �շ���Ŀid
    '    income_item_id  N   1   ������Ŀid
    '    quantity    N   1   ����
    '    fee_amrcvb  N   1   Ӧ�ս��
    '    fee_ampaid  N   1   ʵ�ս��
    '    placer  C   1   ������
    '    operator_code   C   1   ����Ա���
    '    operator_name   C   1   ����Ա����
    '    create_time D   1   �Ǽ�ʱ��
    '    happen_time D   1   ����ʱ��
    '    rec_status  N   1   ��¼״̬
    '    mrbkfee_sign N   1   �Ƿ�����:1-�ǲ�����;0-���ǲ�����
    '    invoice_no  N   1   ��Ʊ��
    '    kpbooks_sign N   1   ���ʱ�־:1-�Ǽ���;0-����
    '    fee_status   N   1   ����״̬:1-�쳣״̬;0-��������
    '    cardtype_id N   1   �����ID
    '    card_no C   1   ����
    '    sendcard_reg    N   1   �Ƿ�ҹҺ�ͬ������:1-�ǹҺ�ͬʱ����;0-�ǹҺ�ͬʱ����

    For i = 1 To cllCardFee.Count
        Set cllTemp = cllCardFee(i)
        
        If Not bln���� Then bln���� = Val(Nvl(cllTemp("_kpbooks_sign"))) = 1
        With rsCardFee_Out
            .AddNew
            !���ݺ� = strNO
            !����id = Val(Nvl(cllTemp("_fee_id")))
            !��� = Val(Nvl(cllTemp("_fee_num")))
            !����ID = Val(Nvl(cllTemp("_pati_id")))
            !���� = Nvl(cllTemp("_pati_name"))
            !�Ա� = Nvl(cllTemp("_pati_sex"))
            !���� = Nvl(cllTemp("_pati_age"))
            !�ѱ� = Nvl(cllTemp("_fee_category"))
            !�շ���Ŀid = Val(Nvl(cllTemp("_item_id")))
            !������ĿID = Val(Nvl(cllTemp("_income_item_id")))
            !���� = Val(Nvl(cllTemp("_quantity")))
            !Ӧ�ս�� = Val(Nvl(cllTemp("_fee_amrcvb")))
            !ʵ�ս�� = Val(Nvl(cllTemp("_fee_ampaid")))
            !������ = Nvl(cllTemp("_placer"))
            !����Ա��� = Nvl(cllTemp("_operator_code"))
            !����Ա���� = Nvl(cllTemp("_operator_name"))
            !�Ǽ�ʱ�� = Nvl(cllTemp("_create_time"))
            !����ʱ�� = Nvl(cllTemp("_happen_time"))
            !��¼״̬ = Val(Nvl(cllTemp("_rec_status")))
            
            !�Ƿ����� = Val(Nvl(cllTemp("_mrbkfee_sign")))
            !��Ʊ�� = Nvl(cllTemp("_invoice_no"))
            !�Ƿ���� = Val(Nvl(cllTemp("_kpbooks_sign")))
            !����״̬ = Val(Nvl(cllTemp("_fee_status")))
            !�����ID = Val(Nvl(cllTemp("_cardtype_id")))
            !���� = Nvl(cllTemp("_card_no"))
            !�Ƿ�Һŷ��� = Val(Nvl(cllTemp("_sendcard_reg")))
            .Update
            dblMoney_Out = RoundEx(dblMoney_Out + Val(Nvl(rsCardFee_Out!ʵ�ս��)), 5)
        End With
    Next
    If bln���� Then
        objBalanceItems_Out.���� = gEM_���ʵ�
    End If
    objBalanceItems_Out.������ = dblMoney_Out
    objBalanceItems_Out.���ݺ� = strNO
    Set cllTemp = Nothing
    zlGetCardFeeDataFromColl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardFeeBalanceFromColl(ByVal cllCardFeeBalance As Collection, ByRef rsBalance_Out As Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��񷵻صļ���,����¼����ʽ���ؽ�����Ϣ
    '���:cllCardFeeBalance-��ǰ����
    '
    '����:rsBalance_out-���صĿ��ѽ�����Ϣ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long
    On Error GoTo errHandle
    
    Set rsBalance_Out = New ADODB.Recordset
    With rsBalance_Out
        If .State = adStateOpen Then .Close
        .Fields.Append "���㷽ʽ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ID", adBigInt, , adFldIsNullable
        .Fields.Append "���ʽ��", adDouble, , adFldIsNullable
        
        .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "������ˮ��", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "����˵��", adLongVarChar, 4000, adFldIsNullable
        .Fields.Append "��������id", adBigInt, , adFldIsNullable
        
        .Fields.Append "�����id", adBigInt, , adFldIsNullable
        .Fields.Append "�Ƿ����ѿ�", adSmallInt, , adFldIsNullable
        .Fields.Append "��������", adSmallInt, , adFldIsNullable
        .Fields.Append "����״̬", adSmallInt, , adFldIsNullable
        .Fields.Append "���ѿ�id", adBigInt, , adFldIsNullable
         
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    If cllCardFeeBalance Is Nothing Then Exit Function
    '    blnc_mode   C   1   ���㷽ʽ����
    '    balance_id  N   1   ����ID
    '    blnc_money  N   1   ���ʽ��
    '    pay_cardno  N   1   ֧������
    '    pay_swapno  C   1   ������ˮ��
    '    pay_swapmemo    C   1   ����˵��
    '    relation_id N   1   ��������id
    '    cardtype_id N   1   �����id
    '    consume_card    N   1   �Ƿ����ѿ�:1-��;0-����
    '    blnc_nature N   1   ��������:1-�ֽ���㷽ʽ,2-������ҽ������ , 8-���㿨���� ,9-����
    '    blnc_statu  N   1   ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
    '    consume_card_id N   1   ���ѿ�id
    For i = 1 To cllCardFeeBalance.Count
        Set cllTemp = cllCardFeeBalance(i)
        With rsBalance_Out
            .AddNew
                !���㷽ʽ = Nvl(cllTemp("_blnc_mode"))
                !����id = Val(Nvl(cllTemp("_balance_id")))
                !���ʽ�� = Val(Nvl(cllTemp("_blnc_money")))
                !���� = Nvl(cllTemp("_pay_cardno"))
                !������ˮ�� = Nvl(cllTemp("_pay_swapno"))
                !����˵�� = Nvl(cllTemp("_pay_swapmemo"))
                !��������ID = Val(Nvl(cllTemp("_relation_id")))
                !�����ID = Val(Nvl(cllTemp("_cardtype_id")))
                !�Ƿ����ѿ� = Val(Nvl(cllTemp("_consume_card")))
                !�������� = Val(Nvl(cllTemp("_blnc_nature")))
                !����״̬ = Val(Nvl(cllTemp("_blnc_statu")))
                !���ѿ�ID = Val(Nvl(cllTemp("_consume_card_id")))
            .Update
        End With
    Next
    zlGetCardFeeBalanceFromColl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceItemsFromCardFeeColl(ByVal strNO As String, ByVal cllCardFeeBalance As Collection, ByRef objBalanceItems_Out As clsBalanceItems, _
    Optional ByVal bln�鿴���� As Boolean, Optional blnDelFee As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��񷵻صļ���,����¼����ʽ���ؽ�����Ϣ
    '���:cllCardFeeBalance-��ǰ����
    '     strNo-���õ��ݺ�
    '     bln�鿴����-��ǰ���ĵ������ϵ���
    '     blnDelFee-��ǰΪ�˷Ѳ���
    '����:objBalanceItems_Out-���صĿ��ѽ�����Ϣ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long
    Dim objItem As clsBalanceItem, objCard As Card
    Dim dbl���� As Double
    On Error GoTo errHandle
    
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    
    If cllCardFeeBalance Is Nothing Then Exit Function
    If objBalanceItems_Out.���� = gEM_���ʵ� Then zlGetBalanceItemsFromCardFeeColl = True: Exit Function
    
    '    blnc_mode   C   1   ���㷽ʽ����
    '    balance_id  N   1   ����ID
    '    blnc_money  N   1   ���ʽ��
    '    pay_cardno  N   1   ֧������
    '    pay_swapno  C   1   ������ˮ��
    '    pay_swapmemo    C   1   ����˵��
    '    relation_id N   1   ��������id
    '    cardtype_id N   1   �����id
    '    consume_card    N   1   �Ƿ����ѿ�:1-��;0-����
    '    blnc_nature N   1   ��������:1-�ֽ���㷽ʽ,2-������ҽ������ , 8-���㿨���� ,9-����
    '    blnc_statu  N   1   ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
    '    consume_card_id N   1   ���ѿ�id
    '    blnc_no C   1   �������
    '    blnc_memo   C   1   ժҪ
    
    objBalanceItems_Out.������ = 0
    For i = 1 To cllCardFeeBalance.Count
        Set cllTemp = cllCardFeeBalance(i)
        Set objItem = New clsBalanceItem
        Set objCard = zlGetCardFromCardType(Val(Nvl(cllTemp("_cardtype_id"))), Val(Nvl(cllTemp("_consume_card"))) = 1, Nvl(cllTemp("_blnc_mode")))
        If Val(Nvl(cllTemp("_blnc_nature"))) = 9 Then
            dbl���� = RoundEx(dbl���� + Val(Nvl(cllTemp("_blnc_money"))), 6)
        Else
            With objItem
                Set .objCard = objCard
                .���ݺ� = strNO
                .�������� = 5   ' 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ�
                .���㷽ʽ = Nvl(cllTemp("_blnc_mode"))
                .������ = Val(Nvl(cllTemp("_blnc_money")))
                .��������ID = Val(Nvl(cllTemp("_relation_id")))
                .������ˮ�� = Nvl(cllTemp("_pay_swapno"))
                .����˵�� = Nvl(cllTemp("_pay_swapmemo"))
                .������� = Nvl(cllTemp("_blnc_no"))
                .�������� = Val(Nvl(cllTemp("_blnc_nature")))
                .����ժҪ = Nvl(cllTemp("_blnc_memo"))
                .���� = Nvl(cllTemp("_pay_cardno"))
                
                .�����ID = Val(Nvl(cllTemp("_cardtype_id")))
                .���ѿ�ID = Val(Nvl(cllTemp("_consume_card_id")))
                .���ѿ� = Val(Nvl(cllTemp("_consume_card"))) = 1
                .�Ƿ����� = objCard.�������Ĺ��� <> ""
                .ԭʼ��� = .������
                .δ�˽�� = .������
                .�Ƿ�����༭ = False
                .�Ƿ�����ɾ�� = False
                .У�Ա�־ = Val(Nvl(cllTemp("_blnc_statu")))
                .�Ƿ���� = .У�Ա�־ = 2 Or .У�Ա�־ = 0
                .���� = ""
                .�ʻ���� = 0
                If .�����ID = 0 Then   '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                      .�������� = 0
                ElseIf .�����ID <> 0 And .���ѿ� = False Then
                      .�������� = 3
                ElseIf .�����ID <> 0 And .���ѿ� Then
                      .�������� = 5
                Else
                     .�������� = 0
                End If
                .�Ƿ��˿� = blnDelFee
                If bln�鿴���� Then
                    .����ID = Val(Nvl(cllTemp("_balance_id")))
                    .����ID = Val(Nvl(cllTemp("_original_id"))) 'ԭ����ID
                   
                Else
                    .����ID = Val(Nvl(cllTemp("_balance_id")))
                    .����ID = Val(Nvl(cllTemp("_original_id"))) 'ԭ����ID
                End If
                .�Ƿ�Ԥ�� = False
            End With
            objBalanceItems_Out.AddItem objItem
            objBalanceItems_Out.������ = RoundEx(objBalanceItems_Out.������ + objItem.������, 6)
            objBalanceItems_Out.���ݺ� = objItem.���ݺ�
            If objItem.�����ID <> 0 Then
                objBalanceItems_Out.���� = IIf(objItem.���ѿ�, gEM_���ѿ�, gEM_һ��ͨ)
            Else
                objBalanceItems_Out.���� = gEM_��ͨ����
            End If
        End If
    Next
    
    objBalanceItems_Out.���� = dbl����
    objBalanceItems_Out.δ�˽�� = objBalanceItems_Out.������
    objBalanceItems_Out.ԭʼ��� = objBalanceItems_Out.������ '�ݶ�Ϊδ�˲���
    
    zlGetBalanceItemsFromCardFeeColl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceItemsFromDepositColl(ByVal cllDeposits As Collection, ByRef objBalanceItems_Out As clsBalanceItems) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��񷵻صļ���,����¼����ʽ���ؽ�����Ϣ
    '���:cllDeposits-��ǰ����
    '
    '����:objBalanceItems-���صĿ��ѽ�����Ϣ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long
    Dim objItem As clsBalanceItem, objCard As Card
    
    On Error GoTo errHandle
    
    If objBalanceItems_Out Is Nothing Then Set objBalanceItems_Out = New clsBalanceItems
    
    If cllDeposits Is Nothing Then zlGetBalanceItemsFromDepositColl = True: Exit Function
    If cllDeposits.Count = 0 Then zlGetBalanceItemsFromDepositColl = True: Exit Function
    
    '    deposit_id  N   1   Ԥ��id
    '    deposit_no  C   1   Ԥ�����ݺ�
    '    deposit_money   N   1   Ԥ�����
    '    blnc_mode   C   1   ���㷽ʽ
    '    pay_cardno  N   1   ֧������
    '    pay_swapno  C   1   ������ˮ��
    '    pay_swapmemo    C   1   ����˵��
    '    relation_id N   1   ��������id
    '    cardtype_id N   1   �����id
    '    consume_card    N   1   �Ƿ����ѿ�:1-��;0-����
    '    blnc_nature N   1   ��������
    '    blnc_statu  N   1   ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
    '    consume_card_id N   1   ���ѿ�id
    '    blnc_no C   1   �������
    '    blnc_memo   C   1   ժҪ

     Set cllTemp = cllDeposits
     Set objItem = New clsBalanceItem
     Set objCard = zlGetCardFromCardType(Val(Nvl(cllTemp("_cardtype_id"))), Val(Nvl(cllTemp("_consume_card"))) = 1, Nvl(cllTemp("_blnc_mode")))
     With objItem
         .�������� = 1 ' 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ�
         .���ݺ� = Nvl(cllTemp("_deposit_no"))
         .���㷽ʽ = Nvl(cllTemp("_blnc_mode"))
         .Ԥ��ID = Val(Nvl(cllTemp("_deposit_id")))
         .������ = Val(Nvl(cllTemp("_deposit_money")))
         .��������ID = Val(Nvl(cllTemp("_relation_id")))
         .������ˮ�� = Nvl(cllTemp("_pay_swapno"))
         .����˵�� = Nvl(cllTemp("_pay_swapmemo"))
         .������� = Nvl(cllTemp("_blnc_no"))
         .�������� = Val(Nvl(cllTemp("_blnc_nature")))
         .����ժҪ = Nvl(cllTemp("_blnc_memo"))
         .���� = Nvl(cllTemp("_pay_cardno"))
         
         .�����ID = Val(Nvl(cllTemp("_cardtype_id")))
         .���ѿ�ID = Val(Nvl(cllTemp("_consume_card_id")))
         .���ѿ� = Val(Nvl(cllTemp("_consume_card"))) = 1
         .�Ƿ����� = objCard.�������Ĺ��� <> ""
         .ԭʼ��� = .������
         .δ�˽�� = .������
         .�Ƿ��˿� = .������ < 0
         .�Ƿ�����༭ = False
         .�Ƿ�����ɾ�� = False
         .У�Ա�־ = Val(Nvl(cllTemp("_blnc_statu")))
         .�Ƿ���� = .У�Ա�־ = 2 Or .У�Ա�־ = 0
         .���� = ""
         .�ʻ���� = 0
         If .�����ID = 0 Then   '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
               .�������� = 0
         ElseIf .�����ID <> 0 And .���ѿ� = False Then
               .�������� = 3
         ElseIf .�����ID <> 0 And .���ѿ� Then
               .�������� = 5
         Else
              .�������� = 0
         End If
         .����ID = Val(Nvl(cllTemp("_deposit_id")))
         .����ID = 0
         .�Ƿ�Ԥ�� = False
     End With
    objBalanceItems_Out.AddItem objItem
    objBalanceItems_Out.������ = RoundEx(objBalanceItems_Out.������ + objItem.������, 6)
 
    zlGetBalanceItemsFromDepositColl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetDepsoitFromColl(ByVal cllDeposit As Collection, ByRef rsDeposit_Out As Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��񷵻صļ���,����¼����ʽ���ؽ�����Ϣ
    '���:cllDeposit-��ǰ����
    '
    '����:rsBalance_out-���صĿ��ѽ�����Ϣ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection, i As Long
    On Error GoTo errHandle
    
    Set rsDeposit_Out = New ADODB.Recordset
    With rsDeposit_Out
        If .State = adStateOpen Then .Close
        .Fields.Append "Ԥ�����ݺ�", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ԥ�����", adDouble, , adFldIsNullable
        .Fields.Append "���㷽ʽ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "������ˮ��", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "����˵��", adLongVarChar, 4000, adFldIsNullable
        .Fields.Append "��������id", adBigInt, , adFldIsNullable
        
        .Fields.Append "�����id", adBigInt, , adFldIsNullable
        .Fields.Append "�Ƿ����ѿ�", adSmallInt, , adFldIsNullable
        .Fields.Append "��������", adSmallInt, , adFldIsNullable
        .Fields.Append "����״̬", adSmallInt, , adFldIsNullable
        .Fields.Append "���ѿ�id", adBigInt, , adFldIsNullable
         
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    
    If cllDeposit Is Nothing Then Exit Function
    '    deposit_no  C   1   Ԥ�����ݺ�
    '    deposit_money   N   1   Ԥ�����
    '    blnc_mode   C   1   ���㷽ʽ
    '    pay_cardno  N   1   ֧������
    '    pay_swapno  C   1   ������ˮ��
    '    pay_swapmemo    C   1   ����˵��
    '    relation_id N   1   ��������id
    '    cardtype_id N   1   �����id
    '    consume_card    N   1   �Ƿ����ѿ�:1-��;0-����
    '    blnc_nature N   1   ��������
    '    blnc_statu  N   1   ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
    '    consume_card_id N   1   ���ѿ�id

    For i = 1 To cllDeposit.Count
        Set cllTemp = cllDeposit(i)
        With rsDeposit_Out
            .AddNew
                
                !Ԥ�����ݺ� = Nvl(cllTemp("_deposit_no"))
                !Ԥ����� = Val(Nvl(cllTemp("_deposit_money")))
                !���㷽ʽ = Nvl(cllTemp("_blnc_mode"))
                !���� = Nvl(cllTemp("_pay_cardno"))
                !������ˮ�� = Nvl(cllTemp("_pay_swapno"))
                !����˵�� = Nvl(cllTemp("_pay_swapmemo"))
                !��������ID = Val(Nvl(cllTemp("_relation_id")))
                !�����ID = Val(Nvl(cllTemp("_cardtype_id")))
                !�Ƿ����ѿ� = Val(Nvl(cllTemp("_consume_card")))
                !�������� = Val(Nvl(cllTemp("_blnc_nature")))
                !����״̬ = Val(Nvl(cllTemp("_blnc_statu")))
                !���ѿ�ID = Val(Nvl(cllTemp("_consume_card_id")))
            .Update
        End With
    Next
    zlGetDepsoitFromColl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Public Function Zl_Exsesvr_UpdatePatiBaseInfo(ByVal lng����ID As Long, Optional ByVal cllUpdateInfo As Collection, _
    Optional ByVal str����Ա���� As String, Optional ByVal str����Ա���� As String, Optional blnShowErrMsg As Boolean, _
    Optional ByRef str����ID As String = "", Optional int���� As Integer = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������Ϣ
    '���:cllPatiUpdate-�޸ĵĲ�����Ϣ:array(����,ֵ)
    '                ���ư���������,,�Ա�,����,������(�����))
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '     int����-����:1-����;2-סԺ
    '     str����ID-��ҳID����Ϊ"",������ֵ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    Dim strJsonPatiInfo  As String
    
    If cllUpdateInfo Is Nothing Then Exit Function
    If cllUpdateInfo.Count = 0 Then Exit Function
    If lng����ID = 0 Then Exit Function
    
    If blnShowErrMsg Then On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If Not blnShowErrMsg Then Err.Raise -1001, strErrMsg, strErrMsg: Exit Function
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strJsonPatiInfo = ""
    For i = 1 To cllUpdateInfo.Count
        varTemp = cllUpdateInfo(i)
        Select Case UCase(varTemp(0))
        Case "����", "��������"
            strJsonPatiInfo = strJsonPatiInfo & "," & GetJsonNodeString("pati_name", Trim(varTemp(1)), Json_Text)
        Case "�Ա�"
            strJsonPatiInfo = strJsonPatiInfo & "," & GetJsonNodeString("pati_sex", Trim(varTemp(1)), Json_Text)
        Case "����"
            strJsonPatiInfo = strJsonPatiInfo & "," & GetJsonNodeString("pati_age", Trim(varTemp(1)), Json_Text)
        Case "�����"
            strJsonPatiInfo = strJsonPatiInfo & "," & GetJsonNodeString("outpatient_num", Trim(varTemp(1)), Json_Text)
        Case Else
        End Select
    Next
    If strJsonPatiInfo = "" Then Exit Function
    strJsonPatiInfo = Mid(strJsonPatiInfo, 2)
    
    strJsonPatiInfo = GetNodeString("update_info") & ":{" & strJsonPatiInfo & "}"
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng����ID, Json_num, True)
    If str����ID <> "" Then
        strJson = strJson & "," & GetJsonNodeString("visit_id", Val(str����ID), Json_num, True)
    End If
    strJson = strJson & "," & GetJsonNodeString("occasion", int����, Json_num, True)
    
    strJson = strJson & "," & GetJsonNodeString("operator_name", str����Ա����, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_code", str����Ա����, Json_Text)
    strJson = strJson & "," & strJsonPatiInfo
    'Zl_Exsesvr_UpdatePatiBaseInfo
    '    pati_id N   1   ����id
    '    visit_id    N   1   ����id ���ﲡ��Ϊ�Һ�ID;סԺ����Ϊ��ҳID;Ϊ0˵����������ǹҺž���Ĳ���(����id_InΪ��ʱ,�����ĸò��˵ķ��ò��ֵ�ҵ������)
    '    occasion    N       ����,1-����;2-סԺ
    '    update_info
    '        outpatient_num  C       �����
    '        pati_name   C       ����
    '        pati_age        C       ����
    '        pati_sex        C       �Ա�
    '        pati_birthdate  C       ��������
    '        explain C       ˵��

   ' strJson = Mid(strJson, 2)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Exsesvr_UpdatePatiBaseInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, blnShowErrMsg, , , , blnShowErrMsg) = False Then Exit Function
    '    output
    '        code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '        message C   1   "Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    Zl_Exsesvr_UpdatePatiBaseInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Exsesvr_Addcardfeeinfo(ByVal int����״̬ As Integer, _
    cllCardData As Collection, ByRef lng����ID_Out As Long, lngԤ��ID_Out As Long, _
    Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ӿ��Ѽ�Ԥ������
    '���:int����״̬-����״̬:0-������Ԥ����򿨷ѽɿ�;1-����Ϊδ��Ч��Ԥ������쳣�Ŀ���;2-����Ϊ���ʵ�;3-����Ϊ���۵�
    '     cllData: �����ݶ���
    '          |--billinfo:(����ϼ�,����Ա���,����Ա����,�Ǽ�ʱ��),Key="_billinfo"
    '          |--patinfo:(����ID,��ҳID,��������,�Ա�,����,�����,סԺ��,���ʽ���,�ѱ�,����),Key="_patinfo"
    '          |--cardinfo:������Ϣ(����,�����ID,������ʽ(0-����,1-����,2-����),��������,����id,ԭ����),key="_cardinfo"
    '          |--cardfeelists:key="_cardfeelists"
    '               |---cardfeelist:(���ѵ��ݺ�,���,�۸񸸺�,��������,�շ����,�շ�ϸĿid,������Ŀid,��׼����,�վݷ�Ŀ,Ӧ�ս��,ʵ�ս��,���˿���id,��������id,���˲���id,
    '                                 ִ�в���id,�Ƿ�����,���ձ���,������Ŀ��,ͳ����,ժҪ,��������,���������ID,������ʽ(0-����,1-����,2-����)) ,Key="_" & ���
    
    '          |--balanceinfo:(���㷽ʽ,�������,�����id,���㿨���,���ѿ�ID,֧������,������ˮ��,����˵��,������λ,�Ƿ����Ʊ��) Key="_balanceinfo"
    '          |--depositinfo:(Ԥ�����ݺ�,��Ʊ��,Ԥ�����,��ҳid,�ɿ����id,�ɿ���,�ɿλ,��λ������,ժҪ,����id,Ԥ������Ʊ��),Key="_depositinfo",��Ԥ��ʱ��������
    '          ���ϣ���ʽΪ:,��ʽ��array(����,ֵ)
    '          int����״̬=2-����Ϊ���ʵ�;3-����Ϊ���۵� �ģ�����"balanceinfo"��"depositinfo"�ڵ�
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '     blnDepositStartEinvoice-�Ƿ�����Ԥ������Ʊ��
    '     blnStartEinvoice-�Ƿ����ÿ��ѵ���Ʊ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllRow As Collection, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    Dim strJsonTemp  As String, strJsonFee As String, blnHaveDeposit As Boolean
    Dim j As Long
    
    If cllCardData Is Nothing Then Exit Function
    If cllCardData.Count = 0 Then Exit Function
    
    If blnShowErrMsg Then On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If Not blnShowErrMsg Then Err.Raise -1001, strErrMsg, strErrMsg: Exit Function
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    '1.��ȡ���ݽ�����Ϣ
    '    oper_fun    N   1   ����״̬:0-������Ԥ����򿨷ѽɿ�;1-����Ϊδ��Ч��Ԥ������쳣�Ŀ���;2-����Ϊ���ʵ�;3-����Ϊ���۵�
    '    blnc_total  N   1   ����ϼ�:Ԥ��+����
    '    operator_name   C   1   ����Ա����
    '    operator_code   C   1   ����Ա���
    '    create_time C   1   �Ǽ�ʱ����տ�ʱ��:yyyy-mm-dd hh:mi:ss
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_fun", int����״̬, Json_num)
    Set clldata = cllCardData("_billinfo")
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "����ϼ�"
            strJson = strJson & "," & GetJsonNodeString("blnc_total", Val(Trim(varTemp(1))), Json_num)
        Case "����Ա����", "����Ա"
            strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
        Case "����Ա���"
            strJson = strJson & "," & GetJsonNodeString("operator_code", Trim(varTemp(1)), Json_Text)
        Case "�Ǽ�ʱ��", "�տ�ʱ��"
            strJson = strJson & "," & GetJsonNodeString("create_time", Trim(varTemp(1)), Json_Text)
        Case Else
        End Select
    Next
    
    '2.ȡ������Ϣ
    Set clldata = cllCardData("_patinfo")
    
    '    pati_info   C       ������Ϣ
    '        pati_id N   1   ����ID
    '        pati_pageid N   1   ��ҳid
    '        pati_name   C   1   ��������
    '        pati_sex    C   1   �Ա�
    '        pati_age    C   1   ����
    '        outpatient_num  C   1   �����
    '        inpatient_num   C   1   סԺ��
    '        mdlpay_name    C   1   ���ʽ����
    '        fee_category    C   1   �ѱ�
    '        insurance_type  N   1   ����
    strJsonTemp = ""
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "����ID"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_id", Val(Trim(varTemp(1))), Json_num)
        Case "��ҳID"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_pageid", Val(Trim(varTemp(1))), Json_num)
        Case "����", "��������"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_name", Trim(varTemp(1)), Json_Text)
        Case "�Ա�"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_sex", Trim(varTemp(1)), Json_Text)
        Case "����"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_age", Trim(varTemp(1)), Json_Text)
        Case "�ѱ�"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_category", Trim(varTemp(1)), Json_Text)
        Case "����"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_type", Val(varTemp(1)), Json_num, True)
        Case "�����"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("outpatient_num", Trim(varTemp(1)), Json_Text)
        Case "סԺ��"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("inpatient_num", Trim(varTemp(1)), Json_Text)
        Case "���ʽ����"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("mdlpay_name", Trim(Trim(varTemp(1))), Json_Text)
 
        End Select
    Next
    If strJsonTemp = "" Then Exit Function
    
    
    strJsonTemp = Mid(strJsonTemp, 2)
    strJsonTemp = GetNodeString("pati_info") & ":{" & strJsonTemp & "}"
    strJson = strJson & "," & strJsonTemp
    
    '3.ȡ������Ϣ
    '    card_info   C       ҽ�ƿ���Ϣ
    '        cardno  C   1   ��������
    '        cardtype_id N   1   ���������ID
    '        send_mode   N   1   ������ʽ;0-����,1-����,2-����
    '        cardno_reusing  N   1   ��������:1-����;0-����������
    '        recv_id N   1   ����id:����Id
    '        cardno_old  C   1   ԭ������:����ʱ����Ҫ����ԭ����
    Set clldata = cllCardData("_cardinfo")
    strJsonTemp = ""
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "��������", "����"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
        Case "���������ID", "�����ID"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(Trim(varTemp(1))), Json_num)
        Case "������ʽ"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("send_mode", Val(Trim(varTemp(1))), Json_num)
        Case "��������", "�Ƿ񿨺�����", "�Ƿ񿨺��ظ�ʹ��"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno_reusing", Val(Trim(varTemp(1))), Json_num)
        Case "����ID"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("recv_id", Val(Trim(varTemp(1))), Json_num, True)
        Case "ԭ����"
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno_old", Trim(varTemp(1)), Json_Text)
        Case Else
        End Select
    Next
    If strJsonTemp = "" Then Exit Function
    
    
    strJsonTemp = Mid(strJsonTemp, 2)
    strJsonTemp = GetNodeString("card_info") & ":{" & strJsonTemp & "}"
    strJson = strJson & "," & strJsonTemp
    '2.ȡ������Ϣ
    Set clldata = cllCardData("_cardfeelists")
    '    cardfee_list[]  C   1   �����б�
    '      fee_no  C   1   ���ѵ��ݺ�
    '      serial_num  N   1   ���
    '      price_ftrnum    N   1   �۸񸸺�
    '      subde_ftrnum    N   1   ��������
    '      receipt_type    C   1   �շ����
    '      fitem_id    N   1   �շ�ϸĿid
    '      income_item_id  N   1   ������Ŀid
    '      price   N   1   ��׼����
    '      receipt_fee C   1   �վݷ�Ŀ
    '      fee_amrcvb  N   1   Ӧ�ս��
    '      fee_ampaib  N   1   ʵ�ս��
    '      pati_deptid N   1   ���˿���id
    '      bill_deptid N   1   ��������id
    '      pati_wardarea_id    N   1   ���˲���id
    '      exedept_id  N   1   ִ�в���id
    '      mrbkfee_sign N   1   �Ƿ�����:1-�ǲ�����;0-���ǲ�����
    '      insurance_code  C   1   ���ձ���
    '      insurance_type_id   N   1   ���մ���id
    '      insurance_sign  N   1   ������Ŀ��:1-�Ǳ�����Ŀ;0-���Ǳ�����Ŀ
    '      si_manp_money   N   1   ͳ����
    '      memo    C   1   ժҪ
 
     strJsonFee = ""
    For i = 1 To clldata.Count
        Set cllRow = clldata(i)
        strJsonTemp = ""
        For j = 1 To cllRow.Count
            varTemp = cllRow(j)
            Select Case UCase(varTemp(0))
            Case "���ѵ��ݺ�"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_no", Trim(varTemp(1)), Json_Text)
            Case "���"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("serial_num", Val(Trim(varTemp(1))), Json_num)
            Case "�۸񸸺�"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("price_ftrnum", Val(Trim(varTemp(1))), Json_num, True)
            Case "��������"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("subde_ftrnum", Val(Trim(varTemp(1))), Json_num, True)
            Case "�շ����"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("receipt_type", Trim(varTemp(1)), Json_Text)
            Case "�շ�ϸĿID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fitem_id", Val(Trim(varTemp(1))), Json_num)
            Case "������ĿID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("income_item_id", Val(Trim(varTemp(1))), Json_num)
            Case "��׼����"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("price", Val(Trim(varTemp(1))), Json_num)
            Case "�վݷ�Ŀ"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("receipt_fee", Trim(varTemp(1)), Json_Text)
            Case "Ӧ�ս��"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_amrcvb", Val(Trim(varTemp(1))), Json_num)
            Case "ʵ�ս��"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fee_ampaib", Val(Trim(varTemp(1))), Json_num)
            Case "���˿���ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_deptid", Val(Trim(varTemp(1))), Json_num)
            Case "��������ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("bill_deptid", Val(Trim(varTemp(1))), Json_num)
            Case "���˲���ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_wardarea_id", Val(Trim(varTemp(1))), Json_num)
            Case "ִ�в���ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("exedept_id", Val(Trim(varTemp(1))), Json_num)
            Case "�Ƿ�����"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("mrbkfee_sign", Val(Trim(varTemp(1))), Json_num)
            Case "���ձ���"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_code", Trim(varTemp(1)), Json_Text)
            Case "������Ŀ��"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("insurance_sign", Val(Trim(varTemp(1))), Json_num)
            Case "ͳ����"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("si_manp_money", Val(Trim(varTemp(1))), Json_num)
            Case "ժҪ"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("memo", Trim(varTemp(1)), Json_Text)
            Case "�Ӱ��־"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("overtime_flag", Val(Trim(varTemp(1))), Json_num)
            End Select
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonFee = strJsonFee & ",{" & strJsonTemp & "}"
        End If
    Next
    If strJsonFee = "" Then Exit Function
    strJsonFee = Mid(strJsonFee, 2)
    strJsonFee = GetNodeString("cardfee_list") & ":[" & strJsonFee & "]"
    strJson = strJson & "," & strJsonFee
    
    
    If int����״̬ < 2 Then '2-����Ϊ���ʵ�;3-����Ϊ���۵�
        '�Ǽ��ʻ򻮼�ʱ������Ч
        '4.ȡ������Ϣ
        'balance_info    C       ������Ϣ:Ŀǰֻ֧��һ�ֽ��㷽ʽ
        '    blnc_mode    C   1   ���㷽ʽ
        '    blnc_no C   1   �������
        '    cardtype_id N   1   �����id
        '    consumer_no N   1   ���㿨��ţ��������ѽӿ�Ŀ¼.���
        '    consume_card_id N   1   ���ѿ�ID
        '    cardno  C   1   ֧������
        '    swapno  C   1   ������ˮ��
        '    swapmemo    C   1   ����˵��
        '    cprtion_unit    C   1   ������λ
        '    start_einv  N   1   �Ƿ����õ���Ʊ��:1-����;0-������

        Set clldata = cllCardData("_balanceinfo")
        strJsonTemp = ""
        For i = 1 To clldata.Count
            varTemp = clldata(i)
            Select Case UCase(varTemp(0))
            Case "���㷽ʽ"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("blnc_mode", Trim(varTemp(1)), Json_Text)
            Case "�������"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("blnc_no", Trim(varTemp(1)), Json_Text)
            Case "�����ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(Trim(varTemp(1))), Json_num, True)
            Case "���㿨���"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("consumer_no", Val(Trim(varTemp(1))), Json_num, True)
            Case "֧������"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
            Case "������ˮ��"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swapno", Trim(varTemp(1)), Json_Text)
            Case "����˵��"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swapmemo", Trim(varTemp(1)), Json_Text)
            Case "������λ"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cprtion_unit", Trim(varTemp(1)), Json_Text)
            Case "�Ƿ����Ʊ��"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("start_einv", Val(varTemp(1)), Json_num, True)
            Case "���ѿ�ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("consume_card_id", Val(varTemp(1)), Json_num, True)
            Case Else
            End Select
        Next
        If strJsonTemp = "" Then Exit Function
        
        strJsonTemp = Mid(strJsonTemp, 2)
        strJsonTemp = GetNodeString("balance_info") & ":{" & strJsonTemp & "}"
        strJson = strJson & "," & strJsonTemp
        
        '5.��ȡԤ����
        
        Err = 0: On Error Resume Next
        Set clldata = cllCardData("_depositinfo")
        
        blnHaveDeposit = True
        If Err <> 0 Then
            Err = 0: On Error GoTo 0
            blnHaveDeposit = False
        End If
        
        On Error GoTo errHandle
        
        If blnHaveDeposit Then
            'deposit_info    C   1   Ԥ�����б�
            '    deposit_no  C   1   Ԥ�����ݺ�
            '    fact_no C   1   ��Ʊ��
            '    deposit_type    N       Ԥ�����:1-����;2-סԺ
            '    pati_pageid N   1   ��ҳid
            '    dept_id N   1   �ɿ����id
            '    money   N   1   �ɿ���
            '    emp_name    C   1   �ɿλ
            '    emp_bank_name   C   1   ��λ������
            '    emp_bank_actno  C   1   �������˺�
            '    memo    C   1   ժҪ
            '    recv_id N   1   ����id:����Id
            '    start_einv  N   1   �Ƿ����õ���Ʊ��:1-����;0-������

            strJsonTemp = ""
            For i = 1 To clldata.Count
                varTemp = clldata(i)
                Select Case UCase(varTemp(0))
                
                Case "Ԥ�����ݺ�"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("deposit_no", Trim(varTemp(1)), Json_Text)
                Case "��Ʊ��"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("fact_no", Trim(varTemp(1)), Json_Text)
                Case "Ԥ�����"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("deposit_type", Val(Trim(varTemp(1))), Json_num, True)
                Case "��ҳID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_pageid", Val(Trim(varTemp(1))), Json_num, True)
                Case "�ɿ����ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("dept_id", Val(Trim(varTemp(1))), Json_num, True)
                Case "�ɿ���"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("money", Val(Trim(varTemp(1))), Json_num, True)
                Case "�ɿλ"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("emp_bank_name", Trim(varTemp(1)), Json_Text)
                Case "��λ������"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("emp_name", Trim(varTemp(1)), Json_Text)
                Case "ժҪ"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("memo", Trim(varTemp(1)), Json_Text)
                Case "Ԥ������Ʊ��"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("start_einv", Val(varTemp(1)), Json_num, True)
                Case "����ID"
                    strJsonTemp = strJsonTemp & "," & GetJsonNodeString("recv_id", Val(Trim(varTemp(1))), Json_num, True)
                End Select
            Next
            
            If strJsonTemp = "" Then Exit Function
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonTemp = "" & GetNodeString("deposit_info") & ":{" & strJsonTemp & "}"
            strJson = strJson & "," & strJsonTemp
        End If
    End If

    If strJson = "" Then Exit Function
     
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    
    strServiceName = "Zl_Exsesvr_Addcardfeeinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then
        Exit Function
    End If
    '    output
    '       code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '       message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '       deposit_id  N   1   Ԥ��ID
    '       balance_id  N   1   ����ID
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        mstrErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If mstrErrMsg = "" Then
            mstrErrMsg = "���ô���ʧ�ܣ����飡"
        End If
        If blnShowErrMsg Then MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    

    lngԤ��ID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.deposit_id")))
    lng����ID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.balance_id")))
    Zl_Exsesvr_Addcardfeeinfo = True
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



Public Function Zl_Exsesvr_CheckCardnoIsUsed(ByVal lng�����ID As Long, ByVal strCardNo As String, ByRef bln�Ƿ����_Out As Boolean, ByRef lng����ID_Out As Long, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ�������Ƿ���Ʊ��ʹ����ϸ�д��ڣ�����ʱ����������ID
    '���:lng�����ID- ����
    '     strCardNo-����
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    
    '����:bln�Ƿ����_Out-���Ŵ��ڣ����� true,���򷵻�False
    '     lng����ID_Out-���ڿ���ʱ������ԭ����ID
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    bln�Ƿ����_Out = False
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'Zl_Exsesvr_Checkcardnoisused
    'input
    '    cardtype_id N   1   �����id
    '    cardno  C   1   ����


    strJson = strJson & "" & GetJsonNodeString("cardtype_id", lng�����ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardno", strCardNo, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Exsesvr_Checkcardnoisused"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    'output
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    isexsit N   1   �Ƿ����:1-����;0-������
    '    recv_id N   1   ����id

  
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    bln�Ƿ����_Out = objServiceCall.GetJsonNodeValue("output.isexsit") = 1
    lng����ID_Out = objServiceCall.GetJsonNodeValue("output.recv_id")
    Zl_Exsesvr_CheckCardnoIsUsed = True
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


Public Function Zl_Exsesvr_UpdCardFeeBlncInfo(ByVal int����״̬ As Integer, ByVal clsSendCardInfo As Collection, ByVal cllUpdateDate As Collection, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ŀ��ѽ�������
    '���:int����״̬:0-��ɽ���;1-�ӿڵ���ǰ����;2-�ӿڵ��ú�����
    '   clsSendCardInfo-������Ϣ(�����ID,�䶯����,����,ԭ����,IC����,����,��������,��ֹʹ��ʱ��,����,������,ժҪ,��������,����ID),��ʽ:array(����,ֵ),"_����"
    '   cllUpdateDate-�޸ĵĽ�������
    '         |--billinfo-������Ϣ,"_billinfo"
    '              |-Ԥ������,Ԥ��ID,�շѵ���,����ID,����Ա���,����Ա����,�տ�ʱ��,����ID,�Ƿ����Ʊ��,�Ƿ�Ԥ������Ʊ��)
    '         |--balanceinfo-������Ϣ,"_balanceinfo"
    '                |--(���㷽ʽ,�������,�����id,���㿨���,����,������ˮ��,����˵��,ժҪ,������λ)
    '                |--������Ϣ��,
    '                |-----������Ϣ:��������,��������
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, j As Long, m As Long, strServiceName  As String
    Dim clldata As Collection, cllTemp As Collection, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strSendCardJson As String, strBalanceJson As String, strOthersJson As String
 
    Dim strJsonTemp As String, cllOthers As Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    'Zl_Exsesvr_UpdCardFeeBlncInfo
    'input
    '   oper_fun    N   1   ����״̬:0-��ɽ���;1-�ӿڵ���ǰ����;2-�ӿڵ��ú�����
    '    pati_id N   1   ����id
    '    fee_no  C   1   ���õ��ţ�����Ҫ�����ķ��õ���
    '    balance_id  N       ����ID
    '    operator_name   C   1   ����Ա����
    '    operator_code   C   1   ����Ա���
    '    create_time C   1   ����ʱ��:yyyy-mm-dd hh:mi:ss
    '    fee_einvoice    N   1   ���ѻ������Ƿ����õ���Ʊ��:1-����;0-������
    '    sendcard_info ������Ϣ
    '        send_mode   N   1   ������ʽ;0-����,1-����,2-����
    '        cardtype_id C   1   �����id
    '        cardno  C   1   ����:���η��Ż�󶨻򲹿��Ŀ���
    '        recv_id N   1   ����id:Ʊ������ID(����)
    '        cardno_reusing  N   1   ��������:1-���������ظ�ʹ����;0-�������ظ�ʹ��
    '        cardno_old  C   1   ԭ������:����ʱ����Ҫ����ԭ����
    '    balance_info    C       ������Ϣ
    '        deposit_no  C       Ԥ������
    '        deposit_id  N       Ԥ��ID
    '        deposit_einvoice    N       Ԥ�����õ���Ʊ��:1-����;0-������
    '        pay_mode    C   1   ���㷽ʽ
    '        blnc_no C   1   �������
    '        cardtype_id N   1   �����id
    '        consumer_no N   1   ���㿨��ţ��������ѽӿ�Ŀ¼.���
    '        cardno  C   1   ����
    '        swapno  C   1   ������ˮ��
    '        swapmemo    C   1   ����˵��
    '        memo    C   1   ժҪ
    '        cprtion_unit    C   1   ������λ
    '        other_list[]    C   1   ����������Ϣ
    '            swap_name   C   1   ��������
    '            swap_note   C   1   ��������

    Set clldata = cllUpdateDate("_billinfo")
    strJson = ""
    strSendCardJson = ""
    
    strJson = strJson & "," & GetJsonNodeString("oper_fun", int����״̬, Json_num)
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "Ԥ������"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("deposit_no", Trim(varTemp(1)), Json_Text)
        Case "Ԥ��ID"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("deposit_id", Val(varTemp(1)), Json_num, True)
        Case "����ID"
            strJson = strJson & "," & GetJsonNodeString("pati_id", Val(varTemp(1)), Json_num, True)
        Case "�շѵ���"
            strJson = strJson & "," & GetJsonNodeString("fee_no", Trim(varTemp(1)), Json_Text)
        Case "����ID"
            strJson = strJson & "," & GetJsonNodeString("balance_id", Val(varTemp(1)), Json_num, True)
        Case "����Ա���"
            strJson = strJson & "," & GetJsonNodeString("operator_code", Trim(varTemp(1)), Json_Text)
        Case "����Ա����"
            strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
        Case "�տ�ʱ��"
            strJson = strJson & "," & GetJsonNodeString("create_time", Trim(varTemp(1)), Json_Text)
        Case "�Ƿ����Ʊ��"
            strJson = strJson & "," & GetJsonNodeString("fee_einvoice", Val(varTemp(1)), Json_num, True)
        Case "�Ƿ�Ԥ������Ʊ��"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("deposit_einvoice", Val(varTemp(1)), Json_num, True)
        End Select
    Next
    
    '        send_mode   N   1   ������ʽ;0-����,1-����,2-������3-�˿�
    '        cardtype_id C   1   �����id
    '        cardno  C   1   ����:���η��Ż�󶨻򲹿��Ŀ���
    '        recv_id N   1   ����id:Ʊ������ID(����)
    '        cardno_reusing  N   1   ��������:1-���������ظ�ʹ����;0-�������ظ�ʹ��
    '        cardno_old  C   1   ԭ������:����ʱ����Ҫ����ԭ����
    If Not clsSendCardInfo Is Nothing Then
        strSendCardJson = ""
        For i = 1 To clsSendCardInfo.Count
            varTemp = clsSendCardInfo(i)
            Select Case UCase(varTemp(0))
            Case "�����ID"
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num, True)
            Case "�䶯����"
                '1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ);7-��ֹʱ�����
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("send_mode", decode(Val(varTemp(1)), 1, 0, 11, 0, 2, 2, 3, 1, 3), Json_num, True)
            Case "����"
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
            Case "ԭ����"
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("cardno_old", Trim(varTemp(1)), Json_Text)
            Case "IC����"
            Case "����"
            Case "��������"
            Case "��ֹʹ��ʱ��"
            Case "����"
            Case "������"
            Case "ժҪ"
            Case "��������"
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("cardno_reusing", Val(varTemp(1)), Json_num, True)
            Case "����ID"
                strSendCardJson = strSendCardJson & "," & GetJsonNodeString("recv_id", Val(varTemp(1)), Json_num, True)
            End Select
        Next
        If strSendCardJson = "" Then Exit Function
        strSendCardJson = Mid(strSendCardJson, 2)
        strJson = strJson & "," & GetNodeString("sendcard_info") & ":{" & strSendCardJson & "}"
    End If
    '������Ϣ
    '        deposit_no  C       Ԥ������
    '        deposit_id  N       Ԥ��ID
    '        pay_mode    C   1   ���㷽ʽ
    '        blnc_no C   1   �������
    '        cardtype_id N   1   �����id
    '        consumer_no N   1   ���㿨��ţ��������ѽӿ�Ŀ¼.���
    '        cardno  C   1   ����
    '        swapno  C   1   ������ˮ��
    '        swapmemo    C   1   ����˵��
    '        memo    C   1   ժҪ
    '        statu   N   1   0-��ɽ���;1-�ӿڵ���ǰ,2-�ӿڵ��óɹ�
    '        cprtion_unit    C   1   ������λ
    '        other_list[]    C   1   ����������Ϣ
    '            swap_name   C   1   ��������
    '            swap_note   C   1   ��������

    Set clldata = cllUpdateDate("_balanceinfo")
    For i = 1 To clldata.Count
        varTemp = clldata(i)
        Select Case UCase(varTemp(0))
        Case "���㷽ʽ"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("pay_mode", Trim(varTemp(1)), Json_Text)
        Case "�������"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("blnc_no", Trim(varTemp(1)), Json_Text)
        Case "�����ID"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num)
        Case "���㿨���"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("consumer_no", Val(varTemp(1)), Json_num)
        Case "����"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
        Case "������ˮ��"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("swapno", Trim(varTemp(1)), Json_Text)
        Case "����˵��"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("swapmemo", Trim(varTemp(1)), Json_Text)
        Case "ժҪ"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("memo", Trim(varTemp(1)), Json_Text)
        Case "������λ"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("cprtion_unit", Trim(varTemp(1)), Json_Text)
        Case "У�Ա�־"
            strBalanceJson = strBalanceJson & "," & GetJsonNodeString("statu", Val(varTemp(1)), Json_num)
        Case UCase("������Ϣ��") '������չ������Ϣ
            Set cllOthers = varTemp(1)
            strOthersJson = ""
            For j = 1 To cllOthers.Count
                 Set cllTemp = cllOthers(j)
                 strJsonTemp = ""
                 For m = 1 To cllTemp.Count
                    varTemp = cllTemp(m)
          
                    Select Case UCase(varTemp(0))
                    Case "��������"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swap_name", Trim(varTemp(1)), Json_Text)
                    Case "��������"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swap_note", Trim(varTemp(1)), Json_Text)
                    End Select
                 Next
                 If strJsonTemp <> "" Then
                    strJsonTemp = Mid(strJsonTemp, 2)
                    strOthersJson = strOthersJson & ",{" & strJsonTemp & "}"
                 End If
            Next
        End Select
    Next
    If strOthersJson <> "" Then
        strOthersJson = Mid(strOthersJson, 2)
       strBalanceJson = strBalanceJson & "," & GetNodeString("other_list") & ":[" & strOthersJson & "]"
    End If
     
    If strBalanceJson <> "" Then
        strBalanceJson = Mid(strBalanceJson, 2)
        strJson = strJson & "," & GetNodeString("balance_info") & ":{" & strBalanceJson & "}"
    End If
    If strJson = "" Then Exit Function
    strJson = Mid(strJson, 2)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    
    strServiceName = "Zl_Exsesvr_UpdCardFeeBlncInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    'output
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    Zl_Exsesvr_UpdCardFeeBlncInfo = True
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


Public Function zl_ExseSvr_GetNextNo(ByVal int��� As Integer, ByRef strNo_Out As String, Optional ByVal lng����ID As Long, _
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
    strServiceName = "zl_ExseSvr_GetNextNo"
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
    zl_ExseSvr_GetNextNo = True
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
 
Public Function zl_ExseSvr_GetBillStatuByNo(ByVal strNO As String, ByVal int�������� As Integer, _
    ByRef int�շ�״̬_Out As Integer, ByRef int�쳣״̬_Out As Integer, Optional ByRef int���ʱ�־_Out As Integer, _
    Optional ByRef intԤ�����ѱ�־_Out As Integer, Optional blnShowErrMsg As Boolean = True, Optional ByRef strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݺŻ�õ��ݵ��շѡ��쳣�����ʵ�״̬
    '���:strNO-���ݺ�
    '     int��������-��������:1-�շѵ�;2-���ʵ�;3-�Զ����ʵ�;4-�Һŵ�;5-���￨;6-Ԥ����
    '����:int�շ�״̬_Out:0-δ�շѻ򻮼�;1-���շѻ��Ѽ���;2-��ȫ�˻�ȫ����;3-�����˷ѻ򲿷�����
    '    int�쳣״̬_Out-:0-��������;1-�տ���쳣;2-�˿���쳣
    '    int���ʱ�־_Out:��Լ��ʵ���Ч;0-δ����;1-�Ѿ�����
    '    intԤ�����ѱ�־_Out:1-����������;0-δ��������
    '     strErrMsg_Out-������Ϣ
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
    '    fee_no  C   1   ���ݺ�
    '    bill_prop   N   1   ��¼����:1-�շѵ�;2-���ʵ�;3-�Զ����ʵ�;4-�Һŵ�;5-���￨;6-Ԥ����
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("fee_no", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("bill_prop", int��������, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetBillStatuByNo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    '    output
    '        code    N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '        message C   1   "Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '        statu   N   1   �շ�״̬:0-δ�շѻ򻮼�;1-���շѻ��Ѽ���;2-��ȫ�˻�ȫ����;3-�����˷ѻ򲿷�����
    '        err_sign    N   1   �쳣��־:0-��������;1-�տ���쳣;2-�˿���쳣
    '        blnc_sign   N   1   ���ʱ�־:��Լ��ʵ���Ч;0-δ����;1-�Ѿ�����
    '        consumeed_sign  N   1   Ԥ�����ѱ�־:1-����������;0-δ��������


    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡ���ݺţ����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    int�շ�״̬_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.statu")))
    int�쳣״̬_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.err_sign")))
    int���ʱ�־_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.blnc_sign")))
    intԤ�����ѱ�־_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.consumeed_sign")))
    zl_ExseSvr_GetBillStatuByNo = True
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
 
Public Function zl_ExseSvr_UpdateCardFeeInfo(ByVal int������־ As Integer, ByVal strCardFeeNo As String, ByVal cllSendCard As Collection, _
    ByVal str����Ա���� As String, ByVal str����Ա��� As String, ByVal str�Ǽ�ʱ�� As String, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ŀ��ѽ�������
    '���:int������־-0-ֻ�޸�סԺ���ü�¼;1-�޸ķ��ü�¼��Ʊ��ʹ����ϸ
    '     cllSendCard -���ط�������(�����ID,�䶯����,����,ԭ����,IC����,����,��������,��ֹʹ��ʱ��,����,������,ժҪ,��������,����ID),��ʽ:array(����,ֵ),"_����"
    '     str�Ǽ�ʱ��-��ʽ:yyyy-mm-dd hh24:mi:ss
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strJsonTemp As String, strServiceName    As String, varTemp As Variant
    Dim objServiceCall As Object, i As Long
    Dim intReturn As Integer
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Err = 0: On Error GoTo errHandle:
    'zl_ExseSvr_UpdateCardFeeInfo
    ' input
    ' input
    '    oper_fun    N   1   ������־:0-ֻ�޸�סԺ���ü�¼;1-�޸ķ��ü�¼��Ʊ��ʹ����ϸ
    '    fee_no  C   1   ���õ��ţ�����Ҫ�����ķ��õ���
    '    operator_name   C   1   ����Ա����
    '    operator_code   C   1   ����Ա���
    '    create_time C   1   �Ǽ�ʱ��:yyyy-mm-dd hh:mi:ss
    '    sendcard_info ������Ϣ
    '        send_mode   N   1   ������ʽ;0-����,1-����,2-����;3-�˿�
    '        cardtype_id C   1   �����id
    '        cardno  C   1   ����:���η��Ż�󶨻򲹿��Ŀ���
    '        recv_id N   1   ����id:Ʊ������ID(����)
    '        cardno_reusing  N   1   ��������:1-���������ظ�ʹ����;0-�������ظ�ʹ��
    '        cardno_old  C   1   ԭ������:����ʱ����Ҫ����ԭ����

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_fun", int������־, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_no", strCardFeeNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_name", str����Ա����, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_code", str����Ա���, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("create_time", str�Ǽ�ʱ��, Json_Text)
    strJsonTemp = ""
    If Not cllSendCard Is Nothing Then
        For i = 1 To cllSendCard.Count
            varTemp = cllSendCard(i)
            Select Case UCase(varTemp(0))
            Case "�����ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num, True)
            Case "�䶯����"
                ''0-��ȷ����1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ),7-��ֹʱ�����
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("send_mode", decode(Val(varTemp(1)), 1, 0, 3, 1, 2, 2, 4, 3, 0), Json_num, True)
            Case "����"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
            Case "ԭ����"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno_old", Trim(varTemp(1)), Json_Text)
            Case "IC����"
            Case "����"
            Case "��������"
            Case "��ֹʹ��ʱ��"
            Case "����"
            Case "������"
            Case "ժҪ"
            Case "��������"
            Case "����ID"
                strJsonTemp = strJsonTemp & "," & GetJsonNodeString("recv_id", Val(varTemp(1)), Json_num, True)
            End Select
        Next
        If strJsonTemp <> "" Then
            strJsonTemp = Mid(strJsonTemp, 2)
            strJsonTemp = "," & GetNodeString("sendcard_info") & ":{" & strJsonTemp & "}"
        End If
    End If
    strJson = strJson & strJsonTemp
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_UpdateCardFeeInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, blnShowErrMsg, , , , Not blnShowErrMsg) = False Then Exit Function
    'output
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
    zl_ExseSvr_UpdateCardFeeInfo = True
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

Public Function Zl_Exsesvr_CardfeeIsBalance(ByVal strCardFeeNo As String, ByVal intQueryType As Boolean, _
    ByRef str���ʵ���_Out As String, ByRef bln����_Out As Boolean, _
    Optional blnShowErrMsg As Boolean = True, Optional ByRef strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϿ������Ƿ��Ѿ�����
    '���:strCardFeeNo-���ѵ��ݺ�
    '     intQueryType-0-��ȡ����;1-������;2-���Ѽ�������
    '����:strErrMsg_Out-������Ϣ
    '     str���ʵ���_Out-���ؽ��ʵ��ݺ�
    '     bln����_Out-�Ƿ��Ѿ�����:�Ѿ����ʷ���true,���򷵻�False
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
    '   cardfee_no  C   1   ���Ѷ�Ӧ�ķ��õ��ݺ�
    '   strCardFeeNo  N   1   ��ȡ���ѱ�־:0-��ȡ����,1-������;2-���ѻ�������
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("cardfee_no", strCardFeeNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("strCardFeeNo", intQueryType, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_CardfeeIsBalance"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    'code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    'message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    'isbalanced  N   1   �Ƿ��Ѿ�����:1-�ѽ����;0-δ����
    'blnc_no C   1   ���ʵ��ݺ�

    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "������Ч��ȡ���ѵ��ݣ����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln����_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.isbalanced"))) = 1
    str���ʵ���_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.blnc_no"))) = 1
    Zl_Exsesvr_CardfeeIsBalance = True
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


Public Function zl_ExseSvr_DelCardFeeCheck(ByVal int�˷Ѳ��� As Integer, ByVal strCardFeeNo As String, ByVal bln�쳣���� As Boolean, _
    ByVal cllBalanceInf As Collection, Optional ByVal strDepositNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˿����˲�����ǰ���ݼ��
    '���:int�˷Ѳ���-0-���˿���;1-���˲�����;2-�����Ѽ�����
    '   cllBalanceInf-(�˿���,���㷽ʽ,�����ID,���㿨���,�Ƿ�ȫ��),array(����,ֵ ),"_����"
    '����:strErrMsg-������Ϣ
    '     str���ʵ���_Out-���ؽ��ʵ��ݺ�
    '     bln����_Out-�Ƿ��Ѿ�����:�Ѿ����ʷ���true,���򷵻�False
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Collection, cllTemp As Collection, varTemp As Variant
    Dim strErrMsg As String, strJsonBalance As String
    Dim objServiceCall As Object
    Dim intReturn As Integer
  
    If cllBalanceInf Is Nothing Then
        strErrMsg = "δ�����Ҫ�ļ�����������ܽ���" & decode(int�˷Ѳ���, 1, "�˲�����", "�˿�") & "������!"
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If GetServiceCall(objServiceCall) = False Then
        strErrMsg = "���ӷ��������ʧ�ܣ��޷����" & decode(int�˷Ѳ���, 1, "�˲�����", "�˿�") & "����Ч��,����ϵͳ����Ա��ϵ!"
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'input
    '    cardfee_no  C   1   ���ѵ���
    '    deposit_no  C   1   Ԥ�����ݺ�
    '    reretruned  N   1   �Ƿ��쳣����:1-���쳣����;0-���쳣����
    '    delfee_sign N   1   �˷ѱ�־��0-���˿���;1-���˲�����;2-�����Ѽ�����
    '    balance_info    C       �˿ʽ
    '        delmoney    N   1   �����˿���
    '        pay_mode    C   1   ���㷽ʽ
    '        cardtype_id N   1   �����id
    '        consumer_no N   1   ���㿨��ţ��������ѽӿ�Ŀ¼.���
    '        must_allreturn  N   1   �Ƿ�ȫ��:1-����ȫ��;0-��������
    
    
    strJson = "": strJsonBalance = ""
    If Not cllBalanceInf Is Nothing Then
        
        For i = 1 To cllBalanceInf.Count
            varTemp = cllBalanceInf(i)
            Select Case UCase(varTemp(0))
            Case "�˿���"
                 strJsonBalance = strJsonBalance & "," & GetJsonNodeString("delmoney", Val(varTemp(1)), Json_num)
            Case "���㷽ʽ"
                strJsonBalance = strJsonBalance & "," & GetJsonNodeString("deposit_no", Trim(varTemp(1)), Json_Text)
            Case "�����ID"
                 strJsonBalance = strJsonBalance & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num)
            Case "���㿨���"
                 strJsonBalance = strJsonBalance & "," & GetJsonNodeString("consumer_no", Val(varTemp(1)), Json_num)
            Case "�Ƿ�ȫ��"
                 strJsonBalance = strJsonBalance & "," & GetJsonNodeString("must_allreturn", Val(varTemp(1)), Json_num)
            End Select
        Next
    End If
    

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("delfee_sign", int�˷Ѳ���, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardfee_no", strCardFeeNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("reretruned", IIf(bln�쳣����, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("deposit_no", strDepositNo, Json_Text)
    If strJsonBalance <> "" Then
        strJsonBalance = Mid(strJsonBalance, 2)
        strJsonBalance = "," & GetNodeString("balance_info") & ":{" & strJsonBalance & "}"
    End If
    strJson = strJson & strJsonBalance
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_DelCardFeeCheck"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False, , , , False) = False Then Exit Function
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    'output
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    tip_list[]  C   1   ��ʾ�б�:��Ҫ�ǿ��ܴ��ڶ����ʾѯ�ʷ�ʽ���������б�,��ֹʱ������һ����Ϣ
    '        tip_mode    C   1   ���Ʒ�ʽ:1-��ʾѯ��;2-��ֹ
    '        tip_message C   1   ��ʾ��Ϣ

    If intReturn = 0 Then
        strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg = "" Then
            strErrMsg = "��������Ԥ֪�Ĵ���"
        End If
        MsgBox strErrMsg & ",���ܽ���" & decode(int�˷Ѳ���, 1, "�˲�����", "�˿�") & "����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set clldata = objServiceCall.GetJsonListValue("output.tip_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            strErrMsg = Nvl(cllTemp("_tip_message"))
              
              Select Case Val(Nvl(cllTemp("_tip_mode")))
              Case 1  '��ʾ
                    If MsgBox(strErrMsg & ",���Ƿ����Ҫ" & decode(int�˷Ѳ���, 1, "�˲�����", "�˿�") & "����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                      Exit Function
                    End If
              Case 2  '��ֹ
                  MsgBox strErrMsg & ",���ܽ���" & decode(int�˷Ѳ���, 1, "�˲�����", "�˿�") & "����!", vbInformation + vbOKOnly, gstrSysName
                  Exit Function
              End Select
        Next
    End If
    zl_ExseSvr_DelCardFeeCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function Zl_Exsesvr_Delcardfeeinfo(ByVal int����״̬ As Integer, cllDelFeeData As Collection, ByRef lng����ID_Out As Long, lngԤ��ID_Out As Long, _
    Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˿��ѡ������Ѽ�Ԥ������
    '���:int����״̬-����״̬:0-������Ԥ����򿨷ѵ��˿��¼;1-����Ϊ�쳣���˿��¼;2-�����쳣����
    '     cllDelFeeData-�˷�����
    '        |-(���ѵ���,Ԥ������,�Ƿ��˿���,�Ƿ��˲�����,����Ա����,����Ա���,�˷�ʱ��,������Ϣ) array(����,ֵ) ,"_����)
    '        |-������Ϣ:(�˿���,���㷽ʽ,�������,�����id,���㿨���,֧������,������ˮ��,����˵��,������λ,��������ID,����ժҪ) Key="_������Ϣ"
    '     blnShowErrMsg-�Ƿ���ʾ������Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Collection, varTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
    Dim strJsonTemp  As String
    Dim j As Long
    
    If cllDelFeeData Is Nothing Then Exit Function
    If cllDelFeeData.Count = 0 Then Exit Function
    
    If blnShowErrMsg Then On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If Not blnShowErrMsg Then Err.Raise -1001, strErrMsg, strErrMsg: Exit Function
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    '1.��ȡ���ݽ�����Ϣ
    'oper_fun    N   1   ����״̬:0-������Ԥ����򿨷ѵ��˿��¼;1-����Ϊ�쳣���˿��¼;2-�����쳣����
    'cardfee_no  C   1   ���ѵ���
    'deposit_no  C   1   Ԥ������
    'cardfee_sign    N   1   �Ƿ��˿���:1-���˿���;0-���˿���
    'mrbkfee_sign N   1   �Ƿ��˲�����:1-�˲�����;0-���˲�����
    'operator_name   C   1   ����Ա����
    'operator_code   C   1   ����Ա���
    'del_time    C   1   �˷�ʱ��:yyyy-mm-dd hh:mi:ss
    'balance_info    C       ֻ����һ������
    '    moeny   N   1   �˿���
    '    blnc_mode    C   1   ���㷽ʽ
    '    blnc_no C   1   �������
    '    memo    C   1   ժҪ
    '    cardtype_id N   1   �����id
    '    consumer_no N   1   ���㿨��ţ��������ѽӿ�Ŀ¼.���
    '    cardno  C   1   ����
    '    swapno  C   1   ������ˮ��
    '    swapmemo    C   1   ����˵��
    '    cprtion_unit    C   1   ������λ
    '    relation_id N   1   ��������ID

    strJson = "": strJsonTemp = ""
    strJson = strJson & "" & GetJsonNodeString("oper_fun", int����״̬, Json_num)
    For i = 1 To cllDelFeeData.Count
        varTemp = cllDelFeeData(i)
        Select Case UCase(varTemp(0))
        Case "���ѵ���"
            strJson = strJson & "," & GetJsonNodeString("cardfee_no", Trim(varTemp(1)), Json_Text)
        Case "Ԥ������"
            strJson = strJson & "," & GetJsonNodeString("deposit_no", Trim(varTemp(1)), Json_Text)
        Case "�Ƿ��˿���"
            strJson = strJson & "," & GetJsonNodeString("cardfee_sign", Val(varTemp(1)), Json_num)
        Case "�Ƿ��˲�����"
            strJson = strJson & "," & GetJsonNodeString("mrbkfee_sign", Val(varTemp(1)), Json_num)
        Case "����Ա����"
            strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
        Case "����Ա���"
            strJson = strJson & "," & GetJsonNodeString("operator_code", Trim(varTemp(1)), Json_Text)
        Case "�˷�ʱ��"
            strJson = strJson & "," & GetJsonNodeString("del_time", Trim(varTemp(1)), Json_Text)
        Case "������Ϣ"
            Set cllTemp = varTemp(1)
            If Not cllTemp Is Nothing Then
                For j = 1 To cllTemp.Count
                     varTemp = cllTemp(j)
                    Select Case UCase(varTemp(0))
                    Case "�˿���"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("moeny", Val(varTemp(1)), Json_num)
                    Case "���㷽ʽ"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("blnc_mode", Trim(varTemp(1)), Json_Text)
                    Case "�������"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("blnc_no", Trim(varTemp(1)), Json_Text)
                    Case "����ժҪ"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("memo", Trim(varTemp(1)), Json_Text)
                    Case "�����ID"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardtype_id", Val(varTemp(1)), Json_num)
                    Case "���㿨���"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("consumer_no", Val(varTemp(1)), Json_num)
                    Case "֧������"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cardno", Trim(varTemp(1)), Json_Text)
                    Case "������ˮ��"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swapno", Trim(varTemp(1)), Json_Text)
                    Case "����˵��"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("swapmemo", Trim(varTemp(1)), Json_Text)
                    Case "������λ"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("cprtion_unit", Trim(varTemp(1)), Json_Text)
                    Case "��������ID"
                        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("relation_id", Val(varTemp(1)), Json_num)
                    End Select
                Next
            
            End If
        End Select
    Next
    If strJsonTemp <> "" Then
        strJsonTemp = Mid(strJsonTemp, 2)
        strJsonTemp = "," & GetNodeString("balance_info") & ":{" & strJsonTemp & "}"
    End If
    strJson = strJson & strJsonTemp
    If strJson = "" Then Exit Function
     
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    
    strServiceName = "Zl_Exsesvr_Delcardfeeinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then
        Exit Function
    End If
    '    output
    '       code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '       message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '       deposit_id  N   1   Ԥ��ID
    '       balance_id  N   1   ����ID
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        mstrErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If mstrErrMsg = "" Then
            mstrErrMsg = "���ô���ʧ�ܣ����飡"
        End If
        If blnShowErrMsg Then MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lngԤ��ID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.deposit_id")))
    lng����ID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.balance_id")))
    Zl_Exsesvr_Delcardfeeinfo = True
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

Public Function zlGetReceiveInvoiceRecFromCollect(ByVal cllInvoice As Collection, ByRef rsInvoice_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��񷵻ص�Ʊ����Ϣ����,����¼����ʽ������Ϣ
    '���:cllCardFee-��ǰ����
    '
    '����:rsCardFee_Out-���صĿ����ü���
    '     objBalanceItems_out-������Ϣ�б���Ҫ�ǿ��ܴ��ڼ��ʣ���Ҫ��objBalanceItems_out
    '     dblMoney_Out:ʵ�ս��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-07 15:22:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection
    Dim i As Long
    
    On Error GoTo errHandle
      
    Set rsInvoice_Out = New ADODB.Recordset
    With rsInvoice_Out
        If .State = adStateOpen Then .Close
        
        .Fields.Append "ID", adBigInt, , adFldIsNullable
        .Fields.Append "ʹ��������", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "ʹ�����ID", adBigInt, , adFldIsNullable
        .Fields.Append "ʹ�����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "�Ǽ�ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ʼ����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��ֹ����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "ʣ������", adDouble, , adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    If cllInvoice Is Nothing Then Exit Function
    If cllInvoice.Count = 0 Then Exit Function
    
    '    recv_id N   1   ����ID
    '    use_mode    N   1   ʹ�÷�ʽ:1-���ã���Ʊ�ݽ����������Լ�ʹ�ã�2-���ã���Ʊ���ɶ����Ա��ͬʹ��
    '    use_type    C   1   Ʊ��ʹ�����:1,4: Ʊ��ʹ�����.����;2Ԥ��:1-����Ԥ��;2-סԺԤ��;5:�洢����ҽ�ƿ����.ID
    '    prefix_text C   1   ǰ׺�ı�
    '    start_no    C   1   ��ʼ����
    '    end_no  C   1   ��ֹ����
    '    inv_no_cur  C   1   ��ǰ����
    '    surplus_num C   1   ʣ������
    '    create_time C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
    '    use_time    C   1   ʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
    '    recvtr  C   1   ������
    '    use_typecode    C   1   ʹ��������
    '    use_typeid  N   1   ʹ�����id
    For i = 1 To cllInvoice.Count
        Set cllTemp = cllInvoice(i)
        With rsInvoice_Out
            .AddNew
            !id = Val(Nvl(cllTemp("_recv_id")))
            !ʹ�������� = Nvl(cllTemp("_use_typecode"))
            !ʹ�����ID = Val(Nvl(cllTemp("_use_typeid")))
            !ʹ����� = Nvl(cllTemp("_use_type"))
            !������ = Nvl(cllTemp("_recvtr"))
            !�Ǽ�ʱ�� = Nvl(cllTemp("_create_time"))
            !��ʼ���� = Nvl(cllTemp("_start_no"))
            !��ֹ���� = Nvl(cllTemp("_end_no"))
            !ʣ������ = Val(Nvl(cllTemp("_surplus_num")))
            .Update
        End With
    Next
    If rsInvoice_Out.RecordCount <> 0 Then rsInvoice_Out.MoveFirst
    zlGetReceiveInvoiceRecFromCollect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
Public Function zl_ExseSvr_GetRelatedTransInfo(ByVal str��������Ids As String, ByRef rsSwap_Out As ADODB.Recordset, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݹ�������ID,��ȡ������Ϣ
    '���:str��������Ids-����ö��ŷ���
 
    '����: rsSwap_Out-���ر��ν�����Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    If blnShowErrMsg Then On Error GoTo errHandle:
    
    Set rsSwap_Out = New ADODB.Recordset
    With rsSwap_Out
        If .State = adStateOpen Then .Close
        .Fields.Append "��������ID", adBigInt, , adFldIsNullable
        .Fields.Append "�����ID", adBigInt, , adFldIsNullable
        .Fields.Append "���㷽ʽ", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "������ˮ��", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "����˵��", adLongVarChar, 4000, adFldIsNullable
        .Fields.Append "ԭʼ���", adDouble, , adFldIsNullable
        .Fields.Append "���˽��", adDouble, , adFldIsNullable
        .Fields.Append "ʣ��δ�˽��", adDouble, , adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    'zl_ExseSvr_GetRelatedTransInfo
    '   input
    '        related_ids    C   1   ��������ID:����ö��ŷ���
 
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("related_ids", str��������Ids, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetRelatedTransInfo"
    
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    'output
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    swap_list[]   C   1   ������Ϣ�б�
    '       related_id N   1   ��������ID
    '       cardtype_id N   1   �����ID
    '       blnc_mode   C   1   ���㷽ʽ
    '       swapno  C   1   ������ˮ��
    '       swapmemo    C   1   ����˵��
    '       original_money  N   1   ԭʼ���
    '       return_money    N   1   ���˽��

 

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "δ�ҵ����������Ľ�����Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set clldata = objServiceCall.GetJsonListValue("output.swap_list")
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then Exit Function
 
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsSwap_Out
            .AddNew
            !��������ID = Val(Nvl(cllTemp("_related_id")))
            !�����ID = Val(Nvl(cllTemp("_cardtype_id")))
            !���㷽ʽ = Nvl(cllTemp("_blnc_mode"))
            !������ˮ�� = Nvl(cllTemp("_swapno"))
            !����˵�� = Nvl(cllTemp("_swapmemo"))
            !ԭʼ��� = Val(Nvl(cllTemp("_original_money")))
            !���˽�� = Val(Nvl(cllTemp("_return_money")))
            !ʣ��δ�˽�� = RoundEx(Val(Nvl(!ԭʼ���)) - Val(Nvl(!���˽��)), 5)
            .Update
        End With
    Next
    zl_ExseSvr_GetRelatedTransInfo = True
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
 
Public Function Zl_Exsesvr_Geteinvoicesinfo(ByVal frmMain As Object, ByVal strNO As String, ByRef cllPati_Out As Collection, _
    ByRef lngEInvoiceID_out As Long, ByRef strInvoiceNO_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ʊ����Ϣ
    '���:
    '����:cllPati_Out-���ز��˼�
    '     lngEInvoiceID_out-���ص���Ʊ��ID
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object, intReturn As Integer
    Dim blnShowErrMsg As Boolean, strErrmsg_Out As String
    
    blnShowErrMsg = True
    
    Dim cllTemp As Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '--    query_type  N    ��ѯ��Χ:0-����;1-ֻ��ѯ��Ч�ĵ���Ʊ��;2-��ѯԭʼ����Ʊ����Ϣ
    '--    occasion  N  1  ���㳡��:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨,6-����ҽ������
    '--    fee_nos  C    query_type=2ʱ��Ч:���ݺ�:���㳡��=2ʱ��ΪԤ��NO, ����idδ���룬�ýڵ�ش�
    '--    balance_id  N    ����ID�����㳡��=2ʱ��ΪԤ��ID
    '--    read_oldbill  N  1  �Ƿ�ֻ��ȡԭʼ���ݵĵ���Ʊ��:1-��;2-��
    '--    invoice_type  N    Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", 2, Json_num)
    strJson = strJson & "," & GetJsonNodeString("occasion", 5, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_nos", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("invoice_type", 5, Json_num)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Geteinvoicesinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, True) = False Then Exit Function
    '    output
    '--    pati_info  C    ������Ϣ
    '--      pati_id  N  1  ����ID
    '--      pati_pageid  N    ��ҳID
    '--      pati_name  C  1  ����
    '--      pati_sex  C  1  �Ա�
    '--      pati_age  C  1  ����
    '--      outpatient_num  C  1  �����
    '--      inpatient_num  C  1  סԺ��
    '--    einvoice_info  C    ����Ʊ����Ϣ:query_type=2ʱ����
    '--      einv_id  N  1  ����Ʊ��ID
    '--      paper_nos  C  1  δ���յ�ֽ�ʷ�Ʊ��Ϣ,����ö��ŷ���
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "���ܻ�ȡ���ݺţ����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllTemp = objServiceCall.GetJsonListValue("output.data.pati_info")
    Set cllPati_Out = New Collection
    cllPati_Out.Add Val(Nvl(cllTemp("_pati_id"))), "_����ID"
    cllPati_Out.Add Val(Nvl(cllTemp("_pati_pageid"))), "_��ҳID"
    
    cllPati_Out.Add Trim(Nvl(cllTemp("_pati_name"))), "_����"
    cllPati_Out.Add Trim(Nvl(cllTemp("_pati_sex"))), "_�Ա�"
    cllPati_Out.Add Trim(Nvl(cllTemp("_pati_age"))), "_����"
    
    cllPati_Out.Add Trim(Nvl(cllTemp("_outpatient_num"))), "_�����"
    cllPati_Out.Add Trim(Nvl(cllTemp("_inpatient_num"))), "_סԺ��"
    cllPati_Out.Add 0, "_����"
    
    Set cllTemp = objServiceCall.GetJsonListValue("output.data.einvoice_info")
    
    lngEInvoiceID_out = Val(Nvl(cllTemp("_einv_id")))
    strInvoiceNO_Out = Trim(Nvl(cllTemp("_paper_nos")))
   
    Zl_Exsesvr_Geteinvoicesinfo = True
    Exit Function
errHandle:
 
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

 
Public Function zl_ExseSvr_GetUseBillInfo(ByVal strNO As String, ByRef rsInvoice_Out As ADODB.Recordset, _
    Optional blnShowErrMsg As Boolean, Optional strErrmsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡֽ��Ʊ��ʹ����Ϣ
    '���:���õ��ݺ�

    '����: rsInvoice_Out-���ر��ν�����Ϣ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim clldata As Variant, cllTemp As Variant
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    If blnShowErrMsg Then On Error GoTo errHandle:
        
    Set rsInvoice_Out = New ADODB.Recordset
    With rsInvoice_Out
        If .State = adStateOpen Then .Close
        .Fields.Append "ID", adBigInt, , adFldIsNullable
        .Fields.Append "Ʊ�ݺ�", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "ʹ��ԭ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "ʹ��ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ʹ����", adLongVarChar, 100, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
   
    'zl_ExseSvr_GetUseBillInfo
    '   input
    '        occasion    N   1   ҵ�񳡺�:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
    '        inv_type    N   1   Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    '        fee_nos C   1   ���õ��ݺ�,����ö��ŷ���
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("occasion", 5, Json_num)
    strJson = strJson & "," & GetJsonNodeString("inv_type", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_nos", strNO, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_ExseSvr_GetUseBillInfo"
    
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, False) = False Then Exit Function
    'output
    '    code    C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '    message C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '    data[]  C   1   ʹ����ϸ����
    '        use_id  N   1   ʹ��id
    '        invoice_no  C   1   ��Ʊ��
    '        use_note    C   1   ʹ��ԭ��
    '        use_time    C   1   ʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
    '        inv_user    C   1   ��Ʊʹ����
    '        recv_id C   1   ����ID


    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrmsg_Out = "" Then
            strErrmsg_Out = "δ�ҵ����������Ľ�����Ϣ�����飡"
        End If
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Set clldata = objServiceCall.GetJsonListValue("output.data")
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then Exit Function
 
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsInvoice_Out
            .AddNew
            !id = Val(Nvl(cllTemp("_use_id")))
            !Ʊ�ݺ� = Trim(Nvl(cllTemp("_invoice_no")))
            !ʹ��ԭ�� = Nvl(cllTemp("_use_note"))
            !ʹ��ʱ�� = Mid(Nvl(cllTemp("_use_time")), 6)
            !ʹ���� = Nvl(cllTemp("_inv_user"))
            .Update
        End With
    Next
    zl_ExseSvr_GetUseBillInfo = True
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


Public Function Zl_Exsesvr_Getbalanceinfo(ByVal strNos As String, ByRef cllSwapData_Out As Collection, ByRef blnStartEinvoice_out As Boolean, _
    Optional ByVal strInvoiceNO As String, Optional ByVal lng����ID As Long, Optional ByVal byt���� As Byte = 5, _
    Optional ByRef dblEInvoice_Out As Double, Optional ByRef lngԭ����ID_Out As Long, Optional str�Ǽ�ʱ�� As String) As Boolean
 
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ʊ����Ϣ
    '���:
    '����:cllPati_Out-���ز��˼�
    '     lngEInvoiceID_out-���ص���Ʊ��ID
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object, intReturn As Integer
    Dim blnShowErrMsg As Boolean, strErrmsg_Out As String
    Dim cllPati As Collection, cllBalanceInfo As Collection
    
    blnShowErrMsg = True
    
    Dim cllTemp As Collection
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Err = 0: On Error GoTo errHandle
    
    '       query_type  N  1  ��ѯ��Χ:0-����ʣ����;1-������ԭʼ������Ϣ
    '--    occasion  N  1  ���㳡��:1-�շ�,2-Ԥ��(����Ѻ��),3-����(������),4-�Һ�,5-���￨,6-����ҽ������
    '--    fee_nos  C    query_type=2ʱ��Ч:���ݺ�:���㳡��=2ʱ��ΪԤ��NO, ����idδ���룬�ýڵ�ش�
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", 0, Json_num)
    strJson = strJson & "," & GetJsonNodeString("occasion", byt����, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_nos", strNos, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Getbalanceinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, True) = False Then Exit Function
    '  --  output
    '  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --    message  C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    data  C    ������Ϣ
    '  --      pati_info  C    ������Ϣ
    '  --        pati_id  N  1  ����ID
    '  --        pati_pageid  N    ��ҳID
    '  --        pati_name  C  1  ����
    '  --        pati_sex  C  1  �Ա�
    '  --        pati_age  C  1  ����
    '  --        outpatient_num  C  1  �����
    '  --        inpatient_num  C  1  סԺ��
    '  --        insurance_type  N  1  ����
    '  --      balance_info  C    ������Ϣ
    '  --        invoice_no  C  1  ��Ʊ��
    '  --        balance_oldid  N  1  ԭ����ID
    '  --        create_time  C  1  �շ�ʱ��:yyyy-mm-dd hh:mi:ss
    '  --        total  N  1  �����ܶ�
    '  --        balance_unit  N  1  �Ƿ��Լ��λ����
    '  --        balance_type  N  1  "Ԥ��ʱ��Ԥ�����:1-����;2-סԺ ;3-�����סԺ;����ʱ����������:1-����;2-סԺ ;3-�����סԺ;
    '  --        start_einv  N  1  �Ƿ����õ���Ʊ��
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If blnShowErrMsg And strErrmsg_Out <> "" Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllTemp = objServiceCall.GetJsonListValue("output.data.pati_info")
    Set cllPati = New Collection
    '1.����������Ϣ(����ID,��ҳID,����,�Ա�,����,�����,סԺ��,���ࣩ,key("_",�ڵ�����)
     cllPati.Add Val(Nvl(cllTemp("_pati_id"))), "_����ID"
    cllPati.Add Val(Nvl(cllTemp("_pati_pageid"))), "_��ҳID"
    
    cllPati.Add Trim(Nvl(cllTemp("_pati_name"))), "_����"
    cllPati.Add Trim(Nvl(cllTemp("_pati_sex"))), "_�Ա�"
    cllPati.Add Trim(Nvl(cllTemp("_pati_age"))), "_����"
    
    cllPati.Add Trim(Nvl(cllTemp("_outpatient_num"))), "_�����"
    cllPati.Add Trim(Nvl(cllTemp("_inpatient_num"))), "_סԺ��"
    cllPati.Add Val(Nvl(cllTemp("_insurance_type"))), "_����"
    
    
   
    Set cllTemp = objServiceCall.GetJsonListValue("output.data.balance_info")

    dblEInvoice_Out = RoundEx(Val(Nvl(cllTemp("_total"))), 6)
    
     
    blnStartEinvoice_out = Val(Nvl(cllTemp("_start_einv"))) = 1
    lngԭ����ID_Out = Val(Nvl(cllTemp("_balance_oldid")))


    '2.����������Ϣ:(��Ʊ��,����ID,����ID,���ݺ�(����ö���),�Ǽ�ʱ��(yyyy-mm-dd hh24:mi:ss),�Ƿ񲹽���,�Ƿ񲿷��˿�,����Ա���,����Ա����,������,����ID,��Լ��λ����,��������)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add strInvoiceNO, "_��Ʊ��"
    cllBalanceInfo.Add lngԭ����ID_Out, "_����ID"
    cllBalanceInfo.Add 0, "_����ID"
    cllBalanceInfo.Add strNos, "_���ݺ�"
    If str�Ǽ�ʱ�� <> "" Then
        cllBalanceInfo.Add Nvl(cllTemp("_create_time")), "_�Ǽ�ʱ��"
    Else
        cllBalanceInfo.Add str�Ǽ�ʱ��, "_�Ǽ�ʱ��"
    End If
    
    cllBalanceInfo.Add 0, "_�Ƿ񲹽���"
    cllBalanceInfo.Add 0, "_�Ƿ񲿷��˿�"
    cllBalanceInfo.Add UserInfo.���, "_����Ա���"
    cllBalanceInfo.Add UserInfo.����, "_����Ա����"
    cllBalanceInfo.Add dblEInvoice_Out, "_������"
    cllBalanceInfo.Add lng����ID, "_����ID"
    cllBalanceInfo.Add Val(Nvl(cllTemp("_balance_unit"))), "_��Լ��λ����"
    cllBalanceInfo.Add Val(Nvl(cllTemp("_balance_type"))), "_��������"
 
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
      
   
    Zl_Exsesvr_Getbalanceinfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function Zl_Exsesvr_GetbalanceinfoFromNos(ByVal strNos As String, Optional ByVal byt���� As Byte = 5, _
    Optional ByRef dblEInvoice_Out As Double, Optional ByRef lngԭ����ID_Out As Long, _
    Optional ByRef blnStartEinvoice_out As Boolean) As Boolean
 
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ʊ����Ϣ
    '���:strNos-���ݺ�
    '����:lngԭ����ID_Out-����ԭ����ID
    '     dblEInvoice_Out-����ԭʼ���
    '     blnStartEinvoice_Out-�Ƿ����õ���Ʊ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object, intReturn As Integer
    Dim blnShowErrMsg As Boolean, strErrmsg_Out As String
    Dim cllTemp As Collection
    
    blnShowErrMsg = True
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrmsg_Out = "���ӷ��������ʧ�ܣ��޷���ȡ��Ч�Ĳ�����Ϣ!"
        If blnShowErrMsg Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '   query_type  N  1  ��ѯ��Χ:0-����ʣ����;1-������ԭʼ������Ϣ
    '--    occasion  N  1  ���㳡��:1-�շ�,2-Ԥ��(����Ѻ��),3-����(������),4-�Һ�,5-���￨,6-����ҽ������
    '--    fee_nos  C    query_type=2ʱ��Ч:���ݺ�:���㳡��=2ʱ��ΪԤ��NO, ����idδ���룬�ýڵ�ش�
    strJson = ""
    
    strJson = strJson & "" & GetJsonNodeString("query_type", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("occasion", byt����, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_nos", strNos, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "Zl_Exsesvr_Getbalanceinfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, True) = False Then Exit Function
    '  --  output
    '  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
    '  --    message  C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    '  --    data  C    ������Ϣ
    '  --      balance_info  C    ������Ϣ
    '  --        invoice_no  C  1  ��Ʊ��
    '  --        balance_oldid  N  1  ԭ����ID
    '  --        create_time  C  1  �շ�ʱ��:yyyy-mm-dd hh:mi:ss
    '  --        total  N  1  �����ܶ�
    '  --        balance_unit  N  1  �Ƿ��Լ��λ����
    '  --        balance_type  N  1  "Ԥ��ʱ��Ԥ�����:1-����;2-סԺ ;3-�����סԺ;����ʱ����������:1-����;2-סԺ ;3-�����סԺ;
    '  --        start_einv  N  1  �Ƿ����õ���Ʊ��
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrmsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If blnShowErrMsg And strErrmsg_Out <> "" Then MsgBox strErrmsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
     
    
    Set cllTemp = objServiceCall.GetJsonListValue("output.data.balance_info")

    dblEInvoice_Out = RoundEx(Val(Nvl(cllTemp("_total"))), 6)
    blnStartEinvoice_out = Val(Nvl(cllTemp("_start_einv"))) = 1
    lngԭ����ID_Out = Val(Nvl(cllTemp("_balance_oldid")))
    Zl_Exsesvr_GetbalanceinfoFromNos = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

