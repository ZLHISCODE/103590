Attribute VB_Name = "mdlPubInvoice"
Option Explicit
Public Type Ty_FactProperty
    lngShareUseID As Long   '������������ID
    strUseType As String ' ʹ�����
    intInvoiceFormat As Integer '��ӡ�ķ�Ʊ��ʽ,��Ʊ��ʽ���
    intInvoicePrint As Integer     '��ӡ��ʽ:0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
    bln�ϸ���� As Boolean
End Type

Public Function GetShareInvoiceGroupID(ByVal bytKind As Byte) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��Ʊ�ֵĹ���Ʊ������
    '����:���˺�
    '����:2011-04-29 10:24:48
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, cllBillInfor As Collection
    On Error GoTo errH
    If zl_ExseSvr_GetReceiveInvoice(bytKind, "", cllBillInfor, False, "", , , , , , 1, True, rsTemp) = False Then Exit Function
    Set GetShareInvoiceGroupID = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "") As Long
'���ܣ���ȡ�������ò���ָ��Ʊ��������÷�Χ�ڵ�����ID
'������bytKind      =   Ʊ��
'      intNum       =   Ҫ��ӡ��Ʊ������
'      lngLastUseID =   �ϴ�ʹ�õ�����ID
'      lngShareUseID=   ���ز���ָ���Ĺ���ID
'      strBill      =   ��ǰƱ�ݺţ����ڼ���������ε�Ʊ�ݷ�Χ
'      strUseType-ʹ�����
'���أ�
'      >0   =   �ɹ������õ�����ID
'      =0   =   ʧ��
'      -1   =   û������(����򲻹�����δ����),δ���ù���
'      -2   =   û������(����򲻹�����δ����),���õĹ���������򲻹�
'      -3   =   ָ��Ʊ�ݺŲ��ڵ�ǰ���п����������ε���ЧƱ�ݺŷ�Χ��
'      -4   =   ָ�����ε�Ʊ�ݲ�����
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    Dim cllBillInfo As Collection, cllTemp As Collection
 
    On Error GoTo errH
    '1.�ϴε����������Ƿ���ò�����
    If lngLastUseID > 0 Then
    
        If zl_ExseSvr_GetReceiveInvoice(bytKind, lngLastUseID, cllBillInfo, False, strUseType, "", intNum, True, , glngModul) = False Then Exit Function
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
        If Not cllBillInfo Is Nothing Then
            If cllBillInfo.Count <> 0 Then
                 If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '����û�е�ǰƱ�ݺ�
                
                Set cllTemp = cllBillInfo(1)
                blnTmp = False
                strPre = nvl(cllTemp("_prefix_text"))
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(nvl(cllTemp("_start_no"))) And UCase(strBill) <= UCase(nvl(cllTemp("_end_no"))) And Len(strBill) = Len(nvl(cllTemp("_start_no")))) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
            ElseIf intNum > 1 Then  '����ȷ���������ε���ʱ,��ǰƱ�ݺ��������β�����
                GetInvoiceGroupID = -4: Exit Function
            End If
        ElseIf intNum > 1 Then  '����ȷ���������ε���ʱ,��ǰƱ�ݺ��������β�����
            GetInvoiceGroupID = -4: Exit Function
        End If
    End If
    
    '2.�ϴε��������β����û򲻿���ʱ,ȡ������Ĳ������õ�
    '  �ж��������ʹ�õ�����,�ٵ�����,��������
    If zl_ExseSvr_GetReceiveInvoice(bytKind, 0, cllBillInfo, True, strUseType, UserInfo.����, intNum, True, , glngModul) Then
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
        For i = 1 To cllBillInfo.Count
            Set cllTemp = cllBillInfo(i)
            
            If strBill = "" Then GetInvoiceGroupID = Val(nvl(cllTemp("_recv_id"))): Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
            blnTmp = False
            strPre = nvl(cllTemp("_prefix_text"))
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(nvl(cllTemp("_start_no"))) And UCase(strBill) <= UCase(nvl(cllTemp("_end_no"))) And Len(strBill) = Len(UCase(nvl(cllTemp("_start_no"))))) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = Val(nvl(cllTemp("_recv_id"))): Exit Function
        Next
        lngReturn = IIf(cllBillInfo.Count > 0, -3, -1)
    Else
        lngReturn = -1
    End If

    '3.û�����õ�,ʹ�ñ��ز���ָ���Ĺ�������
    If lngShareUseID > 0 Then
        If zl_ExseSvr_GetReceiveInvoice(bytKind, lngLastUseID, cllBillInfo, False, strUseType, "", intNum, True, , glngModul) = False Then
            GetInvoiceGroupID = lngReturn   '����δ�ҵ���ԭ�����
            Exit Function
        End If
        
        If cllBillInfo.Count = 0 Then
            lngReturn = -2
            GetInvoiceGroupID = lngReturn
            Exit Function
        End If
        
        Set cllTemp = cllBillInfo(1)
        If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '��һ��ʹ��ʱû�е�ǰƱ�ݺ�
        blnTmp = False
        strPre = nvl(cllTemp("_prefix_text"))
        If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
            blnTmp = True
        ElseIf Not (UCase(strBill) >= UCase(nvl(cllTemp("_start_no"))) And UCase(strBill) <= UCase(nvl(cllTemp("_end_no"))) And Len(strBill) = Len(UCase(nvl(cllTemp("_start_no"))))) Then
            blnTmp = True
        End If
        If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
        lngReturn = -3
    End If
    GetInvoiceGroupID = lngReturn   '����δ�ҵ���ԭ�����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function CheckUsedBill(bytKind As Byte, ByVal lng����ID As Long, _
    Optional ByVal strBill As String, _
     Optional ByVal strUseType As String = "") As Long
    '���ܣ���鵱ǰ����Ա�Ƿ��п���Ʊ������(���û���),�����ؿ��õ�����ID
    '������bytKind=Ʊ��
    '      lng����ID=��һ�μ��ʱΪ�������õĹ�������ID,�Ժ�Ϊ�ϴ�ʹ�õ�����ID
    '      strBill=Ҫ��鷶Χ��Ʊ�ݺ�
    '˵����
    '    1.�ڼ�鷶Χʱ,��������ж�������Ʊ��,��ֻҪ������һ��֮�о�����
    '    2.�ڼ�鷶Χʱ,����Ҳ�ڼ�鷶Χ֮�ڡ�
    '    3.���ж�������ʱ,ȱʡ���ٵ�����,��������,"���ʹ�õ�����"ԭ��
    '���أ�
    '      ������Ʊ������ID>0
    '      0=ʧ��
    '      -1:û������(�����δ����)��Ҳû�й���(δ����)
    '      -2:���õĹ���������
    '      -3:ָ��Ʊ�ݺŲ��ڵ�ǰ���÷�Χ��(������������Ʊ�ݵ����)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    Dim cllBillInfos As Collection, cllBillItem As Collection
    Dim cllSharess  As Collection, cllSharessItem As Collection
    Dim i As Long
    
    On Error GoTo errH
    
      '2.�ϴε��������β����û򲻿���ʱ,ȡ������Ĳ������õ�
    '  �ж��������ʹ�õ�����,�ٵ�����,��������
    If Not zl_ExseSvr_GetReceiveInvoice(bytKind, "", cllBillInfos, True, strUseType, UserInfo.����, 1, True, , glngModul) Then
        Set cllBillInfos = New Collection
    End If
    If cllBillInfos Is Nothing Then Set cllBillInfos = New Collection


    If lng����ID = 0 Then
         '�����е�һ�μ��,��û�����ñ��ع���
         If cllBillInfos.Count = 0 Then CheckUsedBill = -1: Exit Function  'Ҳû������Ʊ��
         '������Ʊ�� , ������ԭ�򷵻�
         Set cllBillItem = cllBillInfos(1)
         lngReturn = Val(nvl(cllBillItem("_recv_id")))
    Else
        '�ϴ�ʹ�õ�����ID���һ�μ��Ĺ���ID,���ж�����
        If Not zl_ExseSvr_GetReceiveInvoice(bytKind, lng����ID, cllSharess, True, strUseType, "", 0, True, , glngModul) Then
            Set cllSharess = New Collection
        End If
        
        If cllSharess.Count = 0 Then CheckUsedBill = -2: Exit Function
        Set cllSharessItem = cllSharess(1)
        
        
        If Val(nvl(cllSharessItem("_use_mode"))) = 2 Then '����,Ҫ�ȿ���û������
            If cllBillInfos.Count <> 0 Then
                '�����õģ�����
                Set cllBillItem = cllBillInfos(1)
                lngReturn = Val(cllBillItem("_recv_id"))
            Else
                'û������ȡ����
                If Val(nvl(cllSharessItem("_surplus_num"))) = 0 Then CheckUsedBill = -2: Exit Function '�����Ѿ�����
                lngReturn = Val(cllSharessItem("_recv_id"))
                blnTmp = True
            End If
        Else
            '����Ʊ��
            If Val(nvl(cllSharessItem("_surplus_num"))) > 0 Then
                '��ʣ��
                lngReturn = Val(cllSharessItem("_recv_id"))
            Else
                '������ʣ�������
                If cllBillInfos.Count = 0 Then CheckUsedBill = -1: Exit Function      '��������Ҳû��ʣ��
                Set cllBillItem = cllBillInfos(1)
                lngReturn = Val(cllBillItem("_recv_id"))
            End If
        End If
    End If
    
    '���Ʊ�ŷ�Χ�Ƿ���ȷ
    If strBill <> "" Then
        If blnTmp Then
            '�ڹ��÷�Χ�ڷ�Χ�ж�
            If UCase(Left(strBill, Len(nvl(cllSharessItem("_prefix_text"))))) <> UCase(nvl(cllSharessItem("_prefix_text"))) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(nvl(cllSharessItem("_start_no"))) And UCase(strBill) <= UCase(nvl(cllSharessItem("_end_no"))) And Len(strBill) = Len(UCase(nvl(cllSharessItem("_start_no"))))) Then
                lngReturn = -3
            End If
        Else
            '�ڿ������÷�Χ���ж�
            blnTmp = False
            Set cllBillItem = mdlPubJson.zlGetNodeObjectFromCollect(cllBillInfos, "_" & lngReturn)
           ' rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(nvl(cllBillItem("_prefix_text"))))) <> UCase(nvl(cllBillItem("_prefix_text"))) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(nvl(cllBillItem("_start_no"))) And UCase(strBill) <= UCase(nvl(cllBillItem("_end_no"))) And Len(strBill) = Len(UCase(nvl(cllBillItem("_start_no"))))) Then
                blnTmp = True
            End If
            
            If blnTmp Then
                '����������,�������������м��
                lngReturn = -3
                For i = 1 To cllBillInfos.Count
                    Set cllBillItem = cllBillInfos(i)
                    
                    If nvl(cllBillItem("_recv_id")) <> lngReturn Then
                        blnTmp = False
                        
                        If UCase(Left(strBill, Len(nvl(cllBillItem("_prefix_text"))))) <> UCase(nvl(cllBillItem("_prefix_text"))) Then
                            blnTmp = True
                        ElseIf Not (UCase(strBill) >= UCase(nvl(cllBillItem("_start_no"))) And UCase(strBill) <= UCase(nvl(cllBillItem("_end_no"))) And Len(strBill) = Len(UCase(nvl(cllBillItem("_start_no"))))) Then
                            blnTmp = True
                        End If
                        If Not blnTmp Then lngReturn = Val(cllBillItem("_recv_id")): Exit For
                    End If
                Next
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    CheckUsedBill = 0
End Function

Public Function GetNextBill(lng����ID As Long) As String
'���ܣ�������������ID,��ȡ��һ��ʵ��Ʊ�ݺ�
'˵����1.��ȡ������Χ�ڵ���ЧƱ��ʱ,���ؿ����û�����
'      2.�ſ��ѱ���ĺ���
    Dim strSQL As String, strBill As String
    Dim cllBillInfos As Collection, cllTemp As Collection
    Dim str��һ�ŷ�Ʊ As String
    
    On Error GoTo errH
    
    If Not zl_ExseSvr_GetReceiveInvoice(0, lng����ID, cllBillInfos, True, , UserInfo.����, 1, True, , glngModul) Then
        Set cllBillInfos = New Collection
    End If
    
    If cllBillInfos.Count = 0 Then Exit Function
    
    Set cllTemp = cllBillInfos(1)
    'ȡ��һ������
    If nvl(cllTemp("_inv_no_cur")) = "" Then
        strBill = UCase(nvl(cllTemp("_start_no")))
    Else
        strBill = UCase(zlCommFun.IncStr(nvl(cllTemp("_inv_no_cur"))))
    End If
    
    '���ʹ����ϸ�Ƿ�ʹ�ø�Ʊ��
    If Not zl_ExseSvr_GetNextInvoice(lng����ID, strBill, str��һ�ŷ�Ʊ) Then Exit Function
    If str��һ�ŷ�Ʊ = "" Then Exit Function
    strBill = str��һ�ŷ�Ʊ
    GetNextBill = strBill
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function zl_GetInvoicePreperty(ByVal lngModule As Long, _
    ByVal intƱ�� As Integer, Optional strʹ����� As String) As Ty_FactProperty
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ʽ
    '���:intƱ��:1- �շ��վ�, 2 - Ԥ���վ�, 3 - �����վ�, 4 - �Һ��վ�, 5 - ���￨
    '����:��Ʊ���������
    '����:���˺�
    '����:2011-07-19 16:43:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Ty_Fact As Ty_FactProperty, strFactType As String, varData As Variant, varTemp As Variant
    Dim strShareTypeUseID As String, lng����Ʊ�� As Long, lngʹ��Ʊ�� As Long
    Dim strFactTypeFormat As String, strFacePrintMode As String
    Dim intPrintMode As Long, intPrintMode1 As Long, lng����ID As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset, strValue As String
    Dim i As Long, lngFormat As Long, lngFormat1 As Long
    
    strFactType = Switch(intƱ�� = 1, "�����շ�Ʊ������", intƱ�� = 2, "����Ԥ��Ʊ������", intƱ�� = 3, "���ý���Ʊ������", intƱ�� = 4, "���ùҺ�Ʊ������", intƱ�� = 5, "����ҽ�ƿ�����", True, "")
    strFactTypeFormat = Switch(intƱ�� = 1, "�շѷ�Ʊ��ʽ", intƱ�� = 2, "Ԥ����Ʊ��ʽ", intƱ�� = 3, "���ʷ�Ʊ��ʽ", intƱ�� = 4, "�Һŷ�Ʊ��ʽ", intƱ�� = 5, "ҽ�ƿ���Ʊ��ʽ", True, "")
    strFacePrintMode = Switch(intƱ�� = 1, "�շѷ�Ʊ��ӡ��ʽ", intƱ�� = 2, "Ԥ����Ʊ��ӡ��ʽ", intƱ�� = 3, "���˽��ʴ�ӡ", intƱ�� = 4, "�Һŷ�Ʊ��ӡ��ʽ", intƱ�� = 5, "ҽ�ƿ���Ʊ��ӡ��ʽ", True, "")
    
    If strFactType = "" Then Exit Function
    
    
    'Ʊ���ϸ����
    If intƱ�� >= 1 And intƱ�� <= 4 Then
        strValue = zlDatabase.GetPara("Ʊ���ϸ����", glngSys, , "00000")
        Ty_Fact.bln�ϸ���� = Mid(strValue, intƱ��, 1) = "1"
    End If
    
    '����λ����ʾ, ÿһλ����ͬ��ҵ������:
    '��һλ:         �շ�
    '�ڶ�λ:         Ԥ��
    '����λ:         ����
    '����λ:         �Һ�
    'ÿλ��1��0��ʾ,1��ʾ�ϸ����;0-��ʾ���ϸ����

    
    '78751:���ϴ�,2014/10/20,����Ԥ��Ʊ�ݴ�ӡ��ʽ
    Ty_Fact.strUseType = strʹ�����
 
    strFactTypeFormat = Trim(zlDatabase.GetPara(strFactTypeFormat, glngSys, lngModule, ""))
    '��ʽ:ʹ�����1,��ʽ1|ʹ�����2,��ʽ2...
    varData = Split(strFactTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lngFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then lngFormat1 = lngFormat
        If Trim(varTemp(0)) = strʹ����� And lngFormat <> 0 Then
            Ty_Fact.intInvoiceFormat = lngFormat: Exit For
        End If
    Next
    If Ty_Fact.intInvoiceFormat = 0 And lngFormat1 <> 0 Then Ty_Fact.intInvoiceFormat = lngFormat
 
    '��ӡ��ʽ(0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ)
    '����50656
'    If intƱ�� = 2 Then
'        'Ԥ����Ϊ�Զ���ӡ
'        Ty_Fact.intInvoicePrint = 1
'    Else
        '��ΪGetpara�ͻ����˵�,���Բ������ñ������м�¼
        strFacePrintMode = Trim(zlDatabase.GetPara(strFacePrintMode, glngSys, lngModule, ""))
        Ty_Fact.intInvoicePrint = -1
        '��ʽ:ʹ�����1,��ӡ��ʽ1|ʹ�����2,��ӡ��ʽ2...
        varData = Split(strFacePrintMode, "|")
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",,", ",")
            intPrintMode = Val(varTemp(1))
            If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
            If Trim(varTemp(0)) = strʹ����� Then
                Ty_Fact.intInvoicePrint = intPrintMode: Exit For
            End If
        Next
        If Ty_Fact.intInvoicePrint < 0 Then Ty_Fact.intInvoicePrint = intPrintMode1
'    End If
    '��������
    
    '��ʽ:����ID1,ʹ�����1|....
    strShareTypeUseID = Trim(zlDatabase.GetPara(strFactType, glngSys, lngModule, "0"))
    varData = Split(strShareTypeUseID, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lng����ID = Val(varTemp(0))
        If intƱ�� = 2 Or intƱ�� = 5 Then
            If Val(varTemp(1)) = 0 Then lng����Ʊ�� = lng����ID    '���õ�.
            If Val(varTemp(1)) = Val(strʹ�����) And lng����ID <> 0 Then
                lngʹ��Ʊ�� = lng����ID
            End If
        Else
            If Trim(varTemp(1)) = "" Then lng����Ʊ�� = lng����ID    '���õ�.
            If Trim(varTemp(1)) = strʹ����� And lng����ID <> 0 Then
                lngʹ��Ʊ�� = lng����ID
            End If
        End If
    Next
    
    On Error GoTo errHandle
    '����˳��
    '1.��ʹ��
    '2.ʹ��������ֵ�
    '3.����ʹ������
    Dim cllBillInfo As Collection
    If zl_ExseSvr_GetReceiveInvoice(0, lng����Ʊ�� & "," & lngʹ��Ʊ��, cllBillInfo, False) = False Then
        zl_GetInvoicePreperty = Ty_Fact
        Exit Function
    End If
    If cllBillInfo.Count <> 0 Then
         Ty_Fact.lngShareUseID = Val(nvl(cllBillInfo(1)("_recv_id"))) ' '���õ�����ID
    End If
    zl_GetInvoicePreperty = Ty_Fact
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zl_GetInvoiceUserType(ByVal lng����ID As Long, ByVal lng��ҳid As Long, Optional intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ʊ��ʹ�����
    '����:��Ʊ��ʹ�����
    '����:���˺�
    '����:2011-04-29 11:03:35
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strʹ����� As String
    On Error GoTo errHandle
    If zl_ExseSvr_GetPatiInvoiceClass(lng����ID, lng��ҳid, intInsure, strʹ�����) = False Then Exit Function
    zl_GetInvoiceUserType = strʹ�����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlStartFactUseType(ByVal intƱ�� As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�ʹ����ʹ������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-10 16:11:47
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln�Ƿ����� As Boolean
    
    On Error GoTo errHandle
    If zl_ExseSvr_InvoiceClassUsed(intƱ��, bln�Ƿ�����, True, , glngModul) = False Then Exit Function
    zlStartFactUseType = bln�Ƿ�����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


