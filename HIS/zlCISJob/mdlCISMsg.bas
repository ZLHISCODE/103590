Attribute VB_Name = "mdlCISMsg"
Option Explicit

Public Function ZLHIS_CIS_001(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'�¿�����ҽ��
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    Dim lng�������id As Long, str����������� As String
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim bln��Ϣƽ̨ As Boolean
    
    On Error GoTo errH
    
    If Not objMip Is Nothing Then
        If objMip.IsConnect Then bln��Ϣƽ̨ = True
    End If
    
    With objXML
        .ClearXmlText
        .AppendNode "patient_info"  '���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic"  '���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        
        .AppendData "clinic_dept_id", arrInput(Index) '<�������id>���ͣ�N
        lng�������id = arrInput(Index): Index = Index + 1
        str����������� = arrInput(Index): Index = Index + 1 '<�������>���ͣ�S
        
        If str����������� = "" And bln��Ϣƽ̨ Then
            strSql = "select ���� from ���ű� where id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISMsg", lng�������id)
            str����������� = rsTmp!���� & ""
        End If
        .AppendData "clinic_dept_title", str�����������
        
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "new_order" '���ڵ�[����ҽ��]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "order_urgency", arrInput(Index): Index = Index + 1 '<������־>���ͣ�N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<ҽ����Ч>���ͣ�N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ�����>���ͣ�S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ����������>���ͣ�S
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<����ҽ��>���ͣ�S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<��������id>���ͣ�N
        .AppendData "create_dept_title", arrInput(Index): Index = Index + 1 '<������������>���ͣ�S
        .AppendNode "new_order", True
        
        strMsgNo = "ZLHIS_CIS_001"
     
        If bln��Ϣƽ̨ Then Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_001 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_002(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'ֹͣ����ҽ��
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_dept_title", arrInput(Index) '<��������>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "stop_order" ', True'���ڵ�[����ҽ��]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<ҽ����Ч>���ͣ�N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ�����>���ͣ�S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ����������>���ͣ�S
        .AppendData "stop_doctor", arrInput(Index): Index = Index + 1 '<ͣ��ҽ��>���ͣ�S
        .AppendData "stop_time", arrInput(Index): Index = Index + 1 '<ͣ��ʱ��>���ͣ�D
        .AppendData "order_urgency", arrInput(Index): Index = Index + 1 '<�Ƿ��ǽ���ҽ��>���ͣ�N '0-��ͨ,1-������2-��¼
        .AppendNode "stop_order", True
        
        strMsgNo = "ZLHIS_CIS_002"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_002 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_003(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'���ϻ���ҽ��
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "cancel_order" ', True'���ڵ�[����ҽ��]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<ҽ����Ч>���ͣ�N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ�����>���ͣ�S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ����������>���ͣ�S
        .AppendData "execute_kind", arrInput(Index): Index = Index + 1 '<ִ�з���>���ͣ�N
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<ִ�п���id>���ͣ�N
        .AppendData "cancel_doctor", arrInput(Index): Index = Index + 1 '<����ҽ��>���ͣ�S
        .AppendData "cancel_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendNode "cancel_order", True
        
        strMsgNo = "ZLHIS_CIS_003"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_003 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_004(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'ҽ�����밲��
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "arrange_order" ', True'���ڵ�[����ҽ��]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<ҽ����Ч>���ͣ�N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ�����>���ͣ�S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ����������>���ͣ�S
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<ִ�п���id>���ͣ�N
        .AppendNode "arrange_order", True
        
        strMsgNo = "ZLHIS_CIS_004"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_004 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_005(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'ҽ��ִ�а������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" '', True'���ڵ�'<������Ϣ>
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" '', True'���ڵ�'<������Ϣ>
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "patient_order" '', True'���ڵ�'<����ҽ��>
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<ҽ����Ч>���ͣ�N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ�����>���ͣ�S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ����������>���ͣ�S
        .AppendData "order_item_id", arrInput(Index): Index = Index + 1 '<ҽ����Ŀid>���ͣ�N
        .AppendData "order_item_title", arrInput(Index): Index = Index + 1 '<ҽ����Ŀ>���ͣ�S
        .AppendNode "patient_order", True
        .AppendNode "arrange_result" ', True'���ڵ�'<�������>
        .AppendData "arrange_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "arrange_room", arrInput(Index): Index = Index + 1 '<���ŷ���>���ͣ�S
        .AppendNode "arrange_result", True
        
        strMsgNo = "ZLHIS_CIS_005"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_005 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_006(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'����ҩƷҽ������
    Dim objXML As New zl9ComLib.clsXML
    Dim rsTmp As ADODB.Recordset
    Dim strXML As String
    Dim Index As Integer
    Dim i As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "drug_bill" ', True'���ڵ�[]
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "send_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "send_person", arrInput(Index): Index = Index + 1 '<������Ա>���ͣ�S
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<���ݺ���>���ͣ�S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "charge_state", arrInput(Index): Index = Index + 1 '<�շ�״̬>���ͣ�N
        .AppendNode "send_order" ', True'�ü�¼�����ַ��� '���ڵ�[����ҽ��]
        Set rsTmp = arrInput(Index)
        For i = 1 To rsTmp.RecordCount
            .AppendData "order_id", rsTmp!ҽ��ID '<ҽ��id>���ͣ�N
            .AppendData "order_relevant_id", rsTmp!���ID '<���ID>���ͣ�N
            .AppendData "order_info", rsTmp!ҽ������ & "" '<ҽ������>���ͣ�S
            .AppendData "order_rate", rsTmp!ִ��Ƶ�� & ""  '<ִ��Ƶ��>���ͣ�S
            .AppendData "order_route_id", rsTmp!��ҩ;��ID '<��ҩ;��id>���ͣ�N
            .AppendData "order_route", rsTmp!��ҩ;�� & ""  '<��ҩ;��>���ͣ�S
            .AppendData "order_starttime", rsTmp!��ʼʱ�� & ""  '<��ʼʱ��>���ͣ�D
            .AppendData "order_single", Format(Val(rsTmp!���� & ""), "0.0") '<����>���ͣ�N
            .AppendData "order_total", Format(Val(rsTmp!���� & ""), "0.0")  '<����>���ͣ�N
            .AppendData "order_entrust", rsTmp!ҽ������ & ""  '<ҽ������>���ͣ�S
            .AppendData "order_item_id", rsTmp!Ʒ��ID '<Ʒ��id>���ͣ�N
            .AppendData "Drug_item_kind", rsTmp!ҩƷ��� & ""  '<ҩƷ���>���ͣ�S
            .AppendData "Drug_item_id", rsTmp!ҩƷID '<ҩƷid>���ͣ�N
            .AppendData "execute_dept_id", rsTmp!ִ�в���ID '<ִ�в���id>���ͣ�N
            rsTmp.MoveNext
        Next
        .AppendNode "send_order", True
        .AppendNode "drug_bill", True
        
        strMsgNo = "ZLHIS_CIS_006"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_006 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_007(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'���ﻼ��ת��
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    If Not (objMip Is Nothing) Then
        If objMip.IsConnect Then
            Exit Function
        End If
    End If
    
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" '', True'���ڵ�'<������Ϣ>
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" '', True'���ڵ�'<������Ϣ>
        .AppendData "clinic_id", arrInput(Index): Index = Index + 1 '<����Һ�id>���ͣ�N
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_dept_title", arrInput(Index) '<��������>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "change_clinic" '', True'���ڵ�'<ת����Ϣ>
        .AppendData "change_dept_id", arrInput(Index): Index = Index + 1 '<ת�����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_dept_title", arrInput(Index) '<ת���������>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_doctor_id", arrInput(Index) '<ת��ҽ��id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_doctor", arrInput(Index) '<ת��ҽ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_room", arrInput(Index) '<ת������>���ͣ�S
        Index = Index + 1
        .AppendData "change_exedoctor", arrInput(Index): Index = Index + 1 '<ת��ִ��ҽ��>���ͣ�S
        .AppendNode "change_clinic", True
        
        strMsgNo = "ZLHIS_CIS_007"
        
        Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
    End With
    ZLHIS_CIS_007 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_008(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'��Һ���ε���
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<��ҳid>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<���ﲡ������>���ͣ�S
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "order_info" ', True'���ڵ�[ҽ������]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<ҽ����Ч>���ͣ�N
        .AppendNode "order_info", True
        .AppendNode "transfusion_info" ', True'���ڵ�[��Һ��Ϣ]
        .AppendData "transfusion_id", arrInput(Index): Index = Index + 1 '<��Һid>���ͣ�N
        .AppendData "transfusion_batch", arrInput(Index): Index = Index + 1 '<��Һ����>���ͣ�N
        .AppendNode "transfusion_info", True
        
        strMsgNo = "ZLHIS_CIS_008"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
    End With
    ZLHIS_CIS_008 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_009(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'���ﻼ�߽���
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lng����ID  As Long
    Dim strTmp As String
    Dim strMsgNo As String
    
    On Error GoTo errH
    If Not (objMip Is Nothing) Then
        If objMip.IsConnect Then
            Exit Function
        End If
    End If
    strSql = "Select ���֤��, סԺ��, ���￨��, ������, ҽ����, �Ա�, ����, ��������, ����״��, �ѱ�, ����, ����, ְҵ, ѧ��, ��ͥ��ַ, ��ͥ�绰, ��ͥ��ַ�ʱ� As ��ͥ�ʱ�, ��ϵ�˵绰," & _
        " ��ϵ�˹�ϵ, ��ϵ�˵�ַ, ������λ, ��λ�绰, ��λ�ʱ�, ����֤�� As ����֤��, �����ص�, �໤��, ���ڵ�ַ As ������ַ, ���ڵ�ַ�ʱ� As �����ʱ�, ����, ҽ�Ƹ��ʽ" & _
        " From ������Ϣ Where ����id = [1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISMsg", lng����ID)  'If Not rsTmp.EOF Then '����������жϣ�����д�Ӧ�ñ�¶����
    If rsTmp.EOF Then Exit Function '�������ﲡ��û����
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" '', True'���ڵ�'<������Ϣ>
        .AppendData "patient_id", arrInput(Index) '<����id>���ͣ�N
            lng����ID = Val(arrInput(Index)): Index = Index + 1
         If TypeName(arrInput(Index)) <> "Error" Then
            .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Else
            .AppendData "patient_name", ""
        End If
        Index = Index + 1
        .AppendData "identity_card", "" & rsTmp!���֤��
        .AppendData "in_number", "" & rsTmp!סԺ��
        If TypeName(arrInput(Index)) <> "Error" Then
            .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Else
            .AppendData "out_number", ""
        End If
        Index = Index + 1
        .AppendData "visit_number", "" & rsTmp!���￨��
        .AppendData "health_number", "" & rsTmp!������
        .AppendData "medical_number", "" & rsTmp!ҽ����
        .AppendData "patient_sex", "" & rsTmp!�Ա�
        .AppendData "patient_age", "" & rsTmp!����
        strTmp = Decode("" & rsTmp!�Ա�, "��", 1, "Ů", 2, "δ֪", 3, "")
        .AppendData "patient_age_code", strTmp ' �Ա����
        .AppendData "patient_birthday", "" & rsTmp!��������
        .AppendData "patient_marriage", "" & rsTmp!����״��
        strTmp = Decode("" & rsTmp!����״��, "δ��", 1, "�ѻ�", 2, "ɥż", 3, "���", 4, "����", 9, "")
        .AppendData "patient_marriage_code", strTmp ' ����״������
        .AppendData "patient_chargetype", "" & rsTmp!�ѱ�
        .AppendData "patient_nationality", "" & rsTmp!����
        .AppendData "patient_nation", "" & rsTmp!����
        .AppendData "patient_profession", "" & rsTmp!ְҵ
        .AppendData "patient_education", "" & rsTmp!ѧ��
        .AppendData "home_addr", "" & rsTmp!��ͥ��ַ
        .AppendData "home_telephone", "" & rsTmp!��ͥ�绰
        .AppendData "home_postcode", "" & rsTmp!��ͥ�ʱ�
        .AppendData "contact_telephone", "" & rsTmp!��ϵ�˵绰
        .AppendData "contact_relation", "" & rsTmp!��ϵ�˹�ϵ
        .AppendData "contact_addr", "" & rsTmp!��ϵ�˵�ַ
        .AppendData "patient_work", "" & rsTmp!������λ
        .AppendData "work_telephone", "" & rsTmp!��λ�绰
        .AppendData "work_addr", "" ' ��λ��ַ"--����ֶ���ʱû��
        .AppendData "work_postcode", "" & rsTmp!��λ�ʱ�
        .AppendData "patient_other_papers", "" & rsTmp!����֤��
        .AppendData "patient_birthplace", "" & rsTmp!�����ص�
        .AppendData "patient_guardian", "" & rsTmp!�໤��
        .AppendData "patient_height", arrInput(Index): Index = Index + 1 '���"
        .AppendData "patient_weight", arrInput(Index): Index = Index + 1 ' ����"
        .AppendData "residence_addr", "" & rsTmp!������ַ
        .AppendData "residence_postcode", "" & rsTmp!�����ʱ�
        .AppendData "native_place", "" & rsTmp!����
        .AppendData "patient_payment", "" & rsTmp!ҽ�Ƹ��ʽ
        strTmp = Decode("" & rsTmp!ҽ�Ƹ��ʽ, "������ҽ�Ʊ���", 1, "����ҽ��", 2, "��ͳ��", 3, "��ҵ����", 4, "�Է�ҽ��", 5, "����", 6, "")
        .AppendData "patient_payment_code", strTmp ' ���ʽ����
        .AppendNode "patient_info", True
        .AppendNode "receive_clinic" '', True'���ڵ�'<������Ϣ>
        .AppendData "clinic_id", arrInput(Index): Index = Index + 1 '<�Һ�id>���ͣ�N
        .AppendData "return_visit", arrInput(Index): Index = Index + 1 '<�Ƿ���>���ͣ�N
        .AppendData "emergency_treatment", arrInput(Index): Index = Index + 1 '<�Ƿ���>���ͣ�N
        .AppendData "receive_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "receive_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then
            .AppendData "receive_dept_title", arrInput(Index) '<�����������>���ͣ�S
        Else
            .AppendData "receive_dept_title", ""
        End If
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then
            .AppendData "receive_room", arrInput(Index) '<��������>���ͣ�S
        Else
            .AppendData "receive_room", ""
        End If
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "receive_doctor", arrInput(Index) '<����ҽ��>���ͣ�S
        Index = Index + 1
        .AppendNode "receive_clinic", True
        
        strMsgNo = "ZLHIS_CIS_009"
        

        Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_009 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_010(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'�´ﻼ�����
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1   '<������Դ>���ͣ�N
        .AppendData "clinic_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendNode "patient_clinic", True
        .AppendNode "diagnose_info" ', True'���ڵ�[]
        .AppendData "diagnose_id", Val(arrInput(Index)): Index = Index + 1 '<���id>���ͣ�N
        .AppendData "diagnose_kind", arrInput(Index): Index = Index + 1 '<�������>���ͣ�S
        .AppendData "diagnose_question", arrInput(Index): Index = Index + 1 '<�Ƿ�����>���ͣ�N
        .AppendData "diagnose_serial", arrInput(Index): Index = Index + 1 '<��ϴ���>���ͣ�N
        .AppendData "diagnose_code", arrInput(Index): Index = Index + 1 '<��ϱ���>���ͣ�S
        .AppendData "illness_code", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        .AppendData "illness_addition_code", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        .AppendData "illness_kind", arrInput(Index): Index = Index + 1 '<�������>���ͣ�S
        .AppendData "syndrome_code", arrInput(Index): Index = Index + 1 '<֤�����>���ͣ�S
        .AppendData "syndrome_title", arrInput(Index): Index = Index + 1 '<֤������>���ͣ�S
        .AppendData "record_time", Format(arrInput(Index), "YYYY-MM-DD HH:MM:SS"): Index = Index + 1 '<��¼����>���ͣ�D
        .AppendData "record_person", arrInput(Index): Index = Index + 1 '<��¼��Ա>���ͣ�S
        .AppendNode "diagnose_info", True
        
        strMsgNo = "ZLHIS_CIS_010"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_010 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_011(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'�����������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        .AppendData "clinic_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendNode "patient_clinic", True
        .AppendNode "diagnose_info" ', True'���ڵ�[]
        .AppendData "diagnose_id", arrInput(Index): Index = Index + 1 '<���id>���ͣ�N
        .AppendData "diagnose_code", arrInput(Index): Index = Index + 1 '<��ϱ���>���ͣ�S
        .AppendData "illness_code", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        .AppendNode "diagnose_info", True
        
        strMsgNo = "ZLHIS_CIS_011"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_011 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_013(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'סԺ������Һ��������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<��ҳid>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_area_id", arrInput(Index): Index = Index + 1 '<���ﲡ��id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<���ﲡ������>���ͣ�S
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_dept_title", arrInput(Index) '<�����������>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "cancel_reqeust" ', True'���ڵ�[]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "transfusion_id", arrInput(Index): Index = Index + 1 '<��Һid>���ͣ�N
        .AppendData "request_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "request_person", arrInput(Index): Index = Index + 1 '<������Ա>���ͣ�S
        .AppendData "request_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "request_dept_title", arrInput(Index) '<�������>���ͣ�S
        Index = Index + 1
        .AppendData "audit_dept_id", arrInput(Index): Index = Index + 1 '<��˲���id>���ͣ�N
        .AppendNode "cancel_reqeust", True
        
        strMsgNo = "ZLHIS_CIS_013"
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_013 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_015(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'ҽ���ܾ�ִ��֪ͨ
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" '', True'���ڵ�'<������Ϣ>
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index)  '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" '', True'���ڵ�'<������Ϣ>
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "refuse_order" '', True'���ڵ�'<>
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<ҽ����Ч>���ͣ�N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ�����>���ͣ�S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ����������>���ͣ�S
        .AppendData "order_item_id", arrInput(Index): Index = Index + 1 '<ҽ����Ŀid>���ͣ�N
        .AppendData "order_item_title", arrInput(Index): Index = Index + 1 '<ҽ����Ŀ>���ͣ�S
        .AppendNode "refuse_order", True
        
        strMsgNo = "ZLHIS_CIS_015"
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_015 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_016(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'���߼�������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim varArr As Variant
    Dim strLine As String
    Dim strLine1 As String
    Dim strTmp As String
    Dim Index As Integer
    Dim i As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_dept_title", arrInput(Index) '<��������>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "lis_request" ', True'���ڵ�[��������]
        .AppendData "request_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "lis_sample", arrInput(Index): Index = Index + 1 '<����걾>���ͣ�S
        .AppendData "collect_item_id", arrInput(Index): Index = Index + 1 '<�ɼ���ʽid>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "collect_item_title", arrInput(Index) '<�ɼ���ʽ����>���ͣ�S
        Index = Index + 1
        .AppendData "collect_dept_id", arrInput(Index): Index = Index + 1 '<�ɼ�����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "collect_dept_title", arrInput(Index) '<�ɼ���������>���ͣ�S
        Index = Index + 1
        '������Ŀ lis_item_title����ݲ�����������λ�������ŵ�
        strLine = arrInput(Index): Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then strLine1 = arrInput(Index)
        Index = Index + 1
'        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "lis_item_title", arrInput(Index): Index = Index + 1 '<������Ŀ����>���ͣ�S
        If InStr(strLine, ",") > 0 Then
            varArr = Split(strLine, ",")
            For i = 0 To UBound(varArr)
                .AppendNode "lis_item"
                .AppendData "lis_item_id", Val(varArr(i))
                .AppendNode "lis_item", True
            Next
        Else
            .AppendNode "lis_item"
            .AppendData "lis_item_id", Val(strLine)
            .AppendNode "lis_item", True
        End If
        
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<ִ�п���id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "execute_dept_title", arrInput(Index) '<ִ�п�������>���ͣ�S
        Index = Index + 1
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<���ݺ���>���ͣ�S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<����ҽ��>���ͣ�S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<��������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "create_dept_title", arrInput(Index) '<������������>���ͣ�S
        Index = Index + 1
        .AppendNode "lis_request", True
        
        strMsgNo = "ZLHIS_CIS_016"
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_016 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_017(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'���߼������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        .AppendNode "patient_clinic", True
        .AppendNode "check_request" ', True'���ڵ�[�������]
        .AppendData "request_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "check_item_id", arrInput(Index): Index = Index + 1 '<ҽ����Ŀid>���ͣ�N
        .AppendData "check_item_title", arrInput(Index): Index = Index + 1 '<ҽ����Ŀ����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "check_parts", arrInput(Index) '<��λ�嵥>���ͣ�S
        Index = Index + 1
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<ִ�п���id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "execute_dept_title", arrInput(Index) '<ִ�п�������>���ͣ�S
        Index = Index + 1
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<���ݺ���>���ͣ�S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<����ҽ��>���ͣ�S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<��������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "create_dept_title", arrInput(Index) '<������������>���ͣ�S
        Index = Index + 1
        .AppendNode "check_request", True
        
        strMsgNo = "ZLHIS_CIS_017"
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_017 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_018(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'������������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim varArr As Variant
    Dim strLine As String
    Dim strLine1 As String
    Dim strTmp As String
    Dim Index As Integer
    Dim i As Integer
    Dim j As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index)
        Index = Index + 1 '<סԺ��>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index)
        Index = Index + 1 '<�����>���ͣ�S
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index)
        Index = Index + 1 '<����id>���ͣ�N
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        .AppendNode "patient_clinic", True
        .AppendNode "oper_request" ', True'���ڵ�[��������]
        .AppendData "request_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        
        strLine = arrInput(Index): Index = Index + 1 '<������Ŀid>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then strLine1 = arrInput(Index) '<������Ŀ����>���ͣ�S���˽���ݲ����
        Index = Index + 1
        If InStr(strLine, ",") > 0 Then
            varArr = Split(strLine, ",")
            For i = 0 To UBound(varArr)
                .AppendNode "oper_item"
                .AppendData "oper_item_id", Val(varArr(i))
                .AppendNode "oper_item", True
            Next
        Else
            .AppendNode "oper_item"
            .AppendData "oper_item_id", Val(strLine)
            .AppendNode "oper_item", True
        End If
        
        strLine = ""
        If TypeName(arrInput(Index)) <> "Error" Then strLine = arrInput(Index)  '<������Ŀid>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then strLine1 = arrInput(Index)  '<������Ŀ����>���ͣ�S���˽���ݲ����
        Index = Index + 1
        If strLine <> "" Then
            If InStr(strLine, ",") > 0 Then
                varArr = Split(strLine, ",")
                For i = 0 To UBound(varArr)
                    .AppendNode "narcosis_item"
                    .AppendData "narcosis_item_id", Val(varArr(i))
                    .AppendNode "narcosis_item", True
                Next
            Else
                .AppendNode "narcosis_item"
                .AppendData "narcosis_item_id", Val(strLine)
                .AppendNode "narcosis_item", True
            End If
        End If
         
        strLine = ""
        If TypeName(arrInput(Index)) <> "Error" Then strLine = arrInput(Index) '<����ҽ��>���ͣ�S
        Index = Index + 1
        If strLine <> "" Then
            If InStr(strLine, ",") > 0 Then
                varArr = Split(strLine, ",")
                For i = 0 To UBound(varArr)
                    .AppendData "major_doctor", varArr(i)
                Next
            Else
                .AppendData "major_doctor", strLine
            End If
        End If
        
        strLine = ""
        
        If TypeName(arrInput(Index)) <> "Error" Then strLine = arrInput(Index) '<����ҽ��>���ͣ�S
        Index = Index + 1
        If strLine <> "" Then
            If InStr(strLine, ",") > 0 Then
                varArr = Split(strLine, ",")
                For i = 0 To UBound(varArr)
                    .AppendData "assistant_doctor", varArr(i)
                Next
            Else
                .AppendData "assistant_doctor", strLine
            End If
        End If
         
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<ִ�п���id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "execute_dept_title", arrInput(Index)
        Index = Index + 1 '<ִ�п�������>���ͣ�S
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<���ݺ���>���ͣ�S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<����ҽ��>���ͣ�S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<��������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "create_dept_title", arrInput(Index) '<������������>���ͣ�S
        Index = Index + 1
        .AppendNode "oper_request", True
        
        strMsgNo = "ZLHIS_CIS_018"
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_018 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_019(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'������Ѫ����
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        .AppendNode "patient_clinic", True
        .AppendNode "blood_request" ', True'���ڵ�[��������]
        .AppendData "request_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<ִ�п���id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "execute_dept_title", arrInput(Index) '<ִ�п�������>���ͣ�S
        Index = Index + 1
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<���ݺ���>���ͣ�S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<����ҽ��>���ͣ�S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<��������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "create_dept_title", arrInput(Index) '<������������>���ͣ�S
        Index = Index + 1
        .AppendNode "blood_request", True
        
        strMsgNo = "ZLHIS_CIS_019"
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_019 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_024(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'����ҽ������ ��������������ϲ�����סԺ��������˲���
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        .AppendNode "patient_clinic", True
        .AppendNode "cancel_order" ', True'���ڵ�[]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ�����>���ͣ�S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ����������>���ͣ�S
        .AppendNode "cancel_order", True
        
        strMsgNo = "ZLHIS_CIS_024"
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_024 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_MsgReadAfter(ByRef objMip As zl9ComLib.clsMipModule, ByVal strMsgNo As String, ParamArray arrInput() As Variant) As String
'���ܣ�Σ��ֵ��Ϣ�Ķ����͵���Ϣ������Ϣƽ̨���ӿ���ʱ�Żᷢ��
'������strMsgNo ZLHIS_CIS_025-���Σ��ֵ�Ķ�֪ͨ��ZLHIS_CIS_014������Σ��ֵ�Ķ�֪ͨ
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>S
        .AppendData "clinic_id", arrInput(Index): Index = Index + 1  '<��ҳid���Һ�ID>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "critical_read" ', True'���ڵ�[]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendNode "critical_read", True
        Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
    End With
    ZLHIS_CIS_MsgReadAfter = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_026(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'����������ѣ�ʵϰҽ���´��ҽ����Ҫ��ʽҽ�����
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
 
    On Error GoTo errH
    
    With objXML
        .ClearXmlText
        .AppendNode "patient_info"  '���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic"  '���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<�������>���ͣ�S
        
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "new_order" '���ڵ�[����ҽ��]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "order_urgency", arrInput(Index): Index = Index + 1 '<������־>���ͣ�N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<ҽ����Ч>���ͣ�N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ�����>���ͣ�S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ����������>���ͣ�S
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<����ҽ��>���ͣ�S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<��������id>���ͣ�N
        .AppendData "create_dept_title", arrInput(Index): Index = Index + 1 '<������������>���ͣ�S
        .AppendNode "new_order", True
        
        strMsgNo = "ZLHIS_CIS_026"
     
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_026 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_027(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'��ͣ�������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_dept_title", arrInput(Index) '<��������>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "stop_order" ', True'���ڵ�[����ҽ��]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<ҽ����Ч>���ͣ�N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ�����>���ͣ�S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<��Ҫҽ����������>���ͣ�S
        .AppendData "stop_doctor", arrInput(Index): Index = Index + 1 '<ͣ��ҽ��>���ͣ�S
        .AppendData "stop_time", arrInput(Index): Index = Index + 1 '<ͣ��ʱ��>���ͣ�D
        .AppendData "order_urgency", arrInput(Index): Index = Index + 1 '<�Ƿ��ǽ���ҽ��>���ͣ�N '0-��ͨ,1-������2-��¼
        .AppendNode "stop_order", True
        
        strMsgNo = "ZLHIS_CIS_027"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_027 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_031(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'��Ѫ��Ѫ����
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
 
    On Error GoTo errH
    
    With objXML
        .ClearXmlText
        .AppendNode "patient_info"  '���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic"  '���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<�������>���ͣ�S
        
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "new_order" '���ڵ�[����ҽ��]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<����ҽ��>���ͣ�S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<��������id>���ͣ�N
        .AppendData "create_dept_title", arrInput(Index): Index = Index + 1 '<������������>���ͣ�S
        .AppendNode "new_order", True
        strMsgNo = "ZLHIS_CIS_031"
        Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
    End With
    ZLHIS_CIS_031 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_CIS_Audit(ByVal strMsgNo As String, ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'���ܣ�ҽ�������Ϣ����
'�����⼸�� 'ZLHIS_CIS_028������������ѣ�ZLHIS_CIS_029-����ҩ��������ѣ�ZLHIS_CIS_030 -��Ѫ�������

    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer

    On Error GoTo errH
    
    With objXML
        .ClearXmlText
        .AppendNode "patient_info"  '���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic"  '���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<���ﲡ��id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<�������>���ͣ�S
        
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<���ﲡ��>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "new_order" '���ڵ�[����ҽ��]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<����ҽ��>���ͣ�S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<��������id>���ͣ�N
        .AppendData "create_dept_title", arrInput(Index): Index = Index + 1 '<������������>���ͣ�S
        .AppendNode "new_order", True
     
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_CIS_Audit = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function RecMsgToBub(ByRef objMip As zl9ComLib.clsMipModule, ByVal lngDeptID As Long, ByVal int���� As Integer, ByVal strMsgNo As String, ByVal strXML As String, Optional ByVal intView As Integer) As Boolean
'���ܣ�ð������
'������
'      lngDeptID ����ǻ�ʿվ����Ϊ����id��ҽ��վ����ʱ����� intView �����ж��ǲ���id���ǿ���id
'      int���� 1������ҽ��վ��2��סԺҽ��վ��3��סԺ��ʿվ��4��ҽ��վ
'      intView ��ʾ��ʽ��0-��������ʾ��1-��������ʾ ֻ����סԺҽ��վ����ʱʹ�������
    Dim objXML As zl9ComLib.clsXML
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strMsg As String '�����е�����
    Dim strTitle As String '����
    Dim strPar As String '���Ӳ���
    Dim i As Integer
    Dim strTmp1 As String
    Dim strTmp2 As String
    Dim strDeptNode As String
    Dim strAreaNode As String
    Dim blnTmp As Boolean
    
    On Error GoTo errH

    Set objXML = New zl9ComLib.clsXML
    Call objXML.OpenXMLDocument(strXML)

    If InStr(",ZLHIS_PATIENT_002,ZLHIS_PATIENT_012,", "," & strMsgNo & ",") > 0 Then
        strAreaNode = "in_area_id"
        strDeptNode = "in_dept_id"
    ElseIf InStr(",ZLHIS_PATIENT_009,ZLHIS_PATIENT_010,ZLHIS_PATIENT_012,", "," & strMsgNo & ",") > 0 Then
        strAreaNode = "out_area_id"
        strDeptNode = "out_dept_id"
    ElseIf strMsgNo = "ZLHIS_PATIENT_006" Then
        strAreaNode = "before_area_id"
        strDeptNode = "before_dept_id"
    ElseIf strMsgNo = "ZLHIS_PATIENT_003" And int���� = 3 Then
        strAreaNode = "change_area_id"
        strDeptNode = "change_dept_id"
    Else
        Exit Function
    End If

    Call objXML.GetSingleNodeValue(strAreaNode, strTmp1)  '����id
    Call objXML.GetSingleNodeValue(strDeptNode, strTmp2)  '����id
    
    '���ڲ����Ϳ���֮������Ŷ�Ӧ��ϵ������Ҫ���⴦����
    If int���� = 2 Then
        '��������ʾ���ÿ��ҽ���жϣ����ҽ���Ǳش���
        If intView = 0 And Val(strTmp2) = lngDeptID Then blnTmp = True
        
        If Not blnTmp And intView = 1 And Val(strTmp1) <> 0 Then     '����ǰ�������ʾ
            strSql = "Select ����id as id From �������Ҷ�Ӧ Where ����id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, RecMsgToBub, Val(lngDeptID))
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    If InStr("," & strTmp1 & "," & strTmp2 & ",", "," & rsTmp!ID & ",") > 0 Then blnTmp = True: Exit For
                    rsTmp.MoveNext
                Next
            End If
        End If
    ElseIf int���� = 3 Then
        '���ж��Ƿ��ǵ�ǰ����������ͨ�� ���Ҷ�Ӧ�Ĳ����������ж�
        If Val(strTmp1) = lngDeptID Then blnTmp = True
        
        If Not blnTmp And Val(strTmp2) <> 0 Then
            strSql = "Select ����id as id From �������Ҷ�Ӧ Where ����id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, RecMsgToBub, Val(strTmp2))
            If Not rsTmp.EOF Then
                strTmp1 = "," & strTmp1 & "," & strTmp2 & ","
                For i = 1 To rsTmp.RecordCount
                    If InStr(strTmp1, "," & rsTmp!ID & ",") > 0 Then blnTmp = True: Exit For
                    rsTmp.MoveNext
                Next
            End If
        End If
    End If
    
    If Not blnTmp Then Exit Function
    
    strTmp1 = "": strTmp2 = ""
    Call objXML.GetSingleNodeValue("patient_id", strTmp1)   '����id
    Call objXML.GetSingleNodeValue("page_id", strTmp2)      '��ҳid
    strPar = strTmp1 & "," & strTmp2
    
    Call objXML.GetSingleNodeValue("patient_name", strTmp1) '����
    Call objXML.GetSingleNodeValue("in_number", strTmp2)    'סԺ��
    
    'ƴ����������strMsg��
    strMsg = "������" & strTmp1: strTmp1 = ""
    
    If strMsgNo = "ZLHIS_PATIENT_002" Or (strMsgNo = "ZLHIS_PATIENT_012" And strAreaNode = "in_area_id") Or (strMsgNo = "ZLHIS_PATIENT_003" And int���� = 3) Then
        
        strSql = "Select a.סԺ��, a.����, a.�Ա�, a.����, a.��ǰ���� As ����, a.���� From ������Ϣ A Where a.����id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "RecMsgToBub", Val(Split(strPar, ",")(0)))
       
        strMsg = strMsg & IIf(rsTmp!���� & "" = "", "", "�����䣺" & rsTmp!����)
        
        strMsg = strMsg & IIf(rsTmp!�Ա� & "" = "", "", "���Ա�" & rsTmp!�Ա�)
        
        strMsg = strMsg & "��סԺ�ţ�" & strTmp2: strTmp2 = ""
        
        'ĿǰstrMsg��ʽ  ������XXX�����䣺XXX���Ա�XXX��סԺ�ţ�XXX
        If strMsgNo = "ZLHIS_PATIENT_003" And int���� = 3 Then
            strTitle = "סԺ����ת������֪ͨ"
            strMsg = "���µĴ���ס���ˣ�" & strMsg & "��"
        Else
            If strMsgNo = "ZLHIS_PATIENT_002" Then
                strTitle = "סԺ������Ժ���֪ͨ"
                strMsg = "���²�����Ժ��ƣ�" & strMsg
            ElseIf strMsgNo = "ZLHIS_PATIENT_012" Then
                strTitle = "סԺ����ת�����֪ͨ"
                strMsg = "���²���ת����ƣ�" & strMsg
            End If
            
            strTmp1 = "": strTmp2 = ""
            Call objXML.GetSingleNodeValue("in_bed", strTmp1)  '��ס����
            Call objXML.GetSingleNodeValue("in_tendgrade", strTmp2) '����ȼ�
            strMsg = strMsg & IIf(strTmp1 = "", "", "����ס������" & strTmp1) & IIf(strTmp2 = "", "", "������ȼ���" & strTmp2)
            
            strTmp1 = "": strTmp2 = ""
            Call objXML.GetSingleNodeValue("in_doctor", strTmp1)  'סԺҽʦ
            Call objXML.GetSingleNodeValue("duty_nurse", strTmp2) '���λ�ʿ
            strMsg = strMsg & IIf(strTmp1 = "", "", "��סԺҽʦ��" & strTmp1) & IIf(strTmp2 = "", "", "�����λ�ʿ��" & strTmp2)
            
            strMsg = strMsg & "��"
        End If
    ElseIf strMsgNo = "ZLHIS_PATIENT_012" And strAreaNode = "out_area_id" Then
        strTitle = "סԺ����ת�����֪ͨ"
        Call objXML.GetSingleNodeValue("out_dept_title", strTmp1)  'ת����������
        strMsg = strMsg & "��סԺ�ţ�" & strTmp2 & "����ת��" & strTmp1 & "���ҡ�"
    ElseIf strMsgNo = "ZLHIS_PATIENT_006" Then
        strTitle = "סԺ���߱������֪ͨ"
        strMsg = strMsg & ",סԺ�ţ�" & strTmp2 & "��������": strTmp1 = ""
        Call objXML.GetSingleNodeValue("cancel_kind", strTmp1) '������ʽ
        strMsg = strMsg & strTmp1 & "��"
    ElseIf strMsgNo = "ZLHIS_PATIENT_009" Then
        strTitle = "סԺ����Ԥ��Ժ֪ͨ"
        strMsg = strMsg & "��סԺ�ţ�" & strTmp2 & "����Ԥ��Ժ��"
    ElseIf strMsgNo = "ZLHIS_PATIENT_010" Then
        strTitle = "סԺ���߳�Ժ֪ͨ"
        strMsg = strMsg & "��סԺ�ţ�" & strTmp2 & "����Ժ��"
    End If
 
    If strTitle <> "" Then Call objMip.ShowMessage(strMsgNo, strMsg, strTitle, "�鿴", strPar)
    
    RecMsgToBub = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SendMsg(ByVal strMsgNo As String, ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'ZLHIS_CIS_020-���߻�������,ZLHIS_CIS_021-��������ҽ��,ZLHIS_CIS_022-��������ҽ��,ZLHIS_CIS_023-������������ҽ��,
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strNode1 As String
    Dim strNode2 As String
    
    If strMsgNo = "ZLHIS_CIS_020" Then
        strNode1 = "consultation_request"
        strNode2 = "request_id"
    ElseIf strMsgNo = "ZLHIS_CIS_021" Then
        strNode1 = "rescue_order"
        strNode2 = "order_id"
    ElseIf strMsgNo = "ZLHIS_CIS_022" Then
        strNode1 = "die_order"
        strNode2 = "order_id"
    ElseIf strMsgNo = "ZLHIS_CIS_023" Then
        strNode1 = "treat_order"
        strNode2 = "order_id"
    End If
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<סԺ��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<�����>���ͣ�S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<������Դ>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<����id>���ͣ�N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<�������id>���ͣ�N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<��������>���ͣ�S
        .AppendNode "patient_clinic", True
        .AppendNode strNode1 ', True'���ڵ�[��������]
        .AppendData strNode2, arrInput(Index): Index = Index + 1 '<ҽ��id>���ͣ�N
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<ִ�п���id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "execute_dept_title", arrInput(Index) '<ִ�п�������>���ͣ�S
        Index = Index + 1
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<���ݺ���>���ͣ�S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<��������>���ͣ�N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<����ҽ��>���ͣ�S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<����ʱ��>���ͣ�D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<��������id>���ͣ�N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "create_dept_title", arrInput(Index) '<������������>���ͣ�S
        Index = Index + 1
        .AppendNode strNode1, True
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    SendMsg = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_PATIENT_003(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'סԺ����ת������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "in_patient" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "page_id", arrInput(Index): Index = Index + 1 '<��ҳid>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_sex", arrInput(Index) '<�Ա�>���ͣ�S
        Index = Index + 1
        .AppendData "in_number", arrInput(Index): Index = Index + 1 '<סԺ��>���ͣ�S
        .AppendNode "in_patient", True
        .AppendNode "current_state" ', True'���ڵ�[ת����Ϣ]
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "current_area_id", arrInput(Index) '<ת������id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "current_area_title", arrInput(Index) '<ת������>���ͣ�S
        Index = Index + 1
        .AppendData "current_dept_id", arrInput(Index): Index = Index + 1 '<ת������id>���ͣ�N
        .AppendData "current_dept_title", arrInput(Index): Index = Index + 1 '<ת������>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "current_room", arrInput(Index) '<ת������>���ͣ�S
        Index = Index + 1
        .AppendData "current_bed", arrInput(Index): Index = Index + 1 '<ת������>���ͣ�S
        .AppendNode "current_state", True
        .AppendNode "change_state" ', True'���ڵ�[ת����Ϣ]
        .AppendData "change_id", arrInput(Index): Index = Index + 1 '<ת�Ʊ��id>���ͣ�N
        .AppendData "change_date", arrInput(Index): Index = Index + 1 '<���ʱ��>���ͣ�D
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_area_id", arrInput(Index) '<ת�벡��id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_area_title", arrInput(Index) '<ת�벡��>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_dept_id", arrInput(Index) '<ת�����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_dept_title", arrInput(Index) '<ת�����>���ͣ�S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "order_id", arrInput(Index) '<ת��ҽ��id>���ͣ�N
        Index = Index + 1
        .AppendNode "change_state", True
        
        strMsgNo = "ZLHIS_PATIENT_003"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_PATIENT_003 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_PATIENT_005(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'סԺ���߲�����
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "in_patient" ', True'���ڵ�[������Ϣ]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "page_id", arrInput(Index): Index = Index + 1 '<��ҳid>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_sex", arrInput(Index) '<�Ա�>���ͣ�S
        Index = Index + 1
        .AppendData "in_number", arrInput(Index): Index = Index + 1 '<סԺ��>���ͣ�S
        .AppendNode "in_patient", True
        .AppendNode "current_state" ', True'���ڵ�[��ǰ���]
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "current_area_id", arrInput(Index) '<��ǰ����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "current_area_title", arrInput(Index) '<��ǰ����>���ͣ�S
        Index = Index + 1
        .AppendData "current_dept_id", arrInput(Index): Index = Index + 1 '<��ǰ����id>���ͣ�N
        .AppendData "current_dept_title", arrInput(Index): Index = Index + 1 '<��ǰ����>���ͣ�S
        .AppendData "current_situation", arrInput(Index): Index = Index + 1 '<��ǰ����>���ͣ�S
        .AppendNode "current_state", True
        .AppendNode "change_state" ', True'���ڵ�[]
        .AppendData "change_id", arrInput(Index): Index = Index + 1 '<���id>���ͣ�N
        .AppendData "change_date", arrInput(Index): Index = Index + 1 '<���ʱ��>���ͣ�D
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_situation", arrInput(Index) '<�������>���ͣ�S
        Index = Index + 1
        .AppendData "change_operator", arrInput(Index): Index = Index + 1 '<�������Ա>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "order_id", arrInput(Index) '<ҽ��id>���ͣ�N
        Index = Index + 1
        .AppendNode "change_state", True
        
        strMsgNo = "ZLHIS_PATIENT_005"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_PATIENT_005 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_PATIENT_006(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'סԺ���߱䶯����
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "in_patient" ', True'���ڵ�[]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "page_id", arrInput(Index): Index = Index + 1 '<��ҳid>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_sex", arrInput(Index) '<�Ա�>���ͣ�S
        Index = Index + 1
        .AppendData "in_number", arrInput(Index): Index = Index + 1 '<סԺ��>���ͣ�S
        .AppendNode "in_patient", True
        .AppendNode "change_cancel" ', True'���ڵ�[]
        .AppendData "change_id", arrInput(Index): Index = Index + 1 '<�䶯id>���ͣ�N
        .AppendData "cancel_kind", arrInput(Index): Index = Index + 1 '<������ʽ>���ͣ�S
        .AppendData "before_area_id", arrInput(Index): Index = Index + 1 '<�����䶯ǰ����id>���ͣ�N
        .AppendData "before_dept_id", arrInput(Index): Index = Index + 1 '<�����䶯ǰ����Id>���ͣ�N
        .AppendData "after_area_id", arrInput(Index): Index = Index + 1 '<�����䶯����id>���ͣ�N
        .AppendData "after_dept_id", arrInput(Index): Index = Index + 1 '<�����䶯�����id>���ͣ�N
        .AppendNode "change_cancel", True
        
        strMsgNo = "ZLHIS_PATIENT_006"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_PATIENT_006 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLHIS_PATIENT_009(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'סԺ����Ԥ��Ժ
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "in_patient" ', True'���ڵ�[]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<����id>���ͣ�N
        .AppendData "page_id", arrInput(Index): Index = Index + 1 '<��ҳid>���ͣ�N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_sex", arrInput(Index) '<�Ա�>���ͣ�S
        Index = Index + 1
        .AppendData "in_number", arrInput(Index): Index = Index + 1 '<סԺ��>���ͣ�S
        .AppendNode "in_patient", True
        .AppendNode "out_prehospital" ', True'���ڵ�[����Ԥ��Ժ]
        .AppendData "change_id", arrInput(Index): Index = Index + 1 '<���id>���ͣ�N
        .AppendData "out_date", arrInput(Index): Index = Index + 1 '<Ԥ��Ժʱ��>���ͣ�D
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_area_id", arrInput(Index) '<��ǰ����id>���ͣ�N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_area_title", arrInput(Index) '<��ǰ����>���ͣ�S
        Index = Index + 1
        .AppendData "out_dept_id", arrInput(Index): Index = Index + 1 '<��ǰ����>���ͣ�N
        .AppendData "out_dept_title", arrInput(Index): Index = Index + 1 '<��ǰ����id>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_room", arrInput(Index) '<��ǰ����>���ͣ�S
        Index = Index + 1
        .AppendData "out_bed", arrInput(Index): Index = Index + 1 '<��ǰ����>���ͣ�S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "order_id", arrInput(Index) '<ҽ��id>���ͣ�N
        Index = Index + 1
        .AppendNode "out_prehospital", True
        
        strMsgNo = "ZLHIS_PATIENT_009"
        
        If Not (objMip Is Nothing) Then
            If objMip.IsConnect Then
                Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
            End If
        End If
        
        If strXML = "" Then strXML = .XmlText
        
        Call zlDatabase.SendMsg(strMsgNo, strXML)
        
    End With
    ZLHIS_PATIENT_009 = strXML
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatChange(ByVal lngҽ��ID As Long, ByVal intԭ�� As Integer, ByRef lng�䶯id As Long, ByRef str���� As String) As String
'���ܣ���ȡָ��ҽ��ʱ�����ı䶯��¼
'���أ�id �� ���˲���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    lng�䶯id = 0: str���� = ""
    strSql = "Select b.Id, b.���� From ����ҽ����¼ A, ���˱䶯��¼ B Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.��ʼʱ�� = a.��ʼִ��ʱ�� And b.��ʼԭ�� = [2] And a.Id = [1]"
    If intԭ�� = 3 Then
        strSql = "Select b.Id, b.���� From ����ҽ����¼ A, ���˱䶯��¼ B Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.��ʼʱ�� is null and b.��ֹʱ�� is null And b.��ʼԭ�� = [2] And a.Id = [1]"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISMsg", lngҽ��ID, intԭ��)
    If Not rsTmp.EOF Then lng�䶯id = rsTmp!ID: str���� = rsTmp!���� & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���˵�ǰ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ���ȡ���˵�ǰ�Ĳ���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select ���� From (Select ���� From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] Order By ��ʼʱ�� Desc) Where Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISMsg", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then Get���˵�ǰ���� = rsTmp!���� & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



