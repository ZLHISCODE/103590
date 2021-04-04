Attribute VB_Name = "mdlCISMsg"
Option Explicit

Public Function ZLHIS_CIS_001(ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'新开患者医嘱
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    Dim lng就诊科室id As Long, str就诊科室名称 As String
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim bln消息平台 As Boolean
    
    On Error GoTo errH
    
    If Not objMip Is Nothing Then
        If objMip.IsConnect Then bln消息平台 = True
    End If
    
    With objXML
        .ClearXmlText
        .AppendNode "patient_info"  '父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic"  '父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<就诊病区>类型：S
        Index = Index + 1
        
        .AppendData "clinic_dept_id", arrInput(Index) '<就诊科室id>类型：N
        lng就诊科室id = arrInput(Index): Index = Index + 1
        str就诊科室名称 = arrInput(Index): Index = Index + 1 '<就诊科室>类型：S
        
        If str就诊科室名称 = "" And bln消息平台 Then
            strSql = "select 名称 from 部门表 where id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISMsg", lng就诊科室id)
            str就诊科室名称 = rsTmp!名称 & ""
        End If
        .AppendData "clinic_dept_title", str就诊科室名称
        
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "new_order" '父节点[病人医嘱]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "order_urgency", arrInput(Index): Index = Index + 1 '<紧急标志>类型：N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<医嘱期效>类型：N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<主要医嘱类别>类型：S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<主要医嘱操作类型>类型：S
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<开单医生>类型：S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<开单时间>类型：D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<开单科室id>类型：N
        .AppendData "create_dept_title", arrInput(Index): Index = Index + 1 '<开单科室名称>类型：S
        .AppendNode "new_order", True
        
        strMsgNo = "ZLHIS_CIS_001"
     
        If bln消息平台 Then Call objMip.CommitMessage(strMsgNo, .XmlText, strXML)
        
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
'停止患者医嘱
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<就诊病区>类型：S
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_dept_title", arrInput(Index) '<科室名称>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "stop_order" ', True'父节点[病人医嘱]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<医嘱期效>类型：N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<主要医嘱类别>类型：S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<主要医嘱操作类型>类型：S
        .AppendData "stop_doctor", arrInput(Index): Index = Index + 1 '<停嘱医生>类型：S
        .AppendData "stop_time", arrInput(Index): Index = Index + 1 '<停嘱时间>类型：D
        .AppendData "order_urgency", arrInput(Index): Index = Index + 1 '<是否是紧急医嘱>类型：N '0-普通,1-紧急，2-补录
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
'作废患者医嘱
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<就诊病区>类型：S
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "cancel_order" ', True'父节点[病人医嘱]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<医嘱期效>类型：N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<主要医嘱类别>类型：S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<主要医嘱操作类型>类型：S
        .AppendData "execute_kind", arrInput(Index): Index = Index + 1 '<执行分类>类型：N
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<执行科室id>类型：N
        .AppendData "cancel_doctor", arrInput(Index): Index = Index + 1 '<作废医生>类型：S
        .AppendData "cancel_time", arrInput(Index): Index = Index + 1 '<作废时间>类型：D
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
'医嘱申请安排
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "arrange_order" ', True'父节点[安排医嘱]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<医嘱期效>类型：N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<主要医嘱类别>类型：S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<主要医嘱操作类型>类型：S
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<发送批号>类型：N
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<执行科室id>类型：N
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
'医技执行安排完成
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" '', True'父节点'<病人信息>
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" '', True'父节点'<就诊信息>
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "patient_order" '', True'父节点'<病人医嘱>
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<医嘱期效>类型：N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<主要医嘱类别>类型：S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<主要医嘱操作类型>类型：S
        .AppendData "order_item_id", arrInput(Index): Index = Index + 1 '<医嘱项目id>类型：N
        .AppendData "order_item_title", arrInput(Index): Index = Index + 1 '<医嘱项目>类型：S
        .AppendNode "patient_order", True
        .AppendNode "arrange_result" ', True'父节点'<安排情况>
        .AppendData "arrange_time", arrInput(Index): Index = Index + 1 '<安排时间>类型：D
        .AppendData "arrange_room", arrInput(Index): Index = Index + 1 '<安排房间>类型：S
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
'患者药品医嘱发送
    Dim objXML As New zl9ComLib.clsXML
    Dim rsTmp As ADODB.Recordset
    Dim strXML As String
    Dim Index As Integer
    Dim i As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "drug_bill" ', True'父节点[]
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<发送批号>类型：N
        .AppendData "send_time", arrInput(Index): Index = Index + 1 '<发送时间>类型：D
        .AppendData "send_person", arrInput(Index): Index = Index + 1 '<发送人员>类型：S
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<单据号码>类型：S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<单据性质>类型：N
        .AppendData "charge_state", arrInput(Index): Index = Index + 1 '<收费状态>类型：N
        .AppendNode "send_order" ', True'用记录集或字符串 '父节点[发送医嘱]
        Set rsTmp = arrInput(Index)
        For i = 1 To rsTmp.RecordCount
            .AppendData "order_id", rsTmp!医嘱ID '<医嘱id>类型：N
            .AppendData "order_relevant_id", rsTmp!相关ID '<相关ID>类型：N
            .AppendData "order_info", rsTmp!医嘱内容 & "" '<医嘱内容>类型：S
            .AppendData "order_rate", rsTmp!执行频率 & ""  '<执行频率>类型：S
            .AppendData "order_route_id", rsTmp!给药途径ID '<给药途径id>类型：N
            .AppendData "order_route", rsTmp!给药途径 & ""  '<给药途径>类型：S
            .AppendData "order_starttime", rsTmp!开始时间 & ""  '<开始时间>类型：D
            .AppendData "order_single", Format(Val(rsTmp!单量 & ""), "0.0") '<单量>类型：N
            .AppendData "order_total", Format(Val(rsTmp!总量 & ""), "0.0")  '<总量>类型：N
            .AppendData "order_entrust", rsTmp!医嘱嘱托 & ""  '<医嘱嘱托>类型：S
            .AppendData "order_item_id", rsTmp!品种ID '<品种id>类型：N
            .AppendData "Drug_item_kind", rsTmp!药品类别 & ""  '<药品类别>类型：S
            .AppendData "Drug_item_id", rsTmp!药品ID '<药品id>类型：N
            .AppendData "execute_dept_id", rsTmp!执行部门ID '<执行部门id>类型：N
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
'门诊患者转诊
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
        .AppendNode "patient_info" '', True'父节点'<病人信息>
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" '', True'父节点'<就诊信息>
        .AppendData "clinic_id", arrInput(Index): Index = Index + 1 '<就诊挂号id>类型：N
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_dept_title", arrInput(Index) '<科室名称>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "change_clinic" '', True'父节点'<转诊信息>
        .AppendData "change_dept_id", arrInput(Index): Index = Index + 1 '<转诊科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_dept_title", arrInput(Index) '<转诊科室名称>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_doctor_id", arrInput(Index) '<转诊医生id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_doctor", arrInput(Index) '<转诊医生>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_room", arrInput(Index) '<转诊诊室>类型：S
        Index = Index + 1
        .AppendData "change_exedoctor", arrInput(Index): Index = Index + 1 '<转诊执行医生>类型：S
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
'输液批次调整
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<主页id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<就诊病区名称>类型：S
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "order_info" ', True'父节点[医嘱内容]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<医嘱期效>类型：N
        .AppendNode "order_info", True
        .AppendNode "transfusion_info" ', True'父节点[输液信息]
        .AppendData "transfusion_id", arrInput(Index): Index = Index + 1 '<输液id>类型：N
        .AppendData "transfusion_batch", arrInput(Index): Index = Index + 1 '<输液批次>类型：N
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
'门诊患者接诊
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lng病人ID  As Long
    Dim strTmp As String
    Dim strMsgNo As String
    
    On Error GoTo errH
    If Not (objMip Is Nothing) Then
        If objMip.IsConnect Then
            Exit Function
        End If
    End If
    strSql = "Select 身份证号, 住院号, 就诊卡号, 健康号, 医保号, 性别, 年龄, 出生日期, 婚姻状况, 费别, 国籍, 民族, 职业, 学历, 家庭地址, 家庭电话, 家庭地址邮编 As 家庭邮编, 联系人电话," & _
        " 联系人关系, 联系人地址, 工作单位, 单位电话, 单位邮编, 其他证件 As 其它证件, 出生地点, 监护人, 户口地址 As 户籍地址, 户口地址邮编 As 户籍邮编, 籍贯, 医疗付款方式" & _
        " From 病人信息 Where 病人id = [1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISMsg", lng病人ID)  'If Not rsTmp.EOF Then '不用做这个判断，如果有错应该暴露出来
    If rsTmp.EOF Then Exit Function '可能门诊病人没建档
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" '', True'父节点'<病人信息>
        .AppendData "patient_id", arrInput(Index) '<病人id>类型：N
            lng病人ID = Val(arrInput(Index)): Index = Index + 1
         If TypeName(arrInput(Index)) <> "Error" Then
            .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Else
            .AppendData "patient_name", ""
        End If
        Index = Index + 1
        .AppendData "identity_card", "" & rsTmp!身份证号
        .AppendData "in_number", "" & rsTmp!住院号
        If TypeName(arrInput(Index)) <> "Error" Then
            .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Else
            .AppendData "out_number", ""
        End If
        Index = Index + 1
        .AppendData "visit_number", "" & rsTmp!就诊卡号
        .AppendData "health_number", "" & rsTmp!健康号
        .AppendData "medical_number", "" & rsTmp!医保号
        .AppendData "patient_sex", "" & rsTmp!性别
        .AppendData "patient_age", "" & rsTmp!年龄
        strTmp = Decode("" & rsTmp!性别, "男", 1, "女", 2, "未知", 3, "")
        .AppendData "patient_age_code", strTmp ' 性别编码
        .AppendData "patient_birthday", "" & rsTmp!出生日期
        .AppendData "patient_marriage", "" & rsTmp!婚姻状况
        strTmp = Decode("" & rsTmp!婚姻状况, "未婚", 1, "已婚", 2, "丧偶", 3, "离婚", 4, "其他", 9, "")
        .AppendData "patient_marriage_code", strTmp ' 婚姻状况编码
        .AppendData "patient_chargetype", "" & rsTmp!费别
        .AppendData "patient_nationality", "" & rsTmp!国籍
        .AppendData "patient_nation", "" & rsTmp!民族
        .AppendData "patient_profession", "" & rsTmp!职业
        .AppendData "patient_education", "" & rsTmp!学历
        .AppendData "home_addr", "" & rsTmp!家庭地址
        .AppendData "home_telephone", "" & rsTmp!家庭电话
        .AppendData "home_postcode", "" & rsTmp!家庭邮编
        .AppendData "contact_telephone", "" & rsTmp!联系人电话
        .AppendData "contact_relation", "" & rsTmp!联系人关系
        .AppendData "contact_addr", "" & rsTmp!联系人地址
        .AppendData "patient_work", "" & rsTmp!工作单位
        .AppendData "work_telephone", "" & rsTmp!单位电话
        .AppendData "work_addr", "" ' 单位地址"--这个字段暂时没有
        .AppendData "work_postcode", "" & rsTmp!单位邮编
        .AppendData "patient_other_papers", "" & rsTmp!其它证件
        .AppendData "patient_birthplace", "" & rsTmp!出生地点
        .AppendData "patient_guardian", "" & rsTmp!监护人
        .AppendData "patient_height", arrInput(Index): Index = Index + 1 '身高"
        .AppendData "patient_weight", arrInput(Index): Index = Index + 1 ' 体重"
        .AppendData "residence_addr", "" & rsTmp!户籍地址
        .AppendData "residence_postcode", "" & rsTmp!户籍邮编
        .AppendData "native_place", "" & rsTmp!籍贯
        .AppendData "patient_payment", "" & rsTmp!医疗付款方式
        strTmp = Decode("" & rsTmp!医疗付款方式, "社会基本医疗保险", 1, "公费医疗", 2, "大病统筹", 3, "商业保险", 4, "自费医疗", 5, "其他", 6, "")
        .AppendData "patient_payment_code", strTmp ' 付款方式编码
        .AppendNode "patient_info", True
        .AppendNode "receive_clinic" '', True'父节点'<接诊信息>
        .AppendData "clinic_id", arrInput(Index): Index = Index + 1 '<挂号id>类型：N
        .AppendData "return_visit", arrInput(Index): Index = Index + 1 '<是否复诊>类型：N
        .AppendData "emergency_treatment", arrInput(Index): Index = Index + 1 '<是否急诊>类型：N
        .AppendData "receive_time", arrInput(Index): Index = Index + 1 '<接诊时间>类型：D
        .AppendData "receive_dept_id", arrInput(Index): Index = Index + 1 '<接诊科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then
            .AppendData "receive_dept_title", arrInput(Index) '<接诊科室名称>类型：S
        Else
            .AppendData "receive_dept_title", ""
        End If
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then
            .AppendData "receive_room", arrInput(Index) '<接诊诊室>类型：S
        Else
            .AppendData "receive_room", ""
        End If
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "receive_doctor", arrInput(Index) '<接诊医生>类型：S
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
'下达患者诊断
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1   '<病人来源>类型：N
        .AppendData "clinic_id", arrInput(Index): Index = Index + 1 '<就诊id>类型：N
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendNode "patient_clinic", True
        .AppendNode "diagnose_info" ', True'父节点[]
        .AppendData "diagnose_id", Val(arrInput(Index)): Index = Index + 1 '<诊断id>类型：N
        .AppendData "diagnose_kind", arrInput(Index): Index = Index + 1 '<诊断类型>类型：S
        .AppendData "diagnose_question", arrInput(Index): Index = Index + 1 '<是否疑诊>类型：N
        .AppendData "diagnose_serial", arrInput(Index): Index = Index + 1 '<诊断次序>类型：N
        .AppendData "diagnose_code", arrInput(Index): Index = Index + 1 '<诊断编码>类型：S
        .AppendData "illness_code", arrInput(Index): Index = Index + 1 '<疾病编码>类型：S
        .AppendData "illness_addition_code", arrInput(Index): Index = Index + 1 '<疾病附码>类型：S
        .AppendData "illness_kind", arrInput(Index): Index = Index + 1 '<疾病类别>类型：S
        .AppendData "syndrome_code", arrInput(Index): Index = Index + 1 '<证候编码>类型：S
        .AppendData "syndrome_title", arrInput(Index): Index = Index + 1 '<证候名称>类型：S
        .AppendData "record_time", Format(arrInput(Index), "YYYY-MM-DD HH:MM:SS"): Index = Index + 1 '<记录日期>类型：D
        .AppendData "record_person", arrInput(Index): Index = Index + 1 '<记录人员>类型：S
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
'撤消患者诊断
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        .AppendData "clinic_id", arrInput(Index): Index = Index + 1 '<就诊id>类型：N
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendNode "patient_clinic", True
        .AppendNode "diagnose_info" ', True'父节点[]
        .AppendData "diagnose_id", arrInput(Index): Index = Index + 1 '<诊断id>类型：N
        .AppendData "diagnose_code", arrInput(Index): Index = Index + 1 '<诊断编码>类型：S
        .AppendData "illness_code", arrInput(Index): Index = Index + 1 '<疾病编码>类型：S
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
'住院患者输液销帐申请
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<主页id>类型：N
        Index = Index + 1
        .AppendData "clinic_area_id", arrInput(Index): Index = Index + 1 '<就诊病区id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<就诊病区名称>类型：S
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_dept_title", arrInput(Index) '<就诊科室名称>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "cancel_reqeust" ', True'父节点[]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "transfusion_id", arrInput(Index): Index = Index + 1 '<配液id>类型：N
        .AppendData "request_time", arrInput(Index): Index = Index + 1 '<申请时间>类型：D
        .AppendData "request_person", arrInput(Index): Index = Index + 1 '<申请人员>类型：S
        .AppendData "request_dept_id", arrInput(Index): Index = Index + 1 '<申请科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "request_dept_title", arrInput(Index) '<申请科室>类型：S
        Index = Index + 1
        .AppendData "audit_dept_id", arrInput(Index): Index = Index + 1 '<审核部门id>类型：N
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
'医技拒绝执行通知
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" '', True'父节点'<病人信息>
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index)  '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" '', True'父节点'<就诊信息>
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "refuse_order" '', True'父节点'<>
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<医嘱期效>类型：N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<主要医嘱类别>类型：S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<主要医嘱操作类型>类型：S
        .AppendData "order_item_id", arrInput(Index): Index = Index + 1 '<医嘱项目id>类型：N
        .AppendData "order_item_title", arrInput(Index): Index = Index + 1 '<医嘱项目>类型：S
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
'患者检验申请
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
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_dept_title", arrInput(Index) '<科室名称>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "lis_request" ', True'父节点[检验申请]
        .AppendData "request_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "lis_sample", arrInput(Index): Index = Index + 1 '<检验标本>类型：S
        .AppendData "collect_item_id", arrInput(Index): Index = Index + 1 '<采集方式id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "collect_item_title", arrInput(Index) '<采集方式名称>类型：S
        Index = Index + 1
        .AppendData "collect_dept_id", arrInput(Index): Index = Index + 1 '<采集科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "collect_dept_title", arrInput(Index) '<采集科室名称>类型：S
        Index = Index + 1
        '检验项目 lis_item_title结点暂不产生，但是位置是留着的
        strLine = arrInput(Index): Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then strLine1 = arrInput(Index)
        Index = Index + 1
'        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "lis_item_title", arrInput(Index): Index = Index + 1 '<检验项目名称>类型：S
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
        
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<执行科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "execute_dept_title", arrInput(Index) '<执行科室名称>类型：S
        Index = Index + 1
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<发送批号>类型：N
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<单据号码>类型：S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<单据性质>类型：N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<开单医生>类型：S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<开单时间>类型：D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<开单科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "create_dept_title", arrInput(Index) '<开单科室名称>类型：S
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
'患者检查申请
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        .AppendNode "patient_clinic", True
        .AppendNode "check_request" ', True'父节点[检查申请]
        .AppendData "request_id", arrInput(Index): Index = Index + 1 '<申请id>类型：N
        .AppendData "check_item_id", arrInput(Index): Index = Index + 1 '<医嘱项目id>类型：N
        .AppendData "check_item_title", arrInput(Index): Index = Index + 1 '<医嘱项目名称>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "check_parts", arrInput(Index) '<部位清单>类型：S
        Index = Index + 1
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<执行科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "execute_dept_title", arrInput(Index) '<执行科室名称>类型：S
        Index = Index + 1
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<发送批号>类型：N
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<单据号码>类型：S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<单据性质>类型：N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<开单医生>类型：S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<开单时间>类型：D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<开单科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "create_dept_title", arrInput(Index) '<开单科室名称>类型：S
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
'患者手术申请
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
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index)
        Index = Index + 1 '<住院号>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index)
        Index = Index + 1 '<门诊号>类型：S
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index)
        Index = Index + 1 '<就诊id>类型：N
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        .AppendNode "patient_clinic", True
        .AppendNode "oper_request" ', True'父节点[手术申请]
        .AppendData "request_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        
        strLine = arrInput(Index): Index = Index + 1 '<手术项目id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then strLine1 = arrInput(Index) '<手术项目名称>类型：S，此结点暂不添加
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
        If TypeName(arrInput(Index)) <> "Error" Then strLine = arrInput(Index)  '<麻醉项目id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then strLine1 = arrInput(Index)  '<麻醉项目名称>类型：S，此结点暂不添加
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
        If TypeName(arrInput(Index)) <> "Error" Then strLine = arrInput(Index) '<主刀医生>类型：S
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
        
        If TypeName(arrInput(Index)) <> "Error" Then strLine = arrInput(Index) '<助手医生>类型：S
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
         
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<执行科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "execute_dept_title", arrInput(Index)
        Index = Index + 1 '<执行科室名称>类型：S
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<发送批号>类型：N
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<单据号码>类型：S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<单据性质>类型：N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<开单医生>类型：S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<开单时间>类型：D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<开单科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "create_dept_title", arrInput(Index) '<开单科室名称>类型：S
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
'患者输血申请
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        .AppendNode "patient_clinic", True
        .AppendNode "blood_request" ', True'父节点[手术申请]
        .AppendData "request_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<执行科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "execute_dept_title", arrInput(Index) '<执行科室名称>类型：S
        Index = Index + 1
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<发送批号>类型：N
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<单据号码>类型：S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<单据性质>类型：N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<开单医生>类型：S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<开单时间>类型：D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<开单科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "create_dept_title", arrInput(Index) '<开单科室名称>类型：S
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
'患者医嘱撤消 门诊调用则是作废操作，住院调用则回退操作
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        .AppendNode "patient_clinic", True
        .AppendNode "cancel_order" ', True'父节点[]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<主要医嘱类别>类型：S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<主要医嘱操作类型>类型：S
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
'功能：危机值消息阅读后发送的消息，仅消息平台连接可用时才会发送
'参数：strMsgNo ZLHIS_CIS_025-检查危急值阅读通知；ZLHIS_CIS_014－检验危急值阅读通知
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>S
        .AppendData "clinic_id", arrInput(Index): Index = Index + 1  '<主页id，挂号ID>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "critical_read" ', True'父节点[]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
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
'新嘱审核提醒，实习医生下达的医嘱需要正式医生审核
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
 
    On Error GoTo errH
    
    With objXML
        .ClearXmlText
        .AppendNode "patient_info"  '父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic"  '父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<就诊病区>类型：S
        Index = Index + 1
        
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<就诊科室>类型：S
        
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "new_order" '父节点[病人医嘱]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "order_urgency", arrInput(Index): Index = Index + 1 '<紧急标志>类型：N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<医嘱期效>类型：N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<主要医嘱类别>类型：S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<主要医嘱操作类型>类型：S
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<开单医生>类型：S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<开单时间>类型：D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<开单科室id>类型：N
        .AppendData "create_dept_title", arrInput(Index): Index = Index + 1 '<开单科室名称>类型：S
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
'新停审核提醒
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<就诊病区>类型：S
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_dept_title", arrInput(Index) '<科室名称>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "stop_order" ', True'父节点[病人医嘱]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "order_expiry", arrInput(Index): Index = Index + 1 '<医嘱期效>类型：N
        .AppendData "order_kind", arrInput(Index): Index = Index + 1 '<主要医嘱类别>类型：S
        .AppendData "operation_kind", arrInput(Index): Index = Index + 1 '<主要医嘱操作类型>类型：S
        .AppendData "stop_doctor", arrInput(Index): Index = Index + 1 '<停嘱医生>类型：S
        .AppendData "stop_time", arrInput(Index): Index = Index + 1 '<停嘱时间>类型：D
        .AppendData "order_urgency", arrInput(Index): Index = Index + 1 '<是否是紧急医嘱>类型：N '0-普通,1-紧急，2-补录
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
'输血配血申请
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
 
    On Error GoTo errH
    
    With objXML
        .ClearXmlText
        .AppendNode "patient_info"  '父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic"  '父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<就诊病区>类型：S
        Index = Index + 1
        
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<就诊科室>类型：S
        
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "new_order" '父节点[病人医嘱]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<开单医生>类型：S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<开单时间>类型：D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<开单科室id>类型：N
        .AppendData "create_dept_title", arrInput(Index): Index = Index + 1 '<开单科室名称>类型：S
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
'功能：医嘱审核消息发送
'包含这几个 'ZLHIS_CIS_028－手术审核提醒；ZLHIS_CIS_029-抗菌药物审核提醒；ZLHIS_CIS_030 -输血审核提醒

    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer

    On Error GoTo errH
    
    With objXML
        .ClearXmlText
        .AppendNode "patient_info"  '父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_name", arrInput(Index) '<姓名>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic"  '父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_id", arrInput(Index) '<就诊病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_area_title", arrInput(Index) '<就诊病区>类型：S
        Index = Index + 1
        
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<就诊科室>类型：S
        
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_room", arrInput(Index) '<就诊病房>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_bed", arrInput(Index) '<就诊病床>类型：S
        Index = Index + 1
        .AppendNode "patient_clinic", True
        .AppendNode "new_order" '父节点[病人医嘱]
        .AppendData "order_id", arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<开单医生>类型：S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<开单时间>类型：D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<开单科室id>类型：N
        .AppendData "create_dept_title", arrInput(Index): Index = Index + 1 '<开单科室名称>类型：S
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

Public Function RecMsgToBub(ByRef objMip As zl9ComLib.clsMipModule, ByVal lngDeptID As Long, ByVal int场合 As Integer, ByVal strMsgNo As String, ByVal strXML As String, Optional ByVal intView As Integer) As Boolean
'功能：冒泡提醒
'参数：
'      lngDeptID 如果是护士站调用为病区id，医生站调用时则跟据 intView 进行判断是病区id还是科室id
'      int场合 1－门诊医生站，2－住院医生站，3－住院护士站，4－医技站
'      intView 显示方式，0-按科室显示，1-按病区显示 只有在住院医生站调用时使用这参数
    Dim objXML As zl9ComLib.clsXML
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strMsg As String '气泡中的内容
    Dim strTitle As String '标题
    Dim strPar As String '链接参数
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
    ElseIf strMsgNo = "ZLHIS_PATIENT_003" And int场合 = 3 Then
        strAreaNode = "change_area_id"
        strDeptNode = "change_dept_id"
    Else
        Exit Function
    End If

    Call objXML.GetSingleNodeValue(strAreaNode, strTmp1)  '病区id
    Call objXML.GetSingleNodeValue(strDeptNode, strTmp2)  '科室id
    
    '由于病区和科室之间存在着对应关系，这里要特殊处理下
    If int场合 = 2 Then
        '按科室显示，用科室结点判断，科室结点是必传的
        If intView = 0 And Val(strTmp2) = lngDeptID Then blnTmp = True
        
        If Not blnTmp And intView = 1 And Val(strTmp1) <> 0 Then     '如果是按病区显示
            strSql = "Select 科室id as id From 病区科室对应 Where 病区id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, RecMsgToBub, Val(lngDeptID))
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    If InStr("," & strTmp1 & "," & strTmp2 & ",", "," & rsTmp!ID & ",") > 0 Then blnTmp = True: Exit For
                    rsTmp.MoveNext
                Next
            End If
        End If
    ElseIf int场合 = 3 Then
        '先判断是否是当前病区，否则通过 科室对应的病区来进行判断
        If Val(strTmp1) = lngDeptID Then blnTmp = True
        
        If Not blnTmp And Val(strTmp2) <> 0 Then
            strSql = "Select 病区id as id From 病区科室对应 Where 科室id = [1]"
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
    Call objXML.GetSingleNodeValue("patient_id", strTmp1)   '病人id
    Call objXML.GetSingleNodeValue("page_id", strTmp2)      '主页id
    strPar = strTmp1 & "," & strTmp2
    
    Call objXML.GetSingleNodeValue("patient_name", strTmp1) '姓名
    Call objXML.GetSingleNodeValue("in_number", strTmp2)    '住院号
    
    '拼接姓名存入strMsg中
    strMsg = "姓名：" & strTmp1: strTmp1 = ""
    
    If strMsgNo = "ZLHIS_PATIENT_002" Or (strMsgNo = "ZLHIS_PATIENT_012" And strAreaNode = "in_area_id") Or (strMsgNo = "ZLHIS_PATIENT_003" And int场合 = 3) Then
        
        strSql = "Select a.住院号, a.姓名, a.性别, a.年龄, a.当前床号 As 床号, a.险类 From 病人信息 A Where a.病人id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "RecMsgToBub", Val(Split(strPar, ",")(0)))
       
        strMsg = strMsg & IIf(rsTmp!年龄 & "" = "", "", "，年龄：" & rsTmp!年龄)
        
        strMsg = strMsg & IIf(rsTmp!性别 & "" = "", "", "，性别：" & rsTmp!性别)
        
        strMsg = strMsg & "，住院号：" & strTmp2: strTmp2 = ""
        
        '目前strMsg格式  姓名：XXX，年龄：XXX，性别：XXX，住院号：XXX
        If strMsgNo = "ZLHIS_PATIENT_003" And int场合 = 3 Then
            strTitle = "住院患者转出科室通知"
            strMsg = "有新的待入住病人，" & strMsg & "。"
        Else
            If strMsgNo = "ZLHIS_PATIENT_002" Then
                strTitle = "住院患者入院入科通知"
                strMsg = "有新病人入院入科，" & strMsg
            ElseIf strMsgNo = "ZLHIS_PATIENT_012" Then
                strTitle = "住院患者转入入科通知"
                strMsg = "有新病人转科入科，" & strMsg
            End If
            
            strTmp1 = "": strTmp2 = ""
            Call objXML.GetSingleNodeValue("in_bed", strTmp1)  '入住病床
            Call objXML.GetSingleNodeValue("in_tendgrade", strTmp2) '护理等级
            strMsg = strMsg & IIf(strTmp1 = "", "", "，入住病床：" & strTmp1) & IIf(strTmp2 = "", "", "，护理等级：" & strTmp2)
            
            strTmp1 = "": strTmp2 = ""
            Call objXML.GetSingleNodeValue("in_doctor", strTmp1)  '住院医师
            Call objXML.GetSingleNodeValue("duty_nurse", strTmp2) '责任护士
            strMsg = strMsg & IIf(strTmp1 = "", "", "，住院医师：" & strTmp1) & IIf(strTmp2 = "", "", "，责任护士：" & strTmp2)
            
            strMsg = strMsg & "。"
        End If
    ElseIf strMsgNo = "ZLHIS_PATIENT_012" And strAreaNode = "out_area_id" Then
        strTitle = "住院患者转入科室通知"
        Call objXML.GetSingleNodeValue("out_dept_title", strTmp1)  '转出科室名称
        strMsg = strMsg & "，住院号：" & strTmp2 & "，已转往" & strTmp1 & "科室。"
    ElseIf strMsgNo = "ZLHIS_PATIENT_006" Then
        strTitle = "住院患者变更撤消通知"
        strMsg = strMsg & ",住院号：" & strTmp2 & "，撤消了": strTmp1 = ""
        Call objXML.GetSingleNodeValue("cancel_kind", strTmp1) '撤消方式
        strMsg = strMsg & strTmp1 & "。"
    ElseIf strMsgNo = "ZLHIS_PATIENT_009" Then
        strTitle = "住院患者预出院通知"
        strMsg = strMsg & "，住院号：" & strTmp2 & "，已预出院。"
    ElseIf strMsgNo = "ZLHIS_PATIENT_010" Then
        strTitle = "住院患者出院通知"
        strMsg = strMsg & "，住院号：" & strTmp2 & "，出院。"
    End If
 
    If strTitle <> "" Then Call objMip.ShowMessage(strMsgNo, strMsg, strTitle, "查看", strPar)
    
    RecMsgToBub = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SendMsg(ByVal strMsgNo As String, ByRef objMip As zl9ComLib.clsMipModule, ParamArray arrInput() As Variant) As String
'ZLHIS_CIS_020-患者会诊申请,ZLHIS_CIS_021-患者抢救医嘱,ZLHIS_CIS_022-患者死亡医嘱,ZLHIS_CIS_023-患者特殊治疗医嘱,
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
        .AppendNode "patient_info" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "in_number", arrInput(Index) '<住院号>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_number", arrInput(Index) '<门诊号>类型：S
        Index = Index + 1
        .AppendNode "patient_info", True
        .AppendNode "patient_clinic" ', True'父节点[就诊信息]
        .AppendData "patient_source", arrInput(Index): Index = Index + 1 '<病人来源>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "clinic_id", arrInput(Index) '<就诊id>类型：N
        Index = Index + 1
        .AppendData "clinic_dept_id", arrInput(Index): Index = Index + 1 '<就诊科室id>类型：N
        .AppendData "clinic_dept_title", arrInput(Index): Index = Index + 1 '<科室名称>类型：S
        .AppendNode "patient_clinic", True
        .AppendNode strNode1 ', True'父节点[会诊申请]
        .AppendData strNode2, arrInput(Index): Index = Index + 1 '<医嘱id>类型：N
        .AppendData "execute_dept_id", arrInput(Index): Index = Index + 1 '<执行科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "execute_dept_title", arrInput(Index) '<执行科室名称>类型：S
        Index = Index + 1
        .AppendData "send_serial", arrInput(Index): Index = Index + 1 '<发送批号>类型：N
        .AppendData "bill_no", arrInput(Index): Index = Index + 1 '<单据号码>类型：S
        .AppendData "bill_kind", arrInput(Index): Index = Index + 1 '<单据性质>类型：N
        .AppendData "create_doctor", arrInput(Index): Index = Index + 1 '<开单医生>类型：S
        .AppendData "create_time", arrInput(Index): Index = Index + 1 '<开单时间>类型：D
        .AppendData "create_dept_id", arrInput(Index): Index = Index + 1 '<开单科室id>类型：N
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "create_dept_title", arrInput(Index) '<开单科室名称>类型：S
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
'住院患者转出科室
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "in_patient" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "page_id", arrInput(Index): Index = Index + 1 '<主页id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_sex", arrInput(Index) '<性别>类型：S
        Index = Index + 1
        .AppendData "in_number", arrInput(Index): Index = Index + 1 '<住院号>类型：S
        .AppendNode "in_patient", True
        .AppendNode "current_state" ', True'父节点[转出信息]
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "current_area_id", arrInput(Index) '<转出病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "current_area_title", arrInput(Index) '<转出病区>类型：S
        Index = Index + 1
        .AppendData "current_dept_id", arrInput(Index): Index = Index + 1 '<转出科室id>类型：N
        .AppendData "current_dept_title", arrInput(Index): Index = Index + 1 '<转出科室>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "current_room", arrInput(Index) '<转出病房>类型：S
        Index = Index + 1
        .AppendData "current_bed", arrInput(Index): Index = Index + 1 '<转出病床>类型：S
        .AppendNode "current_state", True
        .AppendNode "change_state" ', True'父节点[转入信息]
        .AppendData "change_id", arrInput(Index): Index = Index + 1 '<转科变更id>类型：N
        .AppendData "change_date", arrInput(Index): Index = Index + 1 '<变更时间>类型：D
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_area_id", arrInput(Index) '<转入病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_area_title", arrInput(Index) '<转入病区>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_dept_id", arrInput(Index) '<转入科室id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_dept_title", arrInput(Index) '<转入科室>类型：S
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "order_id", arrInput(Index) '<转科医嘱id>类型：N
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
'住院患者病情变更
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "in_patient" ', True'父节点[病人信息]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "page_id", arrInput(Index): Index = Index + 1 '<主页id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_sex", arrInput(Index) '<性别>类型：S
        Index = Index + 1
        .AppendData "in_number", arrInput(Index): Index = Index + 1 '<住院号>类型：S
        .AppendNode "in_patient", True
        .AppendNode "current_state" ', True'父节点[当前情况]
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "current_area_id", arrInput(Index) '<当前病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "current_area_title", arrInput(Index) '<当前病区>类型：S
        Index = Index + 1
        .AppendData "current_dept_id", arrInput(Index): Index = Index + 1 '<当前科室id>类型：N
        .AppendData "current_dept_title", arrInput(Index): Index = Index + 1 '<当前科室>类型：S
        .AppendData "current_situation", arrInput(Index): Index = Index + 1 '<当前病况>类型：S
        .AppendNode "current_state", True
        .AppendNode "change_state" ', True'父节点[]
        .AppendData "change_id", arrInput(Index): Index = Index + 1 '<变更id>类型：N
        .AppendData "change_date", arrInput(Index): Index = Index + 1 '<变更时间>类型：D
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "change_situation", arrInput(Index) '<变更病况>类型：S
        Index = Index + 1
        .AppendData "change_operator", arrInput(Index): Index = Index + 1 '<变更操作员>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "order_id", arrInput(Index) '<医嘱id>类型：N
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
'住院患者变动撤消
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "in_patient" ', True'父节点[]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "page_id", arrInput(Index): Index = Index + 1 '<主页id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_sex", arrInput(Index) '<性别>类型：S
        Index = Index + 1
        .AppendData "in_number", arrInput(Index): Index = Index + 1 '<住院号>类型：S
        .AppendNode "in_patient", True
        .AppendNode "change_cancel" ', True'父节点[]
        .AppendData "change_id", arrInput(Index): Index = Index + 1 '<变动id>类型：N
        .AppendData "cancel_kind", arrInput(Index): Index = Index + 1 '<撤消方式>类型：S
        .AppendData "before_area_id", arrInput(Index): Index = Index + 1 '<撤销变动前病区id>类型：N
        .AppendData "before_dept_id", arrInput(Index): Index = Index + 1 '<撤销变动前科室Id>类型：N
        .AppendData "after_area_id", arrInput(Index): Index = Index + 1 '<撤销变动后病区id>类型：N
        .AppendData "after_dept_id", arrInput(Index): Index = Index + 1 '<撤销变动后科室id>类型：N
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
'住院患者预出院
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim Index As Integer
    Dim strMsgNo As String
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "in_patient" ', True'父节点[]
        .AppendData "patient_id", arrInput(Index): Index = Index + 1 '<病人id>类型：N
        .AppendData "page_id", arrInput(Index): Index = Index + 1 '<主页id>类型：N
        .AppendData "patient_name", arrInput(Index): Index = Index + 1 '<姓名>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "patient_sex", arrInput(Index) '<性别>类型：S
        Index = Index + 1
        .AppendData "in_number", arrInput(Index): Index = Index + 1 '<住院号>类型：S
        .AppendNode "in_patient", True
        .AppendNode "out_prehospital" ', True'父节点[病人预出院]
        .AppendData "change_id", arrInput(Index): Index = Index + 1 '<变更id>类型：N
        .AppendData "out_date", arrInput(Index): Index = Index + 1 '<预出院时间>类型：D
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_area_id", arrInput(Index) '<当前病区id>类型：N
        Index = Index + 1
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_area_title", arrInput(Index) '<当前病区>类型：S
        Index = Index + 1
        .AppendData "out_dept_id", arrInput(Index): Index = Index + 1 '<当前科室>类型：N
        .AppendData "out_dept_title", arrInput(Index): Index = Index + 1 '<当前科室id>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "out_room", arrInput(Index) '<当前病房>类型：S
        Index = Index + 1
        .AppendData "out_bed", arrInput(Index): Index = Index + 1 '<当前病床>类型：S
        If TypeName(arrInput(Index)) <> "Error" Then .AppendData "order_id", arrInput(Index) '<医嘱id>类型：N
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

Public Function GetPatChange(ByVal lng医嘱ID As Long, ByVal int原因 As Integer, ByRef lng变动id As Long, ByRef str病情 As String) As String
'功能：获取指定医嘱时产生的变动记录
'返回：id 和 病人病情
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    lng变动id = 0: str病情 = ""
    strSql = "Select b.Id, b.病情 From 病人医嘱记录 A, 病人变动记录 B Where a.病人id = b.病人id And a.主页id = b.主页id And b.开始时间 = a.开始执行时间 And b.开始原因 = [2] And a.Id = [1]"
    If int原因 = 3 Then
        strSql = "Select b.Id, b.病情 From 病人医嘱记录 A, 病人变动记录 B Where a.病人id = b.病人id And a.主页id = b.主页id And b.开始时间 is null and b.终止时间 is null And b.开始原因 = [2] And a.Id = [1]"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISMsg", lng医嘱ID, int原因)
    If Not rsTmp.EOF Then lng变动id = rsTmp!ID: str病情 = rsTmp!病情 & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人当前病情(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：获取病人当前的病情
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select 病情 From (Select 病情 From 病人变动记录 Where 病人id = [1] And 主页id = [2] Order By 开始时间 Desc) Where Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISMsg", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then Get病人当前病情 = rsTmp!病情 & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



