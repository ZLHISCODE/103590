Attribute VB_Name = "mdlPubJson"
Option Explicit
Private mobjServiceCall As Object

'JSON节点类型
Public Enum JSON_TYPE
    Json_Text = 0 '字符
    Json_num = 1 '数值
End Enum


Public Function zlGetNodeValueFromCollect(ByVal cllData As Collection, ByVal strKey As String, ByVal strType As String) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定节点的数据集
    '入参:cllData-当前个集合
    '     strKey-Key
    '     strType-"N"-数字;"C"字符
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-14 16:20:53
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
    '功能:获取指定节点的对象集
    '入参:cllData-当前个集合
    '     strKey-Key
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-14 16:20:53
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
    '功能:获取Json接点串
    '入参:strNodeName-接点名
    '     strValue-值
    '     intType-类型:0-字符;1-数字
    '     blnZeroToEmpty-是否将数值0转换为Null，仅类型为数字时有效
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 18:59:04
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
    '功能：获取Json数组返回的集合数据中指定行或指定元素的值
    '参数：
    '  varRow=行索引或行关键字
    '  strElement=元素名
    '返回：
    '  当未传入strElement参数时，返回指定行的集合对象；当传入strElement参数时，返回指定行指定元素的值
    '  失败时返回Nothing或Empty，但不会报错
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
    '根据关键字判断元素是否存在于集合中
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
    '功能:获取公共服务对象
    '出参:objServiceCall_Out-返回公共服务对象
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String
    If Not mobjServiceCall Is Nothing Then Set objServiceCall_Out = mobjServiceCall: GetServiceCall = True: Exit Function
    
    err = 0: On Error Resume Next
    Set mobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
    If err <> 0 Then
        strErrMsg = "部件【zlServiceCall】丢失，请与系统管理员联系，恢复该部件！"
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

Public Function zlDrugSvr_GetPharmacyWindows(ByVal str药房IDs As String, ByRef rsData As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药房的发药窗口
    '入参:
    '   str药房IDs 药房ID，多个用英文逗号分隔
    '出参:
    '   rsData 字段：药房ID,发药窗口,是否专家
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim cllData As Collection, cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    
    Set rsData = New ADODB.Recordset
    With rsData.Fields
        .Append "药房ID", adBigInt, 18, adFldIsNullable
        .Append "发药窗口", adLongVarChar, 50, adFldIsNullable
        .Append "是否专家", adInteger, 2, adFldIsNullable
    End With
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open
    
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接药品服务失败，无法获取药房的发药窗口！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Drugsvr_Getpharmacywindows
    '  --功能：获取药房所涉及的发药窗口
    '  --入参：Json_In:格式
    '  --  input
    '  --    pharmacy_ids            C   1  药房ID1,药房ID2…
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                    N   1   应答码：0-失败；1-成功
    '  --    message                 C   1   每个药房id对应的发药窗口[数组]
    '  --    window_list[]    更新数据列表[数组]
    '  --        pharmacy_id             N 1 药房ID
    '  --        pharmacy_window         C 1 发药窗口
    '  --        expert_window           N 1 是否专家窗口：1-是，0-不是
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pharmacy_ids", str药房IDs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("zl_DrugSvr_Getpharmacywindows", strJson, , "", glngModul) = False Then Exit Function
    
    Set cllData = objServiceCall.GetJsonListValue("output.window_list")
    If cllData Is Nothing Then Exit Function
    
    For i = 1 To cllData.Count
        Set cllTemp = cllData(i)
        rsData.AddNew
        rsData!药房Id = cllTemp("_pharmacy_id")
        rsData!发药窗口 = cllTemp("_pharmacy_window")
        rsData!是否专家 = cllTemp("_expert_window")
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

Public Function zlExseSvr_UpdRgstArrangeMent(ByVal int操作类型 As Integer, ByVal lng医生ID As Long, _
                Optional ByVal str撤档时间 As String, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整号源、有效的安排、有效的出诊记录中的医生姓名。
    '入参:int操作类型-1-修改姓名,2-停用人员,3-启用人员
    '     str撤档时间-停用和启用时传入，启用时传入原撤档时间
    '出参:strErrMsg_Out
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim intReturn As Integer
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接费用服务失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
'    Zl_ExseSvr_UpdRgstArrangement
'    --功能：调整号源、有效的安排、有效的出诊记录中的医生姓名。
'    --入参
'    --input      调整号源、有效的安排、有效的出诊记录中的医生姓名
'    --  oper_type     N  1  操作方式：1-修改姓名,2-停用人员,3-启用人员
'    --  rgst_dr_id      N  1  病人id
'    --  revoke_time   C         撤档时间
'    --出参
'    --output
'    --  code          C    1  应答码：0-失败；1-成功
'    --  message         C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_type", int操作类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rgst_dr_id", lng医生ID, Json_num)
    If str撤档时间 <> "" Then
        strJson = strJson & "," & GetJsonNodeString("revoke_time", str撤档时间, Json_Text)
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_ExseSvr_UpdRgstArrangement", strJson, , "", glngModul, False) = False Then Exit Function
    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    If intReturn <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out <> "" Then strErrMsg_Out = "更新挂号安排失败！"
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
