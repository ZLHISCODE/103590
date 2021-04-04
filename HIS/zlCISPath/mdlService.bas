Attribute VB_Name = "mdlService"
Option Explicit

Public gobjService As Object        'HIS拆分服务对象
Public Enum JSON_TYPE
    Json_Text = 0 '字符
    Json_num = 1 '数值
End Enum

Public Function OpenJson(ByVal strJson As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:设置一个Json串
'入参:strJson-Json串
'出参:
'返回:设置成功,返回true,否则返回False
'编制:余伟节
'日期:2019-08-11 19:36:34
'---------------------------------------------------------------------------------------------------------------------------------------------
    OpenJson = gobjService.SetJsonString(strJson)
End Function

Public Function GetNode(ByVal strKey As String, ByVal varValue As Variant, Optional ByVal blnFirst As Boolean, _
Optional ByVal bytType As Byte) As String
'功能:获取单个JSON元素键值对
'   strKey-   Json key值
'   strValue- Json value值 数字\字符串
'   blnFirst  = T 第一个节点;F-非第一节点
'   bytType   =1 拼接字符"{}";=2 拼接字符"[]"
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
'功能：获取Json中的数组数据或子结点数据到集合中
'参数：
'  strList=Json数组结点或父结点名及路径，如：output，output.pati_list，output.pati_list[0].baby_list
'  strKeys=数组中作为关键字的结点名，可以多个用","号分隔，如"pati_id,pati_pageid"。注意关键字结点的数据不允许存在重复
'  varNullValue=当数组中的结点值为为null时，返回的转换值
    Set GetJsonListValue = gobjService.GetJsonListValue(strListPathNode, strKeyNodes, varNullValue)
End Function

Public Function GetJsonNodeValue(ByVal strPathNode As String, Optional ByVal varNullValue As Variant) As Variant
'功能：获取Json指定结点的值
'参数：
'  strElement=结点及路径，如：output.message，output.pati_list[0].phone_number,output.num_list
'  varNullValue=当结点值为为null时，返回的转换值
    GetJsonNodeValue = gobjService.GetJsonNodeValue(strPathNode, varNullValue)
End Function

Public Function InitSvr() As Boolean
'功能：初始化服务接口部件
    If gobjService Is Nothing Then
        On Error Resume Next
        Set gobjService = CreateObject("zlServiceCall.clsServiceCall")
        If Not gobjService.InitService(gcnOracle, gstrDBUser, glngSys) Then
            Set gobjService = Nothing
        End If
        Err.Clear: On Error GoTo 0
    End If
    If gobjService Is Nothing Then
        MsgBox "zlServiceCall.clsServiceCall创建失败!", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not gobjService Is Nothing Then InitSvr = True
End Function

Public Function CallService(ByVal strServiceName As String, _
    ByVal strJson_In As String, Optional ByRef strJson_out As String, Optional ByVal strTittle As String, _
    Optional lngModule As Long, Optional blnShowErrMsg As Boolean = True, Optional ByVal strAskDate As String, _
    Optional varExpend As String, Optional lngSys As Long, Optional blnReadServiceErr As Boolean) As Boolean
'功能：调用服务
'相关说明见 zlServiceCall.clsServiceCall.CallService 接口
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
'功能:获取病人ID
'参数:
'       _pati_id 病人ID
'       _pati_pageid 主页ID
    Dim strIn As String
    Dim strOut As String
    Dim strTemp As String
   
    
    strIn = GetNode("find_name", strFindName, True)
    strIn = strIn & GetNode("find_text", strFindValue)
    strIn = GetNode("other_cons_find", "{" & strIn & "}", True, 1)
    strIn = GetNode("input", strIn, True, 1)
    If Not CallService("Zl_Patisvr_Getpatiid", strIn, strOut, "获取病人ID", P临床路径应用) Then Exit Function
    Set colPati = GetJsonNodeValue("output.pati_list[0]")
End Function

Public Function ZL_PatiSvr_GetPatiInfo(ByVal lngPatiID As Long, Optional ByVal bytQueryType As Byte, Optional ByVal bytCard As Byte, _
    Optional ByVal bytFamily As Byte, Optional ByVal bytDrug As Byte, Optional ByVal bytImmune As Byte, _
    Optional ByVal strPatiIds As String, Optional ByVal strPatiName As String, Optional dblOutNum As Double, _
    Optional strIdCard As String, Optional strContactId As String, Optional dblCardTypeId As Double, _
    Optional strMedcCardName As String, Optional strCardNO As String, Optional strQRcode As String, _
    Optional strICCardNo As String, Optional strVisitCard As String, Optional strInsuranceNum As String, _
    Optional intStatu As Integer = -1, Optional strPhoneNumber As String, Optional strBed As String) As Boolean
'功能:获取病人信息
'参数:
'     lngPatiID           N 1 病人id  病人ID<>0时，查询列表中的条件无效
'     bytQueryType        N 1 查询类型:如：0-基本;1-基本+联系人;2-所有
'     bytCard           N 1 是否包含卡信息:1-包含医疗卡;0-不包含医疗卡
'     bytFamily         N 1 是否包含家属:1-包含家属信息，0-不包含家属信息
'     bytDrug           N 1 是否包含过敏药物:1-包含，0-不包含
'     bytImmune         N 1 是否包含免疫修:1-包含;0-不包含
'       strPatiIds          C   病人IDs:多个用逗号
'       strPatiName         C   姓名:可以代%分号表表按姓名匹配
'       dblOutNum           N   门诊号
'       strIdCard           C   身份证号
'       strContactId        C   联系人身份证号
'       dblCardTypeId       N   医疗卡类别ID
'       strMedcCardName     C   医疗卡名称
'       strCardNo           C   卡号
'       strQRcode           C   二维码
'       strICCardNo         C   Ic卡号
'       strVisitCard        C   就诊卡号
'       strInsuranceNum     C   医保号
'       intStatu            C   查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
'       strPhoneNumber      C   手机号
'       strBed              C   当前床号

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
    ZL_PatiSvr_GetPatiInfo = CallService("Zl_Patisvr_Getpatiinfo", strIn, strOut, "获取病人基本信息", P临床路径应用)
End Function

Public Function GetMoneyInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef dblRemainMoney As Double, ByRef dblPrePayMoney As Double) As Boolean
'功能：获取指定病人的剩余额
'参数：
'   dblRemainMoney-剩余款
'   dblExpectedMoney-预结费用
    Dim strIn As String
    Dim strOut As String
    Dim strTemp As String
    
    On Error GoTo errH
 
    strIn = GetNode("pati_id", lng病人ID, True)
    strIn = strIn & GetNode("pati_pageid", lng主页ID)
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
              
    If Not CallService("Zl_Exsesvr_Getremainmoney", strIn, strOut, "费用余额", P临床路径应用) Then Exit Function
    
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
'功能:获取库存不足的药品ID
'参数:  strDrugIds 药品IDs 多个药品ID用逗号分隔
'返回:
'  --入参：Json_In:格式
'  --  input
'  --   drug_ids       C   1   药品ID，多个用英文的逗号分隔
'  --   pharmacy_ids   C   0   库房ID，多个用英文的逗号分隔;空字符串,查询所有库房
'  --   return_price   N   0   是否返回售价：1-返回价格信息(售价);0-不返回
'  --   return_dept    N   0   按科室返回库存：1-按科室返回库存;0-按药品返回库存;2-返回科室所有药品的库存
'  --   query_type     N   1   查询类型:如：0-查询库存不等于0,1-查询库存小于等于0
'  --出参: Json_Out,格式如下
'  --  output
'  --    code                 N   1 应答吗：0-失败；1-成功
'  --    message              C   1 应答消息：失败时返回具体的错误信息
'  --    item_list
'  --    drug_id              N   1   药品ID
'  --    pharmacy_id          N   1   库房ID(按科室返回库存才有此项)
'  --    stock                N   1   可用数量
'  --    price                N   1   零售价(返回价格时才有此项)
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
    If Not CallService("zl_DrugSvr_GetStockBatch", strJson, strJsonOut, App.ProductName, p临床路径管理, True, , , , True) Then
        Exit Function
    End If
    Set colList = gobjService.GetJsonListValue("output.item_list", strKeyNodes)
    GetDrugStockBatch = True
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
            strJson = strJson & ":" & IIf(Mid(strValue, 1, 1) = ".", "0", "") & strValue
        End If
    End If
    GetJsonNodeString = strJson
End Function

Public Function GetNodeString(ByVal strNodeName As String) As String
    GetNodeString = Chr(34) & strNodeName & Chr(34)
End Function

Public Function GetStrByRS(ByVal rsTemp As ADODB.Recordset, Optional ByVal strFiledName As String = "ID") As String
'功能:根据传入记录集返回指定字段用逗号分隔的字符串
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
'功能:通集合关键字获取集合的值,基本数据类型,数字或字符
'入参：strType  N/n  表示数字类型，c表示字符串
'      strDef  缺省值，当出错时以这个值为缺省值返回
'出参:lngExist 集合中是否存在这个结点值,0-存在,-1不存在
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
    '集合访问不到不提示继续处理
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
'功能:通集合关键字获取集合中的集合对象
    On Error GoTo errH
    Set GetColObj = colData(strKey)
    Exit Function
errH:
    Err.Clear
    Set GetColObj = New Collection
End Function
