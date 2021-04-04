Attribute VB_Name = "mdlEInvoice_BS"
Option Explicit
'*********************************************************************************************************************************************
'博思电子票据相关处理
'一、电子票据公共接口处理:
'   1.zlInitIFacePara :初始化博思接口配置
'   2.zlGet业务标识:对标His业务场合与博思业务标识
'二、电子票据接口调用相关数据处理:
'   1.zlGetJson_CreateEInvoice:博思开具电子票据需要的Json数据
'     1.1:zlGetJson_CreateEInvoiceByCharge:收费电子票据
'     1.2:zlGetJson_CreateEInvoiceByDeposit:预交电子票据
'     1.3:zlGetJson_CreateEInvoiceByMzBalance:门诊结帐电子票据
'     1.4:zlGetJson_CreateEInvoiceByZyBalance:住院结帐电子票据
'     1.5:zlGetJson_CreateEInvoiceByRegsit:挂号电子票据
'     1.6:zlGetJson_CreateEInvoiceBySendCard:发卡电子票据
'   3.zlGetJson_PrintEInvoice:获取打印电子票据Json格式数据
'   4.zlGetJson_SendNotice:获取发生告知单Json格式数据
'   5.zlGetJson_CheckCancelEInvoice:票据作废检查的Json
'   6.zlGetJson_CancelEInvoice:票据作废Json
'三、纸质票据相关接口
'   1.zlGetJson_GetNextInvoiceNo:获取纸质票号Json格式数据
'   2.zlGetJson_TurnPaper:获取换开纸质票据Json格式数据
'   3.zlGetJson_TurnPaperPrint:获取换开纸质票据打印Json格式数据
'   4.zlGetJson_CancelPaper:获取纸质票据作废Json格式数据

'目前涉及博思的接口
'1.invoiceEBillOutpatient:门诊收费电子票据
'2.invEBillHospitalized:住院电子票据
'1.getEBillAccountStatus
'2.invoicePayMentVoucher
'编制:李南春
'日期:2020-03-03 14:11:42
'*********************************************************************************************************************************************
Public Enum BS_Version
    V2_0_3 = 0
    V3_1_0
    V3_2_0
End Enum

Private Type IFaceBs
    URL_Type                As String
    URL_Address             As String
    应用帐号                As String
    签名私钥                As String
    支持版本                As BS_Version
    数据传输方式            As String
    字符编码                As String
    缺省卡类别ID            As Long
    医疗卡类型编号          As String
    身份证作卡类型编号      As String
    病人无卡的卡类别编号    As String
    病人无卡的卡号          As String
    录入冲红原因            As Boolean
    误差费对照编码          As String
    误差费对照名称          As String
    零费用开票              As Boolean
    收费纸质票据代码        As String
    挂号纸质票据代码        As String
    结账纸质票据代码        As String
    预交纸质票据代码        As String
End Type
Public gBs_Type As IFaceBs

Private mlngSys As Long
Private mstrOperatorCode As String
Private mstrOperatorName As String
Private mcllJsonKey As Collection
Private mcllJsonFormat As Collection


Public Function zlInitIFacePara(ByVal lngSys As Long, ByVal strOperatorCode As String, ByVal strOperatorName As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取博思接口配置
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/21 15:35
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    mlngSys = lngSys
    mstrOperatorCode = strOperatorCode: mstrOperatorName = strOperatorName
    
    gBs_Type.URL_Type = ""
    gBs_Type.URL_Address = ""
    gBs_Type.应用帐号 = ""
    gBs_Type.签名私钥 = ""
    gBs_Type.支持版本 = 0
    gBs_Type.数据传输方式 = ""
    gBs_Type.字符编码 = ""
    gBs_Type.缺省卡类别ID = 0
    gBs_Type.身份证作卡类型编号 = ""
    gBs_Type.病人无卡的卡类别编号 = ""
    gBs_Type.病人无卡的卡号 = ""
    gBs_Type.医疗卡类型编号 = "" '博思证件类型对照
    gBs_Type.误差费对照编码 = ""
    gBs_Type.误差费对照名称 = ""
    gBs_Type.零费用开票 = False
    gBs_Type.收费纸质票据代码 = ""
    gBs_Type.挂号纸质票据代码 = ""
    gBs_Type.结账纸质票据代码 = ""
    gBs_Type.预交纸质票据代码 = ""
    
    strSQL = "Select 参数号, 参数名, 参数值 From 三方接口配置 Where 接口名 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlInitIFacePara", gobjEinvProvider.提供者)
    With rsTmp
        Do While Not .EOF
            Select Case UCase(Nvl(!参数名))
                Case "URL_TYPE"
                    gBs_Type.URL_Type = Nvl(!参数值)
                Case "URL_ADDRESS"
                    gBs_Type.URL_Address = Nvl(!参数值)
                Case "应用帐号"
                    gBs_Type.应用帐号 = Nvl(!参数值)
                Case "签名私钥"
                    gBs_Type.签名私钥 = Nvl(!参数值)
                Case "支持版本"
                    If Nvl(!参数值) = "V3.2.0" Then
                        gBs_Type.支持版本 = BS_Version.V3_2_0
                    ElseIf Nvl(!参数值) = "V3.1.0" Then
                        gBs_Type.支持版本 = BS_Version.V3_1_0
                    Else
                        gBs_Type.支持版本 = BS_Version.V2_0_3
                    End If
                    
                Case "数据传输方式"
                    gBs_Type.数据传输方式 = Nvl(!参数值)
                Case "字符编码"
                    gBs_Type.字符编码 = Nvl(!参数值)
                Case "缺省卡类别ID"
                    gBs_Type.缺省卡类别ID = Val(Nvl(!参数值))
                Case "身份证作卡类型编号"
                    gBs_Type.身份证作卡类型编号 = Nvl(!参数值)
                Case "病人无卡的卡类别编号"
                    gBs_Type.病人无卡的卡类别编号 = Nvl(!参数值)
                Case "病人无卡的卡号"
                    gBs_Type.病人无卡的卡号 = Nvl(!参数值)
                Case "录入冲红原因"
                    gBs_Type.录入冲红原因 = Val(Nvl(!参数值)) = 1
                Case "医疗卡类型编号"
                    gBs_Type.医疗卡类型编号 = Nvl(!参数值)
                Case "误差费对照编码"
                    gBs_Type.误差费对照编码 = Nvl(!参数值)
                Case "误差费对照名称"
                    gBs_Type.误差费对照名称 = Nvl(!参数值)
                Case "零费用开具电子票据"
                    gBs_Type.零费用开票 = Val(Nvl(!参数值)) = 1
                Case "收费纸质票据代码"
                    gBs_Type.收费纸质票据代码 = Nvl(!参数值)
                Case "挂号纸质票据代码"
                    gBs_Type.挂号纸质票据代码 = Nvl(!参数值)
                Case "结账纸质票据代码"
                    gBs_Type.结账纸质票据代码 = Nvl(!参数值)
                Case "预交纸质票据代码"
                    gBs_Type.结账纸质票据代码 = Nvl(!参数值)
            End Select
            .MoveNext
        Loop
    End With
    
    Call InitVersionDiff
    zlInitIFacePara = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitVersionDiff() As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 设置版本差异
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/6/2 13:51
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    '1.节点差异
    Set mcllJsonKey = New Collection
    If gBs_Type.支持版本 > V3_1_0 Then
        Call mcllJsonKey.Add("patientCategory", "_就诊科室") '收费、挂号
        Call mcllJsonKey.Add("patientCategoryCode", "_就诊科室编码") '挂号
    Else
        Call mcllJsonKey.Add("category", "_就诊科室")
        Call mcllJsonKey.Add("patientCategory", "_就诊科室编码")
    End If
    
    '2.数据格式差异
    Set mcllJsonFormat = New Collection
    If gBs_Type.支持版本 > V3_1_0 Then
        Call mcllJsonFormat.Add(4, "_费用小数") '收费、挂号、结帐
        Call mcllJsonFormat.Add(2, "_数量小数") '收费、挂号、结帐
        Call mcllJsonFormat.Add("yyyyMMdd", "_就诊日期") '包括出入院日期， 收费、挂号、结帐
    Else
        Call mcllJsonFormat.Add(6, "_费用小数")
        Call mcllJsonFormat.Add(6, "_数量小数")
        Call mcllJsonFormat.Add("yyyy-MM-dd", "_就诊日期")
    End If
    
    '3.数据差异
    '此处无法处理
    
    InitVersionDiff = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetVersionDiff(ByVal bytType As Byte, ByVal strKey As String) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取版本差异
    ' 入参 : bytType:1-节点差异；2-节点格式差异
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/6/2 13:51
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    
    If bytType = 1 Then
        GetVersionDiff = mcllJsonKey("_" & strKey)
    Else
        GetVersionDiff = mcllJsonFormat("_" & strKey)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlGet业务标识(ByVal byt场合 As Byte) As String
    '业务标识:01  住院,02  门诊, 03  急诊, 04  门特, 05  体检中心, 06  挂号, 07  住院预交金, 08  体检预交金
    zlGet业务标识 = Decode(byt场合, 1, "02", 2, "07", 3, "01", 4, "06", 5, "02", "02")
End Function

Public Function zlGetJson_CreateEInvoice(ByVal bytInvoiceType As Byte, ByVal lngEInvoiceID As Long, ByVal lng结帐ID As Long, ByVal lng销账ID As Long, _
                ByVal strEInvoiceClientCode As String, strServiceCode As String, dbl票据金额_Out As Double, strJson_Out As String, _
                Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取开具发票Json格式数据
    ' 入参 : bytInvoiceType-调用场合
    '        lngEInvoiceID -电子票据使用记录.ID
    '        strEInvoiceClientCode-开票点编号
    ' 出参 : strJson_Out-开票信息
    '        strServiceCode-服务标识
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    Select Case bytInvoiceType
        Case 1
            If Not zlGetJson_CreateEInvoiceByCharge(lngEInvoiceID, lng结帐ID, lng销账ID, strEInvoiceClientCode, dbl票据金额_Out, strJson_Out, strErrMsg_Out) Then Exit Function
            strServiceCode = "invoiceEBillOutpatient"
        Case 2
            If Not zlGetJson_CreateEInvoiceByDeposit(lngEInvoiceID, lng结帐ID, lng销账ID, strEInvoiceClientCode, dbl票据金额_Out, strJson_Out, strErrMsg_Out) Then Exit Function
            If lng销账ID <> 0 Then
                strServiceCode = "writeOffPayMentVoucher"
            Else
                strServiceCode = "invoicePayMentVoucher"
            End If
        Case 3
            strSQL = "Select Max(结帐类型) As 结帐类型 From 病人结帐记录 Where ID = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "判断结帐类型", lng结帐ID)
            If rsTmp.RecordCount = 0 Then
                strErrMsg_Out = "未找到结算数据，不能打印电子票据": Exit Function
            End If
            If Val(Nvl(rsTmp!结帐类型)) = 1 Then
                If Not zlGetJson_CreateEInvoiceByMzBalance(lngEInvoiceID, lng结帐ID, lng销账ID, strEInvoiceClientCode, dbl票据金额_Out, strJson_Out, strErrMsg_Out) Then Exit Function
                strServiceCode = "invoiceEBillOutpatient"
            Else
                If Not zlGetJson_CreateEInvoiceByZyBalance(lngEInvoiceID, lng结帐ID, lng销账ID, strEInvoiceClientCode, dbl票据金额_Out, strJson_Out, strErrMsg_Out) Then Exit Function
                strServiceCode = "invEBillHospitalized"
            End If
        Case 4
            If Not zlGetJson_CreateEInvoiceByRegsit(lngEInvoiceID, lng结帐ID, lng销账ID, strEInvoiceClientCode, dbl票据金额_Out, strJson_Out, strErrMsg_Out) Then Exit Function
            strServiceCode = "invEBillRegistration"
        Case 5
            If Not zlGetJson_CreateEInvoiceBySendCard(lngEInvoiceID, lng结帐ID, lng销账ID, strEInvoiceClientCode, dbl票据金额_Out, strJson_Out, strErrMsg_Out) Then Exit Function
            strServiceCode = "invoiceEBillOutpatient"
        Case Else
            strErrMsg_Out = "无效的应用场合": Exit Function
    End Select
    zlGetJson_CreateEInvoice = True
End Function

Public Function zlGetJson_PrintEInvoice(ByVal lngEInvoiceID As Long, strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取打印电子票据Json格式数据
    ' 入参 : lngEInvoiceID -电子票据使用记录.ID
    ' 出参 : strJson_Out-开票信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    strJson_Out = ""
    
    Set rsTmp = GetEInvoiceInfo(lngEInvoiceID, strErrMsg_Out)
    If rsTmp Is Nothing Then Exit Function
    
    With rsTmp
        strJson_Out = GetJsonNodeString("billBatchCode", Nvl(!票据代码), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("billNo", Nvl(!票据号码), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("random", Nvl(!票据校验码), Json_Text)
        strJson_Out = "{" & strJson_Out & "}"
    End With
    zlGetJson_PrintEInvoice = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_SendNotice(ByVal lngEInvoiceID As Long, strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取发生告知单Json格式数据
    ' 入参 : lngEInvoiceID -电子票据使用记录.ID
    ' 出参 : strJson_Out-开票信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strJsonList As String, strJson As String
    On Error GoTo ErrHand
    strJson_Out = ""
    
    Set rsTmp = GetEInvoiceWithPatiInfo(lngEInvoiceID, strErrMsg_Out)
    If rsTmp Is Nothing Then Exit Function
    
    
    With rsTmp
        strJson_Out = GetJsonNodeString("billBatchCode", Nvl(!票据代码), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("billNo", Nvl(!票据号码), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("random", Nvl(!票据校验码), Json_Text)
        
        If Nvl(!手机号) <> "" Then
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("noticeType", 1201, Json_Text)
            strJson = strJson & "," & GetJsonNodeString("noticeValue", Nvl(!手机号), Json_Text)
            strJsonList = ",{" & strJson & "}"
        End If
        
        If Nvl(!email) <> "" Then
            strJson = ""
            strJson = strJson & "" & GetJsonNodeString("noticeType", 1202, Json_Text)
            strJson = strJson & "," & GetJsonNodeString("noticeValue", Nvl(!email), Json_Text)
            strJsonList = strJsonList & ",{" & strJson & "}"
        End If
        If strJsonList = "" Then Exit Function '没有消息推送途径，直接退出
        
        strJson_Out = strJson_Out & "," & GetNodeString("noticeList") & ":[" & Mid(strJsonList, 2) & "]"
        strJson_Out = "{" & strJson_Out & "}"
    End With
    zlGetJson_SendNotice = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_GetNextInvoiceNo(ByVal bytInvoiceType As Byte, ByVal strEInvoiceNodeCode As String, _
                strPaperCode_Out As String, strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取纸质票号Json格式数据
    ' 入参 : bytInvoiceType-票种
    '        strEInvoiceNodeCode-开票点编号
    ' 出参 : strJson_Out-开票信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strJsonList As String, strJson As String
    On Error GoTo ErrHand
    strPaperCode_Out = GetPaperCode(bytInvoiceType)
    
    strJson_Out = ""
    strJson_Out = GetJsonNodeString("placeCode", strEInvoiceNodeCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillBatchCode", strPaperCode_Out, Json_Text)
    '返回完整的Json
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_GetNextInvoiceNo = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_TurnPaper(ByVal bytInvoiceType As Byte, ByVal strEInvoiceNodeCode As String, _
    ByVal strInvoiceNO As String, ByVal lngEInvoiceID As Long, ByVal strEInvoiceCode As String, ByVal strEInvoiceNO As String, _
    ByVal strCreateTime As String, ByVal strOperatorCode As String, ByVal strOperatorName As String, _
    strServiceCode As String, strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取换开纸质票据Json格式数据
    ' 入参 : bytInvoiceType-票种
    '        strEInvoiceNodeCode-开票点编号
    '        bytInvoiceType-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '        strInvoiceNO-发票号
    '        lngEInvoiceID-电子票据使用记录ID
    '        strEInvoiceCode-电子票据代码
    '        strEInvoiceNO-电子票据号码
    '        strCreateTime-电子票据生成时间,格式:YYYYMMDDhhmmssSSS
    '        strOperatorCode-操作员编号
    '        strOperatorName-操作员姓名
    ' 出参 : strJson_Out-开票信息
    '        strServiceCode-业务标识
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strJsonList As String, strJson As String, strPaperCode As String
    On Error GoTo ErrHand
    Set rsTmp = GetEInvoiceWithPatiInfo(lngEInvoiceID, strErrMsg_Out)
    If rsTmp Is Nothing Then Exit Function
    If Val(Nvl(rsTmp!是否换开)) = 1 Then
        strServiceCode = "reTurnPaper"
    Else
        strServiceCode = "turnPaper"
    End If
    strPaperCode = GetPaperCode(bytInvoiceType)
    strInvoiceNO = Mid(strInvoiceNO, Len(strPaperCode) + 1)
    
    strJson_Out = ""
    strJson_Out = GetJsonNodeString("billBatchCode", strEInvoiceCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("billNo", strEInvoiceNO, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillBatchCode", strPaperCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillNo", strInvoiceNO, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("busDateTime", strCreateTime, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("placeCode", strEInvoiceNodeCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("operator", strOperatorName, Json_Text)
    '返回完整的Json
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_TurnPaper = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_TurnPaperPrint(ByVal bytInvoiceType As Byte, ByVal strInvoiceNO As String, _
                strServiceCode As String, strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取换开纸质票据打印Json格式数据
    ' 入参 : bytInvoiceType-票种
    '        strEInvoiceNodeCode-开票点编号
    '        bytInvoiceType-1-收费,2-预交,3-结帐,4-挂号;5-就诊卡
    '        strInvoiceNO-发票号
    '        lngEInvoiceID-电子票据使用记录ID
    '        strEInvoiceCode-电子票据代码
    '        strEInvoiceNO-电子票据号码
    '        strCreateTime-电子票据生成时间,格式:YYYYMMDDhhmmssSSS
    '        strOperatorCode-操作员编号
    '        strOperatorName-操作员姓名
    ' 出参 : strJson_Out-开票信息
    '        strServiceCode-业务标识
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strPaperCode As String
    On Error GoTo ErrHand
    strServiceCode = IIf(bytInvoiceType = 2, "", "printPaperBill")
    
    strPaperCode = GetPaperCode(bytInvoiceType)
    strInvoiceNO = Mid(strInvoiceNO, Len(strPaperCode) + 1)
    
    strJson_Out = ""
    strJson_Out = GetJsonNodeString("pBillBatchCode", strPaperCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillNo", strInvoiceNO, Json_Text)
    '返回完整的Json
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_TurnPaperPrint = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_CheckCancelEInvoice(ByVal lngEInvoiceID As Long, strEInvoiceNo_Out As String, strJson_Out As String, Optional strErrMsg_Out As String, _
                Optional ByVal blnCheckAcc As Boolean, Optional strJsonAcc_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取冲红发票检查Json格式数据
    ' 入参 : lngEInvoiceID -电子票据使用记录.ID
    ' 出参 : strEInvoiceNo_Out-电子票据号
    '        strJson_Out-票据信息
    '        strJsonAcc_Out-票据入账信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    strJson_Out = ""
    
    Set rsTmp = GetEInvoiceInfo(lngEInvoiceID, strErrMsg_Out)
    If rsTmp Is Nothing Then Exit Function
    With rsTmp
        strEInvoiceNo_Out = Nvl(!票据号码)
        strJson_Out = GetJsonNodeString("billBatchCode", Nvl(!票据代码), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("billNo", strEInvoiceNo_Out, Json_Text)
        
        If blnCheckAcc Then
            strJsonAcc_Out = strJson_Out
            strJsonAcc_Out = strJsonAcc_Out & "," & GetJsonNodeString("random", Nvl(!票据校验码), Json_Text)
            strJsonAcc_Out = strJsonAcc_Out & "," & GetJsonNodeString("createTime", Nvl(!生成时间), Json_Text)
        End If
    End With
    '返回完整的Json
    strJson_Out = "{" & strJson_Out & "}"
    If blnCheckAcc Then strJsonAcc_Out = "{" & strJsonAcc_Out & "}"
    zlGetJson_CheckCancelEInvoice = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_CancelEInvoice(ByVal frmMain As Object, ByVal lngEInvoiceID As Long, ByVal strEInvoiceClientCode As String, ByVal blnNoInputReason As Boolean, _
                strServiceCode As String, strJson_Out As String, strReason_Out As String, Optional strErrMsg_Out As String, _
                Optional ByVal strEInvoiceType As String, Optional ByVal strEInvoiceCode As String, Optional ByVal strEInvoiceNO As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取冲红发票Json格式数据
    ' 入参 : lngEInvoiceID -电子票据使用记录.ID
    '        strEInvoiceClientCode-开票点编号
    '        blnNoInputReason-是否输入作废原因，可能作废纸质票据时已录入
    '        strEInvoiceType-业务场合
    '        strEInvoiceCode-电子票据代码
    '        strEInvoiceNo-电子票据号码
    ' 出参 : strJson_Out-开票信息
    '        str业务标识-业务标识
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim byt票种 As Byte
    On Error GoTo ErrHand
    strJson_Out = ""
    
    If strEInvoiceNO = "" Then
        Set rsTmp = GetEInvoiceInfo(lngEInvoiceID, strErrMsg_Out)
        If rsTmp Is Nothing Then Exit Function
        With rsTmp
            byt票种 = Val(Nvl(!票种))
            strEInvoiceNO = Nvl(!票据号码)
            strEInvoiceCode = Nvl(!票据代码)
        End With
    Else
        byt票种 = Decode(strEInvoiceType, "02", 1, "07", 2, "01", 3, "06", 4, "02")
    End If
    
    With rsTmp
        If Not blnNoInputReason Then
            If gBs_Type.录入冲红原因 And mlngSys <> 2600 Then
                If frmInputBox.InputBox(frmMain, "票据冲红", "请录入票据冲红的原因：", 30, 1, False, False, strReason_Out) = False Then Exit Function
            Else
                strReason_Out = Decode(byt票种, 2, "退预交", 3, "结帐作废", 4, "退号", 5, "退卡", "退费")
            End If
        End If
        strServiceCode = IIf(byt票种 = 2, "cancelPayMentVoucherBalance", "writeOffEBill")
        
        strJson_Out = GetJsonNodeString("billBatchCode", strEInvoiceCode, Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("billNo", strEInvoiceNO, Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("reason", strReason_Out, Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("operator", mstrOperatorName, Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("busDateTime", Format(zlDatabase.Currentdate, "YYYYMMDDhhmmss000"), Json_Text)
        strJson_Out = strJson_Out & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
        strJson_Out = "{" & strJson_Out & "}"
    End With
    zlGetJson_CancelEInvoice = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_CancelPaper(ByVal frmMain As Object, ByVal bytInvoiceType As String, ByVal strInvoiceNO As String, ByVal strEInvoiceClientCode As String, ByVal strOperatorName As String, _
                strServiceCode As String, strBusDateTime As String, strJson_Out As String, strReason_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取票据作废Json格式数据
    ' 入参 : bytInvoiceType -票种，暂代纸质票据代码
    '        strInvoiceNO-纸质票据号
    '        lngEInvoiceID - 电子票据ID
    '        strEInvoiceClientCode-开票点编号
    ' 出参 : strJson_Out-开票信息
    '        strServiceCode-业务标识
    '        strBusDateTime-作废时间
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strCurCode As String, strPaperCode As String, strPaperNo As String
    
    On Error GoTo ErrHand
    strJson_Out = ""
    strCurCode = GetPaperCode(bytInvoiceType)
    strPaperCode = Left(strInvoiceNO, Len(strCurCode))
    strPaperNo = Mid(strInvoiceNO, Len(strCurCode) + 1)
    
    If gBs_Type.录入冲红原因 And mlngSys <> 2600 Then
        If frmInputBox.InputBox(frmMain, "票据冲红", "请录入票据冲红的原因：", 30, 1, False, False, strReason_Out) = False Then Exit Function
    Else
        strReason_Out = Decode(bytInvoiceType, 2, "退预交", 3, "结帐作废", 4, "退号", 5, "退卡", "退费")
    End If
    strServiceCode = IIf(bytInvoiceType = 2, "invalidPayMentVoucherPaper", "invalidPaper")
    
    strJson_Out = GetJsonNodeString("pBillBatchCode", strPaperCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillNo", strPaperNo, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("author", strOperatorName, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("reason", strReason_Out, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("busDateTime", Format(zlDatabase.Currentdate, "YYYYMMDDhhmmss000"), Json_Text)
    '返回完整的Json
    strJson_Out = "{" & strJson_Out & "}"
    
    zlGetJson_CancelPaper = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function zlGetJson_CancelBlankInvoice(ByVal strBatchNo As String, ByVal strStartInvoice As String, _
                    ByVal strEndInvoice As String, ByVal strEInvoiceClientCode As String, _
                    ByVal strAuthorName As String, ByVal strReason As String, ByVal strHappenTime As String, _
                    strJson_Out As String, strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取票据报损Json格式数据
    ' 入参 : strBatchNo-批次，暂代纸质票据代码
    '        strStartInvoice-起始纸质票据号
    '        strEndInvoice-终止纸质票据号
    '        strAuthorName-作废人
    '        strReason-作废原因
    '        strEInvoiceClientCode-开票点编号
    '        strHappenTime-作废时间
    ' 出参 : strJson_Out-票据报损信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    strJson_Out = ""
    strJson_Out = GetJsonNodeString("pBillBatchCode", strBatchNo, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillNoStart", strStartInvoice, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("pBillNoEnd", strEndInvoice, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("author", strAuthorName, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("reason", strReason, Json_Text)
    strJson_Out = strJson_Out & "," & GetJsonNodeString("busDateTime", Format(strHappenTime, "YYYYMMDDhhmmss000"), Json_Text)
    '返回完整的Json
    strJson_Out = "{" & strJson_Out & "}"
    
    zlGetJson_CancelBlankInvoice = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceByRegsit(ByVal lngEInvoiceID As Long, ByVal lng结帐ID As Long, ByVal lng销账ID As Long, _
                ByVal strEInvoiceClientCode As String, dbl票据总金额 As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取挂号发票Json格式数据
    ' 入参 : lngEInvoiceID -电子票据使用记录.ID
    '        strEInvoiceClientCode-开票点编号
    ' 出参 : strJson-挂号结算信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim cllInsureInfo As Collection
    Dim bytInvoiceType As Byte
    Dim dbl误差费 As Double
    Dim lng病人ID As Long, lng挂号ID As Long
    Dim str业务标识 As String, strChargeDetail As String, strListDetail As String
    Dim str结帐IDs As String, str登记时间 As String, str业务操作员 As String
    Dim str患者姓名 As String, str患者性别 As String, str患者年龄 As String, str医疗付款方式编码 As String
    Dim strJsonList As String, strData As String, strValue As String
    Dim strJsonKey_就诊科室 As String, strJsonKey_就诊科室编码
    Dim strJsonFormat_就诊日期 As String
    Dim intJsonFormat_费用小数 As Integer, intJsonFormat_数量小数 As Integer
    On Error GoTo ErrHand
    bytInvoiceType = 4
    str业务标识 = zlGet业务标识(bytInvoiceType)
    dbl票据总金额 = 0
    
    '版本差异
    strJsonKey_就诊科室 = GetVersionDiff(1, "就诊科室")
    strJsonKey_就诊科室编码 = GetVersionDiff(1, "就诊科室编码")
    strJsonFormat_就诊日期 = GetVersionDiff(2, "就诊日期")
    intJsonFormat_费用小数 = Val(GetVersionDiff(2, "费用小数"))
    intJsonFormat_数量小数 = Val(GetVersionDiff(2, "数量小数"))
    
    strSQL = "Select Min(a.Id) As 费用id, a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, Max(a.计算单位) As 计算单位," & vbNewLine & _
            "          Sum(a.标准单价) As 价格, Avg(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 数量, Sum(a.应收金额) As 应收金额," & vbNewLine & _
            "          Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.实收金额) - Sum(a.统筹金额) As 自费金额," & vbNewLine & _
            "          Max(s.大类编码) As 医保项目编码, Max(s.大类名称) As 医保项目名称, Max(t.统筹比额) As 医保报销比例, Max(a.摘要) As 备注," & vbNewLine & _
            "          Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.姓名) As 姓名," & vbNewLine & _
            "          Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间," & vbNewLine & _
            "          Max(a.付款方式) As 付款方式编码, Max(Nvl(c.名称, c1.名称)) As 收据费目, Max(Nvl(c.编码, c1.编码)) As 收据费目编码, Max(a.医嘱序号) As 医嘱序号," & vbNewLine & _
            "          Max(B1.Id) As 挂号id, Max(d.编码) As 类别编码, Max(d.类别) As 类别名称, Max(b.编码) As 项目编码, Max(b.名称) As 项目名称," & vbNewLine & _
            "          Max(b.规格) As 规格, Max(q.药品剂型) As 药品剂型" & vbNewLine & _
            "   From 门诊费用记录 A, 病人挂号记录 B1, 收费项目目录 B, 收据费目对照 C, 收据费目 C1, 收费类别 D, 药品规格 M, 药品特性 Q, 诊疗项目目录 J, 保险支付大类 T, 支付类别对照 S" & vbNewLine & _
            "   Where a.No = B1.No And a.No In (Select Distinct NO From 门诊费用记录 Where 结帐id = [1]) And a.记录性质 = 4 And a.记录状态 = 1 And" & vbNewLine & _
            "         a.收费类别 = d.编码(+) And a.收费细目id = b.Id And a.收据费目 = c1.名称(+) And a.收据费目 = c.收据费目(+) and Decode(c.费用场合(+), 0, 1, c.费用场合(+)) = 1 And a.收费细目id = m.药品id(+) And" & vbNewLine & _
            "         m.药名id = q.药名id(+) And q.药名id = j.Id(+) And a.保险大类id = t.Id(+) And t.性质(+) = 1 And" & vbNewLine & _
            "         a.保险大类id = s.保险大类id(+)" & vbNewLine & _
            "   Group By a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号), a.收费细目id, c.编码, c.名称, j.编码, j.名称" & vbNewLine & _
            IIf(gBs_Type.零费用开票, "", " Having Sum(a.结帐金额) <> 0") & vbNewLine & _
            "   Order By NO, 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByRegsit", lng结帐ID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "未找到明细数据，不能打印电子票据"
        Exit Function
    End If
    
    strJsonList = ""
    With rsTmp
        str患者姓名 = Nvl(!姓名)
        str患者性别 = Nvl(!性别)
        str患者年龄 = Nvl(!年龄)
        lng病人ID = Val(Nvl(!病人ID))
        str医疗付款方式编码 = Nvl(!付款方式编码)
        lng挂号ID = Val(Nvl(!挂号id))
        str登记时间 = Format(Nvl(!登记时间), "YYYYMMDDhhmmss000")
        str业务操作员 = Nvl(!操作员姓名)
        
        Do While Not .EOF
            strData = ""
            strData = strData & "" & GetJsonNodeString("listDetailNo", zlStr.LPAD(Nvl(!费用ID), 20, "0"), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!收据费目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!收据费目), Json_Text)
            strData = strData & "," & GetJsonNodeString("prescribeCode", Nvl(!No), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeCode", Nvl(!类别编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeName", Nvl(!类别名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("code", Nvl(!项目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("name", Nvl(!项目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("form", Nvl(!药品剂型), Json_Text)
            strData = strData & "," & GetJsonNodeString("specification", Nvl(!规格), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!计算单位), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!价格)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!数量)), intJsonFormat_数量小数), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!实收金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!自费金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("receivableAmt", FormatEx(Val(Nvl(!应收金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!医保项目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("medCareItemType", Nvl(!医保项目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("medReimburseRate", FormatEx(Val(Nvl(!医保报销比例)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!备注), Json_Text)
            strData = strData & "," & GetJsonNodeString("sortNo", Nvl(!序号), Json_num)
            strData = strData & "," & GetJsonNodeString("chrgtype", Nvl(!费用类型), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dbl票据总金额 = dbl票据总金额 + RoundEx(Nvl(!实收金额), 6)
            .MoveNext
        Loop
        
        str结帐IDs = GetBalanceIDs(lng结帐ID, bytInvoiceType)
        dbl误差费 = GetBalanceErrorFee(str结帐IDs)
        strListDetail = GetNodeString("listDetail") & ":[" & Mid(strJsonList, 2) & "]"
    End With
    
    '分类明细
    If gBs_Type.误差费对照编码 <> "" Then
        dbl票据总金额 = dbl票据总金额 - dbl误差费
    End If
    dbl票据总金额 = RoundEx(dbl票据总金额, 2)
    If Not Get分类明细(str结帐IDs, strData, dbl票据总金额, bytInvoiceType, strErrMsg_Out) Then Exit Function
    strChargeDetail = GetNodeString("chargeDetail") & ":[" & strData & "]"
    
    '票据信息
    '业务流水号:lngEInvoiceID_lng结帐ID
    strData = ""
    strData = strData & "" & GetJsonNodeString("busNo", lng结帐ID & "_" & lngEInvoiceID, Json_Text)
    strData = strData & "," & GetJsonNodeString("busType", str业务标识, Json_Text)
    strData = strData & "," & GetJsonNodeString("payer", str患者姓名, Json_Text)
    strData = strData & "," & GetJsonNodeString("busDateTime", str登记时间, Json_Text)
    strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strData = strData & "," & GetJsonNodeString("payee", str业务操作员, Json_Text)
    strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("checker", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("totalAmt", dbl票据总金额, Json_num)
    strData = strData & "," & GetJsonNodeString("remark", IIf(RoundEx(dbl误差费, 6) <> 0 And gBs_Type.误差费对照编码 = "", "存在" & FormatEx(dbl误差费, 6) & "误差金额不参与结算", ""), Json_Text)
    strJson_Out = strData
    
    
    '移动支付
    If Not Get移动支付信息(lng病人ID, lng结帐ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '通知消息
    If Not Get通知消息(lng病人ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '就诊信息
    Call Get医保信息(bytInvoiceType, lng结帐ID, lng病人ID, cllInsureInfo)
    strSQL = "Select To_Char(a.发生时间, 'yyyy-mm-dd') As 就诊日期, b.编码 As 就诊科室编码," & vbNewLine & _
            "       b.名称 As 就诊科室名称, a.No As 就诊编号" & vbNewLine & _
            "  From 病人挂号记录 A, 部门表 B" & vbNewLine & _
            "  Where a.执行部门id = b.Id And a.Id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByRegsit", lng挂号ID)
    
    strData = ""
    strData = strData & "" & GetJsonNodeString("medicalInstitution", GetUnitInfo("医疗机构类型"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareInstitution", zlGetNodeValueFromCollect(cllInsureInfo, "_保险机构编码", "C"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareTypeCode", str医疗付款方式编码, Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalCareType", Get医疗付款方式名称(str医疗付款方式编码), Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalInsuranceID", zlGetNodeValueFromCollect(cllInsureInfo, "_医保号", "C"), Json_Text)
    With rsTmp
        If .RecordCount > 0 Then
            strData = strData & "," & GetJsonNodeString("consultationDate", Format(Nvl(!就诊日期), strJsonFormat_就诊日期), Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_就诊科室, Nvl(!就诊科室名称), Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_就诊科室编码, Nvl(!就诊科室编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!就诊编号), Json_Text)
        Else
            strData = strData & "," & GetJsonNodeString("consultationDate", "", Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_就诊科室, "", Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_就诊科室编码, "", Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", lng结帐ID, Json_Text)
        End If
    End With
    strData = strData & "," & GetJsonNodeString("patientId", lng病人ID, Json_Text)
    strData = strData & "," & GetJsonNodeString("sex", str患者性别, Json_Text)
    strData = strData & "," & GetJsonNodeString("age", str患者年龄, Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '支付信息
    If Not Get结算信息(str结帐IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '缴费渠道
    If Not Get缴费渠道(str结帐IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & GetNodeString("payChannelDetail") & ":[" & strData & "]"
    
    '其它医保信息-暂无
    '其它扩展信息-暂无
    'eBillRelateNo  业务票据关联号  String  32  否  如一笔业务数据需要开具N张电子票据，则N张电子票对应该值保持一致，用于后期关联查询
    'isArrears  是否可流通  String  1  是  0-否、1-是（如欠费情况根据医院业务要求该票据是否可流通）
    'arrearsReason  不可流通原因  String  200  否  isArrears=0，填写不可流通的原因
    strData = ""
    strData = strData & "" & GetJsonNodeString("eBillRelateNo", "", Json_Text)
    strData = strData & "," & GetJsonNodeString("isArrears", "1", Json_Text)
    strData = strData & "," & GetJsonNodeString("arrearsReason", "", Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '收费项目明细
    strJson_Out = strJson_Out & "," & strChargeDetail
    '清单项目明细
    strJson_Out = strJson_Out & "," & strListDetail
    
    '返回完整的Json串
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceByRegsit = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceByCharge(ByVal lngEInvoiceID As Long, ByVal lng结帐ID As Long, ByVal lng销账ID As Long, _
                ByVal strEInvoiceClientCode As String, dbl票据总金额 As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取挂号发票Json格式数据
    ' 入参 : lngEInvoiceID -电子票据使用记录.ID
    '        strEInvoiceClientCode-开票点编号
    ' 出参 : strJson-挂号结算信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, strWhere As String, rsTmp As ADODB.Recordset
    Dim cllInsureInfo As Collection
    Dim bln补结算 As Boolean
    Dim bytInvoiceType As Byte
    Dim dbl误差费 As Double
    Dim lng病人ID As Long, lng挂号ID As Long, lng医嘱序号 As Long
    Dim str业务标识 As String, strChargeDetail As String, strListDetail As String
    Dim str结帐IDs As String, str登记时间 As String, str业务操作员 As String
    Dim str患者姓名 As String, str患者性别 As String, str患者年龄 As String, str医疗付款方式编码 As String
    Dim str门诊号 As String
    Dim strJsonList As String, strData As String
    Dim strJsonKey_就诊科室 As String
    Dim strJsonFormat_就诊日期 As String
    Dim intJsonFormat_费用小数 As Integer, intJsonFormat_数量小数 As Integer
    On Error GoTo ErrHand
    bytInvoiceType = 1
    str业务标识 = zlGet业务标识(bytInvoiceType)
    bln补结算 = CheckBillExistReplenishData(lng结帐ID)
    dbl票据总金额 = 0
    
    '版本差异
    strJsonKey_就诊科室 = GetVersionDiff(1, "就诊科室")
    strJsonFormat_就诊日期 = GetVersionDiff(2, "就诊日期")
    intJsonFormat_费用小数 = Val(GetVersionDiff(2, "费用小数"))
    intJsonFormat_数量小数 = Val(GetVersionDiff(2, "数量小数"))
    
    If bln补结算 Then
        strWhere = "Select Distinct a.NO From 门诊费用记录 a, 费用补充记录 b Where a.结帐ID = b.收费结帐ID and b.结算id = [1]"
    Else
        strWhere = "Select Distinct NO From 门诊费用记录 Where 结帐id = [1]"
    End If
    strSQL = "Select Min(a.Id) As 费用id, a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, Max(a.计算单位) As 计算单位," & vbNewLine & _
            "        Sum(a.标准单价) As 价格, Avg(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 数量, Sum(a.应收金额) As 应收金额," & vbNewLine & _
            "        Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.实收金额) - Sum(a.统筹金额) As 自费金额," & vbNewLine & _
            "        Max(s.大类编码) As 医保项目编码, Max(s.大类名称) As 医保项目名称, Max(t.统筹比额) As 医保报销比例, Max(a.摘要) As 备注," & vbNewLine & _
            "        Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.姓名) As 姓名," & vbNewLine & _
            "        Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间," & vbNewLine & _
            "        Max(a.付款方式) As 付款方式编码, Max(Nvl(c.名称, c1.名称)) As 收据费目, Max(Nvl(c.编码, c1.编码)) As 收据费目编码, Max(a.医嘱序号) As 医嘱序号," & vbNewLine & _
            "        Max(a.挂号id) As 挂号id, Max(d.编码) As 类别编码, Max(d.类别) As 类别名称, Max(b.编码) As 项目编码, Max(b.名称) As 项目名称," & vbNewLine & _
            "        Max(b.规格) As 规格, Max(q.药品剂型) As 药品剂型" & vbNewLine & _
            " From 门诊费用记录 A, 收费项目目录 B, 收据费目对照 C, 收据费目 C1, 收费类别 D, 药品规格 M, 药品特性 Q, 诊疗项目目录 J, 保险支付大类 T, 支付类别对照 S" & vbNewLine & _
            " Where a.No In (" & strWhere & ") And Mod(a.记录性质, 10) = 1 And" & vbNewLine & _
            "       a.收费类别 = d.编码(+) And a.收费细目id = b.Id And a.收据费目 = c1.名称(+) And a.收据费目 = c.收据费目(+) and Decode(c.费用场合(+), 0, 1, c.费用场合(+)) = 1 And a.收费细目id = m.药品id(+) And" & vbNewLine & _
            "       m.药名id = q.药名id(+) And q.药名id = j.Id(+) And a.保险大类id = t.Id(+) And t.性质(+) = 1 And" & vbNewLine & _
            "       a.保险大类id = s.保险大类id(+)" & vbNewLine & _
            " Group By a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号), a.收费细目id, c.编码, c.名称, j.编码, j.名称" & vbNewLine & _
            " Order By NO, 序号"
    strSQL = "Select Min(a.费用ID) As 费用ID, a.No, a.序号, a.收费细目id, a.计算单位, a.价格, Sum(a.数量) As 数量," & vbNewLine & _
            "       Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.自费金额) As 自费金额," & vbNewLine & _
            "       Max(a.医保项目名称) As 医保项目编码, Max(a.医保项目名称) As 医保项目名称, Max(a.医保报销比例) As 医保报销比例, Max(a.备注) As 备注," & vbNewLine & _
            "       Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.姓名) As 姓名," & vbNewLine & _
            "       Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间," & vbNewLine & _
            "       Max(a.付款方式编码) As 付款方式编码, Max(a.收据费目) As 收据费目, Max(a.收据费目编码) As 收据费目编码, Max(a.医嘱序号) As 医嘱序号," & vbNewLine & _
            "       Max(a.挂号id) As 挂号id, Max(a.类别编码) As 类别编码, Max(a.类别名称) As 类别名称, Max(a.项目编码) As 项目编码, Max(a.项目名称) As 项目名称," & vbNewLine & _
            "       Max(a.规格) As 规格, Max(a.药品剂型) As 药品剂型" & vbNewLine & _
            "From (" & strSQL & ") a" & vbNewLine & _
            "Group By a.No, a.序号, a.收费细目id, a.计算单位, a.价格" & vbNewLine & _
            IIf(gBs_Type.零费用开票, "", " Having Sum(a.结帐金额) <> 0")
        
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng结帐ID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "未找到明细数据，不能打印电子票据"
        Exit Function
    End If
    
    strJsonList = ""
    With rsTmp
        str患者姓名 = Nvl(!姓名)
        str患者性别 = Nvl(!性别)
        str患者年龄 = Nvl(!年龄)
        lng病人ID = Val(Nvl(!病人ID))
        str医疗付款方式编码 = Nvl(!付款方式编码)
        lng挂号ID = Val(Nvl(!挂号id))
        lng医嘱序号 = Val(Nvl(!医嘱序号))
        str登记时间 = Format(Nvl(!登记时间), "YYYYMMDDhhmmss000")
        str业务操作员 = Nvl(!操作员姓名)
        
        Do While Not .EOF
            strData = ""
            strData = strData & "" & GetJsonNodeString("listDetailNo", zlStr.LPAD(Nvl(!费用ID), 20, "0"), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!收据费目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!收据费目), Json_Text)
            strData = strData & "," & GetJsonNodeString("prescribeCode", Nvl(!No), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeCode", Nvl(!类别编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeName", Nvl(!类别名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("code", Nvl(!项目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("name", Nvl(!项目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("form", Nvl(!药品剂型), Json_Text)
            strData = strData & "," & GetJsonNodeString("specification", Nvl(!规格), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!计算单位), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!价格)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!数量)), intJsonFormat_数量小数), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!实收金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!自费金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("receivableAmt", FormatEx(Val(Nvl(!应收金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!医保项目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("medCareItemType", Nvl(!医保项目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("medReimburseRate", FormatEx(Val(Nvl(!医保报销比例)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!备注), Json_Text)
            strData = strData & "," & GetJsonNodeString("sortNo", Nvl(!序号), Json_num)
            strData = strData & "," & GetJsonNodeString("chrgtype", Nvl(!费用类型), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dbl票据总金额 = dbl票据总金额 + RoundEx(Nvl(!实收金额), 6)
            .MoveNext
        Loop
        
        str结帐IDs = GetBalanceIDs(lng结帐ID, bytInvoiceType)
        dbl误差费 = GetBalanceErrorFee(str结帐IDs)
        strListDetail = GetNodeString("listDetail") & ":[" & Mid(strJsonList, 2) & "]"
    End With
    
    '分类明细
    If gBs_Type.误差费对照编码 <> "" Then
        dbl票据总金额 = dbl票据总金额 - dbl误差费
    End If
    dbl票据总金额 = RoundEx(dbl票据总金额, 2)
    If Not Get分类明细(str结帐IDs, strData, dbl票据总金额, bytInvoiceType, strErrMsg_Out) Then Exit Function
    strChargeDetail = GetNodeString("chargeDetail") & ":[" & strData & "]"
    
    '票据信息
    '业务流水号:lng结帐ID_lngEInvoiceID
    strData = ""
    strData = strData & "" & GetJsonNodeString("busNo", lng结帐ID & "_" & lngEInvoiceID, Json_Text)
    strData = strData & "," & GetJsonNodeString("busType", str业务标识, Json_Text)
    strData = strData & "," & GetJsonNodeString("payer", str患者姓名, Json_Text)
    strData = strData & "," & GetJsonNodeString("busDateTime", str登记时间, Json_Text)
    strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strData = strData & "," & GetJsonNodeString("payee", str业务操作员, Json_Text)
    strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("checker", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("totalAmt", dbl票据总金额, Json_num)
    strData = strData & "," & GetJsonNodeString("remark", IIf(RoundEx(dbl误差费, 6) <> 0 <> 0 And gBs_Type.误差费对照编码 = "", "存在" & FormatEx(dbl误差费, 6) & "误差金额不参与结算", ""), Json_Text)
    strJson_Out = strData
    
    
    '移动支付(一致)
    If Not Get移动支付信息(lng病人ID, IIf(bln补结算, str结帐IDs, lng结帐ID), strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '通知消息
    If Not Get通知消息(lng病人ID, strData, str门诊号) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '就诊信息
    Call Get医保信息(bytInvoiceType, lng结帐ID, lng病人ID, cllInsureInfo)
    Set rsTmp = Nothing
    If lng医嘱序号 <> 0 Then
        strSQL = "Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')) As 就诊日期, Max(b.编码) As 就诊科室编码," & vbNewLine & _
                "       Max(b.名称) As 就诊科室名称, Max(a.No) As 就诊编号, Max(d.编码) As 疾病编码" & vbNewLine & _
                "  From 病人挂号记录 A, 部门表 B, 病人诊断记录 C, 疾病编码目录 D" & vbNewLine & _
                "  Where a.执行部门id = b.Id And " & vbNewLine & _
                "   a.病人ID = c.病人ID(+) And a.ID = c.主页ID(+) And c.诊断次序(+) = 1 and Mod(c.诊断类型(+), 10) = 1 And c.疾病ID = d.id(+) And " & vbNewLine & _
                "   a.No = (Select Max(挂号单) From 病人医嘱记录 Where ID = [1] Or 相关id = [1])"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng医嘱序号)
    ElseIf lng挂号ID <> 0 Then
        strSQL = "Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')) As 就诊日期, Max(b.编码) As 就诊科室编码," & vbNewLine & _
                "       Max(b.名称) As 就诊科室名称, Max(a.No) As 就诊编号, Max(d.编码) As 疾病编码" & vbNewLine & _
                "  From 病人挂号记录 A, 部门表 B, 病人诊断记录 C, 疾病编码目录 D" & vbNewLine & _
                "  Where a.执行部门id = b.Id And a.Id = [1] And " & vbNewLine & _
                "   a.病人ID = c.病人ID(+) And a.ID = c.主页ID(+) And c.诊断次序(+) = 1 and Mod(c.诊断类型(+), 10) = 1 And c.疾病ID = d.id(+) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng挂号ID)
    End If
    If rsTmp Is Nothing Then
        strSQL = "Select To_Char(a.发生时间, 'yyyy-mm-dd') As 就诊日期, b.编码 As 就诊科室编码," & vbNewLine & _
                "       b.名称 As 就诊科室名称, a.No As 就诊编号, d.编码 As 疾病编码" & vbNewLine & _
                "  From 病人挂号记录 A, 部门表 B, 病人诊断记录 C, 疾病编码目录 D" & vbNewLine & _
                "  Where a.执行部门id = b.Id And " & vbNewLine & _
                "       a.病人ID = c.病人ID(+) And a.ID = c.主页ID(+) And c.诊断次序(+) = 1 and Mod(c.诊断类型(+), 10) = 1 And c.疾病ID = d.id(+) And " & vbNewLine & _
                "       a.Id = (Select ID" & vbNewLine & _
                "           From (Select ID, 发生时间 From 病人挂号记录 Where 病人id = [1] Order By 发生时间 Desc)" & vbNewLine & _
                "           Where Rownum < 2)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng病人ID)
    End If
    
    strData = ""
    strData = strData & "" & GetJsonNodeString("medicalInstitution", GetUnitInfo("医疗机构类型"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareInstitution", zlGetNodeValueFromCollect(cllInsureInfo, "_保险机构编码", "C"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareTypeCode", str医疗付款方式编码, Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalCareType", Get医疗付款方式名称(str医疗付款方式编码), Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalInsuranceID", zlGetNodeValueFromCollect(cllInsureInfo, "_医保号", "C"), Json_Text)
    With rsTmp
        If .RecordCount > 0 Then
            strData = strData & "," & GetJsonNodeString("consultationDate", Format(Nvl(!就诊日期), strJsonFormat_就诊日期), Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_就诊科室, Nvl(!就诊科室名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("patientCategoryCode", Nvl(!就诊科室编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!就诊编号), Json_Text)
        Else
            strData = strData & "," & GetJsonNodeString("consultationDate", "", Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_就诊科室, "", Json_Text)
            strData = strData & "," & GetJsonNodeString("patientCategoryCode", "", Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", lng结帐ID, Json_Text)
        End If
    End With
    strData = strData & "," & GetJsonNodeString("patientId", lng病人ID, Json_Text)
    strData = strData & "," & GetJsonNodeString("sex", str患者性别, Json_Text)
    strData = strData & "," & GetJsonNodeString("age", str患者年龄, Json_Text)
    strData = strData & "," & GetJsonNodeString("caseNumber", str门诊号, Json_Text)
    strData = strData & "," & GetJsonNodeString("ICD", Nvl(rsTmp!疾病编码), Json_Text)
    strData = strData & "," & GetJsonNodeString("specialDiseasesName", zlGetNodeValueFromCollect(cllInsureInfo, "_病种名称", "C"), Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '支付信息
    If Not Get结算信息(str结帐IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '缴费渠道
    If Not Get缴费渠道(str结帐IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & GetNodeString("payChannelDetail") & ":[" & strData & "]"
    
    '其它医保信息-暂无
    '其它扩展信息-暂无
    'eBillRelateNo  业务票据关联号  String  32  否  如一笔业务数据需要开具N张电子票据，则N张电子票对应该值保持一致，用于后期关联查询
    'isArrears  是否可流通  String  1  是  0-否、1-是（如欠费情况根据医院业务要求该票据是否可流通）
    'arrearsReason  不可流通原因  String  200  否  isArrears=0，填写不可流通的原因
    strData = ""
    strData = strData & "" & GetJsonNodeString("eBillRelateNo", "", Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '收费项目明细
    strJson_Out = strJson_Out & "," & strChargeDetail
    '清单项目明细
    strJson_Out = strJson_Out & "," & strListDetail
    
    '返回完整的Json串
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceByCharge = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceByDeposit(ByVal lngEInvoiceID As Long, ByVal lng预交ID As Long, ByVal lng冲销ID As Long, _
                ByVal strEInvoiceClientCode As String, dbl票据总额 As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取挂号发票Json格式数据
    ' 入参 : lngEInvoiceID -电子票据使用记录.ID
    '        strEInvoiceClientCode-开票点编号
    '        lng冲销ID-冲销预交ID
    ' 出参 : strJson-挂号结算信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsDeposit As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim bytInvoiceType As Byte
    Dim dbl预交余额 As Double
    Dim lng病人ID As Long
    Dim str业务标识 As String, str登记时间 As String
    Dim strJsonList As String, strData As String, strChargeDetail As String
    Dim str卡名称 As String, str卡号 As String, strNO As String
    On Error GoTo ErrHand
    bytInvoiceType = 2
    str业务标识 = zlGet业务标识(bytInvoiceType)
    
    strSQL = "Select a.No, a.收款时间, a.预交类别, a.卡类别id, a.病人id, a.主页id, a.科室id, a.缴款单位, a.单位开户行, a.单位帐号, a.摘要, a.结算方式, a.结算号码, a.卡号," & vbNewLine & _
            "       a.交易流水号, a.交易说明, a.合作单位, a.金额, a.操作员编号, a.操作员姓名, Nvl(b.姓名, c.姓名) As 姓名, Nvl(b.性别, c.性别) As 性别," & vbNewLine & _
            "       Nvl(b.年龄, c.年龄) As 年龄, c.门诊号, Nvl(b.住院号, c.住院号) As 住院号, c.Email, c.身份证号, c.手机号, 1 As 缴款类型," & vbNewLine & _
            "       Decode(Nvl(a.预交类别, 0), 1, '07', '07') As 业务标识, d.编码 As 入院科室编码, d.名称 As 入院科室名称, e.编码 As 出院科室编码," & vbNewLine & _
            "       e.名称 As 出院科室名称, b.入院日期, b.出院日期, Nvl(b.病案号, b.住院号) As 病历号, j.名称 As 医疗卡名称" & vbNewLine & _
            "From 病人预交记录 A, 病案主页 B, 病人信息 C, 部门表 D, 部门表 E, 医疗卡类别 J" & vbNewLine & _
            "Where a.Id = [1] And a.病人id = b.病人id(+) And a.主页id = b.主页id(+) And a.病人id = c.病人id(+) And b.入院科室id = d.Id(+) And" & vbNewLine & _
            "      b.出院科室id = e.Id(+) And a.卡类别id = j.Id(+)"

    Set rsDeposit = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByDeposit", lng预交ID)
    If rsDeposit.RecordCount = 0 Then
        strErrMsg_Out = "未找到结算数据，不能打印电子票据"
        Exit Function
    End If
    
    With rsDeposit
        If Nvl(!预交类别) = 1 Then
            strErrMsg_Out = "博思接口不支持门诊预交票据生成电子票据！"
        End If
        lng病人ID = Val(Nvl(!病人ID))
        strNO = Nvl(!No)
        str登记时间 = Format(Nvl(!收款时间), "YYYYMMDDhhmmss000")
    End With
    dbl票据总额 = Get预交单据总额(strNO)
    dbl预交余额 = Get预交余额(lng病人ID, Val(Nvl(rsDeposit!预交类别)))
    
    If lng冲销ID <> 0 Then
        strSQL = "Select 代码, 号码, 凭证代码, 凭证号码" & vbNewLine & _
                " From 电子票据使用记录" & vbNewLine & _
                " Where ID = [1] And 退款id Is Null And 记录状态 = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByDeposit", lngEInvoiceID)
        If rsTmp.EOF Then
            strErrMsg_Out = "原始预交款未开具电子票据凭证，不允行开具退款票据！"
            Exit Function
        End If
        
        strData = ""
        strData = strData & "" & GetJsonNodeString("busType", str业务标识, Json_Text)
        strData = strData & "," & GetJsonNodeString("billBatchCode", Nvl(rsTmp!代码), Json_Text)
        strData = strData & "," & GetJsonNodeString("reason", "退款", Json_Text)
        strData = strData & "," & GetJsonNodeString("operator", mstrOperatorName, Json_Text)
        strData = strData & "," & GetJsonNodeString("busDateTime", str登记时间, Json_Text)
        strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
        strData = strData & "," & GetJsonNodeString("voucherBatchCode", Nvl(rsTmp!凭证代码), Json_Text)
        strData = strData & "," & GetJsonNodeString("voucherNo", Nvl(rsTmp!凭证号码), Json_Text)
        strData = strData & "," & GetJsonNodeString("amt", -1 * dbl票据总额, Json_num)
        strData = strData & "," & GetJsonNodeString("ownAcBalance", dbl预交余额, Json_num)
        strData = strData & "," & GetJsonNodeString("remark", Nvl(rsDeposit!摘要), Json_Text)
        '返回完整的Json串
        strJson_Out = "{" & strJson_Out & "}"
        zlGetJson_CreateEInvoiceByDeposit = True
        Exit Function
    End If

    '缴费渠道
    Call Get通知卡号(lng病人ID, Nvl(rsDeposit!身份证号), str卡名称, str卡号)
    strSQL = "Select c.渠道编码" & vbNewLine & _
            "  From 收费渠道对照 C" & vbNewLine & _
            "  Where c.卡类别id = [1] And c.结算方式 = [2]" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select c.渠道编码" & vbNewLine & _
            "  From 收费渠道对照 C" & vbNewLine & _
            "  Where c.卡类别id Is Null And c.结算方式 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByDeposit", Val(Nvl(rsDeposit!卡类别ID)), Nvl(rsDeposit!结算方式))
    strData = ""
    If rsTmp.RecordCount > 0 Then
        strData = Nvl(rsTmp!渠道编码)
    End If
    strData = GetJsonNodeString("payChannelCode", strData, Json_Text)
    strData = strData & "," & GetJsonNodeString("payChannelValue", dbl票据总额, Json_num)
    strJsonList = "{" & strData & "}"
    strJson_Out = GetNodeString("payChannelDetail") & ":[" & strJsonList & "]"
    
    With rsDeposit
        strData = ""
        strData = strData & "" & GetJsonNodeString("busType", str业务标识, Json_Text)
        '结算ID_电子票据ID
        strData = strData & "," & GetJsonNodeString("busNo", lng预交ID & "_" & lngEInvoiceID, Json_Text)
        strData = strData & "," & GetJsonNodeString("payer", Nvl(!姓名), Json_Text)
        strData = strData & "," & GetJsonNodeString("busDateTime", str登记时间, Json_Text)
        strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
        strData = strData & "," & GetJsonNodeString("payee", Nvl(!操作员姓名), Json_Text)
        strData = strData & "," & GetJsonNodeString("drawee", Nvl(!姓名), Json_Text)
        strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
        strData = strData & "," & GetJsonNodeString("tel", Nvl(!手机号), Json_Text)
        strData = strData & "," & GetJsonNodeString("email", Nvl(!email), Json_Text)
        strData = strData & "," & GetJsonNodeString("idCardNo", Nvl(!身份证号), Json_Text)
        strData = strData & "," & GetJsonNodeString("cardType", str卡名称, Json_Text)
        strData = strData & "," & GetJsonNodeString("cardNo", str卡号, Json_Text)
        strData = strData & "," & GetJsonNodeString("amt", dbl票据总额, Json_num)
        strData = strData & "," & GetJsonNodeString("ownAcBalance", dbl预交余额, Json_num)
        strData = strData & "," & GetJsonNodeString("category", Nvl(!入院科室名称), Json_Text)
        strData = strData & "," & GetJsonNodeString("categoryCode", Nvl(!入院科室编码), Json_Text)
        strData = strData & "," & GetJsonNodeString("inHospitalDate", Format(Nvl(!入院日期), "yyyy-MM-dd"), Json_Text)
        strData = strData & "," & GetJsonNodeString("hospitalNo", Nvl(!住院号), Json_Text)
        strData = strData & "," & GetJsonNodeString("patientId", Nvl(!病人ID), Json_Text)
        strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!主页id), Json_Text)
        strData = strData & "," & GetJsonNodeString("caseNumber", Nvl(!病历号), Json_Text)
        strData = strData & "," & GetJsonNodeString("accountName", Nvl(!医疗卡名称), Json_Text)
        strData = strData & "," & GetJsonNodeString("accountNo", Nvl(!单位帐号), Json_Text)
        strData = strData & "," & GetJsonNodeString("accountBank", IIf(Nvl(!医疗卡名称) <> "", Nvl(!医疗卡名称), Nvl(!单位开户行)), Json_Text)
        strData = strData & "," & GetJsonNodeString("remark", Nvl(!摘要), Json_Text)
        If gBs_Type.支持版本 > BS_Version.V2_0_3 Then
            strData = strData & "," & GetJsonNodeString("workUnit", Nvl(!缴款单位), Json_Text)
        End If
        strJson_Out = strJson_Out & "," & strData
    End With
    
    '返回完整的Json串
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceByDeposit = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceByMzBalance(ByVal lngEInvoiceID As Long, ByVal lng结帐ID As Long, ByVal lng销账ID As Long, _
                ByVal strEInvoiceClientCode As String, dbl票据总金额 As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取挂号发票Json格式数据
    ' 入参 : lngEInvoiceID -电子票据使用记录.ID
    '        strEInvoiceClientCode-开票点编号
    ' 出参 : strJson-挂号结算信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsBalance As ADODB.Recordset
    Dim cllInsureInfo As Collection
    Dim bytInvoiceType As Byte
    Dim dbl误差费 As Double
    Dim lng病人ID As Long, lng挂号ID As Long, lng医嘱序号 As Long
    Dim str卡名称 As String, str卡号 As String, strChargeDetail As String, strListDetail As String
    Dim str患者姓名 As String, str患者性别 As String, str患者年龄 As String, str医疗付款方式编码 As String
    Dim str门诊号 As String
    Dim strJsonList As String, strData As String
    Dim strJsonKey_就诊科室 As String
    Dim strJsonFormat_就诊日期 As String
    Dim intJsonFormat_费用小数 As Integer, intJsonFormat_数量小数 As Integer
    On Error GoTo ErrHand
    bytInvoiceType = 3
    dbl误差费 = GetBalanceErrorFee(lng结帐ID)
    dbl票据总金额 = 0
    
    '版本差异
    strJsonKey_就诊科室 = GetVersionDiff(1, "就诊科室")
    strJsonFormat_就诊日期 = GetVersionDiff(2, "就诊日期")
    intJsonFormat_费用小数 = Val(GetVersionDiff(2, "费用小数"))
    intJsonFormat_数量小数 = Val(GetVersionDiff(2, "数量小数"))
    
    strSQL = "Select a.No, a.收费时间, a.结帐类型, a.操作员编号, a.操作员姓名, a.病人id, a.主页id, Decode(Nvl(a.病人id, 0), 0, a.原因, c.姓名) As 姓名," & vbNewLine & _
            "       '' As 性别, '' As 年龄, c.门诊号, a.备注, a.结帐金额, Decode(Nvl(a.病人id, 0), 0, q.电子邮件, c.Email) As Email, q.联系人," & vbNewLine & _
            "       Decode(Nvl(a.病人id, 0), 0, q.社会信用代码, c.身份证号) As 身份证号," & vbNewLine & _
            "       Decode(Nvl(a.病人id, 0), 0, Nvl(q.电话, To_Char(j.移动电话)), c.手机号) As 手机号," & vbNewLine & _
            "       Decode(Nvl(a.病人id, 0), 0, 2, 1) As 缴款类型, Decode(Nvl(a.结帐类型, 0), 1, '02', '01') As 业务标识, c.门诊号 As 病历号" & vbNewLine & _
            "From 病人结帐记录 A, 病人信息 C, 合约单位 Q, 人员表 J" & vbNewLine & _
            "Where a.Id = [1] And a.病人id = c.病人id(+) And a.原因 = q.名称(+) And q.联系人 = j.姓名(+)"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByMzBalance", lng结帐ID)
    If rsBalance.RecordCount = 0 Then
        strErrMsg_Out = "未找到结帐数据，不能打印电子票据"
        Exit Function
    End If
    
    strSQL = "      Select Min(a.Id) As 费用id, a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, Max(a.计算单位) As 计算单位," & vbNewLine & _
            "              Sum(a.标准单价) As 价格, Avg(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 数量, Sum(a.应收金额) As 应收金额," & vbNewLine & _
            "              Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.实收金额) - Sum(a.统筹金额) As 自费金额," & vbNewLine & _
            "              Max(s.大类编码) As 医保项目编码, Max(s.大类名称) As 医保项目名称, Max(t.统筹比额) As 医保报销比例, Max(a.摘要) As 备注," & vbNewLine & _
            "              Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.姓名) As 姓名," & vbNewLine & _
            "              Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间," & vbNewLine & _
            "              Max(a.付款方式) As 付款方式编码, Max(Nvl(c.名称, c1.名称)) As 收据费目, Max(Nvl(c.编码, c1.编码)) As 收据费目编码, Max(a.医嘱序号) As 医嘱序号," & vbNewLine & _
            "              Max(a.挂号id) As 挂号id, Max(d.编码) As 类别编码, Max(d.类别) As 类别名称, Max(b.编码) As 项目编码, Max(b.名称) As 项目名称," & vbNewLine & _
            "              Max(b.规格) As 规格, Max(q.药品剂型) As 药品剂型" & vbNewLine & _
            "       From 门诊费用记录 A, 收费项目目录 B, 收据费目对照 C, 收据费目 C1, 收费类别 D, 药品规格 M, 药品特性 Q, 诊疗项目目录 J, 保险支付大类 T, 支付类别对照 S" & vbNewLine & _
            "       Where a.no In (Select No From 门诊费用记录 Where 结帐ID = [1]) And a.记帐费用 = 1 And a.收费类别 = d.编码(+) And a.收费细目id = b.Id And " & vbNewLine & _
            "             a.收据费目 = c1.名称(+) And a.收据费目 = c.收据费目(+) And Decode(c.费用场合(+), 0, 1, c.费用场合(+)) = 1 and" & vbNewLine & _
            "             a.收费细目id = m.药品id(+) And m.药名id = q.药名id(+) And q.药名id = j.Id(+) And a.保险大类id = t.Id(+) And" & vbNewLine & _
            "             t.性质(+) = 1 And a.保险大类id = s.保险大类id(+)" & vbNewLine & _
            "       Group By a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号), a.收费细目id, c.编码, c.名称, j.编码, j.名称" & vbNewLine & _
            "       Order By a.NO, 序号"
    
    strSQL = "Select Min(a.费用ID) As 费用ID, a.No, a.序号, a.收费细目id, a.计算单位, Avg(a.价格) As 价格, Avg(a.数量) As 数量," & vbNewLine & _
            "       Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.自费金额) As 自费金额," & vbNewLine & _
            "       Max(a.医保项目名称) As 医保项目编码, Max(a.医保项目名称) As 医保项目名称, Max(a.医保报销比例) As 医保报销比例, Max(a.备注) As 备注," & vbNewLine & _
            "       Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.姓名) As 姓名," & vbNewLine & _
            "       Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间," & vbNewLine & _
            "       Max(a.付款方式编码) As 付款方式编码, Max(a.收据费目) As 收据费目, Max(a.收据费目编码) As 收据费目编码, Max(a.医嘱序号) As 医嘱序号," & vbNewLine & _
            "       Max(a.挂号id) As 挂号id, Max(a.类别编码) As 类别编码, Max(a.类别名称) As 类别名称, Max(a.项目编码) As 项目编码, Max(a.项目名称) As 项目名称," & vbNewLine & _
            "       Max(a.规格) As 规格, Max(a.药品剂型) As 药品剂型" & vbNewLine & _
            " From (" & strSQL & ") a" & vbNewLine & _
            " Group By a.No, a.序号, a.收费细目id, a.计算单位" & vbNewLine & _
            IIf(gBs_Type.零费用开票, "", " Having Sum(a.结帐金额) <> 0")

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByMzBalance", lng结帐ID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "未找到明细数据，不能打印电子票据"
        Exit Function
    End If
    
    strJsonList = ""
    With rsTmp
        str患者姓名 = Nvl(!姓名)
        str患者性别 = Nvl(!性别)
        str患者年龄 = Nvl(!年龄)
        lng病人ID = Val(Nvl(!病人ID))
        str医疗付款方式编码 = Nvl(!付款方式编码)
        lng挂号ID = Val(Nvl(!挂号id))
        lng医嘱序号 = Val(Nvl(!医嘱序号))
        
        Do While Not .EOF
            strData = ""
            strData = strData & "" & GetJsonNodeString("listDetailNo", zlStr.LPAD(Nvl(!费用ID), 20, "0"), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!收据费目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!收据费目), Json_Text)
            strData = strData & "," & GetJsonNodeString("prescribeCode", Nvl(!No), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeCode", Nvl(!类别编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeName", Nvl(!类别名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("code", Nvl(!项目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("name", Nvl(!项目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("form", Nvl(!药品剂型), Json_Text)
            strData = strData & "," & GetJsonNodeString("specification", Nvl(!规格), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!计算单位), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!价格)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!数量)), intJsonFormat_数量小数), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!实收金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!自费金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("receivableAmt", FormatEx(Val(Nvl(!应收金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!医保项目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("medCareItemType", Nvl(!医保项目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("medReimburseRate", FormatEx(Val(Nvl(!医保报销比例)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!备注), Json_Text)
            strData = strData & "," & GetJsonNodeString("sortNo", Nvl(!序号), Json_num)
            strData = strData & "," & GetJsonNodeString("chrgtype", Nvl(!费用类型), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dbl票据总金额 = dbl票据总金额 + RoundEx(Val(Nvl(!实收金额)), 6)
            .MoveNext
        Loop
        strListDetail = GetNodeString("listDetail") & ":[" & Mid(strJsonList, 2) & "]"
    End With

    '分类明细
    If gBs_Type.误差费对照编码 <> "" Then
        dbl票据总金额 = dbl票据总金额 - dbl误差费
    End If
    dbl票据总金额 = RoundEx(dbl票据总金额, 2)
    If Not Get分类明细(lng结帐ID, strData, dbl票据总金额, 3.1, strErrMsg_Out) Then Exit Function
    strChargeDetail = GetNodeString("chargeDetail") & ":[" & strData & "]"
    
    '票据信息
    '业务流水号:lng结帐ID_lngEInvoiceID
    strData = ""
    strData = strData & "" & GetJsonNodeString("busNo", lng结帐ID & "_" & lngEInvoiceID, Json_Text)
    strData = strData & "," & GetJsonNodeString("busType", Nvl(rsBalance!业务标识), Json_Text)
    If Val(Nvl(rsBalance!病人ID)) = 0 Then
        strData = strData & "," & GetJsonNodeString("payer", Nvl(rsBalance!姓名), Json_Text)
    Else
        strData = strData & "," & GetJsonNodeString("payer", str患者姓名, Json_Text)
    End If
    strData = strData & "," & GetJsonNodeString("busDateTime", Format(Nvl(rsBalance!收费时间), "YYYYMMDDhhmmss000"), Json_Text)
    strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strData = strData & "," & GetJsonNodeString("payee", Nvl(rsBalance!操作员姓名), Json_Text)
    strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("checker", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("totalAmt", dbl票据总金额, Json_num)
    strData = strData & "," & GetJsonNodeString("remark", IIf(RoundEx(dbl误差费, 6) <> 0 And gBs_Type.误差费对照编码 = "", "存在" & FormatEx(dbl误差费, 6) & "误差金额不参与结算", Nvl(rsBalance!备注)), Json_Text)
    strJson_Out = strData
    
    
    '移动支付(一致)
    If Not Get移动支付信息(lng病人ID, lng结帐ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '通知消息
    With rsBalance
        Call Get通知卡号(Val(Nvl(!病人ID)), Nvl(!身份证号), str卡名称, str卡号)
        strData = ""
        strData = strData & "" & GetJsonNodeString("tel", Nvl(!手机号), Json_Text)
        strData = strData & "," & GetJsonNodeString("email", Nvl(!email), Json_Text)
        If gBs_Type.支持版本 > BS_Version.V2_0_3 Then
            strData = strData & "," & GetJsonNodeString("payerType", Nvl(!缴款类型), Json_Text)
        End If
        strData = strData & "," & GetJsonNodeString("idCardNo", Nvl(!身份证号), Json_Text)
        strData = strData & "," & GetJsonNodeString("cardType", str卡名称, Json_Text)
        strData = strData & "," & GetJsonNodeString("cardNo", str卡号, Json_Text)
        strJson_Out = strJson_Out & "," & strData
    End With
    
    '就诊信息
    With rsBalance
        Call Get医保信息(bytInvoiceType, lng结帐ID, Val(Nvl(!病人ID)), cllInsureInfo)
        Set rsTmp = Nothing
        If lng医嘱序号 <> 0 Then
            strSQL = "Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')) As 就诊日期, Max(b.编码) As 就诊科室编码," & vbNewLine & _
                    "       Max(b.名称) As 就诊科室名称, Max(a.No) As 就诊编号, Max(d.编码) As 疾病编码" & vbNewLine & _
                    "  From 病人挂号记录 A, 部门表 B, 病人诊断记录 C, 疾病编码目录 D" & vbNewLine & _
                    "  Where a.执行部门id = b.Id And " & vbNewLine & _
                    "   a.病人ID = c.病人ID(+) And a.ID = c.主页ID(+) And c.诊断次序(+) = 1 and Mod(c.诊断类型(+), 10) = 1 And c.疾病ID = d.id(+) And " & vbNewLine & _
                    "   a.No = (Select Max(挂号单) From 病人医嘱记录 Where ID = [1] Or 相关id = [1])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng医嘱序号)
        ElseIf lng挂号ID <> 0 Then
            strSQL = "Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')) As 就诊日期, Max(b.编码) As 就诊科室编码," & vbNewLine & _
                    "       Max(b.名称) As 就诊科室名称, Max(a.No) As 就诊编号, Max(d.编码) As 疾病编码" & vbNewLine & _
                    "  From 病人挂号记录 A, 部门表 B, 病人诊断记录 C, 疾病编码目录 D" & vbNewLine & _
                    "  Where a.执行部门id = b.Id And a.Id = [1] And " & vbNewLine & _
                    "   a.病人ID = c.病人ID(+) And a.ID = c.主页ID(+) And c.诊断次序(+) = 1 and Mod(c.诊断类型(+), 10) = 1 And c.疾病ID = d.id(+) "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng挂号ID)
        End If
        If rsTmp Is Nothing Then
            strSQL = "Select To_Char(a.发生时间, 'yyyy-mm-dd') As 就诊日期, b.编码 As 就诊科室编码," & vbNewLine & _
                    "       b.名称 As 就诊科室名称, a.No As 就诊编号, d.编码 As 疾病编码" & vbNewLine & _
                    "  From 病人挂号记录 A, 部门表 B, 病人诊断记录 C, 疾病编码目录 D" & vbNewLine & _
                    "  Where a.执行部门id = b.Id And " & vbNewLine & _
                    "       a.病人ID = c.病人ID(+) And a.ID = c.主页ID(+) And c.诊断次序(+) = 1 and Mod(c.诊断类型(+), 10) = 1 And c.疾病ID = d.id(+) And " & vbNewLine & _
                    "       a.Id = (Select ID" & vbNewLine & _
                    "           From (Select ID, 发生时间 From 病人挂号记录 Where 病人id = [1] Order By 发生时间 Desc)" & vbNewLine & _
                    "           Where Rownum < 2)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", Val(Nvl(!病人ID)))
        End If
        
        strData = ""
        strData = strData & "" & GetJsonNodeString("medicalInstitution", GetUnitInfo("医疗机构类型"), Json_Text)
        strData = strData & "," & GetJsonNodeString("medCareInstitution", zlGetNodeValueFromCollect(cllInsureInfo, "_保险机构编码", "C"), Json_Text)
        strData = strData & "," & GetJsonNodeString("medCareTypeCode", str医疗付款方式编码, Json_Text)
        strData = strData & "," & GetJsonNodeString("medicalCareType", Get医疗付款方式名称(str医疗付款方式编码), Json_Text)
        strData = strData & "," & GetJsonNodeString("medicalInsuranceID", zlGetNodeValueFromCollect(cllInsureInfo, "_医保号", "C"), Json_Text)
        With rsTmp
            If .RecordCount > 0 Then
                strData = strData & "," & GetJsonNodeString("consultationDate", Format(Nvl(!就诊日期), strJsonFormat_就诊日期), Json_Text)
                strData = strData & "," & GetJsonNodeString(strJsonKey_就诊科室, Nvl(!就诊科室名称), Json_Text)
                strData = strData & "," & GetJsonNodeString("patientCategoryCode", Nvl(!就诊科室编码), Json_Text)
                strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!就诊编号), Json_Text)
            Else
                strData = strData & "," & GetJsonNodeString("consultationDate", "", Json_Text)
                strData = strData & "," & GetJsonNodeString(strJsonKey_就诊科室, "", Json_Text)
                strData = strData & "," & GetJsonNodeString("patientCategoryCode", "", Json_Text)
                strData = strData & "," & GetJsonNodeString("patientNo", lng结帐ID, Json_Text)
            End If
        End With
        strData = strData & "," & GetJsonNodeString("patientId", Nvl(!病人ID), Json_Text)
        If Val(Nvl(!病人ID)) = 0 Then
            strData = strData & "," & GetJsonNodeString("sex", Nvl(!性别), Json_Text)
            strData = strData & "," & GetJsonNodeString("age", Nvl(!年龄), Json_Text)
        Else
            strData = strData & "," & GetJsonNodeString("sex", str患者性别, Json_Text)
            strData = strData & "," & GetJsonNodeString("age", str患者年龄, Json_Text)
        End If
        strData = strData & "," & GetJsonNodeString("caseNumber", Nvl(!病历号), Json_Text)
        strData = strData & "," & GetJsonNodeString("ICD", Nvl(rsTmp!疾病编码), Json_Text)
        strData = strData & "," & GetJsonNodeString("specialDiseasesName", zlGetNodeValueFromCollect(cllInsureInfo, "_病种名称", "C"), Json_Text)
        
        strJson_Out = strJson_Out & "," & strData
    End With
    
    '支付信息
    If Not Get结算信息(lng结帐ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '缴费渠道
    If Not Get缴费渠道(lng结帐ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & GetNodeString("payChannelDetail") & ":[" & strData & "]"
    
    '其它医保信息-暂无
    '其它扩展信息-暂无
    'eBillRelateNo  业务票据关联号  String  32  否  如一笔业务数据需要开具N张电子票据，则N张电子票对应该值保持一致，用于后期关联查询
    'isArrears  是否可流通  String  1  是  0-否、1-是（如欠费情况根据医院业务要求该票据是否可流通）
    'arrearsReason  不可流通原因  String  200  否  isArrears=0，填写不可流通的原因
    strData = ""
    strData = strData & "" & GetJsonNodeString("eBillRelateNo", "", Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '收费项目明细
    strJson_Out = strJson_Out & "," & strChargeDetail
    '清单项目明细
    strJson_Out = strJson_Out & "," & strListDetail
    
    '返回完整的Json串
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceByMzBalance = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceByZyBalance(ByVal lngEInvoiceID As Long, ByVal lng结帐ID As Long, ByVal lng销账ID As Long, _
                ByVal strEInvoiceClientCode As String, dbl票据总金额 As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取挂号发票Json格式数据
    ' 入参 : lngEInvoiceID -电子票据使用记录.ID
    '        strEInvoiceClientCode-开票点编号
    ' 出参 : strJson-挂号结算信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsBalance As ADODB.Recordset
    Dim cllInsureInfo As Collection
    Dim bytInvoiceType As Byte
    Dim dbl误差费 As Double
    Dim str医疗付款方式编码 As String, str卡名称 As String, str卡号 As String
    Dim str门诊号 As String, str住院次数 As String
    Dim strJsonList As String, strData As String, strChargeDetail As String, strListDetail As String
    Dim strJsonKey_就诊科室 As String
    Dim strJsonFormat_就诊日期 As String
    Dim intJsonFormat_费用小数 As Integer, intJsonFormat_数量小数 As Integer
    On Error GoTo ErrHand
    bytInvoiceType = 3
    dbl误差费 = GetBalanceErrorFee(lng结帐ID)
    dbl票据总金额 = 0
    
    '版本差异
    strJsonKey_就诊科室 = GetVersionDiff(1, "就诊科室")
    strJsonFormat_就诊日期 = GetVersionDiff(2, "就诊日期")
    intJsonFormat_费用小数 = Val(GetVersionDiff(2, "费用小数"))
    intJsonFormat_数量小数 = Val(GetVersionDiff(2, "数量小数"))
    
    strSQL = "Select a.No, a.收费时间, a.结帐类型, a.操作员编号, a.操作员姓名, a.病人id, a.主页id," & vbNewLine & _
            "       Decode(Nvl(a.病人id, 0), 0, a.原因, Nvl(b.姓名, c.姓名)) As 姓名, Nvl(b.性别, c.性别) As 性别, Nvl(b.年龄, c.年龄) As 年龄, c.门诊号," & vbNewLine & _
            "       Nvl(b.住院号, c.住院号) As 住院号, a.开始日期, a.结束日期, a.备注, a.结帐金额, Decode(Nvl(a.病人id, 0), 0, q.电子邮件, c.Email) As Email," & vbNewLine & _
            "       q.联系人, Decode(Nvl(a.病人id, 0), 0, q.社会信用代码, c.身份证号) As 身份证号," & vbNewLine & _
            "       Decode(Nvl(a.病人id, 0), 0, Nvl(q.电话, To_Char(j.移动电话)), c.手机号) As 手机号," & vbNewLine & _
            "       Decode(Nvl(a.病人id, 0), 0, 2, 1) As 缴款类型, Decode(Nvl(a.结帐类型, 0), 1, '02', '01') As 业务标识, b.入院日期, b.出院日期," & vbNewLine & _
            "       m.编码 As 入院科室编码, m.名称 As 入院科室名称, p.编码 As 出院科室编码, p.名称 As 出院科室名称, b.出院病床 As 床号, t.名称 As 病区名称," & vbNewLine & _
            "       Nvl(b.病案号, b.住院号) As 病历号, Nvl(b.医疗付款方式, c.医疗付款方式) As 医疗付款方式, Nvl(b.出院日期, Sysdate) - b.入院日期 As 住院天数, f.编码 As 疾病编码" & vbNewLine & _
            "From 病人结帐记录 A, 病案主页 B, 病人信息 C, 合约单位 Q, 人员表 J, 部门表 M, 部门表 P, 部门表 T, 病人诊断记录 E, 疾病编码目录 F" & vbNewLine & _
            "Where a.Id = [1] And a.病人id = b.病人id(+) And a.主页id = b.主页id(+) And a.病人id = c.病人id(+) And a.原因 = q.名称(+) And" & vbNewLine & _
            "      a.病人ID = e.病人ID(+) And a.主页id = e.主页ID(+) And e.诊断次序(+) = 1 And Mod(e.诊断类型(+), 10) = 1 And e.疾病ID = f.id(+) And " & vbNewLine & _
            "      b.入院科室id = m.Id(+) And b.出院科室id = p.Id(+) And b.当前病区id = t.Id(+)" & vbNewLine & _
            "      And q.联系人 = j.姓名(+)"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByZyBalance", lng结帐ID)
    If rsBalance.RecordCount = 0 Then
        strErrMsg_Out = "未找到结帐数据，不能打印电子票据"
        Exit Function
    End If
    
    strSQL = "Select Min(a.Id) As 费用id, a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, Max(a.计算单位) As 计算单位," & vbNewLine & _
            "        Sum(a.标准单价) As 价格, Avg(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 数量, Sum(a.应收金额) As 应收金额," & vbNewLine & _
            "        Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.实收金额) - Sum(a.统筹金额) As 自费金额," & vbNewLine & _
            "        Max(s.大类编码) As 医保项目编码, Max(s.大类名称) As 医保项目名称, Max(t.统筹比额) As 医保报销比例, Max(a.摘要) As 备注," & vbNewLine & _
            "        Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.病人id) As 病人id," & vbNewLine & _
            "        Max(a.登记时间) As 登记时间, Max(Nvl(c.名称, c1.名称)) As 收据费目, Max(Nvl(c.编码, c1.编码)) As 收据费目编码, Max(a.主页id) As 主页id," & vbNewLine & _
            "        Max(d.编码) As 类别编码, Max(d.类别) As 类别名称, Max(b.编码) As 项目编码, Max(b.名称) As 项目名称, Max(b.规格) As 规格," & vbNewLine & _
            "        Max(q.药品剂型) As 药品剂型" & vbNewLine & _
            " From 住院费用记录 A, 收费项目目录 B, 收据费目对照 C, 收据费目 C1, 收费类别 D, 药品规格 M, 药品特性 Q, 诊疗项目目录 J, 保险支付大类 T, 支付类别对照 S" & vbNewLine & _
            " Where a.No In (Select Distinct NO From 住院费用记录 Where 结帐id = [1]) And a.记帐费用 = 1 And a.收费类别 = d.编码(+) And a.收费细目id = b.Id And" & vbNewLine & _
            "       a.收据费目 = c1.名称(+) And a.收据费目 = c.收据费目(+) And Decode(c.费用场合(+), 0, 2, c.费用场合(+)) = 2 and" & vbNewLine & _
            "       a.收费细目id = m.药品id(+) And m.药名id = q.药名id(+) And q.药名id = j.Id(+) And a.保险大类id = t.Id(+) And" & vbNewLine & _
            "       t.性质(+) = 1 And a.保险大类id = s.保险大类id(+)" & vbNewLine & _
            " Group By a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号), a.收费细目id, c.编码, c.名称, j.编码, j.名称" & vbNewLine & _
            " Order By NO, 序号"
    
    strSQL = "Select Min(a.费用ID) As 费用ID, a.No, a.序号, a.收费细目id, a.计算单位, Avg(a.价格) As 价格, Avg(a.数量) As 数量," & vbNewLine & _
            "       Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.自费金额) As 自费金额," & vbNewLine & _
            "       Max(a.医保项目名称) As 医保项目编码, Max(a.医保项目名称) As 医保项目名称, Max(a.医保报销比例) As 医保报销比例, Max(a.备注) As 备注," & vbNewLine & _
            "       Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, " & vbNewLine & _
            "       Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间, Max(a.收据费目) As 收据费目, Max(a.收据费目编码) As 收据费目编码, Max(a.主页id) As 主页id," & vbNewLine & _
            "       Max(a.类别编码) As 类别编码, Max(a.类别名称) As 类别名称, Max(a.项目编码) As 项目编码, Max(a.项目名称) As 项目名称," & vbNewLine & _
            "       Max(a.规格) As 规格, Max(a.药品剂型) As 药品剂型" & vbNewLine & _
            " From (" & strSQL & ") a" & vbNewLine & _
            " Group By a.No, a.序号, a.收费细目id, a.计算单位" & vbNewLine & _
            IIf(gBs_Type.零费用开票, "", " Having Sum(a.结帐金额) <> 0")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByZyBalance", lng结帐ID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "未找到明细数据，不能打印电子票据"
        Exit Function
    End If
    
    strJsonList = ""
    With rsTmp
        Do While Not .EOF
            If InStr(1, str住院次数 & ",", "," & Nvl(!主页id, 0) & ",") = 0 Then
              str住院次数 = str住院次数 & "," & Nvl(!主页id, 0)
            End If
      
            strData = ""
            strData = strData & "" & GetJsonNodeString("listDetailNo", zlStr.LPAD(Nvl(!费用ID), 20, "0"), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!收据费目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!收据费目), Json_Text)
            strData = strData & "," & GetJsonNodeString("prescribeCode", Nvl(!No), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeCode", Nvl(!类别编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeName", Nvl(!类别名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("code", Nvl(!项目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("name", Nvl(!项目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("form", Nvl(!药品剂型), Json_Text)
            strData = strData & "," & GetJsonNodeString("specification", Nvl(!规格), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!计算单位), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!价格)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!数量)), intJsonFormat_数量小数), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!实收金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!自费金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("receivableAmt", FormatEx(Val(Nvl(!应收金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!医保项目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("medCareItemType", Nvl(!医保项目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("medReimburseRate", FormatEx(Val(Nvl(!医保报销比例)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!备注), Json_Text)
            strData = strData & "," & GetJsonNodeString("sortNo", Nvl(!序号), Json_num)
            strData = strData & "," & GetJsonNodeString("chrgtype", Nvl(!费用类型), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dbl票据总金额 = dbl票据总金额 + RoundEx(Nvl(!实收金额), 6)
            .MoveNext
        Loop
        strListDetail = GetNodeString("listDetail") & ":[" & Mid(strJsonList, 2) & "]"
        str住院次数 = Mid(str住院次数, 2)
    End With

    '分类明细
    If gBs_Type.误差费对照编码 <> "" Then
        dbl票据总金额 = dbl票据总金额 - dbl误差费
    End If
    dbl票据总金额 = RoundEx(dbl票据总金额, 2)
    If Not Get分类明细(lng结帐ID, strData, dbl票据总金额, 3.2, strErrMsg_Out) Then Exit Function
    strChargeDetail = GetNodeString("chargeDetail") & ":[" & strData & "]"
    
    '票据信息
    With rsBalance
        '业务流水号:lng结帐ID_lngEInvoiceID
        strData = ""
        strData = strData & "" & GetJsonNodeString("busNo", lng结帐ID & "_" & lngEInvoiceID, Json_Text)
        strData = strData & "," & GetJsonNodeString("busType", Nvl(rsBalance!业务标识), Json_Text)
        strData = strData & "," & GetJsonNodeString("payer", Nvl(rsBalance!姓名), Json_Text)
        strData = strData & "," & GetJsonNodeString("busDateTime", Format(Nvl(rsBalance!收费时间), "YYYYMMDDhhmmss000"), Json_Text)
        strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
        strData = strData & "," & GetJsonNodeString("payee", Nvl(rsBalance!操作员姓名), Json_Text)
        strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
        strData = strData & "," & GetJsonNodeString("checker", mstrOperatorName, Json_Text)
        strData = strData & "," & GetJsonNodeString("totalAmt", dbl票据总金额, Json_num)
        strData = strData & "," & GetJsonNodeString("remark", IIf(RoundEx(dbl误差费, 6) <> 0 And gBs_Type.误差费对照编码 = "", "存在" & FormatEx(dbl误差费, 6) & "误差金额不参与结算", Nvl(rsBalance!备注)), Json_Text)
        strJson_Out = strData
    End With
    
    '移动支付(一致)
    If Not Get移动支付信息(Val(Nvl(rsBalance!病人ID)), lng结帐ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '通知消息
    With rsBalance
        Call Get通知卡号(Val(Nvl(!病人ID)), Nvl(!身份证号), str卡名称, str卡号)
        strData = ""
        strData = strData & "" & GetJsonNodeString("tel", Nvl(!手机号), Json_Text)
        strData = strData & "," & GetJsonNodeString("email", Nvl(!email), Json_Text)
        If gBs_Type.支持版本 > BS_Version.V2_0_3 Then
            strData = strData & "," & GetJsonNodeString("payerType", Nvl(!缴款类型), Json_Text)
        End If
        strData = strData & "," & GetJsonNodeString("idCardNo", Nvl(!身份证号), Json_Text)
        strData = strData & "," & GetJsonNodeString("cardType", str卡名称, Json_Text)
        strData = strData & "," & GetJsonNodeString("cardNo", str卡号, Json_Text)
        strJson_Out = strJson_Out & "," & strData
    End With
    
    '就诊信息
    With rsBalance
        Call Get医保信息(bytInvoiceType, lng结帐ID, Val(Nvl(!病人ID)), cllInsureInfo)
        strData = ""
        strData = strData & "" & GetJsonNodeString("medicalInstitution", GetUnitInfo("医疗机构类型"), Json_Text)
        strData = strData & "," & GetJsonNodeString("medCareInstitution", zlGetNodeValueFromCollect(cllInsureInfo, "_保险机构编码", "C"), Json_Text)
        strData = strData & "," & GetJsonNodeString("medCareTypeCode", Get医疗付款方式编码(Nvl(!医疗卡付款方式)), Json_Text)
        strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!医疗卡付款方式), Json_Text)
        strData = strData & "," & GetJsonNodeString("medicalInsuranceID", zlGetNodeValueFromCollect(cllInsureInfo, "_医保号", "C"), Json_Text)
        strData = strData & "," & GetJsonNodeString("category", Nvl(!入院科室名称), Json_Text)
        strData = strData & "," & GetJsonNodeString("categoryCode", Nvl(!入院科室编码), Json_Text)
        strData = strData & "," & GetJsonNodeString("leaveCategory", Nvl(!出院科室名称), Json_Text)
        strData = strData & "," & GetJsonNodeString("leaveCategoryCode", Nvl(!出院科室编码), Json_Text)
        strData = strData & "," & GetJsonNodeString("hospitalNo", Nvl(!住院号), Json_Text)
        strData = strData & "," & GetJsonNodeString("visitNo", Nvl(!住院号), Json_Text)
        strData = strData & "," & GetJsonNodeString("consultationDate", Format(Nvl(!入院日期), strJsonFormat_就诊日期), Json_Text)
        strData = strData & "," & GetJsonNodeString("patientId", Nvl(!病人ID), Json_Text)
        strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!主页id), Json_Text)
        strData = strData & "," & GetJsonNodeString("sex", Nvl(!性别), Json_Text)
        strData = strData & "," & GetJsonNodeString("age", Nvl(!年龄), Json_Text)
        strData = strData & "," & GetJsonNodeString("hospitalArea", Nvl(!病区名称), Json_Text)
        strData = strData & "," & GetJsonNodeString("bedNo", Nvl(!床号), Json_Text)
        strData = strData & "," & GetJsonNodeString("caseNumber", Nvl(!病历号), Json_Text)
        strData = strData & "," & GetJsonNodeString("ICD", Nvl(!疾病编码), Json_Text)
        If InStr(1, str住院次数, ",") > 0 Then
            strSQL = "Select Min(入院日期) As 入院日期, Max(出院日期) As 出院日期, Sum(Nvl(出院日期, Sysdate) - 入院日期) As 住院天数" & vbNewLine & _
                    "From 病案主页" & vbNewLine & _
                    "Where 病人id = [1] And" & vbNewLine & _
                    "      主页id In (Select /*+cardinality(A,10)*/" & vbNewLine & _
                    "                Column_Value" & vbNewLine & _
                    "               From Table(f_Num2list([2])) a)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByZyBalance", Val(Nvl(!病人ID)), str住院次数)
            With rsTmp
                If Not .EOF Then
                    strData = strData & "," & GetJsonNodeString("inHospitalDate", Format(Nvl(!入院日期), strJsonFormat_就诊日期), Json_Text)
                    strData = strData & "," & GetJsonNodeString("outHospitalDate", Format(Nvl(!出院日期), strJsonFormat_就诊日期), Json_Text)
                    strData = strData & "," & GetJsonNodeString("hospitalDays", FormatEx(Nvl(!住院天数), 2), Json_num)
                End If
            End With
        Else
            strData = strData & "," & GetJsonNodeString("inHospitalDate", Format(Nvl(!入院日期), strJsonFormat_就诊日期), Json_Text)
            strData = strData & "," & GetJsonNodeString("outHospitalDate", Format(Nvl(!出院日期), strJsonFormat_就诊日期), Json_Text)
            strData = strData & "," & GetJsonNodeString("hospitalDays", FormatEx(Nvl(!住院天数), 2), Json_num)
        End If
        strJson_Out = strJson_Out & "," & strData
    End With
    
    '预交支付
    strSQL = "Select q.凭证代码, q.凭证号码, a.No, Max(a.冲预交) As 冲预交" & vbNewLine & _
            "  From (Select NO, Sum(冲预交) As 冲预交" & vbNewLine & _
            "          From 病人预交记录" & vbNewLine & _
            "          Where 结帐id = [1] And Mod(记录性质, 10) = 1 Group by No) A, 病人预交记录 B, 电子票据使用记录 Q" & vbNewLine & _
            "  Where a.No = b.No And b.记录性质 = 1 And b.Id = q.结算id And q.票种 = 2 And Q.退款ID is null" & vbNewLine & _
            " Group by q.凭证代码, q.凭证号码, a.No"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByZyBalance", lng结帐ID)
    strJsonList = ""
    With rsTmp
        Do While Not .EOF
            strData = ""
            strData = strData & "" & GetJsonNodeString("voucherBatchCode", Nvl(!凭证代码), Json_Text)
            strData = strData & "," & GetJsonNodeString("voucherNo", Nvl(!凭证号码), Json_Text)
            strData = strData & "," & GetJsonNodeString("voucherAmt", FormatEx(Nvl(!凭证号码, 0), 6), Json_num)
            strJsonList = strJsonList & ",{" & strData & "}"
            .MoveNext
        Loop
        strJson_Out = strJson_Out & "," & GetNodeString("payMentVoucher") & ":[" & Mid(strJsonList, 2) & "]"
    End With
    
    '支付信息
    If Not Get结算信息(lng结帐ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '缴费渠道
    If Not Get缴费渠道(lng结帐ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & GetNodeString("payChannelDetail") & ":[" & strData & "]"
    
    '其它医保信息-暂无
    '其它扩展信息-暂无
    'eBillRelateNo  业务票据关联号  String  32  否  如一笔业务数据需要开具N张电子票据，则N张电子票对应该值保持一致，用于后期关联查询
    'isArrears  是否可流通  String  1  是  0-否、1-是（如欠费情况根据医院业务要求该票据是否可流通）
    'arrearsReason  不可流通原因  String  200  否  isArrears=0，填写不可流通的原因
    strData = ""
    strData = strData & "" & GetJsonNodeString("eBillRelateNo", "", Json_Text)
    strData = strData & "," & GetJsonNodeString("isArrears", "1", Json_Text)
    strData = strData & "," & GetJsonNodeString("arrearsReason", "", Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '收费项目明细
    strJson_Out = strJson_Out & "," & strChargeDetail
    '清单项目明细
    strJson_Out = strJson_Out & "," & strListDetail
    
    '返回完整的Json串
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceByZyBalance = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function zlGetJson_CreateEInvoiceBySendCard(ByVal lngEInvoiceID As Long, ByVal lng结帐ID As Long, ByVal lng销账ID As Long, _
                ByVal strEInvoiceClientCode As String, dbl票据总金额 As Double, _
                strJson_Out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取发卡发票Json格式数据
    ' 入参 : lngEInvoiceID -电子票据使用记录.ID
    '        strEInvoiceClientCode-开票点编号
    ' 出参 : strJson-挂号结算信息
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 08:58
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim cllInsureInfo As Collection
    Dim bytInvoiceType As Byte
    Dim dbl误差费 As Double
    Dim lng病人ID As Long, lng医嘱序号 As Long
    Dim str业务标识 As String, strChargeDetail As String, strListDetail As String
    Dim str结帐IDs As String, str登记时间 As String, str业务操作员 As String
    Dim str患者姓名 As String, str患者性别 As String, str患者年龄 As String, str医疗付款方式编码 As String
    Dim strJsonList As String, strData As String
    Dim strJsonKey_就诊科室 As String
    Dim strJsonFormat_就诊日期 As String
    Dim intJsonFormat_费用小数 As Integer, intJsonFormat_数量小数 As Integer
    On Error GoTo ErrHand
    bytInvoiceType = 5
    str业务标识 = zlGet业务标识(bytInvoiceType)
    dbl票据总金额 = 0
    
    '版本差异
    strJsonKey_就诊科室 = GetVersionDiff(1, "就诊科室")
    strJsonFormat_就诊日期 = GetVersionDiff(2, "就诊日期")
    intJsonFormat_费用小数 = Val(GetVersionDiff(2, "费用小数"))
    intJsonFormat_数量小数 = Val(GetVersionDiff(2, "数量小数"))
    
    strSQL = "Select Min(a.Id) As 费用id, a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号) As 序号, a.收费细目id, Max(a.计算单位) As 计算单位," & vbNewLine & _
            "        Sum(a.标准单价) As 价格, Avg(Nvl(a.付数, 1) * Nvl(a.数次, 0)) As 数量, Sum(a.应收金额) As 应收金额," & vbNewLine & _
            "        Sum(a.实收金额) As 实收金额, Sum(a.结帐金额) As 结帐金额, Sum(a.实收金额) - Sum(a.统筹金额) As 自费金额," & vbNewLine & _
            "        Max(s.大类编码) As 医保项目编码, Max(s.大类名称) As 医保项目名称, Max(t.统筹比额) As 医保报销比例, Max(a.摘要) As 备注," & vbNewLine & _
            "        Max(a.费用类型) As 费用类型, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.姓名) As 姓名," & vbNewLine & _
            "        Max(a.性别) As 性别, Max(a.年龄) As 年龄, Max(a.病人id) As 病人id, Max(a.登记时间) As 登记时间, Max('') As 付款方式编码," & vbNewLine & _
            "        Max(Nvl(c.名称, c1.名称)) As 收据费目, Max(Nvl(c.编码, c1.编码)) As 收据费目编码, Max(a.医嘱序号) As 医嘱序号, Max(0) As 挂号id," & vbNewLine & _
            "        Max(d.编码) As 类别编码, Max(d.类别) As 类别名称, Max(b.编码) As 项目编码, Max(b.名称) As 项目名称, Max(b.规格) As 规格," & vbNewLine & _
            "        Max(q.药品剂型) As 药品剂型" & vbNewLine & _
            " From 住院费用记录 A, 收费项目目录 B, 收据费目对照 C, 收据费目 C1, 收费类别 D, 药品规格 M, 药品特性 Q, 诊疗项目目录 J, 保险支付大类 T, 支付类别对照 S" & vbNewLine & _
            " Where a.No In (Select Distinct NO From 住院费用记录 Where 结帐id = [1]) And a.记录性质 = 5 And a.记录状态 = 1 And" & vbNewLine & _
            "       a.收费类别 = d.编码(+) And a.收费细目id = b.Id And a.收据费目 = c1.名称(+) And a.收据费目 = c.收据费目(+) and Decode(c.费用场合(+), 0, 1, c.费用场合(+)) = 1 And a.收费细目id = m.药品id(+) And" & vbNewLine & _
            "       m.药名id = q.药名id(+) And q.药名id = j.Id(+) And a.保险大类id = t.Id(+) And t.性质(+) = 1 And" & vbNewLine & _
            "       a.保险大类id = s.保险大类id(+)" & vbNewLine & _
            " Group By a.No, a.记录状态, a.结帐id, Nvl(a.价格父号, a.序号), a.收费细目id, c.编码, c.名称, j.编码, j.名称" & vbNewLine & _
            IIf(gBs_Type.零费用开票, "", " Having Sum(a.结帐金额) <> 0") & vbNewLine & _
            " Order By NO, 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceBySendCard", lng结帐ID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "未找到明细数据，不能打印电子票据"
        Exit Function
    End If
    
    strJsonList = ""
    With rsTmp
        str患者姓名 = Nvl(!姓名)
        str患者性别 = Nvl(!性别)
        str患者年龄 = Nvl(!年龄)
        lng病人ID = Val(Nvl(!病人ID))
        str医疗付款方式编码 = Nvl(!付款方式编码)
        lng医嘱序号 = Val(Nvl(!医嘱序号))
        str登记时间 = Format(Nvl(!登记时间), "YYYYMMDDhhmmss000")
        str业务操作员 = Nvl(!操作员姓名)
        
        Do While Not .EOF
            strData = ""
            strData = strData & "" & GetJsonNodeString("listDetailNo", zlStr.LPAD(Nvl(!费用ID), 20, "0"), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!收据费目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!收据费目), Json_Text)
            strData = strData & "," & GetJsonNodeString("prescribeCode", Nvl(!No), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeCode", Nvl(!类别编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("listTypeName", Nvl(!类别名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("code", Nvl(!项目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("name", Nvl(!项目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("form", Nvl(!药品剂型), Json_Text)
            strData = strData & "," & GetJsonNodeString("specification", Nvl(!规格), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!计算单位), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!价格)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!数量)), intJsonFormat_数量小数), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!实收金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!自费金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("receivableAmt", FormatEx(Val(Nvl(!应收金额)), intJsonFormat_费用小数), Json_num)
            strData = strData & "," & GetJsonNodeString("medicalCareType", Nvl(!医保项目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("medCareItemType", Nvl(!医保项目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("medReimburseRate", FormatEx(Val(Nvl(!医保报销比例)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!备注), Json_Text)
            strData = strData & "," & GetJsonNodeString("sortNo", Nvl(!序号), Json_num)
            strData = strData & "," & GetJsonNodeString("chrgtype", Nvl(!费用类型), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dbl票据总金额 = dbl票据总金额 + RoundEx(Nvl(!实收金额), 6)
            .MoveNext
        Loop
        
        str结帐IDs = GetBalanceIDs(lng结帐ID, bytInvoiceType)
        dbl误差费 = GetBalanceErrorFee(str结帐IDs)
        strListDetail = GetNodeString("listDetail") & ":[" & Mid(strJsonList, 2) & "]"
    End With
    
    '分类明细
    If gBs_Type.误差费对照编码 <> "" Then
        dbl票据总金额 = dbl票据总金额 - dbl误差费
    End If
    dbl票据总金额 = RoundEx(dbl票据总金额, 2)
    If Not Get分类明细(str结帐IDs, strData, dbl票据总金额, bytInvoiceType, strErrMsg_Out) Then Exit Function
    strChargeDetail = GetNodeString("chargeDetail") & ":[" & strData & "]"
    
    '票据信息
    '业务流水号:lngEInvoiceID_lng结帐ID
    strData = ""
    strData = strData & "" & GetJsonNodeString("busNo", lng结帐ID & "_" & lngEInvoiceID, Json_Text)
    strData = strData & "," & GetJsonNodeString("busType", str业务标识, Json_Text)
    strData = strData & "," & GetJsonNodeString("payer", str患者姓名, Json_Text)
    strData = strData & "," & GetJsonNodeString("busDateTime", str登记时间, Json_Text)
    strData = strData & "," & GetJsonNodeString("placeCode", strEInvoiceClientCode, Json_Text)
    strData = strData & "," & GetJsonNodeString("payee", str业务操作员, Json_Text)
    strData = strData & "," & GetJsonNodeString("author", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("checker", mstrOperatorName, Json_Text)
    strData = strData & "," & GetJsonNodeString("totalAmt", dbl票据总金额, Json_num)
    strData = strData & "," & GetJsonNodeString("remark", IIf(RoundEx(dbl误差费, 6) <> 0 And gBs_Type.误差费对照编码 = "", "存在" & FormatEx(dbl误差费, 6) & "误差金额不参与结算", ""), Json_Text)
    strJson_Out = strData
    
    
    '移动支付
    If Not Get移动支付信息(lng病人ID, lng结帐ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '通知消息
    If Not Get通知消息(lng病人ID, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '就诊信息
    Call Get医保信息(bytInvoiceType, lng结帐ID, lng病人ID, cllInsureInfo)
    Set rsTmp = Nothing
    If lng医嘱序号 <> 0 Then
        strSQL = "Select Max(To_Char(a.发生时间, 'yyyy-mm-dd')) As 就诊日期, Max(b.编码) As 就诊科室编码," & vbNewLine & _
                "       Max(b.名称) As 就诊科室名称, Max(a.No) As 就诊编号" & vbNewLine & _
                "  From 病人挂号记录 A, 部门表 B" & vbNewLine & _
                "  Where a.执行部门id = b.Id And " & vbNewLine & _
                "   a.No = (Select Max(挂号单) From 病人医嘱记录 Where ID = [1] Or 相关id = [1])"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng医嘱序号)
        If rsTmp.RecordCount = 0 Then Set rsTmp = Nothing
    End If
    If rsTmp Is Nothing Then
        strSQL = "Select To_Char(a.发生时间, 'yyyy-mm-dd') As 就诊日期, b.编码 As 就诊科室编码," & vbNewLine & _
                "       b.名称 As 就诊科室名称, a.No As 就诊编号" & vbNewLine & _
                "  From 病人挂号记录 A, 部门表 B" & vbNewLine & _
                "  Where a.执行部门id = b.Id And " & vbNewLine & _
                "       a.Id = (Select ID" & vbNewLine & _
                "           From (Select ID, 发生时间 From 病人挂号记录 Where 病人id = [1] Order By 发生时间 Desc)" & vbNewLine & _
                "           Where Rownum < 2)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetJson_CreateEInvoiceByCharge", lng病人ID)
    End If
    
    strData = ""
    strData = strData & "" & GetJsonNodeString("medicalInstitution", GetUnitInfo("医疗机构类型"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareInstitution", zlGetNodeValueFromCollect(cllInsureInfo, "_保险机构编码", "C"), Json_Text)
    strData = strData & "," & GetJsonNodeString("medCareTypeCode", str医疗付款方式编码, Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalCareType", Get医疗付款方式名称(str医疗付款方式编码), Json_Text)
    strData = strData & "," & GetJsonNodeString("medicalInsuranceID", zlGetNodeValueFromCollect(cllInsureInfo, "_医保号", "C"), Json_Text)
    With rsTmp
        If .RecordCount > 0 Then
            strData = strData & "," & GetJsonNodeString("consultationDate", Format(Nvl(!就诊日期), strJsonFormat_就诊日期), Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_就诊科室, Nvl(!就诊科室名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("patientCategoryCode", Nvl(!就诊科室编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", Nvl(!就诊编号), Json_Text)
        Else
            strData = strData & "," & GetJsonNodeString("consultationDate", "", Json_Text)
            strData = strData & "," & GetJsonNodeString(strJsonKey_就诊科室, "", Json_Text)
            strData = strData & "," & GetJsonNodeString("patientCategoryCode", "", Json_Text)
            strData = strData & "," & GetJsonNodeString("patientNo", lng结帐ID, Json_Text)
        End If
    End With
    strData = strData & "," & GetJsonNodeString("patientId", lng病人ID, Json_Text)
    strData = strData & "," & GetJsonNodeString("sex", str患者性别, Json_Text)
    strData = strData & "," & GetJsonNodeString("age", str患者年龄, Json_Text)
    
    strJson_Out = strJson_Out & "," & strData
    
    '支付信息
    If Not Get结算信息(str结帐IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & strData
    
    '缴费渠道
    If Not Get缴费渠道(str结帐IDs, strData) Then Exit Function
    strJson_Out = strJson_Out & "," & GetNodeString("payChannelDetail") & ":[" & strData & "]"
    
    '其它医保信息-暂无
    '其它扩展信息-暂无
    'eBillRelateNo  业务票据关联号  String  32  否  如一笔业务数据需要开具N张电子票据，则N张电子票对应该值保持一致，用于后期关联查询
    'isArrears  是否可流通  String  1  是  0-否、1-是（如欠费情况根据医院业务要求该票据是否可流通）
    'arrearsReason  不可流通原因  String  200  否  isArrears=0，填写不可流通的原因
    strData = ""
    strData = strData & "" & GetJsonNodeString("eBillRelateNo", "", Json_Text)
    strJson_Out = strJson_Out & "," & strData
    
    '收费项目明细
    strJson_Out = strJson_Out & "," & strChargeDetail
    '清单项目明细
    strJson_Out = strJson_Out & "," & strListDetail
    
    '返回完整的Json串
    strJson_Out = "{" & strJson_Out & "}"
    zlGetJson_CreateEInvoiceBySendCard = True
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function Get分类明细(ByVal str结帐IDs As String, strChargeDetail As String, _
                ByVal dbl票据总金额 As Double, _
                Optional ByVal sng业务类型 As Single, Optional strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取分类明细
    ' 入参 : sng业务类型:1-收费;2-预交;3.1-门诊结帐;3.2-住院结帐,4-挂号;5-发卡
    ' 出参 :
    ' 返回 : chargeDetail节点内容
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lng序号 As Long
    Dim dbl合计金额 As Double, dbl误差费 As Double
    Dim strJsonList As String, strData As String, strTable As String
    On Error GoTo ErrHand
    
    dbl误差费 = 0
    '博思不支持收费项目金额为0，也不支持收费项目为空时开具电子票据
    strTable = IIf(sng业务类型 = 3.2 Or sng业务类型 = 5, "住院费用记录", "门诊费用记录")
    
    strSQL = "Select Rownum As 序号, 收据费目编码, 收据费目名称, 数量, 计算单位, Round(单价, 2) As 单价, Round(结帐金额, 2) As 结帐金额," & vbNewLine & _
            "        Round(自费金额, 2) As 自费金额, 备注, 结帐金额 - Round(结帐金额, 2) As 误差费" & vbNewLine & _
            "  From (Select /*+cardinality(b,10)*/" & vbNewLine & _
            "         Nvl(c.编码, c1.编码) As 收据费目编码, Nvl(c.名称, c1.名称) As 收据费目名称, 1 As 数量, '' As 计算单位, Sum(a.结帐金额) As 单价, a.收据费目," & vbNewLine & _
            "         Sum(a.结帐金额) As 结帐金额, Sum(a.结帐金额) - Sum(a.统筹金额) As 自费金额, '' As 备注" & vbNewLine & _
            "        From " & strTable & " A, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B, 收据费目对照 C, 收据费目 C1" & vbNewLine & _
            "        Where a.结帐id = b.Column_Value And a.收据费目 = c1.名称(+) And a.收据费目 = c.收据费目(+) and Decode(c.费用场合(+), 0, [2], c.费用场合(+)) = [2]" & vbNewLine & _
            "        Group By c.编码, c1.编码, c.名称, c1.名称, a.收据费目" & _
            IIf(gBs_Type.零费用开票, "", " Having Sum(a.结帐金额) <> 0") & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get分类明细", str结帐IDs, IIf(sng业务类型 = 3.2, 2, 1))
    If rsTmp.EOF Then Exit Function
    
    strJsonList = ""
    With rsTmp
        Do While Not .EOF
            strData = ""
            lng序号 = Val(Nvl(!序号, 1))
            strData = strData & "" & GetJsonNodeString("sortNo", Nvl(!序号, 1), Json_num)
            strData = strData & "," & GetJsonNodeString("chargeCode", Nvl(!收据费目编码), Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", Nvl(!收据费目名称), Json_Text)
            strData = strData & "," & GetJsonNodeString("unit", Nvl(!计算单位), Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(Val(Nvl(!单价)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("number", FormatEx(Val(Nvl(!数量)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(Val(Nvl(!结帐金额)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(Val(Nvl(!自费金额)), 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", Nvl(!备注), Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
            dbl合计金额 = dbl合计金额 + RoundEx(Nvl(!结帐金额), 2)
            .MoveNext
        Loop
        
        dbl误差费 = dbl票据总金额 - dbl合计金额
        If RoundEx(dbl误差费, 6) <> 0 And gBs_Type.误差费对照编码 <> "" Then
            strData = ""
            strData = strData & "" & GetJsonNodeString("sortNo", lng序号 + 1, Json_num)
            strData = strData & "," & GetJsonNodeString("chargeCode", gBs_Type.误差费对照编码, Json_Text)
            strData = strData & "," & GetJsonNodeString("chargeName", gBs_Type.误差费对照名称, Json_Text)
            strData = strData & "," & GetJsonNodeString("std", FormatEx(dbl误差费, 2), Json_num)
            strData = strData & "," & GetJsonNodeString("number", 1, Json_num)
            strData = strData & "," & GetJsonNodeString("amt", FormatEx(dbl误差费, 2), Json_num)
            strData = strData & "," & GetJsonNodeString("selfAmt", FormatEx(dbl误差费, 2), Json_num)
            strData = strData & "," & GetJsonNodeString("remark", "", Json_Text)
            strJsonList = strJsonList & ",{" & strData & "}"
        End If
        
        strChargeDetail = Mid(strJsonList, 2)
    End With
    
    Get分类明细 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get通知消息(ByVal lng病人ID As Long, strNotice As String, Optional str门诊号 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取通知消息
    ' 入参 :
    ' 出参 :
    ' 返回 : Collect：成员(险类,医保号,保险机构编码,病种名称)
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    
    If lng病人ID <> 0 Then
        strSQL = "Select Max(a.病人id) As 病人id, Max(a.姓名) As 姓名, Max(a.手机号) As 手机号, Max(a.Email) As Email, Max(1) As 缴款类型," & vbNewLine & _
                "      Max(a.身份证号) As 身份证号, Max(m.卡号) As 卡号, Max(a.门诊号) As 门诊号" & vbNewLine & _
                "From 病人信息 A," & vbNewLine & _
                "    (" & vbNewLine & _
                "      Select 病人id, 卡号" & vbNewLine & _
                "      From (Select b.病人id, b.卡号" & vbNewLine & _
                "              From 病人医疗卡信息 B" & vbNewLine & _
                "              Where b.卡类别id = [2] And b.病人id = [1])" & vbNewLine & _
                "      Where Rownum < 2) M" & vbNewLine & _
                "Where a.病人id = m.病人id(+) And a.病人id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get通知消息", lng病人ID, gBs_Type.缺省卡类别ID)
        
        strNotice = ""
        If rsTmp.RecordCount > 0 Then
            With rsTmp
                strNotice = strNotice & "" & GetJsonNodeString("tel", Nvl(!手机号), Json_Text)
                strNotice = strNotice & "," & GetJsonNodeString("email", Nvl(!email), Json_Text)
                If gBs_Type.支持版本 > BS_Version.V2_0_3 Then
                    strNotice = strNotice & "," & GetJsonNodeString("payerType", Nvl(!缴款类型), Json_Text)
                End If
                strNotice = strNotice & "," & GetJsonNodeString("idCardNo", Nvl(!身份证号), Json_Text)
                If Nvl(!卡号) <> "" Then
                    strNotice = strNotice & "," & GetJsonNodeString("cardType", gBs_Type.医疗卡类型编号, Json_Text)
                    strNotice = strNotice & "," & GetJsonNodeString("cardNo", Nvl(!卡号), Json_Text)
                ElseIf Nvl(!身份证号) <> "" And gBs_Type.身份证作卡类型编号 <> "" Then
                    strNotice = strNotice & "," & GetJsonNodeString("cardType", gBs_Type.身份证作卡类型编号, Json_Text)
                    strNotice = strNotice & "," & GetJsonNodeString("cardNo", Nvl(!身份证号), Json_Text)
                Else
                    strNotice = strNotice & "," & GetJsonNodeString("cardType", gBs_Type.病人无卡的卡类别编号, Json_Text)
                    strNotice = strNotice & "," & GetJsonNodeString("cardNo", gBs_Type.病人无卡的卡号, Json_Text)
                End If
                str门诊号 = Nvl(!门诊号)
            End With
        End If
    End If
    Get通知消息 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get通知卡号(ByVal lng病人ID As Long, ByVal str身份证号 As String, str卡名称 As String, str卡号 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取通知卡号
    ' 入参 :
    ' 出参 :
    ' 返回 : strM_Payment：移动支付信息
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    
    If lng病人ID <> 0 Then
        strSQL = "Select 卡号" & vbNewLine & _
                "From (Select b.病人id, b.卡号" & vbNewLine & _
                "       From 病人医疗卡信息 B " & vbNewLine & _
                "       Where b.卡类别id = [2] And b.病人id = [1])" & vbNewLine & _
                "Where Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get通知卡号", lng病人ID, gBs_Type.缺省卡类别ID)
        If rsTmp.RecordCount > 0 Then
            str卡名称 = gBs_Type.医疗卡类型编号
            str卡号 = Nvl(rsTmp!卡号)
        End If
    End If
    If str卡名称 = "" Then
        If str身份证号 <> "" And gBs_Type.身份证作卡类型编号 <> "" Then
            str卡名称 = gBs_Type.身份证作卡类型编号
            str卡号 = str身份证号
        Else
            str卡名称 = gBs_Type.病人无卡的卡类别编号
            str卡号 = gBs_Type.病人无卡的卡号
        End If
    End If
    Get通知卡号 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get移动支付信息(ByVal lng病人ID As Long, ByVal str结帐IDs As String, strM_Payment As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取移动支付信息
    ' 入参 :
    ' 出参 :
    ' 返回 : strM_Payment：移动支付信息
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    
    strSQL = "Select Max(Decode(信息名, '订单号', 信息值, '支付订单号', 信息值, '')) As 支付定单号," & vbNewLine & _
            "        Max(Decode(信息名, '医保支付订单号', 信息值, '医保订单号', 信息值, '')) As 医保支付定单号," & vbNewLine & _
            "        Max(Decode(Upper(信息名), '支付宝公众号USERID', 信息值, '')) As 支付宝公众号userid," & vbNewLine & _
            "        Max(Decode(Upper(信息名), '支付宝小程序USERID', 信息值, '')) As 支付宝小程序userid," & vbNewLine & _
            "        Max(Decode(Upper(信息名), '微信公众号OPENID', 信息值, '')) As 微信公众号openid," & vbNewLine & _
            "        Max(Decode(Upper(信息名), '微信小程序OPENID', 信息值, '')) As 微信小程序openid" & vbNewLine & _
            " From (Select 信息名, 信息值" & vbNewLine & _
            "        From 病人信息从表" & vbNewLine & _
            "        Where 病人id = [1] And 信息名 In ('支付宝公众号USERID', '支付宝小程序USERID', '微信公众号OPENID', '微信小程序OPENID')" & vbNewLine & _
            "        Union All" & vbNewLine & _
            "        Select 交易项目, 交易内容" & vbNewLine & _
            "        From 三方结算交易" & vbNewLine & _
            "        Where 交易id In (Select ID From 病人预交记录 a, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B Where a.结帐id = b.Column_Value) And 交易项目 Like '%订单号')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get移动支付信息", lng病人ID, str结帐IDs)
    
    strM_Payment = ""
    If rsTmp.RecordCount > 0 Then
        With rsTmp
            strM_Payment = strM_Payment & "" & GetJsonNodeString("alipayCode", Nvl(!支付宝公众号userid), Json_Text)
            strM_Payment = strM_Payment & "," & GetJsonNodeString("weChatOrderNo", Nvl(!支付定单号), Json_Text)
            If gBs_Type.支持版本 > BS_Version.V2_0_3 Then
                strM_Payment = strM_Payment & "," & GetJsonNodeString("weChatMedTransNo", Nvl(!医保支付定单号), Json_Text)
            End If
            If Nvl(!微信公众号openid) <> "" Then
                strM_Payment = strM_Payment & "," & GetJsonNodeString("openID", Nvl(!微信公众号openid), Json_Text)
            Else
                strM_Payment = strM_Payment & "," & GetJsonNodeString("openID", Nvl(!微信小程序openid), Json_Text)
            End If
        End With
    End If
    Get移动支付信息 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get医保信息(ByVal byt场合 As Byte, ByVal lng结算ID As Long, ByVal lng病人ID As Long, _
                cllInsureInfo_Out As Collection, Optional bln住院结帐 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取医保信息
    ' 入参 :
    ' 出参 :
    ' 返回 : Collect：成员(险类,医保号,保险机构编码,病种名称)
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    Set cllInsureInfo_Out = New Collection
    
    strSQL = "Select Max(a.险类) As 险类, Max(b.保险机构编码) As 保险机构编码, Max(Nvl(a.病种名称, c.名称)) As 病种名称" & vbNewLine & _
            "  From 保险结算记录 A, 保险类别 B, 保险病种 C" & vbNewLine & _
            "  Where a.险类 = b.序号 And a.病种id = c.Id(+) And a.记录id = [2] And a.性质 = Decode([1], 2, 3, 3, 2, 1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get医保信息", byt场合, lng结算ID)
    With rsTmp
        If .RecordCount > 0 Then
            cllInsureInfo_Out.Add Nvl(!险类), "_险类"
            cllInsureInfo_Out.Add Nvl(!保险机构编码), "_保险机构编码"
            cllInsureInfo_Out.Add Nvl(!病种名称), "_病种名称"
        End If
    End With
    
    If cllInsureInfo_Out.Count > 0 Then
        If Val(cllInsureInfo_Out("_险类")) <> 0 Then
            strSQL = "Select Max(医保号) As 医保号 From 保险帐户 Where 病人id = [1] And 险类 = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get医保信息", lng病人ID, Val(cllInsureInfo_Out("_险类")))
            With rsTmp
                If .RecordCount > 0 Then
                    cllInsureInfo_Out.Add Nvl(!医保号), "_医保号"
                End If
            End With
            
            If cllInsureInfo_Out("_病种名称") = "" And Not bln住院结帐 Then
                strSQL = "Select Max(病种名称) As 病种名称" & vbNewLine & _
                        "      From (Select Distinct a.名称 As 病种名称" & vbNewLine & _
                        "             From 保险病种 A, 保险特准项目 B" & vbNewLine & _
                        "             Where a.险类 = [1] And a.Id = b.病种id And" & vbNewLine & _
                        "                   b.收费细目id In (Select Distinct 收费细目id From 门诊费用记录 Where 结帐id = [2])" & vbNewLine & _
                        "             Union All" & vbNewLine & _
                        "             Select Distinct a.名称 As 病种名称" & vbNewLine & _
                        "             From 保险病种 A, 保险特准项目 B" & vbNewLine & _
                        "             Where a.险类 = [1] And a.Id = b.病种id And" & vbNewLine & _
                        "                   b.大类 In (Select Distinct 保险大类id From 门诊费用记录 Where 结帐id = [2]))" & vbNewLine & _
                        "      Where Rownum < 2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get医保信息", Val(cllInsureInfo_Out("_险类")), lng结算ID)
                If rsTmp.RecordCount > 0 Then
                    cllInsureInfo_Out.Remove "_病种名称"
                    cllInsureInfo_Out.Add Nvl(rsTmp!病种名称), "_病种名称"
                End If
            End If
        End If
    End If
    Get医保信息 = True
    Exit Function
ErrHand:
    Set cllInsureInfo_Out = New Collection
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get结算信息(ByVal str结帐IDs As String, strPayment As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取结算消息
    ' 入参 :
    ' 出参 :
    ' 返回 : strM_Payment：移动支付信息
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:19
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    strPayment = ""
    
    '如果不支持误差费处理，误差费累计到个人现金支付
    strSQL = "Select 现金预交, 支票预交, 转账预交, 个人帐户支付, 医保统筹基金支付, 其它医保支付, 个人现金支付, Decode(Sign(现金支付), -1, 现金支付, 0) As 现金退款," & vbNewLine & _
            "      Decode(Sign(支票支付), -1, 支票支付, 0) As 支票退款, Decode(Sign(转帐支付), -1, 转帐支付, 0) As 转帐退款," & vbNewLine & _
            "      Decode(Sign(现金支付), -1, 0, 现金支付) As 现金支付, Decode(Sign(支票支付), -1, 0, 支票支付) As 支票支付," & vbNewLine & _
            "      Decode(Sign(转帐支付), -1, 0, 转帐支付) As 转帐支付," & vbNewLine & _
            "      Nvl(个人帐户支付, 0) + Nvl(医保统筹基金支付, 0) + Nvl(其它医保支付, 0) As 报销总额," & vbNewLine & _
            "      Nvl(结算总额, 0) - Nvl(个人帐户支付, 0) - Nvl(医保统筹基金支付, 0) - Nvl(其它医保支付, 0) As 自费金额, 结算总额, 医保结算号码," & vbNewLine & _
            "      0 As 个人帐户余额" & vbNewLine & _
            "From (Select /*+cardinality(b,10)*/" & vbNewLine & _
            "       Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '现金', 1, 0), 0) * a.冲预交) As 现金预交," & vbNewLine & _
            "       Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 1, 0), 0) * a.冲预交) As 支票预交," & vbNewLine & _
            "       Sum(Decode(Mod(a.记录性质, 10), 1, Decode(a.结算方式, '支票', 0, '现金', 0, 1), 0) * a.冲预交) As 转账预交," & vbNewLine & _
            "       Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '个人账户支付', 1, 0)) * a.冲预交) As 个人帐户支付," & vbNewLine & _
            "       Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '医保统筹基金支付', 1, 0)) * a.冲预交) As 医保统筹基金支付," & vbNewLine & _
            "       Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 1, 0)) * a.冲预交) As 其它医保支付," & vbNewLine & _
            "       Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, '其它医保支付', 0, '个人账户支付', 0, '医保统筹基金支付', 0, 1)) *" & vbNewLine & _
            IIf(gBs_Type.误差费对照编码 = "", "", " Decode(D.性质, 9, 0, 1) * ") & " a.冲预交) As 个人现金支付," & vbNewLine & _
            "       Max(Decode(Mod(a.记录性质, 10), 1, 0," & vbNewLine & _
            "                   Decode(c.开票结算方式, '其它医保支付', 结算号码, '个人账户支付', 结算号码, '医保统筹基金支付', 结算号码, ''))) As 医保结算号码," & vbNewLine & _
            "       Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, Null, Decode(a.结算方式, '现金', 1, 0), 0)) * a.冲预交) As 现金支付," & vbNewLine & _
            "       Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, Null, Decode(a.结算方式, '支票', 1, 0), 0)) * a.冲预交) As 支票支付," & vbNewLine & _
            "       Sum(Decode(Mod(a.记录性质, 10), 1, 0, Decode(c.开票结算方式, Null, Decode(a.结算方式, '现金', 0, '支票', 0, 1), 0)) * Decode(D.性质, 9, 0, 1) * a.冲预交) As 转帐支付," & vbNewLine & _
            "       Sum(冲预交) As 结算总额" & vbNewLine & _
            "      From 病人预交记录 A, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B, 开票结算对照 C, 结算方式 D" & vbNewLine & _
            "      Where a.结帐id = b.Column_Value And a.结算方式 = c.结算方式(+) and a.结算方式 = d.名称(+))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get结算信息", str结帐IDs)
    
    With rsTmp
        If .RecordCount > 0 Then
            strPayment = strPayment & "" & GetJsonNodeString("accountPay", FormatEx(Val(Nvl(!个人帐户支付)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("fundPay", FormatEx(Val(Nvl(!医保统筹基金支付)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("otherfundPay", FormatEx(Val(Nvl(!其它医保支付)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("ownPay", FormatEx(Val(Nvl(!自费金额)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("selfConceitedAmt", 0, Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("selfPayAmt", 0, Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("selfCashPay", FormatEx(Val(Nvl(!个人现金支付)), 6), Json_num)
            If gBs_Type.支持版本 > V3_1_0 Then
                strPayment = strPayment & "," & GetJsonNodeString("cashPay", FormatEx(Val(Nvl(!现金预交)) + Val(Nvl(!支票预交)) + Val(Nvl(!转账预交)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("cashRecharge", FormatEx(Val(Nvl(!现金支付)) + Val(Nvl(!支票支付)) + Val(Nvl(!转帐支付)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("cashRefund", FormatEx(Val(Nvl(!现金退款)) + Val(Nvl(!支票退款)) + Val(Nvl(!转帐退款)), 6), Json_num)
            Else
                strPayment = strPayment & "," & GetJsonNodeString("cashPay", FormatEx(Val(Nvl(!现金预交)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("chequePay", FormatEx(Val(Nvl(!支票预交)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("transferAccountPay", FormatEx(Val(Nvl(!转账预交)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("cashRecharge", FormatEx(Val(Nvl(!现金支付)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("chequeRecharge", FormatEx(Val(Nvl(!支票支付)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("transferRecharge", FormatEx(Val(Nvl(!转帐支付)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("cashRefund", FormatEx(Val(Nvl(!现金退款)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("chequeRefund", FormatEx(Val(Nvl(!支票退款)), 6), Json_num)
                strPayment = strPayment & "," & GetJsonNodeString("transferRefund", FormatEx(Val(Nvl(!转帐退款)), 6), Json_num)
            End If
            strPayment = strPayment & "," & GetJsonNodeString("ownAcBalance", FormatEx(Val(Nvl(!个人帐户余额)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("reimbursementAmt", FormatEx(Val(Nvl(!报销总额)), 6), Json_num)
            strPayment = strPayment & "," & GetJsonNodeString("balancedNumber", Nvl(!医保结算号码), Json_Text)
        End If
    End With
    Get结算信息 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get缴费渠道(ByVal str结帐IDs As String, strPayChannelInfo As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 缴费渠道信息
    ' 入参 :
    ' 出参 :
    ' 返回 : strPayChannelInfo：缴费渠道信息
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:19
    ' 说明 : 应用场合：收费、挂号
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strJsonList As String, strData As String
    On Error GoTo ErrHand
    strPayChannelInfo = ""
    
    strSQL = "Select /*+cardinality(b,10)*/" & vbNewLine & _
            "      Nvl(c.渠道编码, Nvl(d.渠道编码, '-')) As 渠道编码, Sum(冲预交) As 结算总额" & vbNewLine & _
            "     From 病人预交记录 A, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B, 收费渠道对照 C," & vbNewLine & _
            "          (Select 结算方式, 渠道编码 From 收费渠道对照 D Where 卡类别id Is Null) D" & vbNewLine & _
            "     Where a.结帐id = b.Column_Value And a.卡类别id = c.卡类别id(+) And a.结算方式 = c.结算方式(+) And a.结算方式 = d.结算方式(+)" & vbNewLine & _
            "     Group By Nvl(c.渠道编码, Nvl(d.渠道编码, '-'))" & vbNewLine & _
            "     Order By 渠道编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get缴费渠道", str结帐IDs)
    
    With rsTmp
        Do While Not .EOF
            strData = ""
            If Nvl(!渠道编码) <> "-" Then
                strData = strData & "" & GetJsonNodeString("payChannelCode", Nvl(!渠道编码), Json_Text)
                strData = strData & "," & GetJsonNodeString("payChannelValue", FormatEx(Nvl(!结算总额), 6), Json_num)
                strJsonList = strJsonList & ",{" & strData & "}"
            End If
            .MoveNext
        Loop
        strPayChannelInfo = Mid(strJsonList, 2)
    End With
    
    Get缴费渠道 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get医疗付款方式名称(ByVal str医疗付款方式编码 As String) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据医疗付款方式编码获取名称
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    If str医疗付款方式编码 = "" Then Exit Function
    strSQL = "Select Max(名称) as 名称 From 医疗付款方式 Where 编码 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get医疗付款方式名称", str医疗付款方式编码)
    If rsTmp.RecordCount > 0 Then
        Get医疗付款方式名称 = Nvl(rsTmp!名称)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get医疗付款方式编码(ByVal str医疗付款方式 As String) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据医疗付款方式编码获取名称
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    If str医疗付款方式 = "" Then Exit Function
    strSQL = "Select Max(编码) as 编码 From 医疗付款方式 Where 名称 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get医疗付款方式编码", str医疗付款方式)
    If rsTmp.RecordCount > 0 Then
        Get医疗付款方式编码 = Nvl(rsTmp!编码)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get预交单据总额(ByVal strNO As String) As Double
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人预交余额
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    strSQL = "Select Sum(金额) As 票据总金额" & vbNewLine & _
            "  From (Select Sum(金额) As 金额" & vbNewLine & _
            "         From 病人预交记录" & vbNewLine & _
            "         Where NO = [1] And 记录性质 = 1" & vbNewLine & _
            "         Union All" & vbNewLine & _
            "         Select Sum(冲预交) As 金额" & vbNewLine & _
            "         From 病人预交记录" & vbNewLine & _
            "         Where 结帐id In (Select Distinct 结帐id From 病人预交记录 Where NO = [1] And Mod(记录性质, 10) = 1) And" & vbNewLine & _
            "               Nvl(金额, 0) < 0 And Mod(记录性质, 10) = 1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get预交余额", strNO)
    If rsTmp.RecordCount > 0 Then
        Get预交单据总额 = Val(Nvl(rsTmp!票据总金额))
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get预交余额(ByVal lng病人ID As Long, ByVal int预交类型 As Integer) As Double
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人预交余额
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/4/22 14:38
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    strSQL = "Select Max(预交余额) As 预交余额 From 病人余额 " & _
            " Where 病人id = [1] And 性质 = 1 And 类型 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get预交余额", lng病人ID, int预交类型)
    If rsTmp.RecordCount > 0 Then
        Get预交余额 = Val(Nvl(rsTmp!预交余额))
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetEInvoiceInfo(ByVal lngEInvoiceID As Long, strErrMsg_Out As String) As ADODB.Recordset
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    strSQL = "Select 记录状态, 票种, 代码 As 票据代码, 号码 As 票据号码, 检验码 As 票据校验码, 生成时间 From 电子票据使用记录 Where Id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetEInvoiceInfo", lngEInvoiceID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "未找到电子票据使用记录，请检查。": Exit Function
    End If
    Set GetEInvoiceInfo = rsTmp
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Private Function GetEInvoiceWithPatiInfo(ByVal lngEInvoiceID As Long, strErrMsg_Out As String) As ADODB.Recordset
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    strSQL = "Select a.记录状态, a.票种, a.代码 As 票据代码, a.号码 As 票据号码, a.检验码 As 票据校验码, b.手机号, b.email," & vbNewLine & _
            "        a.是否换开 " & vbNewLine & _
            "From 电子票据使用记录 a, 病人信息 b" & vbNewLine & _
            "Where a.Id =[1] And a.病人id = b.病人id(+)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetEInvoiceInfo", lngEInvoiceID)
    If rsTmp.RecordCount = 0 Then
        strErrMsg_Out = "未找到电子票据使用记录，请检查。": Exit Function
    End If
    Set GetEInvoiceWithPatiInfo = rsTmp
    Exit Function
ErrHand:
    strErrMsg_Out = Err.Description
End Function

Public Function CheckBillExistReplenishData(ByVal lng结算ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是不是补结算费用
    '返回:True-是补结算 False-反之
    '入参:lng结算ID-结帐ID
    '编制:李南春
    '日期:2020-4-30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    strSQL = "Select 1 From 费用补充记录 A where 结算ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查二次结算", lng结算ID)
    If rsTmp.EOF Then
        CheckBillExistReplenishData = False
    Else
        CheckBillExistReplenishData = True
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalanceErrorFee(ByVal str结算ID As String) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算误差费
    '返回:
    '入参:str结算ID-结帐IDs
    '编制:李南春
    '日期:2020-4-30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    strSQL = "Select /*+cardinality(c,10)*/ Sum(a.冲预交) as 冲预交 From 病人预交记录 A, 结算方式 B, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C " & _
            " where a.结帐id = c.Column_Value and a.结算方式 = b.名称(+) and b.性质 = 9"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查二次结算", str结算ID)
    If Not rsTmp.EOF Then
        GetBalanceErrorFee = Val(Nvl(rsTmp!冲预交))
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalanceIDs(ByVal str结算ID As String, Optional ByVal int业务类型 As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据原结帐ID获取结算涉及的所有结帐ID和冲销ID
    '返回:
    '入参:lng结算ID-结帐ID,1-收费;4-挂号;5-发卡
    '编制:李南春
    '日期:2020-4-30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结帐IDs As String, strTable As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    
    strTable = IIf(int业务类型 = 5, "住院费用记录", "门诊费用记录")
    strSQL = "Select Distinct 结帐ID From " & strTable & _
            " Where (no, 记录性质) In (Select No, 记录性质 From " & strTable & " Where 结帐ID = [1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查二次结算", str结算ID)
    With rsTmp
        Do While Not .EOF
            str结帐IDs = str结帐IDs & "," & Nvl(!结帐ID)
            .MoveNext
        Loop
    End With
    GetBalanceIDs = Mid(str结帐IDs, 2)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

    
Public Function GetPaperCode(ByVal bytInvoiceType As Byte) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取纸质票据代码
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/6/28 17:41
    '---------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    GetPaperCode = Decode(bytInvoiceType, 2, gBs_Type.预交纸质票据代码, 4, gBs_Type.挂号纸质票据代码, 3, gBs_Type.结账纸质票据代码, gBs_Type.收费纸质票据代码)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
