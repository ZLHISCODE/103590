VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*********************************************************************************************************************************************
'药品、卫材、病人及临床相关服务
'一、药品
'二、卫材
'三、临床
'四、病人
'    1.医疗卡相关服务
'    1.1. ZlPatiSvr_ChkCardChangeValid-医疗卡变动合法性检查
'    1.2 zlPatiSvr_GetNextID:获取指定表的下一个序列
'    1.3 zl_PatiSvr_DelCardChangeInfo:删除医疗卡变动记录
'    1.4 zl_PatiSvr_ConfirmCardChange:病人医疗卡变动确认（即生成病人医疗卡信息)
'编制:刘兴洪
'日期:2019-11-25 11:11:39
'*********************************************************************************************************************************************
Private mobjServiceCall  As Object
Private mlngModule As Long
Private mlngErrNum As Long, mstrSource As String, mstrErrMsg As String


Private Function GetServiceCall(ByRef objServiceCall_Out As Object, Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共服务对象
    '出参:objServiceCall_Out-返回公共服务对象
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If Not mobjServiceCall Is Nothing Then Set objServiceCall_Out = mobjServiceCall: GetServiceCall = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
    If Err <> 0 Then
        mstrErrMsg = "部件【zlServiceCall】丢失，请与系统管理员联系，恢复该部件！"
        If blnShowErrMsg Then
            MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
            Err = 0: On Error GoTo 0
        Else
            Err.Raise Err.Number, Err.Source, mstrErrMsg: Exit Function
        End If
        
        Err = 0: On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo errHandle
    If mobjServiceCall.InitService(gcnOracle, gstrDBUser, glngSys, mlngModule) = False Then Set mobjServiceCall = Nothing: Exit Function
    Set objServiceCall_Out = mobjServiceCall
    GetServiceCall = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlInitCommon(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '     objStuffService-卫材服务
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    mlngModule = lngModule
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    Call InitVar
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_GetStock(ByVal lng卫材ID As Long, _
    ByVal lng库房ID As Long, Optional ByVal lng批次 As Long, _
    Optional ByRef dbl可用库存_Out As Double, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定药房指定药品库存(以零售单位)
    '入参:
    '     lng库房ID-药房ID
    '     lng批次:lng批次 = -1 Or lng批次 = 0 表示不分批次查找
    '     lngModule-模块号
    '出参:dbl可用库存_Out-返回可用库存
    '返回:获取返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 17:45:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    On Error GoTo errHandle
    
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接药品服务失败，无法获取库存！", vbInformation, gstrSysName
        Exit Function
    End If
     
    dbl可用库存_Out = 0
    If lng批次 < 0 Then lng批次 = 0
    
    'Zl_Stuffsvr_Getstock
    '  --功能：获取指定卫材在指定库房的可用库存数
    '  --入参：Json_In:格式
    '  --  input
    '  --    stuff_id        N   1   卫材ID
    '  --    warehouse_id    C   1   库房ID
    '  --    batch           N       批次：0-不区分批次，>0只查某批次
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code            N   1 应答吗：0-失败；1-成功
    '  --    message         C   1 应答消息：失败时返回具体的错误信息
    '  --    stock           N   1  可用库存
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("stuff_id", lng卫材ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("warehouse_ids", lng库房ID, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("batch", lng批次, Json_num, True)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_StuffSvr_GetStock", strJson, , "", mlngModule) = False Then Exit Function
    dbl可用库存_Out = objServiceCall.GetJsonNodeValue("output.stock")
    zlStuffSvr_GetStock = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_GetStock(ByVal lng药品ID As Long, _
    ByVal lng库房ID As Long, Optional ByVal lng批次 As Long, _
    Optional ByRef dbl可用库存_Out As Double, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定药房指定药品库存(以零售单位)
    '入参:
    '     lng库房ID-药房ID
    '     lng批次:lng批次 = -1 Or lng批次 = 0 表示不分批次查找
    '     lngModule-模块号
    '出参:dbl可用库存_Out-返回可用库存
    '返回:获取返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 17:45:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    On Error GoTo errHandle
    
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接药品服务失败，无法获取库存！", vbInformation, gstrSysName
        Exit Function
    End If
     
    dbl可用库存_Out = 0
    If lng批次 < 0 Then lng批次 = 0
    
    'zl_DrugSvr_Getstock
    '  --功能：获取指定药品在指定库房的可用库存数
    '  --入参：Json_In:格式
    '  --  input
    '  --    drug_id             N   1   药品ID
    '  --    pharmacy_ids        C   1   库房ID
    '  --    batch               N       批次：<=0-不区分批次，>0只查某批次
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N 1 应答吗：0-失败；1-成功
    '  --    message             C 1 应答消息：失败时返回具体的错误信息
    '  --    stock               N  1  可用库存
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("drug_id", lng药品ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_ids", lng库房ID, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("batch", lng批次, Json_num, True)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_DrugSvr_Getstock", strJson, , "", mlngModule) = False Then Exit Function
    dbl可用库存_Out = objServiceCall.GetJsonNodeValue("output.stock")
    zlDrugSvr_GetStock = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_GetMultiStock(ByVal lng卫材ID As Long, ByVal str库房IDs As String, Optional ByRef dbl可用库存_Out As Double, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定药品在指定药房（多个）的可用库存(以零售单位)(以零售单位)
    '入参:
    '     str库房IDs-库房ID:多个用逗号
    '     lngModule-模块号
    '出参:dbl可用库存_Out-返回可用库存
    '返回:获取返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 21:32:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo errHandle
    dbl可用库存_Out = 0
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
        
    'Zl_Stuffsvr_Getstock
    '  --功能：获取指定卫材在指定库房的可用库存数
    '  --入参：Json_In:格式
    '  --  input
    '  --    stuff_id        N   1   卫材ID
    '  --    warehouse_ids   C   1   库房ids,多个用逗号分离
    '  --    batch           N       批次：0-不区分批次，>0只查某批次
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code            N   1 应答吗：0-失败；1-成功
    '  --    message         C   1 应答消息：失败时返回具体的错误信息
    '  --    stock           N   1  可用库存
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("stuff_id", lng卫材ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("warehouse_ids", str库房IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("batch", -1, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_StuffSvr_GetStock", strJson, , "", mlngModule) = False Then Exit Function
    dbl可用库存_Out = objServiceCall.GetJsonNodeValue("output.stock")
    
    zlStuffSvr_GetMultiStock = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_GetMultiStock(ByVal lng药品ID As Long, ByVal str库房IDs As String, Optional ByRef dbl可用库存_Out As Double, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定药品在指定药房（多个）的可用库存(以零售单位)(以零售单位)
    '入参:lng收费细目ID-药品ID或卫材ID
    '     str库房IDs-库房ID:多个用逗号
    '     lngModule-模块号
    '     bln卫材-是否卫生材料
    '出参:dbl可用库存_Out-返回可用库存
    '返回:获取返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 21:32:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo errHandle
    dbl可用库存_Out = 0
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
        
    'zl_DrugSvr_Getstock
    '  --功能：获取指定药品在指定库房的可用库存数
    '  --入参：Json_In:格式
    '  --  input
    '  --    drug_id             N   1   药品ID
    '  --    pharmacy_ids        C   1   库房ID
    '  --    batch               N       批次：<=0-不区分批次，>0只查某批次
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N 1 应答吗：0-失败；1-成功
    '  --    message             C 1 应答消息：失败时返回具体的错误信息
    '  --    stock               N  1  可用库存
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("drug_id", lng药品ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_ids", str库房IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("batch", -1, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_DrugSvr_Getstock", strJson, , "", mlngModule) = False Then Exit Function
    dbl可用库存_Out = objServiceCall.GetJsonNodeValue("output.stock")
    
    zlDrugSvr_GetMultiStock = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_GetStockInfo(ByVal lng药品ID As Long, str药房性质 As String, ByVal dbl换算系数 As Double, _
    ByRef strStockInfor_out As String, Optional ByVal lngModule As Long, _
    Optional byt库存显示方式 As Byte, Optional ByVal str所属部门ID As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品在各个药房，药库的库存信息
    '入参:lng药品ID-药品ID
    '    str药房性质:如: 中药房 , 西药房, 成药房, 中药库, 西药库, 成药库
    '    dbl换算系数:当前的换算系数
    '    byt库存显示方式:0-直接显示库存数;1-显示有无
    '    str所属部门ID:操作员所属科室ID,用于判断显示库存
    '出参:strStockInfor-返回:库房名称:库存,.... 形式
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 22:28:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata  As Collection, cllTemp As Collection
    Dim i As Long
    On Error GoTo errHandle
    strStockInfor_out = ""
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'zl_DrugSvr_Getstockbydept
    '  --功能：根据指定药品及库房性质获取各库房的库存信息
    '  --入参：Json_In:格式
    '  --  input
    '  --    drug_id             N   1   药品ID
    '  --    pharmacy_nature     C   1  库房性质：中药房，西药房，成药房…
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N   1 应答吗：0-失败；1-成功
    '  --    message             C   1 应答消息：失败时返回具体的错误信息
    '  --    item_list
    '  --        pharmacy_id     N   1   药房ID
    '  --        pharmacy_code   C   1   库房编码
    '  --        pharmacy_name   C   1   库房名称
    '  --        stock           N   1   可用数量
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("drug_id", lng药品ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_nature", str药房性质, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_DrugSvr_GetStockByDept", strJson, , "", mlngModule) = False Then Exit Function
     
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then zlDrugSvr_GetStockInfo = True: Exit Function
     
    If dbl换算系数 = 0 Then dbl换算系数 = 1
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        If InStr("," & str所属部门ID & ",", "," & cllTemp("_pharmacy_id")) > 0 Or byt库存显示方式 = 0 Then
            '显示库存数:所属库房,或其他库房显示为库存数时
            strStockInfor_out = strStockInfor_out & "," & cllTemp("_pharmacy_name") & ":" & FormatEx(Val(cllTemp("_stock")) / dbl换算系数, 2)
        Else
            '非操作员库房,则显示为有无
            strStockInfor_out = strStockInfor_out & "," & cllTemp("_pharmacy_name") & ":" & IIf(Val(cllTemp("_stock")) > 0, "有", "无")
        End If
        
        str所属部门ID = str所属部门ID & "," & cllTemp("_pharmacy_id")
    Next
 
    If strStockInfor_out <> "" Then strStockInfor_out = Mid(strStockInfor_out, 2)
    zlDrugSvr_GetStockInfo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_GetStockBatch(ByVal str卫材IDs As String, ByVal str药房IDs As String, ByRef clldata As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品在各个药房，药库的库存信息
    '入参:
    '    str药房ids:药房id,用逗号分隔
    '    dbl换算系数:当前的换算系数
    '出参:strStockInfor-返回:形式:药品ID:药房ID:库存,....
    '返回:获取成功，返回true,否则返回False
    '编制:李南春
    '日期:2019-08-08 22:28:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    Set clldata = New Collection
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    '---------------------------------------------------------------------------
    '--功能：批量获取多个卫材库存及价格信息:在项目选择器中展示库存及价格信息
    '--入参：Json_In:格式
    '--  input
    '--   stuff_ids            C   1   卫材ID，多个用英文的逗号分隔
    '--   warehouse_ids        C   0   库房IDs，库房IDs为NULL时为所有库房
    '--   return_price         N   1   是否返回售价：1-返回价格信息(售价);0-不返回
    '--   type                 N   1   1-返回库存单位信息
    '--出参: Json_Out,格式如下
    '--  output
    '--    code                N 1 应答吗：0-失败；1-成功
    '--    message             C 1 应答消息：失败时返回具体的错误信息
    '--    item_list
    '--    stuff_id            N   1   卫材ID
    '--    warehouse_id        N    1  库房ID
    '--    stock               N   1   可用数量
    '--    unit                C   1   库存单位
    '--    price               N   1   零售价(返回价格时才有此项)
    '---------------------------------------------------------------------------
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("stuff_ids", str卫材IDs)
    strJson = strJson & "," & GetJsonNodeString("warehouse_ids", str药房IDs)
    strJson = strJson & "," & GetJsonNodeString("return_price", 0, Json_num)
    strJson = strJson & "," & GetJsonNodeString("type", 1, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_StuffSvr_GetStockBatch", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list", "stuff_id,warehouse_id")
    If clldata Is Nothing Then Exit Function
    zlStuffSvr_GetStockBatch = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_GetStockBatch(ByVal str药品IDs As String, ByVal str药房IDs As String, ByRef clldata As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品在各个药房，药库的库存信息
    '入参:
    '    str药房ids:药房id,用逗号分隔
    '    dbl换算系数:当前的换算系数
    '出参:strStockInfor-返回:形式:药品ID:药房ID:库存,....
    '返回:获取成功，返回true,否则返回False
    '编制:李南春
    '日期:2019-08-08 22:28:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    Set clldata = New Collection
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    '---------------------------------------------------------------------------
    '--功能：批量获取多个药品库存及价格信息:在项目选择器中展示库存及价格信息
    '--入参：Json_In:格式
    '--  input
    '--   drug_ids       C   1   药品ID，多个用英文的逗号分隔
    '--   pharmacy_ids   C   0   库房ID，多个用英文的逗号分隔;空字符串,查询所有库房
    '--   return_price   N   0   是否返回售价：1-返回价格信息(售价);0-不返回
    '--   return_dept    N   0   按科室返回库存：1-按科室返回库存;0-按药品返回库存;2-返回科室所有药品的库存
    '--   type           N   0   0-直接返回库存数量;>0返回库存单位：1-按门诊单位;2-按住院单位;3-按零售单位
    '--出参: Json_Out,格式如下
    '--  output
    '--    code                 N   1 应答吗：0-失败；1-成功
    '--    message              C   1 应答消息：失败时返回具体的错误信息
    '--    item_list
    '--    drug_id              N   1   药品ID
    '--    pharmacy_id          N   1   库房ID(按科室返回库存才有此项)
    '--    stock                N   1   可用数量
    '--    unit                 C   1   库存单位
    '--    price                N   1   零售价(返回价格时才有此项)
    '---------------------------------------------------------------------------
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("drug_ids", str药品IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_ids", str药房IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("return_price", 0, Json_num)
    strJson = strJson & "," & GetJsonNodeString("return_dept", 1, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_DrugSvr_GetStockBatch", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list", "drug_id,pharmacy_id")
    If clldata Is Nothing Then Exit Function
    zlDrugSvr_GetStockBatch = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function



Public Function zlStuffSvr_CheckValidity(ByVal lng材料ID As Long, ByVal lng库房ID As Long, ByVal dbl数量 As Double, _
    Optional ByVal blnAsk As Boolean = True, Optional lngModule As Long, Optional ByVal lng批次 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查卫生材料的灭菌效期是否过期
    '入参:objPati-病人信息集
    '     blnAsk=表示是否询问是否继续,否则为提醒
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 23:21:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strMinDate As String, strSql As String, strName As String
    Dim bln存在过期 As Boolean, rsTmp As ADODB.Recordset
    Dim objServiceCall As Object
    Dim strJson As String

    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule

    If GetServiceCall(objServiceCall) = False Then Exit Function


    'zl_StuffSvr_CheckExpiryDate
    '  --功能：检查卫生卫材的灭菌效期是否过期
    '  --入参：Json_In:格式
    '  --    input
    '  --        stuff_id    N   1   卫材ID
    '  --        warehouse_id    N   1   库房ID
    '  --        quantity    N   1   数量
    '  --        batch       N   0   批次
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code    N   1   应答码：0-失败；1-成功
    '  --    message C   1   应答消息：失败时返回具体的错误信息
    '  --    exist_expiried  N   1   是否存在过期:0-不存在已过期项目，1-存在已过期项目
    '  --    min_expirydate  C       最小灭菌效期，exist_expiried=1时返回，格式：yyyy-mm-dd
    If lng批次 < 0 Then lng批次 = 0
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("stuff_id", lng材料ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("warehouse_id", lng库房ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("quantity", dbl数量, Json_num)
    strJson = strJson & "," & GetJsonNodeString("batch", lng批次, Json_num, True)
    strJson = "{""input"":{" & strJson & "}}"

    If objServiceCall.CallService("zl_StuffSvr_CheckExpiryDate", strJson, , "", mlngModule) = False Then Exit Function
 

    bln存在过期 = objServiceCall.GetJsonNodeValue("output.exist_expiried") = 1
    
    If bln存在过期 Then
        strMinDate = objServiceCall.GetJsonNodeValue("output.min_expirydate")
        
        strSql = "Select 名称 From 收费项目目录 where ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取材料名称", lng材料ID)
        If rsTmp.EOF = False Then strName = Nvl(rsTmp!名称)

        If blnAsk Then
            If MsgBox("卫生材料""" & strName & """的灭菌效期""" & strMinDate & """已过期,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            MsgBox "提醒：" & vbCrLf & vbCrLf & "卫生材料""" & strName & """的灭菌效期""" & strMinDate & """已过期。", vbInformation, gstrSysName
        End If
    End If
    zlStuffSvr_CheckValidity = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function zlDrugSvr_PatiCheckExecute(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef strNotSendInfor_out As String, _
    Optional ByVal int检查离院带药 As Integer = 0, Optional int婴儿序号 As Integer = -1, Optional int门诊标志 As Integer = 1, _
    Optional ByVal strNOs As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人在药房是否还有未发药的药品
    '入参:lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     lng检查离院带药-离院带药
    '     int门诊标志-1-门诊;2-住院
    '     lngModule -调用模块号
    '出参:
    '    strNotSendInfor_out-返回未发药信息
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String

    Err = 0: On Error GoTo errHandle:
    strNotSendInfor_out = ""
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function

    'Zl_Drugsvr_CheckPatiExecute
    '  --功能：根据病人信息获取未发药品数据
    '  --入参：Json_In:格式
    '  --  input
    '  --     pati_id        N 1 病人ID
    '  --     pati_pageid        N 1 主页ID
    '  --     baby_num       N   婴儿序号:-1表示不区分;0-母亲的;>0具体婴儿费用
    '  --     check_excutenature       N   检查院带药：1-需要进行离院带药检查;0-不需要离院带药检查
    '  --     fee_source       N 1 费用来源:1-门诊;2-住院;4-体检
    '  --     rcp_nos            处方单据号，数组如：["A0001","A0002"]
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                    N   1   应答吗：0-失败；1-成功
    '  --    message                 C   1   应答消息：失败时返回具体的错误信息
    '  --    isexist       N 1 是否存在: 1-存在;0-不存在
    '  --    drug_notsend_infor        C 1 未发药信息,isexist=1时返回
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("baby_num", int婴儿序号, Json_num)
    strJson = strJson & "," & GetJsonNodeString("check_excutenature", int检查离院带药, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_source", int门诊标志, Json_num)
    If strNOs <> "" Then
        strJson = strJson & ",""rcp_nos"":[""" & Replace(strNOs, ",", """,""") & """]"
    End If
    strJson = "{""input"":{" & strJson & "}}"
   
    If objServiceCall.CallService("Zl_Drugsvr_CheckPatiExecute", strJson, , "", mlngModule) = False Then Exit Function
    strNotSendInfor_out = objServiceCall.GetJsonNodeValue("output.drug_notsend_infor")
    
    zlDrugSvr_PatiCheckExecute = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffsvr_PatiCheckExecute(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef strNotSendInfor_out As String, _
    Optional ByVal int检查离院带药 As Integer = 0, Optional int婴儿序号 As Integer = -1, Optional int门诊标志 As Integer = 1, _
    Optional ByVal strNOs As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人在药房是否还有未发药的卫材
    '入参:lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     lng检查离院带药-离院带药
    '     int门诊标志-1-门诊;2-住院
    '     lngModule -调用模块号
    '出参:
    '    strNotSendInfor_out-返回未发料信息
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String

    Err = 0: On Error GoTo errHandle:
    strNotSendInfor_out = ""
    If lngModule <> 0 Then mlngModule = lngModule

    'Zl_Stuffsvr_CheckPatiExecute
    '  --功能：根据病人信息或获取未发料数据
    '  --入参：Json_In:格式
    '  --  input
    '  --     pati_id        N 1 病人ID
    '  --     pati_pageid        N 1 主页ID
    '  --     baby_num       N   婴儿序号:-1表示不区分;0-母亲的;>0具体婴儿费用
    '  --     fee_source       N 1 费用来源:1-门诊;2-住院;4-体检
    '  --     stuff_nos            处方单据号，数组如：["A0001","A0002"]
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                    N   1   应答吗：0-失败；1-成功
    '  --    message                 C   1   应答消息：失败时返回具体的错误信息
    '  --    isexist       N 1 是否存在: 1-存在;0-不存在
    '  --    stuff_notsend_infor        C 1 未发料信息,isexist=1时返回
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("baby_num", int婴儿序号, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_source", int门诊标志, Json_num)
    If strNOs <> "" Then
        strJson = strJson & ",""stuff_nos"":[""" & Replace(strNOs, ",", """,""") & """]"
    End If
    strJson = "{""input"":{" & strJson & "}}"

    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    If objServiceCall.CallService("Zl_Stuffsvr_CheckPatiExecute", strJson, , "", mlngModule) = False Then Exit Function
    strNotSendInfor_out = objServiceCall.GetJsonNodeValue("output.stuff_notsend_infor")
    
    zlStuffsvr_PatiCheckExecute = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_GetBatchNumber(lng收费细目id As Long, ByVal str单据类型 As String, ByRef strBatchNum_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取卫材批号
    ' 入参 : lng收费细目ID-卫材ID
    ' 出参 : str单据类型-单据
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/8/9 18:46
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String

    On Error GoTo errHandle
    If GetServiceCall(objServiceCall) = False Then Exit Function
    strBatchNum_Out = ""
    If lngModule <> 0 Then mlngModule = lngModule
    
    If str单据类型 = "8" Or str单据类型 = "24" Then
        str单据类型 = "1"
    ElseIf str单据类型 = "9" Or str单据类型 = "25" Then
        str单据类型 = "2"
    ElseIf str单据类型 = "10" Or str单据类型 = "26" Then
        str单据类型 = "3"
    End If

    strJson = Chr(34) & "rcpdtl_id" & Chr(34) & ":" & lng收费细目id & "," & Chr(34) & "billtype" & Chr(34) & ":" & str单据类型

    strJson = "{" & Chr(34) & "input" & Chr(34) & ":{" & strJson & "}}"

    If objServiceCall.CallService("zl_StuffSvr_GetBatchNumber", strJson, , "", mlngModule) = False Then Exit Function
     
    strBatchNum_Out = objServiceCall.GetJsonNodeValue("output.batch_number")
    zlStuffSvr_GetBatchNumber = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_GetPrice(lng收费细目id As Long, ByVal lng药房id As Long, _
    ByVal dbl数量 As Double, ByVal lng批次 As Long, ByVal str价格等级 As String, _
    ByRef dbl单价 As Double, ByRef dbl成本价 As Double, ByRef dbl剩余数量 As Double, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取指定药品的零售价、成本价及可用数量
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/8/9 19:01
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String

    On Error GoTo errHandle
    If GetServiceCall(objServiceCall) = False Then Exit Function
     
    If lngModule <> 0 Then mlngModule = lngModule
    dbl单价 = 0: dbl成本价 = 0: dbl剩余数量 = 0
    If lng批次 < 0 Then lng批次 = 0

    'Zl_Stuffsvr_Getprice
    '  --功能：获取指定卫材的售价、成本价
    '  --入参：Json_In:格式
    '  --  input
    '  --    stuff_id          N 1 卫材ID
    '  --    warehouse_id      N 1 药房ID
    '  --    quantity          N 1 数量
    '  --    batch             N   批次：0-不区分批次，>0只查某批次
    '  --    price_grade       C   价格等级名称
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code              N 1   应答吗：0-失败；1-成功
    '  --    message           C 1   应答消息：失败时返回具体的错误信息
    '  --    price             N 1 售价
    '  --    price_cost        N 1 成本价
    '  --    quantity_remain   N 1 剩余数量，等于0表示数量足够，大于0则表示数量不够
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("stuff_id", lng收费细目id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("warehouse_id", lng药房id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("quantity", dbl数量, Json_num)
    strJson = strJson & "," & GetJsonNodeString("batch", lng批次, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("price_grade", str价格等级, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("zl_StuffSvr_GetPrice", strJson, , "", mlngModule) = False Then Exit Function
    
    dbl单价 = objServiceCall.GetJsonNodeValue("output.price")
    dbl成本价 = objServiceCall.GetJsonNodeValue("output.price_cost")
    dbl剩余数量 = objServiceCall.GetJsonNodeValue("output.quantity_remain")
    
    zlStuffSvr_GetPrice = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_GetPrice(lng收费细目id As Long, ByVal lng药房id As Long, _
    ByVal dbl数量 As Double, ByVal lng批次 As Long, ByVal str价格等级 As String, _
    ByRef dbl单价 As Double, ByRef dbl成本价 As Double, ByRef dbl剩余数量 As Double, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取指定药品的零售价、成本价及可用数量
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/8/9 19:01
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String

    On Error GoTo errHandle
    If GetServiceCall(objServiceCall) = False Then Exit Function
     
    If lngModule <> 0 Then mlngModule = lngModule
    dbl单价 = 0: dbl成本价 = 0: dbl剩余数量 = 0
    If lng批次 < 0 Then lng批次 = 0

    'Zl_Drugsvr_Getprice
    '  --功能：获取指定药品的售价、成本价
    '  --入参：Json_In:格式
    '  --  input
    '  --    drug_id          N 1 药品ID
    '  --    pharmacy_id      N 1 药房ID
    '  --    quantity         N 1 数量
    '  --    batch            N   批次：0-不区分批次，>0只查某批次
    '  --    price_grade      C   价格等级名称
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code              N 1   应答吗：0-失败；1-成功
    '  --    message           C 1   应答消息：失败时返回具体的错误信息
    '  --    price             N 1 售价
    '  --    price_cost        N 1 成本价
    '  --    quantity_remain   N 1 剩余数量，等于0表示数量足够，大于0则表示数量不够
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("drug_id", lng收费细目id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_id", lng药房id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("quantity", dbl数量, Json_num)
    strJson = strJson & "," & GetJsonNodeString("batch", lng批次, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("price_grade", str价格等级, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("zl_DrugSvr_GetPrice", strJson, , "", mlngModule) = False Then Exit Function
    
    dbl单价 = objServiceCall.GetJsonNodeValue("output.price")
    dbl成本价 = objServiceCall.GetJsonNodeValue("output.price_cost")
    dbl剩余数量 = objServiceCall.GetJsonNodeValue("output.quantity_remain")
    
    zlDrugSvr_GetPrice = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugsvr_AutoSendDrugFromNo(ByVal strNo As String, ByVal strCurDate As String, ByVal str操作员姓名 As String, ByVal str操作员编号 As String, _
    Optional intBillType As Integer = 2, Optional lngModule As Long, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号进行自动发药
    '入参:strNo-单据号
    '     strCurDate-当前操作时间
    '     str操作员姓名-操作员姓名
    '     str操作员编号-操作员编号
    '     lngModule -调用模块号
    '     intBilltype-单据类型(1-收费，2-记帐,3-记帐表)
    '出参:strErrMsg_Out-发生错误时，返回错误信息
     '返回:自动发料成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '说明:
    '   不加错误捕获，由上级过程捕获(避免事务站用）
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, varValue As Variant

    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule

    strJson = strJson & "" & GetJsonNodeString("billtype", CStr(intBillType), 1)   'N   1   单据类型:1 -收费处方发药  ;2- 记帐单处方发药;3- 记帐表处方发药;
    strJson = strJson & "," & GetJsonNodeString("rcp_nos", strNo) '  C   1   处方NO串：NO1,NO2...
    strJson = strJson & "," & GetJsonNodeString("operator_name", str操作员姓名) '   C   1   操作员姓名
    strJson = strJson & "," & GetJsonNodeString("operator_code", str操作员编号) '   C   1   操作员编号
    strJson = strJson & "," & GetJsonNodeString("rcpdtl_ids", "")     'rcpdtl_ids           处方明细id串,目前传入的费用ID串，用逗号分隔 ：1,2,3,4
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If objServiceCall.CallService("Zl_Drugsvr_Autosenddrug", strJson, , "", lngModule, False) = False Then Exit Function
    
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlDrugsvr_AutoSendDrugFromNo = True
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function


Public Function zlStuffsvr_AutoSendStuffFromNo(ByVal strNo As String, ByVal str费用IDs As String, ByVal strCurDate As String, ByVal str操作员姓名 As String, ByVal str操作员编号 As String, _
    Optional intBillType As Integer = 2, Optional lngModule As Long, Optional lng病人ID As Long, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号进行自动发药
    '入参:strNo-单据号
    '     strCurDate-当前操作时间
    '     str费用Ids-费用Ids
    '     str操作员姓名-操作员姓名
    '     str操作员编号-操作员编号
    '     lngModule -调用模块号
    '     intBilltype-单据类型(1-收费，2-记帐,3-记帐表)
    '     intSendMode-发药方式( 1-处方发药;2-批量发药;3-部门发药)
    '出参:strErrMsg_Out-发生错误时，返回错误信息
     '返回:自动发料成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '说明:
    '   不加错误捕获，由上级过程捕获(避免事务站用）
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, varValue As Variant

    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule

    strJson = strJson & "" & GetJsonNodeString("stuff_nos", strNo) 'C   1   单据号"
    strJson = strJson & "," & GetJsonNodeString("billtype", intBillType, 1)  '   N       单据类型: 单据类型: 1-收费处方；2-记帐单处方；3-记帐表处方"
    strJson = strJson & "," & GetJsonNodeString("stuffdtl_ids", str费用IDs) ' C   1   处方明细id,目前传入的费用ID,允许多个"
    strJson = strJson & "," & GetJsonNodeString("operator_name", str操作员姓名) '   C   1   操作员姓名
    strJson = strJson & "," & GetJsonNodeString("operator_code", str操作员编号) '   C   1   操作员编号
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num) ' 1   病人ID
    strJson = strJson & "," & GetJsonNodeString("send_time", strCurDate)  'C   1   发料时间：yyyy-mm-dd hh24:mi:ss"
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"

    If GetServiceCall(objServiceCall) = False Then Exit Function
    If objServiceCall.CallService("Zl_Stuffsvr_Autosendstuff", strJson, , "", lngModule, False) = False Then Exit Function
    
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If

    zlStuffsvr_AutoSendStuffFromNo = True
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function

Public Function zlDrugSvr_GetExecutedNum(ByVal strNOs As String, ByVal bytBillType As Byte, _
    ByRef rsSended As ADODB.Recordset, Optional ByVal bln含零的已执行数 As Boolean = True, _
    Optional ByVal str费用IDs As String, Optional lngModule As Long, _
    Optional ByVal blnNotShowErr As Boolean, Optional ByRef strErrMsg_Out As String) As Boolean
    '获取准退数量
    '入参:
    '   strNos:单据号，多个用英文逗号分隔
    '   bytBillType:1-收费单,2-记帐单,3-自动记帐单
    '   bln含零的已执行数-true-包含执行数为零的药品及卫材,否则不包含
    '出参:
    '   rsSended:处方卫材数据：处方单号,药品ID,处方明细ID,已发数量,商品条码,内部条码
    '说明:
    Dim clldata  As Collection, cllTemp As Collection
    Dim strJson As String
    Dim i As Long
    Dim varValue As Variant

    Dim objServiceCall As Object
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
     
    Set rsSended = New ADODB.Recordset
    With rsSended.Fields
        .Append "处方单号", adVarChar, 50, adFldIsNullable
        .Append "处方明细ID", adBigInt, 18, adFldIsNullable
        .Append "药品ID", adBigInt, 18, adFldIsNullable
        .Append "已发数量", adDouble, , adFldIsNullable
        .Append "商品条码", adVarChar, 50, adFldIsNullable
        .Append "内部条码", adVarChar, 50, adFldIsNullable
    End With
    rsSended.CursorLocation = adUseClient
    rsSended.LockType = adLockOptimistic
    rsSended.CursorType = adOpenStatic
    rsSended.Open
        
    'zl_DrugSvr_GetExecutedNum
    '  --功能：获取药品已发放数量
    '  --入参：Json_In:格式
    '  --  input
    '  --    billtype         N   1 单据类型:1-收费处方发药;2-记帐单处方发药;3-记帐表处方发药
    '  --    rcp_nos          C   1 单据号:可以传入多张单据
    '  --    notcontain_zero       N       不包含已发药数为0的药品:1-不包含，0-包含
    '  --    rcpdtl_ids                 C   1 处方明细Ids:可以传入多张单据(费用Id)
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                 N   1   应答码：0-失败；1-成功
    '  --    message              C   1   应答消息：失败时返回具体的错误信息
    '  --    item_list[]
    '  --       rcp_no            C   1   处方单号(费用单据号)
    '  --       rcpdtl_id         N   1   处方明细ID(费用ID)
    '  --       drug_id           N   1   药品ID
    '  --       sended_num        N   1   已发数量
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("billtype", bytBillType, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rcp_nos", strNOs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("notcontain_zero", IIf(bln含零的已执行数, 0, 1), Json_num)
    strJson = strJson & "," & GetJsonNodeString("rcpdtl_ids", str费用IDs, Json_Text)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    If objServiceCall.CallService("zl_DrugSvr_GetExecutedNum", strJson, , "", mlngModule, Not blnNotShowErr) = False Then Exit Function
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If Not blnNotShowErr Then MsgBox strErrMsg_Out, vbInformation, gstrSysName
        Exit Function
    End If

    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        rsSended.AddNew
            rsSended!处方单号 = cllTemp("_rcp_no")
            rsSended!处方明细id = cllTemp("_rcpdtl_id")
            rsSended!药品ID = cllTemp("_drug_id")
            rsSended!已发数量 = cllTemp("_sended_num")
        rsSended.Update
    Next
    If rsSended.RecordCount > 0 Then rsSended.MoveFirst

    zlDrugSvr_GetExecutedNum = True
    Exit Function
ErrHandler:
    If blnNotShowErr Then strErrMsg_Out = Err.Description: Exit Function
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_GetExecutedNum(ByVal strNOs As String, ByVal bytBillType As Byte, _
    ByRef rsSended As ADODB.Recordset, Optional ByVal bln含零的已执行数 As Boolean = True, _
    Optional ByVal str费用IDs As String, Optional lngModule As Long, _
    Optional ByVal blnNotShowErr As Boolean, Optional ByRef strErrMsg_Out As String) As Boolean
    '获取准退数量
    '入参:
    '   strNos:单据号，多个用英文逗号分隔
    '   bytBillType:1-收费单,2-记帐单,3-自动记帐单
    '   bln含零的已执行数-true-包含执行数为零的药品及卫材,否则不包含
    '出参:
    '   rsSended:处方卫材数据：处方单号,药品ID,处方明细ID,已发数量,商品条码,内部条码
    '说明:
    Dim clldata  As Collection, cllTemp As Collection
    Dim strJson As String, rsOutData As ADODB.Recordset
    Dim i As Long
    Dim varValue As Variant

    Dim objServiceCall As Object
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    Set rsSended = New ADODB.Recordset
    With rsSended.Fields
        .Append "处方单号", adVarChar, 50, adFldIsNullable
        .Append "处方明细ID", adBigInt, 18, adFldIsNullable
        .Append "药品ID", adBigInt, 18, adFldIsNullable
        .Append "已发数量", adDouble, , adFldIsNullable
        .Append "商品条码", adVarChar, 50, adFldIsNullable
        .Append "内部条码", adVarChar, 50, adFldIsNullable
    End With
    rsSended.CursorLocation = adUseClient
    rsSended.LockType = adLockOptimistic
    rsSended.CursorType = adOpenStatic
    rsSended.Open
    

    Set rsOutData = Nothing
    'zl_StuffSvr_GetExecutedNum
    '  --功能：获取卫生材料已发放数量
    '  --入参：Json_In:格式
    '  --  input
    '  --     billtype             N 1 单据类型:1-收费处方发料；2-记帐单处方发料；3-记帐表处方发料
    '  --     stuff_nos            C 1 单据号:可以传入多张单据
    '  --     stuffdtl_ids         C 1 卫材单据Ids:可以传入多张单据(费用ID)
    '  --     notcontain_zero      N       不包含已发药数为0的药品:1-不包含，0-包含
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                  N   1   应答吗：0-失败；1-成功
    '  --    message               C   1   应答消息：失败时返回具体的错误信息
    '  --    item_list[]
    '  --       stuff_no           C   1   NO
    '  --       stuffdtl_id        N   1   费用ID
    '  --       stuff_id           N   1   卫材ID
    '  --       sended_num         N   1   已发数量
    '  --       barcode_goods      C   0   商品条码
    '  --       barcode_inside     C   0   内部条码
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("billtype", bytBillType, Json_num)
    strJson = strJson & "," & GetJsonNodeString("stuff_nos", strNOs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("stuffdtl_ids", str费用IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("notcontain_zero", IIf(bln含零的已执行数, 0, 1), Json_num)

    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    If objServiceCall.CallService("zl_StuffSvr_GetExecutedNum", strJson, , "", mlngModule, Not blnNotShowErr) = False Then Exit Function
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If Not blnNotShowErr Then MsgBox strErrMsg_Out, vbInformation, gstrSysName
        Exit Function
    End If
    
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        rsSended.AddNew
            rsSended!处方单号 = cllTemp("_stuff_no")
            rsSended!处方明细id = cllTemp("_stuffdtl_id")
            rsSended!药品ID = cllTemp("_stuff_id")
            rsSended!已发数量 = cllTemp("_sended_num")
            rsSended!商品条码 = cllTemp("_barcode_goods")
            rsSended!内部条码 = cllTemp("_barcode_inside")
        rsSended.Update
    Next
    If rsSended.RecordCount > 0 Then rsSended.MoveFirst

    zlStuffSvr_GetExecutedNum = True
    Exit Function
ErrHandler:
    If blnNotShowErr Then strErrMsg_Out = Err.Description: Exit Function
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub

Public Function zlDrugSvr_Getcargospace(ByVal lng收费项目ID As Long, ByVal lng库房ID As Long, _
    ByRef strCargospace As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定药品/卫材在指定库存的货位信息
    '入参:
    '     lng收费项目ID: bln卫材-true-卫生ID,false=药品ID
    '     lng库房ID-药房ID
    '     lngModule-模块号
    '出参:
    '   strCargospace:货位信息
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String

    On Error GoTo errHandle
    strCargospace = ""
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接药品服务失败，无法获取货位信息！", vbInformation, gstrSysName
        Exit Function
    End If

    'zl_DrugSvr_Getcargospace
    '  --功能：获取指定药品的指定库房的货位信息
    '  --入参：Json_In:格式
    '  --  input
    '  --   drug_id            N   1   药品ID
    '  --   pharmacy_id        N   1   库房ID
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code              N   1   应答吗：0-失败；1-成功
    '  --    message           C   1   应答消息：失败时返回具体的错误信息
    '  --    cargospace        N   1   库房货位
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("drug_id", lng收费项目ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_id", lng库房ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"

    If objServiceCall.CallService("zl_DrugSvr_Getcargospace", strJson, , "", mlngModule) = False Then Exit Function
    
    strCargospace = objServiceCall.GetJsonNodeValue("output.cargospace")
    
    zlDrugSvr_Getcargospace = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_RecipeAffirm(ByVal cllRecipeData As Collection, _
    ByVal lng病人ID As Long, ByVal str操作时间 As String, _
    ByVal str审核人编号 As String, ByVal str审核人姓名 As String, _
    Optional lngModule As Long, Optional strErrMsg_Out As String, _
    Optional ByVal cllExtend As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费、记帐等划价审核药品处方确认
    '入参:cllRecipeData-处方数据,每项是数组:
    '       Array(单据类型,单据号,费用IDs,发药窗口,是否自动发放,自动发放明细IDs)
    '           单据类型:1-收费处方发药;2-记帐单处方发药;3-记帐表处方发药
    '           发药窗口:发药窗口1:药房ID1|…|发药窗口n:药房Idn
    '     lng病人ID-病人ID，针对收费时有效
    '     str操作时间-当前操作的时间,审核时，代表审核时间
    '     str审核人编号、str审核人姓名 审核人信息
    '     lngModule-模块号
    '     cllExtend-扩展信息集，主要针对门诊转入了病人ID时，成员:(姓名,性别,年龄,门诊号)
    '出参:
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, varData As Variant, strJsonList As String
    Dim i As Long, j As Long
    Dim varValue As Variant

    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If cllRecipeData Is Nothing Then Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function

    'Zl_Drugsvr_Recipeaffirm
    '  --功能：药品处方记帐确认或药品处方收费确认
    '  --入参：Json_In:格式
    '  --  input
    '  --    pati_id                  N    病人id：针对收费时有效
    '  --    pati_name                C  0 姓名：针对收费时有效
    '  --    pati_sex                 C  0 性别：针对收费时有效
    '  --    pati_age                 C  0 年龄：针对收费时有效
    '  --    auditor                  C  1 审核人
    '  --    auditor_code             C  1 审核人编号
    '  --    audit_time               C  1 审核时间：yyyy-mm-dd hh24:mi:ss
    '  --    item_list[]                   更新数据列表[数组]
    '  --      billtype               N  1 单据类型:1-收费处方发药;2-记帐单处方发药;3-记帐表处方发药
    '  --      rcp_no                 C  1 单据号
    '  --      rcpdtl_ids             C  0 费用ID,可以传入多个,用逗号分离
    '  --      pharmacy_window        C  0 发药窗口:发药窗口1:药房ID1| …|发药窗口n:药房Idn
    '  --      drug_auto_send         N  0 是否自动发放药品:0-不自动发药,1-自动发药
    '  --      auto_send_ids          N  0 自动发药的明细ids，多个用逗号分隔
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code              N   1   应答吗：0-失败；1-成功
    '  --    message           C   1   应答消息：失败时返回具体的错误信息
    strJsonList = ""
    For i = 1 To cllRecipeData.Count
        varData = cllRecipeData(i)
        strJson = "": j = 0
        strJson = strJson & "" & GetJsonNodeString("billtype", Val(varData(j)), Json_num): j = j + 1
        strJson = strJson & "," & GetJsonNodeString("rcp_no", CStr(varData(j)), Json_Text): j = j + 1
        strJson = strJson & "," & GetJsonNodeString("rcpdtl_ids", CStr(varData(j)), Json_Text): j = j + 1
        strJson = strJson & "," & GetJsonNodeString("pharmacy_window", CStr(varData(j)), Json_Text): j = j + 1
        strJson = strJson & "," & GetJsonNodeString("drug_auto_send", Val(varData(j)), Json_num): j = j + 1
        strJson = strJson & "," & GetJsonNodeString("auto_send_ids", CStr(varData(j)), Json_Text): j = j + 1
        strJsonList = strJsonList & ",{" & strJson & "}"
    Next
    strJsonList = Mid(strJsonList, 2)
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    If lng病人ID <> 0 Then
        strJson = strJson & "," & GetJsonNodeString("pati_name", cllExtend("姓名"))
        strJson = strJson & "," & GetJsonNodeString("pati_sex", cllExtend("性别"))
        strJson = strJson & "," & GetJsonNodeString("pati_age", cllExtend("年龄"))
    End If
    strJson = strJson & "," & GetJsonNodeString("auditor", str审核人姓名)
    strJson = strJson & "," & GetJsonNodeString("auditor_code", str审核人编号)
    strJson = strJson & "," & GetJsonNodeString("audit_time", str操作时间)
    strJson = strJson & ",""item_list"":[" & strJsonList & "]"
    strJson = "{""input"":{" & strJson & "}}"

    If objServiceCall.CallService("zl_DrugSvr_RecipeAffirm", strJson, , "", lngModule, False) = False Then Exit Function
    
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    
    zlDrugSvr_RecipeAffirm = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_BillAffirm(ByVal cllRecipeData As Collection, _
    ByVal lng病人ID As Long, ByVal str操作时间 As String, _
    ByVal str审核人编号 As String, ByVal str审核人姓名 As String, _
    Optional ByVal lngModule As Long, Optional strErrMsg_Out As String, _
    Optional ByVal cllExtend As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费、记帐等划价审核卫材处方确认
    '入参:cllRecipeData-处方数据,每项是数组:
    '       Array(单据类型,单据号,费用IDs,是否自动发放,自动发放明细IDs)
    '           单据类型:1 -收费处方发药;2-记帐单处方发药;3-记帐表处方发药\
    '     lng病人ID-病人ID，针对收费时有效
    '     str操作时间-当前操作的时间,审核时，代表审核时间
    '     str审核人编号、str审核人姓名 审核人信息
    '     lngModule-模块号
    '     cllExtend-扩展信息集，主要针对门诊传入了病人ID时，成员:(姓名,性别,年龄,门诊号)
    '出参:
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 10:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, varData As Variant, strJsonList As String
    Dim i As Long, j As Long
    Dim varValue As Variant

    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If cllRecipeData Is Nothing Then Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Stuffsvr_billffirm
    '  --功能：卫材处方记帐确认或卫材处方收费确认
    '  --入参：Json_In:格式
    '  --  input
    '  --    pati_id                   N    病人id：针对收费时有效
    '  --    pati_name                 C  0 姓名：针对收费时有效
    '  --    pati_sex                  C  0 性别：针对收费时有效
    '  --    pati_age                  C  0 年龄：针对收费时有效
    '  --    pati_outpno               C  0 门诊号：针对收费时有效
    '  --    auditor                   C  1 审核人
    '  --    auditor_code              C  1 审核人编号
    '  --    audit_time                C  1 审核时间：yyyy-mm-dd hh24:mi:ss
    '  --    item_list[]                   更新数据列表[数组]
    '  --      billtype                N  1 单据类型:1 -收费处方发药 ;2-记帐单处方发药;3-记帐表处方发药
    '  --      stuff_no                C  1 单据号
    '  --      stuffdtl_ids            C  0 费用ID,可以传入多个,用逗号分离
    '  --      stuff_auto_send         N  0  卫材自动发料;0-不自动发料;1-自动发料
    '  --      auto_send_ids           N  0  自动发料的明细ids，多个用逗号分隔
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                      N   1   应答吗：0-失败；1-成功
    '  --    message                   C   1   应答消息：失败时返回具体的错误信息
    strJsonList = ""
    For i = 1 To cllRecipeData.Count
        varData = cllRecipeData(i)
        strJson = "": j = 0
        strJson = strJson & "" & GetJsonNodeString("billtype", Val(varData(j)), 1): j = j + 1
        strJson = strJson & "," & GetJsonNodeString("stuff_no", CStr(varData(j))): j = j + 1
        strJson = strJson & "," & GetJsonNodeString("stuffdtl_ids", CStr(varData(j))): j = j + 1
        strJson = strJson & "," & GetJsonNodeString("stuff_auto_send", Val(varData(j)), Json_num): j = j + 1
        strJson = strJson & "," & GetJsonNodeString("auto_send_ids", CStr(varData(j)), Json_Text): j = j + 1
        strJsonList = strJsonList & ",{" & strJson & "}"
    Next
    strJsonList = Mid(strJsonList, 2)

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    If lng病人ID <> 0 And Not cllExtend Is Nothing Then
        strJson = strJson & "," & GetJsonNodeString("pati_name", cllExtend("姓名"))
        strJson = strJson & "," & GetJsonNodeString("pati_sex", cllExtend("性别"))
        strJson = strJson & "," & GetJsonNodeString("pati_age", cllExtend("年龄"))
        strJson = strJson & "," & GetJsonNodeString("pati_outpno", cllExtend("门诊号"))
    End If
    strJson = strJson & "," & GetJsonNodeString("auditor", str审核人姓名)
    strJson = strJson & "," & GetJsonNodeString("auditor_code", str审核人编号)
    strJson = strJson & "," & GetJsonNodeString("audit_time", str操作时间)
    strJson = strJson & ",""item_list"":[" & strJsonList & "]"
    strJson = "{""input"":{" & strJson & "}}"
 
    If objServiceCall.CallService("zl_StuffSvr_billAffirm", strJson, , "", lngModule, False) = False Then Exit Function
    
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If

    zlStuffSvr_BillAffirm = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_CheckStoreLimit(ByVal lng收费项目ID As Long, ByVal lng库房ID As Long, ByVal dbl可用库存 As Double, _
    ByRef blnBelowLimitLower As Boolean, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定卫材在指定库房的库存是否低于储备下限
    '入参:
    '   lng收费项目ID 药品ID
    '   lng库房ID 药房ID
    '   dbl可用库存 当前可用库存数
    '出参:
    '   blnBelowLimitLower:低于储备下限为True，否则为False
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String

    On Error GoTo errHandle
    blnBelowLimitLower = False
    If lngModule <> 0 Then mlngModule = lngModule
    
   
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接药品服务失败，无法检查库存是否低于储备下限！", vbInformation, gstrSysName
        Exit Function
    End If

    'zl_DrugSvr_Checkstorelimit
    '  --功能：检查指定药品在指定库房的库存是否低于储备下限
    '  --入参：Json_In:格式
    '  --  input
    '  --   drug_id            N   1   药品ID
    '  --   pharmacy_id        N   1   库房ID
    '  --   stock              N   1   库存数量
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code              N 1 应答吗：0-失败；1-成功
    '  --    message           C 1 应答消息：失败时返回具体的错误信息
    '  --    below_limit_lower N 1 1-低于储备下限，0-大于等于储备下限
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("drug_id", lng收费项目ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_id", lng库房ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("stock", dbl可用库存, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("zl_DrugSvr_Checkstorelimit", strJson, , "", mlngModule) = False Then Exit Function
    
    blnBelowLimitLower = objServiceCall.GetJsonNodeValue("output.below_limit_lower") = 1
    zlDrugSvr_CheckStoreLimit = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_CheckStoreLimit(ByVal lng收费项目ID As Long, ByVal lng库房ID As Long, ByVal dbl可用库存 As Double, _
    ByRef blnBelowLimitLower As Boolean, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定卫材在指定库房的库存是否低于储备下限
    '入参:
    '   lng收费项目ID  材料ID
    '   lng库房ID 药房ID
    '   dbl可用库存 当前可用库存数
    '出参:
    '   blnBelowLimitLower:低于储备下限为True，否则为False
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String

    On Error GoTo errHandle
    blnBelowLimitLower = False
    If lngModule <> 0 Then mlngModule = lngModule
    
   
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Stuffsvr_Checkstorelimit
    '  --功能：检查指定卫材在指定库房的库存是否低于储备下限
    '  --入参：Json_In:格式
    '  --  input
    '  --    stuff_id            N   1   卫材ID
    '  --    warehouse_id        N   1   库房ID
    '  --    stock               N   1   库存数量
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N 1 应答吗：0-失败；1-成功
    '  --    message             C 1 应答消息：失败时返回具体的错误信息
    '  --    below_limit_lower   N 1 1-低于储备下限，0-大于等于储备下限
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("stuff_id", lng收费项目ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("warehouse_id", lng库房ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("stock", dbl可用库存, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_Stuffsvr_Checkstorelimit", strJson, , "", mlngModule) = False Then Exit Function
    
    blnBelowLimitLower = objServiceCall.GetJsonNodeValue("output.below_limit_lower") = 1
    zlStuffSvr_CheckStoreLimit = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_GetStockCheck(ByRef cllStockCheck As Collection, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品的库存检查方式
    '入参:
    '出参:
    '   cllStockCheck:
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim clldata  As Collection, cllTemp As Collection
    Dim objServiceCall As Object
    Dim strJson As String
    Dim i As Long
    
    On Error GoTo ErrHandler
    Set cllStockCheck = New Collection
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接药品服务失败，无法获取卫生材料的库存检查方式！", vbInformation, gstrSysName
        Exit Function
    End If

    'zl_DrugSvr_Getstockcheck
    '  --功能：获取药品库存检查方式
    '  --入参：Json_In:格式
    '  --  input
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N   1 应答吗：0-失败；1-成功
    '  --    message             C   1 应答消息：失败时返回具体的错误信息
    '  --    item_list                   [数组]
    '  --        pharmacy_id     N   1   药房ID
    '  --        check_type      N   1   检查方式：0-不检查，1-检查提示名，2-检查禁止
    strJson = "{""input"":null}"
    If objServiceCall.CallService("zl_DrugSvr_Getstockcheck", strJson, , "", mlngModule) = False Then Exit Function
    
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then zlDrugSvr_GetStockCheck = True: Exit Function
       
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        cllStockCheck.Add cllTemp("_check_type"), "_" & cllTemp("_pharmacy_id")
    Next
    zlDrugSvr_GetStockCheck = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_GetStockCheck(ByRef cllStockCheck As Collection, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卫材的库存检查方式
    '入参:
    '出参:
    '   cllStockCheck:
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim clldata  As Collection, cllTemp As Collection
    Dim objServiceCall As Object
    Dim strJson As String
    Dim i As Long
    
    On Error GoTo ErrHandler
    Set cllStockCheck = New Collection
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Stuffsvr_Getstockcheck
    '  --功能：获取卫材库存检查方式
    '  --入参：Json_In:格式
    '  --  input
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N 1 应答吗：0-失败；1-成功
    '  --    message             C 1 应答消息：失败时返回具体的错误信息
    '  --    item_list             [数组]
    '  --        warehouse_id    N   1   库房ID
    '  --        check_type      N   1   检查方式：0-不检查，1-检查提示名，2-检查禁止
    strJson = "{""input"":null}"
    If objServiceCall.CallService("Zl_Stuffsvr_Getstockcheck", strJson, , "", mlngModule) = False Then Exit Function
    
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then zlStuffSvr_GetStockCheck = True: Exit Function
       
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        cllStockCheck.Add cllTemp("_check_type"), "_" & cllTemp("_warehouse_id")
    Next
    zlStuffSvr_GetStockCheck = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffsvr_GetStuffPutInBills(ByVal bytQueryType As Byte, ByRef clldata As Collection, Optional ByVal lng库房ID As Long, Optional ByVal strStuffNo As String, _
                Optional ByVal datStart As Date, Optional ByVal datEnd As Date, Optional ByVal lngModule As Long, Optional ByVal int单据类型 As Integer) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取备货卫材的入库单
    ' 入参 : query_type :查询方式：0-仅按入库单号查询,1-查询时包含入库时间,2-最后一次备货且有库存的入库单,3-查询收费、记账单
    '        strNo: 入库单号
    '        lng库房ID :库房ID，虚拟库房
    '        datStart:入库单查询开始时间
    '       end_time：入库单查询终止时间
    '       int单据类型: 药品收发记录.单据，查询方式3时使用
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/8/10 16:38
    '---------------------------------------------------------------------------------------
'    --功能：获取备货卫材的单据信息
'    --入参：Json_In:格式
'    --  input
'    --   query_type         N   1 查询方式：
'    --                                     查询入库单据：0-仅按入库单号查询,1-查询时包含入库时间,2-最后一次备货且有库存的入库单
'    --                                     查询卫材收费、记账单：3-查询收费、记账单，需要配合billType使用
'    --   warehouse_id       N   1 库房ID，虚拟库房
'    --   stuff_no             C   1 入库单号
'    --   begin_time         C   1 入库单查询开始时间
'    --   end_time           C   1 入库单查询终止时间
'    --   billType           N     单据类型 1-收费单;2-记账单
'    --出参: Json_Out,格式如下
'    --  output
'    --    code              N   1   应答吗：0-失败；1-成功
'    --    message           C   1   应答消息：失败时返回具体的错误信息
'    --    item_list                入库记录
'    --      bill_id         N   1 卫材入库单id
'    --      bill_no         C   1 卫材入库单号
'    --      serial_number   N     序号
'    --      origin          C     产地
'    --      provider        C     供应商
'    --      batch_number    C     批号
'    --      production_date C     生产日期
'    --      expiry_date     C     效期
'    --      sterilization_expiry_date
'    --                      C     灭菌效期
'    --      putin_quantity  N     入库数量
'    --      putin_price     N     入库零售价
'    --      putin_money     N     入库零售金额
'    --      audit_date      C     审核日期
'    --      stuff_id        N   1 卫材ID
'    --      batch           C     批次
'    --      barcode_goods   C     商品条码
'    --      barcode_inside  C     内部条码
'    --      stock           N     可用库存
'    --      stuffdtl_id     N     处方明细id

    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo errH
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function

    '获取入参json
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", bytQueryType, Json_num)
    strJson = strJson & "," & GetJsonNodeString("warehouse_id", lng库房ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("strStuffNo", strStuffNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("begin_time", Format(datStart, "yyyy-mm-dd HH:MM:SS"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("end_time", Format(datEnd, "yyyy-mm-dd HH:MM:SS"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("billType", int单据类型, Json_num)
    strJson = "{""input"":{" & strJson & "}}"

     
    If objServiceCall.CallService("Zl_Stuffsvr_Getstuffputinbills", strJson, , "", mlngModule) = False Then Exit Function
    
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then zlStuffsvr_GetStuffPutInBills = True: Exit Function
    
    zlStuffsvr_GetStuffPutInBills = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffsvr_CheckContainStuff(ByVal str费用IDs As String, ByRef blnContainStuff As Boolean, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否含备货卫材
    '入参:str费用IDs-处方明细ID,目前传入的费用ID串，用逗号分隔 ：1,2,3,4
    '出参:blnContainStuff:True-含有备货卫材：false-不含
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String

    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function

    'Zl_Stuffsvr_Checkcontainstuff
    '  --功能：检查单据是否含备货卫材
    '  --入参：Json_In:格式
    '  --  input
    '  --    stuffdtl_ids      C   0  单据明细id串,目前传入的费用ID串，用逗号分隔 ：1,2,3,4
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code              N   1   应答吗：0-失败；1-成功
    '  --    message           C   1   应答消息：失败时返回具体的错误信息
    '  --    contain_stuff     N   1   是否含备货卫材:1-含有备货材料,0-不含备货材料
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("stuffdtl_ids", str费用IDs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("zl_StuffSvr_CheckContainStuff", strJson, , "", mlngModule) = False Then Exit Function

    blnContainStuff = objServiceCall.GetJsonNodeValue("output.contain_stuff") = 1
    zlStuffsvr_CheckContainStuff = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_GetPharmacyWindows(ByVal str药房IDs As String, ByRef rsData As ADODB.Recordset, _
    Optional lngModule As Long) As Boolean
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
    Dim clldata As Collection, cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
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
    If objServiceCall.CallService("zl_DrugSvr_Getpharmacywindows", strJson, , "", mlngModule) = False Then Exit Function
    
    Set clldata = objServiceCall.GetJsonListValue("output.window_list")
    If clldata Is Nothing Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
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
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_GetSendWindows(ByVal bytBillType As Byte, _
    ByVal lng病人ID As Long, ByVal int挂号有效天数 As Integer, ByVal str药房IDs As String, _
    ByRef rsData As ADODB.Recordset, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:确定药房的发药窗口
    '入参:
    '   bytBillType 单据类型：1 -收费处方发药  ;2- 记帐单处方发药;3- 记帐表处方发药
    '   str药房IDs 药房信息，格式：药房ID1,缺省发药窗口|药房ID2,缺省发药窗口2|...
    '出参:
    '   rsData 字段：药房ID,发药窗口
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, cllTemp As Collection, i As Long
    Dim strJsonTemp As String, varData As Variant, varTmp As Variant
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule

    Set rsData = New ADODB.Recordset
    With rsData.Fields
        .Append "药房ID", adBigInt, 18, adFldIsNullable
        .Append "发药窗口", adLongVarChar, 50, adFldIsNullable
    End With
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open

    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接药品服务失败，无法确定发药窗口！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If str药房IDs = "" Then Exit Function
    'Zl_Drugsvr_Getsendwindows
    '  --功能：获取药品的发药窗口
    '  --入参：Json_In:格式
    '  --  input
    '  --    item_list[]
    '  --        billtype        N   1   单据类型: 1 -收费处方发药  ;2- 记帐单处方发药;3- 记帐表处方发药
    '  --        pati_id         N   1   病人id
    '  --        valid_days      N       未发药品记录查询范围的有效天数
    '  --        pharmacy_id     N   1   药房id
    '  --        defaultwindow   C       根据各业务模块参数设置中传入的缺省窗口
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N   1 应答吗：0-失败；1-成功
    '  --    message             C   1 应答消息：失败时返回具体的错误信息
    '  --    item_list                   [数组]每个库房ID对应的发药窗口
    '  --        pharmacy_id     N   1   库房ID
    '  --        pharmacy_window C   1   库房ID对应的发药窗口
    '  --窗口获取规则:
    '  --    1)查找指定病人在指定药房中分配了发药窗口且该窗口处于上班的最后一次未发药品的窗口，存在则取该发药窗口
    '  --    2)检查缺省窗口是否处于上班的，是上班的，取该发药窗口
    '  --    3)根据药房窗口的分配规则（忙闲/平均 ），取发药窗口
    strJson = ""
    varData = Split(str药房IDs, "|")
    For i = 0 To UBound(varData)
        varTmp = Split(varData(i) & ",", ",")
        strJsonTemp = ""
        strJsonTemp = strJsonTemp & "" & GetJsonNodeString("billtype", bytBillType, Json_num)
        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("valid_days", int挂号有效天数, Json_num)
        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("pharmacy_id", varTmp(0), Json_num)
        strJsonTemp = strJsonTemp & "," & GetJsonNodeString("defaultwindow", varTmp(1), Json_Text)
        strJson = strJson & ",{" & strJsonTemp & "}"
    Next
    strJson = "{""input"":{""item_list"":[" & Mid(strJson, 2) & "]}}"
    If objServiceCall.CallService("Zl_Drugsvr_Getsendwindows", strJson, , "", mlngModule) = False Then Exit Function
    
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        rsData.AddNew
        rsData!药房Id = cllTemp("_pharmacy_id")
        rsData!发药窗口 = cllTemp("_pharmacy_window")
        rsData.Update
    Next
    If rsData.RecordCount > 0 Then rsData.MoveFirst

    zlDrugSvr_GetSendWindows = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_NewRecipeBill_Check(ByVal strJsonAsk As String, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '生成药品处方单据数据检查
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    
    On Error GoTo ErrHandler
    If strJsonAsk = "" Then Exit Function
    
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If objServiceCall.CallService("Zl_Drugsvr_Newrecbill_Check", strJsonAsk, , "", mlngModule) = False Then Exit Function
    
    zlDrugSvr_NewRecipeBill_Check = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_NewBill_Check(ByVal strJsonAsk As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '生成材料发料单据检查
    '入参:
    '   cllRcpBillLists-处方明细集,元素为数组：Array(单据类型,材料id,付数,数量,药房ID,病人姓名,医嘱ID,是否备货卫材,批次)
    '       单据类型:1-收费处方,2-记账处方
    '       是否备货卫材:有高值卫材才需要传入，非0表示是高值卫材模式(如扫码时使用)
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    
    On Error GoTo ErrHandler
    If strJsonAsk = "" Then Exit Function
    
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    
    If objServiceCall.CallService("Zl_Stuffsvr_Newbill_Check", strJsonAsk, , "", mlngModule) = False Then Exit Function
    
    zlStuffSvr_NewBill_Check = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_NewRecipeBill(ByVal strJsonAsk As String, Optional strErrMsg_Out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:生成药品处方单据
    '入参:strJsonAsk-请求的Json串
    '出参:strErrMsg_out-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-12 21:55:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, varValue As Variant
    
    Err = 0: On Error GoTo errHandle:
    If strJsonAsk = "" Then Exit Function
    
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
     
    If objServiceCall.CallService("zl_DrugSvr_NewRecipeBill", strJsonAsk, , "", lngModule, False) = False Then Exit Function
    
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlDrugSvr_NewRecipeBill = True
    Exit Function
errHandle:
    strErrMsg_Out = Err.Description
End Function

Public Function zlStuffSvr_NewBill(ByVal strJsonAsk As String, _
    Optional strErrMsg_Out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:生成卫材发料单据
    '入参:strJsonAsk-Json请求串
    '出参:strErrMsg_out-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-12 21:55:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, varValue As Variant
        
    On Error GoTo ErrHandler
    If strJsonAsk = "" Then Exit Function
    
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    If objServiceCall.CallService("zl_StuffSvr_NewBill", strJsonAsk, , "", lngModule, False) = False Then Exit Function
    
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlStuffSvr_NewBill = True
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function

Public Function zlDrugsvr_AutoSplitSpeci(ByVal lng药名ID As Long, ByVal byt中药形态 As Byte, _
    ByVal int付数 As Integer, ByVal dbl数量 As Double, ByVal lng药房id As Long, _
    ByRef strData As String, Optional ByVal byt场合 As Byte = 1, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:针对中草药根据药名（品种）来自动分配药品(自动分解多个规格)
    '入参:
    '   byt中药形态 0-散装;1-中药饮片;2-免煎剂
    '   byt场合 1-门诊 ，2-住院
    '出参:
    '   strData 格式：药品id,数量;药品id,数量;...(散装只选择一个规格)
    '               不能完全分配时返回:剂量为6和10的情况下,17克的分配=23755,6;23756,10|1
    '               不能分配时返回空,例如:剂量为6和10的情况下,3克的分配
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule

    strData = ""
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接药品服务失败，无法自动分配药品！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Drugsvr_Autosplitspeci
    '  --功能：针对中草药根据药名（品种）来自动分配药品(自动分解多个规格)
    '  --入参：Json_In:格式
    '  --  input
    '  --    clinic_drug_id     N 1 药名ID
    '  --    form              N 1 形态：0-散装;1-中药饮片;2-免煎剂
    '  --    quantity          N 1 数量，按剂量单位传入
    '  --    packages_num      N 1 付数
    '  --    pharmacy_id       N 1 药房ID
    '  --    occasion          N 1 场合：1-门诊 ，2-住院
    '  --    drug_ids          C   指定药品的分配
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code              N   1   应答吗：0-失败；1-成功
    '  --    message           C   1   应答消息：失败时返回具体的错误信息
    '  --    item_list                 不能分配时该节点返回空
    '  --      drug_id           N   1   药品ID
    '  --      quantity          N   1   数量
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("clinic_drug_id", lng药名ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("form", byt中药形态, Json_num)
    strJson = strJson & "," & GetJsonNodeString("quantity", dbl数量, Json_num)
    strJson = strJson & "," & GetJsonNodeString("packages_num", int付数, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_id", lng药房id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("occasion", byt场合, Json_num)
    strJson = strJson & "," & GetJsonNodeString("drug_ids", "", Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_Drugsvr_Autosplitspeci", strJson, , "", mlngModule) = False Then Exit Function
    
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        '药品id,数量;药品id,数量;...
        strData = strData & ";" & cllTemp("_drug_id") & "," & cllTemp("_quantity")
    Next
    If strData <> "" Then strData = Mid(strData, 2)

    zlDrugsvr_AutoSplitSpeci = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugsvr_AutoRecipeSendDrug(ByVal bytBillType As Byte, ByVal strNo As String, _
    ByVal lng库房ID As Long, ByVal str操作员姓名 As String, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动处方发药
    '入参:
    '   bytBillType 单据类型:1-收费单,2-记帐单
    '出参:
    '返回:执行成功返回True，失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接药品服务失败，无法自动发药！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Drugsvr_Autorecipesenddrug
    '  --功能：处方自动发药
    '  --入参：Json_In:格式
    '  --  input
    '  --    pharmacy_id           N 1 药房ID
    '  --    billtype              N 1 1 -收费处方发药  ;2- 记帐单处方发药;3- 记帐表处方发药;
    '  --    rcp_no                C 1 药品收发记录.no
    '  --    operator_name         C 1 操作员姓名
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                  N 1   应答吗：0-失败；1-成功
    '  --    message               C 1   应答消息：失败时返回具体的错误信息
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pharmacy_id", lng库房ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("billtype", bytBillType, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rcp_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_name", str操作员姓名, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_Drugsvr_Autorecipesenddrug", strJson, , "", mlngModule) = False Then Exit Function
    
    zlDrugsvr_AutoRecipeSendDrug = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffsvr_AutoRecipeSendStuff(ByVal bytBillType As Byte, ByVal strNo As String, _
    ByVal lng库房ID As Long, ByVal str发料人 As String, ByVal str发料时间 As String, _
    Optional ByVal byt发料方式 As Byte = 1, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动处方发料
    '入参:
    '   bytBillType 单据类型:1-收费单,2-记帐单
    '   byt发料方式 1-处方发料;2-批量发料;3-部门发料
    '   str发料时间 yyyy-mm-dd hh:mi:ss
    '出参:
    '返回:执行成功返回True，失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接卫材服务失败，无法自动发料！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Stuffsvr_Autosendstuff
    '  --功能：收费或记帐后，自动发料
    '  --入参：Json_In:格式
    '  --  input
    '  --    warehouse_ID         N   库房ID
    '  --    billtype             N 1 单据类型:1 -收费处方发药  ;2- 记帐单处方发药;3- 记帐表处方发药;
    '  --    stuff_no             C 1 NO
    '  --    send_name            C 1 发料人
    '  --    send_time            C 1 发料时间:yyyy-mm-dd hh:mi:ss
    '  --    send_mode            N 1 发料方式: 1-处方发料;2-批量发料;3-部门发料 以前有汇总发药号，现在取消，服务内自行处理
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                 N 1   应答吗：0-失败；1-成功
    '  --    message              C 1   应答消息：失败时返回具体的错误信息
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("warehouse_ID", lng库房ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("billtype", bytBillType, Json_num)
    strJson = strJson & "," & GetJsonNodeString("stuff_no", strNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("send_name", str发料人, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("send_time", str发料时间, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("send_mode", byt发料方式, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_Stuffsvr_Autosendstuff", strJson, , "", mlngModule) = False Then Exit Function
    
    zlStuffsvr_AutoRecipeSendStuff = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugsvr_CheckIsputDrug(ByVal bytBillType As Byte, ByVal strNOs As String, _
    ByRef blnIsPutDrug As Boolean, _
    Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查药品是否已进行配药
    '入参:
    '   bytBillType 单据类型:1-收费处方；2-记帐单处方；3-记帐表处方
    '   strNOs 处方单据，多个用英文逗号分隔
    '出参:
    '   blnIsPutDrug 是否已进行配药
    '返回:执行成功返回True，否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    blnIsPutDrug = False
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接药品服务失败，无法确定药品是否已进行配药！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_Drugsvr_Checkisputdrug
    '  --功能：检查药品是否已进行配药
    '  --入参：Json_In:格式
    '  --  input
    '  --   rcp_nos                   C  1  药品收发记录.no，多个用英文逗号分隔
    '  --   billtype                    N  1  1-收费处方；2-记帐单处方；3-记帐表处方
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code              N   1   应答吗：0-失败；1-成功
    '  --    message           C   1   应答消息：失败时返回具体的错误信息
    '  --    isputdrug         N   1   是否已配药：1-已配药,0-未配药
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("billtype", bytBillType, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rcp_nos", strNOs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_Drugsvr_Checkisputdrug", strJson, , "", mlngModule) = False Then Exit Function
    
    blnIsPutDrug = objServiceCall.GetJsonNodeValue("output.isputdrug")
    
    zlDrugsvr_CheckIsputDrug = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_GetIDbyBarcode(ByVal strBarCode As String, ByRef clldata As Collection, _
    Optional ByVal bln仅内部条码 As Boolean, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据输入的条码获取卫生卫材信息(卫材，批次)
    '入参:strbarcode-条码串
    '       bln仅内部条码-仅根据内部条码进行查找
    '出参:cllData-返回:形式:材料D:批次ID
    '返回:获取成功，返回true,否则返回False
    '编制:焦博
    '日期:2019-08-14 25:28:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo errHandle
    
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Stuffsvr_Getidbybarcode
    '  --功能：根据商品条码或内部条码获取卫材ID及批次
    '  --入参：Json_In:格式
    '  --  input
    '  --    barcode             C 1 输入的条码串
    '  --    type                N 0 1-不检查库存及期效,0-检查库存及期效
    '  --    only_barcode_inside N 0 1-仅对内部条码进行查找,0-对商品条码及内部条码进行查找
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code              N 1   应答吗：0-失败；1-成功
    '  --    message           C 1   应答消息：失败时返回具体的错误信息
    '  --    item_list
    '  --      stuff_id        N 1   材料ID
    '  --      batch           N     批次：0-不区分批次，>0只查某批次
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("barcode", strBarCode, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("type", 0, Json_num)
    strJson = strJson & "," & GetJsonNodeString("only_barcode_inside", IIf(bln仅内部条码, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_StuffSvr_GetIDbyBarcode", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list", "")
    If clldata Is Nothing Then Exit Function
        
    zlStuffSvr_GetIDbyBarcode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffsvr_CheckExistsbarcode(ByVal str卫材IDs As String, ByVal strCode As String, _
    ByVal bln仅内部条码 As Boolean, ByRef clldata As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定卫材在药品库存中是否存在商品条码或内部条码
    '入参:str卫材IDs-卫材IDs
    '       strCode-需要检查的条码串
    '       bln仅内部条码-仅根据内部条码进行查找
    '出参:strStockInfor-返回:形式:卫材ID1,卫材ID2...卫材IDn
    '返回:获取成功，返回true,否则返回False
    '编制:焦博
    '日期:2019-08-14 14:28:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objServiceCall As Object
    Dim strJson As String, varData As Variant
    Dim strJson_out As String
    
    On Error GoTo errHandle
    
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Stuffsvr_Checkexistsbarcode
    '  --功能：检查指定卫材在药品库存中是否存在商品条码或内部条码
    '  --入参：Json_In:格式
    '  --  input
    '  --    stuff_ids           C 1 输入的卫材ids,多个用逗号分隔
    '  --    barcode             C 0 当前查询的条码串
    '  --    only_barcode_inside N 0 1-仅对内部条码进行查找,0-对商品条码及内部条码进行查找
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N 1   应答吗：0-失败；1-成功
    '  --    message             C 1   应答消息：失败时返回具体的错误信息
    '  --    stuff_ids           C 1  返回的药品库存中存在商品条码或内部条码的卫材ids,多个用逗号分隔
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("stuff_ids", str卫材IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("barcode", strCode, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("only_barcode_inside", IIf(bln仅内部条码, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Stuffsvr_CheckExistsbarcode", strJson, , "", mlngModule) = False Then Exit Function
    strJson_out = objServiceCall.GetJsonNodeValue("output.stuff_ids")
    If Trim(strJson_out) = "" Then Exit Function
    varData = Split(strJson_out, ",")

    Set clldata = New Collection
    For i = 0 To UBound(varData)
        If Val(varData(i)) <> 0 Then clldata.Add varData(i), "_" & varData(i)
    Next
    zlStuffsvr_CheckExistsbarcode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_DelRecipeBill(ByVal strJsonAsk As String, Optional strErrMsg_Out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废药品处方单据
    '入参:strJsonAsk-请求的Json串
    '出参:strErrMsg_out-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-12 21:55:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, varValue As Variant
    
    Err = 0: On Error GoTo errHandle:
    If strJsonAsk = "" Then Exit Function
    
    If lngModule <> 0 Then mlngModule = lngModule
    
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
     
    If objServiceCall.CallService("zl_DrugSvr_DelRecipeBill", strJsonAsk, , "", lngModule, False) = False Then Exit Function
    
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlDrugSvr_DelRecipeBill = True
    Exit Function
errHandle:
    strErrMsg_Out = Err.Description
End Function

Public Function zlStuffSvr_DelBill(ByVal strJsonAsk As String, _
    Optional strErrMsg_Out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废卫材发料单据
    '入参:strJsonAsk-Json请求串
    '出参:strErrMsg_out-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-12 21:55:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, varValue As Variant
        
    On Error GoTo errHandle
    If strJsonAsk = "" Then Exit Function
    
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    If objServiceCall.CallService("zl_StuffSvr_DelBill", strJsonAsk, , "", lngModule, False) = False Then Exit Function
    
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlStuffSvr_DelBill = True
    Exit Function
errHandle:
    strErrMsg_Out = Err.Description
End Function

Public Function ZlDrugsvr_GetStockByDrugName(ByVal lng药名ID As Long, ByVal lng药房id As Long, int显示单位 As Integer, _
ByRef cllStockData_Out As Collection, ByVal int场合 As Integer, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据药名,获取指定库存信息
    '入参:lng药名ID-
    '     lng药房ID-药房ID,=0时表示不区分药房
    '     int场合-1-门诊;2-住院
    '     int显示单位-0-售价单位;1-门诊单位;2-住院单位
    '出参:cllStockData_Out-返回数据集，每个成员的key(pharmacy_id,drug_id,stock)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 17:30:45
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo errHandle
    
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'input
    '    clinicdrug_id   N   1   药名ID
    '    pharmacy_id N   1   药房ID；药房ID=0，表示所有的
    '    occasion    N   1   场合：1-门诊 ，2-住院
    '    show_unit   N   1   显示单位:0-售价单位;1-住院单位;2-门诊单位
    '    site_no C   1   站点号

    '出参 json
    'output
        'code    N   1   应答码：0-失败；1-成功
    '   message C   1   应答消息：失败时返回具体的错误信息
    '   item_list
    '       pharmacy_id N   1   药房ID
    '       drug_id N   1   药品ID
    '       stock   N   1   可用数量
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("clinicdrug_id", lng药名ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pharmacy_id", lng药房id, 1)
    strJson = strJson & "," & GetJsonNodeString("occasion", int场合, 1) '  病人ID
    strJson = strJson & "," & GetJsonNodeString("show_unit", int显示单位, 1) '  :0-售价单位;1-住院单位;2-门诊单位
    strJson = strJson & "," & GetJsonNodeString("site_no", IIf(gobjComlib.gstrNodeName = "-", "", gobjComlib.gstrNodeName)) '  站点号
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
       
    If objServiceCall.CallService("Zl_Drugsvr_GetStockByDrugName", strJson, , "", mlngModule) = False Then Exit Function
    Set cllStockData_Out = objServiceCall.GetJsonListValue("output.item_list")
      
    ZlDrugsvr_GetStockByDrugName = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPivasSvr_Infusion_Update(ByVal cllUpdataDatas As Collection, Optional strErrMsg_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:该服务主要是更新输液配药记录中的：操作员，操作时间及操作状态。
    '入参:cllUpdataDatas-格式如下
    '        |-元素(配药ID,操作员姓名,操作说明,操作时间,操作类型(0-销帐申请;1-销帐申请取消(删除);2-销帐申请审核),销帐申请时间(销帐审核时传入),
    '        |- advice_list(collect)医嘱信息列表：每行元素为(医嘱ID,发送号) :传入了配药ID的或操作类型<>1时，此列表无效表示: 已申请的批次一起取消
    '
    '出参:strErrMsg_out-失败时出错信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-20 10:34:23
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim clldata As Collection, cllTemp As Collection
    Dim i As Long, objServiceCall As Object
    Dim strJson As String, strJsonTemp As String, strJsonAdviceList As String
        
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    If cllUpdataDatas Is Nothing Then Exit Function
    If cllUpdataDatas.Count = 0 Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input
    '        dispensing_id                   C   0   "配药ID，如果传入空，则 advice_list[]必传 销帐申请时，必传入"
    '        type                    N   1   0-销帐申请;1-销帐申请取消(删除);2-销帐申请审核
    '        operator_name                   C   1    操作员姓名
    '        operator_notes                  C       操作说明
    '        operate_time                    C   1   操作时间：yyyy-mm-dd hh24:mi:ss
    '        apply_time                  C   1   申请时间(销帐申请审核时有效)：yyyy-mm-dd hh24:mi:ss
    '         advice_list[]                          "医嘱信息列表:传入了配药ID的或type<>1时，此列表无效     表示:     已申请的批次一起取消
    '            advice_id               N   1   医嘱Id
    '            send_no             N   1   发送号
    '    output
    '        code                    N   1   应答吗：0-失败；1-成功
    '        message                 C   1   应答消息：失败时返回具体的错误信息
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("dispensing_id", zlGetNodeValueFromCollect(cllUpdataDatas, "配药ID", "N"), 1)
    strJson = strJson & "," & GetJsonNodeString("type", zlGetNodeValueFromCollect(cllUpdataDatas, "操作类型", "N"), 1)
    strJson = strJson & "," & GetJsonNodeString("operator_name", zlGetNodeValueFromCollect(cllUpdataDatas, "操作员姓名", "C"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_notes", zlGetNodeValueFromCollect(cllUpdataDatas, "操作说明", "C"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operate_time", zlGetNodeValueFromCollect(cllUpdataDatas, "操作时间", "C"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operate_time", zlGetNodeValueFromCollect(cllUpdataDatas, "销帐申请时间", "C"), Json_Text)
    
    strJsonAdviceList = ""
    Set clldata = zlGetNodeObjectFromCollect(cllUpdataDatas, "advice_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            strJsonTemp = "{"
            strJsonTemp = strJsonTemp & "" & GetJsonNodeString("advice_id", cllTemp("医嘱ID"), Json_num)
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("send_no", cllTemp("发送号"), Json_num)
            strJsonTemp = strJsonTemp & "}"
            strJsonAdviceList = strJsonAdviceList & "," & strJsonTemp
        Next
        If strJsonAdviceList <> "" Then
            strJsonAdviceList = Mid(strJsonAdviceList, 2)
            strJson = strJson & "," & GetNodeString("advice_list") & ":[" & strJsonAdviceList & "]"
        End If
    End If
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    '执行修正
    If objServiceCall.CallService("zl_PivasSvr_Infusion_Update", strJson, , "", lngModule, False) = False Then Exit Function
    
    If objServiceCall.GetJsonNodeValue("output.code") = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlPivasSvr_Infusion_Update = True
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function

Public Function zlDrugSvr_GetRefuseSendList(ByVal lng病人ID As Long, ByVal str主页IDs As String, ByRef str费用IDs_out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的拒发药清单
    '入参:lng病人ID-病人ID
    '     str主页Ids-多个用逗号
    '出参:str费用IDs_out-返回拒发药的卫生材料及药品有的费用ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim clldata As Collection, cllTemp As Collection
    Dim i As Long, objServiceCall As Object
    Dim strJson As String
        
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    str费用IDs_out = ""
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
  
    'Zl_Drugsvr_Getrefusesendlist
    '  --功能：获取拒发药清单
    '  --入参：Json_In:格式
    '  --  input
    '  --     pati_id            N 1 病人Id
    '  --     pati_pageids       C 1 主页IDs:多次住院 ，用逗号分离
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N   1   应答吗：0-失败；1-成功
    '  --    message             C   1   应答消息：失败时返回具体的错误信息
    '  --    item_list []
    '  --        rcp_no          C   1   费用单据号
    '  --        rcpdtl_id       C   1   处方明细ID,目前传入的是费用ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, 1)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", str主页IDs)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_DrugSvr_GetRefuseSendList", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            If InStr(str费用IDs_out & ",", "," & str费用IDs_out & ",") > 0 Then
                str费用IDs_out = str费用IDs_out & "," & cllTemp("_rcpdtl_id")
            End If
        Next
    End If
    zlDrugSvr_GetRefuseSendList = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter()() = 1 Then
        Resume
    End If
End Function

Public Function zlStuffSvr_GetRefuseSendList(ByVal lng病人ID As Long, ByVal str主页IDs As String, ByRef str费用IDs_out As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的拒发药清单
    '入参:lng病人ID-病人ID
    '     str主页Ids-多个用逗号
    '出参:str费用IDs_out-返回拒发药的卫生材料及药品有的费用ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim clldata As Collection, cllTemp As Collection
    Dim i As Long, objServiceCall As Object
    Dim strJson As String
        
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    str费用IDs_out = ""
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
  
    'Zl_Stuffsvr_Getrefusesendlist
    '  --功能：获取拒发药清单
    '  --入参：Json_In:格式
    '  --  input
    '  --     pati_id  N 1 病人Id
    '  --     pati_pageids C 1 主页IDs:多次住院 ，用逗号分离
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                    N   1   应答吗：0-失败；1-成功
    '  --    message                 C   1   应答消息：失败时返回具体的错误信息
    '  --    item_list []
    '  --        stuff_no          C   1   费用单据号
    '  --        stuffdtl_id          C   1   费用ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, 1)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", str主页IDs)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_StuffSvr_GetRefuseSendList", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            If InStr(str费用IDs_out & ",", "," & cllTemp("_stuffdtl_id") & ",") > 0 Then
                str费用IDs_out = str费用IDs_out & "," & cllTemp("_stuffdtl_id")
            End If
        Next
    End If
    zlStuffSvr_GetRefuseSendList = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter()() = 1 Then
        Resume
    End If
End Function


Public Function zlPivasSvr_Getinfusion_Record(ByVal strNOs As String, _
    ByRef str费用IDs_out As String, Optional ByVal lng病人ID As Long, Optional ByVal str主页IDs As String, _
    Optional ByVal str费用Ids_in As String = "", Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的拒发药清单
    '入参:lng病人ID-按病人ID查(str主页Ids-多个用逗号)
    '     strNos-按单据查
    '     str费用Ids_in-按费用ID查
    '     以上条件，必有一个
    '出参:str费用IDs_out-返回所有涉及进入输液配药记录中的费用IDs
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
        
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    str费用IDs_out = ""
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Pivassvr_Getinfusion_Record
    '  --功能：根据病人信息或处方号信息，获取进入输液配药记录中的处方ids
    '  --入参：Json_In:格式
    '  --    input
    '  --        pati_id                 N   1   病人ID
    '  --        pati_pageids            C   1   主页ID
    '  --        rcpdtl_ids              C   0   多个用逗号
    '  --        rcp_nos                 C   0   多个用逗号，rcpdtl_ids传入时，此参数无效
    '  --出参: Json_Out,格式如下
    '  --    output
    '  --        code                    N   1   应答吗：0-失败；1-成功
    '  --        message                 C   1   应答消息：失败时返回具体的错误信息
    '  --        rcpdtl_ids              C   1   处方明细id,目前传入的费用ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, 1)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", str主页IDs)
    strJson = strJson & "," & GetJsonNodeString("rcpdtl_ids", str费用Ids_in)
    strJson = strJson & "," & GetJsonNodeString("rcp_nos", strNOs)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    '先读取药品
    If objServiceCall.CallService("zl_PivasSvr_Getinfusion_Record", strJson, , "", mlngModule) = False Then Exit Function
    str费用IDs_out = Nvl(objServiceCall.GetJsonNodeValue("output.rcpdtl_ids"))
    zlPivasSvr_Getinfusion_Record = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPivasSvr_Isexsitinfusion(ByVal str费用IDs As String, Optional blnIsExist_Out As Boolean, _
    Optional lngModule As Long, Optional lng医嘱ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查药品是否进行输液配药记录
    '入参: str费用Ids_in-按费用ID查
    '      lng医嘱ID-传入时，按医嘱Id判断
    '出参:blnIsExist_Out-是否存在的，返回true,否则返回False
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intExists As Integer, objServiceCall As Object
    Dim strJson As String, strJson_out As String
        
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input               " {
    '    ""input"": {
    '     "rcpdtl_ids":"2,3,4"}
    '     ,"advice_id":0
    '    }"
    '    rcpdtl_ids  C   1   处方明细ids,目前传入的费用IDs,多个用逗号分离
    '    出参 json
    '    output "{"
    '    ""output"": {
    '     ""code"": 1,
    '     ""message"": """",
    '      ""isexist"":1
    '    }}"
    '    code    N   1   应答吗：0-失败；1-成功
    '    message C   1   应答消息：失败时返回具体的错误信息
    '    isexist N   1   1-存在;0-不存在

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("rcpdtl_ids", str费用IDs)
    strJson = strJson & "," & GetJsonNodeString("advice_id", lng医嘱ID)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"

    '先读取药品
    If objServiceCall.CallService("zl_PivasSvr_Isexsitinfusion", strJson, strJson_out, "", mlngModule) = False Then Exit Function
    
    intExists = Val(Nvl(objServiceCall.GetJsonNodeValue("output.isexist")))
    blnIsExist_Out = intExists = 1
    
    zlPivasSvr_Isexsitinfusion = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlStuffSvr_AutoReturnStuff(ByVal strAutoStuffDatas As String, Optional ByRef strErrMsg_Out As String, _
    Optional ByVal strTittle As String, Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动退料操作
    '入参: strAutoStuffDatas-自动退料数据,格式:费用ID1:数量,费用ID2:数量2,...
    '出参:strErrMsg_out-失败时，返回错误信息
    '返回:退料成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-18 21:27:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
        
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    strErrMsg_Out = ""
    '无自动退料数据，直接返回true
    If strAutoStuffDatas = "" Then ZlStuffSvr_AutoReturnStuff = True: Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '  input
    '    audit_operator        C 1 审核人
    '    stuffdtl_ids           单据明细id,目前传入的费用ID,包括退料数量(冒号加逗号组合)，数量为空表示全部退料 ：费用id1:数量1,费用id2:数量2...
    '  output
    '    code                 N 1 应答吗：0-失败；1-成功
    '    message              C 1 应答消息：失败时返回具体的错误信息

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("stuffdtl_ids", strAutoStuffDatas)
    strJson = strJson & "," & GetJsonNodeString("audit_operator", UserInfo.姓名)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"


    '先读取药品
    If objServiceCall.CallService("Zl_Stuffsvr_Autoreturnstuff", strJson, , strTittle, lngModule, False) = False Then Exit Function
    If Val(Nvl(objServiceCall.GetJsonNodeValue("output.code"), "")) = 0 Then
        strErrMsg_Out = Nvl(objServiceCall.GetJsonNodeValue("output.message"))
        Exit Function
    End If
    ZlStuffSvr_AutoReturnStuff = True
    Exit Function
errHandle:
    strErrMsg_Out = Err.Description
End Function

Public Function zlStuffsvr_GetBakStockBatch(ByVal str卫材IDs As String, ByVal lng库房ID As Long, _
                        ByRef clldata As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取备货卫材的库存信息
    '入参:lng药品ID-药品ID
    '       lng库房ID:虚拟库房id
    '出参:cllData-返回集合:形式:药品ID:药房ID:库存,....
    '返回:获取成功，返回true,否则返回False
    '编制:焦博
    '日期:2019-08-21 22:03:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo errHandle
    
    Set clldata = New Collection
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
'    ---------------------------------------------------------------------------
'    --功能：批量获取多个卫材库存及价格信息:在项目选择器中展示库存及价格信息
'    --入参
'    -- input
'    --     stuff_ids                C   1   卫材ID，多个用英文的逗号分隔,不能超过32767
'    --     warehouse_id        C   0    虚拟库房ID
'    --出参 json
'    -- output
'    --     code                           N   1   应答码：0-失败；1-成功
'    --     message                     C   1   应答消息：失败时返回具体的错误信息
'    --     item_list
'    --       stuff_id                     N   1   卫材ID
'    --         stuff_list[]
'    --             stock                   N   1   可用数量
'    --             batch                  N   1   批次(备货卫材时才返回此项)
'    --             batch_number     N   1   批号(备货卫材时才返回此项)
'    --             barcode_goods    C   1   商品条码备货卫材时才返回此项)
'    --             barcode_inside     C   1   内部条码((备货卫材时才返回此项)
'    --             provider               C   1   供应商
'    ---------------------------------------------------------------------------

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("stuff_ids", str卫材IDs)
    strJson = strJson & "," & GetJsonNodeString("warehouse_id", lng库房ID)
    strJson = "{""input"":{" & strJson & "}}"

    If objServiceCall.CallService("Zl_Stuffsvr_Getbakstockbatch", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list", "stuff_id,batch")
    If clldata Is Nothing Then Exit Function
    zlStuffsvr_GetBakStockBatch = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_AdjustFeeData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal bytBillType As Byte, _
    cllUpdataDatas As Collection, Optional ByVal lngModule As Long, _
    Optional ByVal blnShowErrMsg As Boolean = True, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊费用转住院后调整关联费用信息
    '入参:
    '   bytBillType 单据类型：1-收费单;2-记帐单
    '   cllUpdataDatas(collect)-数据集，格式如下
    '     |-cllUpdataData(_新费用ID)(collect)每行明细数据集，成员：(原单据号,原费用ID,新单据号,新费用ID)
    '   blnShowErrMsg 是否显示错误信息
    '出参:
    '   strErrMsg blnShowErrMsg=False时返回错误信息
    '返回:成功，返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, i As Integer
    Dim strJson As String
    Dim strJsonList As String, cllUpdataData As Collection
    
    On Error GoTo ErrHandler
    If cllUpdataDatas Is Nothing Then Exit Function
    If cllUpdataDatas.Count = 0 Then Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If lngModule <> 0 Then mlngModule = lngModule
    
    For i = 1 To cllUpdataDatas.Count
        Set cllUpdataData = cllUpdataDatas(i)
        'Zl_Stuffsvr_Adjustdata
        '  --功能：门诊费用转住院时调整卫材关联数据
        '  --入参：Json_In:格式
        '  --  input
        '  --    pati_id          N  1 病人ID
        '  --    pati_pageid      N  1 主页ID
        '  --    billtype           N  1 单据类型：1-收费单;2-记帐单
        '  --    item_list
        '  --      stuff_no_old       C  1 原单据号
        '  --      stuffdtl_id_old    N  1 原处方明细ID
        '  --      stuff_no_new       C  1 新单据号
        '  --      stuffdtl_id_new    N  1 新处方明细ID
        '  --出参: Json_Out,格式如下
        '  --  output
        '  --    code              N   1   应答码：0-失败；1-成功
        '  --    message           C   1   应答消息：失败时返回具体的错误信息
        strJson = ""
        strJson = strJson & "" & GetJsonNodeString("stuff_no_old", cllUpdataData("原单据号"), Json_Text)
        strJson = strJson & "," & GetJsonNodeString("stuffdtl_id_old", cllUpdataData("原费用ID"), Json_num)
        strJson = strJson & "," & GetJsonNodeString("stuff_no_new", cllUpdataData("新单据号"), Json_Text)
        strJson = strJson & "," & GetJsonNodeString("stuffdtl_id_new", cllUpdataData("新费用ID"), Json_num)
        strJsonList = strJsonList & ",{" & strJson & "}"
    Next
    If strJsonList <> "" Then strJsonList = Mid(strJsonList, 2)
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("billtype", bytBillType, Json_num)
    strJson = strJson & ",""item_list"":[" & strJsonList & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Stuffsvr_Adjustdata", strJson, , "", lngModule, blnShowErrMsg) = False Then Exit Function
    If blnShowErrMsg = False Then
        If Val(objServiceCall.GetJsonNodeValue("output.code")) = 0 Then
            strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
            Exit Function
        End If
    End If
    zlStuffSvr_AdjustFeeData = True
    Exit Function
ErrHandler:
    If blnShowErrMsg = False Then strErrMsg = Err.Description: Exit Function
    
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugsvr_AdjustFeeData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal bytBillType As Byte, _
    cllUpdataDatas As Collection, Optional ByVal lngModule As Long, _
    Optional ByVal blnShowErrMsg As Boolean = True, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊费用转住院后调整关联费用信息
    '入参:
    '   bytBillType 单据类型：1-收费单;2-记帐单
    '   cllUpdataDatas(collect)-数据集，格式如下
    '     |-cllUpdataData(_新费用ID)(collect)每行明细数据集，成员：(原单据号,原费用ID,新单据号,新费用ID)
    '   blnShowErrMsg 是否显示错误信息
    '出参:
    '   strErrMsg blnShowErrMsg=False时返回错误信息
    '返回:成功，返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, i As Integer
    Dim strJson As String
    Dim strJsonList As String, cllUpdataData As Collection
    
    On Error GoTo ErrHandler
    If cllUpdataDatas Is Nothing Then Exit Function
    If cllUpdataDatas.Count = 0 Then Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    If lngModule <> 0 Then mlngModule = lngModule
    
    For i = 1 To cllUpdataDatas.Count
        Set cllUpdataData = cllUpdataDatas(i)
        'Zl_Drugsvr_Adjustdata
        '  --功能：门诊费用转住院时调整药品关联数据
        '  --入参：Json_In:格式
        '  --  input
        '  --    pati_id          N  1 病人ID
        '  --    pati_pageid      N  1 主页ID
        '  --    billtype         N  1 单据类型：1-收费单;2-记帐单
        '  --    item_list
        '  --      rcp_no_old       C  1 原单据号
        '  --      rcpdtl_id_old    N  1 原处方明细ID(目前传入的是费用id)
        '  --      rcp_no_new       C  1 新单据号
        '  --      rcpdtl_id_new    N  1 新处方明细ID(目前传入的是费用id)
        '  --出参: Json_Out,格式如下
        '  --  output
        '  --    code              N   1   应答码：0-失败；1-成功
        '  --    message           C   1   应答消息：失败时返回具体的错误信息
        strJson = ""
        strJson = strJson & "" & GetJsonNodeString("rcp_no_old", cllUpdataData("原单据号"), Json_Text)
        strJson = strJson & "," & GetJsonNodeString("rcpdtl_id_old", cllUpdataData("原费用ID"), Json_num)
        strJson = strJson & "," & GetJsonNodeString("rcp_no_new", cllUpdataData("新单据号"), Json_Text)
        strJson = strJson & "," & GetJsonNodeString("rcpdtl_id_new", cllUpdataData("新费用ID"), Json_num)
        strJsonList = strJsonList & ",{" & strJson & "}"
    Next
    If strJsonList <> "" Then strJsonList = Mid(strJsonList, 2)
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("billtype", bytBillType, Json_num)
    strJson = strJson & ",""item_list"":[" & strJsonList & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Drugsvr_Adjustdata", strJson, , "", lngModule, blnShowErrMsg) = False Then Exit Function
    If blnShowErrMsg = False Then
        If Val(objServiceCall.GetJsonNodeValue("output.code")) = 0 Then
            strErrMsg = objServiceCall.GetJsonNodeValue("output.message")
            Exit Function
        End If
    End If
    
    zlDrugsvr_AdjustFeeData = True
    Exit Function
ErrHandler:
    If blnShowErrMsg = False Then strErrMsg = Err.Description: Exit Function
    
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlStuffsvr_Getvirwarehouse(ByRef clldata As Collection, Optional ByRef strErrMsg_Out As String, _
                        Optional lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取虚拟库房
    '出参:cllData-返回集合
    '                   形式:科室ID:库房ID:虚拟库房id,....
    '返回:获取虚拟库房成功返回true,否则返回False
    '编制:焦博
    '日期:2019-10-17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object

    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    Set clldata = New Collection
    If GetServiceCall(objServiceCall) = False Then Exit Function

    '先读取药品
    If objServiceCall.CallService("Zl_Stuffsvr_Getvirwarehouse", "", , "", lngModule, False) = False Then Exit Function
    If Val(Nvl(objServiceCall.GetJsonNodeValue("output.code"), "")) = 0 Then
        strErrMsg_Out = Nvl(objServiceCall.GetJsonNodeValue("output.message"), "")
        Exit Function
    End If
    Set clldata = objServiceCall.GetJsonListValue("output.item_list", "dept_id,warehouse_id")
    If clldata Is Nothing Then Exit Function

    ZlStuffsvr_Getvirwarehouse = True
    Exit Function
errHandle:
    strErrMsg_Out = Err.Description
End Function

Public Function ZlPatisvr_UpdateNewPati(ByVal lng病人ID As Long, ByVal str原姓名 As String, _
    ByVal str姓名 As String, ByVal str性别 As String, ByVal str年龄 As String, _
    ByVal str医疗付款方式名称 As String, ByVal str费别 As String, ByVal bln更新费别 As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新病人信息
    '入参:
    '   lng病人ID-需要更新的病人ID
    '   str原姓名-原始病人姓名
    '   str姓名、str性别、str年龄、str医疗付款方式名称、str费别-需要更新的病人基本信息
    '出参:
    '返回:更新成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim blnNewPati As Boolean '是否新病人
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If lng病人ID = 0 Then ZlPatisvr_UpdateNewPati = True: Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    blnNewPati = (str原姓名 = "新病人")
    'Zl_Patisvr_Updatepatiarchives
    '  --功能：修改病人档案信息
    '  --入参：Json_In:格式
    '  --    input
    '  --    oper_fun              N  1   0-要更新病人信息表 1-不更新病人信息表
    '  --    is_realname_check     N  1  是否实名检查:1-实名检查;0-不检查
    '  --    pati_id               N  1  病人id:更新条件
    '  --    pati_pageid           N  1  主页ID
    '  --    pati_name_old         N     病人姓名(未修改前的姓名):比如：新病人
    '  --    pati_name             N  1  病人姓名
    '  --    pati_sex              C  1  性别
    '  --    pati_age              C  1  年龄
    '  --    pati_type             C  1  病人类型(普通，医保，留观)
    '  --    pati_birthdate        C  1  出生日期:yyyy-mm-dd hh24:mi:ss
    '  --    phone_number          C  1  手机号
    '  --    insurance_num         C  1  医保号
    '  --    pati_idcard           C  1  身份证号
    '  --    outpatient_num        C  1  门诊号
    '  --    fee_category          C  1  费别
    '  --    mdlpay_mode_name      C  1  医疗付款方式名称
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                  N 1   应答码：0-失败；1-成功
    '  --    message               C 1   应答消息：失败时返回具体的错误信息
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("is_realname_check", 0, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    If blnNewPati Then
        strJson = strJson & "," & GetJsonNodeString("pati_name_old", str原姓名, Json_Text)
        strJson = strJson & "," & GetJsonNodeString("pati_name", str姓名, Json_Text)
        strJson = strJson & "," & GetJsonNodeString("pati_sex", str性别, Json_Text)
        strJson = strJson & "," & GetJsonNodeString("pati_age", str年龄, Json_Text)
    End If
    If bln更新费别 Then strJson = strJson & "," & GetJsonNodeString("fee_category", str费别, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("mdlpay_mode_name", str医疗付款方式名称, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Patisvr_Updatepatiarchives", strJson, , "", mlngModule) = False Then Exit Function
    ZlPatisvr_UpdateNewPati = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetAdviceInfo(ByVal cllQueryCon As Collection, ByRef rsAdviceData As ADODB.Recordset, _
    Optional ByVal byt返回类型 As Byte, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱ID获取医嘱信息
    '入参:
    '   cllQueryCon=查询条件，Array(Key,Value)，其中:Key=医嘱IDS,挂号单号,病人ID,主页ID
    '   byt返回类型=0-基本数据，1-完整数据
    '出参:
    '   rsAdviceData 医嘱信息：医嘱ID,医嘱内容,挂号单号,诊疗项目ID,诊疗类别,操作类型,皮试结果,医嘱期效,
    '               【完整数据还包含:相关ID,序号,婴儿序号,医嘱状态,开嘱医生,医生嘱托,开嘱时间,开嘱科室ID,
    '                 毒理分类,紧急标志,总量,天数,单量,计算单位,执行频次,用法,执行时间方案,
    '                 开始执行时间,执行终止时间,执行科室ID,执行科室名称,执行性质,上次执行时间,执行标记,
    '                 校对护士,校对时间,停嘱医生,停嘱时间,停嘱护士,确认停嘱时间,
    '                 费用审核,审查结果,审核状态,试管编码,屏蔽打印,是否签名,报告ID,查阅状态】
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, i As Long, cllTemp As Collection
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If cllQueryCon Is Nothing Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    Set rsAdviceData = New ADODB.Recordset
    With rsAdviceData.Fields
        '医嘱ID,医嘱内容,挂号单号,诊疗项目id,诊疗类别,操作类型,皮试结果,医嘱期效,
        .Append "医嘱ID", adBigInt, 18, adFldIsNullable
        .Append "医嘱内容", adLongVarChar, 1000, adFldIsNullable
        .Append "挂号单号", adVarChar, 10, adFldIsNullable
        .Append "诊疗项目ID", adBigInt, 18, adFldIsNullable
        .Append "诊疗类别", adVarChar, 2, adFldIsNullable
        .Append "操作类型", adVarChar, 30, adFldIsNullable
        .Append "皮试结果", adVarChar, 10, adFldIsNullable
        .Append "医嘱期效", adInteger, 2, adFldIsNullable
        If byt返回类型 > 0 Then
            '相关ID,序号,婴儿序号,医嘱状态,开嘱医生,医生嘱托,开嘱时间,开嘱科室ID,
            .Append "相关ID", adBigInt, 18, adFldIsNullable
            .Append "序号", adBigInt, 18, adFldIsNullable
            .Append "婴儿序号", adBigInt, 18, adFldIsNullable
            .Append "医嘱状态", adInteger, 3, adFldIsNullable
            .Append "开嘱医生", adVarChar, 50, adFldIsNullable
            .Append "医生嘱托", adVarChar, 100, adFldIsNullable
            .Append "开嘱时间", adVarChar, 20, adFldIsNullable
            .Append "开嘱科室ID", adBigInt, 18, adFldIsNullable
            '毒理分类,紧急标志,总量,天数,单量,计算单位,执行频次,用法,执行时间方案,
            .Append "毒理分类", adVarChar, 10, adFldIsNullable
            .Append "紧急标志", adInteger, 2, adFldIsNullable
            .Append "总量", adVarChar, 200, adFldIsNullable
            .Append "天数", adInteger, 3, adFldIsNullable
            .Append "单量", adVarChar, 100, adFldIsNullable
            .Append "计算单位", adVarChar, 100, adFldIsNullable
            .Append "执行频次", adVarChar, 20, adFldIsNullable
            .Append "用法", adVarChar, 100, adFldIsNullable
            .Append "执行时间方案", adVarChar, 20, adFldIsNullable
            '开始执行时间,执行终止时间,执行科室ID,执行科室名称,执行性质,上次执行时间,执行标记,
            .Append "开始执行时间", adVarChar, 20, adFldIsNullable
            .Append "执行终止时间", adVarChar, 20, adFldIsNullable
            .Append "执行科室ID", adBigInt, 18, adFldIsNullable
            .Append "执行科室名称", adVarChar, 50, adFldIsNullable
            .Append "执行性质", adInteger, 2, adFldIsNullable
            .Append "上次执行时间", adVarChar, 20, adFldIsNullable
            .Append "执行标记", adInteger, 2, adFldIsNullable
            '校对护士,校对时间,停嘱医生,停嘱时间,停嘱护士,确认停嘱时间,
            .Append "校对护士", adVarChar, 100, adFldIsNullable
            .Append "校对时间", adVarChar, 20, adFldIsNullable
            .Append "停嘱医生", adVarChar, 100, adFldIsNullable
            .Append "停嘱时间", adVarChar, 20, adFldIsNullable
            .Append "停嘱护士", adVarChar, 100, adFldIsNullable
            .Append "确认停嘱时间", adVarChar, 20, adFldIsNullable
            '费用审核,审查结果,审核状态,试管编码,屏蔽打印,是否签名,报告ID,查阅状态
            .Append "费用审核", adInteger, 2, adFldIsNullable
            .Append "审查结果", adInteger, 2, adFldIsNullable
            .Append "审核状态", adInteger, 2, adFldIsNullable
            .Append "试管编码", adVarChar, 10, adFldIsNullable
            .Append "屏蔽打印", adInteger, 2, adFldIsNullable
            .Append "是否签名", adInteger, 2, adFldIsNullable
            .Append "报告ID", adBigInt, 18, adFldIsNullable
            .Append "查阅状态", adInteger, 2, adFldIsNullable
        End If
    End With
    rsAdviceData.CursorLocation = adUseClient
    rsAdviceData.LockType = adLockOptimistic
    rsAdviceData.CursorType = adOpenStatic
    rsAdviceData.Open
    
    'Zl_Cissvr_Getadviceinfo
    '  --功能:获取医嘱信息
    '  --入参：Json_In:格式
    '  -- input
    '  --   query_type                   N 0 查询类型：0:查询基本信息；1:查询基本信息+扩展信息
    '  --   advice_ids                   C 0 多个医嘱ID，可能是药嘱，也可能是主医嘱（给药途径）,用“,”分隔
    '  --   rgst_no                      C 0 挂号单号:挂号单或病人ID或医嘱ID必传其中一个条件
    '  --   pati_id                      N 0 病人ID:挂号单或病人ID或医嘱ID必传其中一个条件
    '  --   pati_pageid                  N 0 主页Id
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", byt返回类型, Json_num)
    For i = 1 To cllQueryCon.Count
        Select Case cllQueryCon(i)(0)
        Case "医嘱IDS"
            strJson = strJson & "," & GetJsonNodeString("advice_ids", cllQueryCon(i)(1), Json_Text)
        Case "挂号单号"
            strJson = strJson & "," & GetJsonNodeString("rgst_no", cllQueryCon(i)(1), Json_Text)
        Case "病人ID"
            strJson = strJson & "," & GetJsonNodeString("pati_id", cllQueryCon(i)(1), Json_num)
        Case "主页ID"
            strJson = strJson & "," & GetJsonNodeString("pati_pageid", cllQueryCon(i)(1), Json_num)
        End Select
    Next
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getadviceinfo", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.advice_list")
    If clldata Is Nothing Then Exit Function
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                        C 1 应答码：0-失败；1-成功
    '  --    message                     C 1 应答消息：失败时返回具体的错误信息
    '  --    advice_list      [数组]每个医嘱信息
    '  --      advice_id                 N    id
    '  --      advice_related_id         N    相关id
    '  --      pati_id                   N    病人id
    '  --      pati_pageid               N    主页id
    '  --      pati_source               N    病人来源
    '  --      advice_statu              N    医嘱状态:-1-未生效的暂存医嘱；1-新开；2-校对疑问；3-已校对；4-已作废；5-已重整；6-已暂停；7-已启用；8-已停止；9-已确认停止
    '  --      serial_num                N    序号
    '  --      advice_day                N    天数
    '  --      advice_dosage             N    单次用量
    '  --      oper_type                 C    操作类型:诊疗项目目录.操作类型
    '  --      clinic_type               C    诊疗类别
    '  --      advice_exe_properties     N    执行性质
    '  --      advice_exe_sign           N    执行标记
    '  --      effective_time            N    医嘱期效
    '  --      advice_record_time        D    开嘱时间
    '  --      advice_doctor             C    开嘱医生
    '  --      advice_purpose            C    用药目的
    '  --      advice_reason             C    用药理由
    '  --      advice_taboonote          C    禁忌药品说明
    '  --      advice_doctor_note        C    医生嘱托
    '  --      rcpdtl_excs_desc          C    超量说明
    '  --      advice_audit_result       N    审查结果
    '  --      advice_audit_sign         N    药师审核标志
    '  --      advice_audit_time         D    药师审核时间
    '  --      advice_interval_unit      C    间隔单位
    '  --      advice_frequency          C    执行频次
    '  --      advice_frequency_times    N    频率次数
    '  --      advice_frequency_interval N    频率间隔
    '  --      advice_exetime_plane      C    执行时间方案
    '  --      advice_begintime          D    开始执行时间
    '  --      advice_endtime            D    执行终止时间
    '  --      rgst_no                       C   挂号单号
    '  --      advice_receipt_name       C   配方名称
    '  --      advice_receipt_issecret   N   是否保密
    '  --      advice_note               C   医嘱内容
    '  --      advice_cisitem_id         N   诊疗项目id
    '  --      advice_item_id            N   收费细目id
    '  --      pati_deptid               N   病人科室id
    '  --      pati_name                 C   姓名
    '  --      pati_sex                  C   性别
    '  --      pati_age                  C   年龄
    '  --      advice_audit_reason       C   药师审核原因
    '  --      skintest_info             C   皮试结果
    '
    '  --      total_qunt                N    总给予量
    '  --      Total_qunt_unit           C    总量:代单位的总量(总给予量+单位)
    '  --      single                    C    单量:单次用量+单位
    '  --      toxicity_type             C    毒理分类:药品有效，药品特性.毒理分类
    '  --      advice_dept_id            N    开嘱科室ID
    '  --      advice_stop_doctor        C    停嘱医生
    '  --      advice_stop_nurse         C    停嘱护士
    '  --      advice_stoptime           C    停嘱时间,格式：yyyy-mm-dd hh24:mi:ss
    '  --      advice_stoptime_confirm   C    确认停嘱时间:格式:yyyy-mm-dd hh24:mi:ss
    '  --      order_chk_nurse           C    校对护士
    '  --      order_chk_time            C    校对时间:yyyy-mm-dd hh24:mi;ss
    '  --      lastexe_time              D    上次执行时间
    '  --      usage                     C    用法
    '  --      emergency_tag             N    紧急标志:0-普通;1-紧急;2-补录(对门诊无效)
    '  --      is_charge_verfy           N   是否费用审核:1-审核;0-未审核
    '  --      baby_num                  N   婴儿序号
    '  --      valuation_nature          N   计价特性:0-正常计价；1-不计价；2-手工计价
    '  --      advice_exedept_id         N   执行科室ID:
    '  --      advice_exedept_name       C   执行科室名称
    '  --      testtube_code             C   试管编码:诊疗项目目录.试管编码
    '  --      hide_print                N   屏蔽打印
    '  --      prerequisite_id           N   前提ID
    '  --      is_staff_sig              N   是否签名:1-签名;0-未签名
    '  --      rpt_id                    N   病历报告id
    '  --      consult_statu             N   查阅状态:0-未查阅,1-已查阅
    '  --      advice_verfy_statu        N   审核状态:Null-无需审核，1-待审核，2-审核通过，3-审核未通过，4－血库待接收，5－血库配血中，6-血库停止配血，7-输血初审核待签发
    '  --      apply_num                 N   申请序号
    '  --      cisitem_unit              C   诊疗项目计算单位
    '  --      pati_wardarea_id          N   当前病区ID
    With rsAdviceData
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            .AddNew
            '医嘱ID,医嘱内容,挂号单号,诊疗项目id,诊疗类别,操作类型,皮试结果,医嘱期效,
            !医嘱id = cllTemp("_advice_id")
            !医嘱内容 = cllTemp("_advice_note")
            !挂号单号 = cllTemp("_rgst_no")
            !诊疗项目ID = cllTemp("_advice_cisitem_id")
            !诊疗类别 = cllTemp("_clinic_type")
            !操作类型 = cllTemp("_oper_type")
            !皮试结果 = cllTemp("_skintest_info")
            !医嘱期效 = cllTemp("_effective_time")
            If byt返回类型 > 0 Then
                '相关ID,序号,婴儿序号,医嘱状态,开嘱医生,医生嘱托,开嘱时间,开嘱科室ID,
                !相关ID = cllTemp("_advice_related_id")
                !序号 = cllTemp("_serial_num")
                !婴儿序号 = cllTemp("_baby_num")
                !医嘱状态 = cllTemp("_advice_statu")
                !开嘱医生 = cllTemp("_advice_doctor")
                !医生嘱托 = cllTemp("_advice_doctor_note")
                !开嘱时间 = cllTemp("_advice_record_time")
                !开嘱科室id = cllTemp("_advice_dept_id")
                '毒理分类,紧急标志,总量,天数,单量,计算单位,执行频次,用法,执行时间方案,
                !毒理分类 = cllTemp("_toxicity_type")
                !紧急标志 = cllTemp("_emergency_tag")
                !总量 = cllTemp("_Total_qunt_unit")
                !天数 = cllTemp("_advice_day")
                !单量 = cllTemp("_single")
                !计算单位 = cllTemp("_cisitem_unit")
                !执行频次 = cllTemp("_advice_frequency")
                !用法 = cllTemp("_usage")
                !执行时间方案 = cllTemp("_advice_exetime_plane")
                '开始执行时间,执行终止时间,执行科室ID,执行科室名称,执行性质,上次执行时间,执行标记,
                !开始执行时间 = cllTemp("_advice_begintime")
                !执行终止时间 = cllTemp("_advice_endtime")
                !执行科室ID = cllTemp("_advice_exedept_id")
                !执行科室名称 = cllTemp("_advice_exedept_name")
                !执行性质 = cllTemp("_advice_exe_properties")
                !上次执行时间 = cllTemp("_lastexe_time")
                !执行标记 = cllTemp("_advice_exe_sign")
                '校对护士,校对时间,停嘱医生,停嘱时间,停嘱护士,确认停嘱时间,
                !校对护士 = cllTemp("_order_chk_nurse")
                !校对时间 = cllTemp("_order_chk_time")
                !停嘱医生 = cllTemp("_advice_stop_doctor")
                !停嘱时间 = cllTemp("_advice_stoptime")
                !停嘱护士 = cllTemp("_advice_stop_nurse")
                !确认停嘱时间 = cllTemp("_advice_stoptime_confirm")
                '费用审核,审查结果,审核状态,试管编码,屏蔽打印,是否签名,报告ID,查阅状态
                !费用审核 = cllTemp("_is_charge_verfy")
                !审查结果 = cllTemp("_advice_audit_result")
                !审核状态 = cllTemp("_advice_verfy_statu")
                !试管编码 = cllTemp("_testtube_code")
                !屏蔽打印 = cllTemp("_hide_print")
                !是否签名 = cllTemp("_is_staff_sig")
                !报告ID = cllTemp("_rpt_id")
                !查阅状态 = cllTemp("_consult_statu")
            End If
            .Update
        Next
        If .RecordCount > 0 Then .MoveFirst
    End With
    ZlCissvr_GetAdviceInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetAdviceSendInfo(ByVal byt查询类型 As Byte, ByVal strValues As String, _
    ByRef rsAdviceSendData As ADODB.Recordset, Optional ByVal bln包含相关医嘱 As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱发送信息
    '入参:
    '   byt查询类型=0-按医嘱id串（advice_ids）查询;1-按医嘱ID+医嘱发送号查询;2-按医嘱ID+记录性质+NO查询;3-仅按医嘱发送号查询
    '   strValues=查询值，具体如下:
    '               byt查询类型=0:医嘱ID串,格式：医嘱ID,医嘱ID,...
    '               byt查询类型=1:单据信息,格式：医嘱ID:NO:记录性质,医嘱ID:NO:记录性质,...
    '               byt查询类型=2:医嘱发送信息,格式:医嘱ID:发送号,医嘱ID:发送号,...
    '               byt查询类型=3:医嘱发送号,格式:发送号,发送号,...
    '   bln包含相关医嘱=是否包含相关医嘱ID
    '出参:
    '   rsAdviceSendData 医嘱发送信息：医嘱ID,发送号,挂号单号,病人ID,主页ID,病人科室ID,开嘱科室ID,
    '                                 病人来源,诊疗类别,计价性质,相关id,门诊记帐,医嘱内容,样本条码,
    '                                 No,记录性质,首次时间,末次时间,执行状态,发送时间,医嘱期效
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, strJsonList As String, strJsonTemp As String
    Dim clldata As Collection, i As Long, cllTemp As Collection
    Dim varNos As Variant, varNo As Variant
    
    On Error GoTo ErrHandler
    If strValues = "" Then Exit Function
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    Set rsAdviceSendData = New ADODB.Recordset
    With rsAdviceSendData.Fields
        .Append "医嘱ID", adBigInt, 18, adFldIsNullable
        .Append "发送号", adBigInt, 18, adFldIsNullable
        .Append "挂号单号", adVarChar, 10, adFldIsNullable
        .Append "病人ID", adBigInt, 18, adFldIsNullable
        .Append "主页ID", adBigInt, 18, adFldIsNullable
        .Append "病人科室ID", adBigInt, 18, adFldIsNullable
        .Append "开嘱科室ID", adBigInt, 18, adFldIsNullable
        .Append "病人来源", adInteger, 2, adFldIsNullable
        .Append "诊疗类别", adVarChar, 5, adFldIsNullable
        .Append "计价性质", adInteger, 2, adFldIsNullable
        .Append "相关id", adBigInt, 18, adFldIsNullable
        .Append "门诊记帐", adInteger, 2, adFldIsNullable
        .Append "医嘱内容", adLongVarChar, 1000, adFldIsNullable
        .Append "样本条码", adVarChar, 20, adFldIsNullable
        .Append "No", adVarChar, 20, adFldIsNullable
        .Append "记录性质", adInteger, 2, adFldIsNullable
        .Append "首次时间", adVarChar, 20, adFldIsNullable
        .Append "末次时间", adVarChar, 20, adFldIsNullable
        .Append "执行状态", adInteger, 2, adFldIsNullable
        .Append "发送时间", adVarChar, 20, adFldIsNullable
        .Append "医嘱期效", adInteger, 2, adFldIsNullable
    End With
    rsAdviceSendData.CursorLocation = adUseClient
    rsAdviceSendData.LockType = adLockOptimistic
    rsAdviceSendData.CursorType = adOpenStatic
    rsAdviceSendData.Open
    
    'Zl_Cissvr_Getadvicesendinfo
    '  --功能：获取医嘱发送信息
    '  --入参      json
    '  --input     根据医嘱ID查询医嘱发送信息
    '  --  query_type                N 1 查询类型：0-按医嘱id串（advice_ids）查询;1-按医嘱ID+医嘱发送号查询;2-按医嘱ID+记录性质+NO查询;3-仅按医嘱发送号查询
    '  --  return_type               N 1 返回类型：0-返回基本信息；1-返回基本信息+扩展信息
    '  --  contain_related           N 1 是否包含相关医嘱记录
    '  --  advice_ids                C 1 医嘱ID串，支持多个医嘱ID，用“，”分隔
    '  --  send_nos                  C 1 医嘱发送号串，支持多个，用“，”分隔,仅查询类型为3时有效
    '  --  item_list[]
    '  --    advice_id               N 1 医嘱ID
    '  --    send_no                 N 1 医嘱发送号
    '  --    bill_no                 C 1 NO
    '  --    bill_prop               N 1 记录性质
    '  --出参      json
    '  --output
    '  --  code                      C 1 应答码：0-失败；1-成功
    '  --  message                   C 1 应答消息：失败时返回具体的错误信息
    '  --  advice_send_list[]
    '  --    advice_id               N   医嘱id
    '  --    send_no                 N   发送号
    '  --    register_no             C   挂号单号
    '  --    pati_id                 N   病人id
    '  --    pati_pageid             N   主页id
    '  --    pati_deptid             N   病人科室id
    '  --    advice_dept_id          N   开嘱科室ID
    '  --    pati_source             N   病人来源
    '  --    clinic_type             C   诊疗类别
    '  --    valuation_nature        N   计价性质:0-正常计价；1-不计价；2-手工计价
    '  --    advice_related_id       N   相关id
    '  --    outpati_account         N   是否门诊记帐
    '  --    advice_note             C   医嘱内容
    '  --    sample_barcode          C   样本条码
    '  --    bill_no                 C   No
    '  --    bill_prop               N   记录性质
    '  --    advice_send_firsttime   D   首次时间
    '  --    advice_send_endtime     D   末次时间
    '  --    advice_send_exestatus   N   执行状态
    '  --    advice_send_sendtime    D   发送时间
    '  --    effective_time          N   医嘱期效
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", byt查询类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("return_type", 1, Json_num)
    strJson = strJson & "," & GetJsonNodeString("contain_related", IIf(bln包含相关医嘱, 1, 0), Json_num)
    Select Case byt查询类型
    Case 0
        strJson = strJson & "," & GetJsonNodeString("advice_ids", strValues, Json_Text)
    Case 1
        strJsonList = ""
        varNos = Split(strValues, ",")
        For i = 0 To UBound(varNos)
            varNo = Split(varNos(i), ":")
            strJsonTemp = ""
            strJsonTemp = strJsonTemp & "" & GetJsonNodeString("advice_id", varNo(0), Json_num)
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("send_no", varNo(1), Json_num)
            strJsonList = strJsonList & ",{" & strJsonTemp & "}"
        Next
        strJson = strJson & ",""item_list"":[" & Mid(strJsonList, 2) & "]"
    Case 2
        strJsonList = ""
        varNos = Split(strValues, ",")
        For i = 0 To UBound(varNos)
            varNo = Split(varNos(i), ":")
            strJsonTemp = ""
            strJsonTemp = strJsonTemp & "" & GetJsonNodeString("advice_id", varNo(0), Json_num)
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("bill_no", varNo(1), Json_Text)
            strJsonTemp = strJsonTemp & "," & GetJsonNodeString("bill_prop", varNo(2), Json_num)
            strJsonList = strJsonList & ",{" & strJsonTemp & "}"
        Next
        strJson = strJson & ",""item_list"":[" & Mid(strJsonList, 2) & "]"
    Case 3
        strJson = strJson & "," & GetJsonNodeString("send_nos", strValues, Json_Text)
    Case Else
        Exit Function
    End Select
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getadvicesendinfo", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.advice_send_list")
    If clldata Is Nothing Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        With rsAdviceSendData
            .AddNew
            !医嘱id = cllTemp("_advice_id")
            !发送号 = cllTemp("_send_no")
            !挂号单号 = cllTemp("_register_no")
            !病人ID = cllTemp("_pati_id")
            !主页ID = cllTemp("_pati_pageid")
            !病人科室ID = cllTemp("_pati_deptid")
            !开嘱科室id = cllTemp("_advice_dept_id")
            !病人来源 = cllTemp("_pati_source")
            !诊疗类别 = cllTemp("_clinic_type")
            !计价性质 = cllTemp("_valuation_nature")
            !相关ID = cllTemp("_advice_related_id")
            !门诊记帐 = cllTemp("_outpati_account")
            !医嘱内容 = cllTemp("_advice_note")
            !样本条码 = cllTemp("_sample_barcode")
            !NO = cllTemp("_bill_no")
            !记录性质 = cllTemp("_bill_prop")
            !首次时间 = cllTemp("_advice_send_firsttime")
            !末次时间 = cllTemp("_advice_send_endtime")
            !执行状态 = cllTemp("_advice_send_exestatus")
            !发送时间 = cllTemp("_advice_send_sendtime")
            !医嘱期效 = cllTemp("_effective_time")
            .Update
        End With
    Next
    If rsAdviceSendData.RecordCount > 0 Then rsAdviceSendData.MoveFirst
    
    ZlCissvr_GetAdviceSendInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_AdviceExecuting(ByVal bytBillType As Byte, ByVal cllAdviceDatas As Collection, _
    ByRef blnExistExecuting As Boolean, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定的医嘱是否正在执行
    '入参:
    '   bytBillType 单据类型：1-收费单;2-记帐单;3-自动记帐单
    '   cllAdviceDatas(collect)-数据集，格式如下
    '     |-cllAdviceData(collect)每行明细数据集，成员：(医嘱ID,费用单号)
    '出参:
    '   blnExistExecuting 是否正在执行
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, strItem As String
    Dim cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    blnExistExecuting = False
    If cllAdviceDatas Is Nothing Then ZlCissvr_AdviceExecuting = True: Exit Function
    If cllAdviceDatas.Count = 0 Then ZlCissvr_AdviceExecuting = True: Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Cissvr_Adviceexecuting
    '  --功能：检查指定的医嘱是否正在执行
    '  --入参：Json_In:格式
    '  --input
    '  --    item_list
    '  --      advice_id                C  1  医嘱ID
    '  --      advice_send_no           C  1  发送单号
    '  --      advice_send_properties   N  1  记录性质
    '  --出参: Json_Out,格式如下
    '  --    output
    '  --        code                   N   1   应答吗：0-失败；1-成功
    '  --        message                C   1   应答消息：失败时返回具体的错误信息
    '  --        isexist                N   1   是否存在，1-存在;0-不存在
    For i = 1 To cllAdviceDatas.Count
        Set cllTemp = cllAdviceDatas(i)
        strItem = ""
        strItem = strItem & "" & GetJsonNodeString("advice_id", cllTemp("医嘱ID"), Json_num)
        strItem = strItem & "," & GetJsonNodeString("advice_send_no", cllTemp("费用单号"), Json_Text)
        strItem = strItem & "," & GetJsonNodeString("advice_send_properties", bytBillType, Json_num)
        strJson = strJson & ",{" & strItem & "}"
    Next
    strJson = "{""input"":{""item_list"":[" & Mid(strJson, 2) & "]}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Adviceexecuting", strJson, , "", mlngModule) = False Then Exit Function
    blnExistExecuting = Val(objServiceCall.GetJsonNodeValue("output.isexist")) = 1
    
    ZlCissvr_AdviceExecuting = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCisSvr_GetAdviceExcutNums(ByVal bytBillType As Byte, ByVal cllAdviceDatas As Collection, _
    ByRef rsExecuted As ADODB.Recordset, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱发送的已执行数量
    '入参:
    '   bytBillType 单据类型：1-收费单;2-记帐单;3-自动记帐单
    '   cllAdviceDatas(collect)-数据集，格式如下
    '     |-cllAdviceData(collect)每行明细数据集，成员：(医嘱ID,费用单号,执行部门ID(N))
    '出参:
    '   rsExecuted 已执行数据：医嘱序号,NO,收费细目ID,已执行数
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, strItem As String
    Dim clldata As Collection, cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    Set rsExecuted = New ADODB.Recordset
    With rsExecuted.Fields
        .Append "医嘱序号", adBigInt, 18, adFldIsNullable
        .Append "NO", adVarChar, 10, adFldIsNullable
        .Append "收费细目ID", adBigInt, 18, adFldIsNullable
        .Append "已执行数", adDouble, , adFldIsNullable
    End With
    rsExecuted.CursorLocation = adUseClient
    rsExecuted.LockType = adLockOptimistic
    rsExecuted.CursorType = adOpenStatic
    rsExecuted.Open
    
    If cllAdviceDatas Is Nothing Then ZlCisSvr_GetAdviceExcutNums = True: Exit Function
    If cllAdviceDatas.Count = 0 Then ZlCisSvr_GetAdviceExcutNums = True: Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Cissvr_Getadviceexcutnums
    '  --功能：获取医嘱发送的已执行数量
    '  --入参      json
    '  --input     根据医嘱ID查询医嘱发送信息
    '  --  item_list                 数据组
    '  --    advice_id               N 1 医嘱ID
    '  --    bill_no                 C 1 NO
    '  --    bill_prop               N 1 记录性质
    '  --    exe_deptid            N 0 执行部门ID
    '  --出参      json
    '  --output
    '  --  code                      C 1 应答码：0-失败；1-成功
    '  --  message                   C 1 应答消息：失败时返回具体的错误信息
    '  --  item_list                 数据组
    '  --    advice_id               N 1 医嘱ID
    '  --    bill_no                 C 1 No
    '  --    fee_item_id             N 1 收费细目ID
    '  --    execute_num             N 1 已执行数
    '  --说明：注意，即使已执行数为0也要返回
    For i = 1 To cllAdviceDatas.Count
        Set cllTemp = cllAdviceDatas(i)
        strItem = ""
        strItem = strItem & "" & GetJsonNodeString("advice_id", cllTemp("医嘱ID"), Json_num)
        strItem = strItem & "," & GetJsonNodeString("bill_no", cllTemp("费用单号"), Json_Text)
        strItem = strItem & "," & GetJsonNodeString("bill_prop", bytBillType, Json_num)
        strItem = strItem & "," & GetJsonNodeString("exe_deptid", zlGetNodeValueFromCollect(cllTemp, "执行部门ID", Json_num), Json_num)
        strJson = strJson & ",{" & strItem & "}"
    Next
    strJson = "{""input"":{""item_list"":[" & Mid(strJson, 2) & "]}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getadviceexcutnums", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        rsExecuted.AddNew
            rsExecuted!医嘱序号 = cllTemp("_advice_id")
            rsExecuted!NO = cllTemp("_bill_no")
            rsExecuted!收费细目ID = cllTemp("_fee_item_id")
            rsExecuted!已执行数 = cllTemp("_execute_num")
        rsExecuted.Update
    Next
    If rsExecuted.RecordCount > 0 Then rsExecuted.MoveFirst
    
    ZlCisSvr_GetAdviceExcutNums = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetDiaginfoFromAdviceID(ByVal strAdviceIDs As String, _
    ByRef rsDiagnose As ADODB.Recordset, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱ID获取医嘱诊断信息
    '入参:
    '   strAdviceIDs=医嘱ID，多个用逗号分隔
    '出参:
    '   rsDiagnose=诊断数据，字段:医嘱ID,诊断描述
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object, clldata As Collection
    Dim strJson As String
    Dim cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    Set rsDiagnose = New ADODB.Recordset
    With rsDiagnose.Fields
        .Append "医嘱ID", adBigInt, 18, adFldIsNullable
        .Append "诊断描述", adLongVarChar, 500, adFldIsNullable
    End With
    rsDiagnose.CursorLocation = adUseClient
    rsDiagnose.LockType = adLockOptimistic
    rsDiagnose.CursorType = adOpenStatic
    rsDiagnose.Open
    
    If strAdviceIDs = "" Then ZlCissvr_GetDiaginfoFromAdviceID = True: Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Cissvr_Getdiaginfo
    '  --功能:获取病人诊断信息
    '  --入参：Json_In:格式
    '  -- input
    '  --   advice_ids           C   医嘱ids,医嘱id拼串
    '  --   query_type           N  1 1-按指定条件查询,2-仅按病人id,主页id查询诊断
    '  --   pati_info            C   病人id等其他信息
    '  --     pati_id            N 1 病人id
    '  --     pati_pageid        N 1 主页id
    '  --     diag_types         C   诊断类型:0-所有类型,1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
    '  --                            可以为多个诊断类型，用逗号分离,如:2,12
    '  --     rec_source         N   记录来源:1-病历；2-入院登记；3-首页整理;4-病案;NULL-不作限制
    '  --     diag_num           N   诊断次序:NULL表示不作限制
    '  --     code_type          C   编码类别:ICD-11的编码编码类别为'E',空时表示读取ICD-10等
    '  --     input_num          C   录入次序:启用了ICD-11编码录入后，诊断的录入次序
    '  --     rec_sources       C   记录来源拼串
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N  1 应答码：0-失败；1-成功
    '  --    message             C  1 应答消息：失败时返回具体的错误信息
    '  --    diag_list     [数组]
    '  --      diag_type         N 1 诊断类型:1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;
    '  --                                     12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
    '  --      diag_num          N 1 诊断序号
    '  --      code_num          N 1 编码序号
    '  --      dz_id             N 1 疾病ID
    '  --      dz_code           C 1 疾病编码
    '  --      diag_note         C 1 诊断描述
    '  --      recoder           C 1 记录人
    '  --      rec_time          C 1 记录时间:yyyy-mm-dd hh24:mi:ss
    '  --      adtd_rsn          C 1 出院情况:治愈、好转、未愈、死亡、其他
    '  --      diag_id           N 1 诊断id
    '  --      diag_rec_id       N 1 诊断记录ID:病人诊断记录.ID
    '  --      diag_doubt        N 1 是否疑诊
    '  --      advice_id         N   医嘱id(根据医嘱ids查询时才返回)
    '  --      advice_main_id    N   组医嘱id(根据医嘱ids查询时才返回)
    '  --      advice_related_id N   相关id(根据医嘱ids查询时才返回)
    '  --      rec_source        N   记录来源:1-病历；2-入院登记；3-首页整理;4-病案;NULL-不作限制
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", 0, Json_num)
    strJson = strJson & "," & GetJsonNodeString("advice_ids", strAdviceIDs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getdiaginfo", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.diag_list")
    If clldata Is Nothing Then Exit Function
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        rsDiagnose.AddNew
            rsDiagnose!医嘱id = cllTemp("_advice_id")
            rsDiagnose!诊断描述 = cllTemp("_diag_note")
        rsDiagnose.Update
    Next
    If rsDiagnose.RecordCount > 0 Then rsDiagnose.MoveFirst
    ZlCissvr_GetDiaginfoFromAdviceID = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCisSvr_UpdAdviceChargeTag(ByVal bytBillType As Byte, _
    ByVal cllAdviceDatas As Collection, ByVal cllFeeDetails As Collection, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更改病人医嘱数据的计费状态及删除附费数据
    '入参:
    '   bytBillType 单据类型：1-收费单;2-记帐单;3-自动记帐单
    '   cllAdviceDatas(collect)-数据集，格式如下
    '     |-cllAdviceData(collect)每行明细数据集
    '        |-成员(医嘱ID,发送号(N),费用单号(N),计费状态,删除附费(N))
    '   cllFeeDetails(collect)费用明细集，仅门诊退费时传入，用于标记 医嘱执行计价.执行状态 为 2-已退费
    '     |-cllFeeDetail(collect)每行明细数据
    '        |-成员(医嘱ID,费用单号,收费明细ID)
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, strItem As String, strJsonList As String
    Dim cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    If cllAdviceDatas Is Nothing Then ZlCisSvr_UpdAdviceChargeTag = True: Exit Function
    If cllAdviceDatas.Count = 0 Then ZlCisSvr_UpdAdviceChargeTag = True: Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Cissvr_Updadvicechargetag
    '  --功能：更改病人医嘱数据的计费状态及删除附费数据
    '  --入参：Json_In:格式
    '  --input
    '  --  item_list
    '  --    advice_id               N   1 医嘱ID
    '  --    send_no                 N   1 发送号
    '  --    bill_no                 C   1 单据号
    '  --    bill_prop               N   1 记录性质
    '  --    del_annex               N   0 是否删除对应的医嘱附费数据:1-删除;0-不删除
    '  --    charge_status           N   1 更新的计费状态:-1-无须计费(通常无执行和院外执行的都无须计费);0-未计费;1-已计费(或记帐)，2-部分收费/退费(记帐/销帐)，3-全部收费(仅门诊收费有此项)，4-全部退费(销帐)
    '  -- fee_detail_list            费用明细，仅门诊退费时传入
    '  --    advice_id               N   1 医嘱ID
    '  --    bill_no                 C   1 单据号
    '  --    bill_prop               N   1 记录性质
    '  --    fee_item_id             N   1 收费细目ID
    '  --出参: Json_Out,格式如下
    '  --output
    '  --   code                     N   1   应答吗：0-失败；1-成功
    '  --   message                  C   1   应答消息：失败时返回具体的错误信息
    strJsonList = ""
    For i = 1 To cllAdviceDatas.Count
        Set cllTemp = cllAdviceDatas(i)
        strItem = ""
        strItem = strItem & "" & GetJsonNodeString("advice_id", cllTemp("医嘱ID"), Json_num)
        strItem = strItem & "," & GetJsonNodeString("send_no", zlGetNodeValueFromCollect(cllTemp, "发送号", "N"), Json_num)
        strItem = strItem & "," & GetJsonNodeString("bill_no", zlGetNodeValueFromCollect(cllTemp, "费用单号", "C"), Json_Text)
        strItem = strItem & "," & GetJsonNodeString("bill_prop", bytBillType, Json_num)
        strItem = strItem & "," & GetJsonNodeString("del_annex", zlGetNodeValueFromCollect(cllTemp, "删除附费", "N"), Json_num)
        strItem = strItem & "," & GetJsonNodeString("charge_status", cllTemp("计费状态"), Json_num)
        strJsonList = strJsonList & ",{" & strItem & "}"
    Next
    strJson = ""
    strJson = strJson & """item_list"":[" & Mid(strJsonList, 2) & "]"
    
    If Not cllFeeDetails Is Nothing Then
        strJsonList = ""
        For i = 1 To cllFeeDetails.Count
            Set cllTemp = cllFeeDetails(i)
            strItem = ""
            strItem = strItem & "" & GetJsonNodeString("advice_id", cllTemp("医嘱ID"), Json_num)
            strItem = strItem & "," & GetJsonNodeString("bill_no", cllTemp("费用单号"), Json_Text)
            strItem = strItem & "," & GetJsonNodeString("bill_prop", bytBillType, Json_num)
            strItem = strItem & "," & GetJsonNodeString("fee_item_id", cllTemp("收费明细ID"), Json_num)
            strJsonList = strJsonList & ",{" & strItem & "}"
        Next
        If strJsonList <> "" Then
            strJson = strJson & ",""fee_detail_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Updadvicechargetag", strJson, , , lngModule) = False Then Exit Function
    
    ZlCisSvr_UpdAdviceChargeTag = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlCisSvr_IsExistAdvice(ByVal str挂号单 As String, ByVal lng病人ID As Long, ByRef bln存在医嘱_Out As Boolean, _
    Optional blnShowErrMsg As Boolean, Optional strErrMsg_Out As String, Optional lngModule As Long, _
    Optional str主页id As String, Optional ByVal blnOnly有效医嘱 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定的挂号单或病人是否存在医嘱数据
    '入参:str挂号单-指定的挂号单
    '     lng病人ID-病人id
    '     blnShowErrMsg-是否显示错误信息
    '     str主页id-主页id=""时，表示不按主页id查找;0时表示只查主页id为零的医嘱,>0表示查询指定主页的医嘱
    '出参:bln存在医嘱_Out
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strServiceName    As String
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg_Out = "连接临床域服务失败，无法获取有效的医嘱数据!"
        If blnShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'Zl_Cissvr_Existadvice
    '  --功能:判断是否存在医嘱数据或判断指定挂号单是否已经开医嘱
    '  --入参：Json_In:格式
    '  -- input
    '  --   pati_id              N 1 病人ID
    '  --   pati_pageid          N   主页Id
    '  --   rgst_no              C 1 挂号单，多个用逗号分隔
    '  --   only_valid           N   只检查没有作废的医嘱
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               C 1 应答码：0-失败；1-成功
    '  --    message            C 1 应答消息：失败时返回具体的错误信息
    '  --    exist              N 1 是否存在，1-存在;0-不存在
    strJson = strJson & "" & GetJsonNodeString("rgst_no", str挂号单, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    If str主页id <> "" Then
        strJson = strJson & "," & GetJsonNodeString("pati_pageid", Val(str主页id), Json_num)
    End If
    strJson = strJson & "," & GetJsonNodeString("only_valid", IIf(blnOnly有效医嘱, 1, 0), Json_num)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Cissvr_Existadvice"
    If objServiceCall.CallService(strServiceName, strJson, , "", lngModule, False) = False Then Exit Function
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "不能获取指定条件下的医嘱数据，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    bln存在医嘱_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.exist"))) = 1
    zlCisSvr_IsExistAdvice = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function ZlPatiSvr_UpdateOutPatiState(ByVal cllPati As Collection, _
                Optional ByRef strErrMsg_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:挂号或者退号成功后更新就诊信息
    '入参:
    '   bytBillType 单据类型：1-收费单;2-记帐单;3-自动记帐单
    '   cllPati(collect)-成员：病人id,手机号(N),就诊诊室(N),就诊状态(N),就诊时间(N),门诊号(N)
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, varValue As Variant
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    If cllPati Is Nothing Then ZlPatiSvr_UpdateOutPatiState = True: Exit Function
    If cllPati.Count = 0 Then ZlPatiSvr_UpdateOutPatiState = True: Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Cissvr_Updadvicechargetag
    '  --功能：更改病人医嘱数据的计费状态及删除附费数据
    '  --入参：Json_In:格式
    '  --input
    '    --    pati_id            N 1 病人id
    '    --    phone_number       C   病人手机号
    '    --    fee_category       C   费别
    '    --    visit_room         C   更新的就诊诊室
    '    --    visit_status       N   更新的就诊状态
    '    --    visit_time         C   更新的就诊时间
    '    --    outpatient_num             C   门诊号
    '  --出参: Json_Out,格式如下
    '  --output
    '  --   code                     N   1   应答吗：0-失败；1-成功
    '  --   message                  C   1   应答消息：失败时返回具体的错误信息
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", cllPati("病人ID"), Json_num)
    
    If CollectionExitsValue(cllPati, "手机号") Then
        strJson = strJson & "," & GetJsonNodeString("phone_number", cllPati("手机号"), Json_Text)
    End If
    If CollectionExitsValue(cllPati, "费别") Then
        strJson = strJson & "," & GetJsonNodeString("fee_category", cllPati("费别"), Json_Text)
    End If
    If CollectionExitsValue(cllPati, "年龄") Then
        strJson = strJson & "," & GetJsonNodeString("pati_age", cllPati("年龄"), Json_Text)
    End If
    If CollectionExitsValue(cllPati, "就诊诊室") Then
        strJson = strJson & "," & GetJsonNodeString("visit_room", cllPati("就诊诊室"), Json_Text)
    End If
    If CollectionExitsValue(cllPati, "就诊状态") Then
        strJson = strJson & "," & GetJsonNodeString("visit_status", cllPati("就诊状态"), Json_num)
    End If
    If CollectionExitsValue(cllPati, "就诊时间") Then
        strJson = strJson & "," & GetJsonNodeString("visit_time", cllPati("就诊时间"), Json_Text)
    End If
    If CollectionExitsValue(cllPati, "门诊号") Then
        strJson = strJson & "," & GetJsonNodeString("outpatient_num", cllPati("门诊号"), Json_Text)
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_PatiSvr_UpdateOutPatiState", strJson, , , lngModule, False) = False Then Exit Function
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    
    ZlPatiSvr_UpdateOutPatiState = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function ZlCisSvr_RegistAffirm(Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 挂号成功后通知临床服务
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 14:48
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    On Error GoTo Errhand

    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '目前是空函数
    strJson = "{""input"":{" & strJson & "}}"
    If objServiceCall.CallService("Zl_CisSvr_RegistAffirm", strJson, , , lngModule) = False Then Exit Function
    
    ZlCisSvr_RegistAffirm = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlCISSvr_NewOutPatiVisitRec(ByVal cllRgst As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 挂号成功后生产临床就诊记录
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 14:48
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    On Error GoTo Errhand

    If lngModule <> 0 Then mlngModule = lngModule
    If cllRgst Is Nothing Then Exit Function
    If cllRgst.Count = 0 Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("rgst_no", cllRgst("挂号单号"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("rgst_type", cllRgst("单据性质"), Json_num)
    strJson = strJson & "," & GetJsonNodeString("rgst_appt_sign", cllRgst("预约标志"), Json_num)
    strJson = strJson & "," & GetJsonNodeString("rgst_code", cllRgst("号别"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("rgst_rec_id", zlGetNodeValueFromCollect(cllRgst, "出诊记录id", "N"), Json_num)
    strJson = strJson & "," & GetJsonNodeString("outptyp_name", cllRgst("号类"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_id", zlGetNodeValueFromCollect(cllRgst, "病人ID", "N"), Json_num)
    strJson = strJson & "," & GetJsonNodeString("outpno", zlGetNodeValueFromCollect(cllRgst, "门诊号", "C"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_name", cllRgst("姓名"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_sex", cllRgst("性别"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_age", cllRgst("年龄"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("revst_sign", zlGetNodeValueFromCollect(cllRgst, "复诊", "N"), Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_category", cllRgst("费别"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("emg_sign", zlGetNodeValueFromCollect(cllRgst, "急诊", "N"), Json_num)
    strJson = strJson & "," & GetJsonNodeString("outproom_name", zlGetNodeValueFromCollect(cllRgst, "诊室", "C"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("exe_deptid", cllRgst("执行部门ID"), Json_num)
    strJson = strJson & "," & GetJsonNodeString("rgst_exetr", zlGetNodeValueFromCollect(cllRgst, "执行人", "C"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("happen_time", cllRgst("发生时间"), Json_Text)
    strJson = strJson & "," & GetJsonNodeString("close_account_type", zlGetNodeValueFromCollect(cllRgst, "结算模式", "N"), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_CISSvr_NewOutPatiVisitRec", strJson, , , lngModule) = False Then Exit Function
    
    zlCISSvr_NewOutPatiVisitRec = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlCISSvr_DelOutPatiVisitRec(ByVal strNo As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 退号成功后作废临床就诊记录
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 14:48
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    On Error GoTo Errhand

    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("rgst_no", strNo, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_CISSvr_DelOutPatiVisitRec", strJson, , , lngModule) = False Then Exit Function
    
    zlCISSvr_DelOutPatiVisitRec = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlCISSvr_UpdateOutPatiVisitRec(ByVal cllVisit As Collection, _
                Optional ByVal blnNoShowErr As Boolean, Optional ByRef strErrMsg_Out As String, _
                Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 修改临床就诊信息
    ' 入参 : cllVisit(Collect):就诊信息
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 14:48
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim intReturn As Integer, i As Integer
    Dim strJson As String
    On Error GoTo Errhand

    If lngModule <> 0 Then mlngModule = lngModule
    If cllVisit Is Nothing Then Exit Function
    If cllVisit.Count = 0 Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'zl_CISSvr_UpdOutPatiVisitRec
    'input
    '    --rgst_no   C   1   挂号单号
    '    --exetr C       执行人
    '    --outproom_name C   1   诊室
    '    --exe_time  C   1   执行时间
    '    --rgst_desc C       摘要
    '    --pnurs_oprtr   N   1   护士执行
    
    strJson = ""
    For i = 1 To cllVisit.Count
        Select Case cllVisit(i)(0)
        Case "挂号单号"
            strJson = strJson & "," & GetJsonNodeString("rgst_no", cllVisit(i)(1), Json_Text)
        Case "诊室"
            strJson = strJson & "," & GetJsonNodeString("outproom_name", cllVisit(i)(1), Json_Text)
        Case "执行时间"
            strJson = strJson & "," & GetJsonNodeString("exe_time", cllVisit(i)(1), Json_Text)
        Case "执行人"
            strJson = strJson & "," & GetJsonNodeString("exetr", cllVisit(i)(1), Json_Text)
        Case "执行状态"
            strJson = strJson & "," & GetJsonNodeString("rgst_exe_status", cllVisit(i)(1), Json_num)
        Case "完成时间"
            strJson = strJson & "," & GetJsonNodeString("rgst_fnsh_time", cllVisit(i)(1), Json_Text)
        Case "摘要"
            strJson = strJson & "," & GetJsonNodeString("rgst_desc", cllVisit(i)(1), Json_Text)
        Case "护士执行"
            strJson = strJson & "," & GetJsonNodeString("pnurs_oprtr", Val(cllVisit(i)(1)), Json_num)
        End Select
    Next
    strJson = "{""input"":{" & Mid(strJson, 2) & "}}"
    
    If objServiceCall.CallService("zl_CISSvr_UpdOutPatiVisitRec", strJson, , , lngModule, Not blnNoShowErr) = False Then Exit Function
    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    If intReturn <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlCISSvr_UpdateOutPatiVisitRec = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlPatiSvr_GetblackRegNos(ByVal lng病人ID As Long, ByRef strLastTime As String, ByRef strRegNos As String, Optional ByRef lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据病人ID,获取进入了黑名单的挂号单号
    ' 入参 : lng病人ID
    ' 出参 : strLastTime-最后一次进入不良记录的时间
    '        strRegNos-进入黑名单的挂号单号
    ' 返回 : T-获取数据成功;F-失败
    ' 编制 : 李南春
    ' 日期 : 2019/10/29 13:39
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    If lngModule <> 0 Then mlngModule = lngModule
    Err = 0: On Error GoTo errHandle
    
    strJson = GetJsonNodeString("pati_id", CStr(lng病人ID), 1) '  病人ID
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
  
    '获取服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    If objServiceCall.CallService("Zl_Patisvr_GetblackRegNos", strJson, , "", mlngModule) = False Then Exit Function
    strLastTime = objServiceCall.GetJsonNodeValue("output.last_time")
    strRegNos = objServiceCall.GetJsonNodeValue("output.regnos")
    ZlPatiSvr_GetblackRegNos = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPatiSvr_Savebadrecord(ByVal cllBlackDatas As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 记录病人的不良记录
    ' 入参 : cllBlackDatas ：每个元素都是一个子集合
    '           |-cllBlackData(collect):元素(病人id,行为类别,发生时间,加入时间,加入原因,加入说明,附加信息(N),登记人(N))
    ' 出参 :
    ' 返回 : T-数据保存成功;F-失败
    ' 编制 : 李南春
    ' 日期 : 2019/10/29 13:58
    '---------------------------------------------------------------------------------------
    Dim cllBlackData As Collection
    Dim objServiceCall As Object
    Dim strJson As String, strTempJson As String
    On Error GoTo Errhand
    
    '获取服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
'    Zl_Patisvr_Savebadrecord
'    --功能:病人不良记录数据保存
'    --入参：Json_In:格式
'    --    input
'    --      badrec_list C   保存的不良记录列表
'    --        pati_id N 1 病人id
'    --        behavior_category C 1 行为类别:如预约挂号
'    --        happen_time C 1 发生时间:yyyy-mm-dd hh24:mi:ss
'    --        add_time  C 1 加入时间:yyyy-mm-dd hh24:mi:ss
'    --        add_note  C 1 加入原因：如预约超期
'    --        add_memo  C 1 加入说明
'    --        additional_info C 1 附加信息
'    --        creator C 1 登记人
'    --出参: Json_Out,格式如下
'    --  output
'    --    code              N   1   应答码：0-失败；1-成功
'    --    message           C   1   应答消息：失败时返回具体的错误信息

    If lngModule <> 0 Then mlngModule = lngModule
    
    For Each cllBlackData In cllBlackDatas
        strTempJson = ""
        strTempJson = strTempJson & "" & GetJsonNodeString("pati_id", cllBlackData("_pati_id"), 1)
        strTempJson = strTempJson & "," & GetJsonNodeString("behavior_category", cllBlackData("_behavior_category"))
        strTempJson = strTempJson & "," & GetJsonNodeString("happen_time", cllBlackData("_happen_time"))
        strTempJson = strTempJson & "," & GetJsonNodeString("add_time", cllBlackData("_add_time"))
        strTempJson = strTempJson & "," & GetJsonNodeString("add_note", cllBlackData("_add_note"))
        strTempJson = strTempJson & "," & GetJsonNodeString("add_memo", cllBlackData("_add_memo"))
        strTempJson = strTempJson & "," & GetJsonNodeString("additional_info", zlGetNodeValueFromCollect(cllBlackData, "_additional_info", "C"))
        strTempJson = strTempJson & "," & GetJsonNodeString("creator", zlGetNodeValueFromCollect(cllBlackData, "_creator", "C"))
        strJson = strJson & "," & "{" & strTempJson & "}"
    Next
    If strJson <> "" Then strJson = Mid(strJson, 2)
    strJson = GetNodeString("badrec_list") & ":[" & strJson & "]"
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Patisvr_Savebadrecord", strJson, , "", mlngModule) = False Then Exit Function
    zlPatiSvr_Savebadrecord = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlPatiSvr_NewPatiArchives(ByVal cllPati As Collection, Optional ByRef strErrMsg_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 病人建档或修改病人档案
    ' 入参 : cllPati:病人信息集合
    '           |-成员:(操作类型,病人id,姓名,性别,年龄,出生日期,身份证号,门诊号,费别,医疗付款方式名称,国籍,民族,婚姻状况,职业,
    '           |-身份,工作单位,单位电话,合同单位id,家庭地址,家庭电话,家庭地址邮编,区域,出生地点,户口地址,户口地址邮编,
    '           |-联系人姓名,联系人身份证号,联系人电话,联系人关系,监护人,手机号,医保号,险类,登记时间,操作员姓名 ,
    '           |-操作员编号,社区id,社区号码,社区操作类型,就诊卡号(N),卡验证码(N),IC卡号(N))
    '           |-addr_list(Collect)(地址信息:操作功能,地址类别,地址_省,地址_市,地址_县,地址_乡,地址_其他,区划代码)
    '           |-slave_list(Collect)(病人从项:信息名,信息值)
    '           |-cert_list(Collect)(证件列表:证件名称,证件号码)
    '           |-allergic_drugs_list(Collect)(病人过敏药物:过敏药品ID,过敏药物名称,过每药物反应)
    '           |-immune_list(Collect)(病人免疫信息:接种时间,接种名称)
    '           |-card_property_list(Collect)(医疗卡属性:医疗卡类别ID,卡号,信息名,信息值)
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 14:48
    '---------------------------------------------------------------------------------------
    Dim clldata As Collection, cllListData As Collection
    Dim objServiceCall As Object
    Dim i As Integer
    Dim strJson As String, strJsonList As String, strData As String
    Dim varValue As Variant
    On Error GoTo Errhand

    If lngModule <> 0 Then mlngModule = lngModule
    If cllPati Is Nothing Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function

    strJson = ""
    For i = 1 To cllPati.Count
        If IsArray(cllPati(i)) Then
            Select Case cllPati(i)(0)
                Case "病人ID"
                    strJson = strJson & "," & GetJsonNodeString("pati_id", cllPati(i)(1), Json_num)
                Case "姓名"
                    strJson = strJson & "," & GetJsonNodeString("pati_name", cllPati(i)(1), Json_Text)
                Case "性别"
                    strJson = strJson & "," & GetJsonNodeString("pati_sex", cllPati(i)(1), Json_Text)
                Case "年龄"
                    strJson = strJson & "," & GetJsonNodeString("pati_age", cllPati(i)(1), Json_Text)
                Case "出生日期"
                    strJson = strJson & "," & GetJsonNodeString("pati_birthdate", cllPati(i)(1), Json_Text)
                Case "身份证号"
                    strJson = strJson & "," & GetJsonNodeString("pati_idcard", cllPati(i)(1), Json_Text)
                Case "门诊号"
                    strJson = strJson & "," & GetJsonNodeString("outpatient_num", cllPati(i)(1), Json_Text)
                Case "费别"
                    strJson = strJson & "," & GetJsonNodeString("fee_category", cllPati(i)(1), Json_Text)
                Case "医疗付款方式名称"
                    strJson = strJson & "," & GetJsonNodeString("mdlpay_mode_name", cllPati(i)(1), Json_Text)
                Case "国籍"
                    strJson = strJson & "," & GetJsonNodeString("country_name", cllPati(i)(1), Json_Text)
                Case "民族"
                    strJson = strJson & "," & GetJsonNodeString("nation_name", cllPati(i)(1), Json_Text)
                Case "婚姻状况"
                    strJson = strJson & "," & GetJsonNodeString("mari_status", cllPati(i)(1), Json_Text)
                Case "职业"
                    strJson = strJson & "," & GetJsonNodeString("ocpt_name", cllPati(i)(1), Json_Text)
                Case "身份"
                    strJson = strJson & "," & GetJsonNodeString("pati_identity", cllPati(i)(1), Json_Text)
                Case "工作单位"
                    strJson = strJson & "," & GetJsonNodeString("emp_name", cllPati(i)(1), Json_Text)
                Case "单位电话"
                    strJson = strJson & "," & GetJsonNodeString("emp_phno", cllPati(i)(1), Json_Text)
                Case "合同单位ID"
                    strJson = strJson & "," & GetJsonNodeString("ctt_unit_id", cllPati(i)(1), Json_num)
                Case "单位邮编"
                    strJson = strJson & "," & GetJsonNodeString("emp_postcode", cllPati(i)(1), Json_Text)
                Case "家庭地址"
                    strJson = strJson & "," & GetJsonNodeString("pat_home_addr", cllPati(i)(1), Json_Text)
                Case "家庭电话"
                    strJson = strJson & "," & GetJsonNodeString("pat_home_phno", cllPati(i)(1), Json_Text)
                Case "家庭地址邮编"
                    strJson = strJson & "," & GetJsonNodeString("pat_home_postcode", cllPati(i)(1), Json_Text)
                Case "区域"
                    strJson = strJson & "," & GetJsonNodeString("region", cllPati(i)(1), Json_Text)
                Case "出生地点"
                    strJson = strJson & "," & GetJsonNodeString("pat_baddr", cllPati(i)(1), Json_Text)
                Case "户口地址"
                    strJson = strJson & "," & GetJsonNodeString("pat_hous_addr", cllPati(i)(1), Json_Text)
                Case "户口地址邮编"
                    strJson = strJson & "," & GetJsonNodeString("pat_hous_postcode", cllPati(i)(1), Json_Text)
                Case "监护人"
                    strJson = strJson & "," & GetJsonNodeString("pat_grdn_name", cllPati(i)(1), Json_Text)
                Case "手机号"
                    strJson = strJson & "," & GetJsonNodeString("phone_number", cllPati(i)(1), Json_Text)
                Case "医保号"
                    strJson = strJson & "," & GetJsonNodeString("insurance_num", cllPati(i)(1), Json_Text)
                Case "险类"
                    strJson = strJson & "," & GetJsonNodeString("insurance_type", cllPati(i)(1), Json_num)
                Case "登记时间"
                    strJson = strJson & "," & GetJsonNodeString("create_time", cllPati(i)(1), Json_Text)
                Case "操作员姓名"
                    strJson = strJson & "," & GetJsonNodeString("operator_name", cllPati(i)(1), Json_Text)
                Case "操作员编号"
                    strJson = strJson & "," & GetJsonNodeString("operator_code", cllPati(i)(1), Json_Text)
            End Select
        End If
    Next
    If strJson <> "" Then strJson = Mid(strJson, 2)
    
    'community_info
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "社区信息")
    If Not cllListData Is Nothing Then
        strData = ""
        strData = strData & "" & GetJsonNodeString("num", cllListData("社区序号"), Json_num)
        strData = strData & "," & GetJsonNodeString("code", cllListData("社区号码"), Json_Text)
        strData = strData & "," & GetJsonNodeString("oper_type", cllListData("社区操作类型"), Json_num)
        strJson = strJson & ",""community_info"":{" & strData & "}"
    End If
    
    'sign_info
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "签约信息")
    If Not cllListData Is Nothing Then
        strData = ""
        strData = strData & "" & GetJsonNodeString("card_type_id", cllListData("卡类别ID"), Json_num)
        strData = strData & "," & GetJsonNodeString("card_no", cllListData("卡号"), Json_Text)
        If CollectionExitsValue(cllListData, "卡密码") Then
            strData = strData & "," & GetJsonNodeString("card_pwd", cllListData("卡密码"), Json_Text)
        End If
        If CollectionExitsValue(cllListData, "二维码") Then
            strData = strData & "," & GetJsonNodeString("qrcode", cllListData("二维码"), Json_Text)
        End If
        If CollectionExitsValue(cllListData, "变动原因") Then
            strData = strData & "," & GetJsonNodeString("card_notes", cllListData("变动原因"), Json_Text)
        End If
        If CollectionExitsValue(cllListData, "终止使用时间") Then
            strData = strData & "," & GetJsonNodeString("card_use_endtime", cllListData("终止使用时间"), Json_Text)
        End If
        strJson = strJson & ",""sign_info"":{" & strData & "}"
    End If
    
    'contacts
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "联系人")
    If Not cllListData Is Nothing Then
        strData = ""
        strData = strData & "" & GetJsonNodeString("name", cllListData("联系人姓名"), Json_Text)
        strData = strData & "," & GetJsonNodeString("idcard", cllListData("联系人身份证号"), Json_Text)
        strData = strData & "," & GetJsonNodeString("phone", cllListData("联系人电话"), Json_Text)
        strData = strData & "," & GetJsonNodeString("relation", cllListData("联系人关系"), Json_Text)
        strJson = strJson & ",""contacts"":{" & strData & "}"
    End If
    
    'addr_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "地址信息")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("oper_fun", clldata("操作功能"), Json_num)
                strData = strData & "," & GetJsonNodeString("type", clldata("地址类别"), Json_num)
                strData = strData & "," & GetJsonNodeString("state", clldata("地址_省"), Json_Text)
                strData = strData & "," & GetJsonNodeString("city", clldata("地址_市"), Json_Text)
                strData = strData & "," & GetJsonNodeString("county", clldata("地址_县"), Json_Text)
                strData = strData & "," & GetJsonNodeString("township", clldata("地址_乡"), Json_Text)
                strData = strData & "," & GetJsonNodeString("other", clldata("地址_其他"), Json_Text)
                strData = strData & "," & GetJsonNodeString("code", clldata("区划代码"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""addr_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    'ext_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "从表信息")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("info_name", clldata("信息名"), Json_Text)
                strData = strData & "," & GetJsonNodeString("upd_info_value", clldata("信息值"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""ext_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    'cert_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "证件信息")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("cert_name", clldata("证件名称"), Json_Text)
                strData = strData & "," & GetJsonNodeString("cert_no", clldata("证件号码"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""cert_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    'allergic_drugs_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "过敏药物信息")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("pat_algc_cadn_id", clldata("过敏药品ID"), Json_num)
                strData = strData & "," & GetJsonNodeString("pat_algc_cadn", clldata("过敏药物名称"), Json_Text)
                strData = strData & "," & GetJsonNodeString("allergy_info", clldata("过敏药物反应"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""allergic_drugs_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    'immune_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "免疫信息")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("vaccinate_time", clldata("接种时间"), Json_Text)
                strData = strData & "," & GetJsonNodeString("vaccinate_name", clldata("接种名称"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""immune_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    'card_property_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "医疗卡属性")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("cardtype_id", clldata("医疗卡类别ID"), Json_num)
                strData = strData & "," & GetJsonNodeString("card_no", clldata("卡号"), Json_Text)
                strData = strData & "," & GetJsonNodeString("info_name", clldata("信息名"), Json_Text)
                strData = strData & "," & GetJsonNodeString("info_value", clldata("信息值"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""card_property_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_NewPatiArchives", strJson, , , lngModule, False) = False Then Exit Function
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    
    ZlPatiSvr_NewPatiArchives = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlPatisvr_UpdatePatiArchives(ByVal cllPati As Collection, Optional ByRef strErrMsg_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 病人建档或修改病人档案
    ' 入参 : cllPati:病人信息集合
    '           |-成员:(操作类型,病人id,姓名,性别,年龄,出生日期,身份证号,门诊号,费别,医疗付款方式名称,国籍,民族,婚姻状况,职业,
    '           |-身份,工作单位,单位电话,合同单位id,家庭地址,家庭电话,家庭地址邮编,区域,出生地点,户口地址,户口地址邮编,
    '           |-联系人姓名,联系人身份证号,联系人电话,联系人关系,监护人,手机号,医保号,险类,登记时间,操作员姓名 ,
    '           |-操作员编号,社区id,社区号码,社区操作类型,就诊卡号(N),卡验证码(N),IC卡号(N))
    '           |-addr_list(Collect)(地址信息:操作功能,地址类别,地址_省,地址_市,地址_县,地址_乡,地址_其他,区划代码)
    '           |-slave_list(Collect)(病人从项:信息名,信息值)
    '           |-cert_list(Collect)(证件列表:证件名称,证件号码)
    '           |-allergic_drugs_list(Collect)(病人过敏药物:过敏药品ID,过敏药物名称,过每药物反应)
    '           |-immune_list(Collect)(病人免疫信息:接种时间,接种名称)
    '           |-card_property_list(Collect)(医疗卡属性:医疗卡类别ID,卡号,信息名,信息值)
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 14:48
    '---------------------------------------------------------------------------------------
    Dim clldata As Collection, cllListData As Collection
    Dim objServiceCall As Object
    Dim i As Integer
    Dim strJson As String, strJsonList As String, strData As String
    Dim varValue As Variant
    On Error GoTo Errhand

    If lngModule <> 0 Then mlngModule = lngModule
    If cllPati Is Nothing Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function

    strJson = ""
    For i = 1 To cllPati.Count
        If IsArray(cllPati(i)) Then
            Select Case cllPati(i)(0)
                Case "病人ID"
                    strJson = strJson & "," & GetJsonNodeString("pati_id", cllPati(i)(1), Json_num)
                Case "姓名"
                    strJson = strJson & "," & GetJsonNodeString("pati_name", cllPati(i)(1), Json_Text)
                Case "性别"
                    strJson = strJson & "," & GetJsonNodeString("pati_sex", cllPati(i)(1), Json_Text)
                Case "年龄"
                    strJson = strJson & "," & GetJsonNodeString("pati_age", cllPati(i)(1), Json_Text)
                Case "出生日期"
                    strJson = strJson & "," & GetJsonNodeString("pati_birthdate", cllPati(i)(1), Json_Text)
                Case "身份证号"
                    strJson = strJson & "," & GetJsonNodeString("pati_idcard", cllPati(i)(1), Json_Text)
                Case "门诊号"
                    strJson = strJson & "," & GetJsonNodeString("outpatient_num", cllPati(i)(1), Json_Text)
                Case "费别"
                    strJson = strJson & "," & GetJsonNodeString("fee_category", cllPati(i)(1), Json_Text)
                Case "医疗付款方式名称"
                    strJson = strJson & "," & GetJsonNodeString("mdlpay_mode_name", cllPati(i)(1), Json_Text)
                Case "国籍"
                    strJson = strJson & "," & GetJsonNodeString("country_name", cllPati(i)(1), Json_Text)
                Case "民族"
                    strJson = strJson & "," & GetJsonNodeString("nation_name", cllPati(i)(1), Json_Text)
                Case "婚姻状况"
                    strJson = strJson & "," & GetJsonNodeString("mari_status", cllPati(i)(1), Json_Text)
                Case "职业"
                    strJson = strJson & "," & GetJsonNodeString("ocpt_name", cllPati(i)(1), Json_Text)
                Case "身份"
                    strJson = strJson & "," & GetJsonNodeString("pati_identity", cllPati(i)(1), Json_Text)
                Case "工作单位"
                    strJson = strJson & "," & GetJsonNodeString("emp_name", cllPati(i)(1), Json_Text)
                Case "单位电话"
                    strJson = strJson & "," & GetJsonNodeString("emp_phno", cllPati(i)(1), Json_Text)
                Case "合同单位ID"
                    strJson = strJson & "," & GetJsonNodeString("ctt_unit_id", cllPati(i)(1), Json_num)
                Case "单位邮编"
                    strJson = strJson & "," & GetJsonNodeString("emp_postcode", cllPati(i)(1), Json_Text)
                Case "家庭地址"
                    strJson = strJson & "," & GetJsonNodeString("pat_home_addr", cllPati(i)(1), Json_Text)
                Case "家庭电话"
                    strJson = strJson & "," & GetJsonNodeString("pat_home_phno", cllPati(i)(1), Json_Text)
                Case "家庭地址邮编"
                    strJson = strJson & "," & GetJsonNodeString("pat_home_postcode", cllPati(i)(1), Json_Text)
                Case "区域"
                    strJson = strJson & "," & GetJsonNodeString("region", cllPati(i)(1), Json_Text)
                Case "出生地点"
                    strJson = strJson & "," & GetJsonNodeString("pat_baddr", cllPati(i)(1), Json_Text)
                Case "户口地址"
                    strJson = strJson & "," & GetJsonNodeString("pat_hous_addr", cllPati(i)(1), Json_Text)
                Case "户口地址邮编"
                    strJson = strJson & "," & GetJsonNodeString("pat_hous_postcode", cllPati(i)(1), Json_Text)
                Case "监护人"
                    strJson = strJson & "," & GetJsonNodeString("pat_grdn_name", cllPati(i)(1), Json_Text)
                Case "手机号"
                    strJson = strJson & "," & GetJsonNodeString("phone_number", cllPati(i)(1), Json_Text)
                Case "医保号"
                    strJson = strJson & "," & GetJsonNodeString("insurance_num", cllPati(i)(1), Json_Text)
                Case "险类"
                    strJson = strJson & "," & GetJsonNodeString("insurance_type", cllPati(i)(1), Json_num)
                Case "操作员姓名"
                    strJson = strJson & "," & GetJsonNodeString("operator_name", cllPati(i)(1), Json_Text)
                Case "操作员编号"
                    strJson = strJson & "," & GetJsonNodeString("operator_code", cllPati(i)(1), Json_Text)
            End Select
        End If
    Next
    If strJson <> "" Then strJson = Mid(strJson, 2)
    
    'visit_info
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "就诊信息")
    If Not cllListData Is Nothing Then
        strData = ""
        If CollectionExitsValue(cllListData, "就诊状态") Then
            strData = strData & "," & GetJsonNodeString("status", cllListData("就诊状态"), Json_num)
        End If
        If CollectionExitsValue(cllListData, "就诊诊室") Then
            strData = strData & "," & GetJsonNodeString("room", cllListData("就诊诊室"), Json_Text)
        End If
        If CollectionExitsValue(cllListData, "就诊时间") Then
            strData = strData & "," & GetJsonNodeString("time", cllListData("就诊时间"), Json_Text)
        End If
        strJson = strJson & ",""visit_info"":{" & Mid(strData, 2) & "}"
    End If
    
    'sign_info
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "签约信息")
    If Not cllListData Is Nothing Then
        strData = ""
        strData = strData & "" & GetJsonNodeString("card_type_id", cllListData("卡类别ID"), Json_num)
        strData = strData & "," & GetJsonNodeString("card_no", cllListData("卡号"), Json_Text)
        If CollectionExitsValue(cllListData, "卡密码") Then
            strData = strData & "," & GetJsonNodeString("card_pwd", cllListData("卡密码"), Json_Text)
        End If
        If CollectionExitsValue(cllListData, "二维码") Then
            strData = strData & "," & GetJsonNodeString("qrcode", cllListData("二维码"), Json_Text)
        End If
        If CollectionExitsValue(cllListData, "变动原因") Then
            strData = strData & "," & GetJsonNodeString("card_notes", cllListData("变动原因"), Json_Text)
        End If
        If CollectionExitsValue(cllListData, "终止使用时间") Then
            strData = strData & "," & GetJsonNodeString("card_use_endtime", cllListData("终止使用时间"), Json_Text)
        End If
        strJson = strJson & ",""sign_info"":{" & strData & "}"
    End If
    
    'community_info
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "社区信息")
    If Not cllListData Is Nothing Then
        strData = ""
        strData = strData & "" & GetJsonNodeString("num", cllListData("社区序号"), Json_num)
        strData = strData & "," & GetJsonNodeString("code", cllListData("社区号码"), Json_Text)
        strData = strData & "," & GetJsonNodeString("oper_type", cllListData("社区操作类型"), Json_num)
        strJson = strJson & ",""community_info"":{" & strData & "}"
    End If
    
    'contacts
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "联系人")
    If Not cllListData Is Nothing Then
        strData = ""
        strData = strData & "" & GetJsonNodeString("name", cllListData("联系人姓名"), Json_Text)
        strData = strData & "," & GetJsonNodeString("idcard", cllListData("联系人身份证号"), Json_Text)
        strData = strData & "," & GetJsonNodeString("phone", cllListData("联系人电话"), Json_Text)
        strData = strData & "," & GetJsonNodeString("relation", cllListData("联系人关系"), Json_Text)
        strJson = strJson & ",""contacts"":{" & strData & "}"
    End If
    
    'addr_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "地址信息")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("oper_fun", clldata("操作功能"), Json_num)
                strData = strData & "," & GetJsonNodeString("type", clldata("地址类别"), Json_num)
                strData = strData & "," & GetJsonNodeString("state", clldata("地址_省"), Json_Text)
                strData = strData & "," & GetJsonNodeString("city", clldata("地址_市"), Json_Text)
                strData = strData & "," & GetJsonNodeString("county", clldata("地址_县"), Json_Text)
                strData = strData & "," & GetJsonNodeString("township", clldata("地址_乡"), Json_Text)
                strData = strData & "," & GetJsonNodeString("other", clldata("地址_其他"), Json_Text)
                strData = strData & "," & GetJsonNodeString("code", clldata("区划代码"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""addr_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    'ext_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "从表信息")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("info_name", clldata("信息名"), Json_Text)
                strData = strData & "," & GetJsonNodeString("upd_info_value", clldata("信息值"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""ext_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    'cert_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "证件信息")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("cert_name", clldata("证件名称"), Json_Text)
                strData = strData & "," & GetJsonNodeString("cert_no", clldata("证件号码"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""cert_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    'allergic_drugs_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "过敏药物信息")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("pat_algc_cadn_id", clldata("过敏药品ID"), Json_num)
                strData = strData & "," & GetJsonNodeString("pat_algc_cadn", clldata("过敏药物名称"), Json_Text)
                strData = strData & "," & GetJsonNodeString("allergy_info", clldata("过敏药物反应"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""allergic_drugs_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    'immune_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "免疫信息")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("vaccinate_time", clldata("接种时间"), Json_Text)
                strData = strData & "," & GetJsonNodeString("vaccinate_name", clldata("接种名称"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""immune_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    
    'card_property_list
    strJsonList = ""
    Set cllListData = zlGetNodeObjectFromCollect(cllPati, "医疗卡属性")
    If Not cllListData Is Nothing Then
        If cllListData.Count <> 0 Then
            For i = 1 To cllListData.Count
                Set clldata = cllListData(i)
                strData = ""
                strData = strData & "" & GetJsonNodeString("cardtype_id", clldata("医疗卡类别ID"), Json_num)
                strData = strData & "," & GetJsonNodeString("card_no", clldata("卡号"), Json_Text)
                strData = strData & "," & GetJsonNodeString("info_name", clldata("信息名"), Json_Text)
                strData = strData & "," & GetJsonNodeString("info_value", clldata("信息值"), Json_Text)
                strJsonList = strJsonList & ",{" & strData & "}"
            Next
            strJson = strJson & ",""card_property_list"":[" & Mid(strJsonList, 2) & "]"
        End If
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Patisvr_Updatepatiarchives", strJson, , , lngModule, False) = False Then Exit Function
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    
    ZlPatisvr_UpdatePatiArchives = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPatiSvr_BindOutPno(ByVal lng病人ID As Long, ByVal str门诊号 As String, _
                Optional ByVal str医保号 As String, Optional ByVal intInsure As Integer, _
                Optional ByVal blnNoShowErr As Boolean, Optional ByRef strErrMsg_Out As String, _
                Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新门诊病案记录
    '入参:cllOutMedRec-门诊病案数据集:array(名称,值)
    '                名称包含（病人id,.病案号(门诊号),建立日期,病案类别,存储状态,存放位置)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strServiceName  As String
    Dim objServiceCall As Object
    Dim intReturn As Integer
    
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Patisvr_Updatepatiarchives
    '    input
    '       pati_id N   1   病人id
    '       outpatient_num        C  1  门诊号
    '       insurance_num         C  1  医保号
    '       insurance_type        N  1  险类
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("outpatient_num", str门诊号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("insurance_num", str医保号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("insurance_type", intInsure, Json_num)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_Patisvr_Updatepatiarchives"
    If objServiceCall.CallService(strServiceName, strJson, , strFrmCaption, glngModul, Not blnNoShowErr, strErrMsg_Out, , True) = False Then Exit Function
    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    If intReturn <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlPatiSvr_BindOutPno = True
End Function

Public Function ZlPatiSvr_GetNextNo(ByVal lng项目序号 As Long, Optional ByVal lng科室id As Long, Optional ByVal lngModule As Long) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : zlDataBase.GetNextNo
    ' 入参 : 病人就诊信息等
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 14:48
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim strValue As String
    On Error GoTo Errhand

    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("item_num", lng项目序号, Json_num)
    strJson = strJson & "," & GetJsonNodeString("dept_id", lng科室id, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_PatiSvr_GetNextNo", strJson, , , lngModule) = False Then Exit Function
    strValue = objServiceCall.GetJsonNodeValue("output.next_no")
    
    ZlPatiSvr_GetNextNo = strValue
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPatiSvr_PatiPhoto_Save(ByVal lng病人ID As Long, ByVal strPhotoCode As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 保存病人照片
    ' 入参 : 病人就诊信息等
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 14:48
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    On Error GoTo Errhand

    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
'   zl_PatiSvr_PatiPhoto_Save
'    --input
'    --pati_id N   1   病人id
'    --pati_photo  C   1   编码:base64
'    --output
'    --code    N   1   应答码：0-失败；1-成功
'    --message C   1   "应答消息：成功时返回成功信息


    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_photo", strPhotoCode, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_PatiPhoto_Save", strJson, , , lngModule) = False Then Exit Function
    
    zlPatiSvr_PatiPhoto_Save = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPatiSvr_PatiPhoto_Delete(ByVal lng病人ID As Long, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 保存病人照片
    ' 入参 : 病人就诊信息等
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 14:48
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    On Error GoTo Errhand

    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
'   zl_PatiSvr_PatiPhoto_Delete
'    --input
'    --pati_id N   1   病人id
'    --output
'    --code    N   1   应答码：0-失败；1-成功
'    --message C   1   "应答消息：成功时返回成功信息


    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_PatiPhoto_Delete", strJson, , , lngModule) = False Then Exit Function
    
    zlPatiSvr_PatiPhoto_Delete = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPatiSvr_PatiPhoto_Read(ByVal lng病人ID As Long, ByRef strPhotoCode As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 保存病人照片
    ' 入参 : 病人就诊信息等
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 14:48
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    On Error GoTo Errhand

    If lngModule <> 0 Then mlngModule = lngModule
    strPhotoCode = ""
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
'   zl_PatiSvr_PatiPhoto_Read
'    --input
'    --pati_id N   1   病人id
'    --output
'    --code    N   1   应答码：0-失败；1-成功
'    --message C   1   "应答消息：成功时返回成功信息
'    --pati_photo  C   1   编码:base64


    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_PatiPhoto_Read", strJson, , , lngModule) = False Then Exit Function
    strPhotoCode = objServiceCall.GetJsonNodeValue("output.pati_photo")
    
    zlPatiSvr_PatiPhoto_Read = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetPatiPageInfo(ByVal byt查询类型 As Byte, _
    ByVal str病人信息 As String, ByRef rsPatiPageInfo As ADODB.Recordset, _
    Optional ByVal bln包含婴儿信息 As Boolean, _
    Optional ByRef cllBabyInfo As Collection, Optional ByVal lngModule As Long, _
    Optional ByVal bln取最大主页ID As Boolean = True, _
    Optional ByVal str病人性质 As String, Optional ByVal lng挂号ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查询病人主页信息
    '入参:
    '   byt查询类型 查询类型:0-基本信息;1-基本信息的扩展;2-仅取主页
    '   str病人信息 病人信息,格式两种:一种是:病人id:主页ID,…;一种：病人id,…
    '   bln包含婴儿信息 是否包含婴儿信息
    '   bln取最大主页ID 是否取最后一次住院
    '   str病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人；多个逗号分隔，不传为所有
    '   lng挂号ID:挂号ID，按挂号ID查询
    '出参:
    '   rsPatiPageInfo 病人病案主页信息：病人id,主页id,住院号,姓名,性别,年龄,费别,病人性质,审核标志,住院状态,入院时间,出院时间,住院医师,
    '                                   医疗付款方式名称,医疗付款方式编码,当前病区id,当前病区名称,当前科室id,当前科室名称,医保号,
    '                                   险类,当前床号,病人类型,[学历,职业,国籍,婚姻状况,编目日期,病人备注,挂号id]
    '   cllBabyInfo 婴儿信息,成员：ADODB.Recordset=cllBabyInfo(_病人ID)，字段：病人id,主页id,序号,姓名,性别,出生时间
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, i As Long, cllTemp As Collection
    Dim rsBabyInfo As ADODB.Recordset, cllBaby As Collection, j As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    Set rsPatiPageInfo = New ADODB.Recordset
    With rsPatiPageInfo.Fields
        '病人id,主页id,住院号,姓名,性别,年龄,费别,病人性质,审核标志,住院状态,入院时间,出院时间,住院医师,
        .Append "病人ID", adBigInt, 18, adFldIsNullable
        .Append "主页ID", adBigInt, 18, adFldIsNullable
        .Append "姓名", adVarChar, 100, adFldIsNullable
        
        If byt查询类型 = 0 Or byt查询类型 = 1 Then
            .Append "住院号", adVarChar, 18, adFldIsNullable
            .Append "性别", adVarChar, 4, adFldIsNullable
            .Append "年龄", adVarChar, 20, adFldIsNullable
            .Append "费别", adVarChar, 10, adFldIsNullable
            .Append "病人性质", adInteger, 1, adFldIsNullable
            .Append "审核标志", adInteger, 1, adFldIsNullable
            .Append "住院状态", adInteger, 3, adFldIsNullable
            .Append "入院时间", adVarChar, 20, adFldIsNullable
            .Append "出院时间", adVarChar, 20, adFldIsNullable
            .Append "住院医师", adVarChar, 100, adFldIsNullable
            
            If byt查询类型 = 1 Then
                '医疗付款方式名称,医疗付款方式编码,当前病区id,当前病区名称,当前科室id,当前科室名称,医保号,
                .Append "医疗付款方式编码", adVarChar, 2, adFldIsNullable
                .Append "医疗付款方式名称", adVarChar, 20, adFldIsNullable
                .Append "当前病区ID", adBigInt, 18, adFldIsNullable
                .Append "当前病区名称", adVarChar, 100, adFldIsNullable
                .Append "当前科室ID", adBigInt, 18, adFldIsNullable
                .Append "当前科室名称", adVarChar, 100, adFldIsNullable
                .Append "医保号", adVarChar, 50, adFldIsNullable
                '险类,当前床号,病人类型,学历,职业,国籍,婚姻状况,编目日期,病人备注,挂号id
                .Append "险类", adInteger, 3, adFldIsNullable
                .Append "当前床号", adVarChar, 10, adFldIsNullable
                .Append "病人类型", adVarChar, 50, adFldIsNullable
                .Append "学历", adVarChar, 10, adFldIsNullable
                .Append "职业", adVarChar, 80, adFldIsNullable
                .Append "国籍", adVarChar, 30, adFldIsNullable
                .Append "婚姻状况", adVarChar, 4, adFldIsNullable
                .Append "适用病人", adVarChar, 50, adFldIsNullable
                .Append "担保额", adDouble, , adFldIsNullable
                .Append "编目日期", adVarChar, 20, adFldIsNullable
                .Append "病人备注", adVarChar, 100, adFldIsNullable
                .Append "挂号id", adBigInt, 18, adFldIsNullable
            End If
        End If
    End With
    rsPatiPageInfo.CursorLocation = adUseClient
    rsPatiPageInfo.LockType = adLockOptimistic
    rsPatiPageInfo.CursorType = adOpenStatic
    rsPatiPageInfo.Open
    Set cllBabyInfo = New Collection
    
    'Zl_Cissvr_Getpatipageinfo
    '  --功能:获取病案主页相关信息
    '  --入参：Json_In:格式
    '  --    input
    '  --      query_type          C 1 查询类型:0-基本信息;1-基本信息的展;2-仅取主页
    '  --      pati_pageids        C 1 病人信息,格式两种:一种是:病人id:主页ID,…;一种：病人id,…
    '  --      is_babyinfo         N 1 是否包含婴儿信息:1-包含;0-不包含
    '  --      is_transdeptinfo    N 1 是否包含转科信息:1-包含;0-不包含
    '  --      is_lastpage         N 1 是否取最后一次住院
    '  --      pati_natures        C 1 病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人；多个逗号分隔，不传为所有
    '  --      rgst_id             N 1 挂号ID,根据挂号ID查询
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", byt查询类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", str病人信息, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("is_babyinfo", IIf(bln包含婴儿信息, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("is_lastpage", IIf(bln取最大主页ID, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_natures", str病人性质, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("rgst_id", lng挂号ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getpatipageinfo", strJson, , "", mlngModule) = False Then Exit Function
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                  N 1 应答码：0-失败；1-成功
    '  --    message               C 1 应答消息：失败时返回具体的错误信息
    '  --    pati_count            N 1 查询的病人信息条数
    '  --    page_list[]             1 数据组
    '  --      pati_id             N 1 病人id
    '  --      pati_pageid         N 1 主页id
    '  --      pati_name           C 1 姓名
    '  --      pati_sex            C 1 性别
    '  --      pati_age            C 1 年龄
    '  --      inpatient_num       C 1 住院号
    '  --      fee_category        C 1 费别
    '  --      mdlpay_mode_name    C 1 医疗付款方式名称
    '  --      mdlpay_mode_code    C 1 医疗付款方式编码
    '  --      pati_bed            C 1 当前床号
    '  --      pati_type           C 1 病人类型(普通，医保，留观)
    '  --      pati_show_color     N 1 病人显示颜色
    '  --      pati_education      C 1 学历
    '  --      ocpt_name           C 1 职业
    '  --      country_name        C 1 国籍
    '  --      pati_marital_cstatus  C 1 婚姻状况
    '  --      pati_nature         N 1 病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
    '  --      audit_sign          N 1 审核标志:病案主页.审核标志
    '  --      si_inp_status       N 1 住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)
    '  --      pati_wardarea_id    N 1 当前病区id
    '  --      pati_wardarea_name  C 1 当前病区名称
    '  --      pati_dept_id        N 1 当前科室id
    '  --      pati_dept_name      C 1 当前科室名称
    '  --      adta_time           C 1 入院时间:yyyy-mm-dd hh24:mi:ss
    '  --      adtd_time           C 1 出院时间:yyyy-mm-dd hh24:mi:ss
    '  --      insurance_type      N 1 险类
    '  --      rgst_id             N 1 挂号id
    '  --      catalog_date        C 1 编目日期:yyyy-mm-dd hh24:mi:ss
    '  --      in_objective        C 1 住院目的
    '  --      reg_name            C 1 登记人
    '  --      reg_date            C 1 住院登记时间
    '  --      pat_rsdpscn         C 1 住院医师
    '  --      pati_desc           C 1 病人备注
    '  --      insurance_num       C 1 医保号
    '  --      outpatient_doctor   C 1 门诊医师
    '  --      responsible_nurse   C 1 责任护士
    '  --      hospital_admissions C 1 入院病况
    '  --      current_conditions  C 1 当前病况
    '  --      hospital_days       N 1 住院天数
    '  --      hospital_dept       C 1 入院科室
    '  --      level_of_care       C 1 护理等级
    '  --      level_of_bed        C 1 床位等级
    '  --      in_dept             N 1 是否已入科
    '  --      baby_list[]           1 婴儿信息，[数组]
    '  --        pati_id           N 1 病人id
    '  --        pati_pageid       N 1 主页id
    '  --        baby_num          N 1 婴儿序号
    '  --        baby_name         C 1 婴儿姓名
    '  --        baby_sex          C 1 婴儿性别
    '  --        baby_date         C 1 出生时间
    '  --      trans_list[]        C   转科列表信息
    '  --        start_reason      C 1 开始原因
    '  --        start_time        C 1 开始时间:yyyy-mm-dd hh24:mi:ss
    '  --        dept_name         C 1 科室名称
    Set clldata = objServiceCall.GetJsonListValue("output.page_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            With rsPatiPageInfo
                .AddNew
                '病人id,主页id,住院号,姓名,性别,年龄,费别,病人性质,审核标志,住院状态,入院时间,出院时间,住院医师,
                !病人ID = Val(Nvl(cllTemp("_pati_id")))
                !主页ID = Val(Nvl(cllTemp("_pati_pageid")))
                !姓名 = cllTemp("_pati_name")
                    
                If byt查询类型 = 0 Or byt查询类型 = 1 Then
                    !性别 = cllTemp("_pati_sex")
                    !年龄 = cllTemp("_pati_age")
                    !住院号 = CStr(Nvl(cllTemp("_inpatient_num")))
                    !费别 = cllTemp("_fee_category")
                    !病人性质 = Val(Nvl(cllTemp("_pati_nature")))
                    !审核标志 = Val(Nvl(cllTemp("_audit_sign")))
                    !住院状态 = Val(Nvl(cllTemp("_si_inp_status")))
                    !入院时间 = cllTemp("_adta_time")
                    !出院时间 = cllTemp("_adtd_time")
                    !住院医师 = cllTemp("_pat_rsdpscn")
                    
                    If byt查询类型 = 1 Then
                        '医疗付款方式名称,医疗付款方式编码,当前病区id,当前病区名称,当前科室id,当前科室名称,医保号,
                        !医疗付款方式编码 = cllTemp("_mdlpay_mode_code")
                        !医疗付款方式名称 = cllTemp("_mdlpay_mode_name")
                        !当前病区id = Val(Nvl(cllTemp("_pati_wardarea_id")))
                        !当前病区名称 = cllTemp("_pati_wardarea_name")
                        !当前科室id = Val(Nvl(cllTemp("_pati_dept_id")))
                        !当前科室名称 = cllTemp("_pati_dept_name")
                        !医保号 = Nvl(cllTemp("_insurance_num"))
                        '险类,当前床号,病人类型,学历,职业,国籍,婚姻状况,适用病人,病人备注,挂号id
                        !险类 = Val(Nvl(cllTemp("_insurance_type")))
                        !当前床号 = cllTemp("_pati_bed")
                        !病人类型 = cllTemp("_pati_type")
                        !学历 = cllTemp("_pati_education")
                        !职业 = cllTemp("_ocpt_name")
                        !国籍 = cllTemp("_country_name")
                        !婚姻状况 = cllTemp("_pati_marital_cstatus")
                        !编目日期 = cllTemp("_catalog_date")
                        !病人备注 = cllTemp("_pati_desc")
                        !挂号ID = Val(Nvl(cllTemp("_rgst_id")))
                    End If
    
                    If bln包含婴儿信息 Then
                        Set rsBabyInfo = New ADODB.Recordset
                        With rsBabyInfo.Fields
                            '病人id,主页id,序号,姓名,性别,出生时间
                            .Append "病人ID", adBigInt, 18, adFldIsNullable
                            .Append "主页ID", adBigInt, 18, adFldIsNullable
                            .Append "序号", adInteger, 2, adFldIsNullable
                            .Append "姓名", adVarChar, 100, adFldIsNullable
                            .Append "性别", adVarChar, 4, adFldIsNullable
                            .Append "出生时间", adVarChar, 20, adFldIsNullable
                        End With
                        rsBabyInfo.CursorLocation = adUseClient
                        rsBabyInfo.LockType = adLockOptimistic
                        rsBabyInfo.CursorType = adOpenStatic
                        rsBabyInfo.Open
                        
                        Set cllBaby = objServiceCall.GetJsonListValue("output.page_list[" & i - 1 & "].baby_list")
                        If Not cllBaby Is Nothing Then
                            For j = 1 To cllBaby.Count
                                Set cllTemp = cllBaby(j)
                                With rsBabyInfo
                                    .AddNew
                                    '病人id,主页id,序号,姓名,性别,出生时间
                                    !病人ID = cllTemp("_pati_id")
                                    !主页ID = cllTemp("_pati_pageid")
                                    !序号 = cllTemp("_baby_num")
                                    !姓名 = cllTemp("_baby_name")
                                    !性别 = cllTemp("_baby_sex")
                                    !出生时间 = cllTemp("_baby_date")
                                    .Update
                                End With
                            Next
                            If rsBabyInfo.RecordCount > 0 Then rsBabyInfo.MoveFirst
                        End If
                        cllBabyInfo.Add rsBabyInfo, "_" & !病人ID
                    End If
                End If

                .Update
            End With
        Next
        If rsPatiPageInfo.RecordCount > 0 Then rsPatiPageInfo.MoveFirst
    End If
    
    ZlCissvr_GetPatiPageInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetPatiPageInfByRange(ByVal cllFilter As Collection, _
    ByRef rsPatiPageInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long, Optional ByVal bln包含婴儿信息 As Boolean, _
    Optional ByRef cllBabyInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查询病人主页信息
    '入参:
    '   cllFilter 查询条件集:成员(Array(Key,Value),Array(Key,Value),,...)
    '       Key:病区IDS,科室IDS,病人IDS,主页IDS,入院开始时间,入院结束时间,出院开始时间,出院结束时间,
    '           费别,住院状态,病人性质,姓名,站点编号,查询转科病人,最后一次住院,险类,病区站点编号
    '       住院状态:0-在院病人;1-出院病人;2-在院或出院
    '       病人性质：多个用逗号分0-普通住院病人,1-门诊留观病人,2-住院留观病人，NULL-表示不区分
    '       姓名:可以代%分号表表按姓名匹配
    '       已出院天数，住院状态为1和2时有效
    '       站点编号:科室对应的站点编号
    '       险类:>0:指定险类医保病人,0:医保和普通病人,-1:普通病人,-2:医保病人
    '   bln包含婴儿信息 是否包含婴儿信息
    '出参:
    '   rsPatiPageInfo 病人病案主页信息：病人ID,主页ID,姓名,性别,年龄,住院号,床号,险类,费别,病人类型,医保号,
    '                                   入院时间,出院时间,住院状态,病人性质,当前病区ID,当前病区名称,当前科室ID,当前科室名称,
    '                                   医疗付款方式名称,医疗付款方式编码,住院医师,病人备注,编目日期,护理等级,
    '                                   数据转出,审核标志,审核人,预出院时间,上次催款金额
    '       住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)
    '       病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
    '       数据转出:0-未转出，1-已转出
    '       审核标志:0或空-未审核,1-已审核或开始审核;2-完成审核
    '   cllBabyInfo 婴儿信息,成员：ADODB.Recordset=cllBabyInfo(_病人ID_主页ID)，字段：病人id,主页id,序号,姓名,性别,出生时间
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, i As Long, cllTemp As Collection
    Dim rsBabyInfo As ADODB.Recordset, cllBaby As Collection, j As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
     
    Set rsPatiPageInfo = New ADODB.Recordset
    With rsPatiPageInfo.Fields
        '病人ID,主页ID,姓名,性别,年龄,住院号,床号,险类,费别,病人类型,医保号,
        .Append "病人ID", adBigInt, 18, adFldIsNullable
        .Append "主页ID", adBigInt, 18, adFldIsNullable
        .Append "姓名", adVarChar, 100, adFldIsNullable
        .Append "性别", adVarChar, 4, adFldIsNullable
        .Append "年龄", adVarChar, 20, adFldIsNullable
        .Append "住院号", adVarChar, 18, adFldIsNullable
        .Append "床号", adVarChar, 10, adFldIsNullable
        .Append "险类", adInteger, 3, adFldIsNullable
        .Append "费别", adVarChar, 10, adFldIsNullable
        .Append "病人类型", adVarChar, 50, adFldIsNullable
        .Append "医保号", adVarChar, 50, adFldIsNullable
        '入院时间,出院时间,住院状态,病人性质,当前病区ID,当前病区名称,当前科室id,当前科室名称,
        .Append "入院时间", adVarChar, 20, adFldIsNullable
        .Append "出院时间", adVarChar, 20, adFldIsNullable
        .Append "住院状态", adInteger, 3, adFldIsNullable
        .Append "病人性质", adInteger, 1, adFldIsNullable
        .Append "当前病区ID", adBigInt, 18, adFldIsNullable
        .Append "当前病区名称", adVarChar, 100, adFldIsNullable
        .Append "当前科室ID", adBigInt, 18, adFldIsNullable
        .Append "当前科室名称", adVarChar, 100, adFldIsNullable
        '医疗付款方式名称,医疗付款方式编码,住院医师,病人备注,编目日期,护理等级
        .Append "医疗付款方式编码", adVarChar, 2, adFldIsNullable
        .Append "医疗付款方式名称", adVarChar, 20, adFldIsNullable
        .Append "住院医师", adVarChar, 100, adFldIsNullable
        .Append "病人备注", adVarChar, 100, adFldIsNullable
        .Append "编目日期", adVarChar, 20, adFldIsNullable
        .Append "护理等级", adVarChar, 100, adFldIsNullable
        '数据转出,审核标志,审核人,预出院时间,上次催款金额
        .Append "数据转出", adInteger, 1, adFldIsNullable
        .Append "审核标志", adInteger, 1, adFldIsNullable
        .Append "审核人", adVarChar, 100, adFldIsNullable
        .Append "预出院时间", adVarChar, 20, adFldIsNullable
        .Append "上次催款金额", adCurrency, , adFldIsNullable
    End With
    rsPatiPageInfo.CursorLocation = adUseClient
    rsPatiPageInfo.LockType = adLockOptimistic
    rsPatiPageInfo.CursorType = adOpenStatic
    rsPatiPageInfo.Open
    
    Set cllBabyInfo = New Collection
    
    If cllFilter Is Nothing Then Exit Function
    'Zl_Cissvr_Getpatpageinfbyrange
    '  --功能:获取病案信息
    '  --入参：Json_In:格式
    '  --  input
    '  --    query_type          N 1 查询类型:0-基本;1-基本扩展
    '  --    wararea_ids         C   病区ids:多个用逗号
    '  --    dept_ids            C   科室IDs:个用逗号
    '  --    pati_ids            C   病人ids:多个用逗号分离
    '  --    pati_pageIds        C   主页IDs:病人id:主页id,…
    '  --    adta_start_time     C   入院开始时间:yyyy-mm-dd hh24:mi:ss
    '  --    adta_end_time       C   入院结束时间:yyyy-mm-dd hh24:mi:ss
    '  --    adtd_start_time     C   出院开始时间:yyyy-mm-dd hh24:mi:ss
    '  --    adtd_end_time       C   出院结束时间:yyyy-mm-dd hh24:mi:ss
    '  --    fee_category        C   费别
    '  --    inp_status          N   住院状态:0-在院病人;1-出院病人;2-在院或出院
    '  --    pati_natures        C   病人性质：多个用逗号分0-普通住院病人,1-门诊留观病人,2-住院留观病人，NULL-表示不区分
    '  --    pati_name           C   姓名:可以代%分号表表按姓名匹配
    '  --    pati_out_days       N   已出院天数，住院状态为1和2时有效
    '  --    dept_nodeno         C   科室站点编号
    '  --    change_dept_pati    N   是否查询转科病人
    '  --    is_lastpage         N   是否取最后一次住院 0-最后一次 1-不是最后一次
    '  --    insurance_type      N   根据险类查找:>0:指定险类医保病人,0:医保和普通病人,-1:普通病人,-2:医保病人
    '  --    wararea_nodeno      C   病区站点编号
    '  --    is_babyinfo         N 1 是否包含婴儿信息:1-包含;0-不包含
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", 1, Json_num)
    For i = 1 To cllFilter.Count
        Select Case cllFilter(i)(0)
        Case "病区IDS"
            strJson = strJson & "," & GetJsonNodeString("wararea_ids", cllFilter(i)(1), Json_Text)
        Case "科室IDS"
            strJson = strJson & "," & GetJsonNodeString("dept_ids", cllFilter(i)(1), Json_Text)
        Case "病人IDS"
            strJson = strJson & "," & GetJsonNodeString("pati_ids", cllFilter(i)(1), Json_Text)
        Case "主页IDS"
            strJson = strJson & "," & GetJsonNodeString("pati_pageids", cllFilter(i)(1), Json_Text)
        Case "入院开始时间"
            strJson = strJson & "," & GetJsonNodeString("adta_start_time", cllFilter(i)(1), Json_Text)
        Case "入院结束时间"
            strJson = strJson & "," & GetJsonNodeString("adta_end_time", cllFilter(i)(1), Json_Text)
        Case "出院开始时间"
            strJson = strJson & "," & GetJsonNodeString("adtd_start_time", cllFilter(i)(1), Json_Text)
        Case "出院结束时间"
            strJson = strJson & "," & GetJsonNodeString("adtd_end_time", cllFilter(i)(1), Json_Text)
        Case "费别"
            strJson = strJson & "," & GetJsonNodeString("fee_category", cllFilter(i)(1), Json_Text)
        Case "住院状态"
            strJson = strJson & "," & GetJsonNodeString("inp_status", cllFilter(i)(1), Json_num)
        Case "病人性质"
            strJson = strJson & "," & GetJsonNodeString("pati_natures", cllFilter(i)(1), Json_Text)
        Case "姓名"
            strJson = strJson & "," & GetJsonNodeString("pati_name", cllFilter(i)(1), Json_Text)
        Case "站点编号"
            strJson = strJson & "," & GetJsonNodeString("dept_nodeno", cllFilter(i)(1), Json_Text)
        Case "查询转科病人"
            strJson = strJson & "," & GetJsonNodeString("change_dept_pati", cllFilter(i)(1), Json_num)
        Case "最后一次住院"
            strJson = strJson & "," & GetJsonNodeString("is_lastpage", cllFilter(i)(1), Json_num)
        Case "险类"
            strJson = strJson & "," & GetJsonNodeString("insurance_type", cllFilter(i)(1), Json_num)
        Case "病区站点编号"
            strJson = strJson & "," & GetJsonNodeString("wararea_nodeno", cllFilter(i)(1), Json_Text)
        End Select
    Next
    strJson = strJson & "," & GetJsonNodeString("is_babyinfo", IIf(bln包含婴儿信息, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getpatpageinfbyrange", strJson, , "", mlngModule) = False Then Exit Function
    '  --出参      json
    '  --output
    '  -- code                   N 1 应答码：0-失败；1-成功
    '  -- message                C 1 应答消息： 失败时返回具体的错误信息
    '  --   page_list[]          数据组  √  √
    '  --    pati_id             N    病人id  √  √
    '  --    pati_pageid         N    主页id  √  √
    '  --    pati_name           C    姓名  √  √
    '  --    pati_sex            C    性别  √  √
    '  --    pati_age            C    年龄  √  √
    '  --    inpatient_num       C    住院号  √  √
    '  --    pati_bed            C    出院病床  √  √
    '  --    insurance_type      N    险类  √  √
    '  --    fee_category        C    费别  √  √
    '  --    pati_type           C    病人类型(普通,医保,留观)  √  √
    '  --    adta_time           C    入院时间:yyyy-mm-dd hh24:mi:ss  √  √
    '  --    adtd_time           C    出院时间:yyyy-mm-dd hh24:mi:ss  √  √
    '  --    si_inp_status       N    住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)  √  √
    '  --    pati_nature         N    病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
    '  --    pati_wardarea_id    N    当前病区id
    '  --    pati_wardarea_name  C    当前病区名称
    '  --    pati_dept_id        N    当前科室id
    '  --    pati_dept_name      C    当前科室名称
    '  --    mdlpay_mode_name    C    医疗付款方式名称
    '  --    mdlpay_mode_code    C    医疗付款方式编码
    '  --    pat_rsdpscn         C    住院医师
    '  --    pati_desc           C    病人备注
    '  --    catalog_date        C    编目日期:yyyy-mm-dd hh24:mi:ss
    '  --    insurance_num       C    医保号
    '  --    level_of_care       C    护理等级
    '  --    data_adto_sign      N    数据转出标志:0-未转出，1-已转出
    '  --    fee_auditor_sign    N    费用审核标志:0或空-未审核,1-已审核或开始审核;2-完成审核
    '  --    fee_auditor         C    费用审核人
    '  --    pre_dstat_time      C    预出院时间
    '  --    last_press_money    N    上次催款金额
    '  --    baby_list[]             婴儿信息，[数组]
    '  --      baby_num              N 1 婴儿序号
    '  --      baby_name             C 1 婴儿姓名
    '  --      baby_sex              C 1 婴儿性别
    '  --      baby_date             D 1 出生时间
    Set clldata = objServiceCall.GetJsonListValue("output.page_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            With rsPatiPageInfo
                .AddNew
                '病人ID,主页ID,姓名,性别,年龄,住院号,床号,险类,费别,病人类型,医保号,
                !病人ID = Val(Nvl(cllTemp("_pati_id")))
                !主页ID = Val(Nvl(cllTemp("_pati_pageid")))
                !姓名 = cllTemp("_pati_name")
                !性别 = cllTemp("_pati_sex")
                !年龄 = cllTemp("_pati_age")
                !住院号 = CStr(Nvl(cllTemp("_inpatient_num")))
                !床号 = cllTemp("_pati_bed")
                !险类 = Val(Nvl(cllTemp("_insurance_type")))
                !费别 = cllTemp("_fee_category")
                !病人类型 = cllTemp("_pati_type")
                !医保号 = cllTemp("_insurance_num")
                '入院时间,出院时间,住院状态,病人性质,当前病区ID,当前病区名称,当前科室id,当前科室名称,
                !入院时间 = cllTemp("_adta_time")
                !出院时间 = cllTemp("_adtd_time")
                !住院状态 = Val(Nvl(cllTemp("_si_inp_status")))
                !病人性质 = Val(Nvl(cllTemp("_pati_nature")))
                !当前病区id = Val(Nvl(cllTemp("_pati_wardarea_id")))
                !当前病区名称 = cllTemp("_pati_wardarea_name")
                !当前科室id = Val(Nvl(cllTemp("_pati_dept_id")))
                !当前科室名称 = cllTemp("_pati_dept_name")
                '医疗付款方式名称,医疗付款方式编码,住院医师,病人备注,编目日期,护理等级
                !医疗付款方式编码 = cllTemp("_mdlpay_mode_code")
                !医疗付款方式名称 = cllTemp("_mdlpay_mode_name")
                !住院医师 = cllTemp("_pat_rsdpscn")
                !病人备注 = cllTemp("_pati_desc")
                !编目日期 = cllTemp("_catalog_date")
                !护理等级 = cllTemp("_level_of_care")
                '数据转出,审核标志,审核人,预出院时间,上次催款金额
                !数据转出 = Val(Nvl(cllTemp("_data_adto_sign")))
                !审核标志 = Val(Nvl(cllTemp("_fee_auditor_sign")))
                !审核人 = cllTemp("_fee_auditor")
                !预出院时间 = cllTemp("_pre_dstat_time")
                !上次催款金额 = Val(Nvl(cllTemp("_last_press_money")))
                
                If bln包含婴儿信息 Then
                    Set rsBabyInfo = New ADODB.Recordset
                    With rsBabyInfo.Fields
                        '病人id,主页id,序号,姓名,性别,出生时间
                        .Append "病人ID", adBigInt, 18, adFldIsNullable
                        .Append "主页ID", adBigInt, 18, adFldIsNullable
                        .Append "序号", adInteger, 2, adFldIsNullable
                        .Append "姓名", adVarChar, 100, adFldIsNullable
                        .Append "性别", adVarChar, 4, adFldIsNullable
                        .Append "出生时间", adVarChar, 20, adFldIsNullable
                    End With
                    rsBabyInfo.CursorLocation = adUseClient
                    rsBabyInfo.LockType = adLockOptimistic
                    rsBabyInfo.CursorType = adOpenStatic
                    rsBabyInfo.Open
                    
                    Set cllBaby = objServiceCall.GetJsonListValue("output.page_list[" & i - 1 & "].baby_list")
                    If Not cllBaby Is Nothing Then
                        For j = 1 To cllBaby.Count
                            Set cllTemp = cllBaby(j)
                            With rsBabyInfo
                                .AddNew
                                '病人id,主页id,序号,姓名,性别,出生时间
                                !病人ID = rsPatiPageInfo!病人ID
                                !主页ID = rsPatiPageInfo!主页ID
                                !序号 = cllTemp("_baby_num")
                                !姓名 = cllTemp("_baby_name")
                                !性别 = cllTemp("_baby_sex")
                                !出生时间 = cllTemp("_baby_date")
                                .Update
                            End With
                        Next
                        If rsBabyInfo.RecordCount > 0 Then rsBabyInfo.MoveFirst
                    End If
                    cllBabyInfo.Add rsBabyInfo, "_" & !病人ID & "_" & !主页ID
                End If
                .Update
            End With
        Next
    End If
    If rsPatiPageInfo.RecordCount > 0 Then rsPatiPageInfo.MoveFirst
    
    ZlCissvr_GetPatiPageInfByRange = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetPatiMaxPageID(ByVal lng病人ID As Long, ByRef lngMaxPageID As Long, _
    Optional ByVal lng病区ID As Long, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定条件的最大主页ID
    '入参:
    '   lng病人ID=病人ID
    '   lng病区ID=病区ID
    '出参:
    '   lngMaxPageID=最大主页ID
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    lngMaxPageID = 0
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Getpatimaxpageid
    '  --功能:获取指定条件的最大主页ID
    '  --入参：Json_In:格式
    '  --    input
    '  --      pati_id           N 1 病人id
    '  --      pati_wararea_id   N   当前病区iD
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                  N 1 应答码：0-失败；1-成功
    '  --    message               C 1 应答消息：失败时返回具体的错误信息
    '  --    pati_pageid         N   主页id
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_wardarea_id", lng病区ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getpatimaxpageid", strJson, , "", mlngModule) = False Then Exit Function
    lngMaxPageID = Val(objServiceCall.GetJsonNodeValue("output.pati_pageid"))
    
    ZlCissvr_GetPatiMaxPageID = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetAllGroupAdviceIDs(ByVal str医嘱IDs As String, ByRef rsAdviceData As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱ID查询一组医嘱的所有医嘱信息
    '入参:
    '   str医嘱IDs 医嘱ID,多个英文逗号分隔
    '出参:
    '   rsAdviceData 医嘱信息：医嘱id,相关id,诊疗类别
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, i As Long, cllTemp As Collection
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    Set rsAdviceData = New ADODB.Recordset
    With rsAdviceData.Fields
        .Append "医嘱ID", adBigInt, 18, adFldIsNullable
        .Append "相关ID", adBigInt, 18, adFldIsNullable
        .Append "诊疗类别", adVarChar, 2, adFldIsNullable
    End With
    rsAdviceData.CursorLocation = adUseClient
    rsAdviceData.LockType = adLockOptimistic
    rsAdviceData.CursorType = adOpenStatic
    rsAdviceData.Open
    
    'Zl_Cissvr_Getallgroupadviceids
    '  --根据医嘱ID查询一组医嘱的所有医嘱信息
    '  --入参      json
    '  --input     根据医嘱ID查询医嘱信息
    '  --  advice_id                     C   多个医嘱ID，用“，”分隔
    '  --出参      json
    '  --output
    '  --  code       C  1  应答码：0-失败；1-成功
    '  --  message    C  1  应答消息：
    '  --  advice_list      [数组]每个医嘱信息
    '  --    advice_id                  N    id
    '  --    advice_related_id          N    相关id
    '  --    clinic_type                C    诊疗类别
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("advice_id", str医嘱IDs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getallgroupadviceids", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.advice_list")
    If clldata Is Nothing Then Exit Function
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        rsAdviceData.AddNew
            rsAdviceData!医嘱id = cllTemp("_advice_id")
            rsAdviceData!相关ID = cllTemp("_advice_related_id")
            rsAdviceData!诊疗类别 = cllTemp("_clinic_type")
        rsAdviceData.Update
    Next
    If rsAdviceData.RecordCount > 0 Then rsAdviceData.MoveFirst
    
    ZlCissvr_GetAllGroupAdviceIDs = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPatiSvr_GetCardLastChange(ByVal lng病人ID As Long, ByVal lng卡类别ID As Long, ByVal str卡号 As String, _
                ByRef int变动类型 As Integer, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取卡最后一次变动记录
    ' 入参 : lng卡类别ID-医疗卡类别ID
    '      : str卡号-传入检查的卡号
    ' 出参 : int变动类型-最后一次变动类型
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo Errhand
    If lngModule <> 0 Then mlngModule = lngModule
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardtype_id", lng卡类别ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("card_no", str卡号, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_GetCardLastChange", strJson, , , lngModule) = False Then Exit Function
    int变动类型 = Val(objServiceCall.GetJsonNodeValue("output.change_type"))

    zlPatiSvr_GetCardLastChange = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPatiSvr_CheckOutNoIsExist(ByVal lng病人ID As Long, ByVal str门诊号 As String, _
                ByRef blnUsedByOther As Boolean, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查门诊号是否被他人使用
    ' 入参 : str门诊号-传入检查的门诊号
    ' 出参 : blnUsedByOther:T:被别人使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo Errhand
    If lngModule <> 0 Then mlngModule = lngModule
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("outno", str门诊号, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_CheckOutNoIsExist", strJson, , , lngModule) = False Then Exit Function
    blnUsedByOther = Val(objServiceCall.GetJsonNodeValue("output.isexist")) <> 0

    zlPatiSvr_CheckOutNoIsExist = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPatiSvr_CheckInsNoIsExist(ByVal str医保号 As String, _
                ByRef blnUsed As Boolean, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查医保号是否存在
    ' 入参 : str医保号-传入检查的医保号
    ' 出参 : blnUsed:T:被使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo Errhand
    If lngModule <> 0 Then mlngModule = lngModule
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "," & GetJsonNodeString("outno", str医保号, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_CheckInsNoIsExist", strJson, , , lngModule) = False Then Exit Function
    blnUsed = Val(objServiceCall.GetJsonNodeValue("output.isexist")) <> 0

    zlPatiSvr_CheckInsNoIsExist = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPatiSvr_PatiRelateUpdate(ByVal int操作功能 As Integer, ByVal str关联的病人ids As String, _
                ByVal str操作员姓名 As String, ByVal str操作时间 As String, Optional ByVal lng关联ID As Long, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 处理相同病人关联
    ' 入参 : int操作功能:0-增加关联;1-取消关联;2-更新关联ID;3-入院登记自动关联
    '        lng病人ID:
    '        关联的病人ids:0-增加关联;1-取消关联;2-更新关联ID;3-入院登记自动关联
    ' 出参 : blnUsed:T:被使用
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo Errhand
    If lngModule <> 0 Then mlngModule = lngModule
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_fun", int操作功能, Json_num)
    strJson = strJson & "," & GetJsonNodeString("relate_id", lng关联ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("relate_pati_ids", str关联的病人ids, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_name", str操作员姓名, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_time", str操作时间, Json_Text)

    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_UpdatePatiRelate", strJson, , , lngModule) = False Then Exit Function

    zlPatiSvr_PatiRelateUpdate = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPatisvr_SaveMedcCard(ByVal cllCard As Collection, Optional ByRef strErrMsg_Out As String, Optional blnShowErrMsg As Boolean = False, _
    Optional int操作状态 As Integer = 0, Optional ByVal lng变动id As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 处理医疗卡变动信息，如果是挂号同时发卡，还要处理就诊信息和结算模式
    ' 入参 :cllCard-节点包含:操作类型,病人ID,卡类别ID,原卡号,医疗卡号,二维码,变动原因,密码,IC卡号,挂失方式,终止使用时间,单据号,卡费,操作时间,操作员姓名,操作员编号
    '               其中的操作类型:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
    '       int操作状态:0-正常记录;1-产生异常数据;2-只产生变动记录
    '               每项格式:array("名称",值 )
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim varTemp As Variant
    Dim strJson As String
    Dim varValue As Variant, i As Long
    
    
    On Error GoTo Errhand
    
    If cllCard Is Nothing Then Exit Function
    
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function

    strJson = ""
    strJson = strJson & "," & GetJsonNodeString("oper_state", int操作状态, Json_num, True)
    strJson = strJson & "," & GetJsonNodeString("change_id", lng变动id, Json_num, True)
    For i = 1 To cllCard.Count
        varTemp = cllCard(i)
        Select Case UCase(varTemp(0))
        Case "操作类型" '1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
            strJson = strJson & "," & GetJsonNodeString("oper_type", Val(varTemp(1)), Json_num, True)
        Case "病人ID"
            strJson = strJson & "," & GetJsonNodeString("pati_id", Val(varTemp(1)), Json_num, True)
        Case "卡类别ID"
            strJson = strJson & "," & GetJsonNodeString("card_type_id", Val(varTemp(1)), Json_num, True)
        Case "原卡号"
            strJson = strJson & "," & GetJsonNodeString("card_no_old", Trim(varTemp(1)), Json_Text)
        Case "医疗卡号", "卡号"
            strJson = strJson & "," & GetJsonNodeString("card_no", Trim(varTemp(1)), Json_Text)
        Case "二维码"
            strJson = strJson & "," & GetJsonNodeString("qrcode", Trim(varTemp(1)), Json_Text)
        Case "变动原因"
            strJson = strJson & "," & GetJsonNodeString("card_notes", Trim(varTemp(1)), Json_Text)
        Case "密码"
            strJson = strJson & "," & GetJsonNodeString("card_pwd", Trim(varTemp(1)), Json_Text)
        Case "IC卡号"
            strJson = strJson & "," & GetJsonNodeString("iccard_no", Trim(varTemp(1)), Json_Text)
        Case "终止使用时间"
            strJson = strJson & "," & GetJsonNodeString("card_use_endtime", Trim(varTemp(1)), Json_Text)
        Case "单据号"
            strJson = strJson & "," & GetJsonNodeString("fee_no", Trim(varTemp(1)), Json_Text)
        Case "卡费"
            strJson = strJson & "," & GetJsonNodeString("card_price", Val(varTemp(1)), Json_num, True)
        Case "操作时间"
            strJson = strJson & "," & GetJsonNodeString("operator_time", Trim(varTemp(1)), Json_Text)
        Case "操作员姓名"
            strJson = strJson & "," & GetJsonNodeString("operator_name", Trim(varTemp(1)), Json_Text)
        Case "操作员编号"
            strJson = strJson & "," & GetJsonNodeString("operator_code", Trim(varTemp(1)), Json_Text)
        Case "挂失方式"
            strJson = strJson & "," & GetJsonNodeString("loss_mode", Trim(varTemp(1)), Json_Text)
        End Select
    Next
    If strJson = "" Then
        strErrMsg_Out = "缺失必要的保存数据，请检查!"
        Exit Function
    End If
    strJson = Mid(strJson, 2)
    
    strJson = "{""input"":{" & strJson & "}}"
     
    If objServiceCall.CallService("Zl_Patisvr_Savemedccard", strJson, , , mlngModule, blnShowErrMsg, , , glngSys, False) = False Then Exit Function
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If blnShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    zlPatisvr_SaveMedcCard = True
    Exit Function
Errhand:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlPatisvr_CheckCardExist(ByVal lng病人ID As Long, ByVal lng卡类别ID As Long, ByVal str卡号 As String, _
                ByRef lngPatiId_Out As Long, ByRef blnExisted_out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 病人发卡检查
    ' 入参 :str卡号-不传入时，检查病人id及卡类别id存不存在卡
    '
    ' 出参 :
    '        lngPati_Out-当前使用这张卡的病人id
    '        blnExisted_out-是否存在
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo Errhand
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("card_type_id", lng卡类别ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("card_no", str卡号, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Patisvr_CheckCardExist", strJson, , , mlngModule) = False Then Exit Function
    blnExisted_out = Val(objServiceCall.GetJsonNodeValue("output.exist")) = 1
    lngPatiId_Out = Val(objServiceCall.GetJsonNodeValue("output.pati_id"))
    
    ZlPatisvr_CheckCardExist = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlPatiSvr_DelCard_Check(ByVal byt场合 As Byte, ByVal lng病人ID As Long, ByVal lng卡类别ID As Long, _
                ByVal str卡号 As String, Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String, _
                Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 处理相同病人关联
    ' 入参 :
    ' 出参 : intSendCardLimit_Out:发卡限制：0-满足发卡性质；1-病人存在相同类别的卡，不能再发卡；2-病人存在相同类别的卡，需要提醒是否重复发卡
    '        lngPati_Out-当前使用这张卡的病人id
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim varValue As Variant
    
    On Error GoTo Errhand
    If lngModule <> 0 Then mlngModule = lngModule
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("occasion", byt场合, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("gvcard_type_id", lng卡类别ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("gvcard_no", str卡号, Json_Text)

    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Patisvr_CheckReturnCard", strJson, , , lngModule, Not blnNotShowErrMsg) = False Then Exit Function
    varValue = objServiceCall.GetJsonNodeValue("output.code")
    If varValue <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlPatiSvr_DelCard_Check = True
    Exit Function
Errhand:
    If blnNotShowErrMsg Then
        strErrMsg_Out = Err.Description
    Else
        If gobjComlib.ErrCenter() = 1 Then
            Resume
        End If
    End If
End Function

Public Function zlPatiSvr_GetPatiAddrssInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int地址类别 As Integer, _
                ByRef str地址_省 As String, ByRef str地址_市 As String, ByRef str地址_县 As String, _
                ByRef str地址_乡 As String, ByRef str地址_其他 As String, ByRef str区划代码 As String, _
                Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人结构化地址信息
    ' 入参 : int地址类别:查询的地址类别：1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址；为0时表示查询所有类型的地址信息
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection
    
    On Error GoTo Errhand
    If lngModule <> 0 Then mlngModule = lngModule
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("addr_type", int地址类别, Json_num)

    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_PatiSvr_GetPatiAddrssInfo", strJson, , , lngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.addr_list")
    If Not clldata Is Nothing Then
        If clldata.Count > 0 Then
            str地址_省 = Nvl(clldata(1)("_pat_addr_state"))
            str地址_市 = Nvl(clldata(1)("_pat_addr_city"))
            str地址_县 = Nvl(clldata(1)("_pat_addr_county"))
            str地址_乡 = Nvl(clldata(1)("_pat_addr_township"))
            str地址_其他 = Nvl(clldata(1)("_pat_addr_other"))
            str区划代码 = Nvl(clldata(1)("_pat_region_code"))
        End If
    End If
    
    zlPatiSvr_GetPatiAddrssInfo = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlCISSvr_ChkOutPatiExistOrder(ByVal lng病人ID As Long, ByVal int有效天数 As Integer, _
                ByRef blnContOrder As Boolean, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 挂号将“先诊疗后结算”改为“先结算后诊疗”时要检查是未发送医嘱业务
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/4 10:49
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo Errhand
    If lngModule <> 0 Then mlngModule = lngModule
    '获取费用服务接口
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rgst_expidate", int有效天数, Json_num)

    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_CISSvr_ChkOutPatiExistOrder", strJson, , , lngModule) = False Then Exit Function
    blnContOrder = Val(Nvl(objServiceCall.GetJsonNodeValue("output.exist_order"))) <> 0
    ZlCISSvr_ChkOutPatiExistOrder = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlCissvr_GetMaxBedLen(ByRef lngBedNoMaxLen As Long, Optional ByVal bln按病区查询 As Boolean, _
    Optional ByVal lng病区ID As Long, Optional ByVal lng科室id As Long, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病区床号的最大编号长度
    '入参:
    '   bln按病区查询=True:按病区获取,False:按科室获取
    '出参:
    '   lngBedNoMaxLen=床号最大长度
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Getmaxbedlen
    '  --功能:获取病区床号的最大编号长度
    '  --入参：Json_In:格式
    '  -- input            病区id和科室ID节点二者只能传一个
    '  --   wardarea_id    N  0  病区id
    '  --   dept_id        N  0 科室ID
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               C  1 应答码：0-失败；1-成功
    '  --    message            C  1 应答消息：失败时返回具体的错误信息
    '  --    max_len            N  1 最大长度
    strJson = ""
    If bln按病区查询 Then
        strJson = strJson & "" & GetJsonNodeString("wardarea_id", lng病区ID, Json_num)
    Else
        strJson = strJson & "" & GetJsonNodeString("dept_id", lng科室id, Json_num)
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getmaxbedlen", strJson, , "", mlngModule) = False Then Exit Function
    lngBedNoMaxLen = Val(Nvl(objServiceCall.GetJsonNodeValue("output.max_len")))
    
    ZlCissvr_GetMaxBedLen = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_ExistOutAdvice(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByRef blnExistOutAdvice As Boolean, Optional ByVal lngOutAdviceId As Long, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医生是否下达了出院医嘱
    '入参:
    '出参:
    '   blnExistOutAdvice=是否存在出院医嘱
    '   lngOutAdviceId=已经开了出院医嘱的，返回医嘱的ID
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Existoutadvice
    '  --功能:医生是否下达了出院医嘱
    '  --入参：Json_In:格式
    '  -- input
    '  --   pati_id              N 1 病人ID
    '  --   pati_pageid          N 1 主页Id
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               N  1 应答码：0-失败；1-成功
    '  --    message            C  1 应答消息：失败时返回具体的错误信息
    '  --    is_out             N  1 是否已开出院医嘱 ：1-已开出院;0-未开出院
    '  --    out_advice_id      N  1 已经开了出院医嘱的，返回医嘱的ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Existoutadvice", strJson, , "", mlngModule) = False Then Exit Function
    blnExistOutAdvice = Val(Nvl(objServiceCall.GetJsonNodeValue("output.is_out"))) = 1
    If blnExistOutAdvice Then
        lngOutAdviceId = Val(Nvl(objServiceCall.GetJsonNodeValue("output.out_advice_id")))
    End If
    
    ZlCissvr_ExistOutAdvice = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetBabyData(ByVal lng病人ID As Long, _
    ByVal lng主页ID As Long, ByRef rsBabyInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查询病人主页信息
    '入参:
    '出参:
    '   cllBabyInfo 婴儿信息,字段：序号,姓名
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, i As Long, cllTemp As Collection
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    Set rsBabyInfo = New ADODB.Recordset
    With rsBabyInfo.Fields
        '序号,姓名
        .Append "序号", adInteger, 2, adFldIsNullable
        .Append "姓名", adVarChar, 100, adFldIsNullable
    End With
    rsBabyInfo.CursorLocation = adUseClient
    rsBabyInfo.LockType = adLockOptimistic
    rsBabyInfo.CursorType = adOpenStatic
    rsBabyInfo.Open
    
    'Zl_Cissvr_Getbabydata
    '  --功能:获取病人的适用类型
    '  --入参：Json_In:格式
    '  --    input
    '  --      pati_id           N 1 病人id
    '  --      pati_pageid       N   主页id
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                  N 1 应答码：0-失败；1-成功
    '  --    message               C 1 应答消息：失败时返回具体的错误信息
    '  --    baby_list[]
    '  --      baby_num            N 1 婴儿序号
    '  --      baby_name           C 1 婴儿姓名
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getbabydata", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.baby_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            With rsBabyInfo
                .AddNew
                '序号,姓名
                !序号 = cllTemp("_baby_num")
                !姓名 = cllTemp("_baby_name")
                .Update
            End With
        Next
    End If
    If rsBabyInfo.RecordCount > 0 Then rsBabyInfo.MoveFirst
    
    ZlCissvr_GetBabyData = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetPatiId(ByRef lng病人ID As Long, ByVal lng病区ID As Long, _
    Optional ByVal str住院号 As String, Optional ByVal str床号 As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据床号、住院号获取病人ID及主页ID
    '入参:
    '   str住院号、str床号-二者至少传一个
    '出参:
    '   str病人信息:病人ID:主页ID
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Getpatiid
    '  --根据床号、住院号获取病人ID及主页ID
    '  --input
    '  --   wardarea_id          N 1 当前病区id
    '  --   pati_bed             C 1 当前床号
    '  --   inpatient_num        C 1 住院号
    '  --output
    '  --    code                N 1 应答码：0-失败；1-成功
    '  --    message             C 1 应答消息： 失败时返回具体的错误信息
    '  --    pati_id             N 1 病人ID:未找到时也成功，返回0
    '  --    pati_pageid         N   主页ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("wardarea_id", lng病区ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_bed", str床号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("inpatient_num", str住院号, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getpatiid", strJson, , "", mlngModule) = False Then Exit Function
    lng病人ID = Val(Nvl(objServiceCall.GetJsonNodeValue("output.pati_id")))
    
    ZlCissvr_GetPatiId = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetInDeptInfor(ByVal byt执行功能 As Byte, ByVal byt病人来源 As Byte, _
    ByVal strNodeNo As String, ByVal bln在院病人 As Boolean, ByRef rsDept As ADODB.Recordset, _
    Optional ByVal byt查找方式 As Byte, Optional ByVal bln所有病区 As Boolean, Optional ByVal str病区ids As String, _
    Optional ByVal str服务对象 As String = "1,2,3", Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取所有有病人的住院科室
    '入参:
    '   byt执行功能=0-获取所有有病人的住院科室 1-通过科室id/病区id查找所有病人的入院科室或者病区 2-加载站点
    '   byt病人来源=1-门诊；2-住院
    '   bln在院病人=是否仅取存在在院病人的科室\病区
    '   byt查找方式=0-按科室查找 1-按病区查找
    '   bln所有病区=是否所有病区
    '   str病区ids=非所有病区时有效
    '   str服务对象=科室服务对象，多个逗号分隔,如:1,2,3（1-门诊,2-住院,3-门诊和住院）
    '出参:
    '   rsDept 科室信息,字段：ID,编码,名称,简码
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, i As Long, cllTemp As Collection
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    Set rsDept = New ADODB.Recordset
    With rsDept.Fields
        'ID,编码,名称,简码
        .Append "ID", adBigInt, 18, adFldIsNullable
        .Append "编码", adVarChar, 11, adFldIsNullable
        .Append "名称", adVarChar, 100, adFldIsNullable
        .Append "简码", adVarChar, 100, adFldIsNullable
    End With
    rsDept.CursorLocation = adUseClient
    rsDept.LockType = adLockOptimistic
    rsDept.CursorType = adOpenStatic
    rsDept.Open
    
    'Zl_Cissvr_Getinsdeptinfor
    '  --功能:获取所有有病人的住院科室
    '  --入参：Json_In:格式
    '  -- input
    '  --   opr_fun             N  1 执行功能 0-获取所有有病人的住院科室 1-通过科室id/病区id查找所有病人的入院科室或者病区 2-加载站点
    '  --   pati_source         N  1 病人来源：1-门诊；2-住院
    '  --   nodeno              C  1 站点编号
    '  --   wararea_ids         C  1 病人病区ids
    '  --   find_type           N  1 查找方式 0-按科室查找 1-按病区查找
    '  --   all_wararea         N  1 是否所有病区
    '  --   pati_in             N  1 是否在院
    '  --   dept_srvtype        C  1 服务对象,多个逗号分隔,如:1,2,3
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               C  1 应答码：0-失败；1-成功
    '  --    message            C  1 应答消息：失败时返回具体的错误信息
    '  --    dept_list[]             临床科室列表
    '  --      dept_id          N  1 科室id
    '  --      dept_code        C  1 科室编码
    '  --      dept_name        C  1 科室名称
    '  --      dept_spell       C  1 科室简码
    '  --      nodeno           C  1  站点
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("opr_fun", byt执行功能, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_source", byt病人来源, Json_num)
    strJson = strJson & "," & GetJsonNodeString("nodeno", strNodeNo, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("wararea_ids", str病区ids, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("find_type", byt查找方式, Json_num)
    strJson = strJson & "," & GetJsonNodeString("all_wararea", IIf(bln所有病区, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_in", IIf(bln在院病人, 1, 0), Json_num)
    strJson = strJson & "," & GetJsonNodeString("dept_srvtype", str服务对象, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getinsdeptinfor", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.dept_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            With rsDept
                .AddNew
                'ID,编码,名称,简码
                !ID = cllTemp("_dept_id")
                !编码 = cllTemp("_dept_code")
                !名称 = cllTemp("_dept_name")
                !简码 = cllTemp("_dept_spell")
                .Update
            End With
        Next
    End If
    If rsDept.RecordCount > 0 Then rsDept.MoveFirst
    
    ZlCissvr_GetInDeptInfor = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_PatiIsDead(ByVal lng病人ID As Long, ByRef blnDeath As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id检查病人是否已经死亡
    '入参:
    '出参:
    '   blnDeath:已死亡则为True,否则为False
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Patiisdead
    '  --功能:根据病人id检查病人是否已经死亡
    '  --入参：Json_In:格式
    '  -- input
    '  --   pati_id              N 1 病人id
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               C  1 应答码：0-失败；1-成功
    '  --    message            C  1 应答消息：失败时返回具体的错误信息
    '  --    isdeath            N  1 是否已经死亡：1-是的;0-否
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Patiisdead", strJson, , "", mlngModule) = False Then Exit Function
    blnDeath = Val(Nvl(objServiceCall.GetJsonNodeValue("output.isdeath"))) = 1
    
    ZlCissvr_PatiIsDead = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetPatiChangeStopInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng病区ID As Long, ByVal lng科室id As Long, ByVal str终止原因 As String, _
    ByRef str终止原因_Out As String, ByRef str终止时间_Out As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id、主页id、科室及病区id，获取病人变动的终止信息(终止时间、终止原因等）
    '入参:
    '   str终止原因=终止原因:多个用逗号分离,如:3,15,10,1
    '出参:
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Getpatichangerec
    '  --功能:根据条件获取病人变动信息
    '  --入参：Json_In:格式
    '  -- input
    '  --   query_type           N 1 查询方式：0-根据指定条件查询病人变动信息，1-仅查询病人某次住院的转科信息
    '  --   pati_id              N 1 病人id
    '  --   pati_pageid          N 1 主页id
    '  --   pati_wararea_id      N 0 病区id
    '  --   pati_dept_id         N 0 科室id
    '  --   start_reasons        C 0 开始原因s:多个用逗号分离,如:3,15,10,1
    '  --   stop_reasons         C 0 终止原因s:多个用逗号分离,如:3,15,10,1
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               C  1 应答码：0-失败；1-成功
    '  --    message            C  1 应答消息：失败时返回具体的错误信息
    '  --    change_list[]
    '  --        start_time         C 1 开始时间:yyyy-mm-dd hh24:mi:ss
    '  --        start_reason       N 1 开始原因
    '  --        dept_name          C 1 部门名称
    '  --        stop_time          C 1 终止时间:yyyy-mm-dd hh24:mi:ss
    '  --        stop_reason        N 1 终止原因
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("query_type", 0, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_wararea_id", lng病区ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_dept_id", lng科室id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("stop_reasons", str终止原因, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getpatichangerec", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.change_list")
    If clldata Is Nothing Then Exit Function
    If clldata.Count = 0 Then ZlCissvr_GetPatiChangeStopInfo = True: Exit Function
    
    str终止时间_Out = clldata(1)("_stop_time")
    str终止原因_Out = clldata(1)("_stop_reason")
    
    ZlCissvr_GetPatiChangeStopInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlCissvr_PatiExistMemo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByRef blnIsExist_Out As Boolean, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id及主页id检查是否存在备注信息
    '入参
    '出参:
    '   blnIsExist_Out-存在的，返回true,否则返回False
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
        
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Patiexistmemo
    '  --功能:根据病人id及主页id检查是否存在备注信息
    '  --入参：Json_In:格式
    '  -- input
    '  --   pati_id              N 1 病人id
    '  --   pati_pageid          N 1 主页id
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               C  1 应答码：0-失败；1-成功
    '  --    message            C  1 应答消息：失败时返回具体的错误信息
    '  --    exist              N  1 是否存在备注：1-是的;0-否
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Patiexistmemo", strJson, , "", mlngModule) = False Then Exit Function
    blnIsExist_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.exist")))
    
    zlCissvr_PatiExistMemo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCissvr_GetPatiPageExtendInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str信息名 As String, ByRef cllValue As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id及主页id获取病案主页从表信息
    '入参
    '   str信息名=信息名：多个用逗号
    '出参:
    '   cllValue-信息值集合，成员:Array(信息名,信息值)=cllValue(信息名)
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, clldata As Collection, i As Long, cllTemp As Collection
    Dim strKey As String, strValue As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    
    Set cllValue = New Collection
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Getpatipageextinfo
    '  --功能:根据病人id及主页id获取病案主页从表信息
    '  --入参：Json_In:格式
    '  --    input
    '  --      pati_id           N 1 病人id
    '  --      pati_pageid       N   主页id
    '  --      info_names        C 1 信息名：多个用逗号
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                  N 1 应答码：0-失败；1-成功
    '  --    message               C 1 应答消息：失败时返回具体的错误信息
    '  --    pati_Name             C 1 姓名
    '  --    ext_list[]            C   病案从表信息列表
    '  --      info_name           C 1 信息名
    '  --      info_value          C 1 信息值
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("info_names", str信息名, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getpatipageextinfo", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.ext_list")
    If clldata Is Nothing Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        strKey = cllTemp("_info_name")
        strValue = cllTemp("_info_value")
        cllValue.Add Array(strKey, strValue), strKey
    Next
    
    zlCissvr_GetPatiPageExtendInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlCisSvr_UpdateOutMedRec(ByVal lng病人ID As Long, ByVal str门诊号 As String, _
                ByVal str登记时间 As String, Optional ByVal str病案类型 As String = "一般", _
                Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新门诊病案记录
    '入参:cllOutMedRec-门诊病案数据集:array(名称,值)
    '                名称包含（病人id,.病案号(门诊号),建立日期,病案类别,存储状态,存放位置)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strServiceName  As String
    Dim objServiceCall As Object
    
    On Error GoTo Errhand
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Updateoutmedrecord
    '  --功能：更新门诊档案信息
    '  --入参：json格式
    '  --Input
    '  --   pati_id               N  1 病人id
    '  --   mr_no                 C  1 病案号
    '  --   outpatient_num        C  1 门诊号
    '  --   create_date           C  1 建立日期
    '  --   mr_type               C  1 病案类别
    '  --   strg_status           C  1 存储状态
    '  --   strgloc               C  1 存放位置
    '  --出参：json格式
    '  --Json_Out
    '  --   code                  C  1  应答码：0-失败；1-成功
    '  --   message               C  1  应答消息：失败时返回具体的错误信息
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("mr_no", str门诊号, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("create_date", str登记时间, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("mr_type", str病案类型, Json_Text)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_CisSvr_UpdateOutMedRecord"
    If objServiceCall.CallService(strServiceName, strJson, , strFrmCaption, lngModule) = False Then Exit Function

    zlCisSvr_UpdateOutMedRec = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

    
Public Function zlCissvr_UpdateOutVitalSigns(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, _
                ByVal str病人体征 As String, ByVal str操作员 As String, _
                Optional ByVal strFrmCaption As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 修改病人的生命体征信息
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/25 11:53
    '---------------------------------------------------------------------------------------
    Dim strJson As String, strServiceName  As String
    Dim objServiceCall As Object
    
    On Error GoTo Errhand
    If lngModule <> 0 Then mlngModule = lngModule
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Updateoutvitalsigns
    '--    input
    '--      pati_id           N  1 病人ID
    '--      rgst_id           N  1 挂号ID
    '--      pat_vsign         C  1 体征。格式为：项目ID1|项目值1|单位1,项目ID2|项目值2|单位2
    '--      operator_name     C    操作员姓名
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rgst_id", lng挂号ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pat_vsign", str病人体征, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("operator_name", str操作员, Json_Text)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "zl_Cissvr_UpdateOutVitalSigns"
    If objServiceCall.CallService(strServiceName, strJson, , strFrmCaption, lngModule) = False Then Exit Function
    
    zlCissvr_UpdateOutVitalSigns = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlPatiSvr_ChkCardChangeValid(ByVal cllCheckData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医疗卡变动前检查数据的合法性
    '入参:cllCheckData-格式为Array(保存项名称,保存项值)
    '                   保存的项目名称包含:操作状态,病人ID,卡类别ID,卡号,新卡号
    '                    操作状态:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, strServiceName  As String
    Dim objServiceCall As Object
    Dim strErrMsg As String
    

    On Error GoTo errHandle
    

    If GetServiceCall(objServiceCall, True) = False Then Exit Function
    
    If cllCheckData Is Nothing Then
        strErrMsg = "未传入需要检查的必要条件，请检查!"
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strJson = ""
    For i = 1 To cllCheckData.Count
        Select Case UCase(cllCheckData(i)(0))
        Case "操作状态"
            strJson = strJson & "," & GetJsonNodeString("oper_state", Val(cllCheckData(i)(1)), Json_num)
        Case "病人ID"
            strJson = strJson & "," & GetJsonNodeString("pati_id", Val(cllCheckData(i)(1)), Json_num)
        Case "卡类别ID"
            strJson = strJson & "," & GetJsonNodeString("cardtype_id", Val(cllCheckData(i)(1)), Json_num)
        Case "卡号"
            strJson = strJson & "," & GetJsonNodeString("cardno", cllCheckData(i)(1), Json_Text)
        Case "新卡号"
            strJson = strJson & "," & GetJsonNodeString("new_cardno", cllCheckData(i)(1), Json_Text)
        End Select
    Next
    If strJson = "" Then
        strErrMsg = "未传入需要检查的必要条件，请检查!"
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'Zl_PatiSvr_ChkCardChangeValid
    '    input
    '    oper_statu  N   1   操作状态:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
    '    cardtype_id N   1   卡类别ID
    '    cardno  C   1   卡号：发卡、补卡及换卡等的卡号或其他操作的原始卡号
    '    newcardno   C       新卡号:换卡时的新卡号
    '    pati_id N   1   病人ID
    
    strJson = Mid(strJson, 2)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    strServiceName = "Zl_PatiSvr_ChkCardChangeValid"
    If objServiceCall.CallService(strServiceName, strJson, , "", glngModul, True, , , , True) = False Then Exit Function
    '    output
    '        code    C   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：失败时返回具体的错误信息
    ZlPatiSvr_ChkCardChangeValid = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlPatiSvr_GetNextID(ByVal strTableName As String, ByRef lngID_Out As Long, _
    Optional blnShowErrMsg As Boolean = True, Optional ByRef strErrMsg_Out As String, Optional ByVal str字段名 As String = "ID") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取下一个序列值
    '入参:
    '出参:strErrMsg_Out-错误信息
    '     lngID_Out-返回的一下序列值
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strServiceName   As String
    Dim objServiceCall As Object
    Dim intReturn As Integer
  
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg_Out = "连接病人域服务失败，无法获取有效的序列!"
        If blnShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
  
    'input
    '  table_name    C  1 表名
    '  col_name      C  1 字段名  序列名称不一定是ID，例如记录ID
    ' 出参:
    '  output
    '  next_id      N   1  序列
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("table_name", strTableName, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("col_name", str字段名, Json_Text)
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    strServiceName = "zl_PatiSvr_GetNextID"
    If objServiceCall.CallService(strServiceName, strJson, , "", mlngModule, False) = False Then Exit Function
    '    output
    '        code    N   1   应答码：0-失败；1-成功
    '        message C   1   "应答消息：失败时返回具体的错误信息
    '        next_no C       下一个号码

    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "不能获取" & strTableName & "的序列值，请检查！"
        End If
        If blnShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lngID_Out = Val(Nvl(objServiceCall.GetJsonNodeValue("output.next_id")))
    zlPatiSvr_GetNextID = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zl_PatiSvr_DelCardChangeInfo(ByVal lng病人ID As Long, ByVal lng变动id As Long, lngCardTypeID As Long, ByVal strCardNO As String, _
    Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除医疗卡变动记录
    '入参
    '     blnShowErrMsg-是否显示错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strServiceName   As String
    Dim objServiceCall As Object
    Dim intReturn As Integer, strErrMsg As String
      
    If blnShowErrMsg Then On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "连接病人域服务失败，无法获取有效的变动信息!"
        If Not blnShowErrMsg Then Err.Raise -1001, strErrMsg, strErrMsg: Exit Function
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'input
    '    change_id   N   1   变动id
    '    cardtype_id N   1   卡类别id
    '    cardno  C   1   卡号
    '    pati_id N   1   病人ID

    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("change_id", lng变动id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardtype_id", lngCardTypeID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("cardno", strCardNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    
    strServiceName = "zl_PatiSvr_DelCardChangeInfo"
    If objServiceCall.CallService(strServiceName, strJson, , "", mlngModule, False) = False Then
        Exit Function
    End If
    '    output
    '       code    C   1   应答码：0-失败；1-成功
    '       message C   1   应答消息：失败时返回具体的错误信息
    '       deposit_id  N   1   预交ID
    '       balance_id  N   1   结帐ID
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        mstrErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If mstrErrMsg = "" Then
            mstrErrMsg = "变动记录清除失败，请检查！"
        End If
        If blnShowErrMsg Then MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    zl_PatiSvr_DelCardChangeInfo = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zl_PatiSvr_ConfirmCardChange(ByVal lng病人ID As Long, ByVal lng变动id As Long, Optional blnShowErrMsg As Boolean, Optional ByVal cllSendCard As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人医疗卡变动确认
    '入参
    '     blnShowErrMsg-是否显示错误信息
    '     cllSendCard- -返回卡数据集,格式:array(名称,值)
    '         |-操作类型,病人ID,卡类别ID,原卡号,医疗卡号,二维码,变动原因,密码,IC卡号,挂失方式,终止使用时间,单据号,卡费,操作时间,操作员姓名,操作员编号
    '         操作类型:1-发卡(或11绑定卡);2-换卡;3-补卡(13-补卡停用);4-退卡(或14取消绑定); ５-密码调整(只记录);6-挂失(16取消挂失),7-终止时间调整
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, strServiceName   As String
    Dim objServiceCall As Object, varData As Variant
    Dim intReturn As Integer, strErrMsg As String
    Dim i As Long
    
    If blnShowErrMsg Then On Error GoTo errHandle
    
    If GetServiceCall(objServiceCall, blnShowErrMsg) = False Then
        strErrMsg = "连接病人域服务失败，无法确认医疗卡变动!"
        If Not blnShowErrMsg Then Err.Raise -1001, strErrMsg, strErrMsg: Exit Function
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'input
    '    change_id   N   1   变动id
    '    pati_id N   1   病人ID
    '    card_no C       医疗卡号
    '    card_notes  C       变动原因
    '    card_pwd    C       密码
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("change_id", lng变动id, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    If Not cllSendCard Is Nothing Then
        For i = 1 To cllSendCard.Count
            varData = cllSendCard(i)
            Select Case UCase(varData(0))
            Case "卡类别ID"
                strJson = strJson & "," & GetJsonNodeString("cardtype_id", Val(varData(1)), Json_num, True)
            Case "医疗卡号"
                strJson = strJson & "," & GetJsonNodeString("card_no", varData(1), Json_Text)
            Case "变动原因"
                strJson = strJson & "," & GetJsonNodeString("card_notes", varData(1), Json_Text)
            Case "密码"
                strJson = strJson & "," & GetJsonNodeString("card_pwd", varData(1), Json_Text)
            
            End Select
        Next
    End If
    strJson = "{" & GetNodeString("input") & ":{" & strJson & "}}"
    
    
    strServiceName = "zl_PatiSvr_ConfirmCardChange"
    If objServiceCall.CallService(strServiceName, strJson, , "", mlngModule, False) = False Then
        Exit Function
    End If
    '    output
    '       code    C   1   应答码：0-失败；1-成功
    '       message C   1   应答消息：失败时返回具体的错误信息
    '       deposit_id  N   1   预交ID
    '       balance_id  N   1   结帐ID
    
    intReturn = objServiceCall.GetJsonNodeValue("output.code")
    If intReturn = 0 Then
        mstrErrMsg = objServiceCall.GetJsonNodeValue("output.message")
        If mstrErrMsg = "" Then
            mstrErrMsg = "变动记录确认失败，请检查！"
        End If
        If blnShowErrMsg Then MsgBox mstrErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    zl_PatiSvr_ConfirmCardChange = True
    Exit Function
errHandle:
    mlngErrNum = Err.Number: mstrSource = Err.Source: mstrErrMsg = Err.Description
    If blnShowErrMsg = False Then
        Err.Raise mlngErrNum, mstrSource, mstrErrMsg: Exit Function
    End If
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlCisSvr_UpddteAutoAccountSign(ByVal lng病人ID As Long, _
    ByVal str主页IDs As String, ByVal bln禁止自动记帐 As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新自动记帐标记
    '入参:
    '   str主页IDs=主页ID,多个逗号分隔
    '   bln禁止自动记帐=True:禁止自动记帐，False:允许自动记帐
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If str主页IDs = "" Then ZlCisSvr_UpddteAutoAccountSign = True: Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Setautocalcsign
    '  --功能：更新自动记帐标记
    '  --入参：json格式
    '  --Input
    '  --   pati_id               N  1 病人id
    '  --   pati_pageids          C  1 主页id,多个用英文逗号分隔
    '  --   auto_account_sign     N  1 自动记帐标志：0-允许自动记帐,1-禁止自动记帐
    '  --出参：json格式
    '  --Json_Out
    '  --   code                  C  1  应答码：0-失败；1-成功
    '  --   message               C  1  应答消息： 失败时返回具体的错误信息
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageids", str主页IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("auto_account_sign", IIf(bln禁止自动记帐, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Setautocalcsign", strJson, , , lngModule) = False Then Exit Function
    ZlCisSvr_UpddteAutoAccountSign = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCisSvr_UpdatePatiAuditInfo(ByVal lng病人ID As Long, _
    ByVal lng主页ID As Long, ByVal byt审核标记 As Byte, ByVal bln取消审核 As Boolean, _
    Optional ByVal str审核人 As String, Optional ByVal str审核说明 As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新病人审核信息
    '入参:
    '   byt审核标记=审核标记：0或空-未审核,1-已审核或开始审核;2-完成审核
    '   bln取消审核=是否取消审核：1-取消审核,0-审核
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If lng病人ID = 0 Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Updatepatiauditinfo
    '  --功能：更新病人审核信息
    '  --入参：json格式
    '  --Input
    '  --   pati_id               N  1 病人id
    '  --   pati_pageid           N  1 主页id
    '  --   audit_sign            N  1 审核标记：0或空-未审核,1-已审核或开始审核;2-完成审核
    '  --   auditor               C  0 审核人
    '  --   audit_desc            C  0 审核说明
    '  --   cancel_audit          N  0 是否取消审核：1-取消审核,0-审核
    '  --出参：json格式
    '  --Json_Out
    '  --   code                  C  1  应答码：0-失败；1-成功
    '  --   message               C  1  应答消息： 失败时返回具体的错误信息
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("audit_sign", byt审核标记, Json_num)
    strJson = strJson & "," & GetJsonNodeString("auditor", str审核人, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("audit_desc", str审核说明, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("cancel_audit", IIf(bln取消审核, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Updatepatiauditinfo", strJson, , , lngModule) = False Then Exit Function
    ZlCisSvr_UpdatePatiAuditInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCisSvr_UpdateInpatientExtendInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str更新信息 As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改病案主页从表相关信息
    '入参:
    '   str更新信息=格式：信息名:信息值,信息名:信息值,...
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, strJsonList As String
    Dim var更新信息 As Variant, i As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If lng病人ID = 0 Or str更新信息 = "" Then Exit Function
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Updateinpatiextinfo
    '  --功能:修改病案主页从表相关信息
    '  --入参：Json_In:格式
    '  --    input
    '  --    pati_id             N  1  病人id
    '  --    pati_pageid         N  1  主页Id
    '  --    item_list[]               列表
    '  --      info_name         C  1  信息名
    '  --      info_value        C  1  修改的信息值
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                N   1   应答码：0-失败；1-成功
    '  --    message             C   1   应答消息：失败时返回具体的错误信息
    strJsonList = "": var更新信息 = Split(str更新信息, ",")
    For i = 0 To UBound(var更新信息)
        strJson = ""
        strJson = strJson & "" & GetJsonNodeString("info_name", Split(var更新信息(i) & ":", ":")(0), Json_Text)
        strJson = strJson & "," & GetJsonNodeString("info_value", Split(var更新信息(i) & ":", ":")(1), Json_Text)
        strJsonList = ",{" & strJson & "}"
    Next
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & ",""item_list"":[" & Mid(strJsonList, 2) & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Updateinpatiextinfo", strJson, , , lngModule) = False Then Exit Function
    ZlCisSvr_UpdateInpatientExtendInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetMedicalGroupID(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng开单科室ID As Long, ByVal str开单人 As String, _
    ByVal dt发生时间 As Date, ByRef lng医疗小组ID As Long, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据条件获取对应的医疗小组ID
    '入参:
    '   dt发生时间=费用发生时间
    '出参:
    '   lng医疗小组id=获取到的医疗小组ID
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Getmedicalgroupid
    '  --功能:根据病人变动记录来获取对应的医疗小组ID
    '  --入参：Json_In:格式
    '  --    input
    '  --      pati_id           N 1 病人id
    '  --      pati_pageid       N 1 主页id
    '  --      plcdept_id        N 1 开单科室ID
    '  --      placer            C 1 开单人
    '  --      occur_time        C 1 发生时间:yyyy-mm-dd hh24:mi:ss
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                  N 1 应答码：0-失败；1-成功
    '  --    message               C 1 应答消息：失败时返回具体的错误信息
    '  --    group_id              N 0 医疗小组id
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("plcdept_id", lng开单科室ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("placer", str开单人, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("occur_time", Format(dt发生时间, "yyyy-MM-dd HH:mm:ss"), Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getmedicalgroupid", strJson, , "", mlngModule) = False Then Exit Function
    lng医疗小组ID = Val(Nvl(objServiceCall.GetJsonNodeValue("output.group_id")))
    
    ZlCissvr_GetMedicalGroupID = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_CheckNotExcuteItem(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal int婴儿序号 As Integer, ByVal byt费用来源 As String, _
    ByRef strNotExcuteInfo As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人信息获取医技未执行的项目
    '入参:
    '   int婴儿序号=婴儿序号:-1表示不区分;0-母亲的;>0具体婴儿费用
    '   byt费用来源=1-门诊;2-住院;4-体检
    '出参:
    '   strNotExcuteInfo=未执行的项目信息
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    strNotExcuteInfo = ""
    'Zl_Cissvr_Checkpatiexecute
    '  --功能：根据病人信息获取医技未执行的项目
    '  --入参：Json_In:格式
    '  --  input
    '  --     pati_id              N 1 病人ID
    '  --     pati_pageid          N 1 主页ID
    '  --     baby_num             N 0 婴儿序号:-1表示不区分;0-母亲的;>0具体婴儿费用
    '  --     fee_source           N 1 费用来源:1-门诊;2-住院;4-体检
    '  --     check_type           N 0 检查类型，null/0-表示检查检查其他执行项目，1-表示检查未发血
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                  N 1 应答吗：0-失败；1-成功
    '  --    message               C 1 应答消息：失败时返回具体的错误信息
    '  --    isexist               N 1 是否存在: 1-存在;0-不存在
    '  --    notexecute_infor      C 1 未执行的项目信息,isexist=1时返回
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("baby_num", int婴儿序号, Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_source", byt费用来源, Json_num)
    strJson = strJson & "," & GetJsonNodeString("check_type", 0, Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Checkpatiexecute", strJson, , "", mlngModule) = False Then Exit Function
    If Val(Nvl(objServiceCall.GetJsonNodeValue("output.isexist"))) <> 0 Then
        strNotExcuteInfo = objServiceCall.GetJsonNodeValue("output.notexecute_infor")
    End If
    ZlCissvr_CheckNotExcuteItem = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCisSvr_AuditAdviceCharge(ByVal lng医嘱ID As Long, ByVal bln审核 As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对指定医嘱进行费用审核完成
    '入参:
    '   lng医嘱ID=医嘱ID
    '   bln审核=是否审核:True-审核;False-取消审核
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Auditadvicecharge
    '  --功能:对指定医嘱进行费用审核完成
    '  --入参：Json_In:格式
    '  -- input
    '  --   advice_id            N 1 医嘱id
    '  --   verfy_statu          N 1 审核状态:1-审核;0-取消审核
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               N  1 应答码：0-失败；1-成功
    '  --    message            C  1 应答消息：失败时返回具体的错误信息
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("advice_id", lng医嘱ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("verfy_statu", IIf(bln审核, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Auditadvicecharge", strJson, , , lngModule) = False Then Exit Function
    
    ZlCisSvr_AuditAdviceCharge = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetAdviceDefinedInfo(ByRef rsAdviceDefinedInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱内容定义的相关信息
    '入参:
    '出参:
    '   rsAdviceDefinedInfo 医嘱内容定义信息：诊疗类别,医嘱内容
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, i As Long, cllTemp As Collection
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
     
    Set rsAdviceDefinedInfo = New ADODB.Recordset
    With rsAdviceDefinedInfo
        '诊疗类别,医嘱内容
        .Fields.Append "诊疗类别", adVarChar, 2, adFldIsNullable
        .Fields.Append "医嘱内容", adLongVarChar, 500, adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    'Zl_Cissvr_Getadvicedefinedinfo
    '  --功能：获取医嘱内容定义的相关信息
    '  --入参：无
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                      N 1 应答吗：0-失败；1-成功
    '  --    message                   C 1 应答消息：失败时返回具体的错误信息
    '  --    item_list[]
    '  --      clinic_type             C 1 诊疗类别
    '  --      advice_note             C 1 医嘱内容
    strJson = ""
    
    If objServiceCall.CallService("Zl_Cissvr_Getadvicedefinedinfo", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            With rsAdviceDefinedInfo
                .AddNew
                '诊疗类别,医嘱内容
                !诊疗类别 = cllTemp("_clinic_type")
                !医嘱内容 = cllTemp("_advice_note")
                .Update
            End With
        Next
    End If
    If rsAdviceDefinedInfo.RecordCount > 0 Then rsAdviceDefinedInfo.MoveFirst
    
    ZlCissvr_GetAdviceDefinedInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetAdviceOperInfo(ByVal str医嘱IDs As String, ByVal byt操作类型 As Byte, _
    ByRef rsAdviceOper As ADODB.Recordset, Optional ByVal bln最后一次 As Boolean, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱操作的最后一次信息(停用，启用等时间)
    '入参:
    '   str医嘱IDs=医嘱id,多个用逗号分隔
    '   byt操作类型=操作类型:1-新开；2-校对疑问；3-校对通过；4-作废；5-重整；
    '                       6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果；
    '                       11-审核通过；12-审核未通过；13-实习医师停嘱后待审核；14-血库接收；15-血库审核通过；
    '                       16-血库配血拒绝；17-血库停止配血；18-输血初审通过待签发；9-输血初审回退；20-输血医嘱标记未用
    '   bln最后一次=是否仅取最后一次操作信息
    '出参:
    '   rsAdviceOper=医嘱操作信息：医嘱ID,操作时间,操作说明
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, clldata As Collection
    Dim i As Long, cllTemp As Collection
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    Set rsAdviceOper = New ADODB.Recordset
    With rsAdviceOper
        '医嘱ID,操作时间,操作说明
        .Fields.Append "医嘱ID", adBigInt, 18, adFldIsNullable
        .Fields.Append "操作时间", adVarChar, 20, adFldIsNullable
        .Fields.Append "操作说明", adLongVarChar, 500, adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    'Zl_Cissvr_Getadviceoperinfo
    '  --功能:获取医嘱相关的操作说明
    '  --入参：Json_In:格式
    '  -- input
    '  --   advice_id            C 1 医嘱id,多个用逗号分隔
    '  --   oper_type            N 1 操作类型:1-新开；2-校对疑问；3-校对通过；4-作废；5-重整；
    '  --                                     6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果；
    '  --                                     11-审核通过；12-审核未通过；13-实习医师停嘱后待审核；14-血库接收；15-血库审核通过；
    '  --                                     16-血库配血拒绝；17-血库停止配血；18-输血初审通过待签发；9-输血初审回退；20-输血医嘱标记未用
    '  --   oper_last            N 1 是否取最后一次操作：1-取最后一次,0-取所有
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code               N  1 应答码：0-失败；1-成功
    '  --    message            C  1 应答消息：失败时返回具体的错误信息
    '  --    oper_list          C    数组
    '  --      advice_id        N  1 医嘱ID
    '  --      oper_time        C  1 操作时间:yyyy-mm-dd hh24:mi:ss
    '  --      oper_note        C  1 操作说明
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("advice_id", str医嘱IDs, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("oper_type", byt操作类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("oper_last", IIf(bln最后一次, 1, 0), Json_num)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getadviceoperinfo", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.oper_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            With rsAdviceOper
                .AddNew
                '医嘱ID,操作时间,操作说明
                !医嘱id = cllTemp("_advice_id")
                !操作时间 = cllTemp("_oper_time")
                !操作说明 = cllTemp("_oper_note")
                .Update
            End With
        Next
    End If
    If rsAdviceOper.RecordCount > 0 Then rsAdviceOper.MoveFirst
    
    ZlCissvr_GetAdviceOperInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ZlCissvr_GetGroupAdviceInfo(ByVal str医嘱IDs As String, ByRef rsAdvice As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一并给药的医嘱内容
    '入参:
    '   str医嘱IDs=医嘱id,多个用逗号分隔
    '出参:
    '   rsAdvice=医嘱操作信息：医嘱ID,医嘱内容
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, clldata As Collection
    Dim i As Long, cllTemp As Collection
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    Set rsAdvice = New ADODB.Recordset
    With rsAdvice
        '医嘱ID,医嘱内容
        .Fields.Append "医嘱ID", adBigInt, 18, adFldIsNullable
        .Fields.Append "医嘱内容", adLongVarChar, 1000, adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    'Zl_Cissvr_Getgroupadviceinfo
    '  --功能:获取一并给药的医嘱内容
    '  --入参：Json_In:格式
    '  -- input
    '  --   advice_ids           C   1 医嘱ID，多个用英文的逗号分隔
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                C  1 应答码：0-失败；1-成功
    '  --    message             C  1 应答消息：失败时返回具体的错误信息
    '  --    advice_list              [数组]每个医嘱信息
    '  --      advice_id         N   医嘱id，传入医嘱ID
    '  --      advice_note       C   医嘱内容
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("advice_ids", str医嘱IDs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Getgroupadviceinfo", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.advice_list")
    If Not clldata Is Nothing Then
        For i = 1 To clldata.Count
            Set cllTemp = clldata(i)
            With rsAdvice
                .AddNew
                '医嘱ID,医嘱内容
                !医嘱id = cllTemp("_advice_id")
                !医嘱内容 = cllTemp("_advice_note")
                .Update
            End With
        Next
    End If
    If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
    
    ZlCissvr_GetGroupAdviceInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function zlCisSvr_GetAdviceidsFromDiag(ByVal str诊断IDs As String, ByRef strAdviceIds_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据诊断ID获取对应的医嘱Ids
    '入参:str诊断IDs-诊断ID,多个逗号分离
    '出参:strAdviceIds_Out-返回医嘱IDs
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim intMode As Integer, strMessage As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'input           根据医嘱ID查询医嘱发送信息
    '   diag_ids    C   1   诊断id,多个用逗号分离

    strJson = strJson & "" & GetJsonNodeString("diag_ids", str诊断IDs, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_CisSvr_GetAdviceidsFromDiag", strJson, , "", mlngModule) = False Then Exit Function
    'output
    '    code    C   1   应答码：0-失败；1-成功
    '    message C   1   "应答消息：失败时返回具体的错误信息
    '    advice_ids  C   1   医嘱ids,多个用逗号分离
    
    intMode = Val(Nvl(objServiceCall.GetJsonNodeValue("output.code")))
    If intMode > 0 Then
        strMessage = Nvl(objServiceCall.GetJsonNodeValue("output.message"))

       If MsgBox(strMessage, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
           Exit Function
       End If
       Exit Function
    End If
    strAdviceIds_Out = Nvl(objServiceCall.GetJsonNodeValue("output.advice_ids"))
    zlCisSvr_GetAdviceidsFromDiag = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function



Public Function zlPatisvr_GetPatiBlackInfo(ByVal lng病人ID As Long, ByVal str场合 As String, Optional ByVal str预约方式 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人黑名单信息
    '入参:
    '   str场合-应用场合:预约，挂号，结帐，入院，出院
    '出参:
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim intMode As Integer, strMessage As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    'Zl_Cissvr_Checkpatiexecute
    '  --功能：根据病人信息获取医技未执行的项目
    '  --入参：Json_In:格式
    '  --  input
    '  --      pati_id           N 1 病人id
    '  --      occasion          C 1 应用场合:预约，挂号，结帐，入院，出院
    '  --      appt_mode_name    C   预约方式
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                  N 1 应答吗：0-失败；1-成功
    '  --    message               C 1 应答消息：失败时返回具体的错误信息
    '  --    isexist               N 1 是否存在: 1-存在;0-不存在
    '  --    notexecute_infor      C 1 未执行的项目信息,isexist=1时返回
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("occasion", str场合, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("appt_mode_name", str预约方式, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_Patisvr_GetPatiBlackInfo", strJson, , "", mlngModule) = False Then Exit Function
    intMode = Val(Nvl(objServiceCall.GetJsonNodeValue("output.tip_mode")))
    If intMode > 0 Then
        strMessage = Nvl(objServiceCall.GetJsonNodeValue("output.tip_message"))
        If intMode = 1 Then
            MsgBox strMessage, vbInformation, gstrSysName
            Exit Function
        Else
            If MsgBox(strMessage, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    zlPatisvr_GetPatiBlackInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function




Public Function ZlCisSvr_UpdateAdviceExeStatus(ByVal cllAdviceDatas As Collection, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新医嘱发送的执行状态
    '入参:
    '   cllAdviceDatas(collect)-数据集，格式如下
    '     |-cllAdviceData(collect)每行明细数据集
    '        |-成员(医嘱ID,费用单号,单据性质,执行人,执行时间,执行状态,原执行状态)
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, strItem As String, strJsonList As String
    Dim cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If cllAdviceDatas Is Nothing Then ZlCisSvr_UpdateAdviceExeStatus = True: Exit Function
    If cllAdviceDatas.Count = 0 Then ZlCisSvr_UpdateAdviceExeStatus = True: Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Cissvr_Updadviceexestatus
    '  --功能：更新医嘱发送的执行状态
    '  --入参：Json_In:格式
    '  --  input
    '  --    item_list
    '  --      advice_id              N  1 医嘱ID
    '  --      bill_no                C  1 单据号
    '  --      bill_prop              N  1 记录性质
    '  --      exe_status_old         C    执行状态:多个用逗号
    '  --      exe_status             N  1 更新的执行状态
    '  --      exetr                  C  1 更新的执行人
    '  --      exe_time               C  1 更新的执行时间:yyyy-mm-dd hh24:mi:ss
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                      N 1 应答吗：0-失败；1-成功
    '  --    message                   C 1 应答消息：失败时返回具体的错误信息
    strJsonList = ""
    For i = 1 To cllAdviceDatas.Count
        Set cllTemp = cllAdviceDatas(i)
        strItem = ""
        strItem = strItem & "" & GetJsonNodeString("advice_id", cllTemp("医嘱ID"), Json_num)
        strItem = strItem & "," & GetJsonNodeString("bill_no", cllTemp("费用单号"), Json_Text)
        strItem = strItem & "," & GetJsonNodeString("bill_prop", cllTemp("单据性质"), Json_num)
        strItem = strItem & "," & GetJsonNodeString("exe_status_old", cllTemp("原执行状态"), Json_Text)
        strItem = strItem & "," & GetJsonNodeString("exe_status", cllTemp("执行状态"), Json_num)
        strItem = strItem & "," & GetJsonNodeString("exetr", cllTemp("执行人"), Json_Text)
        strItem = strItem & "," & GetJsonNodeString("exe_time", cllTemp("执行时间"), Json_Text)
        strJsonList = strJsonList & ",{" & strItem & "}"
    Next
    strJson = ""
    strJson = strJson & """item_list"":[" & Mid(strJsonList, 2) & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Cissvr_Updadviceexestatus", strJson, , , lngModule) = False Then Exit Function
    
    ZlCisSvr_UpdateAdviceExeStatus = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPatisvr_BatupdoutPativisit(ByVal cllVisitDatas As Collection, _
                Optional ByVal blnNotShowErr As Boolean, Optional ByRef strErrMsg_Out As String, _
                Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:批量更新病人的就诊状态和就诊诊室
    '入参:
    '   cllVisitDatas(collect)-数据集，格式如下
    '        |-成员(病人IDs,就诊诊室)
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, strItem As String, strJsonList As String
    Dim cllTemp As Collection, i As Long
    Dim intReturn As Integer
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule
    If cllVisitDatas Is Nothing Then zlPatisvr_BatupdoutPativisit = True: Exit Function
    If cllVisitDatas.Count = 0 Then zlPatisvr_BatupdoutPativisit = True: Exit Function
    
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Patisvr_Batupdoutpativisit
    '  --功能：批量更新病人的就诊状态和就诊诊室
    '  --入参：Json_In:格式
    '  --    input
    '  --      visit_status  N 1 就诊状态
    '  --      pati_list     C   数组
    '  --      pati_ids      C 1 病人id,多个用","分隔
    '  --      visit_room    C 1 诊室
    '  --出参: Json_Out,格式如下
    '  --  output
    '  --    code                      N 1 应答吗：0-失败；1-成功
    '  --    message                   C 1 应答消息：失败时返回具体的错误信息
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("visit_status", 1, Json_num)
    
    strJsonList = ""
    For i = 1 To cllVisitDatas.Count
        Set cllTemp = cllVisitDatas(i)
        strItem = ""
        strItem = strItem & "" & GetJsonNodeString("pati_ids", cllTemp("_pati_ids"), Json_Text)
        strItem = strItem & "," & GetJsonNodeString("visit_room", cllTemp("_visit_room"), Json_Text)
        strJsonList = strJsonList & ",{" & strItem & "}"
    Next
    
    strJson = strJson & ",""pati_list"":[" & Mid(strJsonList, 2) & "]"
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Patisvr_Batupdoutpativisit", strJson, , , lngModule, Not blnNotShowErr) = False Then Exit Function
    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    If intReturn = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out = "" Then
            strErrMsg_Out = "更新病人就诊信息失败，请检查！"
        End If
        Exit Function
    End If
    
    zlPatisvr_BatupdoutPativisit = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlDrugSvr_GetNotSendRec(ByVal byt单据类型 As Byte, _
    ByVal bln未收费 As Boolean, ByVal str费用来源 As String, _
    ByVal str开始时间 As String, ByVal str终止时间 As String, _
    ByRef rsData As ADODB.Recordset, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取未发药品记录
    '入参:
    '   byt单据类型=1-收费处方;2-记帐单处方;3-记帐表处方
    '   str费用来源=多个用英文逗号分隔:1-门诊,2-住院,4-体检
    '   str开始时间=查询开始时间:yyyy-mm-dd hh:mi:ss
    '   str终止时间=查询终止时间:yyyy-mm-dd hh:mi:ss
    '出参:
    '   rsData 字段：单据类型,处方单号,药房ID
    '       其中，单据类型:1-收费处方,2-记帐单处方,3-记帐表处方
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule

    Set rsData = New ADODB.Recordset
    With rsData.Fields
        '单据类型,处方单号,药房ID
        .Append "单据类型", adInteger, 2, adFldIsNullable
        .Append "处方单号", adVarChar, 10, adFldIsNullable
        .Append "药房ID", adBigInt, 18, adFldIsNullable
    End With
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open

    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Drugsvr_Getnotsendrec
    '  --功能：获取未发药品记录
    '  --入参：JSON格式
    '  --input
    '  --  billtypes             C  1 单据类型，多个用英文逗号分隔: 1-收费处方;2-记帐单处方;3-记帐表处方
    '  --  charge_tag            N  1 收费标志:0-未收费或记帐划价;1-已收费或记帐
'     --  fee_source            C  1 费用来源，多个用英文逗号分隔:1-门诊,2-住院,4-体检
    '  --  start_time            C  0 开始时间:yyyy-mm-dd hh:mi:ss
    '  --  end_time              C  0 结束时间:yyyy-mm-dd hh:mi:ss
    '  --出参：JSON格式
    '  --output
    '  --  code                  N 1 应答吗：0-失败；1-成功
    '  --  message               C 1 应答消息：失败时返回具体的错误信息
    '  --  item_list[]
    '  --    billtype            N  1 单据类型:1-收费处方,2-记帐单处方,3-记帐表处方
    '  --    rcp_no              C  1 处方单号
    '  --    pharmacy_id         N  1 药房ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("billtypes", byt单据类型, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("charge_tag", IIf(bln未收费, 0, 1), Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_source", str费用来源, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("start_time", str开始时间, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("end_time", str终止时间, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Drugsvr_Getnotsendrec", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        rsData.AddNew
            '单据类型,处方单号,药房ID
            rsData!单据类型 = cllTemp("_billtype")
            rsData!处方单号 = cllTemp("_rcp_no")
            rsData!药房Id = cllTemp("_pharmacy_id")
        rsData.Update
    Next
    If rsData.RecordCount > 0 Then rsData.MoveFirst
    zlDrugSvr_GetNotSendRec = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlStuffSvr_GetNotSendRec(ByVal byt单据类型 As Byte, _
    ByVal bln未收费 As Boolean, ByVal str费用来源 As String, _
    ByVal str开始时间 As String, ByVal str终止时间 As String, _
    ByRef rsData As ADODB.Recordset, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取未发料记录
    '入参:
    '   byt单据类型=1-收费发料单;2-记帐发料单;3-记帐表发料单
    '   str费用来源=多个用英文逗号分隔:1-门诊,2-住院,4-体检
    '   str开始时间=查询开始时间:yyyy-mm-dd hh:mi:ss
    '   str终止时间=查询终止时间:yyyy-mm-dd hh:mi:ss
    '出参:
    '   rsData 字段：单据类型,发料单号,库房ID
    '       其中，单据类型:1-收费发料单;2-记帐发料单;3-记帐表发料单
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, cllTemp As Collection, i As Long
    
    On Error GoTo ErrHandler
    If lngModule <> 0 Then mlngModule = lngModule

    Set rsData = New ADODB.Recordset
    With rsData.Fields
        '单据类型,发料单号,库房ID
        .Append "单据类型", adInteger, 2, adFldIsNullable
        .Append "发料单号", adVarChar, 10, adFldIsNullable
        .Append "库房ID", adBigInt, 18, adFldIsNullable
    End With
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open

    If GetServiceCall(objServiceCall) = False Then Exit Function
    'Zl_Stuffsvr_Getnotsendrec
    '  --功能：获取未发料记录
    '  --入参：JSON格式
    '  --input
    '  --  billtypes             C  1 单据类型，多个用英文逗号分隔:  1-收费发料单;2-记帐发料单;3-记帐表发料单
    '  --  charge_tag            N  1 收费标志:0-未收费或记帐划价;1-已收费或记帐
'     --  fee_source            C  1 费用来源，多个用英文逗号分隔:1-门诊,2-住院,4-体检
    '  --  start_time            C  0 开始时间:yyyy-mm-dd hh:mi:ss
    '  --  end_time              C  0 结束时间:yyyy-mm-dd hh:mi:ss
    '  --出参：JSON格式
    '  --output
    '  --  code  N  1  应答吗：0-失败；1-成功
    '  --  message C 1 应答消息：失败时返回具体的错误信息
    '  --  item_list[]
    '  --    billtype              N  1 单据类型: 1-收费发料单;2-记帐发料单;3-记帐表发料单
    '  --    stuff_no              C  1 发料单号
    '  --    warehouse_id          N  1 库房ID
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("billtypes", byt单据类型, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("charge_tag", IIf(bln未收费, 0, 1), Json_num)
    strJson = strJson & "," & GetJsonNodeString("fee_source", str费用来源, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("start_time", str开始时间, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("end_time", str终止时间, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("Zl_Stuffsvr_Getnotsendrec", strJson, , "", mlngModule) = False Then Exit Function
    Set clldata = objServiceCall.GetJsonListValue("output.item_list")
    If clldata Is Nothing Then Exit Function
    
    For i = 1 To clldata.Count
        Set cllTemp = clldata(i)
        rsData.AddNew
            '单据类型,发料单号,库房ID
            rsData!单据类型 = cllTemp("_billtype")
            rsData!发料单号 = cllTemp("_stuff_no")
            rsData!库房ID = cllTemp("_warehouse_id")
        rsData.Update
    Next
    If rsData.RecordCount > 0 Then rsData.MoveFirst
    zlStuffSvr_GetNotSendRec = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlCISSvr_GetErrBillInfo(ByVal strNo As String, _
    ByRef int同步状态 As Integer, Optional ByRef lng病人ID As Long, _
    Optional ByVal str门诊号 As String, _
    Optional ByRef cllPro As Collection, Optional ByRef cllProOther As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取医生站预约接收异常数据
    ' 入参 : strNo-挂号单号
    ' 出参 : int同步状态:-1-病人ID为同步到业务数据;1-三方支付未校验;2-校验完成未支付成功
    '        str交易信息-三方卡结算信息(Json)
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/1/3 16:30
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim clldata As Collection, i As Long
    Dim cllBalance As Collection, cllExpend As Collection
    Dim bln普通结算 As Boolean
    Dim dbl结算金额 As Double
    Dim lng结帐ID As Long, lng关联ID As Long, lng卡类别ID As Long, lng支付卡类别ID As Long
    Dim strSql As String, str扩展信息 As String, str结算信息 As String
    Dim str结算方式 As String, str交易流水号 As String, str交易说明 As String, str结算号码 As String
    Dim str结算摘要 As String, str支付卡号 As String, str交易名称 As String, str交易内容 As String
    
    On Error GoTo ErrHandler
    Set cllPro = New Collection: Set cllProOther = New Collection
    If GetServiceCall(objServiceCall) = False Then Exit Function
    'zl_CISSvr_GetErrBillInfo
    '  --功能：获取医生站预约接收异常数据
    '  --Input
    '  --   rgst_no               C  1 挂号单
    '  --Json_Out
    '  --   code                  C  1  应答码：0-失败；1-成功
    '  --   message               C  1  "应答消息"：
    '  --   snyc_status           N  1  同步状态：-1-病人ID为同步到业务数据;1-三方支付未校验;2-校验完成未支付成功
    '  --   pati_id               N     病人ID
    '  --   outpno                C     门诊号
    '  --   rgst_balance          C     交易信息:同步状态为2时返回三方交易信息
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("rgst_no", strNo, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_CISSvr_GetErrBillInfo", strJson, , "", mlngModule) = False Then Exit Function
    int同步状态 = Val(objServiceCall.GetJsonNodeValue("output.snyc_status"))
    lng病人ID = Val(Nvl(objServiceCall.GetJsonNodeValue("output.pati_id")))
    str门诊号 = Nvl(objServiceCall.GetJsonNodeValue("output.outpno"))
    If int同步状态 = 2 Then
        lng结帐ID = Val(objServiceCall.GetJsonNodeValue("output.rgst_balance.blnc_id"))
        lng关联ID = Val(objServiceCall.GetJsonNodeValue("output.rgst_balance.relation_id"))
        lng卡类别ID = Val(objServiceCall.GetJsonNodeValue("output.rgst_balance.cardtype_id"))
        
        Set cllBalance = objServiceCall.GetJsonListValue("output.rgst_balance.balance_list")
        Set cllExpend = objServiceCall.GetJsonListValue("output.rgst_balance.otherswap_list")
        For i = 1 To cllBalance.Count
            str结算方式 = cllBalance(i)("_blnc_mode")
            dbl结算金额 = Val(cllBalance(i)("_swapmoney"))
            str交易流水号 = cllBalance(i)("_swapno")
            str交易说明 = cllBalance(i)("_swapmemo")
            On Error Resume Next
                str结算号码 = "": str结算摘要 = ""
                str结算号码 = cllBalance(i)("_blnc_no")
                str结算摘要 = cllBalance(i)("_blnc_memo")
            On Error GoTo ErrHandler
            str支付卡号 = cllBalance(i)("_card_no")
            lng支付卡类别ID = Val(cllBalance(i)("_cardtype_id"))
            If lng支付卡类别ID = 0 Then bln普通结算 = True
            
            str结算信息 = str结算方式 & "," & dbl结算金额 & "," & str结算号码 & "," & str结算摘要
            strSql = zlGetRegFeeModifySQL(strNo, lng结帐ID, str结算信息, Pay_ThreePay, 2, False, i <> 1, _
                    lng关联ID, lng支付卡类别ID, str支付卡号, str交易流水号, str交易说明, bln普通结算)
            zlAddArray cllPro, strSql
        Next
        
        For i = 1 To cllExpend.Count
            str交易名称 = cllExpend(i)("_swap_name")
            str交易内容 = cllExpend(i)("_swap_note")
            str扩展信息 = str扩展信息 & "||" & str交易名称 & "|" & str交易内容
        Next
        If str扩展信息 <> "" Then
           Call zlAddThreeSwapSQLToCollection(False, lng结帐ID, lng卡类别ID, False, str支付卡号, Mid(str扩展信息, 3), cllProOther)
        End If
    End If
    
    zlCISSvr_GetErrBillInfo = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlCISSvr_DelErrBillInfo(ByVal strNo As String, _
                Optional ByVal blnShowErrMsg As Boolean, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取医生站预约接收异常数据
    ' 入参 : strNo-挂号单号
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2020/1/3 16:30
    '---------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim intReturn As Integer
    Dim strJson As String
    
    On Error GoTo ErrHandler

    If GetServiceCall(objServiceCall) = False Then Exit Function
    'zlCISSvr_DelErrBillInfo
    '  --功能：获取医生站预约接收异常数据
    '  --Input
    '  --   rgst_no               C  1 挂号单
    '  --Json_Out
    '  --   code                  C  1  应答码：0-失败；1-成功
    '  --   message               C  1  "应答消息"：
    
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("rgst_no", strNo, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_CISSvr_DelErrBillInfo", strJson, , "", mlngModule, blnShowErrMsg) = False Then Exit Function
    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    If intReturn <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlCISSvr_DelErrBillInfo = True
    Exit Function
ErrHandler:
    If blnShowErrMsg Then
        If gobjComlib.ErrCenter() = 1 Then
            Resume
        End If
        Call gobjComlib.SaveErrLog
    Else
        strErrMsg_Out = Err.Description
    End If
End Function

