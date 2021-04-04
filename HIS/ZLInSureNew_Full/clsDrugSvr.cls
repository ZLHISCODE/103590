VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************************************************************************************
'功能:所有涉及药品公共服务调用接口
'接口说明:
'  一、公共部分
'     1.GetServiceCall-获取药品服务部件
'  二、服务接口部分
'     1.ZlDrugsvr_Getadditional_Infor-获取药品的一些扩展或附加的信息，包含：用法，剂量，频次，剂型等
'     2.zlDrugSvr_Check_IsOutTakeDrug-根据病人ID和主页ID,判断该病人是否出院带药
'     3.
'
'出参:
'返回:成功返回true,否则返回False
'编制:刘兴洪
'日期:2019-08-29 19:23:18
'*********************************************************************************************************************************************

Private mobjServiceCall  As Object
Private mlngModule As Long
Private Function GetServiceCall(ByRef objServiceCall_Out As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共服务对象
    '出参:objServiceCall_Out-返回公共服务对象
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    If Not mobjServiceCall Is Nothing Then Set objServiceCall_Out = mobjServiceCall: GetServiceCall = True: Exit Function
    Err = 0: On Error Resume Next
    Set mobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
    If Err <> 0 Then
        MsgBox "部件【zlServiceCall】丢失，请与系统管理员联系，恢复该部件！", vbInformation + vbOKOnly, gstrSysName
        Err = 0: On Error GoTo 0
        Exit Function
    End If
    If mobjServiceCall.InitService(gcnOracle, gstrDbUser, glngSys, glngModul) = False Then Set mobjServiceCall = Nothing: Exit Function
    
    Set objServiceCall_Out = mobjServiceCall
    GetServiceCall = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlDrugsvr_Getadditional_Infor(ByVal strNO As String, ByVal lng费用Id As Long, _
    ByRef cllData_Out As Collection, Optional ByVal byt单据类型 As Byte = 2, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品的一些扩展或附加的信息，包含：用法，剂量，频次，剂型等
    '入参:strNo-单据号
    '    str费用Ids:费用Id,多个用逗号
    '    byt单据类型:1-收费用单;2-记帐单
    '
    '出参::cllData_Out-(频次,用法,剂型,剂量单位,执行性质)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, cllData As Collection, cllTemp As Collection
    Dim objServiceCall As Object
    Dim i As Long, j As Long
 
    On Error GoTo errHandle
    
    Set cllData_Out = New Collection
    
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input
    '        billtype                    N   1   单据类型:1 -收费处方发药  ;2- 记帐单处方发药
    '        rcp_no                  C   1   单据号
    '        rcpdtl_ids                  C       处方明细ids,目前传入的费用ID
    '    出参 json
    '    output
    '        code                    N   1   应答吗：0-失败；1-成功
    '        message                 C   1   应答消息：失败时返回具体的错误信息
    '        item_list[]                         更新数据列表[数组]
    '            rcp_no              C   1   NO
    '            rcpdtl_id               N   1   处方明细id,目前传入的费用ID
    '            frequency               C   1   频次
    '            usage               C   1   用法
    '            si_drug_form                C   1   剂型
    '            loitem_detail_measunit              C   1   剂量单位
    '            advice_exe_properties               N   1   执行性质:0~2-计价特性,3-离院带药,4-自取药
 
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("billtype", byt单据类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rcp_no", strNO, Json_Text)
    strJson = strJson & "," & GetJsonNodeString("rcpdtl_ids", lng费用Id, Json_Text)
    strJson = "{""input"":{" & strJson & "}}"

    If objServiceCall.CallService("zl_DrugSvr_GetAdditional_Infor", strJson, , "", lngModule) = False Then Exit Function
    
    Set cllData = objServiceCall.GetJsonListValue("output.item_list")
    
    For i = 1 To cllData.Count
        Set cllTemp = cllData(i)
        cllData_Out.Add cllTemp("_frequency"), "频次"
        cllData_Out.Add cllTemp("_usage"), "用法"
        cllData_Out.Add cllTemp("_si_drug_form"), "剂型"
        cllData_Out.Add cllTemp("_loitem_detail_measunit"), "剂量单位"
        cllData_Out.Add cllTemp("_advice_exe_properties"), "执行性质"
        Exit For
    Next
    If cllData_Out.Count = 0 Then
        cllData_Out.Add "", "频次"
        cllData_Out.Add "", "用法"
        cllData_Out.Add "", "剂型"
        cllData_Out.Add "", "剂量单位"
        cllData_Out.Add 0, "执行性质"
    End If
    ZlDrugsvr_Getadditional_Infor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlDrugSvr_Check_IsOutTakeDrug(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID和主页ID,判断该病人是否出院带药
    '入参:lng病人ID-病人ID
    '    lng主页Id-主页ID
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String
    Dim i As Long, j As Long
 
    On Error GoTo errHandle
    
    
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input
    '        pati_id                 N   1   病人ID
    '        pati_pageid                 N   1   主页ID
    '    出参 json
    '    output
    '        code                    N   1   应答吗：0-失败；1-成功
    '        message                 C   1   应答消息：失败时返回具体的错误信息
    '        isexist                 N   1   是否存在: 1-存在;0-不存在
    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("pati_id", lng病人ID, Json_num)
    strJson = strJson & "," & GetJsonNodeString("pati_pageid", lng主页ID, Json_num)
    strJson = "{""input"":{" & strJson & "}}"

    If objServiceCall.CallService("Zl_Cissvr_Isouttakedrug", strJson, , "", lngModule) = False Then Exit Function
    zlDrugSvr_Check_IsOutTakeDrug = objServiceCall.GetJsonNodeValue("output.isexist") = 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlPati_Merage(ByVal lng保留病人Id As Long, ByVal lng合并病人ID As Long, ByVal str合并原因 As String, ByVal str操作员姓名 As String, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:合并病人
    '入参:lng保留病人Id-保留病人Id
    '    lng合并病人ID-合并病人ID
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim strJson As String, str保留信息 As String, varData As Variant, strName As String, strBirthDate As String, strCardID As String, strSex As String
    Dim strAge As String, blnTrans As Boolean, strErrMsg As String, lng保留Id As Long, lng合并ID As Long
    Dim i As Long
 
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    
    
 gcnOracle.BeginTrans: blnTrans = True
    str保留信息 = zlDatabase.CallProcedure("zl_病人信息_MERGE", "医保执行病人信息合并", lng合并病人ID, lng保留病人Id, str合并原因, str操作员姓名, 0, Empty)
    
    If str保留信息 <> "" Then   '保留id ,合并id ,姓名,性别,年龄,出生日期,身份证号|新主页id :原病人id:原主页id
        varData = Split(Split(str保留信息, "|")(0), ",")
        lng保留Id = varData(0)
        lng合并ID = varData(1)
        strName = varData(2)
        strSex = varData(3)
        strAge = varData(4)
        strBirthDate = varData(5)
        strCardID = varData(6)
        varData = Split(Split(str保留信息, "|")(1), ",")
        For i = 0 To UBound(varData)
            varData = Split(varData(i), ":")
            strJson = strJson & "," & "{""page_id_new"":" & varData(0) & ",""pati_id_befor"":" & varData(1) & ",""page_id_befor"":" & varData(2) & "}"
        Next
        If strJson <> "" Then
            strJson = Mid(strJson, 2)
            strJson = "{""input"":{""retain_id"":" & lng保留Id & ",""merge_id"":" & lng合并ID & ",""pati_name"":""" & strName & """,""pati_sex"":""" & strSex & """,""pati_age"":""" & strAge & """,""pati_borth_time"":""" & strBirthDate & """,""pati_identity"":""" & strCardID & """,""item_list"":[" & strJson & "]}}"
            If zlDrugSvr_Merage(strJson, strErrMsg, lngModule) = False Then
                 gcnOracle.RollbackTrans: blnTrans = False
                 If strErrMsg <> "" Then MsgBox "病人合并失败，原因如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                 Exit Function
            End If
             
            '暂时不分事务，原因是与周韬协商，这个只是更新信息，基本不出错,以后改为服务后，可能会造成不一致的情况
            If ZlStuffsvr_Merage(strJson, strErrMsg, lngModule) = False Then
                 gcnOracle.RollbackTrans: blnTrans = False
                 If strErrMsg <> "" Then MsgBox "病人合并失败，原因如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                 Exit Function
            End If
            If ZlPivassvr_merage(strJson, strErrMsg, lngModule) = False Then
                 gcnOracle.RollbackTrans: blnTrans = False
                 If strErrMsg <> "" Then MsgBox "病人合并失败，原因如下:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName
                 Exit Function
            End If
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    zlPati_Merage = True
    Exit Function
errHandle:
    If blnTrans Then gcnBJYB.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 

Public Function zlDrugSvr_Merage(strJsonIn As String, Optional ByRef strErrMsg_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:合并病人(药品)
    '入参:strJsonIn-Json串
    '出参:strErrMsg_Out-返回错误信息
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim i As Long, j As Long
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input           病人信息合并调用过程Zl_病人信息_Merge时先将要合并的病人id和主页id清单查询出来，入参的组织
    '       retain_id   N   1   保留id
    '       merge_id     N   1   合并id
    '       pati_name    C   1   姓名
    '       pati_sex C   1   性别
    '       pati_age     C   1   年龄
    '       pati_borth_time  C   1   出生日期
    '       pati_identity    C   1   身份证号
    '       item_list
    '          page_id_new   N   1   新主页id
    '          pati_id_befor     N   1   原病人id
    '          page_id_befor     N   1   原主页id
    '
    '    出参 json
    '    output
    '    code    N   1   应答吗：0-失败；1-成功
    '    message C   1   应答消息：失败时返回具体的错误信息
    If objServiceCall.CallService("zl_DrugSvr_Merage", strJsonIn, , "医保病人信息合并", lngModule, False) = False Then Exit Function
    If Val(Nvl(objServiceCall.GetJsonNodeValue("output.code"))) = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    zlDrugSvr_Merage = True
    Exit Function
errHandle:
    Err.Raise Err.Number, Err.Description
End Function
 
Public Function ZlStuffsvr_Merage(strJsonIn As String, Optional ByRef strErrMsg_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:合并病人(卫材)
    '入参:strJsonIn-Json串
    '出参:strErrMsg_Out-返回错误信息
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim i As Long, j As Long
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input           病人信息合并调用过程Zl_病人信息_Merge时先将要合并的病人id和主页id清单查询出来，入参的组织
    '       retain_id   N   1   保留id
    '       merge_id     N   1   合并id
    '       pati_name    C   1   姓名
    '       pati_sex C   1   性别
    '       pati_age     C   1   年龄
    '       pati_borth_time  C   1   出生日期
    '       pati_identity    C   1   身份证号
    '       item_list
    '          page_id_new   N   1   新主页id
    '          pati_id_befor     N   1   原病人id
    '          page_id_befor     N   1   原主页id
    '
    '    出参 json
    '    output
    '    code    N   1   应答吗：0-失败；1-成功
    '    message C   1   应答消息：失败时返回具体的错误信息
    If objServiceCall.CallService("Zl_Stuffsvr_Merage", strJsonIn, , "医保病人信息合并", lngModule, False) = False Then Exit Function
    If Val(Nvl(objServiceCall.GetJsonNodeValue("output.code"))) = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    ZlStuffsvr_Merage = True
    Exit Function
errHandle:
    Err.Raise Err.Number, Err.Description
End Function
 
 
Public Function ZlPivassvr_merage(strJsonIn As String, Optional ByRef strErrMsg_Out As String, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:合并病人(静配)
    '入参:strJsonIn-Json串
    '出参:strErrMsg_Out-返回错误信息
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-29 17:15:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim i As Long, j As Long
    Err = 0: On Error GoTo errHandle:
    If lngModule <> 0 Then mlngModule = lngModule
    If GetServiceCall(objServiceCall) = False Then Exit Function
    '    input           病人信息合并调用过程Zl_病人信息_Merge时先将要合并的病人id和主页id清单查询出来，入参的组织
    '       retain_id   N   1   保留id
    '       merge_id     N   1   合并id
    '       pati_name    C   1   姓名
    '       pati_sex C   1   性别
    '       pati_age     C   1   年龄
    '       pati_borth_time  C   1   出生日期
    '       pati_identity    C   1   身份证号
    '       item_list
    '          page_id_new   N   1   新主页id
    '          pati_id_befor     N   1   原病人id
    '          page_id_befor     N   1   原主页id
    '
    '    出参 json
    '    output
    '    code    N   1   应答吗：0-失败；1-成功
    '    message C   1   应答消息：失败时返回具体的错误信息
    If objServiceCall.CallService("Zl_Pivassvr_merage", strJsonIn, , "医保病人信息合并", lngModule, False) = False Then Exit Function
    If Val(Nvl(objServiceCall.GetJsonNodeValue("output.code"))) = 0 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        Exit Function
    End If
    ZlPivassvr_merage = True
    Exit Function
errHandle:
    Err.Raise Err.Number, Err.Description
End Function
  
  
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mobjServiceCall = Nothing
End Sub

