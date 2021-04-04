VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNurseIntegrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjXML As clsXML

Public Property Get RelatedUnitID() As String
    '整体护理的病区ID
    RelatedUnitID = gstrRelatedUnitID
End Property

Public Property Get RelatedUserID() As String
    '整体护理的人员ID
    RelatedUserID = gstrRelatedUserID
End Property

Public Property Get RelatedPatientID() As String
    '整体护理的病人ID
    RelatedPatientID = gstrRelatedPatientID
End Property

Public Function zlInitCommon(ByVal cnOracle As ADODB.Connection, Optional ByVal strDBUser As String, Optional ByVal blnAlone As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及数据库连接
    '入参:l
    '     cnOracle-数据库连接对象
    '     strDBUser-当前数据库登录用户名
    '     blnAlone=是否独立程序调用,独立程序调用将调用InitCommon方法.
    '返回:初始化成功,返回true,否则返回False
    '编制:刘鹏飞
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    gblnAlone = blnAlone
    Set gcnOracle = cnOracle: gstrDBUser = strDBUser
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrIntergrateIP = gobjDatabase.GetPara("整体护理IP地址", 100, 1265)
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IPAdreesCheck(ByVal strIP As String, strErrMsg As String) As Boolean
'功能：供参数设置调用，检查IP地址是否正确
    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    
    If strIP = "" Then Exit Function
    strUrl = "http://" & strIP & "/services/sys/login/OAuth" '固定地址校验
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg, , strIP) = False Then
        Exit Function
    End If
    IPAdreesCheck = True
End Function

Public Function IsUseIntergrate(strErrMsg As String, intMode As Integer) As Boolean
'功能：检查设置地址是否正确，用于判定是否启用整理护理
'intMode:0-IP地址为空;1-登录服务器错误;2-服务器登录成功，IE版本过低
'  使用整体移动护理要求IE版本最低为IE11，低于该版本将不能使用
    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    Dim strIEVersion As String, blnIEOk As Boolean
    Dim strExeName As String, strPathName As String
    
    If gstrIntergrateIP = "" Then intMode = 0: Exit Function
    strUrl = "http://" & gstrIntergrateIP & "/services/sys/login/OAuth" '固定地址校验
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg) = False Then
        intMode = 1
        Exit Function
    End If
    blnIEOk = False
    On Error GoTo ErrHand
    strIEVersion = getvalue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer", "svcVersion", REG_SZ)
    If strIEVersion = "" Then
        strIEVersion = getvalue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer", "Version", REG_SZ)
    End If
    If InStr(1, strIEVersion, ".") > 0 Then
        If Val(Mid(strIEVersion, 1, InStr(1, strIEVersion, ".") - 1)) >= 11 Then
            blnIEOk = True
        End If
    End If
    If blnIEOk = False Then
        intMode = 2
        strErrMsg = "您电脑的IE版本过低无法使用整体护理功能，整体护理要求使用的IE版本不能低于IE11，请联系管理员升级您本机的IE版本！"
        Exit Function
    End If
    
    '获取当前程序调用的进程名
    If gobjComlib.OS.IsDesinMode = False Then '非IDE模式获取执行程序exe(调试环境始终是vb6.exe),以及设计webBrowser控件默认使用IE11浏览器(只有在运行环境下该设置才有用)
        Call GetProcessName(GetCurrentProcessId, strExeName, strPathName)
        gstrExeName = strExeName
        If SetWBIEVerSion(gstrExeName, strErrMsg) = False Then
            intMode = 2
            Exit Function
        End If
    End If
    
    IsUseIntergrate = True
    Exit Function
ErrHand:
    intMode = 2
    strErrMsg = "检查IE版本读取注册表失败，请检查以下目录是否存在【HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer】"
End Function

Public Function UserLogin(strErrMsg As String, Optional blnNologn As Boolean = False) As Boolean
'功能：整体护理登录，调用整体护理数据接口前调用（注意：services/pub这种地址都不需要登录）
'URL：192.168.4.61/services/sys/login/OAuth
'   参数: {
'           "Params":{
'                "UserID":"320" // zlhis用户ID
'           }
'           }

'blnNologn：用于判断是否需要再次登录

    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    Dim strRetrun As String
    
    If blnNologn = True And IntergrateUserInfo.Cookie <> "" Then
        UserLogin = True
        Exit Function
    End If
    
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    strUrl = "http://" & gstrIntergrateIP & "/services/sys/login/OAuth" '固定地址校验
    strParams = "{""Params"":{""UserID"":""" & UserInfo.id & """,""UserName"":""" & UserInfo.用户名 & """}}"
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg) = False Then
        Call WriteBusinessLOG("用户登录", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("用户登录", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
            '从消息头中获取登录人员的Cookie
            IntergrateUserInfo.Cookie = oXmlHttp.getResponseHeader("Set-Cookie")
            IntergrateUserInfo.id = "" & .Eval("json.Output.Data.ID")
            IntergrateUserInfo.UserName = "" & .Eval("json.Output.Data.UserName")
            IntergrateUserInfo.Name = "" & .Eval("json.Output.Data.Name")
            IntergrateUserInfo.Sex = "" & .Eval("json.Output.Data.Sex")
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With
    
    UserLogin = True
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------------
'病区、病人业务接口
'------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetRelatedIDToGUID(ByVal strUnitID As String, strErrMsg As String, Optional ByVal strPatiID As String = "") As Boolean
'功能：根据zlhis的病区ID、用户ID和病人住院唯一表示, 获取整体护理的病区ID、用户ID和病人ID
'strPatiID：ZLHIS中的病人标识(病人ID|主页ID|婴儿序号)
'调用时机：病区切换时调用或病人切换时调用
'URL：192.168.4.61/services/pub/OutMethod/RelatedIDToGUID
'参数: {
'          "Params":{
'            "RelatedUserID":"320", // zlhis用户ID
'            "RelatedUnitID":"320",  // zlhis病区ID
'            "RelatedPatientID":"54486", //zlhis病人ID
'            "HomePageID":"1",
'            "Baby":"0"
'          }
'        }
'返回：{"Flag":1,"Msg":"","Output":{"Data":{"UnitID":null,"UserID":"d1a272dd-4282-4464-a8ba-fce13e5ab9ea"}}}
    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    Dim strRetrun As String
    Dim arrParent() As String
    Dim lngRelatedPatientID As Long, lngHomePageID As Long, lngBaby As Long, blnReadParent As Boolean
    
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    strUrl = "http://" & gstrIntergrateIP & "/services/pub/OutMethod/RelatedIDToGUID"
    If InStr(1, strPatiID, "|") = 0 Then
        strParams = "{""Params"":{""RelatedUserID"":""" & UserInfo.id & """,""RelatedUnitID"":""" & strUnitID & """}}"
    Else
        blnReadParent = True
        arrParent = Split(strPatiID, "|")
        lngRelatedPatientID = arrParent(0)
        lngHomePageID = arrParent(1)
        If UBound(arrParent) > 1 Then lngBaby = arrParent(2)
        strParams = "{""Params"":{""RelatedUserID"":""" & UserInfo.id & """,""RelatedUnitID"":""" & strUnitID & """,""RelatedPatientID"":""" & lngRelatedPatientID & """,""HomePageID"":""" & lngHomePageID & """,""Baby"":""" & lngBaby & """}}"
    End If
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg) = False Then
        Call WriteBusinessLOG("获取病区和用户ID", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("获取病区和用户ID", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
            gstrRelatedUnitID = "" & .Eval("json.Output.Data.UnitID")
            gstrRelatedUserID = "" & .Eval("json.Output.Data.UserID")
            If blnReadParent = True Then
                gstrRelatedPatientID = "" & .Eval("json.Output.Data.PatientID")
            End If
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With
    GetRelatedIDToGUID = True
End Function

Public Function GetLesionMethod(strTabs As String, strErrMsg As String) As Boolean
'功能：获取整体护理中的病区业务标签
'URL：192.168.4.61/services/pub/OutMethod/LesionMethod
'参数 {
'  "Params":{
'    "RelateUserID":"320"
'  }
'}
'返回：{"Flag":1,"Msg":"","Output":{"Data":[{"Name":"新版交班报告","Url":"/shiftReport","Params":"1"},{"Name":"输液状态","Url":"/infuState","Params":"2"},{"Name":"护理巡视","Url":"/nurseTour","Params":null},{"Name":"配液情况","Url":"/liquid","Params":"1"},{"Name":"交班报告","Url":"/reportHand","Params":"2"},{"Name":"输液巡视","Url":"/infuTour","Params":"1"}]}}
'说明：URL参数（URL拼接的方式）：
'LessionID 整体护理 病区ID (通过RelatedIDToGUID接口获取)---GetRelatedIDToGUID
'UserID 整体护理 用户ID (通过RelatedIDToGUID接口获取)---GetRelatedIDToGUID
'From his(固定参数不变, 代表从zlhis打开页面)
'Params 通过接口获取病区功能列表时返回的Params字段(需要EncodeUrl, vb6的实现网上应该能搜得到)

    Dim oXmlHttp As XMLHTTP
    Dim strName As String, strUrl As String, strParams As String
    Dim strRetrun As String
    Dim i As Long, lngCount As Long
    
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    
    strUrl = "http://" & gstrIntergrateIP & "/services/pub/OutMethod/LesionMethod"
    strParams = "{""Params"":{""RelateUserID"":""" & UserInfo.id & """}}"
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg) = False Then
        Call WriteBusinessLOG("获取病区业务标签", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("获取病区业务标签", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
            '获取Data数组的长度
            lngCount = .Eval("json.Output.Data").length
            strTabs = ""
            For i = 0 To lngCount - 1
                strName = "" & .Eval("json.Output.Data[" & i & "].Name")
                strUrl = "http://" & gstrIntergrateIP & .Eval("json.Output.Data[" & i & "].Url")
                strParams = encodeURI("" & .Eval("json.Output.Data[" & i & "].Params"))
                strTabs = strTabs & "<Item><Name>" & strName & "</Name>" & "<Url LessionID=""[LPF_UnitID]"" UserID=""[LPF_UserID]""" & _
                        " From=""his"" Params=""" & strParams & """>" & strUrl & "</Url></Item>"
            Next
            If lngCount > 0 Then strTabs = "<Tab>" & strTabs & "</Tab>"
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With
    GetLesionMethod = True
End Function

Public Function GetPatientMethod(strTabs As String, strErrMsg As String) As Boolean
'功能：获取整体护理中的病人业务标签
'URL：192.168.4.61/services/pub/OutMethod/PatientMethod
'参数 {
'  "Params":{
'    "RelateUserID":"320"
'  }
'}
'返回：{"Flag":1,"Msg":"","Output":{"Data":[{"Name":"护理计划","Url":"/nursePlan","Params":null},{"Name":"护理评估","Url":"/nurseAssess","Params":null},{"Name":"护理评分","Url":"/nurseScore","Params":null},{"Name":"健康宣教","Url":"/healthEdu","Params":null}]}}

'说明：URL参数（URL拼接的方式）：
'LessionID 整体护理 病区ID (通过RelatedIDToGUID接口获取)
'UserID; 整体护理; 用户ID(通过RelatedIDToGUID接口获取)
'PatientID; 整体护理; 病人ID
'From; his(固定参数不变, 代表从zlhis打开页面)
'Params; 通过接口获取病区功能列表时返回的Params字段(需要EncodeUrl, vb6的实现网上应该能搜得到)

    Dim oXmlHttp As XMLHTTP
    Dim strName As String, strUrl As String, strParams As String
    Dim strRetrun As String
    Dim i As Long, lngCount As Long
    
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    
    strUrl = "http://" & gstrIntergrateIP & "/services/pub/OutMethod/PatientMethod"
    strParams = "{""Params"":{""RelateUserID"":""" & UserInfo.id & """}}"
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg) = False Then
        Call WriteBusinessLOG("获取病人业务标签", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("获取病人业务标签", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
            '获取Data数组的长度
            lngCount = .Eval("json.Output.Data").length
            strTabs = ""
            For i = 0 To lngCount - 1
                strName = "" & .Eval("json.Output.Data[" & i & "].Name")
                strUrl = "http://" & gstrIntergrateIP & .Eval("json.Output.Data[" & i & "].Url")
                strParams = encodeURI("" & .Eval("json.Output.Data[" & i & "].Params"))
                strTabs = strTabs & "<Item><Name>" & strName & "</Name>" & "<Url LessionID=""[LPF_UnitID]"" UserID=""[LPF_UserID]""" & _
                    " PatientID=""[LPF_PatientID]"" From=""his"" Params=""" & strParams & """>" & strUrl & "</Url></Item>"
            Next
            If lngCount > 0 Then strTabs = "<Tab>" & strTabs & "</Tab>"
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With
    GetPatientMethod = True
End Function

Public Function GetForm(ByVal strName As String, ByVal strUrl As String) As Object
'功能：返回Tab签窗体对象。工作站窗体加载时调用
    Dim objFrom As Object
    
    Set objFrom = New frmMethod
    objFrom.Caption = strName
    objFrom.Tag = strUrl
    Set GetForm = objFrom
End Function

Public Function GetDocForm() As Object
'功能：返回Tab签窗体对象。工作站窗体加载时调用
    Dim objFrom As Object
    
    Set objFrom = New frmDocRead
    objFrom.Caption = "护理评估信息"
    objFrom.Tag = "http://192.168.4.61/ascore?PatientID='12345'&HomePageID=1"
    Set GetDocForm = objFrom
End Function

Public Function RefreshLesionMethod(ByVal objFrom As Object, ByVal strUrl As String, Optional ByVal strUnitID As String, Optional ByVal strUserID As String) As Boolean
'功能：整体护理病区业务功能刷新
'说明：将[LPF_UnitID]替换成整体护理中的病区ID，[LPF_UserID] 替换成整体护理中的人员ID
    Dim strParam As String
    If strUnitID = "" Then strUnitID = gstrRelatedUnitID
    If strUserID = "" Then strUserID = gstrRelatedUserID
    strUrl = Replace(Replace(strUrl, "[LPF_UnitID]", strUnitID), "[LPF_UserID]", strUserID)
    strParam = Split(strUrl, "?")(1)
    strUrl = Split(strUrl, "?")(0)
    Call WriteBusinessLOG(objFrom.Caption, strUrl, strParam)
    RefreshLesionMethod = objFrom.zlRefresh(strUrl, strParam)
End Function

Public Function RefreshPatientMethod(ByVal objFrom As Object, ByVal strUrl As String, ByVal strPatientId As String, _
        Optional ByVal strUnitID As String, Optional ByVal strUserID As String) As Boolean
'功能：整理护理病人业务功能刷新
'说明：替换要素包含：[LPF_UnitID]--整体护理病区ID；[LPF_UserID] --整理护理人员ID；[LPF_PatientID]--整体护理病人ID
    Dim strParam As String
    If strUnitID = "" Then strUnitID = gstrRelatedUnitID
    If strUserID = "" Then strUserID = gstrRelatedUserID
    strUrl = Replace(Replace(strUrl, "[LPF_UnitID]", strUnitID), "[LPF_UserID]", strUserID) '--病区和人员ID替换
    strUrl = Replace(strUrl, "[LPF_PatientID]", strPatientId) '--病人ID替换
    strParam = Split(strUrl, "?")(1)
    strUrl = Split(strUrl, "?")(0)
    Call WriteBusinessLOG(objFrom.Caption, strUrl, strParam)
    RefreshPatientMethod = objFrom.zlRefresh(strUrl, strParam)
End Function

Public Function ExistsCustomer(ByVal lngPaitID As Long, ByVal lngPageID As Long, strErrMsg As String) As Boolean
'功能：检查his中的某个病人是否存在，不存在移动则会自动同步
    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    Dim strRetrun As String
    
    If UserLogin(strErrMsg, True) = False Then Exit Function
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    strUrl = "http://" & gstrIntergrateIP & "/services/pub/customer/ExistsCustomer"
    strParams = "{""Params"":{""RelatID"":""" & lngPaitID & """,""InTimes"":""" & lngPageID & """}}"
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg, True) = False Then
        Call WriteBusinessLOG("病人存在检查", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("病人存在检查", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
            
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With
    
    ExistsCustomer = True
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'插件接口
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetPlugin(Optional ByVal strName As String) As Object
'功能：整体护理扩展面板
    Dim objFrom As Object
    Dim strUrl As String
    
    strUrl = "http://" & gstrIntergrateIP & "/plugin?LessionID=[LPF_UnitID]&UserID=[LPF_UserID]&From=his"
    Set objFrom = New frmMethod
    objFrom.Caption = strName
    objFrom.Tag = strUrl
    Set GetPlugin = objFrom
End Function

Public Function RefreshPlugin(ByVal objFrom As Object, ByVal strUrl As String, Optional ByVal strUnitID As String, Optional ByVal strUserID As String) As Boolean
'功能：扩展面板刷新
'说明：将[LPF_UnitID]替换成整体护理中的病区ID，[LPF_UserID] 替换成整体护理中的人员ID
    Dim strParam As String
    If strUnitID = "" Then strUnitID = gstrRelatedUnitID
    If strUserID = "" Then strUserID = gstrRelatedUserID
    strUrl = Replace(Replace(strUrl, "[LPF_UnitID]", strUnitID), "[LPF_UserID]", strUserID)
    strParam = Split(strUrl, "?")(1)
    strUrl = Split(strUrl, "?")(0)
    Call WriteBusinessLOG(objFrom.Caption, strUrl, strParam)
    RefreshPlugin = objFrom.zlRefresh(strUrl, strParam)
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------------
'具体业务功能接口
'------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetGroupsList(strList As String, strErrMsg As String) As Boolean
'功能：获取护理分组信息
'URL：http://192.168.4.61/Services/pub/UnitHomePage/GroupsList
'   参数：{"Params":{}}
'   返回：
'{"Flag":1,"Msg":"","Output":{"Data":[{"ID":"72ffdb68-64f4-4be5-8a30-515c70dfc574","CodeID":"dcfad45f-8b39-49cc-a9cd-cb082d132b8c","Name":"护理1组","Level":null,"IsFinal":0,"SeqNo":1,"ModifyUser":null,"ModifyTime":null,"IsInvalid":0,"InvalidTime":null,"InputCode":"HL1Z","Remarks":null,"CommonCodeValue":"1"},
'{"ID":"8ea12c48-22ca-487a-9606-c4dfba07e890","CodeID":"dcfad45f-8b39-49cc-a9cd-cb082d132b8c","Name":"护理2组","Level":null,"IsFinal":0,"SeqNo":2,"ModifyUser":null,"ModifyTime":null,"IsInvalid":0,"InvalidTime":null,"InputCode":"HL2Z","Remarks":null,"CommonCodeValue":"2"},
'{"ID":"e1e6d337-91a4-4e57-9de2-6d7fc810f546","CodeID":"dcfad45f-8b39-49cc-a9cd-cb082d132b8c","Name":"护理3组","Level":null,"IsFinal":0,"SeqNo":3,"ModifyUser":null,"ModifyTime":null,"IsInvalid":0,"InvalidTime":null,"InputCode":"HL3Z","Remarks":null,"CommonCodeValue":"3"},
'{"ID":"da0f0020-ac0c-4500-b30e-2e9b48d5d4b7","CodeID":"dcfad45f-8b39-49cc-a9cd-cb082d132b8c","Name":"护理4组","Level":null,"IsFinal":0,"SeqNo":4,"ModifyUser":null,"ModifyTie":null,"IsInvalid":0,"InvalidTime":null,"InputCode":"HL4Z","Remarks":null,"CommonCodeValue":"4"},
'{"ID":"9da6a188-6ffe-49c0-a84f-70114aa8847d","CodeID":"dcfad45f-8b39-49cc-a9cd-cb082d132b8c","Name":"护理5组","Level":null,"IsFinal":0,"SeqNo":5,"ModifyUser":null,"ModifyTime":null,"IsInvalid":0,"InvalidTime":null,"InputCode":"HL5Z","Remarks":null,"CommonCodeValue":"5"},
'{"ID":"8717e580-97e8-4058-bd40-d0c632b8b8a7","CodeID":"dcfad45f-8b39-49cc-a9cd-cb082d132b8c","Name":"护理6组","Level":null,"IsFinal":0,"SeqNo":6,"ModifyUser":null,"ModifyTime":null,"IsInvalid":0,"InvalidTime":null,"InputCode":"HL6Z","Remarks":null,"CommonCodeValue":"6"},
'{"ID":"b15e5397-0221-45d4-9f88-e31de0ad6edf","CodeID":"dcfad45f-8b39-49cc-a9cd-cb082d132b8c","Name":"护理7组","Level":null,"IsFinal":0,"SeqNo":7,"ModifyUser":null,"ModifyTime":null,"IsInvalid":0,"InvalidTime":null,"InputCode":"HL7Z","Remarks":null,"CommonCodeValue":"7"},
'{"ID":"08569cef-e482-4570-ab91-3e8089b90dad","CodeID":"dcfad45f-8b39-49cc-a9cd-cb082d132b8c","Name":"护理8组","Level":null,"IsFinal":0,"SeqNo":8,"ModifyUser":null,"ModifyTime":null,"IsInvalid":0,"InvalidTime":null,"InputCode":"HL8Z","Remarks":null,"CommonCodeValue":"8"}]}}

'接口返回：TRUE调用成功(strList护理分组信息)；FALSE接口调用失败（strErrMsg错误信息）
    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    Dim strRetrun As String
    Dim i As Long, lngCount As Long, strName As String, strID As String
    
    If UserLogin(strErrMsg, True) = False Then Exit Function
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    strUrl = "http://" & gstrIntergrateIP & "/Services/pub/UnitHomePage/GroupsList" '
    strParams = "{""Params"":{}}"
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg, True) = False Then
        Call WriteBusinessLOG("获取护理分组信息", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("获取护理分组信息", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
            lngCount = .Eval("json.Output.Data").length
            strList = ""
            If lngCount <= 0 Then GetGroupsList = True: Exit Function
            mobjXML.ClearXmlText
            Call mobjXML.AppendNode("List")
            For i = 0 To lngCount - 1
                strName = "" & .Eval("json.Output.Data[" & i & "].Name")
                strID = "" & .Eval("json.Output.Data[" & i & "].ID")
                Call mobjXML.AppendNode("Item")
                Call mobjXML.AppendData("ID", strID)
                Call mobjXML.AppendData("Name", strName)
                Call mobjXML.AppendNode("Item", True)
            Next
            Call mobjXML.AppendNode("List", True)
            strList = mobjXML.XmlText
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With
    
    GetGroupsList = True
End Function

Public Function AddorUpdateGroups(ByVal strGroupID As String, ByVal strBedNumber As String, strGroupNumber As String, strErrMsg As String, Optional ByVal strUnitID As String) As Boolean
'功能：设置病区床位所属的护理分组
'URL：http://192.168.4.61/Services/pub/UnitHomePage/AddorUpdateGroups             改接口为整体护理过滤病人综合处理，此处HIS调用只管护理分组
'   参数：{"Params":{
'"LesionID":"01b7e14e-ba9f-43d6-879b-e1b3972d2831",  --整体护理病区ID（GetRelatedIDToGUID函数获取）
'"GroupID":"1395a2fb-225a-4bc3-9ced-97ad60600320", --病人列表中的 GroupID  （PatientList函数获取）
'"GroupNumber":"8ea12c48-22ca-487a-9606-c4dfba07e890" --护理分组ID（GroupsList函数获取）
'"BedNumber:13"  --新增是传入床号(strGroupID:为空的情况)
'}}
'   返回：{"Flag":1,"Msg":"1","Output":""}

'接口返回：TRUE设置成功；FALSE设置失败（strErrMsg错误信息）

    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    Dim strRetrun As String
    
    If strUnitID = "" Then strUnitID = gstrRelatedUnitID
    If UserLogin(strErrMsg, True) = False Then Exit Function
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    strUrl = "http://" & gstrIntergrateIP & "/Services/pub/UnitHomePage/AddorUpdateGroups" '
    If strGroupID = "" Then
        strParams = "{""Params"":{""LesionID"":""" & strUnitID & """,""BedNumber"":""" & strBedNumber & """,""GroupNumber"":""" & strGroupNumber & """}}"
    Else
        strParams = "{""Params"":{""LesionID"":""" & strUnitID & """,""GroupID"":""" & strGroupID & """,""GroupNumber"":""" & strGroupNumber & """}}"
    End If
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg, True) = False Then
        Call WriteBusinessLOG("护理分组设置", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("护理分组设置", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With
    
    AddorUpdateGroups = True
End Function

Public Function GetPatientList(strPatientList As String, strErrMsg As String, Optional ByVal strGroupNumber As String, Optional ByVal strUnitID As String) As Boolean
'功能：护理组条件筛选列表获取
'URL：http://192.168.4.61/Services/pub/UnitHomePage/PatientList
'   参数：{"Params":{
'"LesionID":"01b7e14e-ba9f-43d6-879b-e1b3972d2831",  --整体护理病区ID
'"Type":5, 5/在院病人  3/出院病人
'"StartTime":"2017/03/14",  --传入空
'"EndTime":"2017/03/21",  --传入空
'"IsRefresh":0, --传入0
'"InputCode":"", --病人查找 传入空
'"GroupNumber":"8ea12c48-22ca-487a-9606-c4dfba07e890",  --护理分组ID(查看未分组传入0即可)
'"LstLevel":""  --护理等级 传入空
'},
'"Page":{"PageNum":1,"PageSize":9999}  -默认传入1 9999
'}
'   返回：
'{"Flag":1,"Msg":"","Output":{"Data":{"PatientList":[{"ID":"7ad59f55-004c-423b-a6fd-60d826c1e4bc","RelatID":"52775","Name":"张萌","AdmissionNumber":"201600014","Baby":0,"HomepageID":"1","Age":"23岁","Sex":"男","Nurse":"陈丹","NursingLevel":"特级护理","IsCritical":0,"GroupID":"1b9c8cba-01b5-4333-810d-182752c384a2","GroupNumber":"72ffdb68-64f4-4be5-8a30-515c70dfc574","GroupValue":1,
'"ExtraInfo":"病况_一般,出生日期_1993-05-19,当前病区_外一病区,当前科室_外一科,婚姻状况_未婚,籍贯_,简要病史_,联系人电话_,联系人关系_,联系人姓名_,民族_汉族,入科时间_2016-04-01,入院病区_外一病区,入院方式_门诊,入院科室_外一科,入院原因_治疗,身份_其他,身份证号_500382199305193796,手术部位_,手术名称_,西医入院诊断_(031014)急性阑尾炎,学历_其他,医保号_,医保名称_,职业_工人,中医入院诊断_,主治医生_陈涛",
'"InDay":"2016-04-01 12:00:24","OutDay":null,"EndDay":"2017-04-07 14:10:09","BZ":0,"BedNumber":"         1","ArrayExtraInfo":[{"病况":"一般"},{"出生日期":"1993-05-19"},{"当前病区":"外一病区"},{"当前科室":"外一科"},{"婚姻状况":"未婚"},{"籍贯":""},{"简要病史":""},{"联系人电话":""},{"联系人关系":""},{"联系人姓名":""},{"民族":"汉族"},{"入科时间":"2016-04-01"},{"入院病区":"外一病区"},{"入院方式":"门诊"},
'{"入院科室":"外一科"},{"入院原因":"治疗"},{"身份":"其他"},{"身份证号":"500382199305193796"},{"手术部位":""},{"手术名称":""},{"西医入院诊断":"(031014)急性阑尾炎"},{"学历":"其他"},{"医保号":""},{"医保名称":""},{"职业":"工人"},{"中医入院诊断":""},{"主治医生":"陈涛"}],"NursingLevelCh":"特","Color":"#c57af5"},
'{"ID":"d17231ed-2da9-49ae-901e-c7e069572223","RelatID":"52955","Name":"医嘱测试","AdmissionNumber":"201600042","Baby":0,"HomepageID":"1","Age":"30岁","Sex":"男","Nurse":"陈丹","NursingLevel":"Ⅱ级护理","IsCritical":0,"GroupID":"8da1de46-e4a2-4fea-898f-90eedc215935","GroupNumber":"72ffdb68-64f4-4be5-8a30-515c70dfc574","GroupValue":1,
'"ExtraInfo":"病况_一般,出生日期_1986-01-01,当前病区_外一病区,当前科室_外一科,婚姻状况_,籍贯_,简要病史_,联系人电话_,联系人关系_,联系人姓名_,民族_汉族,入科时间_2016-06-23,入院病区_外一病区,入院方式_门诊,入院科室_外一科,入院原因_治疗,身份_其他,身份证号_,手术部位_,手术名称_,西医入院诊断_(022001)急性上呼吸道感染,学历_其他,医保号_,医保名称_,职业_工人,中医入院诊断_,主治医生_陈涛","InDay":"2016-06-23 10:41:23","OutDay":null,"EndDay":"2017-04-07 14:10:09","BZ":0,"BedNumber":"        41",
'"ArrayExtraInfo":[{"病况":"一般"},{"出生日期":"1986-01-01"},{"当前病区":"外一病区"},{"当前科室":"外一科"},{"婚姻状况":""},{"籍贯":""},{"简要病史":""},{"联系人电话":""},{"联系人关系":""},{"联系人姓名":""},{"民族":"汉族"},
'{"入科时间":"2016-06-23"},{"入院病区":"外一病区"},{"入院方式":"门诊"},{"入院科室":"外一科"},{"入院原因":"治疗"},{"身份":"其他"},{"身份证号":""},{"手术部位":""},{"手术名称":""},{"西医入院诊断":"(022001)急性上呼吸道感染"},{"学历":"其他"},{"医保号":""},{"医保名称":""},{"职业":"工人"},{"中医入院诊断":""},{"主治医生":"陈涛"}],"NursingLevelCh":"II","Color":"#5c61ea"},
'{"ID":"357526be-0786-475b-b8c5-2343e386296e","RelatID":"54044","Name":"彭云","AdmissionNumber":"2009003663","Baby":0,"HomepageID":"1","Age":"16岁","Sex":"男","Nurse":null,"NursingLevel":"Ⅱ级护理","IsCritical":0,"GroupID":"470cf438-dc0a-4d72-8459-6648f0de0dff","GroupNumber":"72ffdb68-64f4-4be5-8a30-515c70dfc574","GroupValue":1,
'"ExtraInfo":"病况_一般,出生日期_2000-01-01,当前病区_外一病区,当前科室_外一科,婚姻状况_,籍贯_,简要病史_,联系人电话_,联系人关系_,联系人姓名_,民族_汉族,入科时间_,入院病区_外一病区,入院方式_门诊,入院科室_外一科,入院原因_治疗,身份_其他,身份证号_897867542334567,手术部位_,手术名称_,西医入院诊断_,学历_其他,医保号_,医保名称_,职业_工人,中医入院诊断_,主治医生_","InDay":"2016-12-13 15:44:11","OutDay":null,"EndDay":"2017-04-07 14:10:09","BZ":0,"BedNumber":"       ▲51",
'"ArrayExtraInfo":[{"病况":"一般"},{"出生日期":"2000-01-01"},{"当前病区":"外一病区"},{"当前科室":"外一科"},{"婚姻状况":""},{"籍贯":""},{"简要病史":""},{"联系人电话":""},{"联系人关系":""},{"联系人姓名":""},{"民族":"汉族"},{"入科时间":""},{"入院病区":"外一病区"},{"入院方式":"门诊"},{"入院科室":"外一科"},{"入院原因":"治疗"},{"身份":"其他"},{"身份证号":"897867542334567"},{"手术部位":""},{"手术名称":""},{"西医入院诊断":"",{"学历":"其他"},{"医保号":""},{"医保名称":""},{"职业":"工人"},{"中医入院诊断":""},{"主治医生":""}],"NursingLevelCh":"II","Color":"#5c61ea"},
'{"ID":"ff582f29-c155-4777-88da-38de4e1b7a6e","RelatID":"54084","Name":"郭锐013","AdmissionNumber":"2009003686","Baby":0,"HomepageID":"1","Age":"26岁","Sex":"男","Nurse":null,"NursingLevel":"特级护理","IsCritical":0,"GroupID":"06597f45-01c1-4ead-b353-a2744de55410","GroupNumber":"72ffdb68-64f4-4be5-8a30-515c70dfc574","GroupValue":1,"ExtraInfo":"病况_一般,出生日期_1990-01-01,当前病区_外一病区,当前科室_外一科,婚姻状况_,籍贯_,简要病史_,联系人电话_,联系人关系_,联系人姓名_,民族_汉族,入科时间_,入院病区_外一病区,入院方式_门诊,
'入院科室_外一科,入院原因_治疗,身份_其他,身份证号_8974931273921,手术部位_,手术名称_,西医入院诊断_,学历_其他,医保号_,医保名称_,职业_工人,中医入院诊断_,主治医生_","InDay":"2016-12-16 17:39:25","OutDay":null,"EndDay":"2017-04-07 14:10:09","BZ":0,"BedNumber":"       ▲72","ArrayExtraInfo":[{"病况":"一般"},{"出生日期":"1990-01-01"},{"当前病区":"外一病区"},{"当前科室":"外一科"},{"婚姻状况":""},{"籍贯":""},{"简要病史":""},{"联系人电话":""},{"联系人关系":""},{"联系人姓名":""},{"民族":"汉族"},{"入科时间":""},{"入院病区":"外一病区"},{"入院方式":"门诊"},{"入院科室":"外一科"},{"入院原因":"治疗"},{"身份":"其他"},{"身份证号":"8974931273921"},{"手术部位":""},{"手术名称":""},{"西医入院诊断":""},{"学历":"其他"},{"医保号":""},{"医保名称":""},{"职业":"工人"},{"中医入院诊断":""},{"主治医生":""}],"NursingLevelCh":"特","Color":"#c57af5"}],
'"GroupList":[{"GroupName":"全部病人","GroupCount":133,"GroupNumber":"AllPatients","SortNo":0},{"GroupNumber":"72ffdb68-64f4-4be5-8a30-515c70dfc574","SortNo":1,"GroupName":"护理1组","GroupCount":4},{"GroupNumber":"8ea12c48-22ca-487a-9606-c4dfba07e890","SortNo":2,"GroupName":"护理2组","GroupCount":1},{"GroupNumber":"0","SortNo":9999999999,"GroupName":"未分组","GroupCount":128}],"FocusList":[],"ColorList":{"Ⅰ级护理":"#ec5d64","Ⅱ级护理":"#5c61ea","Ⅲ级护理":"#58ba47","特级护理":"#c57af5"}}}}

'接口返回：TRUE调用成功(strPatientList病人信息)；FALSE接口调用失败（strErrMsg错误信息）
    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    Dim strRetrun As String
    Dim i As Long, lngCount As Long
    Dim lngPatiID As String, lngPageID As String, strGroupID As String, lngBaby As Long, strID As String
    Dim strName As String, strAge As String, strSex As String, strPageNo As String, strNursingLevel As String, strBedNumber As String
    Dim intIsBlock As Integer, intIsHighRisk As Integer, intIsHot As Integer
    
    If strUnitID = "" Then strUnitID = gstrRelatedUnitID
    If UserLogin(strErrMsg, True) = False Then Exit Function
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    strUrl = "http://" & gstrIntergrateIP & "/Services/pub/UnitHomePage/PatientList"
    strParams = "{""Params"":{""LesionID"":""" & strUnitID & """,""Type"":""" & 5 & """,""StartTime"":""" & "" & """," & _
            """EndTime"":""" & "" & """,""IsRefresh"":""" & 0 & """,""InputCode"":""" & "" & IIf(strGroupNumber = "", "", """,""GroupNumber"":""" & strGroupNumber) & """," & _
            """LstLevel"":""" & "" & """},""Page"":{""PageNum"":""" & 1 & """,""PageSize"":""" & 999 & """}}"
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg, True) = False Then
        Call WriteBusinessLOG("病人筛选", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("病人筛选", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
            lngCount = .Eval("json.Output.Data.PatientList").length
            strPatientList = ""
            If lngCount <= 0 Then GetPatientList = True: Exit Function
            mobjXML.ClearXmlText
            Call mobjXML.AppendNode("List")
            
            For i = 0 To lngCount - 1
                strID = "" & .Eval("json.Output.Data.PatientList[" & i & "].ID")  '整体护理的病人ID
                lngPatiID = Val("" & .Eval("json.Output.Data.PatientList[" & i & "].RelatID")) 'HIS病人ID
                lngPageID = Val("" & .Eval("json.Output.Data.PatientList[" & i & "].HomepageID"))
                lngBaby = Val("" & .Eval("json.Output.Data.PatientList[" & i & "].Baby"))
                strGroupID = "" & .Eval("json.Output.Data.PatientList[" & i & "].GroupID")
                strGroupNumber = "" & .Eval("json.Output.Data.PatientList[" & i & "].GroupNumber")
                strName = "" & .Eval("json.Output.Data.PatientList[" & i & "].Name")
                strPageNo = "" & .Eval("json.Output.Data.PatientList[" & i & "].AdmissionNumber")
                strAge = "" & .Eval("json.Output.Data.PatientList[" & i & "].Age")
                strSex = "" & .Eval("json.Output.Data.PatientList[" & i & "].Sex")
                strNursingLevel = "" & .Eval("json.Output.Data.PatientList[" & i & "].NursingLevel")
                strBedNumber = "" & .Eval("json.Output.Data.PatientList[" & i & "].BedNumber")
                
                intIsBlock = Val("" & .Eval("json.Output.Data.PatientList[" & i & "].IsBlock"))
                intIsHighRisk = Val("" & .Eval("json.Output.Data.PatientList[" & i & "].IsHighRisk"))
                intIsHot = Val("" & .Eval("json.Output.Data.PatientList[" & i & "].IsHot"))
                
                Call mobjXML.AppendNode("Item")
                Call mobjXML.AppendData("ID", strID)
                Call mobjXML.AppendData("Name", strName)
                Call mobjXML.AppendData("Age", strAge)
                Call mobjXML.AppendData("Sex", strSex)
                Call mobjXML.AppendData("PageNo", strPageNo)
                Call mobjXML.AppendData("PatiID", lngPatiID)
                Call mobjXML.AppendData("PageID", lngPageID)
                Call mobjXML.AppendData("Baby", lngBaby)
                Call mobjXML.AppendData("GroupID", strGroupID)
                Call mobjXML.AppendData("GroupNumber", strGroupNumber)
                Call mobjXML.AppendData("NursingLevel", strNursingLevel)
                Call mobjXML.AppendData("BedNumber", strBedNumber)
                Call mobjXML.AppendData("IsBlock", intIsBlock)
                Call mobjXML.AppendData("IsHighRisk", intIsHighRisk)
                Call mobjXML.AppendData("IsHot", intIsHot)
                Call mobjXML.AppendNode("Item", True)
            Next
            Call mobjXML.AppendNode("List", True)
            strPatientList = mobjXML.XmlText
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With

    GetPatientList = True
End Function

Public Function GetPatientInfo(ByVal lngHwnd As Long, ByVal strPatientId As String, ByVal lngHomePageID As Long, lngBaby As Long, strErrMsg As String, Optional ByVal lngMinTipInfoWith As Long = 3000, Optional ByVal blnShowTip As Boolean = True) As Boolean
'功能：获取病显示病人状态(过敏药物、今日手术、风险、代办事项)
'URL：http://192.168.4.61/Services/pub/UnitHomePage/GetPatientInfo
'参数 ： {"Params":
'{
'"Baby":0,
'"HomepageID":"1",
'"PatientID":"8db9a24d-264d-4402-947f-683bf5529422"
'}}
'返回：Patientinfo:病人信息：TodayOperation：今日手术；TemperatureData：体温数据；BacklogData代办事项；AssessmentData：风险评估
'{"Flag":1,"Msg":"","Output":{"Data":{"Patientinfo":{"AdmissionNumber":2009003434,"InTimes":1,"Sex":"男","Age":"23岁","InsuranceType":"普通病人","Phone":null,"InHospitalDays":2940,"Doctor":"张强","Nurse":"陈丹","FollowUp":null,"Diet":"软食","InDay":"2009-03-23","OutDay":"2009-03-23",
'"Diagnosis":"(S72.301)股骨干骨折,(S72.301)股骨干骨折,(D16.051)肱骨骨软骨瘤","TodayOperation":null,"AllCost":4511.10,"RestCost":null,"AllergyDrugs":null},"TemperatureData":{"Temperature":null,"IsFever":"0"},"BacklogData":[],"transFusionData":{"Temp":[],"All":[]},
'"AssessmentData":[{"ID":"758741bf-ebc1-45cb-a32b-ff03d7ac7914","MarkTabID":"4","UnitID":"01b7e14e-ba9f-43d6-879b-e1b3972d2831","PrintRecordID":null,"RelatedReportID":"604642ce-71a7-40f1-bf3b-6670758c2fd9","ReportName":"HWH","FormName":"Braden评分量表","MarkUser":"李兴贵","EditTime":null,"MarkTime":"2016-06-24 14:28:56","MarkResult":6,"FormID":null,"MarkVerdict":null,"QuestionID":null,"AllQuestionID":"129,130,131,133","DiagnosisID":null,"ModifyTime":null,"CreatedDate":"2016
'06-24 14:28:56","CreatedUser":"李兴贵","NursingEvaluatID":null,"PrecautionID":"D7651FC74D64453493BC8B8928102BAE","AllPrecautionID":"525,527,526,528,529,530,537,538,539,540,541,548,549,550,551,552,553","Option":3,"IsRisk":1,"Current":0}]}}}

'接口返回：TRUE调用成功(弹出浮动窗体),False调用失败
    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    Dim strRetrun As String
    Dim i As Long, lngCount As Long
    Dim lngID As Long, lngParentId As Long
    Dim strName As String, strAge As String, strSex As String, strBedNum As String, strDescription As String, strContentName As String, strTime As String
    Dim strID As String
    Dim rsValue As New ADODB.Recordset
    
    With rsValue
        .fields.Append "id", adBigInt, , adFldKeyColumn
        .fields.Append "parent_id", adBigInt
        .fields.Append "content", adVarChar, 4000
        .fields.Append "color", adBigInt, 10
        .open
    End With
    
    If UserLogin(strErrMsg, True) = False Then Exit Function
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    strUrl = "http://" & gstrIntergrateIP & "/Services/pub/UnitHomePage/GetPatientInfo"
    strParams = "{""Params"":{""Baby"":""" & lngBaby & """,""HomepageID"":""" & lngHomePageID & """,""PatientID"":""" & strPatientId & """}}"
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg, True) = False Then
        Call WriteBusinessLOG("获取病人状态信息", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("获取病人状态信息", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
            '过敏药物
            lngID = lngID + 1
            lngParentId = lngID
            rsValue.AddNew
            rsValue("id").Value = lngParentId
            rsValue("content").Value = "过敏药物"
            rsValue("color").Value = vbBlue
            
            If Not IsNull(.Eval("json.Output.Data.Patientinfo.AllergyDrugs")) Then
                lngCount = .Eval("json.Output.Data.Patientinfo.AllergyDrugs").length
                For i = 0 To lngCount - 1
                    strName = "" & .Eval("json.Output.Data.Patientinfo.AllergyDrugs[" & i & "]")
                    lngID = lngID + 1
                    rsValue.AddNew
                    rsValue("id").Value = lngID
                    rsValue("parent_id").Value = lngParentId
                    rsValue("content").Value = strName
                    rsValue("color").Value = vbRed
                Next i
            Else
                strName = "暂无"
                lngID = lngID + 1
                rsValue.AddNew
                rsValue("id").Value = lngID
                rsValue("parent_id").Value = lngParentId
                rsValue("content").Value = strName
                rsValue("color").Value = vbRed
            End If
            
            '今日手术
            lngID = lngID + 1
            lngParentId = lngID
            rsValue.AddNew
            rsValue("id").Value = lngParentId
            rsValue("content").Value = "今日手术"
            rsValue("color").Value = vbBlue
            
            If Not IsNull(.Eval("json.Output.Data.Patientinfo.TodayOperation")) Then
                lngCount = .Eval("json.Output.Data.Patientinfo.TodayOperation").length
                For i = 0 To lngCount - 1
                    strName = "" & .Eval("json.Output.Data.Patientinfo.TodayOperation[" & i & "]")
                    lngID = lngID + 1
                    rsValue.AddNew
                    rsValue("id").Value = lngID
                    rsValue("parent_id").Value = lngParentId
                    rsValue("content").Value = strName
                    rsValue("color").Value = vbBlack
                Next i
            Else
                strName = "暂无"
                lngID = lngID + 1
                rsValue.AddNew
                rsValue("id").Value = lngID
                rsValue("parent_id").Value = lngParentId
                rsValue("content").Value = strName
                rsValue("color").Value = vbBlack
            End If
            
            '风险
            lngID = lngID + 1
            lngParentId = lngID
            rsValue.AddNew
            rsValue("id").Value = lngParentId
            rsValue("content").Value = "风险评估"
            rsValue("color").Value = vbBlue
            If Not IsNull(.Eval("json.Output.Data.AssessmentData")) Then
                lngCount = .Eval("json.Output.Data.AssessmentData").length
                For i = 0 To lngCount - 1
                    strName = "" & .Eval("json.Output.Data.AssessmentData[" & i & "].FormName") & Space(4) & .Eval("json.Output.Data.AssessmentData[" & i & "].MarkVerdict")
                    lngID = lngID + 1
                    rsValue.AddNew
                    rsValue("id").Value = lngID
                    rsValue("parent_id").Value = lngParentId
                    rsValue("content").Value = strName
                    rsValue("color").Value = vbBlack
                Next i
            Else
                strName = "暂无"
                lngID = lngID + 1
                rsValue.AddNew
                rsValue("id").Value = lngID
                rsValue("parent_id").Value = lngParentId
                rsValue("content").Value = strName
                rsValue("color").Value = vbBlack
            End If
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With
    
    '获取代办事项
    strUrl = "http://" & gstrIntergrateIP & "/services/flow/Task/GetTaskByPatientNo"
    strParams = "{""Params"":{""PatientId"":""" & strPatientId & """}}"
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg, True) = True Then
        strRetrun = oXmlHttp.responseText
        Call WriteBusinessLOG("获取代办事项", strUrl, strParams, strRetrun)
        With gobjScriptControl
            .Language = "javascript"
            .AddCode "var json = " & strRetrun & ";"
            If .Eval("json.Flag") = "1" Then
                lngID = lngID + 1
                lngParentId = lngID
                rsValue.AddNew
                rsValue("id").Value = lngParentId
                rsValue("content").Value = "待办事项"
                rsValue("color").Value = vbBlue
                lngCount = Val("" & .Eval("json.Output.Data.List").length)
                If lngCount > 0 Then
                    For i = 0 To lngCount - 1
                        strName = "" & .Eval("json.Output.Data.List[" & i & "].content.Name")
                        lngID = lngID + 1
                        rsValue.AddNew
                        rsValue("id").Value = lngID
                        rsValue("parent_id").Value = lngParentId
                        rsValue("content").Value = strName
                        rsValue("color").Value = vbBlack
                    Next
                Else
                    strName = "暂无"
                    lngID = lngID + 1
                    rsValue.AddNew
                    rsValue("id").Value = lngID
                    rsValue("parent_id").Value = lngParentId
                    rsValue("content").Value = strName
                    rsValue("color").Value = vbBlack
                End If
            End If
        End With
    Else
        Call WriteBusinessLOG("获取代办事项失败", strUrl, strParams)
    End If
    If blnShowTip = True Then
        rsValue.Filter = ""
        Call frmPatiInfo.ShowTipInfo(lngHwnd, rsValue, lngMinTipInfoWith)
    End If
    rsValue.Filter = ""
    strErrMsg = ""
    Do While Not rsValue.EOF
        strErrMsg = IIf(strErrMsg = "", "", strErrMsg & "<Split>") & rsValue("id").Value & "'" & rsValue("parent_id").Value & "'" & rsValue("content").Value & "'" & rsValue("color").Value
        rsValue.MoveNext
    Loop
    GetPatientInfo = True
End Function

Public Function ShowPaitentInfo(ByVal lngHwnd As Long, ByVal strTipInfo As String, Optional ByVal lngMinTipInfoWith As Long = 3000) As Boolean
    '功能：可根据之前GetPatientInfo获取的病人状态生成的字符串，直接显示病人信息提示窗口（注意：需要保证之前存储格式完整性）
    Dim rsValue As New ADODB.Recordset
    Dim arrInfo, arrItem
    Dim i As Integer, j As Integer
    
    If strTipInfo = "" Then Exit Function
    
    On Error GoTo ErrHand
    With rsValue
        .fields.Append "id", adBigInt, , adFldKeyColumn
        .fields.Append "parent_id", adBigInt
        .fields.Append "content", adVarChar, 4000
        .fields.Append "color", adBigInt, 10
        .open
    End With
    arrInfo = Split(strTipInfo, "<Split>")
    For i = 0 To UBound(arrInfo)
        arrItem = Split(arrInfo(i), "'")
        rsValue.AddNew
        rsValue("id").Value = Val(arrItem(0))
        rsValue("parent_id").Value = Val(arrItem(1))
        rsValue("content").Value = CStr(arrItem(2))
        rsValue("color").Value = Val(arrItem(3))
        rsValue.Update
    Next
    rsValue.Filter = ""
    Call frmPatiInfo.ShowTipInfo(lngHwnd, rsValue, lngMinTipInfoWith)
    
    ShowPaitentInfo = True
    Exit Function
ErrFaild:
    rsValue.AddNew
    rsValue("id").Value = 1
    rsValue("content").Value = "出错啦"
    rsValue("color").Value = vbBlue
    rsValue.Update
    rsValue.AddNew
    rsValue("id").Value = 2
    rsValue("parent_id").Value = 1
    rsValue("content").Value = strTipInfo
    rsValue("color").Value = vbRed
    rsValue.Update
    If rsValue.RecordCount > 0 Then rsValue.MoveFirst
    Call frmPatiInfo.ShowTipInfo(lngHwnd, rsValue, lngMinTipInfoWith)
    Exit Function
ErrHand:
    If Err <> 0 Then Err.Clear
    GoTo ErrFaild
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------
'消息提醒
'------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetMsg(ByVal strRUnitID As String, strMsg As String, strErrMsg As String) As Boolean
'功能：读取消息信息
'URL：http://192.168.4.61/Services/sys/message/GetMsg
'参数：无
'返回：
'{"Flag":1,"Msg":"","Output":{"Data":[{"ID":"7d79dae5-eed0-47b9-814f-2527b5dd2108","fromUser":"Server","toUser":"ZLHIS","type":2,"content":{"mutil":true,"form":"3845ba97-9d8a-4cca-bfb2-e30e58a324c2","version":"1.6","type":1,"FlowId":"9968ce05551a415b9986c786735a0c9b","Name":"填写交接记录单(转出)","NodeId":"555f903858154022b1b09eb3f3d1f427","CreateTime":"2017-04-21 11:44:54","Description":"通知：▲64床 李海粟1号  转出，请填写表单","UnitId":"62402223-1b32-457c-8260-d97627bd0404","PatientId":"89f45207-47c0-46a6-bec6-565adb98bdb0","PatientName":"李海粟1号","BedNumber":"▲64","Sex":"男","Age":"17岁"}}]}}

'接口返回：TRUE调用成功(strMsg消息信息)；FALSE接口调用失败（strErrMsg错误信息）
    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    Dim strRetrun As String
    Dim i As Long, lngCount As Long
    Dim strName As String, strAge As String, strSex As String, strBedNum As String, strDescription As String, strContentName As String, strTime As String
    Dim strUnitID As String, strPatientId As String, strRelatID As String, strInTImes As String, strBabyNum As String
    Dim strID As String, strTOUser As String
    
    If UserLogin(strErrMsg, True) = False Then Exit Function
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    strUrl = "http://" & gstrIntergrateIP & "/Services/sys/message/GetMsg"
    If strRUnitID = "" Then
        strParams = ""
    Else
        strParams = "{""Params"":{""DeptRelatID"":""" & strRUnitID & """}}"
    End If
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg, True) = False Then
        Call WriteBusinessLOG("读取消息信息", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("读取消息信息", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
            strMsg = ""
            lngCount = .Eval("json.Output.Data").length
            If lngCount <= 0 Then GetMsg = True: Exit Function
            mobjXML.ClearXmlText
            Call mobjXML.AppendNode("List")
            For i = 0 To lngCount - 1
                strID = "" & .Eval("json.Output.Data[" & i & "].ID")
                strTOUser = "" & .Eval("json.Output.Data[" & i & "].toUser")
                strName = "" & .Eval("json.Output.Data[" & i & "].content.PatientName")
                strAge = "" & .Eval("json.Output.Data[" & i & "].content.Age")
                strSex = "" & .Eval("json.Output.Data[" & i & "].content.Sex")
                strBedNum = "" & .Eval("json.Output.Data[" & i & "].content.BedNumber")
                strContentName = "" & .Eval("json.Output.Data[" & i & "].content.Name")
                strDescription = "" & .Eval("json.Output.Data[" & i & "].content.Description")
                strTime = "" & .Eval("json.Output.Data[" & i & "].content.CreateTime")
                strUnitID = "" & .Eval("json.Output.Data[" & i & "].content.UnitId")
                strPatientId = "" & .Eval("json.Output.Data[" & i & "].content.PatientId")
                strRelatID = "" & .Eval("json.Output.Data[" & i & "].content.RelatID")
                strInTImes = "" & .Eval("json.Output.Data[" & i & "].content.InTimes")
                strBabyNum = "" & .Eval("json.Output.Data[" & i & "].content.BabyNum")
                
                Call mobjXML.AppendNode("Msg")
                Call mobjXML.AppendData("ID", strID)
                Call mobjXML.AppendData("PatientName", strName)
                Call mobjXML.AppendData("Age", strAge)
                Call mobjXML.AppendData("Sex", strSex)
                Call mobjXML.AppendData("BedNumber", strBedNum)
                Call mobjXML.AppendData("ContentName", strContentName)
                Call mobjXML.AppendData("Description", strDescription)
                Call mobjXML.AppendData("CreateTime", strTime)
                Call mobjXML.AppendData("UnitId", strUnitID)
                Call mobjXML.AppendData("PatientId", strPatientId)
                Call mobjXML.AppendData("RelatID", strRelatID)
                Call mobjXML.AppendData("InTimes", strInTImes)
                Call mobjXML.AppendData("BabyNum", strBabyNum)
                Call mobjXML.AppendData("ToUser", strTOUser)
                Call mobjXML.AppendNode("Msg", True)
            Next i
            Call mobjXML.AppendNode("List", True)
            strMsg = mobjXML.XmlText
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With
    GetMsg = True
End Function

Public Function ReplyMsg(ByVal strMsgIDs As String, strErrMsg As String) As Boolean
'功能：设置消息为已读取状态
'URL：http://192.168.4.61/Services/sys/message/ReplyMsg
'参数：{"Params":{Ids:"1222,51555"}}
'返回：{"Flag":1,"Msg":"1","Output":""}
    Dim oXmlHttp As XMLHTTP
    Dim strUrl As String, strParams As String
    Dim strRetrun As String
    
    If UserLogin(strErrMsg, True) = False Then Exit Function
    If gobjScriptControl Is Nothing Then Set gobjScriptControl = New MSScriptControl.ScriptControl
    strUrl = "http://" & gstrIntergrateIP & "/Services/sys/message/ReplyMsg" '
    strParams = "{""Params"":{""Ids"":""" & strMsgIDs & """}}"
    If SendPostUrl(strUrl, strParams, oXmlHttp, strErrMsg, True) = False Then
        Call WriteBusinessLOG("设置消息信息", strUrl, strParams, strErrMsg)
        Exit Function
    End If
    strRetrun = oXmlHttp.responseText
    Call WriteBusinessLOG("设置消息信息", strUrl, strParams, strRetrun)
    With gobjScriptControl
        .Language = "javascript"
        .AddCode "var json = " & strRetrun & ";"
        If .Eval("json.Flag") = "1" Then
        Else
            strErrMsg = .Eval("json.Msg")
            Exit Function
        End If
    End With
    ReplyMsg = True
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub UnloadWebKixt()
'功能：卸载webkixt控件进程(调用主exe程序关闭时调用)
    Call frmUnloadWebKixt.CloseWebKitX
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set gobjScriptControl = New MSScriptControl.ScriptControl
    Set mobjXML = New clsXML
End Sub

Private Sub Class_Terminate()
    Set gobjScriptControl = Nothing
    Set gobjComlib = Nothing
    Set mobjXML = Nothing
End Sub
