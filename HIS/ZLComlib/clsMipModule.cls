VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'######################################################################################################################

'API
'----------------------------------------------------------------------------------------------------------------------
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'变量
'----------------------------------------------------------------------------------------------------------------------
Private mlngSystem As Long                                                          '系统编号
Private mlngModule As Long                                                          '模块编号
Private mstrPrivs As String
Private mstrXmlSchemaFile As String
Private mstrComputer As String
Private mclsMipClient As clsMipClient
Private mstrSendInstance As String
Private mstrMsgKey As String
Private mobjFso As FileSystemObject
Private mobjTextStream As TextStream
Private mobjReceiveObject As Object
Private mstrProgram As String
Private mstrXmlVersion As String
Private mobjCommonMessageSchema As Object
Private WithEvents mfrmMipModule As frmMipModule
Attribute mfrmMipModule.VB_VarHelpID = -1

'事件
'----------------------------------------------------------------------------------------------------------------------
Public Event OpenReport(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
Public Event OpenModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
Public Event OpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)
Public Event ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
Public Event ConnectStateChanged(ByVal IsConnected As Boolean)

'######################################################################################################################
'公共方法属性
Public Function InitMessage(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPrivs As String, Optional ByVal strProgram As String) As Boolean
'    '******************************************************************************************************************
'    '功能：初始化模块消息（初始化模块内变量、记录模块在线等）
'    '参数：lngSystem-系统编号；lngModule-模块编号
'    '返回：初始化成功返回True；否则返回False
'    '******************************************************************************************************************
'    Dim strSchema As String
'    Dim strSQL As String
'    Dim rsTmp As New ADODB.Recordset
'
'    On Error GoTo errHand
'
'    '------------------------------------------------------------------------------------------------------------------
'    '初始化
'    If gclsMipClient.IsConnect = False Then Exit Function
'    mstrXmlVersion = GetXMLVersion
'    Set mobjReceiveObject = Nothing
'
'    mlngSystem = lngSystem
'    mlngModule = lngModule
'    mstrPrivs = strPrivs
'    mstrProgram = strProgram
'
'    Set mfrmMipModule = New frmMipModule
'    Set mobjFso = New FileSystemObject
'    If mlngModule = 0 Then
'
'        Set mobjCommonMessageSchema = CreateObject("MSXML2.XMLSchemaCache" & mstrXmlVersion)
''        Set mobjCommonMessageSchema = New MSXML2.XMLSchemaCache60
'        mstrXmlSchemaFile = App.Path & "\zlMipCommonMesage.xsd"
'        Set mobjTextStream = mobjFso.CreateTextFile(mstrXmlSchemaFile, True)
'        Call mobjTextStream.Write(GetCommonSchema)
'        mobjTextStream.Close
'    End If
'
'    mstrComputer = AnalyseComputer
'
'    InitMessage = LoginModule(lngSystem, lngModule, mstrSendInstance)
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    InitMessage = False
End Function

Public Function CloseMessage() As Boolean
'    '******************************************************************************************************************
'    '功能：关闭模块消息（绡卸载模块内变量、记录模块离线等）
'    '返回：关闭成功返回True；否则返回False
'    '******************************************************************************************************************
'
'    CloseMessage = LogoutModule(mlngSystem, mlngModule, mstrSendInstance)
'
'    If Not (mfrmMipModule Is Nothing) Then
'        Unload mfrmMipModule
'        Set mfrmMipModule = Nothing
'    End If
End Function

Public Property Get IsConnect() As Boolean
'    '******************************************************************************************************************
'    '功能：判断消息服务平台是否连接
'    '返回：已连接返回True；否则返回False
'    '******************************************************************************************************************
'    If Not (mclsMipClient Is Nothing) Then
'        IsConnect = mclsMipClient.IsConnect
'    End If
End Property

Public Function CommitMessage(ByVal strMsgIdentity As String, ByVal strMsgContent As String, Optional ByRef strMsgResult As String) As Boolean
'    '******************************************************************************************************************
'    '功能：提交消息到ZLHIS导航台消息收发中心
'    '参数：strMsgIdentity - 消息标识；strMsgContent - 消息内容(XML)；strMsgResult - 返回XML结果
'    '返回：成功返回True；否则返回False
'    '*****************************************************************************************************************
'    If Not (mclsMipClient Is Nothing) Then
'        CommitMessage = mclsMipClient.SendMessage(strMsgIdentity, strMsgContent, strMsgResult, mlngSystem, mlngModule, mstrSendInstance, mstrProgram)
'    End If
End Function

Public Function ShowMessage(ByVal strMsgKey As String, ByVal strShowContent As String, Optional ByVal strShowTitle As String, Optional ByVal strShowLinkTitle As String, Optional ByVal strShowLinkPara As String) As Boolean
'    '******************************************************************************************************************
'    '功能：在操作系统任务栏上冒泡提醒消息
'    '参数：strMsgKey-消息id;strShowContent-显示内容；strShowTitle-冒泡提醒标题；strShowLinkTitle-链接显示标题；strShowLinkPara-链接相关参数
'    '返回：显示成功返回True；否则返回False
'    '*****************************************************************************************************************
'    Dim strXml As String
'
'    If Not (mclsMipClient Is Nothing) Then
'
'        strXml = "<?xml version=""1.0"" encoding=""gbk""?>" & vbCrLf
'        strXml = strXml & "<message>" & vbCrLf
'
'        strXml = strXml & "<topic>" & strShowTitle & "</topic>"
'        strXml = strXml & "<text>" & strShowContent & "</text>"
'        strXml = strXml & "<link>"
'        strXml = strXml & "     <type>99</type>"
'        strXml = strXml & "     <title>" & strShowLinkTitle & "</title>"
'        strXml = strXml & "     <para>" & strShowLinkPara & "</para>"
'        strXml = strXml & "</link>"
'
'        strXml = strXml & "</message>"
'
'        ShowMessage = mclsMipClient.ShowMessage(strMsgKey, strXml, mstrSendInstance)
'    End If
'
End Function

Public Property Set ReceiveObject(objReceiveObject As Object)
'    '******************************************************************************************************************
'    '功能：
'    '返回：
'    '******************************************************************************************************************
'    Set mobjReceiveObject = objReceiveObject
End Property

Public Property Get Instance() As String
'    '******************************************************************************************************************
'    '功能：获取实例化此类时的实例值
'    '返回：
'    '******************************************************************************************************************
'    Instance = mstrSendInstance
End Property

Public Property Get Program() As String
'    '******************************************************************************************************************
'    '功能：
'    '返回：
'    '******************************************************************************************************************
'    Program = mstrProgram
End Property
'
''######################################################################################################################
''友元方法、属性，在同一个工程内可以访问
'Friend Function ValidCommonMessageSchema(ByVal strXmlMessage As String, Optional ByRef strErrorReason As String) As Boolean
'    '******************************************************************************************************************
'    '功能：校验XML格式
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim strFile As String
'    Dim objXmlSchema As Object
'    Dim objXmlMessage As Object
'
'    Set objXmlSchema = CreateObject("MSXML2.XMLSchemaCache" & mstrXmlVersion)
''    Set objXmlSchema = New MSXML2.XMLSchemaCache60
'    objXmlSchema.Add "", mstrXmlSchemaFile
'
'    Set objXmlMessage = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
''    Set objXmlMessage = New MSXML2.DOMDocument60
'    objXmlMessage.async = False
'    objXmlMessage.validateOnParse = True
'    objXmlMessage.resolveExternals = False
'    Set objXmlMessage.schemas = objXmlSchema
'
''    strFile = "zlCommonMessage.xml"
'    strFile = App.Path & "\" & mobjFso.GetTempName & ".xml"
'
'    strXmlMessage = Replace(LCase(strXmlMessage), "encoding=""utf-8""?", "encoding=""gbk""?")
'
'    Set mobjTextStream = mobjFso.CreateTextFile(strFile, True)
'    mobjTextStream.Write strXmlMessage
'    mobjTextStream.Close
'
'    Call objXmlMessage.Load(strFile)
'    Call objXmlMessage.Validate
'    If objXmlMessage.parseError.errorCode <> 0 Then
'        strErrorReason = objXmlMessage.parseError.reason
'        ValidCommonMessageSchema = False
'    Else
'        ValidCommonMessageSchema = True
'    End If
'
'    Call mobjFso.DeleteFile(strFile, True)
'End Function
'
'Friend Sub ReceiveMessage(ByVal strMessageItemKey As String, ByVal strMsgContent As String)
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim objEventPara As Collection
'
'    On Error GoTo errHand
'
'    If Not (mobjReceiveObject Is Nothing) Then
'        Set objEventPara = New Collection
'        Call objEventPara.Add(strMessageItemKey)
'        Call objEventPara.Add(strMsgContent)
'
'        On Error Resume Next
'        Call mobjReceiveObject.ActionEvent("ReceiveMessage", objEventPara)
'
'        Call mobjReceiveObject.ReceiveMessage(strMessageItemKey, strMsgContent)
'    Else
'        RaiseEvent ReceiveMessage(strMessageItemKey, strMsgContent)
'    End If
'
'    Exit Sub
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'
'End Sub
'
'Friend Sub OpenLink(ByVal strMsgKey As String, ByVal bytLinkType As Byte, ByVal strLinkPara As String)
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    mstrMsgKey = strMsgKey
'    If Not (mfrmMipModule Is Nothing) Then
'        Call mfrmMipModule.OpenLink(bytLinkType, strLinkPara)
'    End If
'End Sub
'
'Friend Sub ConnectStateChanged(ByVal IsConnected As Boolean)
'    '******************************************************************************************************************
'    '功能：消息服务平台连接状态变化
'    '参数：IsConnected-是否已经连接
'    '******************************************************************************************************************
'    Dim objEventPara As Collection
'
'    If Not (mobjReceiveObject Is Nothing) Then
'        Set objEventPara = New Collection
'        Call objEventPara.Add(IsConnected)
'        Call mobjReceiveObject.ActionEvent("ConnectStateChanged", objEventPara)
'    Else
'        RaiseEvent ConnectStateChanged(IsConnected)
'    End If
'
'End Sub
'
'Friend Property Get System() As Long
'    '属性:系统编号
'    System = mlngSystem
'End Property
'
'Friend Property Get Module() As Long
'    '属性:模块编号
'    Module = mlngModule
'End Property
'
'Friend Property Get Privs() As String
'    '属性:模块权限
'    Privs = mstrPrivs
'End Property
'
'Friend Property Set MipClient(objData As Object)
'    Set mclsMipClient = objData
'End Property
'
'Friend Property Get MipClient() As Object
'    Set MipClient = mclsMipClient
'End Property
'
''######################################################################################################################
''私有方法、属性
'
'Private Function LoginModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strInstance As String) As Boolean
'    '******************************************************************************************************************
'    '功能：连接消息服务平台
'    '参数：无
'    '说明：成功返回True,否则返回False
'    '******************************************************************************************************************
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    If gclsMipClient.IsConnect = False Then Exit Function
'
'    strSQL = "zltools.zlMip_OnLineModule_Insert('" & mstrComputer & "'," & lngSystem & "," & lngModule & ",'" & strInstance & "')"
'
'    Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "LoginModule")
'
'    LoginModule = True
'
'    Exit Function
'errHand:
'
'End Function
'
'Private Function LogoutModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strInstance As String) As Boolean
'    '******************************************************************************************************************
'    '功能：连接消息服务平台
'    '参数：无
'    '说明：成功返回True,否则返回False
'    '******************************************************************************************************************
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    If gclsMipClient.IsConnect = False Then Exit Function
'
'    strSQL = "zltools.zlMip_OnLineModule_Delete('" & mstrComputer & "'," & lngSystem & "," & lngModule & ",'" & strInstance & "')"
'
'    Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "LogoutModule")
'
'    LogoutModule = True
'
'    Exit Function
'errHand:
'
'End Function
'
'Private Function GetCommonSchema() As String
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'
'    Dim strSchema As String
'
'    '"<?xml version=""1.0"" encoding=""gbk""?>" & vbNewLine & _
'
'    strSchema = _
'        "<?xml version=""1.0"" encoding=""utf-8""?>" & vbNewLine & _
'        "<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema""  elementFormDefault=""qualified"" attributeFormDefault=""unqualified"">" & vbNewLine & _
'            "<xs:element name=""message"">" & vbNewLine & _
'                "<xs:complexType>" & vbNewLine & _
'                    "<xs:sequence>"
'
'    strSchema = strSchema & _
'                        "<xs:element name=""head"" minOccurs=""1"" maxOccurs=""1"">" & vbNewLine & _
'                            "<xs:complexType>" & vbNewLine & _
'                                "<xs:sequence>" & vbNewLine & _
'                                    "<xs:element name=""zlhis_pub_004"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""msg_id"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""msg_item_code"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_station_name"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_station_ip"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_program"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_instance"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_system_code"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_module_code"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_mipuser"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_hisuser"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""send_time"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                "</xs:sequence>" & vbNewLine & _
'                            "</xs:complexType>" & vbNewLine & _
'                        "</xs:element>" & vbNewLine
'
'    strSchema = strSchema & _
'                        "<xs:element name=""topic"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1""/>" & vbNewLine & _
'                        "<xs:element name=""text"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine
'
'    strSchema = strSchema & _
'                        "<xs:element name=""link"" minOccurs=""0"" maxOccurs=""1"">" & vbNewLine & _
'                            "<xs:complexType>" & vbNewLine & _
'                                "<xs:sequence>" & vbNewLine & _
'                                    "<xs:element name=""type"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""title"" type=""xs:string"" minOccurs=""1"" maxOccurs=""1""/>" & vbNewLine & _
'                                    "<xs:element name=""para"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1""/>" & vbNewLine & _
'                                "</xs:sequence>" & vbNewLine & _
'                            "</xs:complexType>" & vbNewLine & _
'                        "</xs:element>" & vbNewLine
'
'    strSchema = strSchema & _
'                    "</xs:sequence>" & vbNewLine & _
'                "</xs:complexType>" & vbNewLine & _
'            "</xs:element>" & vbNewLine & _
'        "</xs:schema>"
'
'    GetCommonSchema = strSchema
'End Function
'
'Private Function AnalyseComputer() As String
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim strComputer As String * 256
'
'    Call GetComputerName(strComputer, 255)
'    AnalyseComputer = strComputer
'    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
'End Function
'
'Private Sub Class_Initialize()
'    mstrSendInstance = GetGUID
'End Sub
'
'Private Sub Class_Terminate()
'
'    If Not (mobjFso Is Nothing) Then Set mobjFso = Nothing
'    If Not (mobjTextStream Is Nothing) Then Set mobjTextStream = Nothing
'    If Not (mobjCommonMessageSchema Is Nothing) Then Set mobjCommonMessageSchema = Nothing
'    If Not (mfrmMipModule Is Nothing) Then
'        Unload mfrmMipModule
'        Set mfrmMipModule = Nothing
'    End If
'    If Not (mobjReceiveObject Is Nothing) Then Set mobjReceiveObject = Nothing
'
'End Sub
'
'Private Sub mfrmMipModule_OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)
'
'    Dim rsTmp As ADODB.Recordset
'    Dim strSQL As String
'    Dim varLinkPara As Variant
'    Dim strLink As String
'    Dim objEventPara As Collection
'
'    Select Case bytLinkType
'    '------------------------------------------------------------------------------------------------------------------
'    Case 1
'        '报表
'        strSQL = "Select 系统,程序id From zlreports a Where a.编号=[1]"
'        varLinkPara = Split(strLinkPara, "|")
'        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, gstrSysName, CStr(varLinkPara(0)))
'        If rsTmp.BOF = False Then
'            If Not (mobjReceiveObject Is Nothing) Then
'                Set objEventPara = New Collection
'                Call objEventPara.Add(bytLinkType)
'                Call objEventPara.Add(gobjComLib.zlCommFun.NVL(rsTmp("系统").value, 0))
'                Call objEventPara.Add(gobjComLib.zlCommFun.NVL(rsTmp("程序id").value, 0))
'                If UBound(varLinkPara) >= 1 Then
'                    strLinkPara = Mid(strLinkPara, InStr(strLinkPara, "|") + 1)
'                    Call objEventPara.Add(strLinkPara)
'                Else
'                    Call objEventPara.Add("")
'                End If
'                Call mobjReceiveObject.ActionEvent("OpenReport", objEventPara)
'            Else
'                If UBound(varLinkPara) >= 1 Then
'                    strLinkPara = Mid(strLinkPara, InStr(strLinkPara, "|") + 1)
'                    RaiseEvent OpenReport(gobjComLib.zlCommFun.NVL(rsTmp("系统").value, 0), gobjComLib.zlCommFun.NVL(rsTmp("程序id").value, 0), strLinkPara)
'                Else
'                    RaiseEvent OpenReport(gobjComLib.zlCommFun.NVL(rsTmp("系统").value, 0), gobjComLib.zlCommFun.NVL(rsTmp("程序id").value, 0), "")
'                End If
'            End If
'        End If
'    '------------------------------------------------------------------------------------------------------------------
'    Case 2
'        '模块
'        varLinkPara = Split(strLinkPara, ";")
'        If UBound(varLinkPara) = 1 Then
'            If Not (mobjReceiveObject Is Nothing) Then
'                Set objEventPara = New Collection
'                Call objEventPara.Add(bytLinkType)
'                Call objEventPara.Add(Val(varLinkPara(0)))
'                Call objEventPara.Add(Val(varLinkPara(1)))
'                Call mobjReceiveObject.ActionEvent("OpenModule", objEventPara)
'            Else
'                RaiseEvent OpenModule(Val(varLinkPara(0)), Val(varLinkPara(1)), "")
'            End If
'        End If
'    '------------------------------------------------------------------------------------------------------------------
'    Case 99             '自定义
'        If Not (mobjReceiveObject Is Nothing) Then
'            Set objEventPara = New Collection
'            Call objEventPara.Add(mstrMsgKey)
'            Call objEventPara.Add(strLinkPara)
'            Call mobjReceiveObject.ActionEvent("OpenLink", objEventPara)
'        Else
'            RaiseEvent OpenLink(mstrMsgKey, strLinkPara)
'        End If
'    End Select
'End Sub
'
'Private Function GetXMLNode(ByRef rsXML As ADODB.Recordset, ByVal objNodeList As Object, ByRef lngRow As Long, ByVal lngParentRow As Long) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim objChildNodeList As Object
'    Dim objNode As Object
'
'    On Error GoTo errHand
'
'    For Each objNode In objNodeList
'
'        lngRow = lngRow + 1
'
'        rsXML.AddNew
'        rsXML("id").value = lngRow
'        If lngParentRow > 0 Then rsXML("parent_id").value = lngParentRow
'        rsXML("node_name").value = objNode.nodeName
'
'        If objNode.childNodes.Length = 1 Then
'            rsXML("node_value").value = objNode.Text
'        Else
'            rsXML("node_value").value = ""
'            Set objChildNodeList = objNode.childNodes
'            Call GetXMLNode(rsXML, objChildNodeList, lngRow, lngRow)
'        End If
'    Next
'
'    GetXMLNode = True
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    GetXMLNode = False
'End Function
'
'Private Function GetXMLVersion() As String
'
'    Dim varXMLVersion As Variant
'    Dim strXMLVer As String
'    Dim intLoop As Integer
'    Dim objXML As Object
'
'    On Error GoTo errHand
'
'    varXMLVersion = Split(".6.0,.4.0", ",")
'
'    On Error Resume Next
'    If gobjComLib.OS.IsDesinMode = True Or gobjComLib.zlRegInfo("授权性质") <> "1" Then
'        For intLoop = 0 To UBound(varXMLVersion)
'            Err = 0
'            Set objXML = CreateObject("MSXML2.DOMDocument" & varXMLVersion(intLoop))
'            If Err = 0 Then
'                strXMLVer = varXMLVersion(intLoop)
'                Exit For
'            End If
'        Next
'        On Error GoTo errHand
'
'        If strXMLVer = "" Then
'            MsgBox "创建MSXML2.DOMDocument对象失败"
'            Exit Function
'        End If
'    Else
'        strXMLVer = ""
'    End If
'    GetXMLVersion = strXMLVer
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    MsgBox Err.Description
'End Function
