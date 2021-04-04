VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueueMsgCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const M_STR_ROOT_NODE As String = "message"


Private WithEvents mobjMsgProcess As clsMipModule
Attribute mobjMsgProcess.VB_VarHelpID = -1
Private mobjXml As clsXML

Private mblnMsgOpened As Boolean    '是否打开消息中心
Private mlngBusineesType As Long    '业务类型
Private mlngModule As Long          '模块号
Private mstrComputerName As String  '计算机名
Private mblnDebugMsg As Boolean     '启用消息调试

Public Event OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset)
Public Event OnOpenModule(ByVal lngSystem As Long, ByVal lngModule As Long)
Public Event OnOpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)


Private Function GetSendTag() As String
'获取发送标记
    Dim strXml As String
    
    
    strXml = "<send_computer>" & mstrComputerName & "</send_computer>" & vbCrLf & _
             "<send_program>" & mlngModule & "</send_program>"
             
    GetSendTag = strXml
End Function

Private Function SendMsg(ByVal lngQueueId As Long, ByVal strMsgTag As String) As Boolean
'发送消息
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngStartTime As Long
    
    SendMsg = False
    
    If Not mblnMsgOpened Then Exit Function
    
    strSql = "select zl_排队消息_XML内容获取([1], [2]) as XML from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询排队消息内容文本", lngQueueId, strMsgTag)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If mblnDebugMsg Then
        lngStartTime = GetTickCount
    End If
    
    SendMsg = mobjMsgProcess.CommitMessage(strMsgTag, Nvl(rsData!XML) & GetSendTag)
    
    If mblnDebugMsg Then
        OutputDebugString "消息 [" & strMsgTag & "] 处理时长为" & GetTickCount - lngStartTime & "毫秒。"
    End If
    
    If SendMsg = False Then
        Debug.Print "消息" & strMsgTag & "发送失败。"
    End If
End Function


Public Function SendMsg_LineQueue(ByVal lngQueueId As Long) As Boolean
'发送排队消息
    SendMsg_LineQueue = SendMsg(lngQueueId, G_STR_MSG_QUEUE_001)
End Function


Public Function SendMsg_CompleteQueue(ByVal lngQueueId As Long) As Boolean
'发送完成消息
    SendMsg_CompleteQueue = SendMsg(lngQueueId, G_STR_MSG_QUEUE_002)
End Function


Public Function SendMsg_StateSync(ByVal lngQueueId As Long) As Boolean
'状态同步消息
    SendMsg_StateSync = SendMsg(lngQueueId, G_STR_MSG_QUEUE_003)
End Function


Public Function SendMsg_VoicePlay(ByVal lngQueueId As Long) As Boolean
'语音播放消息
    SendMsg_VoicePlay = SendMsg(lngQueueId, G_STR_MSG_QUEUE_004)
End Function


Public Sub ConfigMsgBusinessType(ByVal lngBusinessType As Long)
'初始化消息处理中心
    mlngBusineesType = lngBusinessType
End Sub


Public Sub OpenMsgCenter(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal lngBusineesType As Long, Optional ByVal strPrivs As String = "")
'打开消息中心
    Dim blnInit As Boolean
    Dim blnAdd As Boolean
    
    mlngBusineesType = lngBusineesType
    mlngModule = lngModule
    
    Set mobjMsgProcess = New clsMipModule

    blnInit = mobjMsgProcess.InitMessage(lngSys, lngModule, strPrivs)
    blnAdd = AddMipModule(mobjMsgProcess)
    
    mstrComputerName = AnalyseComputer & "(" & GetGUID & ")"
    
    If blnInit = True And blnAdd = True Then
        mblnMsgOpened = True
    Else
        mblnMsgOpened = False
    End If
    
    mblnDebugMsg = Val(GetSetting("ZLSOFT", "公共模块", "消息调试", 0))
End Sub

Public Sub CloseMsgCenter()
'关闭消息中心
    If Not mblnMsgOpened Then Exit Sub
    
    Call DelMipModule(mobjMsgProcess)
    
    Call mobjMsgProcess.CloseMessage
    Set mobjMsgProcess = Nothing
End Sub


Private Sub Class_Initialize()
'类初始化
    mblnMsgOpened = False
    
    Set mobjXml = New clsXML
End Sub



Private Sub Class_Terminate()
'释放xml处理对象
    Set mobjXml = Nothing
End Sub

Private Sub mobjMsgProcess_OpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)
'提示框连接打开事件
On Error Resume Next
    RaiseEvent OnOpenLink(strMsgKey, strLinkPara)
End Sub

Private Sub mobjMsgProcess_OpenModule(ByVal lngSystem As Long, ByVal lngModule As Long, ByVal strPara As String)
'模块打开事件
On Error Resume Next
    RaiseEvent OnOpenModule(lngSystem, lngModule)
End Sub

Private Sub mobjMsgProcess_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
'消息接收事件
On Error GoTo errHandle
    Dim rsData As ADODB.Recordset
    Dim strValue As String
    
    
    
    '过滤是否排队相关的消息，如果不是，则退出处理
    If strMsgItemIdentity <> G_STR_MSG_QUEUE_001 _
        And strMsgItemIdentity <> G_STR_MSG_QUEUE_002 _
        And strMsgItemIdentity <> G_STR_MSG_QUEUE_003 _
        And strMsgItemIdentity <> G_STR_MSG_QUEUE_004 Then
        Exit Sub
    End If
        
        
    Call mobjXml.OpenXMLDocument(strMsgContent)
    
    '判断业务类型是否相同，如果消息中的业务类型与当前对象的业务类型不同，则不进行消息处理
    If Not mobjXml.GetSingleNodeValue("business_type", strValue, xsNumber) Then
        Debug.Print "获取排队业务类型时失败，终止消息处理。"
        Exit Sub
    End If
    
    If Val(strValue) <> mlngBusineesType Then
        Debug.Print "获取消息的业务类型与当前业务类型不同，忽略消息处理。"
        Exit Sub
    End If
    
    '相同对象发过来的消息不进行处理
    If Not mobjXml.GetSingleNodeValue("send_instance", strValue, xsString) Then
        Debug.Print "获取排队业务来源站点时失败，终止消息处理。"
        Exit Sub
    End If

    If strValue = mobjMsgProcess.Instance Then
        '获取消息的对象与当前对象相同，忽略消息处理
        Debug.Print "获取消息的对象与当前对象相同，忽略消息处理。"
        Exit Sub
    End If
    
    
    If Not mobjXml.GetAllNodeRecord(rsData) Then
        Debug.Print "调用GetAllNodeRecord方法失败，终止消息处理。"
        Exit Sub
    End If
    
    If rsData.RecordCount <= 0 Then
        Debug.Print "未获取到有效的节点数据，终止消息处理。"
        Exit Sub
    End If
    
    '触发事件接收消息
    RaiseEvent OnRecevieMsg(strMsgItemIdentity, strMsgContent, rsData)
    
Exit Sub
errHandle:
    Debug.Print Err.Description
End Sub


