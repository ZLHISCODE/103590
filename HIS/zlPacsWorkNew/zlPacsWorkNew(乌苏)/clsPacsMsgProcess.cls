VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacsMsgProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mobjMsgProcess As clsMipModule
Attribute mobjMsgProcess.VB_VarHelpID = -1
Private mobjXml As clsXML

Private mlngModule As Long
Private mlngCurDeptId As Long
Private mstrPrivs As String
Private mblnMsgOpened As Boolean
Private mblnDebugMsg As Boolean


Public Event OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset, _
                            objMsgPro As clsMipModule, objXML As clsXML)
                            
Public Event OnOpenModule(ByVal lngSystem As Long, ByVal lngModule As Long)
Public Event OnOpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)


Public Property Get XML() As clsXML
    Set XML = mobjXml
End Property

Private Function SendMsg(ByVal lngAdviceID As Long, ByVal strMsgTag As String, Optional strReportId As String = "") As Boolean
'发送消息
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngStartTime As Long
    
    SendMsg = False
    
    If Not mblnMsgOpened Then Exit Function
    
    If Trim(strReportId) = "" Then
        strSQL = "select zl_影像消息_XML内容获取([1], [2], [3]) as xml from dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询危急值消息内容", lngAdviceID, strMsgTag, UserInfo.用户名)
    Else
        strSQL = "select zl_影像消息_XML内容获取([1], [2], [3], [4]) as xml from dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询危急值消息内容", lngAdviceID, strMsgTag, UserInfo.用户名, strReportId)
    End If
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If mblnDebugMsg Then
        lngStartTime = GetTickCount
    End If
    
    SendMsg = mobjMsgProcess.CommitMessage(strMsgTag, Nvl(rsData!XML))
    
    If strMsgTag = G_STR_MSG_ZLHIS_PACS_005 Or strMsgTag = G_STR_MSG_ZLHIS_CIS_005 Then
        SendMsg = zl9comlib.zlDatabase.SendMsg(strMsgTag, Nvl(rsData!XML))
    End If
    
    If mblnDebugMsg Then
        OutputDebugString "消息 [" & strMsgTag & "] 处理时长为" & GetTickCount - lngStartTime & "毫秒。"
    End If
    
    If SendMsg = False Then
        Debug.Print "消息" & strMsgTag & "发送失败。"
    End If
End Function


Public Function Send_Msg_Request(ByVal lngAdviceID As Long) As Boolean
'发送检查申请消息

    Send_Msg_Request = SendMsg(lngAdviceID, G_STR_MSG_ZLHIS_CIS_017)
End Function


Public Function Send_Msg_CancelAdvice(ByVal lngAdviceID As Long) As Boolean
'发送医嘱撤销消息

    Send_Msg_CancelAdvice = SendMsg(lngAdviceID, G_STR_MSG_ZLHIS_CIS_024)
End Function


Public Function Send_Msg_StudyComplete(ByVal lngAdviceID As Long, Optional ByVal strReportId As String = "") As Boolean
'发送检查完成消息
'strReportId,新版报告ID
    Send_Msg_StudyComplete = SendMsg(lngAdviceID, G_STR_MSG_ZLHIS_PACS_001, strReportId)
End Function


Public Function Send_Msg_StateSync(ByVal lngAdviceID As Long) As Boolean
'发送状态同步消息

    Send_Msg_StateSync = SendMsg(lngAdviceID, G_STR_MSG_ZLHIS_PACS_002)
End Function

Public Function Send_Msg_Arrange(ByVal lngAdviceID As Long) As Boolean
'发送安排消息

    Send_Msg_Arrange = SendMsg(lngAdviceID, G_STR_MSG_ZLHIS_CIS_005)
End Function


Public Function Send_Msg_StateCancel(ByVal lngAdviceID As Long) As Boolean
'发送状态回退消息

    Send_Msg_StateCancel = SendMsg(lngAdviceID, G_STR_MSG_ZLHIS_PACS_003)
End Function


Public Function Send_Msg_CancelComplete(ByVal lngAdviceID As Long) As Boolean
'发送状态回退消息

    Send_Msg_CancelComplete = SendMsg(lngAdviceID, G_STR_MSG_ZLHIS_PACS_004)
End Function


Public Function Send_Msg_Critical(ByVal lngAdviceID As Long) As Boolean
'发送危急值通知消息

    Send_Msg_Critical = SendMsg(lngAdviceID, G_STR_MSG_ZLHIS_PACS_005)
End Function

Public Sub ChangeMsgReceiveDept(ByVal lngCurDeptId As Long)
'改变消息接收科室
    mlngCurDeptId = lngCurDeptId
End Sub


Public Sub OpenMsgCenter(ByVal lngModule As Long, ByVal lngCurDeptId As Long, ByVal strPrivs As String)
'初始化消息对象
    Dim blnInit As Boolean
    Dim blnAdd As Boolean
    
    mlngModule = lngModule
    mlngCurDeptId = lngCurDeptId
    mstrPrivs = strPrivs
    
    Set mobjMsgProcess = New clsMipModule
    
    blnInit = mobjMsgProcess.InitMessage(glngSys, mlngModule, strPrivs)
    blnAdd = AddMipModule(mobjMsgProcess)
    
    mblnDebugMsg = Val(GetSetting("ZLSOFT", "公共模块", "消息调试", 0))
    
    If blnInit = True And blnAdd = True Then
        mblnMsgOpened = True
    Else
        mblnMsgOpened = False
    End If
End Sub


Public Sub CloseMsgCenter()
'关闭消息中心
    If mblnMsgOpened = False Then Exit Sub
    
    
    Call DelMipModule(mobjMsgProcess)
    
    Call mobjMsgProcess.CloseMessage
    
    Set mobjMsgProcess = Nothing
    
    mblnMsgOpened = False
End Sub

Private Sub Class_Initialize()
'初始化消息对象
    
    Set mobjXml = New clsXML
    
    mblnMsgOpened = False
End Sub

Private Sub Class_Terminate()
    Call CloseMsgCenter
    
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
'接收到的消息处理
'消息接收事件
On Error GoTo ErrHandle
    Dim rsData As ADODB.Recordset
    Dim strValue As String
        
    '过滤是否影像系统需要接收的消息，如果不是则直接退出
    If strMsgItemIdentity <> G_STR_MSG_ZLHIS_CIS_017 _
        And strMsgItemIdentity <> G_STR_MSG_ZLHIS_CIS_024 _
        And strMsgItemIdentity <> G_STR_MSG_ZLHIS_PACS_001 _
        And strMsgItemIdentity <> G_STR_MSG_ZLHIS_PACS_002 _
        And strMsgItemIdentity <> G_STR_MSG_ZLHIS_PACS_003 _
        And strMsgItemIdentity <> G_STR_MSG_ZLHIS_PACS_004 _
        And strMsgItemIdentity <> G_STR_MSG_ZLHIS_PACS_005 _
        And strMsgItemIdentity <> G_STR_MSG_ZLHIS_CHARGE_003 _
        And strMsgItemIdentity <> G_STR_MSG_ZLHIS_CIS_025 Then
        Exit Sub
    End If
        
        
    Call mobjXml.OpenXMLDocument(strMsgContent)
    


    '读取消息发送方的实例数据
    Call mobjXml.GetSingleNodeValue("send_instance", strValue)

    '如果实例相同，则不接收消息处理
    If mobjMsgProcess.Instance = strValue Then Exit Sub
          
          
    
    '将所有节点转换为ADO数据集
    If Not mobjXml.GetAllNodeRecord(rsData) Then
        Debug.Print "调用GetAllNodeRecord方法失败，终止消息处理。"
        Exit Sub
    End If
    
    If rsData.RecordCount <= 0 Then
        Debug.Print "未获取到有效的节点数据，终止消息处理。"
        Exit Sub
    End If
    
    '判断对应的执行科室节点是否存在，如果存在，则需要判断当前科室与执行科室是否一致
    rsData.Filter = "node_name='study_execute_id'"
    If rsData.RecordCount > 0 Then
        '判断执行部门ID与当前部门ID是否相同
        If Nvl(rsData!node_value) <> mlngCurDeptId Then Exit Sub
    End If
    
    
    '触发事件接收消息
    RaiseEvent OnRecevieMsg(strMsgItemIdentity, strMsgContent, rsData, mobjMsgProcess, mobjXml)
    
Exit Sub
ErrHandle:
    Debug.Print err.Description
End Sub
