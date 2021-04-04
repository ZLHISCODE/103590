VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mfrmWinsock As frmWinsock
Attribute mfrmWinsock.VB_VarHelpID = -1

Public Event DataArrival(ByVal lngNoticeCode As Long, ByVal intChangeType As Integer, ByVal strTableOwner As String, ByVal TableName As String, ByVal strRowid As String)
Private mlngPort As Long

Private Sub Class_Initialize()
    '初始化类时, 开启UDP端口
    Dim strErrMsg As String
    
    Set mfrmWinsock = New frmWinsock
    
    If StartUdp(strErrMsg) = False Then
        MsgBox """接收数据变动通知""服务启动失败。" & vbNewLine & strErrMsg, , "提示"
    End If
End Sub

Private Sub Class_Terminate()
    Unload mfrmWinsock
    Set mfrmWinsock = Nothing
End Sub

Private Sub mfrmWinsock_DataArrival(ByVal lngNoticeCode As Long, ByVal intChangeType As Integer, ByVal strTableOwner As String, ByVal strTableName As String, ByVal strRowid As String)
    'winsock收到消息后,将事件抛出,由其他业务部件处理
    '返回参数 lngNoticeCode = NoticeCode  intChangeType = 变动类型 strTableOwner = 表所有者 strTableName = 表名  strRowid = 变动行Rowid
    RaiseEvent DataArrival(lngNoticeCode, intChangeType, strTableOwner, strTableName, strRowid)
End Sub

Public Function StartUdp(Optional ByRef strErr As String) As Boolean
    '开启UDP监听,成功返回True
    '如果开启监听中发生错误,返回Flase, strErr  = 错误描述
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '首先判断消息收发器是否打开,如果没有打开,不需要开启端口
    If Not IsServerStart Then
        StartUdp = True
        mlngPort = 0
        Exit Function
    End If
    
    strSql = "Select 消息端口 From Zltools.Zlclientsession Where 会话号 = [1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取导航台UDP端口", glngSessionID)
    
    If rsTmp.RecordCount = 0 Then
        strErr = "未找到当前会话对应的消息端口。"
    Else
        If IsNull(rsTmp!消息端口) Then
            strErr = "未找到当前会话对应的消息端口。"
        Else
            mlngPort = rsTmp!消息端口   '此处获取端口,作用类似于 "会话序号"
            
            '判断参数，是否开启导航台端口
            Set rsTmp = GetZLOptions(9)
            If rsTmp.RecordCount <> 0 Then
                If rsTmp!参数值 = "0" Then
                    mlngPort = 0
                    StartUdp = True
                    Exit Function
                End If
            End If
            
            '如果端口开启失败 ,就将端口修改为0
            If mfrmWinsock.StartUdp(mlngPort, strErr) = False Then
                mlngPort = 0
            Else
                StartUdp = True
            End If
        End If
    End If
    Exit Function
errH:
    strErr = Err.Description
End Function

Public Function IsWinsockStart() As Boolean
    '检查winsock是否成功打开
    
    IsWinsockStart = mlngPort <> 0
End Function

Private Function IsServerStart() As Boolean
    '检查消息收发器是否打开
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim intState As Integer, lngSid As Long
    
    On Error GoTo errH
    '检查消息收发器状态
    Set rsTmp = GetZLOptions(27)
    If rsTmp.RecordCount = 0 Then Exit Function
    If rsTmp!参数值 & "" = "" Then Exit Function
    
    'IP;端口;状态;会话号
    intState = Split(rsTmp!参数值, ";")(2)
    lngSid = Split(rsTmp!参数值, ";")(3)
    
    If intState = 0 Then Exit Function
    
    strSql = "Select 1 From gv$Session Where AUDSID = [1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "检查消息收发器状态", lngSid)
    
    IsServerStart = rsTmp.RecordCount > 0
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Sub ChangePosition(Optional ByVal lng部门ID As Long, Optional ByVal str当前位置 As String, Optional ByRef strErr As String)
'兼容部件处理（将错就错），恢复该方法

    '将当前会话信息保存至表 zlClientSession
    Dim str工作站 As String, lng端口号 As Long
    Dim strSql As String
    
    On Error GoTo errH
    str工作站 = gobjComLib.OS.ComputerName
    lng端口号 = mlngPort
    
    strSql = "Zltools.Zl_Zlclientsession_Save('" & str工作站 & "',0," & lng端口号 & ",'','',''," & lng部门ID & ",'" & str当前位置 & "',1)"
    gobjComLib.zlDatabase.ExecuteProcedure strSql, "保存zlclientsession信息"
    
    Exit Sub
errH:
    strErr = Err.Description
End Sub

Public Sub ChangeSessionInfo(Optional ByVal lng部门ID As Long, Optional ByVal str部门性质 As String, _
    Optional ByVal str人员性质 As String, Optional ByVal str当前位置 As String, _
    Optional ByRef strErr As String)
'兼容部件处理（将错就错），该方法2019.3月新增，同时ChangePosition方法被删除

    '将当前会话信息保存至表 zlClientSession
    Dim str工作站 As String, lng端口号 As Long
    Dim strSql As String
    
    On Error GoTo errH
    str工作站 = gobjComLib.OS.ComputerName
    lng端口号 = mlngPort
    
    strSql = "Zltools.Zl_Zlclientsession_Save('" & str工作站 & "', 0, " & lng端口号 & ", '', '', '', " _
           & lng部门ID & ", '" & str当前位置 & "', 1, '" & str部门性质 & "', '" & str人员性质 & "')"
    gobjComLib.zlDatabase.ExecuteProcedure strSql, "保存zlclientsession信息"
    
    Exit Sub
errH:
    strErr = Err.Description
End Sub

Public Function CheckDcnEnable(ByVal lngNoticeCode As Long) As Boolean
    '功能: 根据传入的NoticeCode检查Dcn是否可用
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 1 From zltools.zlNoticeLists Where NoticeCode = [1] And Status = 1"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "检查DCN配置可用性", lngNoticeCode)
    
    CheckDcnEnable = rsTmp.RecordCount > 0
    Exit Function
errH:
    gobjComLib.ErrCenter
End Function

Public Function CheckDcnState() As Boolean
    '功能:用于实时检查DCN的状态
    CheckDcnState = mfrmWinsock.blnDcnState
End Function

Public Sub ChangeDepts(ByVal strDepts As String, Optional ByRef strErr As String)
    '功能:如果工作站属于多个部门,保存多个部门至zlClientDepts表中
    'strDepts = 多个部门的ID, 各个部门之间用逗号作为间隔
    '如果失败,strErr返回失败信息
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Zltools.Zl_Zlclientdepts_Save(" & glngSessionID & ",'" & strDepts & "')"
    gobjComLib.zlDatabase.ExecuteProcedure strSql, "保存zlclientdepts信息"
    
    Exit Sub
errH:
    strErr = Err.Description
End Sub

