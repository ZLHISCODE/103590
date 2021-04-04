VERSION 5.00
Begin VB.Form frmTrayIcon 
   Caption         =   "托盘图标"
   ClientHeight    =   1455
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   2070
   Icon            =   "frmTrayIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   2070
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer tmrBroadCast 
      Left            =   360
      Top             =   360
   End
   Begin VB.Menu mnuRight 
      Caption         =   "右键菜单"
      Begin VB.Menu mnuSetup 
         Caption         =   "配置"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "退出"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjIcon As clsTaskIcon
Attribute mobjIcon.VB_VarHelpID = -1
Private WithEvents mobjMsgCenter As clsQueueMsgCenter
Attribute mobjMsgCenter.VB_VarHelpID = -1

Private mblnUserMsgCenter As Boolean
Private mblnUseSound As Boolean

Private mdtLastVoiceDate As Date
Private mobjQueueManage As New clsQueueOperation
Private mlngLoopInterval As Long
Private mstrComputerName As String
'private

Private Sub Form_Load()
On Error GoTo ErrorHand
    '打开托盘图标
    Set mobjIcon = New clsTaskIcon
    mobjIcon.frmHwnd = Me.hwnd ' hwnd
    mobjIcon.Icon = Me.Icon.Handle
    mobjIcon.Message = "排队显示控制"
    mobjIcon.AddIcon
    
    If gstrCompareVersion < "010.034.000" Then
        mblnUserMsgCenter = False
    Else
        mblnUserMsgCenter = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "启用消息服务中心", 1)) = 1
    End If
    
    mblnUseSound = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "启用语音呼叫", 1)) = 1
    
    '根据参数判断是否启用消息处理
    If mblnUserMsgCenter Then  '判断是否需要使用消息中心
    
        '连接消息中心
        Call gobjComLib.ConnectMip(Me.hwnd)
        
        Set mobjMsgCenter = New clsQueueMsgCenter
        Call mobjMsgCenter.setComLib(gobjComLib)
        Call mobjMsgCenter.OpenMsgCenter(100, 1160, glngBusinessType)  '最后一个参数需要传入实际的业务类型
        
        tmrBroadCast.Tag = 1
        tmrBroadCast.Enabled = False
    Else
        If mblnUseSound Then
            tmrBroadCast.Tag = 0
            tmrBroadCast.Enabled = True
        End If
    End If
    
    Call mobjQueueManage.setComLib(gobjComLib)
    
    '应用语音配置
    Call ApplyVoiceConfig
    
    If Not mblnUserMsgCenter Then StartVoice

    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Public Sub setMsgBusinessType(ByVal lngBusinessType As Long)
    If mobjMsgCenter Is Nothing Then Exit Sub
    Call mobjMsgCenter.ConfigMsgBusinessType(lngBusinessType)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorHand
    mobjIcon.MouseState x
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHand
    '如果启用了消息处理，则需要断开消息平台的连接
    If mblnUserMsgCenter Then
        Call gobjComLib.DisConnectMip
        Call mobjMsgCenter.CloseMsgCenter
    End If
    
    '清除托盘图标
    mobjIcon.DelIcon
    Set mobjIcon = Nothing
    Set mobjMsgCenter = Nothing
    Set mobjQueueManage = Nothing
    
    Call mnuQuit_Click
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'完全退出
Private Sub mnuQuit_Click()
    Dim objForm As Form
On Error GoTo ErrorHand
    '卸载所有窗体
    For Each objForm In Forms
        Unload objForm
    Next
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'打开配置窗口
Private Sub mnuSetup_Click()
On Error GoTo ErrorHand
    frmMain.zlShowMe
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'打开配置窗口
Private Sub mobjIcon_MouseLeftDBClick()
On Error GoTo ErrorHand
    Call mnuSetup_Click
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'显示右键菜单
Private Sub mobjIcon_MouseRightUp()
On Error GoTo ErrorHand
    PopupMenu mnuRight
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub mobjMsgCenter_OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset)
'处理接收到的消息
    Dim strValue As String
    Dim i As Integer
    
    'G_STR_MSG_QUEUE_001:排队消息
    'G_STR_MSG_QUEUE_002:完成消息
    'G_STR_MSG_QUEUE_003:状态改变消息
    'G_STR_MSG_QUEUE_004:请求呼叫消息
On Error GoTo ErrorHand
        
    Select Case strMsgItemIdentity
        Case G_STR_MSG_QUEUE_001, G_STR_MSG_QUEUE_002, G_STR_MSG_QUEUE_003
            If SafeArrayGetDim(gobjStyleWindow) <= 0 Then Exit Sub
            
            '对不同样式所接收的消息进行同步处理......
            For i = 1 To UBound(gobjStyleWindow)
                If Not gobjStyleWindow(i) Is Nothing Then
                    Call gobjStyleWindow(i).ISty_MsgProcess(gobjStyleWindow(i).ISty_WindNo, strMsgItemIdentity, strXmlContext, rsData)
                End If
            Next
            
        Case G_STR_MSG_QUEUE_004
            '进行呼叫处理......
            If mblnUseSound Then
                If PlayVoice(rsData) Then
                    tmrBroadCast.Tag = 1
                    tmrBroadCast.Enabled = False
                Else    '呼叫失败则进行轮询呼叫
                    tmrBroadCast.Tag = 0
                    tmrBroadCast.Enabled = True
                End If
            End If
            
    End Select
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub tmrBroadCast_Timer()
On Error GoTo ErrorHand
    If Val(tmrBroadCast.Tag) = 1 Then Exit Sub
    
    '停止轮训
    Call AbortCall
    
    '调用轮训方法
    Call LoopPlayVoice
    
    If Val(tmrBroadCast.Tag) = 1 Then Exit Sub
    
    '开始轮训
    Call StartCall
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''语音呼叫部分代码''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub StartVoice()
'开始语音播放
    mdtLastVoiceDate = Timer - mlngLoopInterval
    
    tmrBroadCast.Interval = mlngLoopInterval * 1000
    tmrBroadCast.Enabled = True
    tmrBroadCast.Tag = 0
End Sub

Private Sub StopVoice()
'结束语音播放
On Error GoTo errHandle
    tmrBroadCast.Tag = 1
    tmrBroadCast.Enabled = False
    
    If Not mobjQueueManage Is Nothing Then Call mobjQueueManage.StopVoice
Exit Sub
errHandle:
    Debug.Print "StopVoice Err:" & Err.Description
End Sub

Private Sub ApplyVoiceConfig()
'应用语音配置
    Dim str呼叫站点名称 As String
    
   '读取叫号方式
    If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\mlngWindowNo", "播放方式", 1)) = 1 Then
         str呼叫站点名称 = GetSetting("ZLSOFT", G_STR_REGPATH, "远端呼叫站点", "")
         
         If Trim(str呼叫站点名称) = "" Then str呼叫站点名称 = AnalyseComputer
    Else
        str呼叫站点名称 = AnalyseComputer
    End If
    
    mstrComputerName = AnalyseComputer
    
    mobjQueueManage.PlayStation = str呼叫站点名称
    mobjQueueManage.LocalStation = mstrComputerName
    mobjQueueManage.BusinessType = glngBusinessType
    mobjQueueManage.PlayTimeLength = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "语音播放时长", 15))
    mobjQueueManage.PlayCount = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "语音播放次数", 2))
    mobjQueueManage.VoiceType = GetSetting("ZLSOFT", G_STR_REGPATH, "语音类型", "")
    mobjQueueManage.IsPlayHintSound = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "语音呼叫前播放提示音", False))
    mobjQueueManage.PlaySpeed = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "语音播放语速", 0))
    mobjQueueManage.UseVbsPlay = IIf(Val(GetSetting("ZLSOFT", G_STR_REGPATH, "启用VBS自定义呼叫", 1)) = 0, False, True)
    mobjQueueManage.CusVoiceScript = GetSetting("ZLSOFT", G_STR_REGPATH, "VBS脚本", "")
    
    mlngLoopInterval = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "轮询间隔时间", 30))
End Sub

Private Function PlayVoice(ByVal rsData As ADODB.Recordset) As Boolean
'功能：进行语音呼叫
'返回值：呼叫成功返回True,反之返回False
    Dim lngVoiceId As Long
    Dim lngQueueId As Long
    Dim strVoiceContext As String
    
    PlayVoice = False
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    '判断是否是本站点呼叫数据
    rsData.Filter = "node_name='voice_station'"
    If rsData.RecordCount <= 0 Then Exit Function
    If Nvl(rsData!node_value) <> AnalyseComputer Then Exit Function
    
    rsData.Filter = "node_name='voice_context'"
    If rsData.RecordCount > 0 Then strVoiceContext = Nvl(rsData!node_value)
    
    rsData.Filter = "node_name='voice_id'"
    If rsData.RecordCount > 0 Then lngVoiceId = Val(rsData!node_value)
    
    rsData.Filter = "node_name='queue_id'"
    If rsData.RecordCount > 0 Then lngQueueId = Val(rsData!node_value)
    
    '播放语音
    PlayVoice = mobjQueueManage.PlayQueueVoice(mobjMsgCenter, lngVoiceId, lngQueueId, False, strVoiceContext)

    '呼叫成功后删除呼叫过的内容
    Call mobjQueueManage.DelVoiceData(lngVoiceId)
End Function

'轮询呼叫
Private Sub LoopPlayVoice()
    Dim i As Integer
    Dim strSql As String
    Dim lngVoiceId As Long
    Dim lngQueueId As Long
    Dim strVoiceContext As String
    Dim blnAllowQuery As Boolean
    Dim rsVoiceContext As ADODB.Recordset  '待播放的语音数据集
    
    '判断是否需要查询数据库数据
    blnAllowQuery = IIf(rsVoiceContext Is Nothing, True, False)
    
    If Not rsVoiceContext Is Nothing Then
        blnAllowQuery = IIf(rsVoiceContext.RecordCount <= 0 Or rsVoiceContext.EOF, True, False)
    End If
    
    If blnAllowQuery Then
        If Timer < mdtLastVoiceDate + mlngLoopInterval Then Exit Sub
        mdtLastVoiceDate = Timer
        
        '查询需要播放的语音数据
        If gstrCompareVersion < "010.034.000" Then
            strSql = "select id,队列ID,呼叫内容 from 排队语音呼叫  where 站点=[1] order by id"
        Else
            strSql = "select id,队列ID,呼叫内容,生成时间 from 排队语音呼叫  where 站点=[1] order by 生成时间"
        End If
        
        Set rsVoiceContext = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "查询语音呼叫内容", mstrComputerName)
    End If
    
    If rsVoiceContext Is Nothing Then Exit Sub
    If rsVoiceContext.RecordCount <= 0 Or rsVoiceContext.EOF Then Exit Sub
    
    lngVoiceId = Val(Nvl(rsVoiceContext!ID))
    lngQueueId = Val(Nvl(rsVoiceContext!队列ID))
    strVoiceContext = Nvl(rsVoiceContext!呼叫内容)
    
    Call rsVoiceContext.MoveNext
    
    '刷新界面数据信息
    If Not mblnUserMsgCenter Then
        If SafeArrayGetDim(gobjStyleWindow) > 0 Then
            For i = 1 To UBound(gobjStyleWindow)
                Call gobjStyleWindow(i).ISty_RefreshQueueData(lngQueueId)
            Next
        End If
    End If
    
    If lngQueueId <= 0 Then
        '播放自定义的呼叫内容
        Call mobjQueueManage.PlayCustomVoice(lngVoiceId, False, strVoiceContext)
    Else
        '播放语音
        Call mobjQueueManage.PlayQueueVoice(mobjMsgCenter, lngVoiceId, lngQueueId, False, strVoiceContext)
    End If

    '呼叫成功后删除呼叫过的内容
    Call mobjQueueManage.DelVoiceData(lngVoiceId)
End Sub

'开始数据轮询显示
Private Sub StartCall()
    tmrBroadCast.Enabled = True
End Sub

'终止数据轮询显示
Private Sub AbortCall()
    tmrBroadCast.Enabled = False
End Sub
