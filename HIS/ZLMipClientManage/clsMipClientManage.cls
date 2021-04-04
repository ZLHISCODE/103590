VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipClientManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Public Version As String

Private mfrmActiveForm As Object
Private WithEvents mfrmTable As frmTable
Attribute mfrmTable.VB_VarHelpID = -1
Private WithEvents mfrmItem As frmItem
Attribute mfrmItem.VB_VarHelpID = -1
Private WithEvents mfrmRunControl As frmRunControl
Attribute mfrmRunControl.VB_VarHelpID = -1
'Private WithEvents mfrmSendLog As frmSendLog
'Private WithEvents mfrmReceiveLog As frmReceiveLog
Private WithEvents mfrmMessageMonitor As frmMessageMonitor
Attribute mfrmMessageMonitor.VB_VarHelpID = -1
Private mobjCommunicateProxy As zlMipClient.CommunicateProxy
Private mstrIP As String
Private mintCount As Integer

Public Event AfterClose(ByVal lngModual As Long)
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
Public Event AfterCommunicateChange(ByVal strTitle As String, ByVal strNumber As String)

Public Function CommunicateProxyLogin(ByVal strIp As String, ByVal strPort As String, ByVal strUser As String, ByVal strPwd As String) As Boolean
    '连接到服务器
    '---------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    If mobjCommunicateProxy Is Nothing Then
        Set mobjCommunicateProxy = New zlMipClient.CommunicateProxy
        Call mobjCommunicateProxy.Init(strIp, Val(strPort))
    End If
    
    If mobjCommunicateProxy Is Nothing Then Exit Function
    
    Call mobjCommunicateProxy.Login(strUser, strPwd, "", , True)
    If mobjCommunicateProxy.IsConnected = False Then Exit Function
    
    mstrIP = strIp
    CommunicateProxyLogin = True
    Exit Function
errHand:
    MsgBox Err.Description, vbInformation + vbOKOnly, "信息提示"
End Function

Public Function CommunicateProxyInstall(ByVal strFilePath As String, Optional ByRef rsErr As ADODB.Recordset) As Boolean
    '导入数据
    
    On Error GoTo errHand
    
    Dim strParam() As String
    Dim bytAttachData() As Byte
    Dim bytEmpty() As Byte
    Dim strRemoteFile As String
    Dim strLock As String
    Dim strResult As String
    Dim objXML As New clsMipXML
    Dim rs As New ADODB.Recordset
    Dim strError As String
    Dim blnRet As Boolean
    ReDim strParam(1) As String
    Dim intCount As Integer
    Dim intLoop As Integer
    
    If OpenSQLiteFile(strFilePath, bytAttachData) Then
    
        '上传文件
        strRemoteFile = mobjCommunicateProxy.InvokeCustomService("UploadImportFile", strParam, bytAttachData)
        
        RaiseEvent AfterCommunicateChange("正在锁定...", 100)
        '锁定
        ReDim strParam(1)
        strParam(0) = mstrIP
        strParam(1) = "30"
        strLock = mobjCommunicateProxy.InvokeCustomService("BeginImport", strParam, bytEmpty)
        
        RaiseEvent AfterCommunicateChange("正在设置服务器导入目录下的文件名称...", 100)
        '设置服务器导入目录下的文件名称
        ReDim strParam(1)
        strParam(0) = strRemoteFile
        strParam(1) = strLock
        Call mobjCommunicateProxy.InvokeCustomService("SetImportFile", strParam, bytEmpty)
        
        RaiseEvent AfterCommunicateChange("正在获取导入文件流程信息...", 100)
        '获取导入文件中的流程信息
        ReDim strParam(0)
        strParam(0) = strLock
        strResult = mobjCommunicateProxy.InvokeCustomService("GetImportFlowInfo", strParam, bytEmpty)
        If objXML.OpenXMLDocument(strResult) Then
            If objXML.GetMultiNodeRecord("UpdateFlow/ID", rs) Then
                
            End If
            Call objXML.CloseXMLDocument
        End If
        
        RaiseEvent AfterCommunicateChange("正在停止流程...", 100)

        '停止流程
        ReDim strParam(0)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            On Error Resume Next
            For intLoop = 0 To rs.RecordCount - 1
                RaiseEvent AfterCommunicateChange("正在停止流程...", CStr(CInt(100 * (intLoop + 1) / (rs.RecordCount))))
            
                strParam(0) = rs("node_value").Value
                Call mobjCommunicateProxy.InvokeCustomService("StopFlows", strParam, bytEmpty)
                If Err.Number > 0 Then
                    If rsErr.State = adStateClosed Then
                        rsErr.Open
                    End If
                    mintCount = mintCount + 1
                    rsErr.AddNew
                    rsErr("序号").Value = mintCount
                    rsErr("内容").Value = Err.Description
                    Err.Clear
                End If
                DoEvents
                rs.MoveNext
            Next
            
            On Error GoTo errHand
        End If

        RaiseEvent AfterCommunicateChange("正在验证...", 100)
        '验证
        ReDim strParam(0)
        strParam(0) = strLock
        strError = mobjCommunicateProxy.InvokeCustomService("ValidateData", strParam, bytEmpty)
        
        RaiseEvent AfterCommunicateChange("正在导入...", 100)
        '导入
        ReDim strParam(0)
        strParam(0) = strLock
        Call mobjCommunicateProxy.InvokeCustomService("ImportData", strParam, bytEmpty)
        
        '重新获取导入文件中的流程信息
        ReDim strParam(0)
        strParam(0) = strLock
        strResult = mobjCommunicateProxy.InvokeCustomService("GetImportFlowInfo", strParam, bytEmpty)
        If objXML.OpenXMLDocument(strResult) Then
            If objXML.GetMultiNodeRecord("UpdateFlow/ID", rs) Then
                
            End If
            Call objXML.CloseXMLDocument
        End If
        
        '解锁
        ReDim strParam(0)
        strParam(0) = strLock
        blnRet = mobjCommunicateProxy.InvokeCustomService("EndImport", strParam, bytEmpty)
        
        '删除
        ReDim strParam(0)
        strParam(0) = strRemoteFile
        Call mobjCommunicateProxy.InvokeCustomService("DeleteImportFile", strParam, bytEmpty)
        
        RaiseEvent AfterCommunicateChange("正在重启流程...", 100)

        '启用流程
        ReDim strParam(0)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            On Error Resume Next
            For intLoop = 0 To rs.RecordCount - 1
                RaiseEvent AfterCommunicateChange("正在启用流程...", CStr(CInt(100 * (intLoop + 1) / (rs.RecordCount))))
            
                strParam(0) = rs("node_value").Value
                Call mobjCommunicateProxy.InvokeCustomService("StartFlows", strParam, bytEmpty)
                If Err.Number > 0 Then
                    If rsErr.State = adStateClosed Then
                        rsErr.Open
                    End If
                    mintCount = mintCount + 1
                    rsErr.AddNew
                    rsErr("序号").Value = mintCount
                    rsErr("内容").Value = Err.Description
                    Err.Clear
                End If
                DoEvents
                rs.MoveNext
            Next
            On Error GoTo errHand
        End If
    End If
    
    CommunicateProxyInstall = True
    
    Exit Function
errHand:
    MsgBox Err.Description, vbInformation + vbOKOnly, "信息提示"
End Function

Private Function OpenSQLiteFile(strFile As String, ByRef bytData() As Byte) As Boolean
    '-------------------------------------------------------------
    '功能：将指定的图形文件保存到指定的记录集字段中
    '参数：
    '       strFile：图形文件名称
    '       rsTable：图形存储记录集
    '       strField：图形字段
    '返回：
    '-------------------------------------------------------------
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, FileNum, j
    Dim aryChunk() As Byte
    
    On Error GoTo errH
    
    FileNum = FreeFile
    Open strFile For Binary Access Read As FileNum
    lngFileSize = LOF(FileNum)
    ReDim bytData(lngFileSize - 1) As Byte
    Get FileNum, , bytData()
    
    Close FileNum
    OpenSQLiteFile = True
    Exit Function

errH:
    Close FileNum
    OpenSQLiteFile = False

End Function

Public Function CommunicateProxyLogout()
    '断开服务器
    '---------------------------------------------------------------------------------------------------------------------------
    If Not (mobjCommunicateProxy Is Nothing) Then
        Call mobjCommunicateProxy.Logout
        Set mobjCommunicateProxy = Nothing
    End If
    
End Function

Public Function GetForm(ByVal lngModual As Long)
    
    ParamInfo.模块号 = lngModual
    
'    ParamInfo.系统号 = lngSys

    Select Case lngModual
    Case 1001  '业务信息定义
        If mfrmTable Is Nothing Then Set mfrmTable = New frmTable
        Set GetForm = mfrmTable
    Case 1002  '消息项目管理
        If mfrmItem Is Nothing Then Set mfrmItem = New frmItem
        Set GetForm = mfrmItem
    Case 1003  '消息站点控制
        If mfrmRunControl Is Nothing Then Set mfrmRunControl = New frmRunControl
        Set GetForm = mfrmRunControl
    Case 1004  '消息收发监控
        If mfrmMessageMonitor Is Nothing Then Set mfrmMessageMonitor = New frmMessageMonitor
        Set GetForm = mfrmMessageMonitor
    End Select
End Function

Public Function ActiveForm()
    On Error Resume Next
    DoEvents
    mfrmActiveForm.SetFocus
End Function

Public Sub ShowForm(ByVal lngModul As Long, objDataOracle As zlDataOracle.clsDataOracle, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
    '******************************************************************************************************************
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '   strDBUser:当前登录数据库用户名
    '******************************************************************************************************************
    
    Set gclsDataOracle = objDataOracle
    
    ParamInfo.系统名称 = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    ParamInfo.产品名称 = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")

    Set gclsBusiness = New clsBusiness
    Call gclsBusiness.InitBusiness(gclsDataOracle, ParamInfo.系统名称, "ZLTOOLS")
    
    UserInfo.用户名 = strDBUser
    
'    If Not GetUserInfo Then
'        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, ParamInfo.系统名称
'        Exit Sub
'    End If

    Set gfrmMain = frmMain

'    UserInfo.单位名称 = zlComLib.zlRegInfo("单位名称")
    
    zlComLib.SetDbUser UserInfo.用户名
    
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    Call InitSysPara
    
    Select Case lngModul
    '------------------------------------------------------------------------------------------------------------------
    Case 1001               '业务信息定义
        Call mfrmTable.ShowForm
        DoEvents
        mfrmTable.SetFocus
        Set mfrmActiveForm = mfrmTable
    '------------------------------------------------------------------------------------------------------------------
    Case 1002               '消息项目管理
        Call mfrmItem.ShowForm
        DoEvents
        mfrmItem.SetFocus
        Set mfrmActiveForm = mfrmItem
    '------------------------------------------------------------------------------------------------------------------
    Case 1003               '运行控制
        Call mfrmRunControl.ShowForm
        DoEvents
        mfrmRunControl.SetFocus
        Set mfrmActiveForm = mfrmRunControl
    '------------------------------------------------------------------------------------------------------------------
    Case 1004           '消息监控
        Call mfrmMessageMonitor.ShowForm
        DoEvents
        mfrmMessageMonitor.SetFocus
        Set mfrmActiveForm = mfrmMessageMonitor
    End Select
    
    Exit Sub
    
errHand:

    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
End Sub

Public Sub UnloadForm()
    
    If Not (mfrmTable Is Nothing) Then
        Unload mfrmTable
        Set mfrmTable = Nothing
    End If
    
    If Not (mfrmItem Is Nothing) Then
        Unload mfrmItem
        Set mfrmItem = Nothing
    End If
    
    If Not (mfrmRunControl Is Nothing) Then
        Unload mfrmRunControl
        Set mfrmRunControl = Nothing
    End If
        
    If Not (mfrmMessageMonitor Is Nothing) Then
        Unload mfrmMessageMonitor
        Set mfrmMessageMonitor = Nothing
    End If
End Sub

Public Function CloseWindows() As Boolean
    '******************************************************************************************************************
    '功能:关闭所有子窗口
    '******************************************************************************************************************
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
    
End Function

Private Sub Class_Initialize()
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
        
    Version = App.Major & "." & App.Minor & "." & App.Revision
    
    Set gfrmPubResource = New frmPubResource
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
'    If Not (gclsMsgSystem Is Nothing) Then Set gclsMsgSystem = Nothing
    If Not (gclsBusiness Is Nothing) Then Set gclsBusiness = Nothing
    If Not (gfrmPubResource Is Nothing) Then
        Unload gfrmPubResource
        Set gfrmPubResource = Nothing
    End If
    If Not (mobjCommunicateProxy Is Nothing) Then Set mobjCommunicateProxy = Nothing
End Sub

Private Sub mfrmMessageMonitor_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmMessageMonitor_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub

Private Sub mfrmTable_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmTable_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub

Private Sub mfrmItem_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmItem_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub

Private Sub mfrmRunControl_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmRunControl_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub


