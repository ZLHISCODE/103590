VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMspManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################

Public Version As String

Private mfrmActiveForm As Object
Private WithEvents mfrmTable As frmTable
Attribute mfrmTable.VB_VarHelpID = -1
Private WithEvents mfrmEvent As frmEvent
Attribute mfrmEvent.VB_VarHelpID = -1
Private WithEvents mfrmService As frmService
Attribute mfrmService.VB_VarHelpID = -1

Public Event AfterClose(ByVal lngModual As Long)
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)

Public Function GetForm(ByVal lngModual As Long)
    
'    UserInfo.模块权限 = zlComLib.GetPrivFunc(lngSys, lngModual)
    ParamInfo.模块号 = lngModual
    
'    ParamInfo.系统号 = lngSys

    Select Case lngModual
    Case 1000  '服务配置定义
        If mfrmService Is Nothing Then Set mfrmService = New frmService
        Set GetForm = mfrmService
        
    Case 1001  '业务信息定义
        If mfrmTable Is Nothing Then Set mfrmTable = New frmTable
        Set GetForm = mfrmTable
        
    Case 1002  '业务事件定义
        If mfrmEvent Is Nothing Then Set mfrmEvent = New frmEvent
        Set GetForm = mfrmEvent
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

    Set gclsMsgBase = New clsBusiness
    Call gclsMsgBase.InitBusiness(gclsDataOracle, ParamInfo.系统名称, strDBUser)
    
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
    Case 1000               '服务配置定义
        Call mfrmService.ShowForm
        DoEvents
        mfrmService.SetFocus
        Set mfrmActiveForm = mfrmService
    '------------------------------------------------------------------------------------------------------------------
    Case 1001               '业务信息定义
        Call mfrmTable.ShowForm
        DoEvents
        mfrmTable.SetFocus
        Set mfrmActiveForm = mfrmTable
    '------------------------------------------------------------------------------------------------------------------
    Case 1002               '业务事件定义
        Call mfrmEvent.ShowForm
        DoEvents
        mfrmEvent.SetFocus
        Set mfrmActiveForm = mfrmEvent
    End Select
    
    Exit Sub
    
errHand:

    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
End Sub

Public Sub UnloadForm()
    If Not (mfrmService Is Nothing) Then Unload mfrmService
    If Not (mfrmTable Is Nothing) Then Unload mfrmTable
    If Not (mfrmEvent Is Nothing) Then Unload mfrmEvent
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
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
        
    Version = App.Major & "." & App.Minor & "." & App.Revision
    
    Set gfrmPubResource = New frmPubResource
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
'    If Not (gclsMsgSystem Is Nothing) Then Set gclsMsgSystem = Nothing
    If Not (gclsMsgBase Is Nothing) Then Set gclsMsgBase = Nothing
    If Not (gfrmPubResource Is Nothing) Then
        Unload gfrmPubResource
        Set gfrmPubResource = Nothing
    End If
End Sub

Private Sub mfrmService_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmService_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub

Private Sub mfrmTable_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmTable_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub

Private Sub mfrmEvent_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmEvent_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub


