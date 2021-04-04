VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAppTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mblnRestart As Boolean '是否重启
Public Property Get IsRestart() As Boolean
    IsRestart = mblnRestart
End Property

Public Property Let IsRestart(ByVal blnRestart As Boolean)
    mblnRestart = blnRestart
End Property

Public Function GetLBoundIcon() As Long
'图标索引的下界
    GetLBoundIcon = glngLBound
End Function

Public Function GetUBoundIcon() As Long
'图标索引的上界
    GetUBoundIcon = glngUBound
End Function

Public Sub CodeMan(ByVal lngSysIN As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String, Optional ByVal strParameter As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '返回：
    '------------------------------------------------
    Set gcnOracle = cnMain
    glngSys = lngSysIN
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstr单位名称 = zlRegInfo("单位名称", , -1)
    gstrDbUser = strDbUser
    
    Set gfrmMain = frmMain
    
    '-------------------------------------------------
    '1   字典管理工具
    '2   邮件收发工具
    
    Set gclsAppTool = Me
    Select Case lngModul
    Case 1
        Call GetUserInfo
        gstrPrivs = GetPrivFunc(0, 11)
        frmDictManager.字典管理
    Case 2
        Call GetUserInfo
        frmMessageManager.Show , frmMain
    Case 3
        Call GetUserInfo
        gstrPrivs = GetPrivFunc(0, 13)
        gstrMenuSys = strParameter
        frmOptions.Show vbModal, frmMain
    Case 4
        NotifyMessage (strParameter = "Open")
    Case 5
        
        Call GetUserInfo
        ShowWindow frmAlert.hWnd, 0
        Call frmAlert.InitAlert
        
    Case 6
        
        frmAlert.Show , frmMain
        Call frmAlert.InitData
        
    End Select
End Sub

Private Sub NotifyMessage(blnOpen As Boolean)
'功能：自动消息通知程序的开关
    If gblnMessageGet = blnOpen Then Exit Sub
    
    gblnMessageGet = blnOpen
    If blnOpen = True Then
        Call GetUserInfo
        '打开关闭
        If gblnMessageShow = True Then
            '已经打开
            Exit Sub
        Else
            Load frmMessageRead
        End If
    Else
        If gblnMessageShow = True Then
            '消息收发程序还要用
            Exit Sub
        Else
            Unload frmMessageRead
        End If
    End If
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
        
    Call RemoveIcon(frmAlert.picNotify.hWnd)
    
    For Each frmThis In Forms
    
        If frmThis.Caption = "提醒消息" Then
            frmAlert.mblnUnLoad = True
        End If
        
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
    Set gclsAppTool = Nothing
End Function

Public Function GetIcon(ByVal IntIcon As Integer) As IPictureDisp
    Set GetIcon = LoadResPicture(IntIcon, vbResIcon)
End Function

Public Function GetResource(ByVal StrName As String) As IPictureDisp
    Set GetResource = LoadCustomPicture(StrName)
End Function

Public Function zlAppointDict(ByVal strTables As String, Optional ByVal lngSys As Long = 0)
    Call frmDictManager.字典管理(strTables, lngSys)
End Function

