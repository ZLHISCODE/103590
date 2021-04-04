VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmNativateStart 
   BackColor       =   &H8000000C&
   Caption         =   "消息集成平台ZLHIS客户端管理"
   ClientHeight    =   10080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15630
   Icon            =   "frmNativateStart.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   10080
   ScaleWidth      =   15630
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   9720
      Width           =   15630
      _ExtentX        =   27570
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmNativateStart.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22701
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgTry 
      Height          =   240
      Left            =   2805
      Picture         =   "frmNativateStart.frx":70E6
      Top             =   2505
      Width           =   240
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   510
      Top             =   300
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   3
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmNativateStart.frx":D938
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmNativateStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mclsMipClientManage As clsMipClientManage
Attribute mclsMipClientManage.VB_VarHelpID = -1
Private mfrmNativate As New frmNativate
Private mstrKind As String
Private mrsMenus As ADODB.Recordset

Private maryModualPara As Variant
Private mcolOpenModual As New Collection
Private mlngSys As Long
Private mblnOpening As Boolean

Public Sub SetEnvironment(gstrSysNameIn As String, gstrVersionIn As String, gstrAviPathIn As String, _
                          gstrUserFlagIn As String, gstrDbUserIn As String, glngUserIdIn As Long, _
                          gstrUserCodeIn As String, gstrUserNameIn As String, gstrUserAbbrIn As String, _
                          glngDeptIdIn As Long, gstrDeptCodeIn As String, gstrDeptNameIn As String, _
                          gstrStationIn As String, gstrMenusysIn As String, Optional strCommand As String)
    '******************************************************************************************************************
    '功能：设置环境变量
    '参数：
    '返回：
    '******************************************************************************************************************

    gstrSysName = gstrSysNameIn
    gstrVersion = gstrVersionIn
    gstrAviPath = gstrAviPathIn
    gstrUserFlag = gstrUserFlagIn
    gstrDbUser = gstrDbUserIn
    glngUserId = glngUserIdIn
    gstrUserCode = gstrUserCodeIn
    gstrUserName = gstrUserNameIn
    gstrUserAbbr = gstrUserAbbrIn
    glngDeptId = glngDeptIdIn
    gstrDeptCode = gstrDeptCodeIn
    gstrDeptName = gstrDeptNameIn
    gstrStation = gstrStationIn
    gstrMenuSys = gstrMenusysIn
    gstrCommand = strCommand
End Sub

Public Sub InitBrower(StartForm As Object, cnOracle As ADODB.Connection, rsMenu As ADODB.Recordset)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error Resume Next
'    Set StartObj = StartForm
'    StartObj.Caption = "中联软件"
'    Set gcnOracle = cnOracle
'    Set mrsMenus = rsMenu.Clone
    
    Set mclsMipClientManage = New clsMipClientManage
    Me.Caption = "消息集成平台ZLHIS客户端管理(" & UCase(gstrDbUser) & ")"
    Me.Show
End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objcbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    Dim intPostion As Integer
    Dim strProductSimple As String
    
    
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    
    
    ParamInfo.产品名称 = zlRegInfo("产品简名")
        
    Call zlCommFun.CommandBarInit(cbsMain)
'    cbsMain.VisualTheme = xtpThemeWhidbey
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    cbsMain.Options.LargeIcons = True
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap

    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_LoadData, "安装消息数据(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_UnLoadData, "卸载消息数据(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Upgrade, "升级应用数据(&G)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_UnLoad, "卸载应用数据(&U)")
        
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '功能
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "功能(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, 1, "业务信息(&A)", , enumIcon.Data, "1001;业务信息定义")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, 1, "消息配置(&C)", , enumIcon.Document, "1002;消息项目配置")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, 1, "运行监控(&D)", , enumIcon.Workstation, "1003;站点运行监控")

    Set objControl = NewCommandBar(objMenu, xtpControlButton, 1, "消息监控(&S)", , enumIcon.HistoryMessage, "1004;消息收发监控")
     
    
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
'    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    
    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & ParamInfo.产品名称)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.产品名称 & "主页(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.产品名称 & "论坛(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = True
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, 1, "业务信息", True, , xtpButtonIconAndCaption)
    objControl.IconId = enumIcon.Data
    objControl.Parameter = "1001;业务信息定义"
    objControl.DescriptionText = "配置定义业务数据信息"
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, 1, "消息配置", False, , xtpButtonIconAndCaption)
    objControl.IconId = enumIcon.Document
    objControl.Parameter = "1002;消息项目配置"
    objControl.DescriptionText = "配置定义消息项目及内容和投递目标"
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, 1, "运行监控", False, , xtpButtonIconAndCaption)
    objControl.IconId = enumIcon.Workstation
    objControl.Parameter = "1003;站点运行监控"
    objControl.DescriptionText = "配置消息集成平台客户端运行启用的相关参数"
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, 1, "消息监控", False, , xtpButtonIconAndCaption)
    objControl.IconId = enumIcon.HistoryMessage
    objControl.Parameter = "1004;消息收发监控"
    objControl.DescriptionText = "查阅从客户端发送出去的消息和客户端接收到的消息"
        
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
        
    '命令的快键绑定:公共部份主界面已处理
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help_Help              '帮助
        .Add FCONTROL, vbKeyX, conMenu_File_Exit
    End With
    
    stbThis.Panels(2).Text = UCase(GetSetting("ZLSOFT", "注册信息\登陆信息", "USER")) & "@" & UCase(GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER"))
    mstrKind = ""
End Function

Private Function CommandBarExecutePublic(Control As Object, frmMain As Object, ByVal lngSys As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Help              '帮助主题
    
        Call zlComLib.ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((lngSys) / 100))
        
    Case conMenu_Help_Web_Home          'Web上的中联
        
        Call zlComLib.zlHomePage(frmMain.hWnd)
        
    Case conMenu_Help_Web_Forum         'Web上的论坛
    
        Call zlComLib.zlWebForum(frmMain.hWnd)
        
    Case conMenu_Help_Web_Mail          '发送反馈
        
        Call zlComLib.zlMailTo(frmMain.hWnd)
            
    Case conMenu_Help_About             '关于
        
        Call zlComLib.ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '按钮文字
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '大图标
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '状态栏
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_File_Exit             '退出
    
        Unload frmMain
        
    End Select
    
    CommandBarExecutePublic = True
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************

    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "标准"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_LoadData
        Call frmAppDataLoad.ShowDialog
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_UnLoadData
        Call frmAppDataUnload.ShowDialog
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Upgrade
        Call frmAppUpgrade.ShowDialog
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_UnLoad
        
        If MsgBox("您确定要删除消息集成平台客户端吗？删除后所有已有和消息有关的数据将丢失！", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If gclsMsgSystem.UnloadMipClient = True Then
            MsgBox "已经成功卸装了消息集成平台客户端，确定后将自动退出！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        If mblnOpening = False Then Call ExecuteModual(Control.Parameter, 1)
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        If mblnOpening = False Then Call ExecuteModual(Control.Parameter, 2)
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        Call CommandBarExecutePublic(Control, Me, mlngSys)
    End Select
End Sub

Private Sub ExecuteModual(ByVal strParameter As String, Optional ByVal bytMode As Byte = 1)
    Dim lngLoop As Long
    Dim lngHwnd As Long
    Dim blnExist As Boolean
    Dim aryTemp As Variant
    
    mblnOpening = True
'    stbThis.Panels(2).Text = ""
    If bytMode = 1 Then

        maryModualPara = Split(strParameter, ";")
                
        Me.Caption = "消息集成平台客户端管理(" & UCase(gstrDbUser) & ") - " & CStr(maryModualPara(1))

        blnExist = False
        For lngLoop = 1 To mcolOpenModual.Count
            aryTemp = Split(mcolOpenModual.Item(lngLoop), ";")
            If Val(aryTemp(0)) = Val(maryModualPara(0)) Then
                blnExist = True
                Exit For
            End If
        Next
        If blnExist = False Then
            mcolOpenModual.Add strParameter, "K" & Val(maryModualPara(0))
        End If
                
        dkpMain.Panes(1).Handle = mclsMipClientManage.GetForm(Val(maryModualPara(0))).hWnd
        Call mclsMipClientManage.ShowForm(Val(maryModualPara(0)), gclsMsgOracle, Me, gstrDbUser)

    Else
        aryTemp = Split(strParameter, ";")
        
        lngHwnd = mclsMipClientManage.GetForm(Val(aryTemp(0))).hWnd
        Call mclsMipClientManage.ShowForm(Val(aryTemp(0)), gclsMsgOracle, Me, gstrDbUser)
        Call mclsMipClientManage.ActiveForm
    End If
    
    mblnOpening = False
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 1
        If IsEmpty(maryModualPara) Or IsNull(maryModualPara) Then
            Control.Checked = False
        Else
            Control.Checked = (Join(maryModualPara, ";") = Control.Parameter)
        End If
    Case conMenu_View_ToolBar_Button            '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = stbThis.Visible
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = mfrmNativate.hWnd
    End Select
End Sub

Private Sub Form_Load()
    
    Call InitCommandBar
    Call InitDockPannel
    
    Call RestoreWinState(Me)
    Me.WindowState = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsMipClientManage Is Nothing) Then
        mclsMipClientManage.UnloadForm
        Set mclsMipClientManage = Nothing
    End If
    Unload mfrmNativate
    
    Dim frmThis As Form
    
    On Error Resume Next
    
    '关闭本部件窗体
    For Each frmThis In Forms
        If frmThis.Caption <> Me.Caption Then
            Unload frmThis
        End If
    Next
    
    Set gclsMsgSystem = Nothing
End Sub

Private Sub mclsMipClientManage_AfterClose(ByVal lngModual As Long)
    Dim lngLoop As Long
    Dim aryTemp As Variant
        
    For lngLoop = 1 To mcolOpenModual.Count
        aryTemp = Split(mcolOpenModual.Item(lngLoop), ";")
        If Val(aryTemp(0)) = lngModual Then
            mcolOpenModual.Remove lngLoop
            Exit For
        End If
    Next
    
    If mcolOpenModual.Count > 0 Then
        Call ExecuteModual(CStr(mcolOpenModual.Item(mcolOpenModual.Count)), 1)
    Else
        maryModualPara = Null
    End If
    
End Sub

Private Sub mclsMipClientManage_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
'    stbThis.Panels(intIndex).Text = strContent
End Sub
