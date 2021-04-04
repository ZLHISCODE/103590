VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmClinicPlanMainManage 
   Caption         =   "出诊安排管理"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "frmClinicPlanMainManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      Top             =   10260
      Visible         =   0   'False
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10590
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmClinicPlanMainManage.frx":1082
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18124
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   89
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   953
            MinWidth        =   882
            Text            =   "职称"
            TextSave        =   "职称"
            Key             =   "DoctorsTitle"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "安排颜色"
            TextSave        =   "安排颜色"
            Key             =   "PlanColor"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   720
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmClinicPlanMainManage.frx":1916
      Left            =   1260
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmClinicPlanMainManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModule As Long
Private mblnUnload As Boolean
Private mblnFirst As Boolean

Private mWorkPan As Pane '当前功能
Private mfrmCurForm As Form '当前功能窗体
Public mFunListActived As Boolean

Private mrs职称 As ADODB.Recordset  '所有医生专业技术职称和对应的标识符，缓存

Private mfrmClinicPlanMainFun As frmClinicPlanMainFun

Private mfrmClinicWorkTimeManage As frmClinicWorkTimeManage
Private mfrmClinicHolidayManage As frmClinicHolidayManage
Private mfrmClinicOfficeManage As frmClinicOfficeManage
Private mfrmClinicSignalSourceManage As frmClinicSignalSourceManage
    
Private mfrmClinicFixedPlanManage As frmClinicFixedPlanManage
Private mfrmClinicPlanDaysManage As frmClinicPlanDaysManage
Private mfrmClinicPlanTempletManage As frmClinicPlanTempletManage
Private mfrmClinicPlanStopVisitManage As frmClinicPlanStopVisitManage
Private mfrmClinicPlanTempletByDayManage As frmClinicPlanTempletByDayManage

Private Sub cbsThis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If Me.Visible = False Then Exit Sub

    Err = 0: On Error Resume Next
    Select Case CommandBar.Parent.id
    Case conMenu_View_FindType
        If Not mfrmCurForm Is Nothing Then Call mfrmCurForm.InitCommandsPopup(CommandBar)
    End Select
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
    mblnUnload = False
    If mblnFirst Then mblnFirst = False: Exit Sub
    
    Err = 0: On Error Resume Next
    '添加mFunListActived变量和ActiveFormChange事件是为了控制焦点
    If Not mfrmCurForm Is Nothing Then
        If mFunListActived = False And mfrmCurForm.Visible Then mfrmCurForm.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    '固定安排自动生成出诊记录
    On Error Resume Next
    mblnUnload = False
    zlCommFun.ShowFlash "正在加载数据，请稍等...", Me
    zlDatabase.ExecuteProcedure "zl1_auto_buildingregisterplan(Null)", Me.Caption
    
    Err = 0: On Error GoTo errHandler
    '如果登录站点无固定出诊表记录，则自动生成临床出诊表记录
    'Zl_临床出诊表_Add(
    strSQL = "Zl_临床出诊表_Add("
    '  操作类型_In         Number,
    strSQL = strSQL & "" & "2" & ","
    '  出诊id_In           临床出诊表.Id%Type,
    strSQL = strSQL & "" & "NULL" & ","
    '  出诊表名_In         临床出诊表.出诊表名%Type,
    strSQL = strSQL & "'" & "固定出诊表" & "',"
    '  站点_In             部门表.站点%Type,
    strSQL = strSQL & "'" & gstrNodeNo & "',"
    '  全院号源归属站点_In 部门表.站点%Type,
    strSQL = strSQL & "'" & gVisitPlan_ModulePara.str号源维护站点 & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    mblnFirst = True
    Set gobjRegist = New clsRegist
    gobjRegist.zlInitCommon glngSys, gcnOracle, gstrDBUser
    
    mstrPrivs = gstrPrivs
    mlngModule = glngModul
    
    Set mfrmClinicPlanMainFun = New frmClinicPlanMainFun
    Call mfrmClinicPlanMainFun.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)

    Call DefMainCommandBars
    Call InitPanel '初始化dkpMain
    Call RestoreWinState(Me, App.ProductName)
    Call Load职称  '加载医生职称及标识符
    
    zlCommFun.StopFlash
    Exit Sub
errHandler:
    zlCommFun.StopFlash
    mblnUnload = True
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnUnload = True
End Sub

Private Sub InitPanel()
    Dim objPane As Pane

    Err = 0: On Error GoTo errHandler
    Set objPane = dkpMain.CreatePane(Pane_FunFace, 150, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable
    objPane.MinTrackSize.Width = 130
    objPane.MaxTrackSize.Width = 240
    objPane.Tag = Pane_FunFace

    Set mWorkPan = dkpMain.CreatePane(Pane_Face, 700, 400, DockRightOf, objPane)
    mWorkPan.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    mWorkPan.Tag = Pane_Face

    With dkpMain
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar

    Err = 0: On Error GoTo errHandler

    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False

    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.id = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Sign, "职称标识设置(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ImportPlan, "导入“挂号安排”(&I)", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.id = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
        cbrSubControl.Checked = True
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
        cbrSubControl.Checked = True
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
        cbrSubControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.id = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    '显示自定义报表菜单
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs, _
        "ZL" & glngSys \ 100 & "_INSIDE_1114_1", "ZL" & glngSys \ 100 & "_INSIDE_1114_2", _
        "ZL" & glngSys \ 100 & "_INSIDE_1114_3", "ZL" & glngSys \ 100 & "_INSIDE_1114_4")

    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")

        'Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    '快键绑定
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With

    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
    End With

    DefMainCommandBars = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub DefSubCommandBars(ByVal ObjItem As Pane)
    '功能：刷新子窗体菜单及工具条
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String

    Err = 0: On Error GoTo errHandler
    '记录现有菜单样式
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsThis.Count >= 2 Then
        idx = GetFirstCommandBar(cbsThis(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsThis(2).Visible
            bytStyle = cbsThis(2).Controls(idx).Style
        End If
    End If

    '刷新子窗口菜单
    Call LockWindowUpdate(Me.Hwnd)
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsThis.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsThis.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsThis.Count To 2 Step -1
        cbsThis(lngCount).Delete
    Next

    '主窗口重新加入
    Call DefMainCommandBars

    '子窗口重新加入
    If Not mfrmCurForm Is Nothing Then
        Call mfrmCurForm.zlDefCommandBars
    End If

    '恢复及固定的一些菜单设置
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop + xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsThis.Count
        cbsThis(lngCount).ContextMenuPresent = False
        cbsThis(lngCount).ShowTextBelowIcons = False
        cbsThis(lngCount).EnableDocking xtpFlagStretched + xtpFlagHideWrap
        For Each objControl In cbsThis(lngCount).Controls
            If objControl.Type <> xtpControlLabel _
                And objControl.Type <> xtpControlEdit Then
                objControl.Style = bytStyle
            End If
        Next
        cbsThis(lngCount).Visible = blnShowBar
    Next

    '如果用了RecalcLayout反而不正常
    Call LockWindowUpdate(0)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    stbThis.Top = Me.ScaleHeight - stbThis.Height
    stbThis.Width = Me.ScaleWidth
End Sub

Private Sub SetProgressBarPostion(ByVal lngValue As Long, _
    Optional ByVal blnInit As Boolean, Optional ByVal lngMax As Long)
    
    With ProgressBar
        If blnInit Then
            .Max = lngMax
            .Left = stbThis.Panels(2).Left + 50
            .Top = stbThis.Top + (stbThis.Height - .Height) / 2 + 20
            .Width = stbThis.Panels(2).Width - 100
            .ZOrder
        Else
            .Value = lngValue
        End If
    End With
End Sub

Public Sub ActiveFormChange(objForm As Form)
    Err = 0: On Error Resume Next
    mFunListActived = objForm Is mfrmClinicPlanMainFun
End Sub

Public Sub NodeChanged(ByVal strKey As String)
    Call mfrmClinicPlanMainFun.RefreshVisitTable(strKey)
End Sub

Public Sub StatusShowInfoChanged(ByVal PanelIndex As Integer, ByVal strInfo As String)
    If PanelIndex = 2 Then
        stbThis.Panels(2).Text = strInfo
    Else
        stbThis.Panels(3).Text = strInfo
    End If
End Sub

Public Sub SelectedChange(ByVal bytMode As RegistPlanFun, _
    Optional ByVal lng出诊ID As Long, _
    Optional ByVal intYear As Integer, Optional ByVal intMonth As Integer, _
    Optional ByVal strTitle As String, Optional ByVal byt模板类型 As Byte)
    
    Err = 0: On Error Resume Next
    stbThis.Panels(2).Text = ""
    stbThis.Panels(3).Visible = False
    stbThis.Panels("PlanColor").Visible = False
    stbThis.Panels("DoctorsTitle").Visible = False
    
    Select Case bytMode
    Case Pane_WorkTime
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicWorkTimeManage Is Nothing Then
                Set mfrmClinicWorkTimeManage = New frmClinicWorkTimeManage
                Call mfrmClinicWorkTimeManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicWorkTimeManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            '刷新子窗体菜单及工具条
            Call DefSubCommandBars(mWorkPan)
        End If
        Call mfrmCurForm.LoadData
    Case Pane_Holiday
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicHolidayManage Is Nothing Then
                Set mfrmClinicHolidayManage = New frmClinicHolidayManage
                Call mfrmClinicHolidayManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicHolidayManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        Call mfrmCurForm.RefrashData(Year(zlDatabase.Currentdate))
    Case Pane_DoctorOffice
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicOfficeManage Is Nothing Then
                Set mfrmClinicOfficeManage = New frmClinicOfficeManage
                Call mfrmClinicOfficeManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicOfficeManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        Call mfrmCurForm.LoadData
    Case Pane_SignalSource
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicSignalSourceManage Is Nothing Then
                Set mfrmClinicSignalSourceManage = New frmClinicSignalSourceManage
                Call mfrmClinicSignalSourceManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicSignalSourceManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        Call mfrmCurForm.LoadData
        If mrs职称.RecordCount <> 0 Then stbThis.Panels("DoctorsTitle").Visible = True
    Case Pane_StopPlan '停诊管理
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicPlanStopVisitManage Is Nothing Then
                Set mfrmClinicPlanStopVisitManage = New frmClinicPlanStopVisitManage
                Call zlControl.FormSetCaption(mfrmClinicPlanStopVisitManage, False, False)
                Call mfrmClinicPlanStopVisitManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicPlanStopVisitManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        mfrmCurForm.RefreshData
    Case Pane_PlanTemplet
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicPlanTempletManage Is Nothing Then
                Set mfrmClinicPlanTempletManage = New frmClinicPlanTempletManage
                Call mfrmClinicPlanTempletManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicPlanTempletManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        mfrmCurForm.RefreshData IIf(byt模板类型 = 0, 2, 1), lng出诊ID, True
        stbThis.Panels(3).Visible = True
        If mrs职称.RecordCount <> 0 Then stbThis.Panels("DoctorsTitle").Visible = True
    Case Pane_MonthTemplet
        '2-按天排班的月排班模板
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicPlanTempletByDayManage Is Nothing Then
                Set mfrmClinicPlanTempletByDayManage = New frmClinicPlanTempletByDayManage
                Call mfrmClinicPlanTempletByDayManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicPlanTempletByDayManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        mfrmCurForm.RefreshData 1, lng出诊ID, True, intYear, intMonth, strTitle
        stbThis.Panels(3).Visible = True
        If mrs职称.RecordCount <> 0 Then stbThis.Panels("DoctorsTitle").Visible = True
    Case Pane_FixedPlan
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicFixedPlanManage Is Nothing Then
                Set mfrmClinicFixedPlanManage = New frmClinicFixedPlanManage
                Call mfrmClinicFixedPlanManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicFixedPlanManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        mfrmCurForm.RefreshData lng出诊ID, True
        stbThis.Panels("PlanColor").Visible = True
        stbThis.Panels(3).Visible = True
        If mrs职称.RecordCount <> 0 Then stbThis.Panels("DoctorsTitle").Visible = True
    Case Pane_MonthPlan, Pane_WeekPlan
        If mWorkPan.Tag <> bytMode Then
            If mfrmClinicPlanDaysManage Is Nothing Then
                Set mfrmClinicPlanDaysManage = New frmClinicPlanDaysManage
                Call mfrmClinicPlanDaysManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmClinicPlanDaysManage
            mWorkPan.Handle = mfrmCurForm.Hwnd
            mWorkPan.Tag = bytMode
            Call DefSubCommandBars(mWorkPan)
        End If
        mfrmCurForm.RefreshData IIf(bytMode = Pane_MonthPlan, 1, 2), lng出诊ID, True, intYear, intMonth, strTitle
        stbThis.Panels("PlanColor").Visible = True
        stbThis.Panels(3).Visible = True
        If mrs职称.RecordCount <> 0 Then stbThis.Panels("DoctorsTitle").Visible = True
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnUnload = False
    Call SaveWinState(Me, App.ProductName)
    Set mWorkPan = Nothing
    Set mfrmCurForm = Nothing
    '卸载所有窗体
    If Not mfrmClinicWorkTimeManage Is Nothing Then Unload mfrmClinicWorkTimeManage: Set mfrmClinicWorkTimeManage = Nothing
    If Not mfrmClinicHolidayManage Is Nothing Then Unload mfrmClinicHolidayManage: Set mfrmClinicHolidayManage = Nothing
    If Not mfrmClinicOfficeManage Is Nothing Then Unload mfrmClinicOfficeManage: Set mfrmClinicOfficeManage = Nothing
    If Not mfrmClinicSignalSourceManage Is Nothing Then Unload mfrmClinicSignalSourceManage: Set mfrmClinicSignalSourceManage = Nothing

    If Not mfrmClinicPlanDaysManage Is Nothing Then Unload mfrmClinicPlanDaysManage: Set mfrmClinicPlanDaysManage = Nothing
    If Not mfrmClinicFixedPlanManage Is Nothing Then Unload mfrmClinicFixedPlanManage: Set mfrmClinicFixedPlanManage = Nothing
    If Not mfrmClinicPlanTempletManage Is Nothing Then Unload mfrmClinicPlanTempletManage: Set mfrmClinicPlanTempletManage = Nothing
    If Not mfrmClinicPlanStopVisitManage Is Nothing Then Unload mfrmClinicPlanStopVisitManage: Set mfrmClinicPlanStopVisitManage = Nothing
    If Not mfrmClinicPlanTempletByDayManage Is Nothing Then Unload mfrmClinicPlanTempletByDayManage: Set mfrmClinicPlanTempletByDayManage = Nothing
    Unload mfrmClinicPlanMainFun: Set mfrmClinicPlanMainFun = Nothing
    
    On Error Resume Next
    Unload frmClinicPlanTemp '关闭临时窗口
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub

    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_File_ImportPlan
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排;所有科室", True)
        Control.Enabled = Control.Visible
    Case conMenu_File_Sign
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "职称标识设置")
        Control.Enabled = Control.Visible
    Case Else
        If Not mfrmCurForm Is Nothing Then
            Call mfrmCurForm.zlUpdateCommandBars(Control)
        End If
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl

    Err = 0: On Error GoTo errHandler
    Select Case Control.id
    Case conMenu_File_Sign
        If frmClinicDoctorTitleSet.ShowMe(Me) = False Then Exit Sub
        Call Load职称
        '刷新出诊表数据
        If mWorkPan.Tag = Pane_FixedPlan Or mWorkPan.Tag = Pane_MonthPlan _
            Or mWorkPan.Tag = Pane_WeekPlan _
            Or mWorkPan.Tag = Pane_PlanTemplet _
            Or mWorkPan.Tag = Pane_MonthTemplet Then
            Call mfrmClinicPlanMainFun.RefreshVisitTable
        ElseIf mWorkPan.Tag = Pane_SignalSource Then
            Call SelectedChange(Pane_SignalSource)
        End If
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_View_StatusBar
        Control.Checked = Not Control.Checked
        stbThis.Visible = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        Control.Checked = Not Control.Checked
        cbsThis(2).Visible = Control.Checked
        Set objControl = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Text, , True)
        objControl.Enabled = Control.Checked
        Set objControl = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Size, , True)
        objControl.Enabled = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not Control.Checked
        For Each objControl In cbsThis(2).Controls
            objControl.Style = IIf(Control.Checked, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Control.Checked = Not Control.Checked
        cbsThis.Options.LargeIcons = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_Help_Help: Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.Hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.Hwnd)
    Case conMenu_Help_About: Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zlCallCustomReprot(Me, Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))

        ElseIf Control.id = conMenu_File_Parameter Then
            '参数设置
            Dim frmPara As New frmClinicPlanParaSet
            If frmPara.ShowMe(Me, mlngModule, mstrPrivs) Then
                Call InitLocVisitPlanPar(mlngModule)
            End If
        ElseIf Control.id = conMenu_File_ImportPlan Then
            '导入安排
            If ImportPlan() Then
                MsgBox "“挂号安排”导入完成！", vbInformation, gstrSysName
                mfrmClinicPlanMainFun.RefreshVisitTable  '重新读取
            End If
        Else
            If Control.id = conMenu_View_Refresh And mFunListActived And Val(mWorkPan.Tag) > 5 Then
                '刷新出诊表列表
                mfrmClinicPlanMainFun.RefreshVisitTable
                Exit Sub
            End If

            If Val(mWorkPan.Tag) = Pane_PlanTemplet _
                Or Val(mWorkPan.Tag) = Pane_FixedPlan _
                Or Val(mWorkPan.Tag) = Pane_MonthPlan _
                Or Val(mWorkPan.Tag) = Pane_WeekPlan _
                Or Val(mWorkPan.Tag) = Pane_StopPlan _
                Or Val(mWorkPan.Tag) = Pane_MonthTemplet Then
                
                If ExecuteAddNewPlan(Control) Then Exit Sub
            End If
            If Not mfrmCurForm Is Nothing Then Call mfrmCurForm.zlExecuteCommandBars(Control)
        End If
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ExecuteAddNewPlan(ByVal Control As CommandBarControl) As Boolean
    '新增安排
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim strKey As String

    Err = 0: On Error GoTo errHandler
    Select Case Control.id
    Case conMenu_Edit_AddTemplet '模板
        If mfrmClinicPlanTempletManage Is Nothing Then
            Set mfrmClinicPlanTempletManage = New frmClinicPlanTempletManage
            Call mfrmClinicPlanTempletManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
        End If
        strKey = mfrmClinicPlanTempletManage.AddNewPlanTemplet
    Case conMenu_Edit_AddMonthPlan '月安排
        If mfrmClinicPlanDaysManage Is Nothing Then
            Set mfrmClinicPlanDaysManage = New frmClinicPlanDaysManage
            Call mfrmClinicPlanDaysManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
        End If
        strKey = mfrmClinicPlanDaysManage.AddNewPlan(True)
    Case conMenu_Edit_AddWeekPlan '周安排
        If mfrmClinicPlanDaysManage Is Nothing Then
            Set mfrmClinicPlanDaysManage = New frmClinicPlanDaysManage
            Call mfrmClinicPlanDaysManage.InitCommVariable(Me, cbsThis, mstrPrivs, mlngModule)
        End If
        strKey = mfrmClinicPlanDaysManage.AddNewPlan(False)
    Case Else
        Exit Function
    End Select
    If strKey = "" Then Exit Function
    Call mfrmClinicPlanMainFun.RefreshVisitTable(strKey)
    ExecuteAddNewPlan = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Err = 0: On Error GoTo errHandler
    If Item.Tag = Pane_FunFace Then
        Item.Handle = mfrmClinicPlanMainFun.Hwnd
    ElseIf Not mfrmCurForm Is Nothing Then
        Item.Handle = mfrmCurForm.Hwnd
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ImportPlan() As Boolean
    '导入历史安排
    Dim i As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim strTemp As String, cllSQL As Collection, blnDo As Boolean
    Dim rsPlanAll As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    '检查:
    '1.如果存在安排，则不允许再导入
    '2.如果号源中存在数据，不存在安排，则提醒“导入时这些号源将会被覆盖，是否继续导入？”
    '3.原挂号安排中的上班时间段如果有不存在的，则不允许导入，要求必须先添加（如安排中使用了"上午"，但是上班时间段里面没有"上午"）
    strSQL = "Select 号码 From 挂号安排 Order By ID Desc"
    Set rsPlanAll = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsPlanAll.EOF Then
        MsgBox "不存在挂号安排，不需要导入！", vbInformation, gstrSysName
        Exit Function
    End If
    
    strSQL = "Select 1 From 临床出诊表 A, 临床出诊安排 B Where a.Id = b.出诊id And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊号源 B, 部门表 C" & vbNewLine & _
                " Where a.号源id = b.Id And b.科室id = c.Id And (c.站点 Is Null Or c.站点 = [1]) And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)
        If rsTemp.EOF Then
            '不是本站点的
            MsgBox "当前其它院区已经存在临床出诊安排了，请先删除，否则不允许导入！", vbInformation, gstrSysName
        Else
            MsgBox "当前已经存在临床出诊安排了，请先删除，否则不允许导入！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    strSQL = "Select 1 From 临床出诊号源 Where Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        If MsgBox("当前存在临床出诊号源，在导入时这些号源将会被覆盖，是否继续导入？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    strSQL = "Select f_List2str(Cast(Collect(s.时间段) As t_Strlist)) As 时间段" & vbNewLine & _
            " From (Select 时间段, Row_Number() Over(Partition By 时间段 Order By 时间段) As 组号" & vbNewLine & _
            "        From (Select Decode(b.行号, 1, a.周一, 2, a.周二, 3, a.周三, 4, a.周四, 5, a.周五, 6, a.周六, a.周日) As 时间段" & vbNewLine & _
            "               From (Select 周一, 周二, 周三, 周四, 周五, 周六, 周日" & vbNewLine & _
            "                      From 挂号安排" & vbNewLine & _
            "                      Union All" & vbNewLine & _
            "                      Select 周一, 周二, 周三, 周四, 周五, 周六, 周日 From 挂号安排计划) A," & vbNewLine & _
            "                    (Select Level As 行号 From Dual Connect By Level <= 7) B)" & vbNewLine & _
            "        Where 时间段 Is Not Null) S, 时间段 T" & vbNewLine & _
            " Where s.时间段 = t.时间段(+) And t.时间段 Is Null And s.组号 = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!时间段) <> "" Then
            MsgBox "原挂号安排中的上班时间段【" & Nvl(rsTemp!时间段) & "】不存在，请先在“基础设置>上班时间管理”中添加！", vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '导入安排
    zlCommFun.ShowFlash "正在导入安排，请稍等...", Me
    Set cllSQL = New Collection
    Do While Not rsPlanAll.EOF
        cllSQL.Add "Zl_临床出诊表_导入('" & rsPlanAll!号码 & "'," & IIf(cllSQL.Count = 0, 1, 0) & ")"
        rsPlanAll.MoveNext
    Loop
    
    '执行SQL语句
    Call SetProgressBarPostion(0, True, cllSQL.Count)
    Me.ProgressBar.Visible = True
    blnDo = True: gcnOracle.BeginTrans
    For i = 1 To cllSQL.Count
        zlDatabase.ExecuteProcedure cllSQL(i), Me.Caption
        Call SetProgressBarPostion(i)
    Next
    gcnOracle.CommitTrans: blnDo = False
    Me.ProgressBar.Visible = False
    
    zlCommFun.StopFlash
    ImportPlan = True
    Exit Function
errHandler:
    If blnDo Then gcnOracle.RollbackTrans
    Me.ProgressBar.Visible = False
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub zlCallCustomReprot(ByVal frmMain As Form, ByVal lngSys As Long, strReprotName As String)
    '功能:调用相关的自定义报表

    Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain)
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If frmClinicPlanTemp.Visible Then Exit Sub
    Select Case Panel.Key
    Case "PlanColor"
        Call frmClinicPlanTemp.ShowPlanColor(Me)
    Case "DoctorsTitle"
        Call frmClinicPlanTemp.ShowDoctorsTitle(Me, mrs职称)
    End Select
End Sub

Public Function GetPopupCommandBarSub() As CommandBar
    '获取弹出菜单
    If mfrmCurForm Is Nothing Then Exit Function
    Set GetPopupCommandBarSub = GetPopupCommandBar(mfrmCurForm, cbsThis)
End Function

'功能:获取医生专业技术职务的名称及标识符
Public Sub Load职称()
    Dim strSQL As String
    strSQL = "Select 名称, 标识符 From 专业技术职务" & vbNewLine & _
             "Where 编码 like '23%'and 编码<>'23'" & vbNewLine & _
             "And 标识符 Is Not Null"
    Set mrs职称 = zlDatabase.OpenSQLRecord(strSQL, "获取职称和标识符")
    
    If mrs职称.RecordCount = 0 Then
        stbThis.Panels("DoctorsTitle").Visible = False
    Else
        If mWorkPan.Tag = Pane_FixedPlan Or mWorkPan.Tag = Pane_MonthPlan _
            Or mWorkPan.Tag = Pane_WeekPlan _
            Or mWorkPan.Tag = Pane_MonthTemplet _
            Or mWorkPan.Tag = Pane_PlanTemplet _
            Or mWorkPan.Tag = Pane_SignalSource Then
            stbThis.Panels("DoctorsTitle").Visible = True
        Else
            stbThis.Panels("DoctorsTitle").Visible = False
        End If
    End If
End Sub


