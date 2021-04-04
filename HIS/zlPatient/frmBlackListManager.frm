VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJOCK.DOCKINGPANE.UNICODE.9600.OCX"
Begin VB.Form frmBlackListManager 
   Caption         =   "病人不良记录管理"
   ClientHeight    =   11070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15240
   Icon            =   "frmBlackListManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11070
   ScaleWidth      =   15240
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10710
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBlackListManager.frx":06EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21802
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmBlackListManager.frx":0F7E
      Left            =   1260
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBlackListManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModule As Long
Private mblnUnload As Boolean
Private mblnFirst As Boolean
Private mobjWorkPan As Pane '当前功能
Private mfrmCurForm As Form '当前功能窗体

Public mblnFunListActived As Boolean

Private WithEvents mfrmBlackListMainFun As frmBlackListMainFun    '主要功能菜单
Attribute mfrmBlackListMainFun.VB_VarHelpID = -1
Private WithEvents mfrmBlackTypeManage As frmBlackTypeManage   '不良行为分类管理
Private WithEvents mfrmBlackListReasonManage As frmBlackListReasonManage   '不良行为常用的原因管理
Private WithEvents mfrmBlackListRecordManage As frmBlackListRecordManage   '不良行为记录管理
Attribute mfrmBlackListRecordManage.VB_VarHelpID = -1


Private Sub cbsThis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If Me.Visible = False Then Exit Sub
    Err = 0: On Error Resume Next
    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType
        If Not mfrmCurForm Is Nothing Then Call mfrmCurForm.InitCommandsPopup(CommandBar)
    End Select
End Sub

Private Sub Form_Activate()

    If mblnUnload Then Unload Me: Exit Sub
    mblnUnload = False
    If mblnFirst Then mblnFirst = False: Exit Sub
    
    Err = 0: On Error Resume Next
    '添加mblnFunListActived变量和ActiveFormChange事件是为了控制焦点
    If Not mfrmCurForm Is Nothing Then
        If mblnFunListActived = False And mfrmCurForm.Visible Then mfrmCurForm.SetFocus
    End If
End Sub
Private Sub InitVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关模块级变量
    '编制:刘兴洪
    '日期:2018-11-08 10:36:07
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    Set mfrmBlackListMainFun = New frmBlackListMainFun
    Call mfrmBlackListMainFun.zlInitComm(Me, cbsThis, mstrPrivs, mlngModule)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
     
    Err = 0: On Error GoTo errHandle
    mblnFirst = True: mstrPrivs = gstrPrivs: mlngModule = glngModul
    Call InitVar
    Call DefMainCommandBars
    Call InitPanel '初始化dkpMain
    Call RestoreWinState(Me, App.ProductName)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
     mblnUnload = True
End Sub

Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面各相关区域面版
    '编制:刘兴洪
    '日期:2018-11-08 10:38:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane

    Err = 0: On Error GoTo ErrHandler
    
    Set objPane = dkpMain.CreatePane(gEM_BlackListFun.Em_Pane_FunFace, 150, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable
    objPane.MinTrackSize.Width = 130
    objPane.MaxTrackSize.Width = 240
    objPane.Tag = Em_Pane_FunFace

    Set mobjWorkPan = dkpMain.CreatePane(gEM_BlackListFun.Em_Pane_Face, 700, 400, DockRightOf, objPane)
    mobjWorkPan.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    mobjWorkPan.Tag = Em_Pane_Face

    With dkpMain
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 10:41:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar

    Err = 0: On Error GoTo ErrHandler

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
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        'Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
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
    cbrMenuBar.ID = conMenu_HelpPopup
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
    Set cbrToolBar = GetCommbarFromName(cbsThis, "工具栏")
    If cbrToolBar Is Nothing Then
        Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    End If
    
   ' Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
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
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub DefSubCommandBars(ByVal objItem As Pane)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定义子窗体菜单及工具条
    '入参:ObjItem-当前功能页对象
    '编制:刘兴洪
    '日期:2018-11-08 10:42:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objControl As CommandBarControl, bytStyle As XTPButtonStyle, blnShowBar As Boolean
    Dim lngCount As Long, lngIndex As Long, objCustom As CommandBarControlCustom
    Dim strName As String, cbrToolBar As CommandBar

    Err = 0: On Error GoTo ErrHandler
    '记录现有菜单样式
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsThis.Count >= 2 Then
        lngIndex = zlGetFirstCommandBar(cbsThis(2).Controls)
        If lngIndex > 0 Then
            blnShowBar = cbsThis(2).Visible
            bytStyle = cbsThis(2).Controls(lngIndex).Style
        End If
    End If

    '刷新子窗口菜单
    Call LockWindowUpdate(Me.hWnd)
    
    '删除现在的工具栏及顶级菜单项
    cbsThis.ActiveMenuBar.Controls.DeleteAll
    If Not mfrmCurForm Is Nothing Then mfrmCurForm.zlCancelBands
    
    Set cbrToolBar = GetCommbarFromName(cbsThis, "工具栏")
    cbrToolBar.Controls.DeleteAll

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
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    stbThis.Top = Me.ScaleHeight - stbThis.Height
    stbThis.Width = Me.ScaleWidth
End Sub

Private Sub mfrmBlackListMainFun_zlActivate(ByVal frmSubForm As Form)
    '子窗体集合时触发该事件
    mblnFunListActived = frmSubForm Is mfrmBlackListMainFun
End Sub
Private Sub mfrmBlackListReasonManage_zlActivate(ByVal frmSubForm As Form)
  mblnFunListActived = frmSubForm Is mfrmBlackListMainFun
End Sub

Private Sub mfrmBlackTypeManage_zlActivate(ByVal frmSubForm As Form)
    mblnFunListActived = frmSubForm Is mfrmBlackListMainFun
End Sub

Private Sub mfrmBlackListMainFun_SelectedChange(ByVal bytFunMode As gEM_BlackListFun, ByVal strBlackLitType As String)
    '功能选择改变后触发该事件
    
    On Error GoTo errHandle
    stbThis.Panels(2).Text = ""
 
    
    Select Case bytFunMode
    Case Em_Pane_Type   '不良行类分类
        If mobjWorkPan.Tag <> bytFunMode Then
        
            If Val(mobjWorkPan.Tag) = Em_Pane_Record And Not mfrmBlackListRecordManage Is Nothing Then
                Call mfrmBlackListRecordManage.zlCancelBands
            End If
            
            If mfrmBlackTypeManage Is Nothing Then
                Set mfrmBlackTypeManage = New frmBlackTypeManage
                Call mfrmBlackTypeManage.zlInitComm(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmBlackTypeManage
            mobjWorkPan.Handle = mfrmCurForm.hWnd
            mobjWorkPan.Tag = bytFunMode
            '刷新子窗体菜单及工具条
            Call DefSubCommandBars(mobjWorkPan)
        End If
        Call mfrmCurForm.zlLoadData
    Case Em_Pane_Reason  '不良原因
        If mobjWorkPan.Tag <> bytFunMode Then
            If Val(mobjWorkPan.Tag) = Em_Pane_Record And Not mfrmBlackListRecordManage Is Nothing Then
                Call mfrmBlackListRecordManage.zlCancelBands
            End If
            If mfrmBlackListReasonManage Is Nothing Then
                Set mfrmBlackListReasonManage = New frmBlackListReasonManage
                Call mfrmBlackListReasonManage.zlInitComm(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmBlackListReasonManage
            mobjWorkPan.Handle = mfrmCurForm.hWnd
            mobjWorkPan.Tag = bytFunMode
            Call DefSubCommandBars(mobjWorkPan)
        End If
        Call mfrmCurForm.zlLoadData
        
    Case Em_Pane_Record '不良记录管理
        
        If mobjWorkPan.Tag <> bytFunMode Then
            If mfrmBlackListRecordManage Is Nothing Then
                Set mfrmBlackListRecordManage = New frmBlackListRecordManage
                Call mfrmBlackListRecordManage.zlInitComm(Me, cbsThis, mstrPrivs, mlngModule)
            End If
            Set mfrmCurForm = mfrmBlackListRecordManage
            mobjWorkPan.Handle = mfrmCurForm.hWnd
            mobjWorkPan.Tag = bytFunMode
            Call DefSubCommandBars(mobjWorkPan)
        End If
        Call mfrmCurForm.zlLoadData(strBlackLitType)
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Sub mfrmBlackListRecordManage_zlShowStatusText(ByVal bytPancel As Byte, ByVal strText As String)
     stbThis.Panels(bytPancel).Text = strText
End Sub
 


Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    mblnUnload = False
    Call SaveWinState(Me, App.ProductName)
    Set mobjWorkPan = Nothing
    Set mfrmCurForm = Nothing
    '卸载所有窗体
    If Not mfrmBlackTypeManage Is Nothing Then Unload mfrmBlackTypeManage: Set mfrmBlackTypeManage = Nothing
    If Not mfrmBlackListReasonManage Is Nothing Then Unload mfrmBlackListReasonManage: Set mfrmBlackListReasonManage = Nothing
    If Not mfrmBlackListRecordManage Is Nothing Then Unload mfrmBlackListRecordManage: Set mfrmBlackListRecordManage = Nothing
    If Not mfrmBlackListMainFun Is Nothing Then Unload mfrmBlackListMainFun: Set mfrmBlackListMainFun = Nothing
End Sub


Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If mfrmCurForm Is Nothing Then Exit Sub
    Call mfrmCurForm.zlUpdateCommandBars(Control)
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl

    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Parameter '参数设置
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
    Case conMenu_View_Refresh   '刷新
            
            
        If mfrmBlackListMainFun Is Nothing Then Exit Sub
        
        If Not (mblnFunListActived And Val(mobjWorkPan.Tag) = 13) Then
            If mfrmCurForm Is Nothing Then Exit Sub
            Call mfrmCurForm.zlExecuteCommandBars(Control)
            Exit Sub
        End If
        Call mfrmBlackListMainFun.zlRefresh
    
    Case conMenu_Help_Help: Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About: Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zlOpenCustomReport(Me, Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
        Else
            If mfrmCurForm Is Nothing Then Exit Sub
            Call mfrmCurForm.zlExecuteCommandBars(Control)
        End If
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Err = 0: On Error GoTo ErrHandler
    If Item.Tag = Em_Pane_FunFace Then
        Item.Handle = mfrmBlackListMainFun.hWnd
    ElseIf Not mfrmCurForm Is Nothing Then
        Item.Handle = mfrmCurForm.hWnd
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub zlOpenCustomReport(ByVal frmMain As Form, ByVal lngSys As Long, strReprotName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开相关的自定义报表
    '入参:frmMain-调用的父窗体
    '     lngSys-系统号
    '     strReprotName-报名名称
    '编制:刘兴洪
    '日期:2018-11-08 11:16:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain)
End Sub

Public Function GetPopupCommandBarSub() As CommandBar
    '获取弹出菜单
    If mfrmCurForm Is Nothing Then Exit Function
    Set GetPopupCommandBarSub = zlGetPopupCommandBar(mfrmCurForm, cbsThis)
End Function
Private Sub mfrmBlackTypeManage_zlChangeType()
    If mfrmBlackListMainFun Is Nothing Then Exit Sub
    Call mfrmBlackListMainFun.zlRefresh(True)
End Sub
