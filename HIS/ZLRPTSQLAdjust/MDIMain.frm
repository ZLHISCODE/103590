VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "大表拆分配套工具"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBack 
      Align           =   3  'Align Left
      Height          =   6405
      Left            =   0
      ScaleHeight     =   6345
      ScaleWidth      =   1980
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2040
      Begin VB.PictureBox picSel 
         Height          =   5205
         Left            =   0
         ScaleHeight     =   5145
         ScaleWidth      =   10200
         TabIndex        =   2
         Top             =   0
         Width           =   10260
         Begin XtremeSuiteControls.TaskPanel tkpMain 
            Height          =   3735
            Left            =   210
            TabIndex        =   3
            Top             =   870
            Width           =   1710
            _Version        =   589884
            _ExtentX        =   3016
            _ExtentY        =   6588
            _StockProps     =   64
            ItemLayout      =   2
            HotTrackStyle   =   1
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6405
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   661
      SimpleText      =   $"MDIMain.frx":058A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "MDIMain.frx":05D1
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13018
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
            AutoSize        =   2
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
   Begin MSComctlLib.ImageList ils24 
      Left            =   6120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":0E63
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1577
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1C8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":239F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   6300
      Top             =   750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.ImageManager imgMgr 
      Left            =   4980
      Top             =   285
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "MDIMain.frx":2AB3
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   5556
      Top             =   276
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "MDIMain.frx":6D3B
      Left            =   4590
      Top             =   375
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmActive As Form
Private WithEvents Workspace As TabWorkspace
Attribute Workspace.VB_VarHelpID = -1
Private WithEvents mfrmExportReport As frmExportReport
Attribute mfrmExportReport.VB_VarHelpID = -1
Private WithEvents mfrmRptSQLMgr As frmRptSQLMgr
Attribute mfrmRptSQLMgr.VB_VarHelpID = -1
Private WithEvents mfrmCheckScrip As frmCheckScrip
Attribute mfrmCheckScrip.VB_VarHelpID = -1


Private Sub mfrmExportReport_StatusTextUpdate(ByVal strMSG As String)
    stbThis.Panels(2).Text = strMSG
End Sub
Private Sub mfrmRptSQLMgr_StatusTextUpdate(ByVal strMSG As String)
    stbThis.Panels(2).Text = strMSG
End Sub
Private Sub mfrmCheckScrip_StatusTextUpdate(ByVal strMSG As String)
    stbThis.Panels(2).Text = strMSG
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
        Dim objControl As CommandBarControl
        If Action = PaneActionClosed Then
            Set objControl = cbsMain.FindControl(, conMenu_View_Navigation)
            If Not objControl Is Nothing Then
                objControl.Checked = Not objControl.Checked
            End If
        End If
End Sub

Private Sub MDIForm_Load()
    Dim objNode As Node
    Me.Caption = Me.Caption & " [" & gstrUserName & IIf(gstrServer = "", "", "@" & gstrServer) & "]"
    gstrSysName = gstrProductName & "软件"
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
    RestoreWinState Me, App.ProductName
    Call InitControl
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   Dim frmChild As Form
    For Each frmChild In Forms
        If frmChild.Name <> Me.Name Then
            Unload frmChild
        End If
    Next
    SaveWinState Me, App.ProductName
    Set mfrmActive = Nothing
    Set mfrmExportReport = Nothing
    Set mfrmCheckScrip = Nothing
    Set mfrmRptSQLMgr = Nothing
    
End Sub
Private Function InitControl() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单、工具栏及tkpMain的相关对象
    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Err = 0: On Error GoTo ErrHand:
    With CommandBarsGlobalSettings
        .App = App
        .ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
        .ColorManager.SystemTheme = xtpSystemThemeAuto
    End With
    
    'cbrMain控件相关属性设置
    With cbsMain
        .VisualTheme = xtpThemeOffice2003
        With .Options
            .ShowExpandButtonAlways = False
            .UseDisabledIcons = True
            .AlwaysShowFullMenus = False
            .LargeIcons = True
            .SetIconSize True, 24, 24
            .SetIconSize False, 16, 16
        End With
        .EnableCustomization False
        Set .Icons = frmPubIco.imgPublic.Icons
        '菜单属性设置
        .ActiveMenuBar.EnableDocking (xtpFlagAlignTop + xtpFlagHideWrap)
        .TabWorkspace.PaintManager.Appearance = xtpTabAppearanceFlat
    End With
    
    '--------------------------------------------------------------------------------------------------------------------------
    '第一部分:加载菜单
    
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    With cbsMain.ActiveMenuBar.Controls
        '----------------------------------------------------------------------------------------------------------------------------
        '1.增加文件下菜单
        Set objMenu = .Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
        objMenu.Id = conMenu_FilePopup
        With objMenu.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)", -1, False)
'            Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "打印预览(&V)", -1, False)
'            Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)", -1, False)
'            Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_File_LogOut, "注销(&L)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)", -1, False)
        End With
        '----------------------------------------------------------------------------------------------------------------------------
        '2.增加查看等菜单
        Set objMenu = .Add(xtpControlPopup, conMenu_FilePopup, "查看(&V)", -1, False)
        objMenu.Id = conMenu_ViewPopup
        With objMenu.CommandBar.Controls
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
'            With objPopup.CommandBar.Controls
'                Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
'                objControl.Checked = True
'                Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
'                objControl.Checked = True
'                Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
'                objControl.Checked = True
'            End With
            Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
            Set objControl = .Add(xtpControlButton, conMenu_View_Navigation, "功能导航(&D)"): objControl.BeginGroup = True
            objControl.IconId = 7921
            
            Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
            
        End With
        '----------------------------------------------------------------------------------------------------------------------------
        '3.增加帮助主题等菜单
        Set objMenu = .Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
        objMenu.Id = conMenu_HelpPopup
        With objMenu.CommandBar.Controls
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrWebSustainer)
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, conMenu_Help_Web_Home, gstrWebSustainer & "主页(&H)", -1, False
                .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
            End With
        End With
    End With
    '--------------------------------------------------------------------------------------------------------------------------
    '第二部分:加载工具栏
'    Dim objBar As CommandBar
'
'    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
'    With objBar
'        '相关属性设置
'        .ContextMenuPresent = False
'        .ShowTextBelowIcons = False
'        .EnableDocking xtpFlagHideWrap
'        With .Controls
'            Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
'            Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
'
'            Set objControl = .Add(xtpControlButton, conMenu_View_Navigation, "隐藏导航"): objControl.BeginGroup = True
'            objControl.IconId = 7921
'
'            Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
'        End With
'    End With
        
    '设置各控件的样式:图标->标题
'    For Each objControl In objBar.Controls
'        objControl.Style = xtpButtonIconAndCaption
'    Next
    
    
    '--------------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print   '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
        .Add FCONTROL, vbKeyX, conMenu_File_Exit
    End With
    
    cbsMain.EnableCustomization (True)
    
    '--------------------------------------------------------------------------------------------------------------------------
    '第三部分:加载面版
    
    Dim objPane As Pane
    Set objPane = dkpMain.CreatePane(1, 100, 120, DockLeftOf, Nothing)
    
    objPane.Handle = picSel.hwnd
    objPane.Select
    objPane.Title = ""
    
    Set dkpMain.ImageList = ils24
    dkpMain.SetCommandBars Me.cbsMain
    
    Call LoadFunctionMenu
    Set Workspace = cbsMain.ShowTabWorkspace(True)
    Workspace.EnableGroups
    InitControl = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
 
Private Sub LoadFunctionMenu()
    '--------------------------------------------------------------------------------------
    '功能:加载功能菜单
    '--------------------------------------------------------------------------------------
    Dim objGroup As TaskPanelGroup
    Dim i As Long
    Dim lngMarg As Long
    
    Call tkpMain.Icons.AddIcons(imgMgr.Icons)
    
    With tkpMain
        .SetIconSize 32, 32
        .AllowDrag = True
        .VisualTheme = xtpTaskPanelThemeToolboxWhidbey
        .HotTrackStyle = xtpTaskPanelHighlightDefault
         
        .ItemLayout = xtpTaskItemLayoutImagesWithTextBelow
        .Behaviour = xtpTaskPanelBehaviourToolbox
        lngMarg = 1 * 15 / 10
        .SetItemOuterMargins lngMarg, lngMarg, lngMarg, lngMarg
        lngMarg = 7 * 2  '7*20/10
        .SetGroupInnerMargins lngMarg, lngMarg, lngMarg, lngMarg
        .SelectItemOnFocus = True
    End With
    
    Set objGroup = tkpMain.Groups.Add(0, "任务列表")
    With objGroup
        .Items.Add 1, "报表导出备份", xtpTaskItemTypeLink, 3
        .Items.Add 2, "报表SQL调整", xtpTaskItemTypeLink, 4
        .Items.Add 3, "过程和函数检查", xtpTaskItemTypeLink, 2
    End With
    objGroup.Expanded = True
    
End Sub
 Private Sub RunByModule(ByVal strModule As String)
    '----------------------------------------------------------------------------------------------------------------------------------------
    '功能:相关功能运行
    '参数:strModule-模块号
    '----------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim frmChild As Form
    
    For Each frmChild In Forms
        If frmChild Is Me Then
        ElseIf frmChild.MDIChild = True And frmChild.Enabled = True Then
            Unload frmChild
        End If
    Next
    Set mfrmActive = Nothing
    stbThis.Panels(2).Text = ""
    If strModule = "2" Then
        If CheckLogTab = False Then
            For i = 1 To tkpMain.Groups(1).Items.Count
                tkpMain.Groups(1).Items(i).Selected = False
            Next
            tkpMain.Groups(1).Items(1).Selected = True
            Call tkpMain_ItemClick(tkpMain.Groups(1).Items(1))
            Exit Sub
        End If
    End If
    
    Select Case strModule
        Case "1"
            Set mfrmActive = frmExportReport
            Set mfrmExportReport = mfrmActive
        Case "2" '自定义报表SQL调整
            Set mfrmActive = frmRptSQLMgr
            Set mfrmRptSQLMgr = mfrmActive
        Case "3"
            Set mfrmActive = frmCheckScrip
            Set mfrmCheckScrip = mfrmActive
    End Select
    If Not mfrmActive Is Nothing Then
        Call FindWindowAndSetActive(mfrmActive)
        mfrmActive.Show
        mfrmActive.ZOrder 0
    End If
End Sub

Private Sub tkpMain_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
        Me.dkpMain.Panes(1).Title = Item.Group.Caption
        RunByModule Item.Id
End Sub


Private Function CheckLogTab() As Boolean
    Dim strSQL As String, rstmp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select 1 From All_Tables Where Table_Name = Upper('zlrptadjustlog') And Owner = 'ZLTOOLS'"
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "查找日志表")
    If rstmp.RecordCount = 1 Then
        strSQL = "Select 1 From zltools.zlrptadjustlog Where rownum<2"
        Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "查找日志表")
    End If
    
    If rstmp.RecordCount = 0 Then
        MsgBox "请先执行报表导出备份，创建待调整的报表清单!", vbInformation, gstrSysName
        
        CheckLogTab = False
    Else
        CheckLogTab = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    '调整大小
   If stbThis.Visible Then
        Bottom = stbThis.Height
    End If
End Sub
Private Sub picSel_Resize()
    '调整大小
    Me.tkpMain.Width = picSel.ScaleWidth
    Me.tkpMain.Height = picSel.ScaleHeight
    Me.tkpMain.Left = picSel.ScaleLeft
    Me.tkpMain.Top = picSel.ScaleTop
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '功能执行
    
    Dim i As Integer, objControl As CommandBarControl
    Select Case Control.Id
    Case conMenu_File_Exit
        Unload Me
    Case conMenu_File_Preview
        mfrmActive.subPrint 2
    Case conMenu_File_Print
        mfrmActive.subPrint 1
    Case conMenu_File_Excel
        mfrmActive.subPrint 3
    Case conMenu_File_PrintSet
        Call zlPrintSet
    
'    Case conMenu_Help_Help
'        MsgBox "不存在帮助!", vbInformation, gstrSysName
    Case conMenu_File_LogOut
            Unload Me
            Call Main
    Case conMenu_Help_Web_Home
        ShellExecute hwnd, "open", "http://" & gstrWebURL, "", "", 1
    Case conMenu_Help_Web_Mail
        ShellExecute hwnd, "open", "mailto:" & gstrWebEmail, "", "", 1
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Not Control.Checked
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        
        Me.cbsMain.RecalcLayout
   Case conMenu_View_ToolBar_Button '工具栏
        Control.Checked = Not Control.Checked
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Control.Checked
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        Control.Checked = Not Control.Checked
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Navigation
        
        If Control.Checked = False Then
            dkpMain.ShowPane (1)
            Control.IconId = 7921
            Control.Caption = "隐藏导航(&D)"
            Control.Checked = True
        Else
            dkpMain.Panes(1).Close
            Control.Caption = "显示导航(&D)"
            Control.IconId = conMenu_View_Navigation
            Control.Checked = False
        End If
'
'        cbsMain.FindControl(, conMenu_View_Navigation).IconId = Control.IconId
'        cbsMain.FindControl(, conMenu_View_Navigation).Caption = Control.Caption
        
    Case conMenu_View_Refresh  '刷新
        Call mfrmActive.RefreshList
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '控件状态设置
    
    If mfrmActive Is Nothing Then
         Me.tkpMain.Enabled = True
    ElseIf mfrmActive.Enabled = False Then
        Me.tkpMain.Enabled = False
    Else
         Me.tkpMain.Enabled = True
    End If
    Select Case Control.Id
        Case conMenu_File_Exit
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
        Case conMenu_View_Navigation
            Control.Checked = dkpMain.Panes(1).Closed = False
    End Select
End Sub

Private Sub setPrintEnable(ByVal Control As CommandBarControl)
    '------------------------------------------------------------------------------
    '--功能:设置打印控件的Enable属性
    '--参数:Control-打印控件
    '------------------------------------------------------------------------------
    Dim blnEnable As Boolean
    
    If mfrmActive Is Nothing Then
        blnEnable = False
    Else
        blnEnable = mfrmActive.SupportPrint()
    End If
    Control.Enabled = blnEnable
End Sub

Private Sub FindWindowAndSetActive(ByVal FrmObj As Form)
    Dim LngTargetHdl As Long
    '--如果该窗体已经打开,则激活它(这样,窗体的大小不会发生变化)--zyb
    LngTargetHdl = FindWindow(vbNullString, FrmObj.Caption)
    If LngTargetHdl <> 0 Then
        If IsIconic(LngTargetHdl) Then
            Call ShowWindow(LngTargetHdl, 9)            '还原指定窗体为原大小
        End If
        Call SetActiveWindow(LngTargetHdl)
    End If
End Sub
