VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.MDIForm frmMDIMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "管理工具"
   ClientHeight    =   10140
   ClientLeft      =   165
   ClientTop       =   60
   ClientWidth     =   16005
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '屏幕中心
   Begin ComCtl3.CoolBar cbarTool 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   1535
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   16005
      _CBHeight       =   870
      _Version        =   "6.7.9816"
      MinHeight1      =   285
      Width1          =   5880
      NewRow1         =   0   'False
      MinHeight2      =   525
      Width2          =   2880
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Begin XtremeCommandBars.CommandBars cbsMain 
         Left            =   780
         Top             =   225
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin MSComDlg.CommonDialog DlgMain 
      Bindings        =   "frmMain.frx":1CFA
      Left            =   3885
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9765
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   661
      SimpleText      =   $"frmMain.frx":1D0E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":1D55
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23151
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   870
      Width           =   3135
      Begin XtremeSuiteControls.ShortcutBar sbFunc 
         Height          =   8640
         Left            =   120
         TabIndex        =   4
         Top             =   90
         Width           =   2550
         _Version        =   589884
         _ExtentX        =   4498
         _ExtentY        =   15240
         _StockProps     =   64
      End
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   4260
         Left            =   2625
         MousePointer    =   9  'Size W E
         ScaleHeight     =   4260
         ScaleWidth      =   45
         TabIndex        =   3
         Top             =   570
         Width           =   45
      End
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   7200
      Top             =   2280
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":25E7
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==变量定义
'==============================================================
Private mstrCurModule    As String          '当前选中模块
Public gstrLastModule   As String           '上次选中模块
Public grsToolsMenu     As ADODB.Recordset  '管理工具菜单
Private mcllModuleBar   As Collection
Private mcllItems       As Collection       '分类下子窗体
'==============================================================
'==公共接口
'==============================================================
Public Sub RunByModule(ByVal strNo As String)
'功能：转到执行模块菜单
    Dim frmChild As Form, strTmp As String
    mstrCurModule = strNo
    strTmp = Mid(strNo, 1, 2)
    sbFunc.Tag = "模块调用"
    If strTmp <> "" Then
        sbFunc.Selected = sbFunc.FindItem(Val(strTmp))
        If sbFunc.Tag <> "" Then '同一分类下不会触发事件，强制调用
            Call sbFunc_SelectedChanged(sbFunc.FindItem(Val(strTmp)))
        End If
    End If
    mstrCurModule = ""
End Sub

Public Function GetIcons() As ImageManager
    Set GetIcons = imgMain
End Function

'==============================================================
'=控件事件
'==============================================================
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Integer, strTmp As String
    
    On Error Resume Next
    Select Case Control.Id
        Case conMenu_File_PrintSet '打印设置
            Call zlPrintSet
        Case conMenu_File_Preview '预览
            gfrmActive.SubPrint 2
        Case conMenu_File_Print '打印
            gfrmActive.SubPrint 1
        Case conMenu_File_Excel '输出到Excel
            gfrmActive.SubPrint 3
        Case conMenu_View_ToolBar_Button        '工具栏
            For i = 2 To cbsMain.Count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
                
                cbarTool.Bands.Item(2).NewRow = Not cbarTool.Bands.Item(2).NewRow
                
                If cbarTool.Bands.Item(2).NewRow = True Then
                    If Me.cbsMain.Options.LargeIcons = True Then
                        cbarTool.Bands.Item(2).MinHeight = 520
                    Else
                        cbarTool.Bands.Item(2).MinHeight = 420
                    End If
                Else
                    cbarTool.Bands.Item(2).MinHeight = cbarTool.Bands(1).MinHeight
                End If
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text            '按钮文字
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size        '大图标
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            If Me.cbsMain.Options.LargeIcons = True Then
                cbarTool.Bands.Item(2).MinHeight = 520
            Else
                cbarTool.Bands.Item(2).MinHeight = 420
            End If
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StatusBar '状态栏
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolsList '工具列表
            Me.picFunc.Visible = Not Me.picFunc.Visible
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolsPwd
            Clipboard.Clear
            Clipboard.SetText gstrLoginUserPwd
        Case conMenu_Help_Help
            Select Case UCase(gfrmActive.name)
                Case UCase("frmAppMan")         '装卸管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmSvrCreate"
                Case UCase("frmAppStart")       '系统装卸管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmAppStart"
                Case UCase("frmAppUpgrade")     '系统升迁
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmAppUpgrade"
                Case UCase("frmAppCheck")       '对象检查修复
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmAppCheck"
                Case UCase("frmAppScript")      '置换安装脚本
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmAppScript"
                Case UCase("frmDataMan")        '数据管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmSvrCreate"
                Case UCase("frmDataMove")       '数据归档转移
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmDataMove"
                Case UCase("frmExp")            '数据导出
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmExp"
                Case UCase("frmImp")            '数据导入
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmImp"
                Case UCase("frmLoadOut")        '数据调出
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmLoadOut"
                Case UCase("frmLoadIn")         '数据调入
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmLoadIn"
                Case UCase("frmClearData")      '数据清除
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmClearData"
                Case UCase("frmRunMan")         '运行管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmSvrCreate"
                Case UCase("frmRegist")         '用户注册管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmRegist"
                Case UCase("frmStatus")         '运行状态监控
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmStatus"
                Case UCase("frmAutoJobs")       '后台作业管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmAutoJobs"
                Case UCase("FrmRunLog")         '运行日志管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "FrmRunLog"
                Case UCase("frmParameters")
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmParameters"
                Case UCase("FrmErrLog")         '错误日志管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "FrmErrLog"
                Case UCase("FrmRunOption")      '系统运行选项
                    ShowHelp Me.hwnd, "zl9svrtools\" & "FrmRunOption"
                Case UCase("frmGrantMan")       '权限管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmSvrCreate"
                Case UCase("frmRole")           '角色授权管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmRole"
                Case UCase("frmUser")           '用户授权管理
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmUser"
                Case UCase("frmMenu")           '菜单重组规划
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmMenu"
                Case UCase("frmMgrGrant") '管理工具授权
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmMgrGrant"
                Case UCase("frmRptMan") '报表管理
                    ShowHelp Me.hwnd, "zlreport\main"
            End Select
        Case conMenu_Help_Web_Home 'Web上的中联
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '发送反馈
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About '关于
            Call frmAbout.ShowAbout
        Case conMenu_File_RemoveTools '卸载管理工具
            Call FileRemove
        Case conMenu_File_LogOut '注销
            Unload Me
            Call Main
        Case conMenu_File_Exit '退出
            Unload Me
        Case Else '打开菜单中的模块
            strTmp = mcllModuleBar("K_" & Mid(Control.Id, 1, 3))
            If strTmp <> "" Then
                Call RunByModule(Mid(Control.Id, Len(strTmp) + 1))
            End If
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    If gfrmActive Is Nothing Then
        blnEnabled = False
    Else
        blnEnabled = gfrmActive.SupportPrint()
    End If
    
    Select Case Control.Id
        Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel '打印,预览,输出到Excel
            Control.Enabled = blnEnabled
        Case conMenu_File_RemoveTools '卸载管理工具的可用性
            Control.Enabled = gblnDBA
        Case conMenu_View_ToolBar_Button '工具栏
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
        Case conMenu_View_ToolsList
            Control.Checked = Me.picFunc.Visible
    End Select
End Sub

Private Sub MDIForm_Load()
    Dim rsTmp           As ADODB.Recordset, strSQL      As String
    Dim rsTmpChild      As ADODB.Recordset
    Dim objFrmTmp       As frmItem, objFrmMain          As frmItem
    Dim sbItem          As ShortcutBarItem, sbItemMain  As ShortcutBarItem
    Dim objPopup        As CommandBarPopup, objPrarentPop As CommandBarPopup, objControl As CommandBarControl
    Dim lngControlID    As Long
    Dim strSort         As String
    
    
    On Error GoTo errH
    Call zl9PrintMode.IniPrintMode(gcnOracle, gstrUserName)
    Call InitCommandBar
    '缓存菜单ID
    Set mcllModuleBar = New Collection
    mcllModuleBar.Add conMenu_Tool_LoadAndUnload, "K_501" '装卸管理
    mcllModuleBar.Add conMenu_Tool_DataMana, "K_502" '数据管理
    mcllModuleBar.Add conMenu_Tool_RunMana, "K_503" '运行管理
    mcllModuleBar.Add conMenu_Tool_Popedom, "K_504" '权限管理
    mcllModuleBar.Add conMenu_Tool_Expert, "K_505" '专项工具
    mcllModuleBar.Add conMenu_Tool_DBA, "K_506" 'DBA工具
    
    Set mcllItems = New Collection
    '变量初始化以及一些基本的界面设置
    Set gcbsMain = Me.cbsMain
    gblnSystemUser = gclsBase.IsStSystemUser(gstrLoginUserName)
    Me.Caption = Me.Caption & " [" & gstrLoginUserName & IIf(gstrServer = "", "", "@" & gstrServer) & "]"
    gstrSysName = gstrProductName & "软件"
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
    Call ApplyOEM(stbThis)
    Call ApplyOEM_Picture(Me, "Icon")
    '读取菜单加载菜单
    Call CheckProcManage    '临时检查\添加变动过程管理界面
    If CheckAndAdjustMustTable("Zlsvrtools", "次序", False) Then
        strSort = "次序,编号"
    Else
        strSort = "编号"
    End If
    strSQL = "Select * From Zlsvrtools" & IIf(gstrHaveProg <> "", " Where 上级 is null or instr('," & gstrHaveProg & ",' ,',' ||  编号  || ',' )>0", "")
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    Set rsTmpChild = CopyNewRec(rsTmp)
    Set grsToolsMenu = CopyNewRec(rsTmp)
    '添加右边快捷分组以及菜单下的模块
    Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    rsTmp.Filter = "上级=NULL": rsTmp.Sort = "编号"
    Do While Not rsTmp.EOF
        rsTmpChild.Filter = "上级 = " & rsTmp!编号
        rsTmpChild.Sort = strSort
        If rsTmpChild.RecordCount > 0 Then
            If rsTmpChild.RecordCount = 1 And rsTmpChild!编号 = "0404" And Not gblnSystemUser Then
                '当只有管理工具授权模块，且不是系统用户时，则不再加载分组以及子项
            Else
                If mstrCurModule = "" Then mstrCurModule = rsTmpChild!编号
                '获取菜单项目,并添加菜单
                If Not objControl Is Nothing Then
                    lngControlID = mcllModuleBar("K_5" & rsTmp!编号)
                    Set objPrarentPop = objControl.CommandBar.Controls.Add(xtpControlButtonPopup, lngControlID, rsTmp!标题 & IIf(rsTmp!快键 & "" = "", "", "(&" & rsTmp!快键 & ")"))
                    If Not objPrarentPop Is Nothing Then
                        Do While Not rsTmpChild.EOF
                            objPrarentPop.CommandBar.Controls.Add xtpControlButton, Val(lngControlID & rsTmpChild!编号), rsTmpChild!标题 & IIf(rsTmpChild!快键 & "" = "", "", "(&" & rsTmpChild!快键 & ")"), -1, False
                            rsTmpChild.MoveNext
                        Loop
                    End If
                End If
                '添加右边快捷导航
                Set objFrmTmp = New frmItem
                objFrmTmp.gstrParentNo = rsTmp!编号 & ""
                objFrmTmp.gstrParentCap = rsTmp!标题 & ""
                Set sbItem = sbFunc.addItem(Val(rsTmp!编号), rsTmp!标题, objFrmTmp.hwnd)
                mcllItems.Add objFrmTmp, "K_" & rsTmp!编号 & ""
                If objFrmMain Is Nothing Then Set objFrmMain = objFrmTmp
                If sbItemMain Is Nothing Then Set sbItemMain = sbItem
            End If
        End If
        rsTmp.MoveNext
    Loop
    '定位第一个模块
    Call sbFunc.Icons.AddIcons(imgMain.Icons)
    sbFunc.ExpandedLinesCount = sbFunc.ItemCount
    Call RunByModule(mstrCurModule)
    Exit Sub
errH:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub MDIForm_Resize()

    On Error Resume Next
    If picVbar.Left < 2200 Then picVbar.Left = 2200
    If picVbar.Left > Width - 3000 Then picVbar.Left = Width - 3000
    picVbar.Top = 0
    picVbar.Height = picFunc.Height
    picFunc.Width = picVbar.Left + picVbar.Width
    
    sbFunc.Left = picFunc.ScaleLeft + 45
    sbFunc.Width = picFunc.ScaleWidth - picVbar.Width - 45
    sbFunc.Top = picFunc.ScaleTop
    sbFunc.Height = picFunc.ScaleHeight
    
    If stbThis.Panels(2) = "" Then
        '特殊处理，不然状态栏的宽度不正确
        stbThis.Panels(2) = " "
        stbThis.Panels(2) = ""
    End If
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim frmChild As Form
    Set grsToolsMenu = Nothing
    Set mcllItems = Nothing
    Set mcllModuleBar = Nothing
    mstrCurModule = ""
    gstrLastModule = ""
    For Each frmChild In Forms
        Unload frmChild
    Next
End Sub

Private Sub picFunc_Resize()
    Call MDIForm_Resize
End Sub

Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        picVbar.Left = IIf(picVbar.Left + x < 2200, 2200, picVbar.Left + x)
        Call MDIForm_Resize
    End If
End Sub

Private Sub sbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub sbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If sbFunc.Tag <> "" Then
        If mstrCurModule <> gstrLastModule Then
            Call mcllItems("K_" & Format(Item.Id, "00")).RunByModule(mstrCurModule)
        End If
        sbFunc.Tag = ""
    Else
        Call mcllItems("K_" & Format(Item.Id, "00")).RunByModule
    End If
End Sub

'==============================================================
'=私有方法
'==============================================================
Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    
    '菜单定义:包括公共部份
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.Id = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "打印预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_RemoveTools, "卸载管理工具(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_LogOut, "注销(&L)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    objMenu.Id = conMenu_ToolPopup
'    With objMenu.CommandBar.Controls
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_LoadAndUnload, "装卸管理(&I)")
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_DataMana, "数据管理(&D)")
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_RunMana, "运行管理(&E)")
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Popedom, "权限管理(&G)")
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Expert, "专项工具(&R)")
'    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.Id = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_View_ToolsList, "工具列表(&L)")
        If gstrUserName <> gstrLoginUserName Then
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolsPwd, gstrLoginUserName & "的数据库密码(点击复制):" & gstrLoginUserPwd)
        End If
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.Id = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrWebSustainer)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrWebSustainer & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrWebSustainer & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyF12, conMenu_File_Parameter '参数设置
        
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新增
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改
        .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除
        
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '折叠所有组
        .Add FCONTROL, vbKeyG, conMenu_View_Filter '过滤
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找
        .Add 0, vbKeyF3, conMenu_View_FindNext '查找下一个
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
    '设置一些公共的不常用命令
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet '打印设置
        .AddHiddenCommand conMenu_File_Excel '输出到Excel
    End With
End Sub

Private Sub FileRemove()
'功能：卸载管理工具
    Dim rsTmp As ADODB.Recordset
    Dim blnReturn As Boolean
    
    '判断是否可以卸载
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    If rsTmp.RecordCount > 0 Then
        MsgBox "当前已经安装了应用系统，不能删除管理工具。", vbExclamation, gstrSysName
        Exit Sub
    End If
    If MsgBox("管理工具是管理应用系统的基础，" & vbCrLf & "将它删除后就不能再做任何工作，继续吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    blnReturn = RemoveServer
    Me.MousePointer = vbDefault
    Me.Enabled = True
    If blnReturn Then Unload Me
End Sub

Private Function RemoveServer() As Boolean
'功能：拆卸管理具体实现
'-------------拆卸算法-----------------
'   删除用户
'   删除回滚段
'   删除表空间
'--------------------------------------
    Dim strSpaces As String, strFiles As String, strErrInfo As String
    Dim aryTbs() As String, aryFile() As String
    Dim rsTemp As New ADODB.Recordset, intVer As Integer
    Dim strSQL As String, intCount As Integer
    
    On Error GoTo 0
    With rsTemp
        .Open "select 1 from gv$session where USERNAME='ZLTOOLS'", gcnOracle
        If .EOF = False Then
            MsgBox "ZLTOOLS用户正连接到数据库上，无法完成卸载操作。", vbExclamation, gstrSysName
            Exit Function
        End If
    End With
    
    '搜索表空间及数据文件
    strSpaces = "'ZLTOOLSTBS','ZLTOOLSTMP'"
    strFiles = ""
    With rsTemp
        strSQL = "select F.NAME from V$TABLESPACE T,V$DATAFILE F where T.TS#=F.TS# and T.NAME in (" & strSpaces & ")"
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        Do Until .EOF
            strFiles = strFiles & ";" & .Fields("NAME").value
            .MoveNext
        Loop
    End With
    If strFiles <> "" Then strFiles = Mid(strFiles, 2)

    On Error Resume Next
    
    '截断数据,才能执行删除所有者操作
    strSQL = "Truncate Table zltools.zlregaudit"
    gcnOracle.Execute strSQL
    If err.Number <> 0 Then Debug.Print err.Description
    
    strSQL = "Truncate Table zltools.zlRegFile"
    gcnOracle.Execute strSQL
    If err.Number <> 0 Then Debug.Print err.Description
   
    '删除本系统所有者
    
    stbThis.Panels(2).Text = "删除管理工具所有者…"
    DoEvents
    intCount = 0
    Do
        gcnOracle.Execute "Drop user ZLTOOLS cascade"
        If err.Number <> 0 Then Debug.Print err.Description
        With rsTemp
            If .State = adStateOpen Then .Close
            .Open "select * from all_users where username='ZLTOOLS'", gcnOracle
            If .EOF Then Exit Do
        End With
        intCount = intCount + 1
        DoEvents
        '最多删除100，如果都失败了就不再继续
        If intCount > 100 Then
            MsgBox "不能删除用户ZLTOOLS，可能其连接状态还未清除。", vbInformation, gstrSysName
            Exit Function
        End If
    Loop
    
    '删除已经建立的公共同义词
    stbThis.Panels(2).Text = "删除公共同义词…"
    DoEvents
    If rsTemp.State = adStateOpen Then rsTemp.Close
    strSQL = "SELECT Synonym_Name FROM All_Synonyms WHERE owner='PUBLIC' AND table_owner='ZLTOOLS'"
    rsTemp.Open strSQL, gcnOracle, adOpenStatic
    Do Until rsTemp.EOF
        strSQL = "drop public Synonym  " & rsTemp("Synonym_Name")
        gcnOracle.Execute strSQL
        rsTemp.MoveNext
    Loop
    
    '删除建立在表空间上的回滚段
    stbThis.Panels(2).Text = "删除表空间中的回滚段…"
    DoEvents
    With rsTemp
        If .State = adStateOpen Then .Close
        strSQL = "select SEGMENT_NAME from DBA_ROLLBACK_SEGS where tablespace_name in(" & strSpaces & ")"
        .Open strSQL, gcnOracle
        Do Until .EOF
            DoEvents
            gcnOracle.Execute "alter rollback segment " & .Fields(0).value & " offline"
            gcnOracle.Execute "drop rollback segment " & .Fields(0).value
            .MoveNext
        Loop
    End With
    
    '删除本系统数据空间
    stbThis.Panels(2).Text = "删除数据表空间…"
    DoEvents
    
    intVer = GetOracleVersion(, True)
    If intVer < 9 Then
        gcnOracle.Execute "alter rollback segment rbs_ZLTOOLS offline"
        gcnOracle.Execute "drop rollback segment rbs_ZLTOOLS"
    End If
    
    aryTbs = Split(strSpaces, ",")
    For intCount = LBound(aryTbs) To UBound(aryTbs)
        DoEvents
        strSpaces = Mid(aryTbs(intCount), 2, Len(aryTbs(intCount)) - 2)
        gcnOracle.Execute "alter tablespace " & strSpaces & " offline"
        gcnOracle.Execute "drop tablespace " & strSpaces & " including contents and datafiles cascade constraints"
    Next
    
    '试图删除无用的数据文件
    stbThis.Panels(2).Text = "删除无用的数据文件…"
    DoEvents
    aryFile = Split(strFiles, ";")
    For intCount = LBound(aryFile) To UBound(aryFile)
        err = 0
        Kill aryFile(intCount)
        If err <> 0 Then
            strErrInfo = strErrInfo & vbCr & "文件：" & aryFile(intCount)
        End If
    Next
    If strErrInfo <> "" Then
        MsgBox "管理工具拆卸完成，请手工删除以下内容：" & strErrInfo, vbExclamation, gstrSysName
    Else
        MsgBox "管理工具拆卸完成", vbExclamation, gstrSysName
    End If
    RemoveServer = True
End Function
