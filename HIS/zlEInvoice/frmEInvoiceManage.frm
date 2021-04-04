VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmEInvoiceManage 
   Caption         =   "电子票据管理"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   Icon            =   "frmEInvoiceManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picFunc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   8268
      Left            =   144
      ScaleHeight     =   8265
      ScaleWidth      =   3330
      TabIndex        =   1
      Top             =   600
      Width           =   3324
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6975
         Left            =   312
         ScaleHeight     =   6975
         ScaleWidth      =   2310
         TabIndex        =   2
         Top             =   264
         Width           =   2304
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   4320
            Left            =   48
            TabIndex        =   3
            Top             =   1272
            Width           =   2208
            _Version        =   589884
            _ExtentX        =   3895
            _ExtentY        =   7620
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   529
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   7608
         Left            =   24
         TabIndex        =   5
         Top             =   48
         Width           =   3000
         _Version        =   589884
         _ExtentX        =   5292
         _ExtentY        =   13414
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10584
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21590
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
   Begin XtremeCommandBars.ImageManager imgFunc 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmEInvoiceManage.frx":6852
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   864
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmEInvoiceManage.frx":1EED4
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmEInvoiceManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSys As Long, mlngModule As Long, mstrDBUser As String
Private mblnFirst As Boolean
Private mstrEInvPrivs As String  '电子票据操作模块权限

Private Enum Panel_Index
    Pane_Fun = 1001
    Pane_Form = 1002
End Enum

Private marrFunc(1) As String
Private Enum FunID_Idex
    FunID_基础数据设置 = 101
    FunID_收据费目对照 = 102
    FunID_收费渠道对照 = 103
    FunID_支付类别对照 = 104
    FunID_开票结算对照 = 105
    FunID_补开电子票据 = 201
    FunID_电子票据打印 = 202
    FunID_电子票据核对 = 203
End Enum

Private mWorkPan As Pane '当前功能
Private mfrmCurForm As Form '当前功能窗体
Private mfrmEInvoicePoint As frmEInvoicePoint
Private mfrmEInvoiceFees As frmEInvoiceFees
Private mfrmEInvoiceChannel As frmEInvoiceChannel
Private mfrmEInvoiceInsure As frmEInvoiceInsure
Private mfrmEInvoiceBalance As frmEInvoiceBalance
Private WithEvents mfrmEInvoiceCheck As frmEInvoiceCheck
Attribute mfrmEInvoiceCheck.VB_VarHelpID = -1
Private WithEvents mfrmEInvoiceCreate As frmEInvoiceCreate
Attribute mfrmEInvoiceCreate.VB_VarHelpID = -1
Private WithEvents mfrmEInvoicePrint As frmEInvoicePrint
Attribute mfrmEInvoicePrint.VB_VarHelpID = -1
Private mobjEInvoice As clsEInvoiceModule, mobjPubEInvoice As Object

Public Sub ShowMe(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, ByVal strDBUser As String, _
    objEInvoice As Object, Optional ByVal bytCheckTimeType As Byte)
    '程序入口
    '入参：
    '
    mlngSys = lngSys: mlngModule = lngModule
    mstrDBUser = strDBUser
    Set mobjEInvoice = objEInvoice
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    Call mobjEInvoice.zlGetEInvoiceProductName(Me, gstrProductName)
    
    On Error Resume Next
    Me.Show , frmMain
End Sub

Public Sub BHShowMe(ByVal lngMain As Long, ByVal lngSys As Long, ByVal lngModule As Long, ByVal strDBUser As String, _
    objEInvoice As Object)
    'BH调用程序入口
    mlngSys = lngSys: mlngModule = lngModule
    mstrDBUser = strDBUser
    Set mobjEInvoice = objEInvoice
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    Call mobjEInvoice.zlGetEInvoiceProductName(Me, gstrProductName)
    
    On Error Resume Next
    zlCommFun.ShowChildWindow Me.hWnd, lngMain
End Sub

Private Sub Form_Activate()
    If mblnFirst Then mblnFirst = False: Exit Sub
End Sub

Private Sub Form_Load()
    mblnFirst = True
    
    zlCommFun.ShowFlash "正在加载数据，请稍等...", Me
    mstrEInvPrivs = GetPrivFunc(mlngSys, 1145)     '电子票据操作权限
    
    Call DefMainCommandBars
    Call InitPanel '初始化dkpMain
    Call InitFunPanel
    
    Call RestoreWinState(Me, App.ProductName)
    
    zlCommFun.StopFlash
End Sub

Private Sub InitFunPanel()
    Dim strCategory As String
    Dim objPic As PictureBox
    
    strCategory = "业务数据管理,基础数据管理"
    
    '图标编号,TaskPanelItem的ID(同时也是参数容器Picture控件数组号),TaskPanelItem的标题;......
    marrFunc(0) = "114,201,补开电子票据;5012,202,电子票据打印;3010,203,电子票据核对"
    marrFunc(1) = "100,101,基础数据设置;105,102,收据费目对照;102,103,收费渠道对照;104,104,支付类别对照;111,105,开票结算对照"
    
    '1.初始化快捷面板的一级分类列表,缺省选中第一个
    Call InitSCBItem(scbFunc, strCategory, picTPL.hWnd)
    Call scbFunc.Icons.AddIcons(imgFunc.Icons)
      
    '2.初始化任务面板的二级分类列表,缺省选中第一个
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
End Sub

Public Sub InitSCBItem(ByRef scb As ShortcutBar, ByVal strItems As String, ByRef lngTPLhwnd As Long, Optional ByVal lngSelectedItem As Long = 1)
'功能：初始化一个快捷面板分类列表
'参数：
'      strItems         - 多个分类列表名称，以逗号分隔,例：基础数据初始,流程与规则,接口配置
'      lngTPLhwnd       - 分类列表上绑定的TaskPanel所在的容器句柄（窗体或Picture）
'      lngSelectedItem  - 缺省选中项的序号,从1开始
    Dim scbItem As ShortcutBarItem
    Dim i As Long
    Dim arrItem As Variant
    
    arrItem = Split(strItems, ",")
    For i = 0 To UBound(arrItem)
        Set scbItem = scb.AddItem(i + 1, arrItem(i), lngTPLhwnd)    '图标序号比指定的小1，所以要加1
        If i + 1 = lngSelectedItem Then Set scb.Selected = scbItem
    Next
    
    scb.ExpandedLinesCount = scb.ItemCount
End Sub

Public Sub InitTPLItem(ByRef scc As ShortcutCaption, ByRef tplFunc As TaskPanel, _
        ByVal strCategory As String, ByVal strItems As String, Optional ByVal lngSelectedItem As Long = 1)
'功能：初始或重新加载一个任务面板列表（仅一个分组）
'参数：
'      strCategory      - 显示在ShotcutCaption上的当前分类名称
'      strItems         - 多个二级分类的名称，以分号分隔,以逗号分隔图标ID、容器数组及二级分类名称,例：401,1,门诊划价管理;412,2,病人收费管理;......
'      lngSelectedItem  - 缺省选中项的序号,从1开始
    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    Dim arrItem As Variant
    Dim i As Long
    Dim lngImg As Long, lngID As Long
    Dim strItem As String
    Dim lngUbound As Long
    
    '增加一个隐藏分组
    scc.Caption = strCategory
    If tplFunc.Groups.Count = 0 Then
        Set tplGroup = tplFunc.Groups.Add(1, "分组")
        tplGroup.CaptionVisible = False
        tplGroup.Expanded = True
        
        tplFunc.SetMargins 1, 2, 0, 2, 2
        tplFunc.SetIconSize 24, 24
        tplFunc.SelectItemOnFocus = True
    Else
        Set tplGroup = tplFunc.Groups(1)    'index是从1开始的
        tplGroup.Items.Clear
    End If
    
    arrItem = Split(strItems, ";")
    lngUbound = UBound(arrItem)
    For i = 0 To lngUbound
        lngImg = Split(arrItem(i), ",")(0) + 1  '图标序号比指定的小1，所以要加1
        lngID = Split(arrItem(i), ",")(1)       'ID（作为参数控件容器的Picture数组编号）
        strItem = Split(arrItem(i), ",")(2)
        Set tplItem = tplGroup.Items.Add(lngID, strItem, xtpTaskItemTypeLink, lngImg)
        If i = lngUbound Then tplItem.SetMargins 0, 0, 0, 0 '不然最后一个选中时的框框不能完全框住内容
        If i + 1 = lngSelectedItem Then tplItem.Selected = True
    Next
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    If Val(tplFunc.Tag) = Item.ID Then Exit Sub
    tplFunc.Tag = Item.ID
    
    Select Case Item.ID
    Case FunID_基础数据设置
        If mfrmEInvoicePoint Is Nothing Then
            Set mfrmEInvoicePoint = New frmEInvoicePoint
            Call mfrmEInvoicePoint.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser)
        End If
        Set mfrmCurForm = mfrmEInvoicePoint
    Case FunID_收据费目对照
        If mfrmEInvoiceFees Is Nothing Then
            Set mfrmEInvoiceFees = New frmEInvoiceFees
            Call mfrmEInvoiceFees.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser)
        End If
        Set mfrmCurForm = mfrmEInvoiceFees
    Case FunID_收费渠道对照
        If mfrmEInvoiceChannel Is Nothing Then
            Set mfrmEInvoiceChannel = New frmEInvoiceChannel
            Call mfrmEInvoiceChannel.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser)
        End If
        Set mfrmCurForm = mfrmEInvoiceChannel
    Case FunID_支付类别对照
        If mfrmEInvoiceInsure Is Nothing Then
            Set mfrmEInvoiceInsure = New frmEInvoiceInsure
            Call mfrmEInvoiceInsure.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser)
        End If
        Set mfrmCurForm = mfrmEInvoiceInsure
    Case FunID_开票结算对照
        If mfrmEInvoiceBalance Is Nothing Then
            Set mfrmEInvoiceBalance = New frmEInvoiceBalance
            Call mfrmEInvoiceBalance.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser)
        End If
        Set mfrmCurForm = mfrmEInvoiceBalance
    Case FunID_电子票据核对
        If mfrmEInvoiceCheck Is Nothing Then
            Set mfrmEInvoiceCheck = New frmEInvoiceCheck
            Call mfrmEInvoiceCheck.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser, mobjEInvoice)
        End If
        Set mfrmCurForm = mfrmEInvoiceCheck
    Case FunID_补开电子票据
        If mfrmEInvoiceCreate Is Nothing Then
            Set mfrmEInvoiceCreate = New frmEInvoiceCreate
            Call mfrmEInvoiceCreate.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser, mstrEInvPrivs, mobjEInvoice, mobjPubEInvoice)
        End If
        Set mfrmCurForm = mfrmEInvoiceCreate
    Case FunID_电子票据打印
        If mfrmEInvoicePrint Is Nothing Then
            Set mfrmEInvoicePrint = New frmEInvoicePrint
            Call mfrmEInvoicePrint.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser, mstrEInvPrivs, mobjEInvoice, mobjPubEInvoice)
        End If
        Set mfrmCurForm = mfrmEInvoicePrint
    Case Else
        Exit Sub
    End Select
    
    mWorkPan.Handle = mfrmCurForm.hWnd
    '刷新子窗体菜单及工具条
    Call DefSubCommandBars(mWorkPan)
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID是从1开始的（因为同时为图标序号）,数组是从0开始
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub

Private Sub picFunc_Resize()
    On Error Resume Next
    scbFunc.Top = picFunc.ScaleTop
    scbFunc.Left = picFunc.ScaleLeft + 45
    scbFunc.Width = picFunc.ScaleWidth - 45
    scbFunc.Height = picFunc.ScaleHeight
End Sub

Private Sub picTPL_Resize()
    On Error Resume Next
    sccFunc.Left = picTPL.ScaleLeft
    sccFunc.Width = picTPL.ScaleWidth
    
    tplFunc.Left = picTPL.ScaleLeft
    tplFunc.Top = sccFunc.Top + sccFunc.Height
    tplFunc.Height = picTPL.ScaleHeight - sccFunc.Height
    tplFunc.Width = picTPL.ScaleWidth
End Sub

Private Sub InitPanel()
    Dim objPane As Pane

    On Error GoTo ErrHandler
    Set objPane = dkpMain.CreatePane(Pane_Fun, 120, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Tag = Pane_Fun
    objPane.MinTrackSize.Width = 60
    objPane.MaxTrackSize.Width = 240

    Set mWorkPan = dkpMain.CreatePane(Pane_Form, 700, 400, DockRightOf, objPane)
    mWorkPan.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    mWorkPan.Tag = Pane_Form

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
    '返回:设置成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar

    On Error GoTo ErrHandler

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
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False: cbrControl.BeginGroup = True
    End With

    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
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
        '.AddHiddenCommand conMenu_File_PrintSet
        '.AddHiddenCommand conMenu_File_Excel
    End With

    DefMainCommandBars = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub DefSubCommandBars(ByVal objItem As Pane)
    '功能：刷新子窗体菜单及工具条
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String

    On Error GoTo ErrHandler
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
    Call LockWindowUpdate(Me.hWnd)
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
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    stbThis.Top = Me.ScaleHeight - stbThis.Height
    stbThis.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    Set mWorkPan = Nothing
    Set mfrmCurForm = Nothing
    
    '卸载所有窗体
    If Not mfrmEInvoicePoint Is Nothing Then Unload mfrmEInvoicePoint: Set mfrmEInvoicePoint = Nothing
    If Not mfrmEInvoiceCheck Is Nothing Then Unload mfrmEInvoiceCheck: Set mfrmEInvoiceCheck = Nothing
    If Not mfrmEInvoiceCreate Is Nothing Then Unload mfrmEInvoiceCreate: Set mfrmEInvoiceCreate = Nothing
    If Not mfrmEInvoicePrint Is Nothing Then Unload mfrmEInvoicePrint: Set mfrmEInvoicePrint = Nothing
    If Not mfrmEInvoiceFees Is Nothing Then Unload mfrmEInvoiceFees: Set mfrmEInvoiceFees = Nothing
    If Not mfrmEInvoiceChannel Is Nothing Then Unload mfrmEInvoiceChannel: Set mfrmEInvoiceChannel = Nothing
    If Not mfrmEInvoiceInsure Is Nothing Then Unload mfrmEInvoiceInsure: Set mfrmEInvoiceInsure = Nothing
    If Not mfrmEInvoiceBalance Is Nothing Then Unload mfrmEInvoiceBalance: Set mfrmEInvoiceBalance = Nothing
    Set mobjEInvoice = Nothing
    
    If Not mobjPubEInvoice Is Nothing Then
        Call mobjPubEInvoice.zlTerminate
        Set mobjPubEInvoice = Nothing
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    On Error Resume Next
    Select Case Control.ID
    Case Else
        If Not mfrmCurForm Is Nothing Then
            Call mfrmCurForm.zlUpdateCommandBars(Control)
        End If
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    
    On Error GoTo ErrHandler
    Select Case Control.ID
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
    Case conMenu_Help_Help: Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((mlngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About: Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case Else
        If Not mfrmCurForm Is Nothing Then Call mfrmCurForm.zlExecuteCommandBars(Control)
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    Select Case Val(Item.Tag)
    Case Pane_Fun
        Item.Handle = picFunc.hWnd
    Case Pane_Form
        If Not mfrmCurForm Is Nothing Then
            Item.Handle = mfrmCurForm.hWnd
        End If
    End Select
End Sub

Private Sub mfrmEInvoiceCheck_ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
    Call ShowPopupMenu(blnAddOutPutExcel)
End Sub

Private Sub mfrmEInvoiceCreate_ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
    Call ShowPopupMenu(blnAddOutPutExcel)
End Sub

Private Sub mfrmEInvoicePrint_ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
    Call ShowPopupMenu(blnAddOutPutExcel)
End Sub

Public Sub ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
    '弹出右键菜单
    Dim objPopup As CommandBarPopup, cbCommandBar As CommandBar
    Dim cbrControl As CommandBarControl, cbrControlNew As CommandBarControl
    Dim i As Integer
    
    Set objPopup = cbsThis.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    Set cbCommandBar = cbsThis.Add("Popup", xtpBarPopup) '弹出菜单
    If cbCommandBar Is Nothing Then Exit Sub
    
    For i = 1 To objPopup.CommandBar.Controls.Count
        Set cbrControl = objPopup.CommandBar.Controls(i)
        Call cbsThis_Update(cbrControl)   '判断是否可见，因为第一次时菜单还没有执行Update
        If cbrControl.Visible Then
            Set cbrControlNew = cbCommandBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
            cbrControlNew.BeginGroup = cbrControl.BeginGroup
            cbrControlNew.IconId = cbrControl.IconId
            cbrControlNew.Enabled = cbrControl.Enabled
        End If
    Next
    
    If blnAddOutPutExcel Then
        Set objPopup = cbsThis.FindControl(xtpControlPopup, conMenu_FilePopup, , True)
        If Not objPopup Is Nothing Then
            Set cbrControl = objPopup.CommandBar.Controls.Find(xtpControlButton, conMenu_File_Excel, , True)
            If cbrControl.Visible Then
                Set cbrControlNew = cbCommandBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
                cbrControlNew.BeginGroup = True
                cbrControlNew.IconId = cbrControl.IconId
                cbrControlNew.Enabled = cbrControl.Enabled
            End If
        End If
    End If
    
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub
    
    cbCommandBar.ShowPopup
End Sub

Private Sub mfrmEInvoiceCheck_ShowInfo(ByVal strInfo As String)
        Call ShowInfoInStatusBar(strInfo)
End Sub

Private Sub mfrmEInvoiceCreate_ShowInfo(ByVal strInfo As String)
        Call ShowInfoInStatusBar(strInfo)
End Sub

Private Sub mfrmEInvoicePrint_ShowInfo(ByVal strInfo As String)
        Call ShowInfoInStatusBar(strInfo)
End Sub

Private Sub ShowInfoInStatusBar(ByVal strInfo As String)
    stbThis.Panels(2).Text = strInfo
End Sub

Public Function GetFirstCommandBar(ByRef objControls As CommandBarControls) As Long
'功能：获取工具栏打印预览按钮后的第一个按钮的index
    Dim objControl As CommandBarControl, idx As Long
    
    For Each objControl In objControls
        If objControl.ID = conMenu_File_Preview Then
            idx = objControl.index + 1
        End If
    Next
    GetFirstCommandBar = idx
End Function
