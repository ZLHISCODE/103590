VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmFinanceSupervise 
   Caption         =   "收费财务监控"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11730
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinanceSupervise.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11730
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8055
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   635
      SimpleText      =   $"frmFinanceSupervise.frx":6852
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmFinanceSupervise.frx":6899
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13044
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "刘兴洪"
            TextSave        =   "刘兴洪"
            Object.ToolTipText     =   "当前操作员:刘兴洪"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   1605
      Left            =   -75
      TabIndex        =   1
      Top             =   1020
      Width           =   4290
      _Version        =   589884
      _ExtentX        =   7567
      _ExtentY        =   2831
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   -30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmFinanceSupervise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mfrmCollect As frmFinanceSupervisePersonList
Private mfrmHistory As frmFinaceSuperviseHistory
Private mfrmStandbyMoney As frmFinanceSuperviseStandbyMoenyList
Private mblnAllowZero As Boolean  '允许归零操作

Private Enum mPgIndex
    EM_PG_收款列表 = 250101
    EM_PG_历史列表 = 250102
    EM_PG_备用金列表 = 250103
End Enum

Private Sub initVar()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量
    '编制:刘兴洪
    '日期:2013-10-14 16:35:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dtDate As Date
    dtDate = Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-mm-dd")
    strSQL = "Select 1 From 人员收缴记录 Where 登记时间>=[1] And Rownum=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtDate)
    '轧帐归零操作:收费员缴款数据清零，主要是因为由些用户使用的是报表打印的方式进行缴款,而做的清零操作
    '   如果在一月内存在人员收款、轧帐记录，则不允许进行轧帐归零操作，直接屏蔽该功能
    mblnAllowZero = rsTemp.EOF  '
    If Not rsTemp Is Nothing Then rsTemp.Close
    Set rsTemp = Nothing
End Sub


Public Sub zlShowFinanceSupervise(ByVal frmMain As Object, _
        ByVal lngModule As Long, ByVal strPrivs As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费财务监控程序入口
    '入参:frmMain-调用的主窗体
    '       lngModule-模块号
    '       strPrivs-模块权限串
    '编制:刘兴洪
    '日期:2013-09-22 16:32:18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
     
    If CheckDepend = False Then Exit Sub
    '初始化数据
    Call InitFace
    If frmMain Is Nothing Then
        Me.Show
    Else
        Me.Show , frmMain
    End If
End Sub

Public Sub BHShowList(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngMain As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '编制:刘兴洪
    '日期:2013-10-17 18:17:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If CheckDepend = False Then Exit Sub
    '初始化数据
    Call InitFace
    mlngModule = lngModule: mstrPrivs = strPrivs
    zlCommFun.ShowChildWindow Me.hWnd, lngMain
    Me.ZOrder 0
End Sub

Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面
    '编制:刘兴洪
    '日期:2013-09-03 14:43:09
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call InitPage
End Sub
Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-28 18:21:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim objComBar As CommandBarComboBox
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
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
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置"): mcbrControl.BeginGroup = True
        If mblnAllowZero And zlStr.IsHavePrivs(mstrPrivs, "轧帐归零") Then '允许轧帐归零操作时，才允许执行该功能
            Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain_Zero, "轧帐归零(&C)"): mcbrControl.BeginGroup = True
        End If
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_PutOut, "发放备用金(&L)")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_OnWork, "上岗备用金")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_PutIn, "收回备用金(&H)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3017
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Manual, "手工收款(&M)")
        mcbrControl.IconId = 6820
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_RollingCurtain, "轧帐收款(&S)")
        mcbrControl.IconId = 3588
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Cancel, "收款作废(&C)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3589
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Other, "其他人员收款(&O)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 228
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CheckCash, "现金点钞(&E)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3590
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Personnel_Group, "成员分组(&G)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeBook_Reprint, "重打收款收据(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DrawBook_Reprint, "重打备用金领用单(&D)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "大图标(&G)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "小图标(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "列表(&L)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "详细资料(&D)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "查看明细数据(&V)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 2322
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("C"), conMenu_Edit_Collect_Cancel
        .Add FCONTROL, Asc("R"), conMenu_Edit_ChargeBook_Reprint
        .Add FCONTROL, Asc("M"), conMenu_Edit_Collect_Manual
        .Add FCONTROL, Asc("O"), conMenu_Edit_Collect_Other
        .Add FCONTROL, Asc("T"), conMenu_View_Detail
        .Add 0, VK_F2, conMenu_Edit_Collect_RollingCurtain
        .Add 0, VK_F6, conMenu_Edit_CheckCash
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Manual, "手工收款"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 6820
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_RollingCurtain, "轧帐收款")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Other, "其他人员收款(&O)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 228
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Cancel, "收款作废")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_PutOut, "发放备用金")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_OnWork, "上岗备用金")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_PutIn, "收回备用金"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3017
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CheckCash, "现金点钞"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3590
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "查询明细")
         mcbrControl.IconId = 2322
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")

    End With
    For Each mcbrControl In mcbrToolBar.Controls
        If mcbrControl.ID <> conMenu_COMBOX_INTERFACE Then
          mcbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    
    If zlStr.IsHavePrivs(mstrPrivs, "收费员收款") _
        Or zlStr.IsHavePrivs(mstrPrivs, "财务组收款") _
        Or zlStr.IsHavePrivs(mstrPrivs, "其他人员收款") Then
        If mfrmCollect Is Nothing Then
            Set mfrmCollect = New frmFinanceSupervisePersonList
            Load mfrmCollect
        End If
        '初始化变量
        Call mfrmCollect.zlInitVar(mlngModule, mstrPrivs, cbsThis)
    End If
    
    If mfrmHistory Is Nothing Then
        Set mfrmHistory = New frmFinaceSuperviseHistory
        Load mfrmHistory
    End If
    Call mfrmHistory.zlInitVar(mlngModule, mstrPrivs)
 
    If mfrmStandbyMoney Is Nothing Then
        Set mfrmStandbyMoney = New frmFinanceSuperviseStandbyMoenyList
        Load mfrmStandbyMoney
    End If
    Call mfrmStandbyMoney.zlInitVar(mlngModule, mstrPrivs)
    If zlStr.IsHavePrivs(mstrPrivs, "收费员收款") _
      Or zlStr.IsHavePrivs(mstrPrivs, "财务组收款") _
      Or zlStr.IsHavePrivs(mstrPrivs, "其他人员收款") Then
        Set objItem = tbPage.InsertItem(EM_PG_收款列表, "收款", mfrmCollect.hWnd, 0)
        objItem.Tag = EM_PG_收款列表
    End If
    Set objItem = tbPage.InsertItem(EM_PG_历史列表, "历史收款信息", mfrmHistory.hWnd, 0)
    objItem.Tag = EM_PG_历史列表
    Set objItem = tbPage.InsertItem(EM_PG_备用金列表, "备用金列表 ", mfrmStandbyMoney.hWnd, 0)
    objItem.Tag = EM_PG_备用金列表
     With tbPage
        Set tbPage.PaintManager.Font = Me.Font
        tbPage.Item(0).Selected = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.StaticFrame = True
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutSizeToFit
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

 
Private Sub cbsThis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Err = 0: On Error Resume Next
    With tbPage
        tbPage.Left = Left
        tbPage.Top = Top
        tbPage.Width = Right - Left
        tbPage.Height = Bottom - Top
    End With
End Sub

Private Sub Form_Activate()
    stbThis.Panels(3).Text = UserInfo.姓名
End Sub

Private Sub Form_Load()
    Call initVar
    RestoreWinState Me, App.ProductName
    Call zlDefCommandBars
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    If Not mfrmCollect Is Nothing Then Unload mfrmCollect
    If Not mfrmHistory Is Nothing Then Unload mfrmHistory
    If Not mfrmStandbyMoney Is Nothing Then Unload mfrmStandbyMoney
    Set mfrmCollect = Nothing
    Set mfrmHistory = Nothing
    Set mfrmStandbyMoney = Nothing
End Sub
 
Private Sub ParameterSet()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:参数设置
    '编制:刘兴洪
    '日期:2013-09-12 15:31:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If frmFinanceSuperviseParaSet.ShowMe(Me, mlngModule, mstrPrivs) = False Then Exit Sub
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Sub SaveCollect(Optional ByVal blnCustomCollect As Boolean = False)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:轧帐处理
    '入参:blnCustomCollect-true-手工收款;false-轧帐收款;
    '编制:刘兴洪
    '日期:2013-09-12 15:34:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String
    If Val(tbPage.Selected.Tag) <> EM_PG_收款列表 Then Exit Sub
    
    If Not (zlStr.IsHavePrivs(mstrPrivs, "收费员收款") _
        Or zlStr.IsHavePrivs(mstrPrivs, "财务组收款") _
        Or zlStr.IsHavePrivs(mstrPrivs, "其他人员收款")) Then Exit Sub
    
    If Not mfrmCollect.zlRollingCurtainCollect(Me, blnCustomCollect) Then Exit Sub
    Call zlRefresh
End Sub
Private Sub SaveCollectCancel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收款作废处理
    '编制:刘兴洪
    '日期:2013-09-12 15:34:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String, blnDel As Boolean
    If Val(tbPage.Selected.Tag) <> EM_PG_历史列表 Then Exit Sub
    If mfrmHistory.CancelData() Then Exit Sub
End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
     Select Case Control.ID
        Case conMenu_File_Exit: Unload Me: '退出(&X)
        Case conMenu_File_PrintSet: Call zlPrintSet '打印设置
        Case conMenu_File_Preview: Call zlPrintRpt(2)  '预览(&V)
        Case conMenu_File_Print: Call zlPrintRpt(1) '打印(&P)
        Case conMenu_File_Excel: Call zlPrintRpt(3)  '输出到&Excel…
        Case conMenu_File_Parameter: Call ParameterSet '参数设置
        Case conMenu_Edit_RollingCurtain_Zero: ExcuteRollingCurtainZero  '轧帐归零(&C)"
        Case conMenu_Edit_StandbyMoeny_PutOut: ExcutePutOutStandbyMoeny '发放备用金
        Case conMenu_Edit_StandbyMoeny_OnWork: ExcuteOnWorkStandbyMoeny
        Case conMenu_Edit_StandbyMoeny_PutIn: ExcutePutINStandbyMoeny '收回备用金
        Case conMenu_Edit_Collect_Manual: Call SaveCollect(True)   '手工收款
        Case conMenu_Edit_Collect_RollingCurtain: Call SaveCollect '轧帐收款
        Case conMenu_Edit_Collect_Cancel: Call SaveCollectCancel '收款作废
        Case conMenu_Edit_Collect_Other: Call SaveCollect  '其他人员收款
        Case conMenu_Edit_CheckCash: Call CheckCash '现金点钞(&E)
        Case conMenu_Edit_Personnel_Group: Call ExcuteSplitGroup '成员分组
        Case conMenu_Edit_ChargeBook_Reprint:  Call RePrintBill(0) '重打收款收据(&R)
        Case conMenu_Edit_DrawBook_Reprint:  Call RePrintBill(1) '重打备用金领用单(&R)
        Case conMenu_View_Detail: Call ShowChargeList '查看明细数据(&V)
        Case conMenu_View_Refresh: zlRefresh '刷新(&R)
        Case conMenu_View_LargeICO: SetPersonListShow (0) '大图标(&G)
        Case conMenu_View_MinICO: SetPersonListShow (1)  '小图标(&M)
        Case conMenu_View_ListICO: SetPersonListShow (2)  '列表(&L)
        Case conMenu_View_DetailsICO:: SetPersonListShow (3)  '详细资料(&D)
        Case conMenu_View_StatusBar '状态栏(&S)
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else
            If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                '执行发布到当前模块的报表
                Call CallCustomRpt(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
            End If
        End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHavePrivs As Boolean, lngPage As Long
    Dim intView As Integer, blnEanbled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    If Not tbPage.Selected Is Nothing Then
        lngPage = Val(tbPage.Selected.Tag)
    End If
    Select Case Control.ID
    Case conMenu_Edit_RollingCurtain_Zero '轧帐归零
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "轧帐归零") And mblnAllowZero
        blnEanbled = lngPage = EM_PG_收款列表
        Control.Visible = blnHavePrivs And blnEanbled
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_StandbyMoeny_PutOut '发放备用金
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "发放备用金")
        blnEanbled = lngPage = EM_PG_备用金列表
        Control.Visible = blnHavePrivs And blnEanbled:
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_StandbyMoeny_OnWork '上岗备用金
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "上岗备用金")
        blnEanbled = lngPage = EM_PG_备用金列表
        Control.Visible = blnHavePrivs And blnEanbled:
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_StandbyMoeny_PutIn '收回备用金
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "收回备用金")
        blnEanbled = lngPage = EM_PG_备用金列表
        Control.Visible = blnHavePrivs And blnEanbled
        If blnEanbled Then blnEanbled = mfrmStandbyMoney.IsAllowCancel
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_Collect_Manual '手工收款
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "手工收款")
        blnEanbled = lngPage = EM_PG_收款列表
        If blnEanbled Then
              blnEanbled = blnEanbled And mfrmCollect.IsAllowCustomCollect
        End If
       Control.Visible = blnHavePrivs And blnEanbled
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_Collect_RollingCurtain '轧帐收款
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "轧帐收款")
        blnEanbled = lngPage = EM_PG_收款列表
        Control.Visible = blnHavePrivs And blnEanbled
        If blnEanbled Then blnEanbled = mfrmCollect.IsAllowCollect
        Control.Enabled = blnEanbled
    Case conMenu_Edit_Collect_Other  '其他人员收款
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "其他人员收款")
        blnEanbled = lngPage = EM_PG_收款列表
        If blnEanbled Then blnEanbled = mfrmCollect.IsAllowOtherCollect
        Control.Visible = blnHavePrivs And blnEanbled
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_Collect_Cancel '收款作废
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "收款作废")
        blnEanbled = lngPage = EM_PG_历史列表
        Control.Visible = blnHavePrivs And blnEanbled
        If blnEanbled Then blnEanbled = mfrmHistory.IsAllowCollectCancel
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_CheckCash ' "现金点钞(&E)")
        blnEanbled = lngPage <> EM_PG_备用金列表
        Control.Visible = blnEanbled
        Control.Enabled = blnEanbled
    Case conMenu_Edit_Personnel_Group '成员分组
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "成员分组")
        Control.Visible = blnHavePrivs: Control.Enabled = blnHavePrivs
    Case conMenu_Edit_ChargeBook_Reprint ' "重打收款收据(&R)")
        blnEanbled = lngPage = EM_PG_历史列表
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "重打收款收据") And zlStr.IsHavePrivs(mstrPrivs, "收款收据打印")
        Control.Visible = blnHavePrivs And blnEanbled
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_DrawBook_Reprint ' "重打备用金领用单
        blnEanbled = lngPage = EM_PG_备用金列表
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "重打备用金领用单") _
            And zlStr.IsHavePrivs(mstrPrivs, "备用金领用单打印")
        Control.Visible = blnHavePrivs And blnEanbled
        If blnEanbled Then blnEanbled = mfrmStandbyMoney.IsAllowCancel
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_View_Detail '查看明细数据
        blnEanbled = lngPage <> EM_PG_备用金列表
        Control.Visible = blnEanbled
        If blnEanbled Then
            If lngPage <> EM_PG_历史列表 Then
                blnEanbled = mfrmCollect.IsAllowViewChargeList
            Else
                blnEanbled = mfrmHistory.IsAllowViewChargeList
            End If
        End If
        Control.Enabled = blnEanbled
    Case conMenu_View_LargeICO  '大图标(&G)
        Control.Visible = lngPage = EM_PG_收款列表
        If Control.Visible Then
            intView = mfrmCollect.zlPersonListShowMode
            Control.Checked = intView = 0
        End If
    Case conMenu_View_MinICO  '小图标(&M)
        Control.Visible = lngPage = EM_PG_收款列表
        If Control.Visible Then
            intView = mfrmCollect.zlPersonListShowMode
            Control.Checked = intView = 1
        End If
    Case conMenu_View_ListICO  '列表(&L)
        Control.Visible = lngPage = EM_PG_收款列表
        If Control.Visible Then
            intView = mfrmCollect.zlPersonListShowMode
            Control.Checked = intView = 2
        End If
    Case conMenu_View_DetailsICO  '详细资料(&D)
        Control.Visible = lngPage = EM_PG_收款列表
        If Control.Visible Then
            intView = mfrmCollect.zlPersonListShowMode
            Control.Checked = intView = 3
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    End Select
End Sub
Private Function CheckDepend() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据依赖性
    '返回:数据合法,返回true，否则返回False
    '编制:刘兴洪
    '日期:2013-09-04 17:10:03
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    CheckDepend = False
    On Error GoTo errHandle
    Set rsTemp = Get结算方式
    rsTemp.Filter = "性质=1"
    If rsTemp.EOF Then
        rsTemp.Filter = 0
        ShowMsgbox "结算方式中不存在一条件有现金性质的结算方式,请在结算方式管理中设置!"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Filter = 0
    rsTemp.Close
    If UserInfo.姓名 = "" Then
        MsgBox "当前登录用户未指定对应的人员,不能使用本功能。", vbExclamation, gstrSysName
        Exit Function
    End If
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub zlPrintRpt(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出列表
    '入参:bytMode=1-打印,2-预览,3-输出到Excel
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-13 10:23:30
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    If Val(tbPage.Selected.Tag) = EM_PG_收款列表 Then
        '打印轧帐信息
        Call mfrmCollect.zlPrint(bytMode)
        Exit Sub
    End If
    If Val(tbPage.Selected.Tag) = EM_PG_备用金列表 Then
        mfrmStandbyMoney.zlPrint (bytMode): Exit Sub
    End If
    Call mfrmHistory.zlPrint(bytMode)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RePrintBill(ByVal bytRePrintType As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打单据
    '入参:bytRePrintType-0-收据;1-备用金领用单
    '编制:刘兴洪
    '日期:2013-09-13 16:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    If Val(tbPage.Selected.Tag) = EM_PG_收款列表 Then Exit Sub
    If Val(tbPage.Selected.Tag) = EM_PG_备用金列表 Then
        mfrmStandbyMoney.RePrintBill: Exit Sub
    End If
    Call mfrmHistory.RePrintBill
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub CheckCash()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:现金点钞
    '编制:刘兴洪
    '日期:2013-09-13 16:08:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim objCash As New clsChargeBill
    If Val(tbPage.Selected.Tag) = EM_PG_收款列表 Then
        dblMoney = mfrmCollect.GetCashMoney
    End If
    objCash.CheckCash Me, dblMoney
    Set objCash = Nothing
End Sub

Private Sub zlRefresh()
    '重新进行数据刷新
    If Val(tbPage.Selected.Tag) = EM_PG_收款列表 Then
        Call mfrmCollect.zlRefresh
    ElseIf Val(tbPage.Selected.Tag) = EM_PG_备用金列表 Then
         Call mfrmStandbyMoney.zlRefresh: Exit Sub
    Else
        Call mfrmHistory.zlRefresh
    End If
End Sub

Private Sub ShowChargeList()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细收款数据
    '编制:刘兴洪
    '日期:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
     If Val(tbPage.Selected.Tag) = EM_PG_备用金列表 Then Exit Sub
    If Val(tbPage.Selected.Tag) = EM_PG_收款列表 Then
         Call mfrmCollect.ShowChargeList(Me)
         Exit Sub
    End If
    '历史数据显示
    Call mfrmHistory.ShowChargeList(Me)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CallCustomRpt(ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用自定义报表
    '入参:lngSys-系统号
    '        strRptCode-报表编号
    '编制:刘兴洪
    '日期:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Val(tbPage.Selected.Tag) = EM_PG_收款列表 Then
         Call mfrmCollect.CallCustomRpt(Me, lngSys, strRptCode)
         Exit Sub
    End If
    If Val(tbPage.Selected.Tag) = EM_PG_备用金列表 Then
         Call mfrmStandbyMoney.CallCustomRpt(Me, lngSys, strRptCode)
    End If
    '历史数据显示
    Call mfrmHistory.CallCustomRpt(Me, lngSys, strRptCode)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetPersonListShow(ByVal intICOType As Integer)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置列表的显示方式
    '入参:intType-图标类型(0-大图标;1-小图标;2-列表;3-详细资料)
    '编制:刘兴洪
    '日期:2013-09-27 15:30:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Val(tbPage.Selected.Tag) <> EM_PG_收款列表 Then Exit Sub
    mfrmCollect.zlPersonListShowMode = intICOType
End Sub

Private Sub ExcuteOnWorkStandbyMoeny()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行发放上岗备用金操作
    '编制:刘尔旋
    '日期:2013-12-4
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmStandbyMoney Is Nothing Then Exit Sub
    If mfrmStandbyMoney.zlPayOnWorkMoney(Me) = False Then Exit Sub
End Sub

 Private Sub ExcutePutOutStandbyMoeny()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行发放备用金操作
    '编制:刘兴洪
    '日期:2013-10-12 16:49:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmStandbyMoney Is Nothing Then Exit Sub
    If mfrmStandbyMoney.zlPayStandbyMoney(Me) = False Then Exit Sub
 End Sub
 Private Sub ExcutePutINStandbyMoeny()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行收回发放备用金操作
    '编制:刘兴洪
    '日期:2013-10-12 16:49:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmStandbyMoney Is Nothing Then Exit Sub
    If mfrmStandbyMoney.CancelStandbyMoney() = False Then Exit Sub
 End Sub
Private Sub ExcuteRollingCurtainZero()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:轧帐归零处理
    '编制:刘兴洪
    '日期:2013-10-14 11:49:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDate As String, strSQL As String, rsTemp As ADODB.Recordset
    Dim str收费员 As String, blnStrans As Boolean, intStep As Integer, intCount As Integer
    Dim strRand As String, strTemp As String, lngTop As Long, lngLeft As Long
    Dim i As Long, intTemp As Integer
    On Error GoTo errHandle
    '提示警告
    If zlStr.IsHavePrivs(mstrPrivs, "轧帐归零") = False Then Exit Sub
    Randomize
    For i = 1 To 3
       intTemp = Asc("A") + Int(Rnd * 10)
        If intTemp > Asc("Z") Then intTemp = Asc("A")
       strRand = strRand & Chr(intTemp)
    Next
    
    lngLeft = Me.Left + 2500: lngTop = Me.Top + 2500
    strTemp = InputBox("  轧帐归零操作将会清除所有收费人员的缴款数据, " & _
                          "如果你不明该功能作用，请不要使用该功能。 " & _
                          "如果你确认要清除轧帐功能,请输入如下字符:" & vbCrLf & " " & vbCrLf & " " & _
                          "" & strRand, "警告", "", lngLeft, lngTop)
    If strTemp = "" Then Exit Sub
    If UCase(strTemp) <> UCase(strRand) Then
         MsgBox "输入错误,不允许继续操作!", vbInformation + vbOKOnly, gstrSysName
         Exit Sub
    End If
    frmWait.OpenWait Me, "轧账归零处理", True
    frmWait.WaitInfo = "正在提取数据..."
    strSQL = "Select Distinct 收款员 From 人员缴款余额 Where 性质=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.EOF Then
        frmWait.CloseWait
        MsgBox "没有要轧帐归零的数据处理，不需要轧帐归零", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        Exit Sub
    End If
    str收费员 = ""
    gcnOracle.BeginTrans
    blnStrans = True
    intStep = 1: intCount = rsTemp.RecordCount
    With rsTemp
        Do While Not .EOF
            frmWait.WaitInfo = "正在清除[" & Nvl(!收款员) & "的缴款数据..."
            If zlCommFun.ActualLen(str收费员 & "," & Nvl(!收款员)) > 4000 Then
                str收费员 = Mid(str收费员, 2)
                ' Zl_轧帐归零记录_Insert
                strSQL = "Zl_轧帐归零记录_Insert("
                '  登记人_In   In 人员收缴记录.登记人%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '  登记时间_In In 人员收缴记录.登记时间%Type,
                strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
                '  收费员_In   In Varchar2 := Null
                '  收费员_in-指定的收费员,为空时,为所有收费员;非空时,为指定的收费员(可以为多个,多个用逗号分隔
                strSQL = strSQL & "'" & str收费员 & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            str收费员 = str收费员 & "," & Nvl(!收款员)
            frmWait.pgb.Value = intStep \ intCount
            intStep = intStep + 1
            .MoveNext
        Loop
    End With
    If intCount = 0 Then intCount = 1
    frmWait.pgb.Value = intStep \ intCount
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If str收费员 <> "" Then
        str收费员 = Mid(str收费员, 2)
        ' Zl_轧帐归零记录_Insert
        strSQL = "Zl_轧帐归零记录_Insert("
        '  登记人_In   In 人员收缴记录.登记人%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  登记时间_In In 人员收缴记录.登记时间%Type,
        strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  收费员_In   In Varchar2 := Null
        '  收费员_in-指定的收费员,为空时,为所有收费员;非空时,为指定的收费员(可以为多个,多个用逗号分隔
        strSQL = strSQL & "'" & str收费员 & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    gcnOracle.CommitTrans: blnStrans = False
    frmWait.CloseWait
    
    MsgBox "轧帐归零操作成功!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
    Exit Sub
errHandle:
    frmWait.CloseWait
    If blnStrans Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub ExcuteSplitGroup()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行财务分组
    '编制:刘兴洪
    '日期:2013-10-15 15:22:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If frmGroupAndPesons.ShowGroups(Me, mlngModule, mstrPrivs) = False Then Exit Sub
    '重新加载财务组
    Call mfrmCollect.zlRefresh
End Sub
