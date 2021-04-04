VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frm卫材发放管理_New 
   Caption         =   "卫材发放管理"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   Icon            =   "frm卫材发放管理_New.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsColSet 
      Height          =   2865
      Left            =   5505
      TabIndex        =   4
      Top             =   1230
      Visible         =   0   'False
      Width           =   2655
      _cx             =   4683
      _cy             =   5054
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm卫材发放管理_New.frx":08CA
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   1
      ExplorerBar     =   2
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   135
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   1
      Top             =   1230
      Width           =   5070
      Begin VB.CheckBox Chk清单 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "显示所有过程单据"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3105
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   1935
      End
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6060
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm卫材发放管理_New.frx":0917
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11959
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
      Left            =   0
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frm卫材发放管理_New.frx":11AB
      Left            =   615
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frm卫材发放管理_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrm未发料  As frm卫材未发料清单
Attribute mfrm未发料.VB_VarHelpID = -1
Private mfrm发料汇总 As frm卫材发料汇总
Private mfrm缺料清单 As frm卫材缺料清单
Private mfrm拒发清单 As frm卫材拒发料清单
Private WithEvents mfrm退料清单 As frm卫材退料清单
Attribute mfrm退料清单.VB_VarHelpID = -1
Private WithEvents mfrmFilter As frm卫材发放过滤
Attribute mfrmFilter.VB_VarHelpID = -1
Private mstrSelectTabItem As String
Private Const ID_PANE_SEARCH = 201
Private Const conMenu_Popu_住院号 = 1011
Private Const conMenu_Popu_床号 = 1012
Private Const conMenu_Popu_姓名 = 1013
Private Const conMenu_Popu_病人ID = 1014
Private Const conMenu_Popu_门诊号 = 1015
Private Const conMenu_Popu_IC卡号 = 1016

Private mArrFilter As Variant   '过滤条件
Private mcbrControl As CommandBarControl
Private mcbrMenuBar As CommandBarPopup
Private mcbrToolBar As CommandBar
Private mrsNotPayStuff As ADODB.Recordset   '发料数据集
Private mrsChargeOff As New ADODB.Recordset                   '用于显示销帐申请记录
Private mrsBakStuff As ADODB.Recordset
Private mstrPrivs As String
Private mlngModule As Long
Private mintUnit As Integer     '单位:0-散装单位;1-包装单位
Private mint字号 As Integer     '当前字号:0-9,1-12,2-15
Private mbln页签 As Boolean
Private mint页签 As Integer

Private Enum mPage
    pag_未发清单 = 0
    pag_汇总发料 = 1
    pag_缺料清单 = 2
    pag_拒发清单 = 3
    pag_退料清单 = 4
End Enum
'------------------------------------------------------------------------------------------------------------
'从材料处方发药传过来的参数
Private mblnTrans As Boolean            'True表示从材料处方发药窗口调用
Private mstrNo  As String               '单据号，仅用于定位
Private mlng库房id As Long              '发药库房ID，一般和发料部门一致
Private mlng病人id As Long              'mlng病人id
Private mstrStuffStartDate As String     '材料单据开始时间
Private mstrStuffEndDate As String       '材料单据结束时间

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString

Private mint输入模式 As Integer

Private mobjPlugIn As Object             '外挂接口对象
'----------------------------------------------------------------------------------------------------------
Public Sub ShowList(ByVal frmMain As Form, ByVal lng病人id As Long, ByVal strNo As String, ByVal lng库房ID As Long, ByVal strStartDate As String, ByVal strEndDate As String)
    '-----------------------------------------------------------------------------------------------------------
    '功能:发药窗品调用
    '入参:frmMain-主窗口
    '     lng病人ID-病人ID
    '     strNo-处方号
    '     lng库房id-库房ID
    '     strStartDate-开始日期
    '     strEndDate-结束日期
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-01 21:51:18
    '-----------------------------------------------------------------------------------------------------------
    
    mlng病人id = lng病人id
    mstrNo = strNo
    mlng库房id = lng库房ID
    mstrStuffStartDate = strStartDate
    mstrStuffEndDate = strEndDate
    mblnTrans = True
    Me.Show , frmMain
    Me.ZOrder 0
End Sub

Private Sub initLocalPara()
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取初始化本地参数变量值
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-01 21:11:08
    '-----------------------------------------------------------------------------------------------------------
    mintUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mint字号 = zlDatabase.GetPara("字体字号", glngSys, mlngModule, "0")
    mbln页签 = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "卫材发放管理", "保持上一次窗体关闭时的页签", 0)) = 1)
    mint页签 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "卫材发放管理", "当前页签", 0))
    
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    
    mint输入模式 = 0
    If Val(zlDatabase.GetPara("使用个性化风格")) <> 0 Then
        mint输入模式 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "卫材发放管理", "输入模式", "0"))
        If mint输入模式 < 0 Then
            mint输入模式 = 0
        End If
    End If
End Sub
 Private Sub InitPage()
    '------------------------------------------------------------------------------
    '功能:初始化页面控件
    '返回:
    '编制:刘兴宏
    '日期:2007/08/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim stdfnt As StdFont
    Dim objItem As TabControlItem
    
    
    Set mfrm未发料 = New frm卫材未发料清单
    Set objItem = tbPage.InsertItem(mPage.pag_未发清单, "未发料清单", mfrm未发料.hwnd, 0)
    objItem.Tag = mPage.pag_未发清单
    Set mfrm发料汇总 = New frm卫材发料汇总
    Set objItem = tbPage.InsertItem(mPage.pag_汇总发料, "汇总发料", mfrm发料汇总.hwnd, 0)
    objItem.Tag = mPage.pag_汇总发料
    
    Set mfrm缺料清单 = New frm卫材缺料清单
    Set objItem = tbPage.InsertItem(mPage.pag_缺料清单, "缺料清单", mfrm缺料清单.hwnd, 0)
    objItem.Tag = mPage.pag_缺料清单
    
    Set mfrm拒发清单 = New frm卫材拒发料清单
    Set objItem = tbPage.InsertItem(mPage.pag_拒发清单, "拒发料清单", mfrm拒发清单.hwnd, 0)
    objItem.Tag = mPage.pag_拒发清单
    
    Set mfrm退料清单 = New frm卫材退料清单
    Set objItem = tbPage.InsertItem(mPage.pag_退料清单, "退料清单", mfrm退料清单.hwnd, 0)
    objItem.Tag = mPage.pag_退料清单
    Call mfrmFilter_zlRefreshCon(mArrFilter)
 
    With tbPage
        If mint页签 <> 0 And mbln页签 Then
            .Item(mint页签).Selected = True
        Else
            .Item(0).Selected = True
        End If
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        Set stdfnt = Me.Font
        Set .PaintManager.Font = stdfnt
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function InitComandBars() As Boolean
    '----------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/1/9
    '----------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl

    Dim panThis As Pane
    err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    
    
'    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.Id = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_BillPrint, "打印发料单据(&B)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_BillPrintView, "打印退料通知单(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.Id = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "发料(&A)"):        mcbrControl.IconId = 3010
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Discard, "拒发(&H)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Recall, "恢复(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Untread, "退料(&T)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelAudit, "销帐(&S)"):  mcbrControl.IconId = 21905:  mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CfPay, "按处方发料(&C)"):  mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BillPay, "按票据号发料(&B)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BillBackPay, "按单据退料(&N)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_OtherPay, "发其他库房处方(&N)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StopPay, "停止发料标记(&S)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "停止发料") Or zlStr.IsHavePrivs(mstrPrivs, "恢复发料")
        mcbrControl.BeginGroup = mcbrControl.Visible
    End With

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.Id = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
    
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_FontSize, "字体(&F)")
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_FontSize_1, "小字体(&S)", -1, False)
        If mint字号 = 0 Then cbrControl.Checked = True
        cbrControl.Parameter = 0
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_FontSize_2, "中字体(&M)", -1, False)
        If mint字号 = 1 Then cbrControl.Checked = True
        cbrControl.Parameter = 1
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_FontSize_3, "大字体(&B)", -1, False)
        If mint字号 = 2 Then cbrControl.Checked = True
        cbrControl.Parameter = 2
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Column, "列设置(&C)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.Id = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With

    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With

    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With

    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "发料"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3010
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Discard, "拒发")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Recall, "恢复")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Untread, "退料"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelAudit, "销帐"): mcbrControl.IconId = 21905: mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    InitComandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitPanel()
    Dim PaneSearch As Pane
    If mfrmFilter Is Nothing Then Set mfrmFilter = New frm卫材发放过滤
    mfrmFilter.Set发料窗口条件 mblnTrans, mstrNo, mstrStuffStartDate, mstrStuffEndDate, mlng病人id, mlng库房id
    Set mArrFilter = mfrmFilter.GetFilterCon
    
    With dkpMan
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        
        '.Options.HideClient = True
        Set PaneSearch = .CreatePane(ID_PANE_SEARCH, 400, 100, DockTopOf, Nothing)
        PaneSearch.Title = "过滤"
       ' PaneSearch.Options = PaneNoCloseable
        '.ImageList = imlPaneIcons '
        .SetCommandBars cbsThis
    End With
End Function
Private Function SendBillPay(ByVal bln票据号发料 As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:按处方发料
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-06 14:55:11
    '-----------------------------------------------------------------------------------------------------------
    With Frm按单号发料
        .In_单据 = 0
        .In_单据IN = mArrFilter("单据")
        .In_发料部门id = Val(mArrFilter("发料部门ID"))
        .In_库存检查 = GetCheckPara()
        .In_允许未配料发料 = 1
        .In_权限 = mstrPrivs
        .mstr配料人 = gstrUserName
        .按票据号发料 = bln票据号发料
        Set .In_PlugIn = mobjPlugIn
        .Show 1, Me
    End With
    SendBillPay = True
    Call mfrmFilter_zlRefreshCon(mArrFilter)
    
End Function
Private Function SendBackPay() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:按单据号退料
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-06 15:07:05
    '-----------------------------------------------------------------------------------------------------------
    Set Frm按单号退料.In_PlugIn = mobjPlugIn
    If Frm按单号退料.ShowCard(Me, Val(mArrFilter("发料部门ID")), mstrPrivs) = False Then Exit Function
    SendBackPay = True
    Call mfrmFilter_zlRefreshCon(mArrFilter)
End Function
Private Sub StopPayStuffFlag()
    '-----------------------------------------------------------------------------------------------------------
    '功能:停止发料标记
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-06 15:09:00
    '-----------------------------------------------------------------------------------------------------------
    '停止发料
    '发药方式=-1
    Dim frmFlag As New Frm不再发药处方标志
    frmFlag.In_库存检查 = GetCheckPara
    
    '--50313（zdt）:对停止发料的发料部门id进行赋值
    frmFlag.In_库房id = Val(mArrFilter("发料部门id"))
    frmFlag.In_请求类型 = Val(mArrFilter("请求类型"))
    
    frmFlag.gstrParentName = Replace(Me.Name, "_New", "")
    frmFlag.Show vbModal
    Call mfrmFilter_zlRefreshCon(mArrFilter)
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long
    Dim blnAskPring As Boolean
    
    Dim cllFind As Collection
    
    '------------------------------------
    Select Case Control.Id
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    
    
    Case conMenu_File_BillPrint
        '发料单打印
        If tbPage.Selected Is Nothing Then
            blnAskPring = True
        Else
            blnAskPring = Not (Val(tbPage.Selected.Tag) = mPage.pag_退料清单)
        End If
        Call mfrm退料清单.zlPrintBill(True, "", IIf(mfrm未发料.cboEdit(0).ListIndex = -1, 1, mfrm未发料.cboEdit(0).ListIndex + 1), mstrPrivs, blnAskPring)
        
    Case conMenu_File_BillPrintView
        '打印退料通知单
        If tbPage.Selected Is Nothing Then
            blnAskPring = True
        Else
            blnAskPring = Not (Val(tbPage.Selected.Tag) = mPage.pag_退料清单)
        End If
        Call mfrm退料清单.zlPrintBill(False, "", 1, mstrPrivs, blnAskPring)
    Case conMenu_File_Parameter:
        '参数设置
        If frmPayExitParaSet.ShowSetPara(Me, mlngModule, mstrPrivs) = False Then Exit Sub
        Call initLocalPara
        Set mArrFilter = mfrmFilter.GetFilterCon
        Call mfrmFilter_zlRefreshCon(mArrFilter)
    
    Case conMenu_File_Exit: Unload Me
    Case conMenu_Edit_NewItem:       '发料
        Set mfrm未发料.In_PlugIn = mobjPlugIn
        If mfrm未发料.zlPayStuff = False Then Exit Sub
        Call mfrmFilter_zlRefreshCon(mArrFilter)
    Case conMenu_Manage_Discard:  '拒发
        '拒发
        If Save拒发 = False Then Exit Sub
        Call mfrmFilter_zlRefreshCon(mArrFilter)
    Case conMenu_Manage_Recall  '恢复
        If mfrm拒发清单.zlRestorePayStuff = False Then Exit Sub
        Call mfrmFilter_zlRefreshCon(mArrFilter)
    Case conMenu_Manage_Untread     '退料
        Set mfrm退料清单.In_PlugIn = mobjPlugIn
       If mfrm退料清单.zlBackPayStuff = False Then Exit Sub
       Call mfrmFilter_zlRefreshCon(mArrFilter)
    Case conMenu_Edit_ChargeDelAudit    '销帐
        Set frm卫材销帐.In_PlugIn = mobjPlugIn
        If frm卫材销帐.ShowList(Me, mstrPrivs, mlngModule, Val(mArrFilter("发料部门ID")), mintUnit) = False Then Exit Sub
        
    Case conMenu_Edit_CfPay    '按处方发料
        Call SendBillPay(False)
    Case conMenu_Edit_BillPay    '按票据号发料
        Call SendBillPay(True)
    Case conMenu_Edit_BillBackPay    '按单据退料
        Call SendBackPay
    Case conMenu_Edit_OtherPay       '发其他库房处方
        Call SendOtherPay
    Case conMenu_Edit_StopPay    '停止发料标记
        Call StopPayStuffFlag
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each mcbrControl In Me.cbsThis(2).Controls
            mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_FontSize_1, conMenu_View_FontSize_2, conMenu_View_FontSize_3
        mint字号 = Val(Control.Parameter)
        Call SetFontSize
        Call zlDatabase.SetPara("字体字号", mint字号, glngSys, mlngModule)
    Case conMenu_View_Refresh   '刷新
        mstrSelectTabItem = "," & Val(tbPage.Selected.Tag)
       Set mArrFilter = mfrmFilter.GetFilterCon
       Call mfrmFilter_zlRefreshCon(mArrFilter)
    Case conMenu_View_Column '列设置
        Call LoadFulltoColSel
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
'    Case conMenu_Popu_住院号
'        mfrmFilter.PatiTittle = 0
'    Case conMenu_Popu_姓名
'        mfrmFilter.PatiTittle = 1
'    Case conMenu_Popu_床号
'        mfrmFilter.PatiTittle = 2
'    Case conMenu_Popu_病人ID    '
'        mfrmFilter.PatiTittle = 3
'    Case conMenu_Popu_门诊号
'        mfrmFilter.PatiTittle = 4
'    Case conMenu_Popu_就诊卡号
'        mfrmFilter.PatiTittle = 5
'    Case conMenu_Popu_IC卡号
'        mfrmFilter.PatiTittle = 6
    Case Else
        If Control.Id > 401 And Control.Id < 499 Then
            '相关报表执行
            Call OpenRpt(Control)
        End If
        
        '弹出菜单
        If Control.Id >= conMenu_Popu_住院号 And Control.Id <= conMenu_Popu_住院号 + 6 + gintCardCount Then
            mint输入模式 = Control.Id - conMenu_Popu_住院号
'            mfrmFilter.PatiTittle = mint输入模式
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub SetFontSize()
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-05-06 17:24:45
    '-----------------------------------------------------------------------------------------------------------
    Dim curFontSize As Currency
    Dim stdfnt As StdFont
    
    curFontSize = Decode(mint字号, 1, 11, 2, 15, 9)
    mfrm发料汇总.zlSetFontSize curFontSize
    mfrm拒发清单.zlSetFontSize curFontSize
    mfrm缺料清单.zlSetFontSize curFontSize
    mfrm退料清单.zlSetFontSize curFontSize
    mfrm未发料.zlSetFontSize curFontSize
     If Not tbPage.PaintManager.Font Is Nothing Then
        With tbPage
            Set stdfnt = .PaintManager.Font
            stdfnt.Size = curFontSize
             Set .PaintManager.Font = stdfnt
              .PaintManager.Layout = xtpTabLayoutAutoSize
        End With
    End If
    Me.FontSize = curFontSize
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then
        Bottom = stbThis.Height
    End If
End Sub

Private Sub cbsThis_Resize()
    Dim sngStatusHeight As Single
    On Error Resume Next
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    cbsThis.GetClientRect Left, Top, Right, Bottom '
    With picList
        .Left = Left
        .Top = Top
        .Width = Right - Left
        .Height = Bottom - Top
    End With
End Sub
Private Function ISHaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取是否有相关的数据
    '入参:
    '出参:
    '返回:有,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-02 00:22:47
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim vsList As VSFlexGrid
    Dim lngRow As Long
    err = 0: On Error GoTo ErrHand:
    If tbPage.Selected Is Nothing Then Exit Function
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_汇总发料
       ISHaveData = mfrm发料汇总.zlHaveData
    Case mPage.pag_拒发清单
       ISHaveData = mfrm拒发清单.zlHaveData
    Case mPage.pag_未发清单
       ISHaveData = mfrm未发料.zlHaveData
    Case mPage.pag_缺料清单
       ISHaveData = mfrm缺料清单.zlHaveData
    Case mPage.pag_退料清单
       ISHaveData = mfrm退料清单.zlHaveData
    End Select
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub 权限控制(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '权限控制
  
  Select Case Control.Id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
    Case conMenu_File_BillPrint, conMenu_File_BillPrintView
    Case conMenu_File_Parameter:
        '参数设置
       ' Control.Enabled = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    Case conMenu_Edit_NewItem, conMenu_Edit_CfPay, conMenu_Edit_BillPay:     '发料
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "卫生材料发料")
    Case conMenu_Edit_StopPay
         Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "停止发料") Or zlStr.IsHavePrivs(mstrPrivs, "恢复发料")
    Case conMenu_Manage_Discard:  '拒发
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "卫生材料拒发")
    Case conMenu_Manage_Recall  '恢复发料
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "卫生材料恢复")
    Case conMenu_Manage_Untread, conMenu_Edit_BillBackPay    '退料
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "卫生材料退料")
    Case conMenu_EditPopup
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "卫生材料发料") Or _
                          zlStr.IsHavePrivs(mstrPrivs, "卫生材料退料") Or _
                          zlStr.IsHavePrivs(mstrPrivs, "卫生材料拒发") Or _
                          zlStr.IsHavePrivs(mstrPrivs, "卫生材料恢复") Or _
                          zlStr.IsHavePrivs(mstrPrivs, "卫生材料退料") Or _
                          zlStr.IsHavePrivs(mstrPrivs, "卫生材料销帐")
    Case conMenu_Edit_ChargeDelAudit    '销帐
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "卫生材料销帐")
    End Select
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '设置控件的相关属性
   Call 权限控制(Control)
   Select Case Control.Id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = ISHaveData
    Case conMenu_File_BillPrintView ' conMenu_File_BillPrint,
        If tbPage.Selected Is Nothing Then
            Control.Enabled = False
        Else
            Control.Enabled = Val(tbPage.Selected.Tag) = mPage.pag_退料清单
        End If
    Case conMenu_File_Parameter:
        '参数设置
       ' Control.Enabled = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    Case conMenu_Edit_NewItem:       '发料
        Control.Enabled = mfrm未发料.zlHaveSel发料 And (Val(tbPage.Selected.Tag) = mPage.pag_未发清单 Or Val(tbPage.Selected.Tag) = mPage.pag_汇总发料)
    Case conMenu_Manage_Discard:  '拒发
        '        mfrm未发料.zl
        Control.Enabled = mfrm未发料.zlHaveSel拒发 And (Val(tbPage.Selected.Tag) = mPage.pag_未发清单)
    Case conMenu_Manage_Recall
        '恢复
        Control.Enabled = mfrm拒发清单.zlHaveSel恢复 And (Val(tbPage.Selected.Tag) = mPage.pag_拒发清单)
    Case conMenu_Manage_Untread     '退料
        Control.Enabled = mfrm退料清单.zlHaveSel退料 And (Val(tbPage.Selected.Tag) = mPage.pag_退料清单)
    Case conMenu_Edit_ChargeDelAudit    '销帐
    Case conMenu_View_FontSize_1, conMenu_View_FontSize_2, conMenu_View_FontSize_3
        Control.Checked = Val(Control.Parameter) = mint字号
    
    End Select
End Sub

Private Sub Chk清单_Click()
     mfrm退料清单.zl显示整过程单据 = Chk清单.Value = 1
End Sub

Private Sub Form_Activate()
    If mfrmFilter.CheckDept = False Then
        ShowMsgBox "至少应该设置一个具有发料部门性质或者" & vbCrLf & "你不是发料部门的工作人员,请查看部门管理！"
        Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs
    mlngModule = glngModul

    '一卡通接口
    On Error Resume Next
    Set gobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Not gobjSquareCard Is Nothing Then
        If gobjSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle) = False Then
            Set gobjSquareCard = Nothing
        Else
            gstrCardType = gobjSquareCard.zlGetIDKindStr
            
            '取“就诊卡”类别和之后的消费卡
            gstrCardType = Mid(gstrCardType, InStr(1, gstrCardType, "就|就诊卡"))
        End If
    End If
    
    err.Clear: On Error GoTo 0
    
    Call initLocalPara
    Call InitComandBars
    Call InitPanel
    Call InitPage
    Call SetFontSize
    RestoreWinState Me, App.ProductName
    '2008-03-12:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '恢复录入状态
'    mfrmFilter.PatiTittle = mint输入模式

    '发药业务外挂部件
    Call zlPlugIn_Ini(glngSys, mlngModule, mobjPlugIn)
End Sub


Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
    Case ID_PANE_SEARCH
        Item.Handle = mfrmFilter.hwnd
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnTrans = False
    
    '卸载一卡通接口
    gstrCardType = ""
    Set gobjSquareCard = Nothing
    
    If Not mfrmFilter Is Nothing Then Unload mfrmFilter
    If Not mfrm发料汇总 Is Nothing Then Unload mfrm发料汇总
    If Not mfrm拒发清单 Is Nothing Then Unload mfrm拒发清单
    If Not mfrm缺料清单 Is Nothing Then Unload mfrm缺料清单
    If Not mfrm退料清单 Is Nothing Then Unload mfrm退料清单
    If Not mfrm未发料 Is Nothing Then Unload mfrm未发料
    
    Set mfrmFilter = Nothing
    Set mfrm发料汇总 = Nothing
    Set mfrm拒发清单 = Nothing
    Set mfrm缺料清单 = Nothing
    Set mfrm退料清单 = Nothing
    Set mfrm未发料 = Nothing
    mstrNo = "": mstrStuffStartDate = "": mstrStuffEndDate = "": mlng病人id = 0: mlng库房id = 0
    
    Call SaveWinState(Me, App.ProductName)
    
    '保存输入模式
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "卫材发放管理", "输入模式", mint输入模式)
    If mbln页签 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\卫材发放管理", "当前页签", tbPage.Selected.Index)
    End If
    
    '卸载外挂接口
    Call zlPlugIn_Unload(mobjPlugIn)
End Sub

Private Sub mfrmFilter_zlPopupMenus(ByVal x As Long, ByVal Y As Long)
    '弹出菜单
'    Dim intType As Integer
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim intCount As Integer
    Dim strCardName As String
    
  '  If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
'    mint输入模式 = mfrmFilter.PatiTittle
    With cbrPopupBar
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_住院号, "住院号(&A)")
        If mint输入模式 = 0 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_床号, "床号(&C)")
        If mint输入模式 = 1 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_姓名, "姓名(&N)")
        If mint输入模式 = 2 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_病人ID, "病人ID(&I)")
        If mint输入模式 = 3 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_门诊号, "门诊号(&M)")
        If mint输入模式 = 4 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_IC卡号, "IC卡号(&K)")
        If mint输入模式 = 5 Then cbrPopupItem.Checked = True
        
        '动态取其他消费卡
        If gstrCardType <> "" Then
            gintCardCount = UBound(Split(gstrCardType, ";")) + 1
            For intCount = 0 To UBound(Split(gstrCardType, ";"))
                '取银行卡名称
                strCardName = Split(Split(gstrCardType, ";")(intCount), "|")(1)
                
                Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_IC卡号 + intCount + 1, strCardName & "(&" & intCount + 1 & ")")
                
                If mint输入模式 = conMenu_Popu_IC卡号 + intCount + 1 Then
                    cbrPopupItem.Checked = True
                End If
                
                '保存卡信息
                cbrPopupItem.Parameter = Split(gstrCardType, ";")(intCount)
                
                If intCount = 0 Then
                    cbrPopupItem.BeginGroup = True
                End If
            Next
        End If
        
    End With
    cbrPopupBar.ShowPopup
End Sub

Private Sub mfrmFilter_zlRefreshCon(ByVal arrFilter As Variant)
    Set mArrFilter = arrFilter
    
    '条件发生了改变
    Set mfrm退料清单.zlArrFilter = mArrFilter
    Call mfrm未发料.zlFullData(Me, mstrPrivs, mlngModule, mintUnit, arrFilter)
    Call mfrm拒发清单.zlRefreshData(Me, mstrPrivs, mlngModule, mintUnit, mArrFilter)
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_退料清单
        mfrm退料清单.zlRefreshData Me, mstrPrivs, mlngModule, mintUnit, mArrFilter
    Case mPage.pag_拒发清单
        mfrm拒发清单.zlFullData mrsNotPayStuff
    Case mPage.pag_汇总发料
        If mfrm发料汇总.zlFullData(mintUnit, mrsNotPayStuff, mrsChargeOff) = False Then Exit Sub
    Case Else
    End Select
    
End Sub

Private Sub mfrm退料清单_zlRefreshDataRecordSet(ByVal rsNotStuffStuff As ADODB.Recordset)
    Set mrsBakStuff = rsNotStuffStuff
    stbThis.Panels(2).Text = "共有" & mrsBakStuff.RecordCount & "条记录,上次汇总发料号为:" & mfrm未发料.zl_上次汇总发料号
End Sub

Private Sub mfrm未发料_zlRefreshDataRecordSet(ByVal rsNotStuffStuff As ADODB.Recordset, ByVal rsChargeOff As ADODB.Recordset)
    Set mrsNotPayStuff = rsNotStuffStuff
    Set mrsChargeOff = rsChargeOff
    stbThis.Panels(2).Text = "共有" & mrsNotPayStuff.RecordCount & "条待发料记录,上次汇总发料号为:" & mfrm未发料.zl_上次汇总发料号
End Sub

Private Sub picList_Resize()
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Width = .ScaleWidth
        tbPage.Top = .ScaleTop
        tbPage.Height = .ScaleHeight
        Chk清单.Top = tbPage.Top
        Chk清单.Left = .ScaleWidth - Chk清单.Width - 100
    End With
End Sub

Private Function Save拒发() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:拒发未生材料
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-24 11:59:21
    '-----------------------------------------------------------------------------------------------------------
    Dim cllProc As Collection
    Set cllProc = New Collection
    
    With mrsNotPayStuff
        .Filter = "执行状态=2"
        If .RecordCount = 0 Then
            ShowMsgBox "不存在相关的拒发情况，请选择拒发材料,操作中止!"
            Exit Function
        End If
        .Sort = "材料id Asc"
        .MoveFirst
        Do While Not .EOF '
            If !执行状态 = 2 Then
                'Zl_卫生材料发放_拒发(Id_In In 材料收发记录.ID%Type)
                gstrSQL = "Zl_卫生材料发放_拒发(" & NVL(!Id) & ")"
                AddArray cllProc, gstrSQL
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllProc, Me.Caption
    Save拒发 = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    
End Function
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, ObjAppRow As zlTabAppRow
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_未发清单
        Set objPrint.Body = mfrm未发料.vsGrid
    Case mPage.pag_汇总发料
        Set objPrint.Body = mfrm发料汇总.vsGrid
    Case mPage.pag_拒发清单
        Set objPrint.Body = mfrm拒发清单.vsGrid
    Case mPage.pag_缺料清单
        Set objPrint.Body = mfrm缺料清单.vsGrid
    Case mPage.pag_退料清单
        Set objPrint.Body = mfrm退料清单.vsGrid
    Case Else
        Exit Sub
    End Select
    
    objPrint.Title.Text = tbPage.Selected.Caption & "清册"
    Set ObjAppRow = New zlTabAppRow
    Call ObjAppRow.Add("打印人:" & gstrUserName)
    Call ObjAppRow.Add("打印时间:" & Format(Sys.Currentdate, "yyyy年MM月DD日"))
    Call objPrint.BelowAppRows.Add(ObjAppRow)
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
Private Function OpenRpt(ByVal Control As XtremeCommandBars.ICommandBarControl) As Boolean
    '------------------------------------------------------------------------------
    '功能:打开报表
    '参数:Control-执行报表的控件
    '返回:
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim arrData As Variant
    Dim strNo As String, intRecodeSta As Integer, lng发料部门ID As Long, lng位置 As Long
    Dim lng材料ID As Long, lng费用ID As Long, str住院号 As String, lng单据 As Long
    
    'CStr(gobjComLib.zlCommFun.NVL(rsTmp!系统, 0) &  "," & rsTmp!编号)
    arrData = Split(Control.Parameter, ",")
    'Set mrs拒发 = zldatabase.OpenSQLRecord(gstrsql, Me.Caption, _
           Val(mArrFilter("发料部门ID")), _
           CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), _
           CStr("," & mArrFilter("单据") & ","), _
           Val(mArrFilter("开单科室ID")), _
           CStr(mArrFilter("单据号")(0)), CStr(mArrFilter("单据号")(1)), _
           Val(mArrFilter("病人ID")), Val(mArrFilter("住院号")), _
           CStr(mArrFilter("姓名")))
    
    str住院号 = Val(mArrFilter("住院号"))
    lng发料部门ID = Val(mArrFilter("开单科室ID"))
        
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_未发清单
        With mfrm未发料.vsGrid
            lng位置 = Val(.Cell(flexcpData, .Row, .ColIndex("单据号")))
            mrsNotPayStuff.Find "位置=" & lng位置
            With mrsNotPayStuff
                If Not mrsNotPayStuff.EOF Then
                    lng材料ID = Val(NVL(!材料ID))
                    intRecodeSta = 1
                    lng费用ID = Val(NVL(!费用ID))
                    str住院号 = NVL(!住院号)
                    lng单据 = NVL(!单据)
                End If
            End With
        End With
    Case mPage.pag_汇总发料
        With mfrm发料汇总.vsGrid
            lng材料ID = Val(.Cell(flexcpData, .Row, .ColIndex("材料名称")))
        End With
    Case mPage.pag_拒发清单
        With mfrm拒发清单.vsGrid
            lng单据 = Val(.Cell(flexcpData, .Row, .ColIndex("单据类型")))
            lng材料ID = Val(.Cell(flexcpData, .Row, .ColIndex("材料名称")))
            str住院号 = .TextMatrix(.Row, .ColIndex("住院号"))
            lng费用ID = Val(.Cell(flexcpData, .Row, .ColIndex("单据号")))
            intRecodeSta = 1
        End With
    Case mPage.pag_缺料清单
        With mfrm缺料清单.vsGrid
            lng位置 = Val(.Cell(flexcpData, .Row, .ColIndex("单据号")))
            mrsNotPayStuff.Find "位置=" & lng位置
            With mrsNotPayStuff
                If Not mrsNotPayStuff.EOF Then
                    lng材料ID = Val(NVL(!材料ID))
                    intRecodeSta = 1
                    lng费用ID = Val(NVL(!费用ID))
                    str住院号 = NVL(!住院号)
                    lng单据 = NVL(!单据)
                End If
            End With
        End With
    Case mPage.pag_退料清单
        With mfrm退料清单.vsGrid
            lng位置 = Val(.Cell(flexcpData, .Row, .ColIndex("单据号")))
            mrsBakStuff.Find "位置=" & lng位置
            With mrsBakStuff
                If Not mrsBakStuff.EOF Then
                    lng材料ID = Val(NVL(!材料ID))
                    intRecodeSta = 1
                    lng费用ID = Val(NVL(!费用ID))
                    str住院号 = NVL(!住院号)
                    lng单据 = NVL(!单据)
                End If
            End With
        End With
    End Select
    '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
    Call ReportOpen(gcnOracle, Val(arrData(0)), arrData(1), Me, "NO=" & strNo, "记录状态=" & intRecodeSta, _
        "发料部门=" & Val(mArrFilter("发料部门ID")), "单据类型=" & lng单据, _
        "材料=" & lng材料ID, "费用=" & lng费用ID, "住院号=" & str住院号, _
        "开始日期=" & mArrFilter("日期范围")(0), "结束日期=" & mArrFilter("日期范围")(1))
End Function
Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
        
        Chk清单.Visible = False
        stbThis.Panels(2) = ""
        Select Case Val(Item.Tag)
        Case mPage.pag_汇总发料
            If mfrm发料汇总.zlFullData(mintUnit, mrsNotPayStuff, mrsChargeOff) = False Then Exit Sub
            stbThis.Panels(2).Text = "上次汇总发料号为:" & mfrm未发料.zl_上次汇总发料号
        Case mPage.pag_拒发清单
            If mfrm拒发清单.zlFullData(mrsNotPayStuff) = False Then Exit Sub
            stbThis.Panels(2).Text = "上次汇总发料号为:" & mfrm未发料.zl_上次汇总发料号
        Case mPage.pag_缺料清单
            If mfrm缺料清单.zlFullData(mintUnit, mrsNotPayStuff) = False Then Exit Sub
            stbThis.Panels(2).Text = "上次汇总发料号为:" & mfrm未发料.zl_上次汇总发料号
        Case mPage.pag_退料清单
            Chk清单.Visible = True
            Call mfrm退料清单.zlRefreshData(Me, mstrPrivs, mlngModule, mintUnit, mArrFilter)
            If mrsBakStuff Is Nothing Then Exit Sub
            If mrsBakStuff.State = 1 Then
                stbThis.Panels(2).Text = "共有" & mrsBakStuff.RecordCount & "条待记录 ,上次汇总发料号为:" & mfrm未发料.zl_上次汇总发料号
            End If
        Case mPage.pag_未发清单
            If mrsNotPayStuff Is Nothing Then Exit Sub
            If mrsNotPayStuff.State = 1 Then
                stbThis.Panels(2).Text = "共有" & mrsNotPayStuff.RecordCount & "条待发料记录,上次汇总发料号为:" & mfrm未发料.zl_上次汇总发料号
            End If
        End Select
        
End Sub


Private Function GetCheckPara() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取库存检查参数
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-06 15:02:49
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = " Select Nvl(检查方式,0) 库存检查 From 材料出库检查 Where 库房ID=[1]"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mArrFilter("发料部门id")))
    With rsTemp
        If Not .EOF Then
            GetCheckPara = NVL(!库存检查, 0)
        End If
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function LoadFulltoColSel() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载列设置
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-09 16:46:43
    '-----------------------------------------------------------------------------------------------------------
    Dim vsGrid As VSFlexGrid, i As Long, lngRow As Long
    Dim sngFrmHeight As Single, sngSelSumHeight As Single
    
    Select Case Val(Me.tbPage.Selected.Tag)
    Case mPage.pag_汇总发料
        Set vsGrid = mfrm发料汇总.vsGrid
    Case mPage.pag_拒发清单
        Set vsGrid = mfrm拒发清单.vsGrid
    Case mPage.pag_缺料清单
        Set vsGrid = mfrm缺料清单.vsGrid
    Case mPage.pag_退料清单
        Set vsGrid = mfrm退料清单.vsGrid
    Case mPage.pag_未发清单
        Set vsGrid = mfrm未发料.vsGrid
    End Select
    vsColSet.Clear 1
    vsColSet.Rows = 2
    With vsGrid
        lngRow = 1
        For i = 0 To .Cols - 1
            '.coldata(i):1-固定,-1-不能选,0-可选
            If Trim(.ColKey(i)) <> "" And (.ColData(i) = 1 Or .ColData(i) = 0) Then
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("列名")) = .ColKey(i)
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("选择")) = IIf(.ColWidth(i) = 0 Or .ColHidden(i), False, True)
                vsColSet.RowData(lngRow) = .ColData(i)
                If .ColData(i) = 1 Then
                    vsColSet.Cell(flexcpForeColor, lngRow, 0, lngRow, vsColSet.Cols - 1) = vbBlue
                End If
                vsColSet.Rows = vsColSet.Rows + 1
                lngRow = lngRow + 1
            End If
        Next
    End With
    If vsColSet.Rows > 2 Then vsColSet.Rows = vsColSet.Rows - 1
    SetParent vsColSet.hwnd, vsGrid.Parent.hwnd
    sngFrmHeight = vsGrid.Parent.ScaleHeight
    With vsColSet
        sngSelSumHeight = (.RowHeight(0) + 60) * (.Rows) + 60
        
        .Cell(flexcpBackColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000001
        .Cell(flexcpForeColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000005
        .BackColorSel = &H8000000D
        .Row = 1
        .Visible = True
        .Editable = flexEDKbdMouse
        .ZOrder 0
        .Left = vsGrid.Left + .Cell(flexcpWidth, 0, 0, 0, 0) + 30
        .Top = vsGrid.Top + vsGrid.RowHeight(0) + 15
        sngFrmHeight = sngFrmHeight - .Top
        
        If sngFrmHeight > sngSelSumHeight Then
            .Height = sngSelSumHeight
        Else
            .Height = IIf(sngFrmHeight < 0, 0, sngFrmHeight)
        End If
        .SetFocus
    End With
End Function
Private Function SetVsGridCol(ByVal strColKey As String, ByVal blnShow As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置显示列
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-09 17:31:22
    '-----------------------------------------------------------------------------------------------------------
    Dim vsGrid As VSFlexGrid, i As Long, lngRow As Long
    Select Case Val(Me.tbPage.Selected.Tag)
    Case mPage.pag_汇总发料
        Set vsGrid = mfrm发料汇总.vsGrid
    Case mPage.pag_拒发清单
        Set vsGrid = mfrm拒发清单.vsGrid
    Case mPage.pag_缺料清单
        Set vsGrid = mfrm缺料清单.vsGrid
    Case mPage.pag_退料清单
        Set vsGrid = mfrm退料清单.vsGrid
    Case mPage.pag_未发清单
        Set vsGrid = mfrm未发料.vsGrid
    End Select
    With vsGrid
        
        .ColHidden(.ColIndex(strColKey)) = Not blnShow
        If .ColWidth(.ColIndex(strColKey)) = 0 Then .ColWidth(.ColIndex(strColKey)) = 1000
    End With
    
End Function
Private Sub vsColSet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '修改后
    Dim strColKey As String, blnShow As Boolean
    With vsColSet
        Select Case Col
        Case .ColIndex("选择")
            blnShow = GetVsGridBoolColVal(vsColSet, Row, .ColIndex("选择"))
            Call SetVsGridCol(.TextMatrix(Row, .ColIndex("列名")), blnShow)
        Case Else
        End Select
    End With
End Sub
Private Sub vsColSet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsColSet
        Select Case Col
        Case .ColIndex("选择")
            'rowdata(i):1-固定,-1-不能选,0-可选
            If .RowData(Row) = 1 Then
                Cancel = True
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Sub vsColSet_LostFocus()
    vsColSet.Visible = False
End Sub

Private Function SendOtherPay() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:按处方发料
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-06 14:55:11
    '-----------------------------------------------------------------------------------------------------------
    With frm代发料
        .In_单据 = 0
        .In_单据IN = mArrFilter("单据")
        .In_发料部门id = Val(mArrFilter("发料部门ID"))
        .In_库存检查 = GetCheckPara()
        .In_允许未配料发料 = 1
        .In_权限 = mstrPrivs
        .mstr配料人 = gstrUserName
        Set .In_PlugIn = mobjPlugIn
        .Show 1, Me
    End With
    SendOtherPay = True
    Call mfrmFilter_zlRefreshCon(mArrFilter)
    
End Function
