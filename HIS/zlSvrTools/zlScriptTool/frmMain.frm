VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmMain 
   Caption         =   "部件收集工具"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":6852
   ScaleHeight     =   9240
   ScaleWidth      =   13575
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chk部件 
      Caption         =   "系统文件"
      Height          =   240
      Index           =   5
      Left            =   3420
      TabIndex        =   17
      Top             =   5715
      Value           =   1  'Checked
      Width           =   1080
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2310
      Index           =   4
      Left            =   7725
      ScaleHeight     =   2310
      ScaleWidth      =   2610
      TabIndex        =   15
      Top             =   6330
      Width           =   2610
      Begin VSFlex8Ctl.VSFlexGrid fgFile 
         Height          =   1515
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   16110
         _cx             =   28416
         _cy             =   2672
         Appearance      =   2
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14737600
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   33023
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483630
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   12
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMain.frx":6B94
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
         Ellipsis        =   0
         ExplorerBar     =   7
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
         WallPaperAlignment=   4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2310
      Index           =   3
      Left            =   4845
      ScaleHeight     =   2310
      ScaleWidth      =   2610
      TabIndex        =   13
      Top             =   6255
      Width           =   2610
      Begin VB.TextBox txtScript 
         BackColor       =   &H00C0FFC0&
         Height          =   2205
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   30
         Width           =   2490
      End
   End
   Begin VB.CheckBox chk部件 
      Caption         =   "三方部件"
      Height          =   240
      Index           =   4
      Left            =   2355
      TabIndex        =   12
      Top             =   5715
      Value           =   1  'Checked
      Width           =   1080
   End
   Begin VB.CheckBox chk部件 
      BackColor       =   &H00FFFFC0&
      Caption         =   "其他文件"
      Height          =   240
      Index           =   3
      Left            =   1245
      TabIndex        =   11
      Top             =   5715
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CheckBox chk部件 
      BackColor       =   &H00FFFFC0&
      Caption         =   "帮助文件"
      Height          =   240
      Index           =   2
      Left            =   3420
      TabIndex        =   10
      Top             =   5370
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CheckBox chk部件 
      BackColor       =   &H00FFFFC0&
      Caption         =   "应用部件"
      Height          =   240
      Index           =   1
      Left            =   2355
      TabIndex        =   9
      Top             =   5385
      Value           =   1  'Checked
      Width           =   1050
   End
   Begin VB.CheckBox chk部件 
      BackColor       =   &H00FFFFC0&
      Caption         =   "公共部件"
      Height          =   240
      Index           =   0
      Left            =   1260
      TabIndex        =   8
      Top             =   5385
      Value           =   1  'Checked
      Width           =   1050
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   2
      Left            =   1200
      ScaleHeight     =   330
      ScaleWidth      =   2970
      TabIndex        =   6
      Top             =   5010
      Width           =   2970
      Begin VB.ComboBox cboSystem 
         Height          =   300
         Left            =   15
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   15
         Width           =   2925
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   0
      Left            =   180
      ScaleHeight     =   3585
      ScaleWidth      =   11970
      TabIndex        =   2
      Top             =   525
      Width           =   11970
      Begin VB.PictureBox picPane 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   1695
         Index           =   1
         Left            =   1140
         ScaleHeight     =   1695
         ScaleWidth      =   10785
         TabIndex        =   4
         Top             =   1725
         Width           =   10785
         Begin VSFlex8Ctl.VSFlexGrid fgMain 
            Height          =   1515
            Left            =   90
            TabIndex        =   5
            Top             =   105
            Width           =   16110
            _cx             =   28416
            _cy             =   2672
            Appearance      =   2
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   14737600
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16761024
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorAlternate=   12648384
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483630
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   2
            SelectionMode   =   2
            GridLines       =   1
            GridLinesFixed  =   12
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmMain.frx":6CC0
            ScrollTrack     =   0   'False
            ScrollBars      =   3
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
            ExplorerBar     =   7
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
            WallPaperAlignment=   4
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   2775
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   11175
         _Version        =   589884
         _ExtentX        =   19711
         _ExtentY        =   4895
         _StockProps     =   64
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   9150
      TabIndex        =   1
      ToolTipText     =   "快捷键：F3"
      Top             =   30
      Width           =   1320
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8880
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   635
      SimpleText      =   $"frmMain.frx":6E21
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":6E3D
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21034
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
   Begin VB.Image imgBG_fg 
      Height          =   1845
      Index           =   1
      Left            =   5325
      Picture         =   "frmMain.frx":76D1
      Top             =   3795
      Visible         =   0   'False
      Width           =   8145
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMain.frx":96DE
      Left            =   525
      Top             =   45
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
      ScaleMode       =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////
'
'       模块：收集升级文件脚本工具
'       功能：收集、组织文件,创建脚本代码到脚本中。
'       编写：祝庆
'       日期：2010年11月15日
'
'///////////////////////////////////////////////////////////////////////////////


Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'打印模式
Enum gzlPrintModeS
    zlPrint = 1         '打印
    zlView = 2          '查看
    zlExcel = 3         '输出到Excel
End Enum

Private mobjFindKey             As CommandBarPopup      '查询
Private mstrFindKey             As String               '查询串
Private m_strCurTypeName        As String               '当前选中的方式
Private m_strCurFileName        As String               '当前选中的名称
Private m_strCurVision          As String               '当前选中的版本
Private m_strCurEditDate        As String               '当前选中的修改日期
Private m_strCurSysNum          As String               '当前选中的系统
Private m_strCurSetupPath       As String               '当前选中的安装路径
Private m_strCurSysOption       As String               '当前选中的系统参数
Private m_strCurFileExplanation As String               '当前选中的文件说明
Private m_strCurSellFile        As String               '当前选中的引用文件
Private m_blnCurReg             As Boolean              '当前选中的文件是否注册
Private m_blnCurUpData          As Boolean              '当前选中的文件是否强制覆盖

Private mzlPrintModeS           As gzlPrintModeS        '打印

Private m_lngCurRow             As Long
Private mstrPrivs               As String               '权限串
Private mcbrPopupBarItem        As CommandBar           '弹出窗口【项目】
Dim cbrPopupItem                As CommandBarControl    '弹出项
Dim mrsTemp      As ADODB.Recordset

'==============================================================================
'=功能： 控件初始化
'==============================================================================
Private Sub InitControl()
    On Error GoTo errH

    '加载TbcPage
    Call InitTbcPagePanel
    '菜单控制
    Call InitCommandBar
    '加载区域
    Call InitDockPannel

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 初始菜单工具栏
'==============================================================================
Private Sub InitCommandBar()
    Dim objMenu         As CommandBarPopup
    Dim objBar          As CommandBar
    Dim objExtendedBar  As CommandBar
    Dim objPopup        As CommandBarPopup
    Dim objControl      As CommandBarControl
    Dim cbrCustom       As CommandBarControlCustom
    
    On Error GoTo errH
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.Id = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "参数设置(&C)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    '------------------------------------------------------------------------------------------------------------------
    '编辑
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.Id = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "增加文件(&A)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_CopyNewItem, "插入文件(&I)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "修改文件(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除文件(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Gather, "收集文件(&S)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Generate, "生成脚本(&C)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Unpack, "复制脚本(&U)")
    
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.Id = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.Id = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & ParamInfo.产品名称)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.产品名称 & "主页(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.产品名称 & "论坛(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)
    '------------------------------------------------------------------------------------------------------------------
    '主菜单右侧的查找
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    
    If mstrFindKey = "" Then mstrFindKey = "名称"
    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    

    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.名称", , , "名称")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.类型", , , "类型")
    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = txtLocation.Hwnd
    cbrCustom.Flags = xtpFlagRightAlign
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "前一条")
    objControl.Flags = xtpFlagRightAlign
    objControl.Style = xtpButtonIcon
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "后一条")
    objControl.Flags = xtpFlagRightAlign
    objControl.Style = xtpButtonIcon
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "增加", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_CopyNewItem, "插入")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Gather, "收集", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Generate, "生成")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_FenFa, "分配")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Unpack, "复制")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
    '工具栏右侧的查找
    '------------------------------------------------------------------------------------------------------------------
'    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
'    cbrCustom.Handle = picPane(2).Hwnd
        '标准过滤栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("过滤栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, conMenu_Edit_System, "系统")
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10005, "")
    cbrCustom.Handle = picPane(2).Hwnd
    cbrCustom.Flags = xtpFlagLeftPopup
'
'    Set objControl = NewToolBar(objBar, xtpControlLabel, 10006, " ")
    
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, 10000, "", True)
    cbrCustom.Handle = chk部件(0).Hwnd
    cbrCustom.Flags = xtpFlagRightAlign
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10001, "")
    cbrCustom.Handle = chk部件(1).Hwnd
    cbrCustom.Flags = xtpFlagRightAlign
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10002, "")
    cbrCustom.Handle = chk部件(2).Hwnd
    cbrCustom.Flags = xtpFlagRightAlign
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10003, "")
    cbrCustom.Handle = chk部件(3).Hwnd
    cbrCustom.Flags = xtpFlagRightAlign
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10004, "")
    cbrCustom.Handle = chk部件(4).Hwnd
    cbrCustom.Flags = xtpFlagRightAlign
    Set cbrCustom = objBar.Controls.Add(xtpControlCustom, 10005, "")
    cbrCustom.Handle = chk部件(5).Hwnd
    cbrCustom.Flags = xtpFlagRightAlign
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '新增
        .Add FCONTROL, vbKeyI, conMenu_Edit_CopyNewItem     '插入
        .Add FCONTROL, vbKeyE, conMenu_Edit_Modify          '修改
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete       '删除
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save     '保存
        .Add 0, vbKeyF3, conMenu_View_Location              '定位
        .Add 0, vbKeyF4, conMenu_View_Option                '选择定位依据
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      '前一条
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '后一条
        
        .Add FCONTROL, vbKeyS, conMenu_Edit_Gather          '收集
        .Add FCONTROL, vbKeyC, conMenu_Edit_Generate          '收集
        .Add FCONTROL, vbKeyU, conMenu_Edit_Unpack          '收集
        .Add FCONTROL, vbKeyF, conMenu_Edit_FenFa            '分配
    End With
    '------------------------------------------------------------------------------------------------------------------
    '弹出菜单分类
    Set mcbrPopupBarItem = cbsMain.Add("弹出项目菜单", xtpBarPopup)
    mcbrPopupBarItem.ContextMenuPresent = False
    mcbrPopupBarItem.ShowTextBelowIcons = False
    mcbrPopupBarItem.EnableDocking xtpFlagStretched
    Set cbrPopupItem = NewToolBar(mcbrPopupBarItem, xtpControlButton, conMenu_Edit_NewItem, "增加文件(&A)")
    Set cbrPopupItem = NewToolBar(mcbrPopupBarItem, xtpControlButton, conMenu_Edit_CopyNewItem, "插入文件(&I)")
    Set cbrPopupItem = NewToolBar(mcbrPopupBarItem, xtpControlButton, conMenu_Edit_Modify, "修改文件(&M)")
    Set cbrPopupItem = NewToolBar(mcbrPopupBarItem, xtpControlButton, conMenu_Edit_Delete, "删除文件(&D)")

    Set cbrPopupItem = NewToolBar(mcbrPopupBarItem, xtpControlButton, conMenu_Edit_Gather, "收集文件(&S)", True)
    Set cbrPopupItem = NewToolBar(mcbrPopupBarItem, xtpControlButton, conMenu_Edit_Generate, "生成脚本(&C)")
    Set cbrPopupItem = NewToolBar(mcbrPopupBarItem, xtpControlButton, conMenu_Edit_Unpack, "复制脚本(&U)")
    
'    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "增加标准(&A)")
'    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_CopyNewItem, "插入标准(&I)")
'    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "修改标准(&M)")
'    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除标准(&D)")
'
'    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Gather, "收集文件(&S)")
'    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Generate, "生成脚本(&C)")
'    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Unpack, "打包上传(&U)")
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboSystem_Click()
    Call refData
End Sub

'==============================================================================
'=功能： 菜单功能控制
'==============================================================================
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnNewCancel        As Boolean
    On Error GoTo errH
    
    Select Case Control.Id
        Case conMenu_Edit_NewItem           '增加文件
            Call StandardAdd
        Case conMenu_Edit_CopyNewItem       '插入文件
            Call StandardCopyAdd
        Case conMenu_Edit_Modify            '修改文件
            Call StandardEdit
        Case conMenu_Edit_Delete            '删除文件
            Call StandardDel
        Case conMenu_View_Refresh           '刷新数据
            Call refData
        Case conMenu_File_Preview           '预览
            Call ItemPrint
        Case conMenu_File_Print             '打印
            mzlPrintModeS = zlPrint
            Call ItemPrint
        Case conMenu_File_Excel             '输出到&Excel
            mzlPrintModeS = zlExcel
            Call ItemPrint
        Case conMenu_File_Parameter         '参数设置
'            frm评分标准参数设置.Show 1
        Case conMenu_Edit_Gather           '收集部件
           
        Case conMenu_Edit_Generate         '生成脚本
            Call GenerateScript
        Case conMenu_Edit_Unpack           '复制脚本
            VB.Clipboard.Clear
            VB.Clipboard.SetText txtScript.Text
        Case conMenu_View_Forward           '上一条
            With fgMain
            If .Row > 1 Then
                .Row = .Row - 1
                .ShowCell .Row, .Col
            End If
        End With
        Case conMenu_View_Backward          '下一条
 
            With fgMain
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    .ShowCell .Row, .Col
                End If
            End With
        Case conMenu_View_Option
            mobjFindKey.Execute
        Case conMenu_View_LocationItem
            mstrFindKey = Control.Parameter
            mobjFindKey.Caption = mstrFindKey
            cbsMain.RecalcLayout
        Case conMenu_View_Location
            LocationObj txtLocation
        Case Else
            Call CommandBarExecutePublic(Control, Me, fgMain, "frmMain")
    End Select
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
        
    Select Case Control.Id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet              '打印设置
    
        Call zlPrintSet
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel               '打印数据,预览数据,输出到Excel
        
        If objPrnVsf Is Nothing Then Exit Function
        
        If Not SearchPrintData(objPrnVsf, frmPubResource.msfPrint) Then
            MsgBox "你打印的网络不存在数据，请重新检视！", vbInformation, ParamInfo.系统名称
            Exit Function
        End If
        
        '调用打印部件处理
        Set objPrint.Body = frmPubResource.msfPrint
        objPrint.Title.Text = strPrintTitle
        Set objAppRow = New zlTabAppRow
        Call objAppRow.Add("")
        Call objAppRow.Add("打印时间:" & Now())
        Call objPrint.BelowAppRows.Add(objAppRow)

        Select Case Control.Id
        Case conMenu_File_Print
            bytMode = zlPrintAsk(objPrint)
            If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
        Case conMenu_File_Preview
            zlPrintOrView1Grd objPrint, 2
        Case conMenu_File_Excel
            zlPrintOrView1Grd objPrint, 3
        End Select
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
    
    Case conMenu_Help_Help              '帮助主题
    
'        Call ShowHelp(App.ProductName, frmMain.Hwnd, frmMain.Name, Int((ParamInfo.系统号) / 100))
        
    Case conMenu_Help_Web_Home          'Web上的中联
        
        Call zlHomePage(frmMain.Hwnd)
        
    Case conMenu_Help_Web_Forum         'Web上的论坛
    
        Call zlWebForum(frmMain.Hwnd)
        
    Case conMenu_Help_Web_Mail          '发送反馈
        
        Call zlMailTo(frmMain.Hwnd)
            
    Case conMenu_Help_About             '关于
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_File_Exit              '退出
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function


'==============================================================================
'=功能： 菜单权限控制
'==============================================================================
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo errH
    
    With fgMain
        Select Case Control.Id
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel
                Control.Enabled = fgMain.Rows > 1
            Case conMenu_Edit_NewItem          '增加文件
                Control.Enabled = (tbcPage.Selected.Index = 0)
            Case conMenu_Edit_CopyNewItem      '插入文件
                Control.Enabled = (tbcPage.Selected.Index = 0)
            Case conMenu_Edit_Modify           '修改文件
    
                Control.Enabled = (fgMain.Rows > 0) And (tbcPage.Selected.Index = 0)
            Case conMenu_Edit_Delete           '删除文件
                Control.Enabled = (fgMain.SelectedRows > 0) And (tbcPage.Selected.Index = 0)
            Case conMenu_Edit_Gather           '收集部件
                Control.Enabled = (fgMain.Rows > 0) And (tbcPage.Selected.Index = 0)
            Case conMenu_Edit_Generate         '生成脚本
                Control.Enabled = (fgMain.Rows > 0) And (tbcPage.Selected.Index = 1)
            Case conMenu_Edit_Unpack           '复制脚本
                Control.Enabled = (fgMain.Rows > 0) And (tbcPage.Selected.Index = 1)
            Case conMenu_View_Forward
                Control.Enabled = (Control.Visible And fgMain.Row > 1)
            Case conMenu_View_Backward
                Control.Enabled = (Control.Visible And fgMain.Row + 1 < fgMain.Rows)
            Case conMenu_View_Refresh
                Control.Enabled = Control.Visible
            Case conMenu_View_LocationItem, conMenu_View_LocationItem, conMenu_View_LocationItem
                If InStr(Control.Caption, mstrFindKey) > 0 Then
                    Control.Checked = True
                Else
                    Control.Checked = False
                End If
            Case conMenu_Edit_FenFa
                Control.Enabled = (fgFile.Rows > 0) And (tbcPage.Selected.Index = 2)
            Case Else
                Call CommandBarUpdatePublic(Control, Me)
        End Select
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub


Private Sub chk部件_Click(Index As Integer)
    Dim strTemp As String
    On Error GoTo errH
    If chk部件(0).Value Then
        strTemp = "0,"
    End If
    
    If chk部件(1).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "1,"
        Else
            strTemp = strTemp & "1,"
        End If
    End If
    
    If chk部件(2).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "2,"
        Else
            strTemp = strTemp & "2,"
        End If
    End If
    
    If chk部件(3).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "3,"
        Else
            strTemp = strTemp & "3,"
        End If
    End If
    
    If chk部件(4).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "4,"
        Else
            strTemp = strTemp & "4,"
        End If
    End If
    
    If chk部件(5).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "5"
        Else
            strTemp = strTemp & "5"
        End If
    End If
    
    
    If Len(strTemp) > 0 Then
        If Right(strTemp, 1) = "," Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
        Call DataLoad(strTemp)
    Else
        Call DataLoad("Clear")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 刷新数据
'==============================================================================
Private Sub refData()
    Dim strTemp As String
    On Error GoTo errH
    If chk部件(0).Value Then
        strTemp = "0,"
    End If
    
    If chk部件(1).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "1,"
        Else
            strTemp = strTemp & "1,"
        End If
    End If
    
    If chk部件(2).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "2,"
        Else
            strTemp = strTemp & "2,"
        End If
    End If
    
    If chk部件(3).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "3,"
        Else
            strTemp = strTemp & "3,"
        End If
    End If
    
    If chk部件(4).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "4,"
        Else
            strTemp = strTemp & "4,"
        End If
    End If
    
    If chk部件(5).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "5"
        Else
            strTemp = strTemp & "5"
        End If
    End If
    
    If Len(strTemp) > 0 Then
        If Right(strTemp, 1) = "," Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
        Call DataLoad(strTemp)
    Else
        Call DataLoad("Clear")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub fgMain_DblClick()
    Call StandardEdit
End Sub

'==============================================================================
'=功能： 窗口初始化时窗口控位位置控制
'==============================================================================
Private Sub Form_Activate()
    On Error GoTo errH
    Call Form_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 窗口初始化
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo errH
    Call InitCommonControls
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 窗口初始化
'==============================================================================
Private Sub Form_Load()
  On Error GoTo errH
    
    KeyPreview = True
    mstrPrivs = "增删改"
    '控件初始化
    Call InitControl
    m_lngCurRow = -1
    
    '填充Combo
    Call InitComBo
    
    '填充数据
'    Call DataLoad
    
    ChDir App.Path
'    SkinFramework1.LoadSkin "Office2007.cjstyles", ""
'    SkinFramework1.ApplyWindow Me.Hwnd
'    SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
  
    '恢复界面位置
    
    If KeyPreview Then
        frmMain.Move Screen.Width / 2 - frmMain.Width / 2, Screen.Height / 2 - frmMain.Height / 2, 1024 * Screen.TwipsPerPixelX, 768 * Screen.TwipsPerPixelY
        Call meRestoreWinState
    End If
    
    Call SetMenu
'    refListView
    Exit Sub
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 窗口位置变化
'==============================================================================
Private Sub Form_Resize()
    On Error GoTo errH

    Call SetPaneRange(dkpMain, 1, 100, 60, 450, Me.ScaleHeight)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 窗口退出
'==============================================================================
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errH
    meSaveWinState
  
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 显示记录数信息
'==============================================================================
Private Sub SetMenu()
    On Error GoTo errH
    stbThis.Panels(2).Text = "列表中共显示有" & fgMain.Rows - 1 & "行数据。"
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub refListView() '刷新更新表
    Dim strSQL As String
    Dim rsFileUp As ADODB.Recordset
    Dim m_item As MSComctlLib.ListItem
    
    strSQL = "select id,filename,filesize from TestFileUpdate"
    Set rsFileUp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    rsFileUp.MoveFirst
    Do Until rsFileUp.EOF
            rsFileUp.MoveNext
    Loop
End Sub

'==============================================================================
'=功能： 界面分割
'==============================================================================
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo errH
    Select Case Item.Id
        Case 1
            Item.Handle = picPane(0).Hwnd
        Case 2
            Item.Handle = picPane(1).Hwnd
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 初始区域划分
'==============================================================================
Private Sub InitDockPannel()
    Dim objPane As Pane
    
    On Error GoTo errH
    
    Set objPane = dkpMain.CreatePane(1, 100, 100, DockTopOf, Nothing)
    objPane.Title = "收集"
    objPane.Options = PaneNoCaption
'    Set objPane = dkpMain.CreatePane(2, 200, 100, DockBottomOf, Nothing)
'    objPane.Title = "控制"
'    objPane.Options = PaneNoCaption
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        tbcPage.Move 0, 0, picPane(Index).Width, picPane(Index).Height
'        picPane(2).Move tbcPage.Width - picPane(2).Width, -15, picPane(2).Width, picPane(2).Height
    Case 1
        fgMain.Move 15, 15, picPane(1).Width - 15 * 2, picPane(1).Height - picPane(1).Top - 15 * 2
    Case 3
        txtScript.Move 15, 15, picPane(3).Width - 15 * 2, picPane(3).Height - picPane(3).Top - 15 * 2
    Case 4
        fgFile.Move 15, 15, picPane(4).Width - 15 * 2, picPane(4).Height - picPane(4).Top - 15 * 2
    End Select
End Sub

'==============================================================================
'=功能： 初始Tbc分页
'==============================================================================
Private Sub InitTbcPagePanel()
    Call TabControlInit(tbcPage)
    With tbcPage
        .PaintManager.BoldSelected = True
        
        Call .InsertItem(0, "文件收集", picPane(1).Hwnd, 3)
        Call .InsertItem(1, "脚本生成", picPane(3).Hwnd, 2)
        Call .InsertItem(2, "文件分配", picPane(4).Hwnd, 3)

        .Item(0).Selected = True
        .Item(2).Visible = False
    End With
End Sub

'==============================================================================
'=功能： 网格fgMain单击后刷新状态信息
'==============================================================================
Private Sub fgMain_Click()
    On Error GoTo errH
    fgMain_SelChange
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 右键菜单 fgMain
'==============================================================================
Private Sub fgMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errH
    Select Case Button
        Case 2          '弹出菜单处理
            Call SendLMouseButton(fgMain.Hwnd, X, Y)
            mcbrPopupBarItem.ShowPopup
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网格行列变化时更新基本信息
'==============================================================================
Private Sub fgMain_RowColChange()
    On Error GoTo errH
    Call fgMain_SelChange
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网格选择行列变化时更新基本信息
'==============================================================================
Private Sub fgMain_SelChange()
    Dim lngID       As Long
    On Error GoTo errH
    
'    fgMain.WallPaper = imgBG_fg(1).Picture
    m_lngCurRow = fgMain.Row
    If m_lngCurRow = 0 Then Exit Sub
    m_strCurTypeName = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 1)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 1))   '获取ID
    m_strCurFileName = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 2)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 2))     '获取ID
    m_strCurVision = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 3)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 3))
    m_strCurEditDate = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 4)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 4))
    m_strCurSysNum = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 5)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 5))
    m_strCurSellFile = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 6)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 6))
    m_strCurSetupPath = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 7)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 7))
    m_strCurSysOption = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 10)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 10))
    m_strCurFileExplanation = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 11)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 11)) '文件说明
    m_blnCurReg = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 12)) = 0, False, fgMain.Cell(flexcpText, m_lngCurRow, 12)) '自动注册
    m_blnCurUpData = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 13)) = 0, False, fgMain.Cell(flexcpText, m_lngCurRow, 13)) '强制覆盖
    
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'==============================================================================
'=功能： 填充系统 ComBo
'==============================================================================
Private Sub InitComBo()
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim lngDefaultNum As Long
    Dim str编号       As String
    
    On Error GoTo errH
    With cboSystem
        .Clear
        strSQL = "select 编号,名称,共享号 from zlSystems"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rs.BOF = False Then
            rs.MoveFirst
            .AddItem "[0]所有系统"
            .ItemData(.NewIndex) = 0
            Do While Not rs.EOF
                str编号 = rs("编号").Value \ 100
                .AddItem "[" & str编号 & "]" & rs("名称").Value
                .ItemData(.NewIndex) = str编号
                If NVL(rs("共享号").Value, 0) = 0 Then
                    lngDefaultNum = .ListCount - 1
                End If
                rs.MoveNext
            Loop
        End If
        .ListIndex = 0 'lngDefaultNum
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 装入对应方案的评分标准
'==============================================================================
Public Sub DataLoad(Optional ByVal strFilter As String)

    Dim i            As Long
    Dim strSQL       As String
    Dim strSystemNum As String
    Dim strTypeID()  As String
    Dim strTemp      As String
    On Error GoTo errH
    
    With fgMain
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        .Cols = 14
'        Exit Sub
        .Cell(flexcpText, 0, 0) = "序号"
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 1) = "文件类型"
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 2) = "文件名"
        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 3) = "版本号"
        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 4) = "修改日期"
        .Cell(flexcpAlignment, 0, 4) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 5) = "所属系统"
        .Cell(flexcpAlignment, 0, 5) = flexAlignCenterCenter
        .ColWidth(5) = 1800
        .Cell(flexcpText, 0, 6) = "业务部件"
        .Cell(flexcpAlignment, 0, 6) = flexAlignCenterCenter
        .ColWidth(6) = 4800
        
        .Cell(flexcpText, 0, 7) = "安装路径"
        .Cell(flexcpAlignment, 0, 7) = flexAlignCenterCenter
        .ColWidth(7) = 0
        
        .Cell(flexcpText, 0, 8) = "类型ID"
        .Cell(flexcpAlignment, 0, 8) = flexAlignCenterCenter
        .ColWidth(8) = 0
        
        .Cell(flexcpText, 0, 9) = "安装路径"
        .Cell(flexcpAlignment, 0, 9) = flexAlignCenterCenter
        .ColWidth(9) = 2000
         
        .Cell(flexcpText, 0, 10) = "系统参数"
        .Cell(flexcpAlignment, 0, 10) = flexAlignCenterCenter
        .ColWidth(10) = 0
        .Cell(flexcpText, 0, 11) = "文件说明"
        .Cell(flexcpAlignment, 0, 11) = flexAlignCenterCenter
        .ColWidth(11) = 1000
        
        .Cell(flexcpText, 0, 12) = "自动注册"
        .Cell(flexcpAlignment, 0, 12) = flexAlignCenterCenter
        .ColWidth(12) = 0
        
        .Cell(flexcpText, 0, 13) = "强制覆盖"
        .Cell(flexcpAlignment, 0, 13) = flexAlignCenterCenter
        .ColWidth(13) = 0
        
        If CheckTable = False Then
            Exit Sub
        End If
        
        If Len(strFilter) <> 0 Then
            If strFilter = "Clear" Then
                Exit Sub
            Else
                strSystemNum = cboSystem.ItemData(cboSystem.ListIndex)
                If strSystemNum = "" Then strSystemNum = "1"
                
                If strSystemNum = "0" Then
                     strSQL = "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                             "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册,a.强制覆盖" & vbNewLine & _
                             "From zlFilesUpgrade A" & vbNewLine & _
                             "Where a.文件类型 In (" & strFilter & ") order by lpad(a.序号,5,'0')"
                              Set mrsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
                              GoTo zt
                End If
                
                If InStrRev(strFilter, "0") > 0 Then
                   strTypeID = Split(strFilter, ",")
                   For i = 0 To UBound(strTypeID)
                        If strTemp = "" Then
                            strTemp = strTypeID(i)
                        Else
                            strTemp = strTemp & "," & strTypeID(i)
                        End If
                   Next
                    strSQL = "Select B.序号,B.类型ID,B.文件类型,B.文件名,B.版本号,B.修改日期,B.所属系统,B.业务部件,B.安装路径,B.文件说明,B.自动注册,B.强制覆盖 From ( " & vbNewLine & _
                                "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                                "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册,a.强制覆盖" & vbNewLine & _
                                "From zlFilesUpgrade A" & vbNewLine & _
                                "Where a.文件类型 In (" & strTemp & ") And (Instr(a.所属系统, ','|| [1] || ',' ) > 0 or a.所属系统 is null )" & vbNewLine & _
                                "Union" & vbNewLine & _
                                "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                                "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册,a.强制覆盖" & vbNewLine & _
                                "From zlFilesUpgrade A" & vbNewLine & _
                                "Where a.文件类型 =0" & vbNewLine & _
                                ") B Order by lpad(B.序号,5,'0')"
                        
                    Set mrsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSystemNum)
                Else
                    strSQL = "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                             "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册,a.强制覆盖" & vbNewLine & _
                             "From zlFilesUpgrade A" & vbNewLine & _
                             "Where a.文件类型 In (" & strFilter & ") And (Instr(a.所属系统, ',' ||  [1] || ',' ) > 0 or a.所属系统 is null ) order by lpad(a.序号,5,'0')"

                    
                    Set mrsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSystemNum)
                End If
            End If
        Else
            strSystemNum = cboSystem.ItemData(cboSystem.ListIndex)
            If strSystemNum = "" Then strSystemNum = "100"
    
            strSQL = "Select B.序号,B.类型ID,B.文件类型,B.文件名,B.版本号,B.修改日期,B.所属系统,B.业务部件,B.安装路径,B.文件说明,B.强制覆盖 From ( " & vbNewLine & _
                        "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                         "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册,a.强制覆盖" & vbNewLine & _
                         "From zlFilesUpgrade A" & vbNewLine & _
                         "Where a.文件类型 In (1, 2, 3,4) And (Instr(a.所属系统,  ',' ||  [1] || ',') > 0 or a.所属系统 is null )" & vbNewLine & _
                         "Union" & vbNewLine & _
                         "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                         "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册,a.强制覆盖" & vbNewLine & _
                         "From zlFilesUpgrade A" & vbNewLine & _
                         "Where a.文件类型 =0" & vbNewLine & _
                         ") B Order by lpad(B.序号,5,'0')"
        
            Set mrsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSystemNum)
        End If
zt:
'    .AllowSelection = False '对齐
'    .Editable = flexEDKbdMouse
'    .AllowUserResizing = flexResizeBoth
'    .AllowUserFreezing = flexFreezeBoth
'    .BackColorFrozen = 14737632
'    .GridLines = flexGridFlatVert
        .ExtendLastCol = True
'    .ScrollTips = True
    
        .FocusRect = flexFocusSolid
        '数据填入
        .Rows = mrsTemp.RecordCount + 1
    
        i = 1
        Do Until mrsTemp.EOF
            .Cell(flexcpText, i, 0) = NVL(mrsTemp.Fields("序号"), 0) 'mrsTemp.AbsolutePosition
            .Cell(flexcpAlignment, i, 0) = flexAlignLeftCenter
            
            
            .Cell(flexcpText, i, 1) = NVL(mrsTemp.Fields("文件类型"))
            .Cell(flexcpAlignment, i, 1) = flexAlignCenterCenter
'            If NVL(mrsTemp.Fields("文件类型")) = "应用部件" Then
'                .Cell(flexcpBackColor, i, 1) = &H80C0FF   '&H8080FF
'            End If
            .Cell(flexcpText, i, 2) = NVL(mrsTemp.Fields("文件名"))
            .Cell(flexcpAlignment, i, 2) = flexAlignLeftCenter
            
            strTemp = NVL(mrsTemp.Fields("版本号"))
            strTemp = GetFileVision(strTemp)
            
            .Cell(flexcpText, i, 3) = strTemp
            .Cell(flexcpAlignment, i, 3) = flexAlignCenterCenter
            
            If NVL(mrsTemp.Fields("修改日期")) <> "" Then
                strTemp = Format(NVL(mrsTemp.Fields("修改日期")), "yyyy-mm-dd hh:mm:ss")
            Else
                strTemp = ""
            End If
            
            .Cell(flexcpText, i, 4) = strTemp
            .Cell(flexcpAlignment, i, 4) = flexAlignCenterCenter
            
            strTemp = NVL(mrsTemp.Fields("所属系统"))
            If strTemp Like ",*," Then
                strTemp = Right(strTemp, Len(strTemp) - 1)
                strTemp = Left(strTemp, Len(strTemp) - 1)
            End If
            If strTemp = "" Then
                strTemp = "所有系统"
            End If
            
            .Cell(flexcpText, i, 5) = strTemp
            .Cell(flexcpAlignment, i, 5) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 6) = NVL(mrsTemp.Fields("业务部件"))
            .Cell(flexcpAlignment, i, 6) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 7) = NVL(mrsTemp.Fields("安装路径"))
            .Cell(flexcpAlignment, i, 7) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 8) = NVL(mrsTemp.Fields("类型ID"))
            .Cell(flexcpAlignment, i, 8) = flexAlignLeftTop
            
            .Cell(flexcpText, i, 9) = NVL(mrsTemp.Fields("安装路径"))
            .Cell(flexcpAlignment, i, 9) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 10) = NVL(mrsTemp.Fields("所属系统")) 'NVL(mrsTemp.Fields("系统参数"))
            .Cell(flexcpAlignment, i, 10) = flexAlignCenterCenter
            
            .Cell(flexcpText, i, 11) = NVL(mrsTemp.Fields("文件说明"), "")
            .Cell(flexcpAlignment, i, 11) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 12) = NVL(mrsTemp.Fields("自动注册"), "")
            .Cell(flexcpAlignment, i, 12) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 13) = NVL(mrsTemp.Fields("强制覆盖"), 0)
            .Cell(flexcpAlignment, i, 13) = flexAlignLeftCenter
            mrsTemp.MoveNext
            
            i = i + 1
        Loop

        '自动换行
        .WordWrap = True
        '合并单元格
        .MergeCells = 0
        .MergeCol(.ColIndex("文件类型")) = True
        .MergeCol(.ColIndex("文件名")) = True
        '隐藏单元格
        .ColWidth(.ColIndex("类型ID")) = 0
        
        '行高设置
        .RowHeightMin = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize .ColIndex("业务部件")
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        
         Call SetMenu
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub meRestoreWinState()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", Me.Left)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", Me.Top)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", Me.Width)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", Me.Height)
End Sub

Private Sub meSaveWinState()
    If Me.WindowState <> vbMinimized Then
      SaveSetting App.Title, "Settings", "MainLeft", Me.Left
      SaveSetting App.Title, "Settings", "MainTop", Me.Top
      SaveSetting App.Title, "Settings", "MainWidth", Me.Width
      SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

'==============================================================================
'=功能： 检查表是否是新表或者表是否存在
'==============================================================================
Private Function CheckTable() As Boolean
    On Error GoTo errH
    Dim mrsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim blnUse As Boolean
    strSQL = "select * from zlFilesUpgrade where rownum =1"
    Set mrsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If mrsTemp.RecordCount >= 0 Then
        For i = 1 To mrsTemp.Fields.Count
            If mrsTemp.Fields.Item(i - 1).Name = "所属系统" Then
                blnUse = True
                Exit For
            End If
        Next
        
        If blnUse Then
            CheckTable = True
        Else
            MsgBox "在zlFilesUpgrade表中,没有找到相应的字段!" & vbCrLf & "请检查表结构是否为最新!", vbInformation
            CheckTable = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=新增文件
'==============================================================================
Private Sub StandardAdd()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    On Error GoTo errH
    If cboSystem.Text = "" Then
        strSysNum = 1
    Else
        strSysNum = cboSystem.ItemData(cboSystem.ListIndex)
    End If
    
    f.ShowForm "新增", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, "0"
    If f.Moded Then
        Call refData
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=插入文件
'==============================================================================
Private Sub StandardCopyAdd()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    Dim str序号 As String
    On Error GoTo errH
    If cboSystem.Text = "" Then
        strSysNum = 1
    Else
        strSysNum = cboSystem.ItemData(cboSystem.ListIndex)
    End If
    
    If fgMain.Row > 0 Then
       str序号 = fgMain.TextMatrix(fgMain.Row, fgMain.ColIndex("序号"))
    End If
    
    f.ShowForm "新增", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, str序号
    If f.Moded Then
        Call refData
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=修改文件
'==============================================================================
Private Sub StandardEdit()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    On Error GoTo errH
    If cboSystem.Text = "" Then
        strSysNum = 100
    Else
        strSysNum = cboSystem.ItemData(cboSystem.ListIndex)
    End If
    
    f.ShowForm "修改", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, "0"
    If f.Moded Then
        Call refData
        Dim lngRow As Long
        lngRow = fgMain.FindRow(CStr(m_strCurFileName), , 2)
        If lngRow <> -1 Then
              fgMain.Select lngRow, 2
              fgMain.ShowCell lngRow, 2
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=删除文件
'==============================================================================
Private Sub StandardDel()
    Dim i         As Long
    Dim strName   As String
    Dim lngCurRow As Long
    Dim rs        As ADODB.Recordset
    Dim strSQL    As String
    Dim strSys    As String
    Dim strSysNum As String
    Dim lngRow    As Long
    On Error GoTo errH
    
    If fgMain.SelectedRows = 0 Then Exit Sub
    
    If fgMain.SelectedRows = 1 Then
        If MsgBox("你确认要删除[" & Right(cboSystem.Text, Len(cboSystem.Text) - InStrRev(cboSystem.Text, "]", -1)) & "]" & vbCrLf & "的部件" & m_strCurFileName & "吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox("你确认要删除选择的" & fgMain.SelectedRows & "个部件吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
'    gcnOracle.BeginTrans
    
    
    lngRow = fgMain.FindRow(CStr(m_strCurFileName), , 2)
    
    For i = 0 To fgMain.SelectedRows
        If fgMain.SelectedRow(i) Then
            lngCurRow = fgMain.SelectedRow(i)
            If lngCurRow <> -1 Then
                strName = IIf(Len(fgMain.Cell(flexcpText, lngCurRow, 2)) = 0, 0, fgMain.Cell(flexcpText, lngCurRow, 2))
                strName = UCase(strName)
                
'                strSQL = "select 所属系统 from zlFilesUpgrade where upper(文件名) = upper([1])"
'                Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strName)
'                If rs.RecordCount = 1 Then
'                    strSys = rs!所属系统
'                    strSysNum = ";" & cboSystem.ItemData(cboSystem.ListIndex)
'                    If strSys = strSysNum Then GoTo zldell
'                    strSys = Replace(strSys, strSysNum, "")
'                    gstrSql = "update Zlfilesupgrade set 所属系统='" & strSys & "' where upper(文件名)=upper('" & strName & "')"
'                    gcnOracle.Execute gstrSql
'                Else
'zldell:
                gstrSql = "delete zlFilesUpgrade where upper(文件名)= upper('" & strName & "')"
                gcnOracle.Execute gstrSql
'                End If
            End If

        End If
    Next
    
'    gcnOracle.CommitTrans
    
    ''Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Call refData
    Call SetMenu
    
    
    If lngRow <> -1 Then
        If lngRow >= 2 And fgMain.Rows > 2 Then
          fgMain.Select lngRow - 1, 2
          fgMain.ShowCell lngRow - 1, 2
        End If
    End If
    Exit Sub
errH:
'    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 打印 ItemPrint
'==============================================================================
Private Sub ItemPrint()
    On Error GoTo errH
    subPrint (mzlPrintModeS)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能:将数据表进行打印,预览和输出到EXCEL
'=参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'==============================================================================
Private Sub subPrint(bytMode As Byte)
    Dim objPrint            As New zlPrint1Grd
    Dim objAppRow           As zlTabAppRow
    Dim bytR                As Byte
    Dim rs                  As ADODB.Recordset
    
    On Error GoTo errH
    
    Set objPrint.Body = fgMain
    objPrint.Title.Text = Right(cboSystem.Text, Len(cboSystem.Text) - InStrRev(cboSystem.Text, "]"))
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    Set objAppRow = New zlTabAppRow
  
    objAppRow.Add "系统编号:" & cboSystem.ItemData(cboSystem.ListIndex)
    
    objPrint.UnderAppRows.Add objAppRow
 
    
    Set objAppRow = New zlTabAppRow
    objAppRow.Add "打印人：" & gstrUserName
    objAppRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objAppRow
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objPrint)
        If bytR <> 0 Then zlPrintOrView1Grd objPrint, bytR
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
    Case 0
        
    Case 1
        
    Case 2
        
    End Select
End Sub

'==============================================================================
'=定位得到焦点选中
'==============================================================================
Private Sub txtLocation_GotFocus()
    On Error GoTo errH
    Call zlControl.TxtSelAll(txtLocation)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=快速定位
'==============================================================================
Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long

    On Error GoTo errH
    
    lngRow = 0
    If txtLocation.Locked Then Exit Sub
    If mstrFindKey = "名称" Then mstrFindKey = "文件名称"
    If KeyAscii = vbKeyReturn Then
        '读取大于当前行的记录数据
        For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
            If InStr(UCase(fgMain.TextMatrix(lngLoop, 2)), UCase(txtLocation.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '读取小于当前行的记录数据
        If lngRow = 0 Then
            For lngLoop = 0 To fgMain.Row
                If InStr(UCase(fgMain.TextMatrix(lngLoop, 2)), UCase(txtLocation.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If fgMain.Rows > 1 And lngRow >= 1 Then
            fgMain.Row = lngRow
            fgMain.ShowCell lngRow, 2
        End If
        
        
        Call LocationObj(txtLocation)
    End If
    If mstrFindKey = "文件名称" Then mstrFindKey = "名称"

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    mstrFindKey = "名称"
    
End Sub


Private Sub txtScript_KeyDown(KeyCode As Integer, Shift As Integer)
 If ((Shift And vbCtrlMask) > 0) And (KeyCode = vbKeyQ) Then  '全选
      txtScript.SelText = txtScript.Text
      Exit Sub
  End If
  
  If ((Shift And vbCtrlMask) > 0) And (KeyCode = vbKeyC) Then  '复制
    
      Exit Sub
  End If
End Sub

Private Sub txtScript_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    On Error GoTo errH
'    If Button = 2 Then mcbrPopupBarItem.ShowPopup
'    Exit Sub
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
End Sub

'==============================================================================
'=功能:收集当前选择的系统部件生成SQL语句
'==============================================================================
Private Sub GenerateScript()
    On Error GoTo errH
    Dim rs As ADODB.Recordset
    Dim strSelectSQL As String '查询用的SQL
    Dim strInert     As String '生成主语句
    Dim strSQL As String
    Dim strDeSQL As String
    Dim i As Integer
    Dim strSystemNum As String
    Dim strFilter As String
    
    strSystemNum = cboSystem.ItemData(cboSystem.ListIndex)
    txtScript.Text = ""
    
    strDeSQL = "delete from zlfilesupgrade;" & vbCrLf '& "commit;" & vbCrLf
    
    strSQL = "Insert Into ZLTOOLS.ZLFILESUPGRADE(序号,文件类型,文件名,版本号,修改日期,所属系统,业务部件,安装路径,文件说明,自动注册,强制覆盖)"
    strFilter = GetFileType
    If strFilter = "" Then Exit Sub
       
    If strSystemNum = "0" Then
        '生成所有的脚本
        strSelectSQL = "select * from ZLFILESUPGRADE Where 文件类型 In (" & strFilter & ")"
        Set rs = zlDatabase.OpenSQLRecord(strSelectSQL, Me.Caption)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do Until rs.EOF
               
                '最后一条语句
                If rs.AbsolutePosition = rs.RecordCount Then
                    strInert = strInert & vbCrLf & "Select " & rs.AbsolutePosition & "," & rs!文件类型 & ",'" & rs!文件名 & "','" & rs!版本号 & "',to_date('" & rs!修改日期 & "','yyyy-mm-dd hh24:mi:ss'),'" & NVL(rs!所属系统, "") & "','" & NVL(rs!业务部件, "") & "','" & NVL(rs!安装路径, "") & "','" & NVL(rs!文件说明) & "'," & NVL(rs!自动注册, 0) & "," & NVL(rs!强制覆盖, 0) & " From Dual"
                Else
                    strInert = strInert & vbCrLf & "Select " & rs.AbsolutePosition & "," & rs!文件类型 & ",'" & rs!文件名 & "','" & rs!版本号 & "',to_date('" & rs!修改日期 & "','yyyy-mm-dd hh24:mi:ss'),'" & NVL(rs!所属系统, "") & "','" & NVL(rs!业务部件, "") & "','" & NVL(rs!安装路径, "") & "','" & NVL(rs!文件说明) & "'," & NVL(rs!自动注册, 0) & "," & NVL(rs!强制覆盖, 0) & " From Dual Union All"
                End If
                rs.MoveNext
            Loop
        End If
    Else
       '生成系统号的脚本

       strSelectSQL = "Select * From zlFilesUpgrade A" & vbNewLine & _
                 "Where a.文件类型 In (" & strFilter & ") and (Instr(a.所属系统, ',' ||  [1] || ',') > 0 or a.所属系统 is null)"
                             
        Set rs = zlDatabase.OpenSQLRecord(strSelectSQL, Me.Caption, strSystemNum)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do Until rs.EOF
               
                '最后一条语句
                If rs.AbsolutePosition = rs.RecordCount Then
                    strInert = strInert & vbCrLf & "Select " & rs.AbsolutePosition & "," & rs!文件类型 & ",'" & rs!文件名 & "','" & rs!版本号 & "',to_date('" & rs!修改日期 & "','yyyy-mm-dd hh24:mi:ss'),'" & NVL(rs!所属系统, "") & "','" & NVL(rs!业务部件, "") & "','" & NVL(rs!安装路径, "") & "','" & NVL(rs!文件说明) & "'," & NVL(rs!自动注册, 0) & "," & NVL(rs!强制覆盖, 0) & " From Dual"
                Else
                    strInert = strInert & vbCrLf & "Select " & rs.AbsolutePosition & "," & rs!文件类型 & ",'" & rs!文件名 & "','" & rs!版本号 & "',to_date('" & rs!修改日期 & "','yyyy-mm-dd hh24:mi:ss'),'" & NVL(rs!所属系统, "") & "','" & NVL(rs!业务部件, "") & "','" & NVL(rs!安装路径, "") & "','" & NVL(rs!文件说明) & "'," & NVL(rs!自动注册, 0) & "," & NVL(rs!强制覆盖, 0) & " From Dual Union All"
                End If
                rs.MoveNext
            Loop
        End If
    End If
    
    If strInert <> "" Then
        txtScript.Text = strDeSQL & strSQL & strInert & ";"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

''''==============================================================================
''''=功能:获得文件是否需要注册
''''==============================================================================
'''Private Function GetRegedit(ByVal strOption As String) As String
'''    Dim i As Integer
'''    Dim strTemp As String
'''    On Error Resume Next
'''    If strOption = "" Then Exit Function
'''    i = InStrRev(strOption, "Z")
'''    If i > 0 Then
'''        strTemp = Right(Left(strOption, i + 1), 1)
'''        If strTemp = "1" Then
'''            GetRegedit = "是"
'''        Else
'''            GetRegedit = ""
'''        End If
'''    End If
'''
'''End Function

Private Function GetFileType() As String
    Dim strTemp As String
    On Error GoTo errH
    If chk部件(0).Value Then
        strTemp = "0,"
    End If
    
    If chk部件(1).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "1,"
        Else
            strTemp = strTemp & "1,"
        End If
    End If
    
    If chk部件(2).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "2,"
        Else
            strTemp = strTemp & "2,"
        End If
    End If
    
    If chk部件(3).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "3,"
        Else
            strTemp = strTemp & "3,"
        End If
    End If
    
    If chk部件(4).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "4,"
        Else
            strTemp = strTemp & "4,"
        End If
    End If
    
    If chk部件(5).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "5"
        Else
            strTemp = strTemp & "5"
        End If
    End If
    
    If Len(strTemp) > 0 Then
        If Right(strTemp, 1) = "," Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
       GetFileType = strTemp
    Else
       GetFileType = ""
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
