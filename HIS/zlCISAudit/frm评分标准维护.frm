VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm评分标准维护 
   Caption         =   "评分标准维护"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   270
   ClientWidth     =   11115
   Icon            =   "frm评分标准维护.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8580
      TabIndex        =   18
      ToolTipText     =   "快捷键：F3"
      Top             =   210
      Width           =   1320
   End
   Begin VB.PictureBox picRightUp 
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   4590
      ScaleHeight     =   2715
      ScaleWidth      =   5010
      TabIndex        =   16
      Top             =   870
      Width           =   5010
      Begin VSFlex8Ctl.VSFlexGrid fgMain 
         Height          =   5475
         Left            =   45
         TabIndex        =   17
         Top             =   165
         Width           =   5235
         _cx             =   9234
         _cy             =   9657
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483643
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm评分标准维护.frx":1272
         ScrollTrack     =   -1  'True
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
         ExplorerBar     =   0
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7680
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   635
      SimpleText      =   $"frm评分标准维护.frx":1389
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm评分标准维护.frx":13D0
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16695
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
   Begin VB.PictureBox pic项目信息_S 
      BackColor       =   &H00FAFAFA&
      Height          =   1935
      Left            =   4605
      Picture         =   "frm评分标准维护.frx":1C64
      ScaleHeight     =   1875
      ScaleWidth      =   5040
      TabIndex        =   13
      Top             =   3765
      Width           =   5100
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "无内容"
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   165
         TabIndex        =   15
         Top             =   420
         Width           =   6360
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "项目内容"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0F0F0&
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   90
         Width           =   1095
      End
   End
   Begin VB.PictureBox picLeft_S 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5610
      Left            =   135
      ScaleHeight     =   5610
      ScaleWidth      =   3255
      TabIndex        =   1
      Top             =   690
      Width           =   3255
      Begin VB.PictureBox pic方案信息 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1695
         ScaleWidth      =   2790
         TabIndex        =   5
         Top             =   1965
         Width           =   2790
         Begin VB.PictureBox picFAXX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2415
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   6
            Top             =   75
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "方案信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   12
            Top             =   90
            Width           =   1095
         End
         Begin VB.Label lbl方案名称 
            BackStyle       =   0  'Transparent
            Caption         =   "方案名称"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   11
            Top             =   450
            Width           =   2580
         End
         Begin VB.Label lbl上值 
            BackStyle       =   0  'Transparent
            Caption         =   "上值:"
            Height          =   195
            Left            =   225
            TabIndex        =   10
            Top             =   1146
            Width           =   2580
         End
         Begin VB.Label lbl下值 
            BackStyle       =   0  'Transparent
            Caption         =   "下值:"
            Height          =   195
            Left            =   225
            TabIndex        =   9
            Top             =   1380
            Width           =   2580
         End
         Begin VB.Label lbl总分 
            BackStyle       =   0  'Transparent
            Caption         =   "总分:"
            Height          =   195
            Left            =   225
            TabIndex        =   8
            Top             =   914
            Width           =   2580
         End
         Begin VB.Label lbl分制 
            BackStyle       =   0  'Transparent
            Caption         =   "分制:"
            Height          =   195
            Left            =   225
            TabIndex        =   7
            Top             =   682
            Width           =   2580
         End
      End
      Begin VB.PictureBox picTree 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   90
         ScaleHeight     =   1590
         ScaleWidth      =   2940
         TabIndex        =   2
         Top             =   45
         Width           =   2940
         Begin MSComctlLib.TreeView tvw方案 
            Height          =   1200
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   2117
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            Sorted          =   -1  'True
            Style           =   7
            Appearance      =   0
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "评分方案"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   225
            TabIndex        =   4
            Top             =   90
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1155
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm评分标准维护.frx":2161
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm评分标准维护.frx":2FB3
            Key             =   "RootSel"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgBG_fg 
      Height          =   2970
      Index           =   2
      Left            =   9825
      Top             =   8475
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   5415
      Picture         =   "frm评分标准维护.frx":3227
      Top             =   8595
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgOpen 
      Height          =   225
      Left            =   2565
      Picture         =   "frm评分标准维护.frx":3276
      Top             =   8565
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgBG 
      Height          =   1695
      Left            =   2940
      Picture         =   "frm评分标准维护.frx":32CB
      Top             =   8520
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image imgBGBlue 
      Height          =   1530
      Left            =   75
      Picture         =   "frm评分标准维护.frx":3489
      Top             =   8505
      Visible         =   0   'False
      Width           =   2790
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   270
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frm评分标准维护.frx":3649
      Left            =   735
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
      ScaleMode       =   1
   End
   Begin VB.Image imgBG_fg 
      Height          =   2970
      Index           =   0
      Left            =   7725
      Picture         =   "frm评分标准维护.frx":365D
      Top             =   8490
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image imgBG_fg 
      Height          =   2970
      Index           =   1
      Left            =   5865
      Picture         =   "frm评分标准维护.frx":3E81
      Top             =   8535
      Visible         =   0   'False
      Width           =   3105
   End
End
Attribute VB_Name = "frm评分标准维护"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////
'
'       模块：评分标准维护
'       功能：病案评分标准的录入、修改、删除、打印、选用等。
'       编写：吴庆伟
'       日期：2005年1月5日
'
'///////////////////////////////////////////////////////////////////////////////


Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private mstrPrivs               As String               '权限串
Private mobjFindKey             As CommandBarPopup      '查询
Private mstrFindKey             As String               '查询串
Private mlngModule              As Long                 '模块号
Private m_lngOldRow             As Long                 '
Private m_lngCurRow             As Long                 '
Private m_lngCurID              As Long                 '记录当前记录ID
Private m_lngCurFAID            As Long                 '方案ID
Private m_lngCurSJID            As Long                 '上级ID
Private m_strTreeKey            As String
Private m_lngOldSJID            As Long
Private mzlPrintModeS           As gzlPrintModeS        '打印
Private mintItemID              As Long                 '标准ID
Private mcbrPopupBarProg        As CommandBar           '弹出窗口【分类】
Private mcbrPopupBarItem        As CommandBar           '弹出窗口【项目】
Private mblnProgUsed            As Boolean              '方案是否已使用
Dim cbrPopupItem                As CommandBarControl    '弹出项

'==============================================================================
'=功能： 控件初始化
'==============================================================================
Private Sub InitControl()
    On Error GoTo errH
    '菜单控制
    Call InitCommandBar
    '加载区域
    Call InitDockPannel
    '填充Tree
    Call InitTreeView
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
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    '------------------------------------------------------------------------------------------------------------------
    '编辑
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewKind, "新增方案(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyKind, "修改方案(&F)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteKind, "删除方案(&L)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Import, "导入方案(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Select, "选用方案(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewParent, "新增项目(&X)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Insert, "插入项目(&G)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyParent, "修改项目(&R)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteParent, "删除项目(&C)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "增加标准(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_CopyNewItem, "插入制标准(&I)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "修改标准目(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除标准(&D)")
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
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
    '主菜单右侧的查找
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    mstrFindKey = Trim(GetPara("定位依据", mlngModule, "项目", True))
    If mstrFindKey = "" Then mstrFindKey = "项目"
    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.项目", , , "项目")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.缺陷", , , "缺陷")
    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = txtLocation.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "前一条")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "后一条")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "增加", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
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
    End With
    '------------------------------------------------------------------------------------------------------------------
    '弹出菜单分类
    Set mcbrPopupBarProg = cbsMain.Add("弹出方案菜单", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBarProg.Controls.Add(xtpControlButton, conMenu_Edit_NewKind, "新增方案(&N)")
    Set cbrPopupItem = mcbrPopupBarProg.Controls.Add(xtpControlButton, conMenu_Edit_ModifyKind, "修改方案(&F)")
    Set cbrPopupItem = mcbrPopupBarProg.Controls.Add(xtpControlButton, conMenu_Edit_DeleteKind, "删除方案(&L)")
    Set cbrPopupItem = mcbrPopupBarProg.Controls.Add(xtpControlButton, conMenu_Edit_Import, "导入方案(&P)")
    Set cbrPopupItem = mcbrPopupBarProg.Controls.Add(xtpControlButton, conMenu_Edit_Select, "选用方案(&S)")
    Set mcbrPopupBarItem = cbsMain.Add("弹出项目菜单", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_NewParent, "新增项目(&X)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Insert, "插入项目(&G)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改项目(&R)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_DeleteParent, "删除项目(&C)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "增加标准(&A)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_CopyNewItem, "插入标准(&I)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "修改标准(&M)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除标准(&D)")
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 界面分割
'==============================================================================
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo errH
    Select Case Item.ID
        Case 1
            Item.Handle = picLeft_S.hWnd
        Case 2
            Item.Handle = picRightUp.hWnd
        Case 3
            Item.Handle = pic项目信息_S.hWnd
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
    
    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "方案"
    objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 200, 100, DockRightOf, Nothing)
    objPane.Title = "项目"
    objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(3, 200, 100, DockBottomOf, objPane)
    objPane.Title = "标准"
    objPane.Options = PaneNoCaption
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
            Call SendLMouseButton(fgMain.hWnd, X, Y)
            mcbrPopupBarItem.ShowPopup
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 双击编辑该评分标准或项目（独立项目时）
'==============================================================================
Private Sub fgMain_DblClick()
    On Error GoTo errH
    
    If InStr(mstrPrivs, "增删改") = 0 Then Exit Sub
    If fgMain.MouseRow = 0 Or mblnProgUsed Then Exit Sub
    Call StandardEdit
    
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
    Dim rsTemp      As ADODB.Recordset
    On Error GoTo errH
    
    fgMain.WallPaper = imgBG_fg(2).Picture
    m_lngCurRow = fgMain.Row
    mblnProgUsed = False
    If m_lngCurRow <= 0 Then
        m_lngCurSJID = 0
        m_lngCurID = 0
        fgMain.WallPaper = imgBG_fg(0).Picture
        Exit Sub
    End If
    
    m_lngCurID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 4)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 4)))    '获取ID
    m_lngCurSJID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 5)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 5)))     '获取ID
    m_lngCurFAID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 6)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 6)))     '获取ID
    If m_lngCurSJID = 0 Then
        lngID = m_lngCurID
    Else
        lngID = m_lngCurSJID
    End If
    
    Show基本要求 lngID, fgMain.Cell(flexcpText, m_lngCurRow, 0), fgMain.Cell(flexcpText, m_lngCurRow, 1)
    m_lngOldRow = m_lngCurRow
    
    gstrSQL = "select count(*) from 病案评分结果 where 方案ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If rsTemp(0).Value > 0 Then
        '该方案已经使用
        mblnProgUsed = True
        fgMain.WallPaper = imgBG_fg(1).Picture
    End If
    Call SetMenu
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
    m_lngOldRow = -1
    m_lngCurRow = -1
    m_lngCurID = -1
    m_lngOldSJID = -1
    mblnProgUsed = False
    mstrPrivs = UserInfo.模块权限
    mlngModule = ParamInfo.模块号
    If GetPersonSet Then
        mstrFindKey = Trim(GetPara("定位依据", mlngModule, "名称", True))
    End If
    '控件初始化
    Call InitControl
    
    '填充列表
    Call DataLoad

    '恢复界面位置
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    Call SetMenu
    
    picFAXX.Picture = imgClose.Picture
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
    Call SetPaneRange(dkpMain, 2, 400, 100, ScaleHeight, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 3, 400, 100, ScaleHeight, Me.ScaleHeight)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 窗口关闭时保存参数
'==============================================================================
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errH
    m_strTreeKey = ""
    '使用个性化设置
    Call SetPara("定位依据", mstrFindKey, mlngModule)
    SaveWinState Me, App.ProductName
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 方案信息关闭或显示
'==============================================================================
Private Sub picFAXX_Click()
    On Error GoTo errH
    
    If picFAXX.Tag = "" Then
        picFAXX.Tag = "Opened"
        picFAXX.Picture = imgOpen.Picture
        pic方案信息.Height = 340
    Else
        picFAXX.Tag = ""
        picFAXX.Picture = imgClose.Picture
        pic方案信息.Height = 1695
    End If
    picFAXX.Refresh
    Call picLeft_S_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 数据统计
'==============================================================================
Private Sub DataUpdate()
    Dim rs              As ADODB.Recordset
    Dim lng总分         As Double
    On Error GoTo errH
    gstrSQL = "select 名称,分制,上值,下值,总分 from 病案评分方案 where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If Not rs.EOF Then
        lbl方案名称.Caption = rs("名称")
        lbl分制.Caption = "分制:" & rs("分制")
        lbl上值.Caption = "上值:" & rs("上值")
        lbl下值.Caption = "下值:" & rs("下值")
        lbl总分.Caption = "总分:" & rs("总分")
        lng总分 = rs("总分")
    Else
        lbl方案名称.Caption = ""
        lbl分制.Caption = ""
        lbl上值.Caption = ""
        lbl下值.Caption = ""
        lbl总分.Caption = ""
    End If

    gstrSQL = "select sum(标准分值) from 病案评分标准 where 上级ID is null and 方案ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If Not rs.EOF Then
        If Abs(lng总分 - rs.Fields(0)) > 0.01 Then
            lbl总分 = lbl总分 + "，项目分数和为:" & rs.Fields(0)
            lbl总分.ForeColor = vbRed
        Else
            lbl总分.ForeColor = vbBlack
        End If
    Else
        lbl总分.ForeColor = vbRed
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 方案信息焦点变色
'==============================================================================
Private Sub picFAXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errH
    If X >= 0 And X <= picFAXX.ScaleWidth And Y >= 0 And Y <= picFAXX.ScaleHeight Then
        SetCapture picFAXX.hWnd
        '鼠标移入！！！
        picFAXX.Line (0, 0)-(picFAXX.ScaleWidth - Screen.TwipsPerPixelX, picFAXX.ScaleHeight - Screen.TwipsPerPixelY), vbBlue, B
    Else
        '鼠标移出！！！
        picFAXX.Cls
        ReleaseCapture
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 左侧picLeft，控件位置控制
'==============================================================================
Private Sub picLeft_S_Resize()
On Error Resume Next
    pic方案信息.Move 135, picLeft_S.ScaleHeight - pic方案信息.Height - 270 * 2, picLeft_S.ScaleWidth - 270
    With picTree
        .Move 135, 135, pic方案信息.Width, Abs(picLeft_S.ScaleHeight - pic方案信息.Height - 270 * 3)
        .Cls
        .PaintPicture imgBGBlue.Picture, 0, 0, picTree.Width, 360, 0, 0, imgBGBlue.Width, 360
        .PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, picTree.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
        .PaintPicture imgBGBlue.Picture, picTree.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picTree.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
        .PaintPicture imgBGBlue.Picture, 0, picTree.ScaleHeight - Screen.TwipsPerPixelY, picTree.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
    End With
    tvw方案.Move Screen.TwipsPerPixelX * 4, 390, Abs(picTree.ScaleWidth - 8 * Screen.TwipsPerPixelX), Abs(picTree.ScaleHeight - 390 - Screen.TwipsPerPixelY * 4)
    With pic方案信息
        .Cls
        .PaintPicture imgBGBlue.Picture, 0, 0, pic方案信息.Width, 360, 0, 0, imgBGBlue.Width, 360
        .PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, pic方案信息.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
        .PaintPicture imgBGBlue.Picture, pic方案信息.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic方案信息.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
        .PaintPicture imgBGBlue.Picture, 0, pic方案信息.ScaleHeight - Screen.TwipsPerPixelY, pic方案信息.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
    End With
    picFAXX.Move pic方案信息.ScaleWidth - picFAXX.Width - 80
    Refresh
End Sub

'==============================================================================
'=功能： 右侧picRightUp，控件位置控制
'==============================================================================
Private Sub picRightUp_Resize()
    On Error GoTo errH
    fgMain.Move 15, 15, picRightUp.Width - 30, picRightUp.Height - 30
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 右侧picRightUp，控件位置控制
'==============================================================================
Private Sub pic项目信息_S_Resize()
    On Error GoTo errH
    lblInfo.Move lblInfo.Left, lblInfo.Top, Abs(pic项目信息_S.ScaleWidth - 2 * lblInfo.Left), Abs(pic项目信息_S.ScaleHeight - lblInfo.Top)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 左上侧picTree，树双击更新
'==============================================================================
Private Sub picTree_DblClick()
    On Error GoTo errH
    If Left(tvw方案.SelectedItem.Key, 4) = "Root" Then Exit Sub
    Call ProgEdit
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 左上侧picTree，树双击更新
'==============================================================================
Private Sub picTree_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    If IsNumeric(Mid(m_strTreeKey, 2)) Then
        If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call ProgEdit
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 右击方案时，弹出方案右键菜单
'==============================================================================
Private Sub tvw方案_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errH
    If InStr(mstrPrivs, "增删改") = 0 Then Exit Sub
    Select Case Button
        Case 2          '弹出菜单处理
            Call SendLMouseButton(tvw方案.hWnd, X, Y)
            mcbrPopupBarProg.ShowPopup
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 方案更新后选择相应的项目和标准
'==============================================================================
Private Sub tvw方案_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim rs              As ADODB.Recordset
    Dim lng总分         As Double
    
    On Error GoTo errH
    
    If m_strTreeKey = Node.Key Then Exit Sub     '避免重复刷新
    m_strTreeKey = Node.Key
    m_lngCurFAID = Val(Mid(m_strTreeKey, 2))
    
    gstrSQL = "select 名称,分制,上值,下值,总分 from 病案评分方案 where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If Not rs.EOF Then
        lbl方案名称 = rs("名称")
        lbl分制 = "分制:" & rs("分制")
        lbl上值 = "上值:" & rs("上值")
        lbl下值 = "下值:" & rs("下值")
        lbl总分 = "总分:" & rs("总分")
        lng总分 = rs("总分")
    Else
        lbl方案名称 = ""
        lbl分制 = ""
        lbl上值 = ""
        lbl下值 = ""
        lbl总分 = ""
    End If
    
    gstrSQL = "select sum(标准分值) from 病案评分标准 where 上级ID is null and 方案ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If Not rs.EOF Then
        If Abs(lng总分 - rs.Fields(0)) > 0.01 Then
            lbl总分 = lbl总分 + "，项目分数和为:" & rs.Fields(0)
            lbl总分.ForeColor = vbRed
        Else
            lbl总分.ForeColor = vbBlack
        End If
    Else
        lbl总分.ForeColor = vbRed
    End If
    '填充列表
    Call DataLoad
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
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
    objPrint.Title.Text = tvw方案.SelectedItem.Text
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    Set objAppRow = New zlTabAppRow
    gstrSQL = "select ID,名称,总分,上值,下值,类型,分制,选用,启用时间,停用时间 from 病案评分方案 where ID= [1]" ' & m_lngCurFAID
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If Not rs.EOF Then
        objAppRow.Add "总分:" & NVL(rs("总分"), 0)
        objAppRow.Add "甲级分数线:" & NVL(rs("上值"), 0)
        objAppRow.Add "乙级分数线:" & NVL(rs("下值"), 0)
        
        objPrint.UnderAppRows.Add objAppRow
    End If
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

'==============================================================================
'=功能： 装入评分方案 目前只考虑住院病案
'==============================================================================
Private Sub InitTreeView()
    Dim rsTemp          As ADODB.Recordset
    Dim nod             As Node
    Dim i               As Long
    Dim FirstKey        As String
    Dim v               As Variant
    
    On Error GoTo errH
    
    fgMain.Tag = ""
    'Tree的初始化
    Set tvw方案.ImageList = ils16
    tvw方案.Nodes.Clear
    
    '注意调用格式：先赋值gstrSQL,然后打开数据集
    gstrSQL = "select ID,名称,选用 from 病案评分方案 where 类型='住院' Order by 选用 desc,名称,启用时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    i = 1
    Do Until rsTemp.EOF
        '添加子节点
        Set nod = tvw方案.Nodes.Add(, , "A" & rsTemp("ID"), rsTemp("名称"), IIf(rsTemp("选用") = 1, "RootSel", "Root"), IIf(rsTemp("选用") = 1, "RootSel", "Root"))
        If rsTemp("选用") = 1 Then
            nod.Bold = True
        Else
            nod.Bold = False
        End If
        If i = 1 Then FirstKey = nod.Key
        If FirstKey = nod.Key Then i = 2
        If FirstKey = "" And i = 1 Then FirstKey = nod.Key: i = 2
        rsTemp.MoveNext
    Loop
    'm_strTreeKey不为空，但是又没有找到。
    If i = 1 Then m_strTreeKey = FirstKey
    For Each v In tvw方案.Nodes
        If v.Key = FirstKey Then
            '设置选中
            v.Selected = True
            v.EnsureVisible
            If picTree.Visible = True Then picTree.SetFocus
        End If
    Next
    If Not tvw方案.SelectedItem Is Nothing Then tvw方案_NodeClick tvw方案.SelectedItem
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 装入对应方案的评分标准
'==============================================================================
Public Sub DataLoad()
    Dim rsTemp      As ADODB.Recordset
    Dim i           As Long
    
    On Error GoTo errH
    
    With fgMain
        .Tag = ""
        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        .Cell(flexcpText, 0, 0) = "项目"
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 1) = "标准分值"
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 2) = "标准名称"
        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 3) = "评分标准"
        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 4) = "ID"
        .Cell(flexcpText, 0, 5) = "上级ID"
        .Cell(flexcpText, 0, 6) = "方案ID"
        .Cell(flexcpText, 0, 7) = "序号"
        
        '确定方案名称
        If tvw方案.SelectedItem Is Nothing Then .Redraw = flexRDDirect: Exit Sub
        With tvw方案.SelectedItem
            Select Case Left(.Key, 1)
                Case "A", "B"
                    m_lngCurFAID = Val(Mid(.Key, 2))
                    gstrSQL = "select 上级序号,序号,ID,上级ID,方案ID,项目,标准分值,基本要求,缺陷内容,扣分标准,隐藏 from 病案评分标准视图 Where 隐藏='否' and 方案ID = [1]"
                Case Else
                    Call SetMenu
                    fgMain.Redraw = flexRDDirect
                    Exit Sub
            End Select
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(CStr(Mid(.Key, 2))))
        End With
        
        .FocusRect = flexFocusSolid
        '数据填入
        .Cols = 8
        .Rows = rsTemp.RecordCount + 1
        i = 1
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, 0) = NVL(rsTemp.Fields("项目"))
            .Cell(flexcpAlignment, i, 0) = flexAlignCenterCenter
            .Cell(flexcpText, i, 1) = IIf(IsNull(rsTemp.Fields("标准分值")), " ", Format(rsTemp.Fields("标准分值"), "####分"))
            .Cell(flexcpAlignment, i, 1) = flexAlignCenterCenter
            .Cell(flexcpText, i, 2) = NVL(rsTemp.Fields("缺陷内容"))
            .Cell(flexcpAlignment, i, 2) = flexAlignLeftTop
            .Cell(flexcpText, i, 3) = IIf(IsNull(rsTemp.Fields("扣分标准")), "", IIf(rsTemp.Fields("扣分标准") = "甲", "甲级", IIf(rsTemp.Fields("扣分标准") = "乙", "乙级", IIf(rsTemp.Fields("扣分标准") = "丙", "丙级", IIf(rsTemp.Fields("扣分标准") = "否", "单项否决", rsTemp.Fields("扣分标准"))))))
            .Cell(flexcpAlignment, i, 3) = flexAlignCenterCenter
            .Cell(flexcpText, i, 4) = NVL(rsTemp.Fields("ID"), 0)
            .Cell(flexcpText, i, 5) = NVL(rsTemp.Fields("上级ID"), 0)
            .Cell(flexcpText, i, 6) = NVL(rsTemp.Fields("方案ID"), 0)
            .Cell(flexcpText, i, 7) = NVL(rsTemp.Fields("序号"), 0)
            rsTemp.MoveNext
            i = i + 1
        Loop
        '自动换行
        .WordWrap = True
        '合并单元格
        .MergeCells = 2
        .MergeCol(.ColIndex("项目")) = True
        .MergeCol(.ColIndex("标准分值")) = True
        '对齐设置
        .ColAlignment(.ColIndex("项目")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("标准分值")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("评分标准")) = flexAlignCenterCenter
        '隐藏单元格
        .ColWidth(.ColIndex("ID")) = 0
        .ColWidth(.ColIndex("上级ID")) = 0
        .ColWidth(.ColIndex("方案ID")) = 0
        .ColWidth(.ColIndex("序号")) = 0
        '宽度设置
        .ColWidth(.ColIndex("项目")) = 1500
        .ColWidth(.ColIndex("标准分值")) = 850
        .ColWidth(.ColIndex("缺陷内容")) = 3700
        .ColWidth(.ColIndex("评分标准")) = 1100
        '行高设置
'        .RowHeightMin = 300
        '最大宽度设置
'        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize .ColIndex("缺陷内容")
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        '选中先前的行
        If m_lngOldRow > 0 And m_lngOldRow < i Then
            .Row = m_lngOldRow
            .Col = 2
            .ShowCell m_lngOldRow, 2
            On Error Resume Next
            If .Visible = True Then .SetFocus
            fgMain_SelChange
        ElseIf fgMain.Tag = "" And i > 1 And .Rows > 1 Then
            m_lngOldRow = 1
            fgMain.Tag = "选中第一行"
            .Row = 1
            .Col = 2
            .ShowCell m_lngOldRow, 2
            On Error Resume Next
            If .Visible = True Then .SetFocus
            fgMain_SelChange
        Else
            lblInfo = "无内容"
        End If
    End With
    
    Call DataUpdate
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
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

'==============================================================================
'=功能： 根据项目ID显示基本要求
'==============================================================================
Private Sub Show基本要求(lngID As Long, 项目 As String, 标准分值 As String)
    Dim rs          As ADODB.Recordset
    
    On Error GoTo errH
    gstrSQL = "select ID,描述 as 基本要求,上级ID from 病案评分标准 Where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
    
    If Not rs.EOF Then
        If m_lngOldSJID > 0 And m_lngOldSJID = lngID Then Exit Sub
        If IsNull(rs.Fields("基本要求")) Then
                lblInfo = "名称：" + 项目 + "  " + IIf(Len(Trim(标准分值)) = 0, "", "(" + 标准分值 + ")")
                lblInfo = lblInfo + vbCrLf
        Else
            If Len(rs.Fields("基本要求")) > 0 Then
                lblInfo = "名称：" + 项目 + "  " + IIf(Len(Trim(标准分值)) = 0, "", "(" + 标准分值 + ")")
                lblInfo = lblInfo + vbCrLf + rs.Fields("基本要求")
            End If
        End If
    Else
        lblInfo.Caption = "无内容":
    End If
    m_lngOldSJID = m_lngCurSJID
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 取得是否保存布局
'==============================================================================
Private Function GetPersonSet() As Boolean
    
    On Error GoTo errH
    
    GetPersonSet = False
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then GetPersonSet = True

    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 菜单功能控制
'==============================================================================
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnNewCancel        As Boolean
    On Error GoTo errH
    
    Select Case Control.ID
        Case conMenu_Edit_NewKind           '增加方案
            Call ProgAdd
        Case conMenu_Edit_ModifyKind        '修改方案
            Call ProgEdit
        Case conMenu_Edit_DeleteKind        '删除方案
            Call ProgDel
        Case conMenu_Edit_Import            '导入方案
            Call ProgImp
        Case conMenu_Edit_Select            '选用方案
            Call ProgSele
        Case conMenu_Edit_NewParent         '新增项目
            Call ItemAdd
        Case conMenu_Edit_Insert            '插入项目
            Call ItemInsert
        Case conMenu_Edit_ModifyParent      '修改项目
            Call ItemEdit
        Case conMenu_Edit_DeleteParent      '删除项目
            Call ItemDel
        Case conMenu_Edit_NewItem           '增加标准
            Call StandardAdd
        Case conMenu_Edit_CopyNewItem       '插入标准
            Call StandardInsrt
        Case conMenu_Edit_Modify            '修改标准
            Call StandardEdit
        Case conMenu_Edit_Delete            '删除标准
            Call StandardDel
        Case conMenu_View_Refresh           '刷新数据
            Call InitTreeView
        Case conMenu_File_Preview           '预览
            Call ItemPrint
        Case conMenu_File_Print             '打印
            mzlPrintModeS = zlPrint
            Call ItemPrint
        Case conMenu_File_Excel             '输出到&Excel
            mzlPrintModeS = zlExcel
            Call ItemPrint
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
            If Control.ID > 400 And Control.ID < 500 Then
                Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
            Else
                 '与业务无关的功能，公共的功能
                Call CommandBarExecutePublic(Control, Me, fgMain, "病案评分标准维护")
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 菜单权限控制
'==============================================================================
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo errH
    
    With fgMain
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel
                Control.Enabled = ((fgMain.Rows > 1) And IsPrivs(mstrPrivs, "增删改"))
            Case conMenu_Edit_NewKind           '增加方案
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = Control.Visible
            Case conMenu_Edit_ModifyKind       '修改方案
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw方案.Nodes.count > 0
            Case conMenu_Edit_DeleteKind       '删除方案
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw方案.Nodes.count > 0
            Case conMenu_Edit_Import           '导入方案
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed And (fgMain.Rows <= 1)) And tvw方案.Nodes.count > 0
            Case conMenu_Edit_Select           '选用方案
                Control.Enabled = (fgMain.Rows > 0) And tvw方案.Nodes.count > 0
            Case conMenu_Edit_NewParent        '新增项目
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw方案.Nodes.count > 0
            Case conMenu_Edit_Insert           '插入项目
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw方案.Nodes.count > 0
            Case conMenu_Edit_ModifyParent     '修改项目
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw方案.Nodes.count > 0
            Case conMenu_Edit_DeleteParent     '删除项目
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw方案.Nodes.count > 0
            Case conMenu_Edit_NewItem          '增加标准
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw方案.Nodes.count > 0
            Case conMenu_Edit_CopyNewItem      '插入标准
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw方案.Nodes.count > 0
            Case conMenu_Edit_Modify           '修改标准
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (fgMain.Row > 0) Or (Control.Visible And Not mblnProgUsed And m_lngCurSJID <> 0)
            Case conMenu_Edit_Delete           '删除标准
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = (Control.Visible And Not mblnProgUsed And m_lngCurSJID <> 0)
            Case conMenu_View_Forward
                Control.Enabled = (Control.Visible And fgMain.Row > 1)
            Case conMenu_View_Backward
                Control.Enabled = (Control.Visible And fgMain.Row + 1 < fgMain.Rows)
            Case conMenu_View_Refresh
                Control.Visible = IsPrivs(mstrPrivs, "增删改")
                Control.Enabled = Control.Visible
            Case conMenu_View_LocationItem, conMenu_View_LocationItem, conMenu_View_LocationItem
                If InStr(Control.Caption, mstrFindKey) > 0 Then
                    Control.Checked = True
                Else
                    Control.Checked = False
                End If
    
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
'=增加评分方案
'==============================================================================
Private Sub ProgAdd()
    Dim f As New frm评分方案编辑
    On Error GoTo errH
    f.ShowForm   '新增
    Call DataUpdate
    If f.Moded Then
        Call InitTreeView
        '填充列表
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=选用方案
'==============================================================================
Private Sub ProgSele()
    Dim intIndex        As Long
    Dim bln已使用       As Boolean
    
    On Error GoTo errH
    
    If m_lngCurFAID < 1 Then Exit Sub
    If MsgBox("注意：评分分案的选用是一件非常慎重的事情，通常不要随意更改！" & vbCrLf & "请确认选用本评分方案吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "select count(*) from 病案评分结果 where 方案ID=(select ID from 病案评分方案 where 类型='住院' and 选用=1)"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    If rsTemp(0).Value > 0 Then
        '默认住院方案已经使用
        If MsgBox("注意：系统默认评分分案正在使用当中，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    rsTemp.Close
    
    gstrSQL = "ZL_病案评分方案_选用(" & CStr(m_lngCurFAID) & ",1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call InitTreeView
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=修改评分方案
'==============================================================================
Private Sub ProgEdit()
    Dim f               As New frm评分方案编辑
    Dim lng总分         As Double
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If m_lngCurFAID < 1 Then Exit Sub
    f.ShowForm m_lngCurFAID   '修改，传入ID
    Call DataUpdate
    If f.Moded Then
        Call InitTreeView
        '填充列表
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=导入已有方案
'==============================================================================
Private Sub ProgImp()
    Dim lID             As Long     '选中的方案ID
    Dim lNewID          As Long
    Dim f               As New frm选择评分方案
    Dim rs              As ADODB.Recordset
    Dim lng总分         As Double
    Dim rsTmp           As ADODB.Recordset
    Dim strT            As String

    On Error GoTo errH
    
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If m_lngCurFAID <= 0 Then Exit Sub
    
    f.FillCmbSelFA m_lngCurFAID
    f.Show 1
    lID = f.ID_From
    
    '执行导入操作！！！
    '源ID为：lID   目的ID为： m_lngCurFAID
    gstrSQL = "Select ID, 上级id, 方案id, 名称, 描述, 标准分值, 缺陷等级, 评分单位, 上级序号, 序号, 判断依据, 否决等级, 数据源" & vbNewLine & _
                "From 病案评分标准" & vbNewLine & _
                "Where 上级id Is Null And 方案id = [1]" & vbNewLine & _
                "Order By 上级序号, 序号, ID"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lID)
    zlCommFun.ShowFlash "请稍候，系统正在导入评分方案……", Me
    DoEvents
    
    On Error GoTo LL
    
    gcnOracle.BeginTrans
    Do While Not rs.EOF
        '找到了项目，添加项目
        lNewID = zlDatabase.GetNextId("病案评分标准")
        gstrSQL = "ZL_病案评分标准_Insert" & _
            "(" & lNewID & "," & NVL(rs!上级ID, "NULL") & "," & m_lngCurFAID & _
            ",'" & NVL(rs!名称) & "','" & NVL(rs!描述) & "'," & NVL(rs!标准分值, "NULL") & ",'" & NVL(rs!缺陷等级) & _
            "','" & NVL(rs!评分单位) & "',0,'" & Replace(NVL(rs!判断依据), "'", "''") & "','" & NVL(rs!否决等级) & "'," & rs!数据源 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        '进一步查找下级项目，循环添加之！
        gstrSQL = "Select ID, 上级id, 方案id, 名称, 描述, 标准分值, 缺陷等级, 评分单位, 上级序号, 序号, 判断依据, 否决等级, 数据源" & vbNewLine & _
                "From 病案评分标准" & vbNewLine & _
                "Where 上级id = [1] And 方案id = [2]" & vbNewLine & _
                "Order By 上级序号, 序号, ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(rs("ID")), lID)
        Do While Not rsTmp.EOF
            gstrSQL = "ZL_病案评分标准_Insert" & _
                "(" & zlDatabase.GetNextId("病案评分标准") & "," & lNewID & "," & m_lngCurFAID & _
                ",'" & NVL(rsTmp!名称) & "','" & NVL(rsTmp!描述) & "'," & NVL(rsTmp!标准分值, "NULL") & ",'" & NVL(rsTmp!缺陷等级) & _
                "','" & NVL(rsTmp!评分单位) & "',0,'" & Replace(NVL(rs!判断依据), "'", "''") & "','" & NVL(rs!否决等级) & "'," & rs!数据源 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            rsTmp.MoveNext
        Loop
        rs.MoveNext
    Loop
    
    '刷新结果！
    gcnOracle.CommitTrans
    
    Call DataLoad
    zlCommFun.StopFlash
    Exit Sub
LL:
    gcnOracle.RollbackTrans
    zlCommFun.StopFlash
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=删除评分方案
'==============================================================================
Private Sub ProgDel()
    Dim intIndex        As Long
    
    On Error GoTo errH
    
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If m_lngCurFAID < 1 Then Exit Sub
    If MsgBox("你确认要删除该条方案吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "ZL_病案评分方案_Delete(" & CStr(m_lngCurFAID) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call InitTreeView
    Call DataLoad
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=新增项目
'==============================================================================
Private Sub ItemAdd()
    Dim f As New frm评分标准编辑
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    f.ShowForm "新增", m_lngCurFAID
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=插入项目
'==============================================================================
Private Sub ItemInsert()
    Dim f As New frm评分标准编辑
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If m_lngCurSJID < 1 Then '为独立评分项
        f.ShowForm "插入", m_lngCurFAID, 0, m_lngCurID
    Else
        f.ShowForm "插入", m_lngCurFAID, 0, m_lngCurSJID
    End If
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=修改项目
'==============================================================================
Private Sub ItemEdit()
    Dim f As New frm评分标准编辑
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If m_lngCurID < 1 Then Exit Sub
    If m_lngCurSJID < 1 Then
        f.ShowForm "修改", m_lngCurFAID, , m_lngCurID
    Else
        f.ShowForm "修改", m_lngCurFAID, , m_lngCurSJID
    End If
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=删除项目
'==============================================================================
Private Sub ItemDel()
    
    Dim intIndex As Long
    
    On Error GoTo errH
    
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    
    If m_lngCurID < 1 Then Exit Sub
    
    If MsgBox("你确认要删除该条评分项目吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If m_lngCurSJID = 0 Then
        gstrSQL = "ZL_病案评分标准_Delete(" & CStr(m_lngCurID) & ",0)"
    Else
        gstrSQL = "ZL_病案评分标准_Delete(" & CStr(m_lngCurSJID) & ",0)"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call DataLoad
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=新增标准
'==============================================================================
Private Sub StandardAdd()
    Dim f As New frm评分标准编辑
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If m_lngCurSJID < 1 Then '为独立评分项
        f.ShowForm "新增", m_lngCurFAID, m_lngCurID
    Else
        f.ShowForm "新增", m_lngCurFAID, m_lngCurSJID
    End If
    
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=插入标准
'==============================================================================
Private Sub StandardInsrt()
    Dim f As New frm评分标准编辑
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If m_lngCurSJID < 1 Then '为独立评分项
        f.ShowForm "插入", m_lngCurFAID, m_lngCurID, m_lngCurSJID
    Else
        f.ShowForm "插入", m_lngCurFAID, m_lngCurSJID, m_lngCurID
    End If
    
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=修改标准
'==============================================================================
Private Sub StandardEdit()
    Dim f As New frm评分标准编辑
    On Error GoTo errH
    If ObjPtr(tvw方案.SelectedItem) = 0 Then Exit Sub
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If m_lngCurID < 1 Then Exit Sub
    If fgMain.Col < 2 Then  '一级项目
        If m_lngCurSJID < 1 Then
            f.ShowForm "修改", m_lngCurFAID, , m_lngCurID, Not mblnProgUsed
        Else
            f.ShowForm "修改", m_lngCurFAID, , m_lngCurSJID, Not mblnProgUsed
        End If
    Else                    '子项目
        f.ShowForm "修改", m_lngCurFAID, m_lngCurSJID, m_lngCurID, Not mblnProgUsed
    End If
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=删除标准
'==============================================================================
Private Sub StandardDel()
    Dim intIndex As Long
    
    On Error GoTo errH
    
    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
    If m_lngCurID < 1 Then Exit Sub
    If MsgBox("你确认要删除该条评分标准吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "ZL_病案评分标准_Delete(" & CStr(m_lngCurID) & ",1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call DataLoad
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
    If mstrFindKey = "缺陷" Then mstrFindKey = "缺陷内容"
    If KeyAscii = vbKeyReturn Then
        '读取大于当前行的记录数据
        For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
            If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '读取小于当前行的记录数据
        If lngRow = 0 Then
            For lngLoop = 0 To fgMain.Row
                If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If fgMain.Rows > 1 And lngRow >= 1 Then fgMain.Row = lngRow
        
        Call LocationObj(txtLocation)
    End If
    If mstrFindKey = "缺陷内容" Then mstrFindKey = "缺陷"
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
