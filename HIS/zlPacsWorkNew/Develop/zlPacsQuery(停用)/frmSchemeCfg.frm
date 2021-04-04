VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSchemeCfg 
   Caption         =   "查询方案配置"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   15120
   Icon            =   "frmSchemeCfg.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8220
   ScaleWidth      =   15120
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2760
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSchemeName 
      Height          =   5295
      Left            =   960
      ScaleHeight     =   5235
      ScaleWidth      =   2715
      TabIndex        =   17
      Top             =   1080
      Width           =   2775
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   4935
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   2415
         _cx             =   4260
         _cy             =   8705
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
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
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picSchemeContent 
      Height          =   7215
      Left            =   3960
      ScaleHeight     =   7155
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   480
      Width           =   10935
      Begin VB.Frame fraBasic 
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   10575
         Begin VB.TextBox txtSchemeName 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   1
            Top             =   270
            Width           =   2475
         End
         Begin VB.TextBox txtSchemeMemo 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtPage 
            Height          =   270
            Left            =   6690
            TabIndex        =   7
            Text            =   "0"
            Top             =   892
            Width           =   495
         End
         Begin VB.TextBox txtDate 
            Height          =   270
            Left            =   8730
            TabIndex        =   8
            Text            =   "0"
            Top             =   892
            Width           =   495
         End
         Begin VB.CheckBox chkTrance 
            Caption         =   "列表功能跟随"
            Height          =   375
            Left            =   4080
            TabIndex        =   6
            Top             =   840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkHistory 
            Caption         =   "显示检查历史"
            Height          =   375
            Left            =   2520
            TabIndex        =   5
            Top             =   840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkGroup 
            Caption         =   "启用分组"
            Height          =   375
            Left            =   1320
            TabIndex        =   4
            Top             =   840
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkCard 
            Caption         =   "启用刷卡"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.Frame fraLine 
            Height          =   30
            Left            =   0
            TabIndex        =   10
            Top             =   720
            Width           =   10215
         End
         Begin VB.Label labSchemeMemo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "方案说明:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   16
            Top             =   330
            Width           =   975
         End
         Begin VB.Label labObj 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "方案名称:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   330
            Width           =   975
         End
         Begin VB.Label lblPage 
            AutoSize        =   -1  'True
            Caption         =   "分页大小："
            Height          =   180
            Left            =   5760
            TabIndex        =   14
            Top             =   937
            Width           =   900
         End
         Begin VB.Label lblRows 
            AutoSize        =   -1  'True
            Caption         =   "行"
            Height          =   180
            Left            =   7290
            TabIndex        =   13
            Top             =   937
            Width           =   180
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "日期范围："
            Height          =   180
            Left            =   7770
            TabIndex        =   12
            Top             =   937
            Width           =   900
         End
         Begin VB.Label lblYears 
            AutoSize        =   -1  'True
            Caption         =   "年"
            Height          =   180
            Left            =   9330
            TabIndex        =   11
            Top             =   937
            Width           =   180
         End
      End
      Begin XtremeSuiteControls.TabControl tbcMain 
         Height          =   4095
         Left            =   480
         TabIndex        =   20
         Top             =   2160
         Width           =   9615
         _Version        =   589884
         _ExtentX        =   16960
         _ExtentY        =   7223
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   7860
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSchemeCfg.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12938
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Bindings        =   "frmSchemeCfg.frx":70E6
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmSchemeCfg.frx":70FA
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   1320
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSchemeCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModuleNo As Long     '所属模块
Private mblnNewState As Boolean     '是否在新增
Private mblnClose As Boolean
Private mblnIsEdit As Boolean   '方案是否已编辑
Private mlngRow As Long
Private mobjIconManage As New frmIconManage
Private mobjQuerySet As New frmScheme_BaseQueryCfg
Private mobjFilterSet As New frmScheme_FilterCfg
Private mobjDisPlaySet As New frmScheme_DisplayCfg
Private mobjSetRelated As New frmSetRelated
Private mobjSqlScheme As New clsSqlScheme

Private Const SW_SHOW = 5
Private Const M_STR_GRIDDATA = "序号|ID|方案名称|使用状态|是否默认|是否常用|方案说明"       '表格显示
Private Const M_STR_CROOK = "√"
Private Enum ColTitle
    ct序号 = 0
    ctID = 1
    ct方案名称 = 2
    ct使用状态 = 3
    ct是否默认 = 4
    ct是否常用 = 5
    ct方案说明 = 6
End Enum

Private Const conMenu_View_ToolBar = 301              '工具栏(&T)
Private Const conMenu_View_ToolBar_Button = 302         '标准按钮(&S)
Private Const conMenu_View_ToolBar_Text = 303           '文本标签(&T)
Private Const conMenu_View_ToolBar_Size = 304           '大图标(&B)
Private Const conMenu_View_StatusBar = 305            '状态栏(&S)

'菜单类型枚举定义
Private Enum TMenuType
    mtFile = 1                  '文件
    mtSave = 101                '保存
    mtCancel = 102              '关闭
    mtImport = 103              '导入
    mtExport = 104              '导出
    mtQuit = 105                '退出
    
    mtEdit = 2                  '编辑
    mtNewScheme = 201           '新增
    mtModifyScheme = 202        '修改
    mtDelScheme = 203           '删除
    mtUsually = 204             '常用
    mtSetDefault = 205           '默认
    mtRecover = 206              '恢复
    mtUseScheme = 207            '启用/禁用
    mtCheckScheme = 208          '测试
    mtMoveLastScheme = 209       '上移
    mtMoveNextScheme = 210       '下移
    mtSetSysQuery = 214          '用户查询调整
    mtResource = 215             '资源管理
    
    mtViewPopup = 3            '查看
    
'    mtHelpPopup = 4            '帮助
End Enum

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ShowMe(ByVal lngModuleNo As Long, ByVal strSysPara As String, owner As Object)
    mlngModuleNo = lngModuleNo
    gstrPara = strSysPara
    Me.Show , owner
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle
    
    Select Case control.Id
        Case TMenuType.mtSave       '保存
            Call SaveScheme
        Case TMenuType.mtCancel     '关闭
            Call CancelScheme
        Case TMenuType.mtImport     '导入
            Call ImportScheme
        Case TMenuType.mtExport     '导出
            Call ExportScheme
            tbcMain.SetFocus
        Case TMenuType.mtQuit       '退出
            Call UnloadMe
        Case TMenuType.mtNewScheme  '新增
            Call NewScheme
        Case TMenuType.mtModifyScheme   '修改
            Call ModifyScheme
        Case TMenuType.mtDelScheme      '删除
            Call DeleteScheme
        Case TMenuType.mtUsually    '常用
            Call SetUsualScheme
        Case TMenuType.mtSetDefault '默认
            Call SetDefaultScheme
        Case TMenuType.mtRecover    '恢复
            Call RecoverScheme
        Case TMenuType.mtUseScheme  '启用/禁用
            Call SetUseScheme
        Case TMenuType.mtCheckScheme    '测试
        Case TMenuType.mtMoveLastScheme     '上移
            Call MoveLastScheme
        Case TMenuType.mtMoveNextScheme     '下移
            Call MoveNextScheme
        Case TMenuType.mtSetSysQuery    '用户查询调整
            Call ShowUserScheme
        Case TMenuType.mtResource       '资源管理
            Call mobjIconManage.ShowIconWindow("", Me)
            tbcMain.SetFocus
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button '工具栏
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '按钮文字
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '大图标
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '状态栏
            Call Menu_View_StatusBar_click(control)
            
''--------------------------帮助-----------------
'        Case conMenu_Help_Help
'            Call Menu_Help_Help_click
'        Case conMenu_Help_Web_Forum
'            Call Menu_Help_Web_Forum_click
'        Case conMenu_Help_Web_Home
'            Call Menu_Help_Web_Home_click
'        Case conMenu_Help_Web_Mail
'            Call Menu_Help_Web_Mail_click
'        Case conMenu_Help_About
'            Call Menu_Help_About_click
    End Select
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle

    Select Case control.Id
        Case TMenuType.mtSave       '保存
            If mblnClose Then
                control.Enabled = IsEdit
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtCancel     '关闭
            control.Enabled = mblnClose
        Case TMenuType.mtImport     '导入
            control.Enabled = Not mblnClose
        Case TMenuType.mtExport     '导出
            control.Enabled = Not mblnClose
        Case TMenuType.mtNewScheme  '新增
            control.Enabled = Not mblnClose
        Case TMenuType.mtModifyScheme   '修改
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
        Case TMenuType.mtDelScheme      '删除
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
        Case TMenuType.mtRecover
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
        Case TMenuType.mtUseScheme
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
            If IsSelectionRow(vsfMain) Then
                control.Caption = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct使用状态) = "启用", "禁用", "启用")
                control.IconId = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct使用状态) = "启用", 211, 207)
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtUsually
            If IsSelectionRow(vsfMain) Then
                control.Enabled = Not (mblnClose Or vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct使用状态) = "禁用")
                control.Caption = IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否常用)) = 0, "常     用", "取消常用")
                control.IconId = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否常用) = M_STR_CROOK, 212, 204)
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtSetDefault

            If IsSelectionRow(vsfMain) Then
                control.Enabled = Not (mblnClose Or vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct使用状态) = "禁用")
                control.Caption = IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否默认)) = 0, "默     认", "取消默认")
                control.IconId = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否默认) = M_STR_CROOK, 213, 205)
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtCheckScheme    '测试
            control.Enabled = Not (vsfMain.Rows <= 1)
        Case TMenuType.mtMoveLastScheme     '上移
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
        Case TMenuType.mtMoveNextScheme     '下移
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
    End Select

    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkCard_Click()
    On Error GoTo errHandle
    
    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkGroup_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkHistory_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkTrance_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    Me.ZOrder
    Call InitCommandBars
    Call InitDockPannel
    Call InitTabControl

    Call GridInit(M_STR_GRIDDATA, vsfMain)
    Call RefreshWindowState(False)
    Call RefreshList
    Call RefreshScheme
    Call ShowScheme
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub


Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '设置菜单栏和工具栏风格
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True                                '显示按钮提示
        .AlwaysShowFullMenus = False                            '不常用的菜单项先隐藏
        .UseFadedIcons = False                                  '图标显示为褪色效果
        .IconsWithShadow = True                                 '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True                                '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True                                      '工具栏显示为大图标
        .SetIconSize True, 24, 24                               '设置大图标的尺寸
        .SetIconSize False, 16, 16                              '设置小图标的尺寸
        .ShowTextBelowIcons = True
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                        '设置控件显示风格
        .EnableCustomization False                               '是否允许自定义设置
        Set .Icons = imgMain.Icons                               '设置关联的图标控件
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '菜单定义
'Begin------------------------编辑菜单--------------------------------------默认可见
    cbrMain.ActiveMenuBar.Title = "菜单"
    
    
        Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtFile, "文件(&F)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSave, "保存(&S)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCancel, "关闭(&C)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtImport, "导入(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtExport, "导出(&E)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtQuit, "退出(&Q)"): cbrControl.BeginGroup = True
    End With
        
        Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtEdit, "编辑(&E)")
        
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtNewScheme, "新增(&N)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtModifyScheme, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelScheme, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSetDefault, "默认(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUsually, "常用(&Y)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtRecover, "恢复(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUseScheme, "启用(&A)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCheckScheme, "测试(&V)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveLastScheme, "上移(&L)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveNextScheme, "下移(&X)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSetSysQuery, "用户查询调整(&U)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtResource, "资源管理(&Z)")
    End With
    
    
    'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mtViewPopup, "查看(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)
    
'    'Begin----------------------帮助菜单--------------------------------------默认可见
'    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mtHelpPopup, "帮助(H)")
'    Call CreateViewAndHelpMenu(Nothing, cbrMenuBar)
    
    
    '---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSave, "保存", "保存方案")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCancel, "关闭", "关闭编辑")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtNewScheme, "新增", "新增方案"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtModifyScheme, "修改", "修改方案")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelScheme, "删除", "删除方案")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSetDefault, "默认", "设置默认方案"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUsually, "常用", "常用方案")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtRecover, "恢复", "恢复方案")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUseScheme, "启用", "禁用方案")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveLastScheme, "上移", "上移方案"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveNextScheme, "下移", "下移方案")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCheckScheme, "测试", "测试方案"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtQuit, "退出", "退出"): cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub CreateViewAndHelpMenu(objViewMenu As Object, objHelpMenu As Object, _
    Optional ByVal strMenuTag As String = "")
    
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    
    
    'Begin----------------------查看菜单--------------------------------------
    If Not (objViewMenu Is Nothing) Then
        Set cbrMenuBar = objViewMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(T)")
                cbrControl.Category = strMenuTag
            
                With cbrControl.CommandBar '二级菜单
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(0)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(1)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                End With
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(S)")
                cbrControl.Checked = True
                        cbrPopControl.Category = strMenuTag
        End With
    End If

'    'Begin----------------------帮助菜单--------------------------------------默认可见
'    If Not (objHelpMenu Is Nothing) Then
'        Set cbrMenuBar = objHelpMenu
'
'        With cbrMenuBar.CommandBar
'            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_Help, "帮助主题(M)")
'                cbrControl.Category = strMenuTag
'
'            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB上的中联(W)")
'                cbrControl.Category = strMenuTag
'
'                With cbrControl.CommandBar
'                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(0)")
'                        cbrPopControl.Category = strMenuTag
'
'                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Home, "中联主页(1)")
'                        cbrPopControl.Category = strMenuTag
'
'                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(2)")
'                        cbrPopControl.Category = strMenuTag
'                End With
'
'            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "关于…(A)")
'                cbrControl.Category = strMenuTag
'        End With
'    End If
End Sub

'布局绑定
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo errHandle
    
    Select Case Item.Id
        Case 1
            Item.Handle = picSchemeName.hwnd
        Case 2
            Item.Handle = picSchemeContent.hwnd
    End Select
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

'界面布局
Private Sub InitDockPannel()
    Dim objPane As Pane
    
    On Error GoTo errHandle

    dkpMain.SetCommandBars cbrMain
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "picSchemeName"
    objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, objPane)
    objPane.Title = "picSchemeContent"
    objPane.Options = PaneNoCaption
    
    Set objPane = Nothing
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub picSchemeContent_Resize()
    On Error Resume Next

    fraBasic.Move picSchemeContent.Left, picSchemeContent.Top - 100, picSchemeContent.Width
    tbcMain.Move picSchemeContent.Left, picSchemeContent.Top + fraBasic.Height - 100, picSchemeContent.Width, picSchemeContent.Height - fraBasic.Height + IIf(Not stbThis.Visible, stbThis.Height, 0) + 100
    fraLine.Move fraBasic.Left, fraBasic.Top + fraBasic.Height / 2 + 200, fraBasic.Width
    txtSchemeMemo.Move txtSchemeMemo.Left, txtSchemeMemo.Top, fraBasic.Width - txtSchemeMemo.Left - 500
End Sub

Private Sub picSchemeName_Resize()
    On Error Resume Next

    vsfMain.Move picSchemeName.Left, picSchemeName.Top, picSchemeName.Width, picSchemeName.Height - IIf(stbThis.Visible, stbThis.Height, 0)
End Sub

Private Sub InitTabControl()
    With tbcMain
        With .PaintManager
            .BoldSelected = True
            .Appearance = xtpTabAppearancePropertyPage2003
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameNone
            .Position = xtpTabPositionTop
            .OneNoteColors = False
            .BoldSelected = True
            .ColorSet.ButtonSelected = &HFFC0C0
            .ColorSet.ButtonNormal = &HE0E0E0
            .Layout = xtpTabLayoutAutoSize
            .ButtonMargin.Top = 3
            .ButtonMargin.Bottom = 4
            .ShowIcons = True
        End With
        .InsertItem 0, "基础查询设置", mobjQuerySet.hwnd, 0
        .Item(.ItemCount - 1).Tag = "基础查询设置"
        .InsertItem 1, "查询过滤设置", mobjFilterSet.hwnd, 0
        .Item(.ItemCount - 1).Tag = "查询过滤设置"
        .InsertItem 2, "数据显示设置", mobjDisPlaySet.hwnd, 0
        .Item(.ItemCount - 1).Tag = "数据显示设置"
        .Item(0).Selected = True
    End With
    
End Sub

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim strResult As String
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As Recordset
    Dim strItem As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    If Item.Tag = "查询过滤设置" Or Item.Tag = "数据显示设置" Then
        If Len(mobjQuerySet.GetQuerySql) = 0 Then
            MsgBox "请先输入查询语句。", vbInformation, Me.Caption
            tbcMain.Item(0).Selected = True
            Exit Sub
        End If
    
        strResult = SqlVerify(mobjQuerySet.GetQuerySql)
        
        If Len(strResult) = 0 Then
            strResult = IsHaveID(mobjQuerySet.GetQuerySql)
        End If
        
        If Len(strResult) > 0 Then
            MsgBox "查询语句验证失败，" & vbCrLf & "原因是：" & strResult, vbInformation, Me.Caption
            tbcMain.Item(0).Selected = True
            Call mobjQuerySet.rtbCheckSQLSetFocus
            Exit Sub
        End If
        
        Call RefreshShowScheme
    End If

    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub RefreshList()
'表格数据刷新
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset

    vsfMain.Clear
    strSql = "select '' 序号,ID,方案名称,使用状态,是否默认,是否常用,方案说明 from 影像查询方案 where 所属模块 = [1] Order By 方案序号"
    Set rsData = ExecuteSql(strSql, "查询方案信息", mlngModuleNo)
    Set vsfMain.DataSource = rsData
    vsfMain.ColHidden(ColTitle.ctID) = True
    
    Call DataConvert
    Call SchemeNo

    vsfMain.ColWidth(ColTitle.ct序号) = 500
    vsfMain.ColWidth(ColTitle.ct方案名称) = 2000
End Sub


Private Sub DataConvert()
'表格数据转换
    Dim i As Long
    
    If vsfMain.Rows <= 1 Then
        Exit Sub
    End If
    
    For i = 1 To vsfMain.Rows - 1
        If Val(vsfMain.TextMatrix(i, ColTitle.ct使用状态)) = 0 Then
            vsfMain.TextMatrix(i, ColTitle.ct使用状态) = "禁用"
        ElseIf Val(vsfMain.TextMatrix(i, ColTitle.ct使用状态)) = 1 Then
            vsfMain.TextMatrix(i, ColTitle.ct使用状态) = "启用"
        End If
        
        If Val(vsfMain.TextMatrix(i, ColTitle.ct是否常用)) = 0 Then
            vsfMain.TextMatrix(i, ColTitle.ct是否常用) = ""
        ElseIf Val(vsfMain.TextMatrix(i, ColTitle.ct是否常用)) = 1 Then
            vsfMain.TextMatrix(i, ColTitle.ct是否常用) = M_STR_CROOK
        End If
        If Val(vsfMain.TextMatrix(i, ColTitle.ct是否默认)) = 0 Then
            vsfMain.TextMatrix(i, ColTitle.ct是否默认) = ""
        ElseIf Val(vsfMain.TextMatrix(i, ColTitle.ct是否默认)) = 1 Then
            vsfMain.TextMatrix(i, ColTitle.ct是否默认) = M_STR_CROOK
        End If
    Next
End Sub

Private Sub txtDate_Change()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And Chr(KeyAscii) <> vbBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPage_Change()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub txtPage_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And Chr(KeyAscii) <> vbBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSchemeMemo_Change()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub txtSchemeMemo_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If Len(txtSchemeMemo.Text) > 511 And Chr(KeyAscii) <> vbBack Then KeyAscii = 0
End Sub

Private Sub txtSchemeName_Change()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub txtSchemeName_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If Len(txtSchemeName.Text) > 29 And Chr(KeyAscii) <> vbBack Then KeyAscii = 0
End Sub

Private Sub vsfMain_DblClick()
    On Error GoTo errHandle
    
    If vsfMain.Rows <= 1 Then Exit Sub
    Call ModifyScheme
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub vsfMain_RowColChange()
    On Error GoTo errHandle
    
    If vsfMain.Row <> mlngRow Then
        mlngRow = vsfMain.Row
        Call ShowScheme
        Call RefreshShowScheme
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub NewScheme()
'新增方案
    mblnNewState = True
    mblnClose = True
    Call RefreshWindowState(True)
    Call RefreshSubWindowState(True)
    Call NewRow(vsfMain)
    Call InitScheme
    Call SetNoneEdit
End Sub

Private Sub ModifyScheme()
'修改方案

    mblnNewState = False
    mblnClose = True
    Call RefreshWindowState(True)
    Call RefreshSubWindowState(True)
    Call SetNoneEdit
End Sub

Private Sub SaveScheme()
    Dim strSql As String
    Dim strText As String
    Dim rsData As Recordset

    If Not IsEnabledToSvae Then Exit Sub
    strText = GetSchemeContent
    If mblnNewState Then
        strSql = "Zl_影像查询_新增方案('" & Replace(txtSchemeName.Text, "'", "''") & "','" & Replace(txtSchemeMemo.Text, "'", "''") & "','" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否默认) & "','" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct使用状态) & "','" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否常用) & "','" & mlngModuleNo & "','" & Replace(strText, "'", "''") & "')"
    Else
        strSql = "Zl_影像查询_更新方案(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & ",'" & Replace(txtSchemeName.Text, "'", "''") & "','" & Replace(txtSchemeMemo.Text, "'", "''") & "'," & IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否默认)) = 0, 0, 1) & "," & IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct使用状态) = "启用", 1, 0) & "," & IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否常用)) = 0, 0, 1) & "," & mlngModuleNo & ",'" & Replace(strText, "'", "''") & "')"
    End If
    Call ExecuteCmd(strSql, "编辑方案")
    If mblnNewState Then
        strSql = "select ID from 影像查询方案 where 方案名称 = [1] and 所属模块 = [2]"
        Set rsData = ExecuteSql(strSql, "查询方案信息", txtSchemeName.Text, mlngModuleNo)
        If rsData.RecordCount < 1 Then
            Exit Sub
        End If
        vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) = rsData.Fields!Id

    End If

    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct方案名称) = txtSchemeName.Text
    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct方案说明) = txtSchemeMemo.Text
    vsfMain.RowData(vsfMain.Row) = strText

    Call ShowScheme
    Call RefreshShowScheme
    If mblnNewState Then
        Call SetUseScheme
        Call SchemeNo
    End If

    mblnNewState = False
    Call SetNoneEdit

End Sub

Private Function GetSchemeContent() As String
    Dim objSqlScheme As clsSqlScheme
    Dim objScSearchCfg As clsScSerachCfg
    Dim strText As String
    Dim strQuery As String
    Dim strDetail As String


    Set objSqlScheme = New clsSqlScheme

    '基本信息
    With objSqlScheme
        If Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)) > 0 Then
            .SchemeId = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)
        End If
        .UseCard = chkCard.value
        .UseGroup = chkGroup.value
        .ShowHistory = chkHistory.value
        .UseFuncFollow = chkTrance.value
        .PageRecord = Val(txtPage.Text)
        .DataRange = Val(txtDate.Text)
        .SchemeName = txtSchemeName
        .Descript = txtSchemeMemo
        .Store = mobjSqlScheme.Store
    End With


    '查询语句模块
    Call mobjQuerySet.SetQueryCfg(objSqlScheme)

    '录入设置模块
    Call mobjFilterSet.SetConditionCfg(objSqlScheme)

    Call mobjDisPlaySet.SetShowCfg(objSqlScheme)

    strText = objSqlScheme.GetScheme

    GetSchemeContent = strText

    Set objSqlScheme = Nothing
    Set objScSearchCfg = Nothing
End Function

Private Sub MoveLastScheme()
'上移方案
    Dim strSql As String

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If

    Call MoveUp(vsfMain)

    strSql = "zl_影像查询_移动方案(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & vsfMain.Row & "," & mlngModuleNo & ")"
    Call ExecuteCmd(strSql, "上移方案")
    Call SchemeNo
End Sub



Private Sub MoveNextScheme()
'下移方案
    Dim strSql As String

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If

    Call MoveDown(vsfMain)

    strSql = "zl_影像查询_移动方案(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & vsfMain.Row & "," & mlngModuleNo & ")"
    Call ExecuteCmd(strSql, "下移方案")
    Call SchemeNo
End Sub

Private Sub SetDefaultScheme()
'设置默认方案
    Dim strSql As String
    Dim strCurDefaultState As String
    Dim i As Long

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strCurDefaultState = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否默认)
    strSql = "Zl_影像查询_默认方案(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & IIf(strCurDefaultState = M_STR_CROOK, 0, 1) & "," & mlngModuleNo & ")"
    Call ExecuteCmd(strSql, "配置默认")
    
    vsfMain.Cell(flexcpText, 1, ColTitle.ct是否默认, vsfMain.Rows - 1, ColTitle.ct是否默认) = ""
    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否默认) = IIf(Len(strCurDefaultState) = 0, M_STR_CROOK, "")
'    cbrMain.RecalcLayout
End Sub

Private Sub SetUsualScheme()
'设置是否常用
    Dim strSql As String
    Dim strCurUsualState As String
    Dim i As Long

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strCurUsualState = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否常用)
    strSql = "Zl_影像查询_常用方案(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & IIf(strCurUsualState = M_STR_CROOK, 0, 1) & ")"
    Call ExecuteCmd(strSql, "配置常用")

    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct是否常用) = IIf(Len(strCurUsualState) = 0, M_STR_CROOK, "")
'    cbrMain.RecalcLayout
End Sub

Private Sub SetUseScheme()
'启用方案
    Dim strSql As String
    Dim strCurUseState As String

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strCurUseState = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct使用状态)

    strSql = "Zl_影像查询_启用方案(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & IIf(strCurUseState = "启用", 0, 1) & ")"
    Call ExecuteCmd(strSql, "使用状态配置")

    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct使用状态) = IIf(strCurUseState = "启用", "禁用", "启用")

End Sub


Private Sub DeleteScheme()
'删除选中方案
    Dim strSql As String
    Dim lngRow As Long

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If

    If Not MsgBox("是否删除选中方案？", vbYesNo, Me.Caption) = vbYes Then
        Exit Sub
    End If

    lngRow = vsfMain.Row
    strSql = "Zl_影像查询_删除方案(" & vsfMain.TextMatrix(lngRow, ColTitle.ctID) & ")"
    Call ExecuteCmd(strSql, "删除方案")

    vsfMain.RemoveItem (lngRow)
End Sub

Private Sub CancelScheme()
    If IsEdit Then
        If MsgBox("方案尚未保存，是否保存？", vbYesNo, Me.Caption) = vbYes Then
            Call SaveScheme
        End If
    End If
    If mblnNewState Then
        vsfMain.RemoveItem vsfMain.Row
    End If
    Call ShowScheme
    Call RefreshShowScheme
    mblnNewState = False
    mblnClose = False
    Call RefreshWindowState(False)
    Call RefreshSubWindowState(False)
    
'    Call RefreshScheme
End Sub

Private Sub RefreshWindowState(blnState As Boolean)
    vsfMain.Enabled = Not blnState
    txtSchemeName.Enabled = blnState
    txtSchemeMemo.Enabled = blnState
    txtDate.Enabled = blnState
    txtPage.Enabled = blnState
    chkCard.Enabled = blnState
    chkGroup.Enabled = blnState
    chkHistory.Enabled = blnState
    chkTrance.Enabled = blnState
End Sub

Private Sub InitScheme()
    chkCard.value = 1
    chkHistory.value = 1
    chkGroup.value = 1
    chkTrance.value = 1
    txtPage.Text = 0
    txtDate.Text = 0
    txtSchemeName.Text = ""
    txtSchemeMemo.Text = ""

    tbcMain.Item(0).Selected = True
End Sub


Private Sub RefreshScheme()
'刷新方案
    Dim i As Long
    Dim lngNumber As Long
    Dim strSchemeXml As String
'
'    Set mobjSqlScheme = Nothing
'    If Not IsSelectionRow(vsfMain) Then
'        Exit Sub
'    End If
'    If Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)) = 0 Then
'        Exit Sub
'    End If
'
'    lngNumber = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)
'
'    strSchemeXml = ReadSchemeXml(lngNumber, "")
'
'    Call mobjSqlScheme.OpenScheme(strSchemeXml)
    If vsfMain.Rows < 2 Then Exit Sub
    For i = 1 To vsfMain.Rows - 1
        If Not Len(vsfMain.TextMatrix(i, ColTitle.ctID)) = 0 Then
            lngNumber = vsfMain.TextMatrix(i, ColTitle.ctID)
            strSchemeXml = ReadSchemeXml(lngNumber, "")
            vsfMain.RowData(i) = strSchemeXml
        End If
    Next

End Sub

Private Sub ShowScheme()
'显示方案的基本信息界面
    Call mobjSqlScheme.OpenScheme(vsfMain.RowData(vsfMain.Row))
    With mobjSqlScheme
        chkCard.value = IIf(.UseCard, 1, 0)
        chkHistory.value = IIf(.ShowHistory, 1, 0)
        chkGroup.value = IIf(.UseGroup, 1, 0)
        chkTrance.value = IIf(.UseFuncFollow, 1, 0)
        txtPage.Text = .PageRecord
        txtDate.Text = .DateRange
        txtSchemeName.Text = .SchemeName
        txtSchemeMemo.Text = .Descript

        '显示查询语句模块
        Call mobjQuerySet.ShowQuerySet(mobjSqlScheme)
        '显示快速过滤模块
        Call mobjFilterSet.ShowFilterSet(mobjSqlScheme)
        '显示显示配置模块
        Call mobjDisPlaySet.ShowDisplaySet(mobjSqlScheme)
    End With
End Sub

Private Sub RefreshShowScheme()
'刷新方案的显示
    Dim strQuerySql As String

    '获取当前查询语句
    strQuerySql = mobjQuerySet.GetQuerySql

    If tbcMain.Item(1).Selected = True Then
        Call mobjFilterSet.RefreshFilterSet(strQuerySql, mobjSqlScheme)
    ElseIf tbcMain.Item(2).Selected = True Then
        Call mobjDisPlaySet.RefreshDisplaySet(strQuerySql)
    End If
End Sub

Private Function ExeSqlTrans(strSql As String) As String
    Dim i As Long
    
    On Error GoTo errRollback

    ExeSqlTrans = ""

    gcnOracle.BeginTrans

    If Len(strSql) > 0 Then
        Call ExecuteCmd(strSql, "保存方案配置")
    End If

    gcnOracle.CommitTrans
    Exit Function
errRollback:
    gcnOracle.RollbackTrans
    ExeSqlTrans = Err.Description
End Function

Private Function IsEnabledToSvae() As Boolean
'方案检验
    Dim i As Long
    Dim strResult As String

    IsEnabledToSvae = False
    If Len(Replace(txtSchemeName.Text, " ", "")) = 0 Then
        MsgBox "方案名称为空，请输入方案名称。", vbInformation, Me.Caption
        txtSchemeName.SetFocus
        Exit Function
    End If

    '方案名不能重复
    For i = 1 To vsfMain.Rows - 2
        If vsfMain.TextMatrix(i, 2) = txtSchemeName.Text And (i <> vsfMain.Row) Then
            MsgBox "方案名已存在，请检查。", vbInformation, Me.Caption
            txtSchemeName.SetFocus
            Exit Function
        End If
    Next

    If Not mobjQuerySet.IsEnabledToSave Then
        MsgBox "请录入查询语句", vbInformation, Me.Caption
        Exit Function
    End If

    strResult = SqlVerify(mobjQuerySet.GetQuerySql)
    If Len(strResult) = 0 Then
        strResult = IsHaveID(mobjQuerySet.GetQuerySql)
    End If
    If Len(strResult) > 0 Then
        MsgBox "查询语句验证失败，原因为：" & strResult, vbInformation, Me.Caption
        Exit Function
    End If

    If Not mobjFilterSet.IsEnabledSave Then
        Exit Function
    End If

    IsEnabledToSvae = True
End Function

Private Sub RefreshSubWindowState(blnState As Boolean)
'刷新子界面可用状态
    Call mobjFilterSet.RefreshWindowState(blnState)
    Call mobjQuerySet.RefreshWindowState(blnState)
    Call mobjDisPlaySet.RefreshWindowState(blnState)
End Sub

Private Sub ExportScheme()
'导出
    Dim objExportScheme As New frmExportScheme
    Dim arrID() As Long
    Dim strFile As String
    Dim blnIcon As Boolean

    objExportScheme.ShowMe mlngModuleNo, True, arrID, strFile, blnIcon, Me
    Set objExportScheme = Nothing
End Sub

Private Sub ImportScheme()
'导入
    Dim objExportScheme As New frmExportScheme
    Dim arrID() As Long
    Dim strFile As String
    Dim blnIcon As Boolean

    dlgFile.Filter = "(*.XML)|*.XML|(*.*)|*.*"
    dlgFile.DefaultExt = "*.XML"

    dlgFile.FileName = ""
    dlgFile.ShowOpen

    If dlgFile.FileName = "" Then Exit Sub
    strFile = dlgFile.FileName

    If objExportScheme.ShowMe(mlngModuleNo, False, arrID, strFile, blnIcon, Me) Then
        Call ImportContent(arrID, strFile, blnIcon)
        Set objExportScheme = Nothing
    End If

    Call RefreshList
    Call RefreshScheme
    Call ShowScheme
End Sub


Private Sub ImportContent(arrID() As Long, strFile As String, blnIcon As Boolean)
'导入方案
    Dim rsData As ADODB.Recordset
    Dim lngOldSchemeId As Long
    Dim lngNewSchemeId As Long
    Dim strSql As String
    Dim strExeSql() As String
    Dim strResult As String
    Dim strText As String
    Dim strLog As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strSchemeName As String
    Dim lngCount As Long
    Dim lngScheme As Long
    Dim blnIsImport As Boolean
    Dim strOldName As String
    Dim blnImportIcon As Boolean
    Dim arrIcon() As String
    Dim strPath As String
    Dim strName As String
    Dim blnIsHave As Boolean
    Dim lngIconNum As Long
    Dim lngDefeated As Long
    Dim objSqlScheme As New clsSqlScheme

    If blnIcon Then
    '导入图标
        If Len(strFile) = 0 Then
            MsgBox "没有图标文件，请检查", vbInformation, Me.Caption
            Exit Sub
        End If
        strPath = Replace(strFile, ".XML", "\")
        If Len(Dir(strPath)) = 0 Then
            MsgBox "未找到图标文件，请检查", vbInformation, Me.Caption
            Exit Sub
        End If
        strName = Dir(strPath & "*.ico", 7)

        strSql = "select 资源名称 from 影像查询资源 where 资源类型 = [1]"
        Set rsData = ExecuteSql(strSql, "查询方案信息", 1)
        lngIconNum = 0
        Do
            If Len(Trim(strName)) = 0 Then Exit Do
            If Not IsHaveIcon(Replace(strName, ".ico", ""), rsData) Then
                strSql = "Zl_影像查询_新增图标('" & Replace(strName, ".ico", "") & "','1')"
                Call ExecuteCmd(strSql, "新增图标")
                Call zlBlobSave(Replace(strName, ".ico", ""), strPath & strName)
                lngIconNum = lngIconNum + 1
            Else
                strLog = "图标名称【" & Replace(strName, ".ico", "") & "】已存在，导入时已忽略该图标"
                Call LogFile(strLog)
            End If
            strName = Dir()
        Loop
    End If

'
    Set rsData = New ADODB.Recordset
    Call rsData.Open(strFile)

    If rsData.RecordCount <= 0 Then
        MsgBox "没有可用于导入的数据，请检查文件是否正确。", vbInformation, Me.Caption
        Exit Sub
    End If


    rsData.Sort = "id"

    lngScheme = 0
    lngDefeated = 0
    lngOldSchemeId = 0
    rsData.MoveFirst
    ReDim Preserve strExeSql(1)

    While Not rsData.EOF
        blnIsHave = False
        For i = 0 To UBound(arrID)
            If Val(NVL(rsData!Id)) = arrID(i) Then
                blnIsHave = True
                Exit For
            End If
        Next

        If blnIsHave Then
            lngCount = 0
            strOldName = ""
            strText = ""
            blnIsImport = True
            If lngOldSchemeId <> Val(NVL(rsData!Id)) Then
                '增加影像查询方案记录

                strSql = "select 影像查询方案_ID.NextVal as ID from dual"
                Set rsTemp = ExecuteSql(strSql, "获取新方案ID")
                If rsTemp.RecordCount <= 0 Then
                    MsgBox "不能获取新增的方案ID，系统将退出保存。", vbExclamation, Me.Caption
                    Exit Sub
                End If

                lngNewSchemeId = Val(NVL(rsTemp!Id))
                strSchemeName = NVL(rsData!方案名称)

                For i = 1 To vsfMain.Rows - 1
                    If vsfMain.TextMatrix(i, ColTitle.ct方案名称) = strSchemeName Then
                        strOldName = strSchemeName
                        If MsgBox("已存在名为：【" & strSchemeName & "】的方案,是否继续导入", vbYesNo, Me.Caption) = vbYes Then
                            Do While True
                                strSchemeName = strSchemeName & lngCount
                                If IsHaveScheme(strSchemeName) Then
                                    Exit Do
                                End If

                                lngCount = lngCount + 1
                            Loop
                        Else
                            blnIsImport = False
                        End If
                    End If
                Next
                If blnIsImport Then
                    strText = strText & rsData.Fields(3).value
                    objSqlScheme.OpenScheme strText
                    strResult = SqlVerify(objSqlScheme.Query)
                    If Len(strResult) = 0 Then
                        If Len(strOldName) > 0 Then
                            strText = Replace(strText, "name=""" & strOldName & """", "name=""" & strSchemeName & """")
                        End If
                        ReDim Preserve strExeSql(UBound(strExeSql) + 1)
                        strExeSql(UBound(strExeSql) - 1) = "zl_影像查询_新增方案('" & _
                                                                    Replace(strSchemeName, "'", "''") & "','" & _
                                                                    NVL(rsData!方案说明) & "'," & _
                                                                    "'','1','','" & mlngModuleNo & "','" & _
                                                                    Replace(NVL(strText), "'", "''") & "')"
                    Else
                        strLog = "方案【" & NVL(rsData!方案名称) & "】导入失败，原因为：方案验证失败（" & strResult & "）"
                        Call LogFile(strLog)
                        lngDefeated = lngDefeated + 1
                    End If
                End If
                lngOldSchemeId = Val(NVL(rsData!Id))
            End If
        End If

        rsData.MoveNext
    Wend

    '写入方案配置相关数据
    For i = 0 To UBound(strExeSql)
        strSql = strExeSql(i)
        If Len(strSql) > 0 Then
            strResult = ExeSqlTrans(strSql)
            
            If Len(strResult) > 0 Then
                strLog = "方案【" & Mid(strSql, InStr(strSql, "('") + 2, InStr(strSql, "',") - InStr(strSql, "('") - 2) & "】导入失败，原因为：" & strResult
                Call LogFile(strLog)
                lngDefeated = lngDefeated + 1
            Else
                lngScheme = lngScheme + 1
            End If
        End If
    Next i
    
    If blnIcon Then
        MsgBox "已导入成功" & lngScheme & "条数据，失败" & lngDefeated & "条数据，" & lngIconNum & "个图标资源。", vbInformation, Me.Caption
    Else
        MsgBox "已导入成功" & lngScheme & "条数据，失败" & lngDefeated & "条数据。", vbInformation, Me.Caption
    End If
    If lngDefeated > 0 Then
        ShellExecute Me.hwnd, "open", App.Path & "\" & "SchemeImport" & ".log", "", vbNullString, SW_SHOW
    End If
End Sub

Private Sub LogFile(ByVal strInfo As String)
    Dim lngFileNum As Long
    Dim FilePath As String
    Dim objFSO As Object
    Dim objLogFile As Object
    
    FilePath = App.Path & "\" & "SchemeImport" & ".log"

    lngFileNum = FreeFile
 
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Len(Dir(FilePath)) = 0 Then
        objFSO.CreateTextFile FilePath, True
    End If
    Set objLogFile = objFSO.GetFile(FilePath)
    If objLogFile = Empty Then
        Open FilePath For Output As #lngFileNum
    Else
        If objLogFile.Size > 2097152 Then
            objLogFile.Copy App.Path & "\" & App.EXEName & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".log"
            Open FilePath For Output As #lngFileNum
        Else
            Open FilePath For Append As #lngFileNum
        End If
    End If
 
    Print #lngFileNum, CStr(Now()) & ": " & strInfo
    Close #lngFileNum
 
End Sub

Private Sub UnloadMe()
'退出
    If cbrMain.FindControl(, TMenuType.mtSave).Enabled Then
        If MsgBox("尚未保存方案，是否退出？", vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If

    mobjFilterSet.UnloadMe
    mobjDisPlaySet.UnloadMe
    mobjSetRelated.UnloadMe
    
    Set mobjQuerySet = Nothing
    Set mobjFilterSet = Nothing
    Set mobjDisPlaySet = Nothing
    Set mobjSqlScheme = Nothing
    Set mobjSetRelated = Nothing
    Set mobjIconManage = Nothing

    Unload Me
End Sub

Private Sub RecoverScheme()
    Dim strStore As String

    If MsgBox("是否确定恢复方案？", vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    strStore = mobjSqlScheme.Store
    If Len(strStore) < 1 Then
        MsgBox "该方案没有设置恢复属性，无法恢复", vbInformation, Me.Caption
        Exit Sub
    End If
    Call mobjSqlScheme.OpenScheme(strStore)
    mobjSqlScheme.Store = strStore

    vsfMain.RowData(vsfMain.Row) = strStore
    Call ShowScheme
    Call SaveScheme
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
'查看
    Dim i As Integer

    On Error GoTo errHandle

    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
'按钮
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    On Error GoTo errHandle

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If

        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
'图标
    On Error GoTo errHandle

    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked

    Me.cbrMain.RecalcLayout
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
'状态栏
    On Error GoTo errHandle

    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    picSchemeName.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    picSchemeContent.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    Me.cbrMain.RecalcLayout
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub
'
'Private Sub Menu_Help_Help_click()
''功能：调用帮助主题
'    On Error GoTo errHandle
'
'    ShowHelp App.ProductName, Me.hWnd, Me.Name
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbExclamation, Me.Caption
'    Err.Clear
'End Sub
'
'Private Sub Menu_Help_Web_Forum_click()
'    On Error GoTo errHandle
'
'    Call zlWebForum(Me.hWnd)
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbExclamation, Me.Caption
'    Err.Clear
'End Sub
'
'Private Sub Menu_Help_Web_Home_click()
'    On Error GoTo errHandle
'
'    zlHomePage hWnd
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbExclamation, Me.Caption
'    Err.Clear
'End Sub
'
'Private Sub Menu_Help_Web_Mail_click()
'    On Error GoTo errHandle
'
'    zlMailTo hWnd
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbExclamation, Me.Caption
'    Err.Clear
'End Sub
'
'Private Sub Menu_Help_About_click()
'    On Error GoTo errHandle
'
'    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbExclamation, Me.Caption
'    Err.Clear
'End Sub

Private Function IsHaveScheme(strName As String) As Boolean
    Dim i As Long

    IsHaveScheme = False
    For i = 1 To vsfMain.Rows - 1
        If UCase(Trim(vsfMain.TextMatrix(i, ColTitle.ct方案名称))) = UCase(Trim(strName)) Then
            Exit Function
        End If
    Next
    IsHaveScheme = True
End Function

Private Sub SchemeNo()
'调整方案序号
    Dim i As Long

    If vsfMain.Rows < 2 Then Exit Sub
    For i = 1 To vsfMain.Rows - 1
        vsfMain.TextMatrix(i, ColTitle.ct序号) = i
    Next
End Sub

Private Function IsEdit() As Boolean
'判断方案内容是否发生改变
    IsEdit = False
    
    If mblnClose Then
        If mblnIsEdit Or mobjQuerySet.mblnIsEdit Or mobjDisPlaySet.mblnIsEdit Or mobjFilterSet.mblnIsEdit Then
            IsEdit = True
        End If
    End If
End Function

Private Sub SetNoneEdit()
    mblnIsEdit = False
    mobjQuerySet.mblnIsEdit = False
    mobjDisPlaySet.mblnIsEdit = False
    mobjFilterSet.mblnIsEdit = False
End Sub

Private Function IsHaveIcon(strName As String, rsRecord As Recordset) As Boolean
    IsHaveIcon = False
    If rsRecord.RecordCount < 1 Then
        IsHaveIcon = False
        Exit Function
    End If
    rsRecord.MoveFirst
    Do While Not rsRecord.EOF
        If UCase(Trim(strName)) = UCase(Trim(NVL(rsRecord.Fields!资源名称))) Then
            IsHaveIcon = True
            Exit Function
        End If
        rsRecord.MoveNext
    Loop
End Function

Public Function ShowUserScheme() As Boolean
'显示用户常用查询方案配置...
'如果有调整返回true，没有调整返回false
    Dim objQueryCfg As New frmUserQueryReleation

    On Error GoTo errHandle

    ShowUserScheme = objQueryCfg.ShowUserScheme(Me, mlngModuleNo, 0)
    Exit Function
errHandle:
    Set objQueryCfg = Nothing
    Err.Raise -1, "clsPacsQuery.ShowUserScheme", "用户查询方案关联调用失败:" & Err.Description
End Function

