VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMediPrice 
   Caption         =   "药品调价管理"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11910
   Icon            =   "frmMediPrice.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      ScaleHeight     =   255
      ScaleWidth      =   1935
      TabIndex        =   8
      Top             =   6240
      Width           =   1935
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblExecute 
         AutoSize        =   -1  'True
         Caption         =   "已生效"
         Height          =   180
         Left            =   1320
         TabIndex        =   12
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblNotExecute 
         AutoSize        =   -1  'True
         Caption         =   "未生效"
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   30
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "查阅(&V)"
      Height          =   350
      Left            =   7560
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin XtremeSuiteControls.TabControl TabDetails 
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   1815
      _Version        =   589884
      _ExtentX        =   3201
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7680
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediPrice.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15240
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "当前大写键状态"
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   885
      Left            =   3000
      TabIndex        =   3
      Top             =   1680
      Width           =   4935
      _cx             =   8705
      _cy             =   1561
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
      BackColorSel    =   16769992
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediPrice.frx":70E6
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPrice 
      Height          =   975
      Left            =   3120
      TabIndex        =   4
      Top             =   4320
      Width           =   3015
      _cx             =   5318
      _cy             =   1720
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
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediPrice.frx":715B
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsfCost 
      Height          =   975
      Left            =   6840
      TabIndex        =   5
      Top             =   4440
      Width           =   3135
      _cx             =   5530
      _cy             =   1720
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
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediPrice.frx":71D0
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1080
      MousePointer    =   7  'Size N S
      ScaleHeight     =   255
      ScaleWidth      =   7455
      TabIndex        =   6
      Top             =   3360
      Width           =   7455
      Begin VB.Label lblScope 
         Caption         =   "日期范围：2012年11月1日至2012年11月31日"
         Height          =   180
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3615
      End
   End
   Begin XtremeCommandBars.ImageManager imgList 
      Left            =   2520
      Top             =   1200
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMediPrice.frx":7245
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMediPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mconMenu_FilePopup As Long = 1 '文件
Private Const mconMenu_ReportPopup As Long = 2 '报表
Private Const mconMenu_EditPopup As Long = 3 '编辑
Private Const mconMenu_ViewPopup As Long = 4 '查看
Private Const mconMenu_HelpPopup As Long = 5 '帮助

'文件
Private Const mconMenu_File_PrintSet = 100           '*打印设置(&S)…
Private Const mconMenu_File_Preview = 101            '*预览(&V)
Private Const mconMenu_File_Print = 102              '*打印(&P)
Private Const mconMenu_File_BillPrint = 103 '单据打印（&B）
Private Const mconMenu_File_BillPreview = 104 '单据预览（&L）
Private Const mconMenu_File_Excel = 105              '输出到&Excel…
Private Const mconMenu_File_Parameter = 106 '参数设置(&R)
Private Const mconMenu_File_Exit = 107 '退出(&E)
'编辑
Private Const mconMenu_Edit_Add = 200 '增加(&A)
Private Const mconMenu_Edit_Update = 201 '修改(&U)
Private Const mconMenu_Edit_Delete = 202 '删除(&D)
Private Const mconMenu_Edit_BatchPrice = 203 '批量执行调价(&B)
'查看
Private Const mconMenu_View_Filter = 300 '过滤(&F)
Private Const mconMenu_View_Refresh = 301 '刷新(&R)
'帮助
Private Const mconMenu_Help_Title = 400 '帮助主题(&H)
Private Const mconMenu_Help_Web = 401 'web上中联
Private Const mconMenu_Help_web_WebHome = 402 '中联主页(&H)
Private Const mconMenu_Help_web_WebForum = 403 '中联论坛(&F)
Private Const mconMenu_Help_web_WebMail = 404 '发送反馈(&K)
Private Const mconMenu_Help_About = 405 '关于(&A)

Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数
Private mstrMoneyFormat As String
Private mintSalePriceDigit As Integer
Private Const MStrCaption As String = "药品调价管理"

Private mlngForeColor As Long '记录当前单元格前景色

Private mintUnit As Integer     '用来记录启用的是什么单位

Private Type Type_Condition '过滤时设置的日期
    date填制时间开始 As Date
    date填制时间结束 As Date
    date执行时间开始 As Date
    date执行时间结束 As Date
End Type
Private mSQLCondition As Type_Condition


Private mstrResult As String '过滤结果
'日期设置
Private mdaStart As Date
Private mdaEnd As Date
Private mdaVerifyStart As Date
Private mdaVerifyEnd As Date
Private mstrPrivs As String

'各种全局变量
Private Const mconlngRowHeight As Long = 300 '表格中各列行高
Private mblnLoad As Boolean     '是否加载完成

'调价汇总表
Private Enum mEnuListCol
    调价号 = 1
    调价类型
    填制人
    填制日期
    执行日期
    说明
    总列数
End Enum
'售价调价表
Private Enum menuPriceCol
    NO = 1
    原价id
    药品信息
    规格
    单位
    单位系数
    原价
    现价
    执行日期
    调价人
    总列数
End Enum
'成本价调价
Private Enum mEnuCostCol
    NO = 1
    药品信息
    库房
    规格
    批号
    产地
    单位
    原价
    现价
    效期
    执行日期
    调价人
    总列数
End Enum

Private Sub initCommandBars()
    With CommandBarsGlobalSettings
        .App = App
        .CompanyName = "重庆中联信息产业有限责任公司" '公司名称
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '设置中文语言资源文件
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '控件整体的颜色方案
    End With

    With cbsMain.Options
        .ShowExpandButtonAlways = False '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True '显示按钮提示
        .AlwaysShowFullMenus = False '不常用的菜单项先隐藏
        .UseFadedIcons = True '图标显示为褪色效果
        .IconsWithShadow = True '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True '工具栏显示为大图标
        .SetIconSize True, 24, 24 '设置大图标的尺寸
        .SetIconSize False, 16, 16 '设置小图标的尺寸
    End With

    With cbsMain
        .VisualTheme = xtpThemeOffice2003 '设置控件显示风格
        .EnableCustomization False '是否允许自定义设置
        Set .Icons = imgList.Icons '设置关联的图标控件
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap '窗体变化时，如果显示不完菜单也不换行
        .ActiveMenuBar.Title = "菜单"
    End With
    
End Sub

Private Sub initMenu()
'创建菜单
    Dim cbrMenuPopup As CommandBarPopup
    Dim cbrMenuControl As CommandBarControl

    With cbsMain
        '文件
        Set cbrMenuPopup = .ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "文件(&F)")
        cbrMenuPopup.id = mconMenu_FilePopup
        With cbrMenuPopup.CommandBar.Controls
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_PrintSet, "打印设置(&S)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_Preview, "打印预览(&V)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_Print, "打印(&P)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_BillPrint, "单据打印(&B)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_BillPreview, "单据预览(&L)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_Excel, "输出到Excel...")
            cbrMenuControl.BeginGroup = True
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_Parameter, "参数设置(&R)")
            cbrMenuControl.BeginGroup = True
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_Exit, "退出(&E)")
        End With
        '报表
        Set cbrMenuPopup = .ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ReportPopup, "报表(&R)")
        cbrMenuPopup.id = mconMenu_ReportPopup
'        cbrMenuPopup.Visible = False
        
        '编辑
        Set cbrMenuPopup = .ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "编辑(&E)")
        cbrMenuPopup.id = mconMenu_EditPopup
        With cbrMenuPopup.CommandBar.Controls
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Edit_Add, "新增(&A)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Edit_Update, "修改(&U)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Edit_Delete, "删除(&D)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Edit_BatchPrice, "零差价批量调价(&B)")
            If gtype_UserSysParms.P275_零差价管理模式 <> 0 Then
                cbrMenuControl.Visible = True
            Else
                cbrMenuControl.Visible = False
            End If
            cbrMenuControl.BeginGroup = True
        End With
        '查看
        Set cbrMenuPopup = .ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "查看(&V)")
        cbrMenuPopup.id = mconMenu_ViewPopup
        With cbrMenuPopup.CommandBar.Controls
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_View_Filter, "过滤(&F)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_View_Refresh, "刷新(&R)")
        End With
        '帮助
        Set cbrMenuPopup = .ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "帮助(&H)")
        cbrMenuPopup.id = mconMenu_HelpPopup
        With cbrMenuPopup.CommandBar.Controls
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Help_Title, "帮助主题(&H)")
            Set cbrMenuControl = .Add(xtpControlPopup, mconMenu_Help_Web, "web上中联")
            cbrMenuControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_web_WebHome, "中联主页(&H)", -1, False
            cbrMenuControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_web_WebForum, "中联论坛(&F)", -1, False
            cbrMenuControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_web_WebMail, "发送反馈(&K)", -1, False
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Help_About, "关于(&A)")
            cbrMenuControl.BeginGroup = True
        End With
    End With
    
End Sub

Private Sub InitTool()
    '创建工具栏
    Dim cbrToolBar As CommandBar
    Dim cbrMenuPopup As CommandBarPopup
    Dim cbrMenuControl As CommandBarControl

    Set cbrToolBar = cbsMain.Add("工具栏", xtpBarTop)
    With cbrToolBar
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_File_Preview, "预览")
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_File_Print, "打印")
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_Edit_Add, "新增")
        cbrMenuControl.BeginGroup = True
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_Edit_Update, "修改")
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_Edit_Delete, "删除")
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_View_Filter, "过滤")
        cbrMenuControl.BeginGroup = True
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_View_Refresh, "刷新")
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_Help_Title, "帮助")
        cbrMenuControl.BeginGroup = True
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_File_Exit, "退出")
    End With

    For Each cbrMenuControl In cbrToolBar.Controls  '让工具栏中按钮同时显示图标和文字
        cbrMenuControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitTabControl()
    '初始化TabControl控件
    Dim objtabctl As TabControlItem

    picSplit.Left = 0
    picSplit.Top = vsfList.Top + vsfList.Height + 400
    With TabDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem 1, "售价调价", vsfPrice.hWnd, 0
        .InsertItem 2, "成本价调价", vsfCost.hWnd, 0
        .Top = picSplit.Top + picSplit.Height + 20
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - picSplit.Top - picSplit.Height - stbThis.Height
        .Item(0).Selected = True
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intMethod As Integer
    Dim blnPrivs As Boolean
    
    Select Case Control.id
    Case mconMenu_Edit_Add '新增
        frmMediPriceCard.ShowME Me, 0, "", 0
        Call getListInfo
        vsfList.SetFocus
    Case mconMenu_Edit_Update '修改
        If vsfList.rows = 1 Then Exit Sub
        blnPrivs = CheckPrivs(intMethod)
        If blnPrivs = True Then
            frmMediPriceCard.ShowME Me, 1, vsfList.TextMatrix(vsfList.Row, mEnuListCol.调价号), intMethod
            Call getListInfo
            vsfList.SetFocus
        Else
            MsgBox "操作员不具备" & IIf(intMethod = 0, "售价调价", IIf(intMethod = 1, "成本价调价", "售价成本价一起调价")) & "权限，请与管理员联系！", vbInformation, gstrSysName
            Exit Sub
        End If
    Case mconMenu_Edit_Delete '删除
        If vsfList.rows = 1 Then Exit Sub
        blnPrivs = CheckPrivs(intMethod)
        If blnPrivs = True Then
            Call deleteNotExecutePirce
        Else
            MsgBox "操作员不具备" & IIf(intMethod = 0, "售价调价", IIf(intMethod = 1, "成本价调价", "售价成本价一起调价")) & "权限，请与管理员联系！", vbInformation, gstrSysName
            Exit Sub
        End If
    Case mconMenu_Edit_BatchPrice '批量调价
        frmMediPriceDiffCard.Show vbModal, Me
        Call getListInfo
        vsfList.SetFocus
    Case mconMenu_File_Exit '退出
        Unload Me
    Case mconMenu_View_Refresh '刷新
        Call getListInfo
        vsfList.SetFocus
    Case mconMenu_View_Filter '过滤
        frmMediPriceSearch.ShowME Me, mstrResult, mSQLCondition.date填制时间开始, mSQLCondition.date填制时间结束, mSQLCondition.date执行时间开始, mSQLCondition.date执行时间结束
        Call getListInfo
        vsfList.SetFocus
    Case mconMenu_File_Parameter '参数设置
        frm参数设置.设置参数 Me, mstrPrivs, MStrCaption
        Call initJinDu
        Call getListInfo
    Case mconMenu_File_PrintSet '打印设置
        Call zlPrintSet
    Case mconMenu_File_Preview '打印预览
        Call PrintView
    Case mconMenu_File_Print '打印
        Call filePrint
    Case mconMenu_File_BillPrint '单据打印
        Call BillPrint(2)
    Case mconMenu_File_BillPreview '单据预览
        Call BillPrint(1)
    Case mconMenu_File_Excel '输出到Excel
        Call billExcel
    Case mconMenu_Help_About    '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case mconMenu_Help_Title '帮助主题
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case mconMenu_Help_web_WebHome '中联主页
        Call zlHomePage(Me.hWnd)
    Case mconMenu_Help_web_WebForum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case mconMenu_Help_web_WebMail '发送反馈
        Call zlMailTo(Me.hWnd)
    Case Else '报表
        Call vsfPrint_Custom(Control)
    End Select
End Sub

Private Function CheckPrivs(ByRef intMethod As Integer) As Boolean
    '功能：判断是否具有对应操作的权限
    '返回值：ture-具有对应操作权限，false-不具有对应操作权限
    '传参：返回操作类型 0-售价调价，1-成本价调价 2-一起调整
    With vsfList
        If .TextMatrix(vsfList.Row, mEnuListCol.调价类型) = "仅售价调价" Then
            intMethod = 0
            If InStr(1, ";" & mstrPrivs & ";", ";售价调价;") > 0 Then CheckPrivs = True
        ElseIf .TextMatrix(vsfList.Row, mEnuListCol.调价类型) = "仅成本价调价" Then
            intMethod = 1
            If InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") > 0 Then CheckPrivs = True
        ElseIf .TextMatrix(vsfList.Row, mEnuListCol.调价类型) = "售价成本价一起调价" Then
            intMethod = 2
            If InStr(1, ";" & mstrPrivs & ";", ";售价调价;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") > 0 Then CheckPrivs = True
        End If
    End With
End Function

Private Sub vsfPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '打印自定义报表：NO=调价汇总单号
    Dim strNo As String
    
    With vsfList
        If .rows > 1 Then
            strNo = .TextMatrix(.Row, mEnuListCol.调价号)
        End If
    End With
    
    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "NO=" & strNo)
End Sub


Private Sub BillPrint(ByVal intType As Integer)
    '单据打印
    Dim strTemp As String
    
    On Error GoTo errHandle
    
    With vsfList
        If .TextMatrix(.Row, mEnuListCol.调价号) = "" Then Exit Sub
        strTemp = .TextMatrix(.Row, mEnuListCol.调价号)
    End With
    
    
'    If vsfPrice.rows = 1 And vsfCost.rows = 1 Then
'        Exit Sub
'    ElseIf vsfPrice.rows <> 1 Then
'        strTemp = vsfPrice.TextMatrix(1, mEnuPriceCol.No)
'    ElseIf vsfCost.rows <> 1 Then
'        strTemp = vsfCost.TextMatrix(1, mEnuCostCol.No)
'    End If

    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1333", Me, "NO=" & strTemp, "包装单位=" & mintUnit, intType)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PrintView()
    '打印预览
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub filePrint()
    '打印
    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub billExcel()
    '输出到Excel
    If Me.ActiveControl Is vsfList Then
        vsfList.Redraw = flexRDNone
        subPrint 3
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    ElseIf Me.ActiveControl Is vsfPrice Then
        vsfPrice.Redraw = flexRDNone
        subExcel 3
        vsfPrice.Redraw = flexRDDirect
        vsfPrice.Col = 0
        vsfPrice.ColSel = vsfPrice.Cols - 1
    End If
End Sub
Private Sub subPrint(ByVal bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    If Format(mdaStart, "yyyy-mm-dd") = "1901-01-01" And Format(mdaVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "审核日期 " & Format(mdaVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdaVerifyEnd, "yyyy年MM月dd日")
    ElseIf Format(mdaVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "填制日期 " & Format(mdaStart, "yyyy年MM月dd日") & "至" & Format(mdaEnd, "yyyy年MM月dd日") & "  审核日期 " & Format(mdaVerifyStart, "yyyy年MM月dd日") & "至" & Format(mdaVerifyEnd, "yyyy年MM月dd日")
    Else
        strRange = "填制日期 " & Format(mdaStart, "yyyy年MM月dd日") & "至" & Format(mdaEnd, "yyyy年MM月dd日")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "药品调价管理"
        
    objRow.Add "时间：" & strRange
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & UserInfo.用户姓名
    objRow.Add "打印日期:" & Format(Sys.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    If vsfList Is ActiveControl Then
        Set objPrint.Body = vsfList
    ElseIf vsfPrice Is ActiveControl Then
        Set objPrint.Body = vsfPrice
    ElseIf vsfCost Is ActiveControl Then
        Set objPrint.Body = vsfCost
    End If
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub subExcel(ByVal bytMode As Byte)
'功能:进行输出到EXCEL
'参数:bytMode3 输出到EXCEL

    Dim objPrint As Object
    Dim objRow As zlTabAppRow
    Dim strRange As String
    
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "药品调价管理"
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "调价号." & Trim(vsfList.TextMatrix(vsfList.Row, mEnuListCol.调价号))
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "填制人:" & vsfList.TextMatrix(vsfList.Row, mEnuListCol.填制人) & "  填制日期:" & vsfList.TextMatrix(vsfList.Row, mEnuListCol.填制日期)

    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfPrice
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub cmdView_Click()
    Dim intMethod As Integer
    
    If vsfList.Row <= 0 Then Exit Sub
    If vsfList.TextMatrix(vsfList.Row, mEnuListCol.调价类型) = "仅售价调价" Then
        intMethod = 0
    ElseIf vsfList.TextMatrix(vsfList.Row, mEnuListCol.调价类型) = "仅成本价调价" Then
        intMethod = 1
    ElseIf vsfList.TextMatrix(vsfList.Row, mEnuListCol.调价类型) = "售价成本价一起调价" Then
        intMethod = 2
    End If
    frmMediPriceCard.ShowME Me, 2, vsfList.TextMatrix(vsfList.Row, mEnuListCol.调价号), intMethod
End Sub

Private Sub Form_Load()
    
'    Me.Height = Screen.Height * (3 / 4)
'    Me.Width = Screen.Width * (3 / 4)
    Me.Height = 768 * 15
    Me.Width = 1024 * 15
    
    mstrPrivs = GetPrivFunc(glngSys, 模块号.药品调价)
    Call initJinDu
    Call initCommandBars
    Call initMenu
    Call InitTool
    Call InitTabControl
    Call initVsflexgrid
    Call SetMenuEnable '用权限来控制菜单
    '添加自定义报表
    Call zldatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call RestoreWinState(Me, App.ProductName, MStrCaption)
    
    Call getListInfo
    stbThis.Panels(2).Picture = picColor
    mblnLoad = True
    
End Sub

Private Sub initJinDu()
    '功能：初始化各个单位下面的精度
    '判断是否以药库单位显示
    '获取设置的单位
    Dim intUnitTemp As Integer
'    Dim strOder As String
    
'    strOder = Val(zlDatabase.GetPara("排序", glngSys, 1333, "00"))
    
    mintUnit = Val(zldatabase.GetPara("药品单位", glngSys, 1333, "1"))
    Select Case mintUnit
        Case 0 '药库
            intUnitTemp = 4
        Case 1 '住院
            intUnitTemp = 3
        Case 2 '门诊
            intUnitTemp = 2
        Case 3 '售价
            intUnitTemp = 1
    End Select
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
    mintNumberDigit = GetDigitTiaoJia(1, 3, intUnitTemp)
    mintMoneyDigit = GetDigitTiaoJia(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    mintSalePriceDigit = GetDigitTiaoJia(1, 2, 1)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If

    If Me.ScaleHeight / 2 < 2000 Then Exit Sub
    vsfList.Move 0, 900, Me.ScaleWidth, Me.ScaleHeight / 2 - 2000
    picSplit.Left = 50
    picSplit.Top = vsfList.Top + vsfList.Height + 50
    picSplit.Width = Me.ScaleWidth
    cmdView.Move Me.ScaleWidth - cmdView.Width - 500, picSplit.Top + 50

    With TabDetails
        .Top = picSplit.Top + picSplit.Height + 20
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - picSplit.Top - picSplit.Height - stbThis.Height
    End With
    vsfPrice.Move 0, 360, TabDetails.Width, TabDetails.Height
    vsfCost.Move 0, 360, TabDetails.Width, TabDetails.Height
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - stbThis.Panels(3).Width - stbThis.Panels(4).Width - .Width - 300
    End With
End Sub

Private Sub initVsflexgrid()
    With vsfList
        .Editable = flexEDNone
        .Cols = mEnuListCol.总列数
        .rows = 1
        .ColWidth(0) = 200
        .Cell(flexcpFontBold, 0, 0, .rows - 1, .Cols - 1) = 50
        .RowHeight(0) = mconlngRowHeight
        .AllowSelection = False '不能多选
        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExSortShowAndMove '排序和移动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
'        .GridLineWidth = 2  '设置单元格边框
'        .GridLines = flexGridInset
'        .GridColor = &H0&
        '设置列宽
        .ColWidth(mEnuListCol.调价号) = 1500
        .ColWidth(mEnuListCol.调价类型) = 2000
        .ColWidth(mEnuListCol.填制人) = 1500
        .ColWidth(mEnuListCol.填制日期) = 2000
        .ColWidth(mEnuListCol.执行日期) = 2000
        .ColWidth(mEnuListCol.说明) = 2000
        '对齐方式
        .ColAlignment(mEnuListCol.调价号) = flexAlignLeftCenter
        .ColAlignment(mEnuListCol.调价类型) = flexAlignLeftCenter
        .ColAlignment(mEnuListCol.填制人) = flexAlignLeftCenter
        .ColAlignment(mEnuListCol.填制日期) = flexAlignLeftCenter
        .ColAlignment(mEnuListCol.执行日期) = flexAlignLeftCenter
        .ColAlignment(mEnuListCol.说明) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter

        .TextMatrix(0, mEnuListCol.调价号) = "调价号"
        .TextMatrix(0, mEnuListCol.调价类型) = "调价类型"
        .TextMatrix(0, mEnuListCol.填制人) = "填制人"
        .TextMatrix(0, mEnuListCol.填制日期) = "填制日期"
        .TextMatrix(0, mEnuListCol.执行日期) = "执行日期"
        .TextMatrix(0, mEnuListCol.说明) = "说明"
    End With

    With vsfPrice
        .Editable = flexEDNone
        .Cols = menuPriceCol.总列数
        .rows = 1
        .colHidden(0) = True
        .Cell(flexcpFontBold, 0, 0, .rows - 1, .Cols - 1) = 50 '字体加粗
        .RowHeight(0) = mconlngRowHeight
        .AllowSelection = False '不能多选
        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExSortShowAndMove '排序和移动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
'        .GridLineWidth = 2  '设置单元格边框
'        .GridLines = flexGridInset
'        .GridColor = &H0&
        '设置列宽
        .ColWidth(menuPriceCol.原价id) = 0
        .ColWidth(menuPriceCol.NO) = 1000
        .ColWidth(menuPriceCol.药品信息) = 3500
        .ColWidth(menuPriceCol.规格) = 1500
        .ColWidth(menuPriceCol.单位) = 800
        .ColWidth(menuPriceCol.单位系数) = 0
        .ColWidth(menuPriceCol.原价) = 1000
        .ColWidth(menuPriceCol.现价) = 1000
        .ColWidth(menuPriceCol.执行日期) = 0
        .ColWidth(menuPriceCol.调价人) = 1000
        '对齐方式
        .ColAlignment(menuPriceCol.NO) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.药品信息) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.规格) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.单位) = flexAlignCenterCenter
        .ColAlignment(menuPriceCol.原价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.现价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.执行日期) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.调价人) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter

        .TextMatrix(0, menuPriceCol.原价id) = "原价id"
        .TextMatrix(0, menuPriceCol.NO) = "NO"
        .TextMatrix(0, menuPriceCol.药品信息) = "药品"
        .TextMatrix(0, menuPriceCol.规格) = "规格"
        .TextMatrix(0, menuPriceCol.单位) = "单位"
        .TextMatrix(0, menuPriceCol.单位系数) = "单位系数"
        .TextMatrix(0, menuPriceCol.原价) = "原价"
        .TextMatrix(0, menuPriceCol.现价) = "现价"
        .TextMatrix(0, menuPriceCol.执行日期) = "执行日期"
        .TextMatrix(0, menuPriceCol.调价人) = "调价人"
    End With

    With vsfCost
        .Editable = flexEDNone
        .Cols = mEnuCostCol.总列数
        .rows = 1
        .colHidden(0) = True
        .Cell(flexcpFontBold, 0, 0, .rows - 1, .Cols - 1) = 50
        .RowHeight(0) = mconlngRowHeight
        .AllowSelection = False '不能多选
        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExSortShowAndMove '排序和移动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
'        .GridLineWidth = 2  '设置单元格边框
'        .GridLines = flexGridInset
'        .GridColor = &H0&
        '设置列宽
        .ColWidth(mEnuCostCol.NO) = 1000
        .ColWidth(mEnuCostCol.药品信息) = 3500
        .ColWidth(mEnuCostCol.库房) = 800
        .ColWidth(mEnuCostCol.规格) = 1500
        .ColWidth(mEnuCostCol.批号) = 1000
        .ColWidth(mEnuCostCol.产地) = 1500
        .ColWidth(mEnuCostCol.单位) = 800
        .ColWidth(mEnuCostCol.原价) = 1000
        .ColWidth(mEnuCostCol.现价) = 1000
        .ColWidth(mEnuCostCol.效期) = 1500
        .ColWidth(mEnuCostCol.执行日期) = 0
        .ColWidth(mEnuCostCol.调价人) = 1000
        '对齐方式
        .ColAlignment(mEnuCostCol.NO) = flexAlignLeftCenter
        .ColAlignment(mEnuCostCol.药品信息) = flexAlignLeftCenter
        .ColAlignment(mEnuCostCol.库房) = flexAlignLeftCenter
        .ColAlignment(mEnuCostCol.规格) = flexAlignLeftCenter
        .ColAlignment(mEnuCostCol.批号) = flexAlignLeftCenter
        .ColAlignment(mEnuCostCol.产地) = flexAlignLeftCenter
        .ColAlignment(mEnuCostCol.单位) = flexAlignCenterCenter
        .ColAlignment(mEnuCostCol.原价) = flexAlignRightCenter
        .ColAlignment(mEnuCostCol.现价) = flexAlignRightCenter
        .ColAlignment(mEnuCostCol.效期) = flexAlignLeftCenter
        .ColAlignment(mEnuCostCol.执行日期) = flexAlignLeftCenter
        .ColAlignment(mEnuCostCol.调价人) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter

        .TextMatrix(0, mEnuCostCol.NO) = "NO"
        .TextMatrix(0, mEnuCostCol.药品信息) = "药品"
        .TextMatrix(0, mEnuCostCol.库房) = "库房"
        .TextMatrix(0, mEnuCostCol.规格) = "规格"
        .TextMatrix(0, mEnuCostCol.批号) = "批号"
        .TextMatrix(0, mEnuCostCol.产地) = "生产商"
        .TextMatrix(0, mEnuCostCol.单位) = "单位"
        .TextMatrix(0, mEnuCostCol.原价) = "原成本价"
        .TextMatrix(0, mEnuCostCol.现价) = "现成本价"
        .TextMatrix(0, mEnuCostCol.效期) = "效期"
        .TextMatrix(0, mEnuCostCol.执行日期) = "执行日期"
        .TextMatrix(0, mEnuCostCol.调价人) = "调价人"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, MStrCaption)
    mblnLoad = False
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    If vsfList.Height + y <= 800 Then Exit Sub
    If TabDetails.Height - y <= 1000 Then Exit Sub
    picSplit.Move 0, picSplit.Top + y
    cmdView.Move Me.ScaleWidth - cmdView.Width - 500, picSplit.Top + 50
    vsfList.Move 0, 900, Me.ScaleWidth, vsfList.Height + y
    
    With TabDetails
        .Top = picSplit.Top + picSplit.Height + 20
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = TabDetails.Height - y
    End With
End Sub

Private Sub vsfCost_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfCost
        .Move 0, 360, TabDetails.Width, TabDetails.Height - 300
    End With
End Sub

Private Sub vsfList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '移动第一栏的标记到当前行！
    With vsfList
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 8
        End If
    End With
End Sub

Private Sub SetMenuEnable()
    '判断权限对菜单的影响
    Dim cbrMenuControl As CommandBarControl
    Dim cbrMenuPop As CommandBarControl

    '控制新增菜单
    Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Add, , True)
    Set cbrMenuControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Add, , True)
    If InStr(1, mstrPrivs, "登记") = 0 Or (InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") = 0) Then
        If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
        If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = False
    End If

    '控制修改菜单
    Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Update, , True)
    Set cbrMenuControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Update, , True)
    If InStr(1, mstrPrivs, "修改") = 0 Or (InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") = 0) Then
        If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
        If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = False
    End If

    '控制删除菜单
    Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Delete, , True)
    Set cbrMenuControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Delete, , True)
    If InStr(1, mstrPrivs, "删除") = 0 Or (InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") = 0) Then
        If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
        If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = False
    End If
    Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_File_Parameter, , True)
    If InStr(1, mstrPrivs, "参数设置") = 0 Then
        If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
    End If
End Sub

Private Sub getListInfo()
    '获取调价汇总信息
    Dim rstemp As ADODB.Recordset
    Dim strClass As String '调价类型
    Dim i As Integer
    Dim dateCurrentDate As Date
    Dim int查询天数 As Integer

    On Error GoTo errHandle
    
    dateCurrentDate = Sys.Currentdate
    int查询天数 = Val(zldatabase.GetPara("查询天数", glngSys, 1333, 7))
    mdaStart = Format(DateAdd("d", -int查询天数, dateCurrentDate), "yyyy-MM-dd")
    mdaEnd = CDate(Format(dateCurrentDate, "yyyy-MM-dd") & " 23:59:59")
    mdaVerifyStart = "1901-01-01"
    mdaVerifyEnd = "1901-01-01"
    If mSQLCondition.date填制时间开始 = "0:00:00" Then
        lblScope.Caption = "日期范围：" & Format(mdaStart, "yyyy-mm-dd") & "至" & Format(mdaEnd, "yyyy-mm-dd")
    Else
        lblScope.Caption = "日期范围：" & Format(mSQLCondition.date填制时间开始, "yyyy-mm-dd") & "至" & Format(mSQLCondition.date填制时间结束, "yyyy-mm-dd")
    End If
    
    vsfList.rows = 1
    vsfPrice.rows = 1
    vsfCost.rows = 1
    gstrSQL = "select a.调价号, a.类型, a.执行日期, a.填制日期, a.填制人, a.说明 from 调价汇总记录 a"
    
    '不等于空代表是过滤
    If mstrResult <> "" Then
        gstrSQL = gstrSQL + " where " + mstrResult + " and a.分类=0 order by a.调价号 desc"
        
    Else '默认只查询填制日期在一个周的调价信息
        gstrSQL = gstrSQL + " where " + " a.填制日期 between to_date('" & mdaStart & "', 'yyyy-mm-dd hh24:mi:ss') and to_date('" & mdaEnd & "', 'yyyy-mm-dd hh24:mi:ss') and a.分类=0 order by a.调价号 desc"
    End If
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "查询调价汇总记录", mSQLCondition.date填制时间开始, mSQLCondition.date填制时间结束, mSQLCondition.date执行时间开始, mSQLCondition.date执行时间结束)

    If rstemp.RecordCount = 0 Then Exit Sub
    rstemp.MoveFirst
    For i = 0 To rstemp.RecordCount - 1
        With vsfList
            .rows = .rows + 1
            .RowHeight(.rows - 1) = mconlngRowHeight
            .TextMatrix(.rows - 1, mEnuListCol.调价号) = rstemp!调价号
            If rstemp!类型 = 0 Then
                strClass = "仅售价调价"
            ElseIf rstemp!类型 = 1 Then
                strClass = "仅成本价调价"
            ElseIf rstemp!类型 = 2 Then
                strClass = "售价成本价一起调价"
            End If
            .TextMatrix(.rows - 1, mEnuListCol.调价类型) = strClass
            .TextMatrix(.rows - 1, mEnuListCol.填制人) = rstemp!填制人
            .TextMatrix(.rows - 1, mEnuListCol.填制日期) = Format(rstemp!填制日期, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.rows - 1, mEnuListCol.执行日期) = Format(rstemp!执行日期, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.rows - 1, mEnuListCol.说明) = IIf(IsNull(rstemp!说明), "", rstemp!说明)
            
            If rstemp!执行日期 > dateCurrentDate Then '未执行的用红色显示
                .Cell(flexcpForeColor, .rows - 1, 0, .rows - 1, .Cols - 1) = vbRed
            End If
            rstemp.MoveNext
        End With
    Next
    
    If vsfList.TextMatrix(1, mEnuListCol.调价号) <> "" Then
        vsfList.Row = 1
        vsfList.Col = 1
        Call getPriceInfo
        Call getCostInfo
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub getCostInfo()
    '获取成本价调价信息
    Dim rstemp As ADODB.Recordset
    Dim i As Integer
    Dim db包装系数 As Double
    Dim strUnit As String

    On Error GoTo errHandle
    
    gstrSQL = " Select B.NO, I.ID As 药品id, '[' || I.编码 || ']' || I.名称 ||  ' ' || I.产地 As 药品, P.名称 As 库房,A.批号,A.效期,A.产地,i.规格, " & _
            " I.计算单位 As 单位, S.药库单位, Nvl(S.药库包装, 1) 药库包装,s.住院单位,s.住院包装,s.门诊单位,s.门诊包装, A.原价 As 原成本价,A.现价 As 成本价, A.执行日期, B.摘要 " & _
            " From 药品收发记录 B, 收费项目目录 I, 药品规格 S, 部门表 P, 药品价格记录 A " & _
            " Where A.收发id = B.ID(+) And A.药品id = I.ID And " & _
            " I.ID = S.药品id And A.库房id = P.ID(+) And a.价格类型=2 And a.调价汇总号=[1] "
    gstrSQL = gstrSQL & "Union All " & _
            " Select B.NO, I.ID As 药品id, '[' || I.编码 || ']' || I.名称 ||  ' ' || I.产地 As 药品, P.名称 As 库房,A.批号,A.效期,A.产地,i.规格, " & _
            " I.计算单位 As 单位, S.药库单位, Nvl(S.药库包装, 1) 药库包装,s.住院单位,s.住院包装,s.门诊单位,s.门诊包装, A.原成本价,A.新成本价 As 成本价, A.执行日期, B.摘要 " & _
            " From 药品收发记录 B, 收费项目目录 I, 药品规格 S, 部门表 P, 成本价调价信息 A " & _
            " Where A.收发id = B.ID(+) And A.药品id = I.ID And " & _
            " I.ID = S.药品id And A.库房id = P.ID(+) And a.调价汇总号=[1] "
    gstrSQL = gstrSQL & " Order By 药品, 执行日期 Desc, NO Desc"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "查询成本价调价", vsfList.TextMatrix(vsfList.Row, mEnuListCol.调价号))
    
    vsfCost.rows = 1
    If rstemp.RecordCount = 0 Then Exit Sub

    With vsfCost
        For i = 0 To rstemp.RecordCount - 1
            .rows = .rows + 1
            Select Case mintUnit
                Case 0
                    db包装系数 = rstemp!药库包装
                    strUnit = rstemp!药库单位
                Case 1
                    db包装系数 = rstemp!住院包装
                    strUnit = rstemp!住院单位
                Case 2
                    db包装系数 = rstemp!门诊包装
                    strUnit = rstemp!门诊单位
                Case 3
                    db包装系数 = 1
                    strUnit = rstemp!单位
            End Select
            .RowHeight(.rows - 1) = mconlngRowHeight
            .TextMatrix(.rows - 1, mEnuCostCol.NO) = IIf(IsNull(rstemp!NO), "", rstemp!NO)
            .TextMatrix(.rows - 1, mEnuCostCol.药品信息) = rstemp!药品
            .TextMatrix(.rows - 1, mEnuCostCol.库房) = IIf(IsNull(rstemp!库房), "", rstemp!库房)
            .TextMatrix(.rows - 1, mEnuCostCol.规格) = IIf(IsNull(rstemp!规格), "", rstemp!规格)
            .TextMatrix(.rows - 1, mEnuCostCol.批号) = IIf(IsNull(rstemp!批号), "", rstemp!批号)
            .TextMatrix(.rows - 1, mEnuCostCol.产地) = IIf(IsNull(rstemp!产地), "", rstemp!产地)
            .TextMatrix(.rows - 1, mEnuCostCol.单位) = strUnit
            .TextMatrix(.rows - 1, mEnuCostCol.原价) = zlStr.FormatEx(IIf(IsNull(rstemp!原成本价), 0, rstemp!原成本价) * db包装系数, mintPriceDigit, , True)
            .TextMatrix(.rows - 1, mEnuCostCol.现价) = zlStr.FormatEx(IIf(IsNull(rstemp!成本价), 0, rstemp!成本价) * db包装系数, mintPriceDigit, , True)
            .TextMatrix(.rows - 1, mEnuCostCol.效期) = Format(IIf(IsNull(rstemp!效期), "", rstemp!效期), "yyyy-mm-dd")
            .TextMatrix(.rows - 1, mEnuCostCol.执行日期) = Format(IIf(IsNull(rstemp!执行日期), "", rstemp!执行日期), "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.rows - 1, mEnuCostCol.调价人) = vsfList.TextMatrix(vsfList.Row, mEnuListCol.填制人)
            rstemp.MoveNext
        Next
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub getPriceInfo()
    '获取售价调价信息
    Dim rstemp As ADODB.Recordset
    Dim i As Integer
    Dim db包装系数 As Double
    Dim strUnit As String
    
    On Error GoTo errHandle

    gstrSQL = "Select p.Id, i.规格,Decode(i.是否变价, 1, '时价', '定价') As 药价类型, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & _
                       " Nvl(i.屏蔽费别, 0) As 屏蔽费别," & _
                       " Decode(Sign(p.执行日期 - Sysdate), 1, 1, Decode(Sign(p.终止日期 - Sysdate), -1, -1, 0)) As 执行情况," & _
                       " '[' || i.编码 || ']' || i.名称 || ' '  || i.产地 As 药品, i.计算单位 As 单位, s.门诊单位, s.门诊包装, s.住院单位, s.住院包装," & _
                       " s.药库单位, Nvl(s.药库包装, 1) 药库包装,p.原价, p.现价 , u.名称 As 收入项目, p.调价说明, To_Char(p.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期,p.调价人," & _
                       " i.Id 药品id, p.No 调价no" & _
                " From 收费价目 P, 收入项目 U, 收费项目目录 I, 药品规格 S" & _
                " Where p.收费细目id = i.Id And p.收入项目id = u.Id And i.Id = s.药品id And p.调价汇总号 = [1] " & _
                GetPriceClassString("P") & _
                " Order By i.编码, p.执行日期 Desc"

    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, vsfList.TextMatrix(vsfList.Row, mEnuListCol.调价号))
    vsfPrice.rows = 1
    If rstemp.RecordCount = 0 Then Exit Sub

    With vsfPrice
        For i = 0 To rstemp.RecordCount - 1
            .rows = .rows + 1
            .RowHeight(.rows - 1) = mconlngRowHeight
            
            Select Case mintUnit
                Case 0
                    db包装系数 = rstemp!药库包装
                    strUnit = rstemp!药库单位
                Case 1
                    db包装系数 = rstemp!住院包装
                    strUnit = rstemp!住院单位
                Case 2
                    db包装系数 = rstemp!门诊包装
                    strUnit = rstemp!门诊单位
                Case 3
                    db包装系数 = 1
                    strUnit = rstemp!单位
            End Select
            
            .TextMatrix(.rows - 1, menuPriceCol.原价id) = rstemp!id
            .TextMatrix(.rows - 1, menuPriceCol.NO) = rstemp!调价no
            .TextMatrix(.rows - 1, menuPriceCol.药品信息) = rstemp!药品
            .TextMatrix(.rows - 1, menuPriceCol.规格) = IIf(IsNull(rstemp!规格), "", rstemp!规格)
            .TextMatrix(.rows - 1, menuPriceCol.单位) = strUnit
            .TextMatrix(.rows - 1, menuPriceCol.单位系数) = IIf(mintUnit = 0, 1, rstemp!药库包装)
            .TextMatrix(.rows - 1, menuPriceCol.原价) = zlStr.FormatEx(IIf(IsNull(rstemp!原价), 0, rstemp!原价) * db包装系数, mintCostDigit, , True)
            .TextMatrix(.rows - 1, menuPriceCol.现价) = zlStr.FormatEx(IIf(IsNull(rstemp!现价), 0, rstemp!现价) * db包装系数, mintCostDigit, , True)
            .TextMatrix(.rows - 1, menuPriceCol.执行日期) = Format(rstemp!执行日期, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.rows - 1, menuPriceCol.调价人) = rstemp!调价人
            rstemp.MoveNext
        Next
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_DblClick()
    If vsfList.MouseRow = 0 Then Exit Sub
    Call cmdView_Click
End Sub

Private Sub vsfList_EnterCell()
    Dim cbrMenuControl As CommandBarControl
    Dim cbrMenuPop As CommandBarControl
        
    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)
        .Redraw = flexRDDirect
         
        If .TextMatrix(.Row, mEnuListCol.调价类型) = "仅售价调价" Then
            TabDetails.Item(1).Visible = False
            TabDetails.Item(0).Visible = True
            TabDetails.Item(0).Selected = True
        ElseIf .TextMatrix(.Row, mEnuListCol.调价类型) = "仅成本价调价" Then
            TabDetails.Item(1).Visible = True
            TabDetails.Item(0).Visible = False
            TabDetails.Item(1).Selected = True
        Else
            TabDetails.Item(1).Visible = True
            TabDetails.Item(0).Visible = True
            TabDetails.Item(0).Selected = True
        End If
        If .TextMatrix(.Row, mEnuListCol.调价号) <> "" And .Row > 0 Then
            '执行日期大于系统当前日期不能修改调价单
            Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Update, , True)
            Set cbrMenuControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Update, , True)
            If CDate(.TextMatrix(.Row, mEnuListCol.执行日期)) <= Sys.Currentdate() Then
                If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
                If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = False
            Else
                If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = True
                If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = True
            End If
        
            '执行日期大于系统当前日期不能删除调价单
            Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Delete, , True)
            Set cbrMenuControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Delete, , True)
            If CDate(.TextMatrix(.Row, mEnuListCol.执行日期)) <= Sys.Currentdate() Then
                If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
                If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = False
            Else
                If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = True
                If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = True
            End If
            Call SetMenuEnable
            
            Call getPriceInfo
            Call getCostInfo
        End If
'        If mblnLoad = True Then
'            vsfList.SetFocus
'        End If
    End With
End Sub

Private Sub deleteNotExecutePirce()
    '清除未执行价格
    Dim rstemp As ADODB.Recordset
    Dim arrSql As Variant
    Dim i As Integer
    Dim int调价类型 As Integer  '0-所有;1-售价;2-成本价
    
    On Error GoTo errHandle
    arrSql = Array()
    With vsfList
        If .TextMatrix(.Row, mEnuListCol.调价号) <> "" Then
            If MsgBox("确定删除该条调价单据？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                
            If .TextMatrix(.Row, mEnuListCol.调价类型) = "仅售价调价" Then
                int调价类型 = 1
            ElseIf .TextMatrix(.Row, mEnuListCol.调价类型) = "仅成本价调价" Then
                int调价类型 = 2
            Else
                int调价类型 = 0
            End If
            gstrSQL = "select 收费细目id as id from 收费价目 where 调价汇总号=[1]" & GetPriceClassString("") & _
                        " union " & _
                      " select 药品id as id from 药品价格记录 where 调价汇总号=[1]"
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "清除价格", .TextMatrix(.Row, mEnuListCol.调价号))
            If rstemp.RecordCount = 0 Then
                MsgBox "该调价信息已经被人删除！", vbInformation, gstrSysName
                Exit Sub
            Else
                rstemp.MoveFirst
                Do While Not rstemp.EOF
                    gstrSQL = "Zl_药品未执行价格_Delete(" & rstemp!id & "," & int调价类型 & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                    rstemp.MoveNext
                Loop
            End If
            
            gcnOracle.BeginTrans
            For i = 0 To UBound(arrSql)
                Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveRestore")
            Next
            gcnOracle.CommitTrans
        End If
    End With
    '删除后刷新界面
    Call getListInfo
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_GotFocus()
    Call SetGridFocus(vsfList, True)
End Sub

Private Sub vsfList_LostFocus()
    Call SetGridFocus(vsfList, False)
End Sub

Private Sub vsfPrice_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfPrice
        .Move 0, 360, TabDetails.Width, TabDetails.Height - 300
    End With
End Sub

Private Sub vsfPrice_GotFocus()
    Call SetGridFocus(vsfPrice, True)
End Sub
Private Sub vsfPrice_LostFocus()
    Call SetGridFocus(vsfPrice, False)
End Sub

Private Sub vsfcost_GotFocus()
    Call SetGridFocus(vsfCost, True)
End Sub
Private Sub vsfcost_LostFocus()
    Call SetGridFocus(vsfCost, False)
End Sub
