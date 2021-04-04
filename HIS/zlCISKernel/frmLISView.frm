VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmLisView 
   Caption         =   "检验结果"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15540
   Icon            =   "frmLISView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   15540
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6465
      Left            =   270
      ScaleHeight     =   6465
      ScaleWidth      =   3705
      TabIndex        =   6
      Top             =   435
      Width           =   3705
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   3765
         Left            =   30
         TabIndex        =   7
         Top             =   1530
         Width           =   2970
         _Version        =   589884
         _ExtentX        =   5239
         _ExtentY        =   6641
         _StockProps     =   0
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picFind 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   45
         ScaleHeight     =   1275
         ScaleWidth      =   3210
         TabIndex        =   8
         Top             =   105
         Width           =   3210
         Begin VB.ComboBox cboPages 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   885
            Width           =   1335
         End
         Begin VB.CommandButton cmd项目 
            Caption         =   "…"
            Height          =   300
            Left            =   2325
            TabIndex        =   15
            Top             =   480
            Width           =   350
         End
         Begin VB.TextBox txt项目 
            Height          =   300
            Left            =   510
            TabIndex        =   13
            Top             =   495
            Width           =   1800
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   300
            Left            =   510
            TabIndex        =   10
            Top             =   105
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   106233859
            CurrentDate     =   39819
         End
         Begin VB.CommandButton cmdOK 
            Height          =   300
            Left            =   2745
            Picture         =   "frmLISView.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   350
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   1830
            TabIndex        =   11
            Top             =   105
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   106233859
            CurrentDate     =   39819
         End
         Begin VB.Label lblPages 
            AutoSize        =   -1  'True
            Caption         =   "住院次数"
            Height          =   180
            Left            =   90
            TabIndex        =   17
            Top             =   930
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "项目"
            Height          =   180
            Left            =   90
            TabIndex        =   14
            Top             =   555
            Width           =   660
         End
         Begin VB.Label lblinfo 
            Caption         =   "日期"
            Height          =   180
            Left            =   90
            TabIndex        =   12
            Top             =   165
            Width           =   660
         End
      End
   End
   Begin VB.PictureBox PicTab 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   4515
      ScaleHeight     =   2535
      ScaleWidth      =   3990
      TabIndex        =   4
      Top             =   2085
      Width           =   3990
      Begin XtremeSuiteControls.TabControl TabCtlWindow 
         Height          =   2280
         Left            =   90
         TabIndex        =   5
         Top             =   105
         Width           =   3765
         _Version        =   589884
         _ExtentX        =   6641
         _ExtentY        =   4022
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox PicImage 
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   9570
      ScaleHeight     =   2595
      ScaleWidth      =   1935
      TabIndex        =   1
      Top             =   360
      Width           =   1935
      Begin VB.VScrollBar VScroll 
         Height          =   1245
         Left            =   1620
         Max             =   0
         TabIndex        =   2
         Top             =   150
         Width           =   225
      End
      Begin C1Chart2D8.Chart2D ChartThis 
         Height          =   735
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   120
         Width           =   885
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   1561
         _ExtentY        =   1296
         _StockProps     =   0
         ControlProperties=   "frmLISView.frx":685E
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7020
      Width           =   15540
      _ExtentX        =   27411
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLISView.frx":6DE1
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24500
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
   Begin MSComctlLib.ImageList Imglist 
      Left            =   135
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":7675
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":7C0F
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":81A9
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":8743
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":8CDD
            Key             =   ""
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":9277
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":9611
            Key             =   ""
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":99AB
            Key             =   ""
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":9D45
            Key             =   ""
            Object.Tag             =   "9"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrThis 
      Left            =   1020
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmLISView.frx":A0DF
      Left            =   2565
      Top             =   180
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLisView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000

Private mlng病人ID As Long
Private mlng主页ID As Long

Private mlng医嘱ID As Long
Private mlng标本ID As Long
Private mlng结果次数 As Long

Private mstrPrivs As String
Private mstrLike As String

Private mfrmLisRptGeneral   As frmLisRptGeneral                  '报告查看
Attribute mfrmLisRptGeneral.VB_VarHelpID = -1
Private mfrmLisRptMicrobiology As frmLisRptMicrobiology                ' 微生物报告查看

Private Const ID_MENU_MOUSE = 90

Private Const Dkp_ID_Request As Integer = 3                         '核对登记窗格
Private Const Dkp_ID_Append As Integer = 4                          '报告附加窗格
Private Const Dkp_ID_Image As Integer = 5                           '显示检验图像

Private Enum mCol
    ID = 0: 类型: 紧急: 申请时间: 检验项目: 来源: 住院次数: 标本号: 微生物标本: 医嘱ID: 结果次数 ': 姓名: 性别: 年龄: 病人id:  检验人: 审核人: 婴儿: 主页ID: 审核时间: 定位
End Enum
Dim blnLoad As Boolean
Private mlngItemID As Long      '上次选中行
Private mstrWhere As String     '过滤条件

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private mlngMod As Integer '调用模块号
Private mblnShowBorder As Boolean       '是否显示窗体的border
Private mlngPageId As Long                '住院病人主页

Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Sub ShowMe(ByVal lng病人ID As Long, ByVal lngMod As Long, ByVal frmMain As Form, Optional ByVal blnShowBorder As Boolean = True, _
                Optional ByRef objOutFrm As Object)
    On Error GoTo errHandle
    
    mblnShowBorder = blnShowBorder
    mlng病人ID = 0
    If lng病人ID = 0 Then Exit Sub
    mlng病人ID = lng病人ID
    mlngMod = lngMod
    If blnShowBorder Then
        Me.Show , frmMain  '如果不显示窗体的边框，则表示该窗体为嵌入式调用，不是调用show方法
    Else
        Call YSystemMenu(Me.hWnd)
    End If
    Set objOutFrm = Me

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CreateCbs()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbrThis.Icons = zlCommFun.GetPubIcons
    With Me.cbrThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrThis.EnableCustomization False

    '-----------------------------------------------------
    '菜单定义
    Me.cbrThis.ActiveMenuBar.Title = "菜单"
'    Me.cbrthis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&T)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)")

       ' Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&O)"): cbrControl.BeginGroup = True

        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With
    

    'conMenu_EditPopup
'    '右键菜单
    Set cbrMenuBar = Me.cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_MENU_MOUSE, "右键菜单", -1, False)
    cbrMenuBar.ID = ID_MENU_MOUSE
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告预览(&V)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "报告查询(&P)"): cbrControl.BeginGroup = True

    End With
    cbrMenuBar.Visible = False

    Set cbrMenuBar = Me.cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_Expend, "展开/折叠组(&X)")
        With cbrControl.CommandBar.Controls
            Set cbrPopControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)", -1, False)
            Set cbrPopControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)", -1, False)
            Set cbrPopControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)", -1, False): cbrPopControl.BeginGroup = True
            Set cbrPopControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&E)", -1, False)
        End With
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Location, "显示分组框(&S)"): cbrPopControl.BeginGroup = True
    
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
            cbrPopControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Backward, "前一条(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Forward, "后一条(&L)")


        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_LeaveMedi, "隐藏检验图形"): 'cbrControl.BeginGroup = True
        
        If zlDatabase.GetPara("隐藏检验图形", glngSys, mlngMod, "True") = "True" Then
            cbrControl.Checked = True
        End If

        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_HideList, "检验列表(&P)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&F)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With

    '快键绑定
    
    With Me.cbrThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_ESCAPE, conMenu_LIS_Cancel
        .Add 0, VK_PAGEUP, conMenu_Tool_Reference_1
        .Add 0, VK_PAGEDOWN, conMenu_Tool_Reference_2

    End With
    Me.cbrThis.ActiveMenuBar.Visible = mblnShowBorder
    '设置不常用菜单
'    With Me.cbrthis.Options
'        .AddHiddenCommand conMenu_File_PrintSet
'    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbrThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

End Sub

Private Sub CreateDockPane()
    Dim Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    Dim lngPane5Width As Long, lngPane2Height As Long, lngPane2Width As Long, lngPane3Height As Long
    

    dkpMain.Options.HideClient = True
    
    Set Pane3 = dkpMain.CreatePane(Dkp_ID_Request, 100, 600, DockLeftOf)
    Pane3.Title = "核收登记"
    Pane3.Handle = Me.PicInfo.hWnd
    Pane3.Options = PaneNoCaption
    
    
    Set Pane4 = dkpMain.CreatePane(Dkp_ID_Append, 9800, 790, DockRightOf, Pane3)
    Pane4.Title = "附加窗体"
    Pane4.Handle = Me.PicTab.hWnd
    Pane4.Options = PaneNoCaption
    
    lngPane5Width = 200
    Set Pane5 = dkpMain.CreatePane(Dkp_ID_Image, lngPane5Width, 200, DockRightOf, Pane4)
    Pane5.Title = "图像显示"
    Pane5.Handle = Me.PicImage.hWnd
'    Pane5.Options = PaneNoCaption
    Pane4.Select
    
End Sub

Private Sub CreateTableControl()
    
    On Error Resume Next

    With Me.TabCtlWindow
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem(0, "检验报告", mfrmLisRptGeneral.hWnd, conMenu_Tool_Report).Tag = "普通报告结果"
        .InsertItem(1, "检验报告", mfrmLisRptMicrobiology.hWnd, conMenu_Tool_Report).Tag = "微生物报告结果"
        
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
        DoEvents
        .Item(1).Visible = False
    End With

End Sub

Private Sub CreateRptListHead()
    Dim Column As ReportColumn
    Dim i As Integer

    With Me.rptList.Columns


        rptList.SetImageList Imglist

        Set Column = .Add(mCol.ID, "ID", 30, True): Column.Visible = False
        
        Set Column = .Add(mCol.类型, "检验类型", 90, True): Column.Groupable = True
        
        Set Column = .Add(mCol.紧急, "", 18, False): Column.Icon = 0

        Set Column = .Add(mCol.检验项目, "检验项目", 90, True): Column.Groupable = True

        Set Column = .Add(mCol.申请时间, "申请时间", 80, True): Column.Groupable = True
        Column.Sortable = True: Column.SortAscending = False: Me.rptList.SortOrder.Add Column
        Set Column = .Add(mCol.来源, "来源", 30, True): Column.Groupable = True
        Set Column = .Add(mCol.住院次数, "住院次数", 65, True): Column.Groupable = False
        
        Set Column = .Add(mCol.标本号, "标本号", 65, True): Column.Groupable = False
        Set Column = .Add(mCol.微生物标本, "微生物标本", 30, True): Column.Visible = False: Column.Groupable = True
        Set Column = .Add(mCol.医嘱ID, "医嘱id", 30, True): Column.Visible = False: Column.Groupable = False
        Set Column = .Add(mCol.结果次数, "结果次数", 30, True): Column.Visible = False: Column.Groupable = False

    End With
    
    With rptList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
        
    End With
        
    '加入分组
    Me.rptList.GroupsOrder.DeleteAll
    Me.rptList.GroupsOrder.Add Me.rptList.Columns.Find(mCol.类型)
    Me.rptList.GroupsOrder(0).SortAscending = True
    Me.rptList.Columns.Find(mCol.类型).Visible = False
    Me.rptList.Populate
End Sub

Private Sub ImageTypeSet(intCount As Integer, Optional blnReset As Boolean = False)
    '功能           对检验图像进行排版
    '参数           intCount = 图像数
    '               blnReset = 是否需要重新读入
    Dim intLoop As Integer
    Dim Pane5 As Pane

    On Error Resume Next
    
    For intLoop = 0 To intCount
        If intLoop = 0 Then
            With Me.ChartThis(intLoop)
                .Visible = True
                .Top = 0
                .Left = 0
                .Width = IIF(Me.PicImage.ScaleWidth - Me.VScroll.Width - 20 <= 100, 100, Me.PicImage.ScaleWidth - Me.VScroll.Width - 20)
                .Height = .Width
            End With
        Else
            If blnReset = True And Me.ChartThis.UBound < intLoop Then
                Load Me.ChartThis(intLoop)
            End If
            With Me.ChartThis(intLoop)
                .Visible = True
                .Top = Me.ChartThis(intLoop - 1).Top + Me.ChartThis(intLoop - 1).Height + 10
                .Left = 0
                .Width = Me.ChartThis(intLoop - 1).Width
                .Height = .Width
                .IsBatched = False
            End With
        End If
    Next
    
    '隐藏多余的Chart控件
    For intLoop = intCount + 1 To Me.ChartThis.UBound
        Me.ChartThis(intLoop).Visible = False
    Next
    
    Set Pane5 = Me.dkpMain.FindPane(Dkp_ID_Image)
    If Not Pane5 Is Nothing Then
'        If intCount < 0 Then
'            Pane5.Close
'        Else
            If Me.cbrThis.FindControl(, conMenu_Manage_LeaveMedi, , True).Checked = False Then

                Me.dkpMain.ShowPane (Dkp_ID_Image)
                Me.dkpMain.FindPane(Dkp_ID_Request).Select
                Me.dkpMain.RecalcLayout
            Else
                Pane5.Close
            End If
'        End If
    End If
    With Me.VScroll
        .Top = 0
        .Left = Me.PicImage.ScaleWidth - .Width - 10
        .Height = Me.PicImage.ScaleHeight
        .Max = intCount
        .SmallChange = 1
        .LargeChange = 1
    End With
End Sub

Private Function GetLastPageId(ByVal lngPatientID As String) As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    GetLastPageId = 0
    
    strSQL = "Select Distinct a.病人id, b.门诊号, a.住院号, a.入院病床, a.姓名, a.主页id," & _
            " a.入院日期, a.出院日期 From 病案主页 a, 病人信息 b where a.病人id=b.病人id and a.病人id=[1] order by 主页id"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检验技师站", lngPatientID)
    
    With Me.cboPages
        .Clear
        .AddItem "所有"
        .ItemData(.NewIndex) = 0
        Do Until rsTmp.EOF
            .AddItem "第 " & rsTmp("主页ID") & " 次"
            .ItemData(.NewIndex) = rsTmp("病人id")
            rsTmp.MoveNext
        Loop
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveLast
            GetLastPageId = rsTmp("主页ID")
            .Text = "第 " & rsTmp("主页ID") & " 次"
        Else
            .Visible = False
            lblPages.Visible = False
        End If
    End With
End Function

Private Sub LoadAllData(ByVal lng病人ID As Long)
    '调入数据
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim objRow As ReportRow
    Dim blnHave As Boolean
    Dim dateMin As Date, dateMax As Date, str项目 As String
    Dim lngPageId As Long
    
    On Error GoTo errHandle
    
    lngPageId = mlngPageId
    
    dateMin = CDate(0)
    dateMax = CDate(0)
    
    Me.rptList.Records.DeleteAll
    If lngPageId > 0 Or (cboPages.Text = "所有" And cboPages.Visible = True) Then
        rptList.Columns(mCol.住院次数).Visible = True
    Else
        rptList.Columns(mCol.住院次数).Visible = False
    End If
    
    If mstrWhere = "" Then
        strSQL = "Select A.ID, A.紧急, A.申请时间, A.检验项目, A.标本序号, Nvl(A.微生物标本,0) as 微生物标本, A.姓名, A.性别, A.年龄, A.报告结果 as 结果次数, A.医嘱id, A.病人id, A.检验人, A.审核人," & vbNewLine & _
                "       A.婴儿, a.主页id 住院次数, A.审核时间, decode(A.病人来源,1,'门诊',2,'住院','其他') as 病人来源, A.操作类型" & vbNewLine & _
                "From 检验标本记录 A,病人信息 B  " & vbNewLine & _
                "Where A.病人id = [1] And A.审核人 is Not null And A.病人id=B.病人id " & IIF(lngPageId > 0, " And (a.主页id = [2] or a.主页id is null) ", "") & vbNewLine & _
                "Order By A.申请时间, A.审核时间 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lngPageId)
    Else
        dateMin = CDate(Split(mstrWhere, "|")(0))
        dateMax = CDate(Split(mstrWhere, "|")(1))
        str项目 = CStr(Split(mstrWhere, "|")(2))
        dateMax = DateAdd("d", 1, dateMax)
        
        strSQL = "Select A.ID, A.紧急, A.申请时间, A.检验项目, A.标本序号,  Nvl(A.微生物标本,0) as 微生物标本, A.姓名, A.性别, A.年龄, A.报告结果 as 结果次数, A.医嘱id, A.病人id, A.检验人, A.审核人," & vbNewLine & _
                "       A.婴儿, a.主页id 住院次数, A.审核时间, decode(A.病人来源,1,'门诊',2,'住院','其他') as 病人来源, A.操作类型" & vbNewLine & _
                "From 检验标本记录  A ,病人信息 B " & vbNewLine & _
                "Where A.病人id = [1] And A.病人id=B.病人id And A.审核人 is Not null And A.申请时间 Between [2] And [3] " & vbNewLine & _
                IIF(str项目 = "", "", " And instr(A.检验项目,[4])>0  ") & IIF(lngPageId > 0, " And (a.主页id = [5] or a.主页id is null) ", "") & _
                "Order By A.申请时间, A.审核时间 Desc"
                
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, dateMin, dateMax, str项目, lngPageId)
    End If
    Do Until rsTmp.EOF
        Set Record = Me.rptList.Records.Add
        For intLoop = 0 To Me.rptList.Columns.Count + 1
            Record.AddItem ""
        Next
        Record.Item(mCol.类型).value = "" & rsTmp("操作类型")
        If Val("" & rsTmp("紧急")) = 1 Then
            Record.Item(mCol.紧急).Icon = 1
        End If
        Record.Item(mCol.检验项目).value = Trim("" & rsTmp("检验项目"))
        Record.Item(mCol.申请时间).Caption = Format("" & rsTmp("申请时间"), "MM-dd HH:mm:ss")
        Record.Item(mCol.申请时间).value = Format("" & rsTmp("申请时间"), "YYYY-MM-dd HH:mm:ss")
        Record.Item(mCol.标本号).value = Trim("" & rsTmp("标本序号"))
        Record.Item(mCol.微生物标本).value = Trim("" & rsTmp("微生物标本"))
        Record.Item(mCol.医嘱ID).value = Val("" & rsTmp("医嘱ID"))
        Record.Item(mCol.ID).value = Val("" & rsTmp!ID)
        Record.Item(mCol.结果次数).value = Val("" & rsTmp("结果次数"))
        Record.Item(mCol.来源).value = "" & rsTmp("病人来源")
        Record.Item(mCol.住院次数).value = "" & rsTmp("住院次数")
        If mstrWhere = "" And IsNull(rsTmp("申请时间")) = False Then
            If dateMin = CDate(0) Then
                dateMin = CDate(Format("" & rsTmp("申请时间"), "YYYY-MM-dd"))
                dateMax = CDate(Format("" & rsTmp("申请时间"), "YYYY-MM-dd"))
            Else
                If CDate(Format("" & rsTmp("申请时间"), "YYYY-MM-dd")) > dateMax Then
                    dateMax = CDate(Format("" & rsTmp("申请时间"), "YYYY-MM-dd"))
                End If
            End If
        End If
        blnHave = True
        rsTmp.MoveNext
    Loop
    
    If mstrWhere = "" Then
        dtpStart.MaxDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
        dtpEnd.MaxDate = dtpStart.MaxDate
        
        txt项目.Text = ""
        txt项目.Tag = ""
        strSQL = "select 姓名,登记时间 from 病人信息 Where 病人id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        Do Until rsTmp.EOF
            Me.Caption = "报告查阅(" & rsTmp.Fields("姓名") & ")"
            dtpEnd.MinDate = rsTmp!登记时间
            dtpStart.MinDate = rsTmp!登记时间
            rsTmp.MoveNext
        Loop
        
        strSQL = "Select Min(入院日期) as 入院日期 From 病案主页 Where 病人ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        
        If Not rsTmp.EOF Then
            If rsTmp!入院日期 < dtpEnd.MinDate Then
                dtpEnd.MinDate = rsTmp!入院日期
                dtpStart.MinDate = rsTmp!入院日期
            End If
        End If
        
        dtpStart.value = dtpStart.MinDate
        dtpEnd.value = dtpEnd.MaxDate
    End If
    
    '1-刷新
    rptList.Populate

    '2-折叠所有组
    For Each objRow In rptList.Rows
        If objRow.GroupRow Then objRow.Expanded = False
    Next
    
    '3-定位到上次选中行
    If mlngItemID <> 0 Then
        For Each objRow In Me.rptList.Rows
            If objRow.GroupRow = False Then
                If Val(objRow.Record(mCol.ID).value) = mlngItemID Then
                    Set Me.rptList.FocusedRow = objRow
                    Exit For
                End If
            End If
        Next
    End If
    
    '4-展开选中行
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Me.rptList.Rows(0).Expanded = True
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    '5-调用事件
    Call rptList_SelectionChanged
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ReadImageData(lngKeyID As Long, blnSave As Boolean) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim DrawIndex As Integer
    Dim strTime As Date
    Dim objLisDev As Object, strFilename As String, strErr As String
    
    On Error GoTo errH
    strTime = Now
    gstrSQL = "select id ,标本ID,图像类型 from 检验图像结果 where 标本id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKeyID)
    '图像排版
    ImageTypeSet rsTmp.RecordCount - 1, True
    '不显示时不更新
    If Me.cbrThis.FindControl(, conMenu_Manage_LeaveMedi, , True).Checked = True Then Exit Function
    
    Set objLisDev = CreateObject("zlLISDev.clsDrawGraph")
    
    If objLisDev.GetSampleImgInit(glngSys, gcnOracle, strErr) Then
        Do Until rsTmp.EOF
            If Dir(App.Path & "\" & rsTmp("ID") & ".cht") = "" Then
                If Not objLisDev Is Nothing Then
                    strFilename = ""
                    
                    strFilename = objLisDev.GetImage(Val("" & rsTmp("ID")), App.Path, False, strErr)
                    If strFilename <> "" Then
                        Me.ChartThis(DrawIndex).Load App.Path & "\" & strFilename
                    Else
                        '读取文件 失败
                         
                    End If
                Else
                    '部件创建失败!
                End If
            Else
                Me.ChartThis(DrawIndex).Load App.Path & "\" & rsTmp("ID") & ".cht"
                 
            End If
            DrawIndex = DrawIndex + 1
            rsTmp.MoveNext
        Loop
        Call objLisDev.GetSampleImgExit(strErr)
    End If
    ImageTypeSet DrawIndex - 1, False

    ReadImageData = True
'    Debug.Print "ID=" & lngKeyID & ",用时:" & DateDiff("s", strTime, Now)
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub PrintSetup()
    '打印设置
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lng医嘱ID As Long, lng发送号 As Long, lng病人ID As Long
    Dim strSQL As String
    
    On Error GoTo errH
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    lng医嘱ID = mlng医嘱ID
    lng病人ID = mlng病人ID
    
    strSQL = "select 发送号 from 病人医嘱发送 a , 病人医嘱记录 b where b.id = a.医嘱id and b.id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng医嘱ID)
    If rsTmp.EOF = False Then
        lng发送号 = Val("" & rsTmp(0))
    End If
    
    If GetReportCode(lng医嘱ID, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        ReportPrintSet gcnOracle, glngSys, strReportCode, Me
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetReportCode(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByRef strCode As String, ByRef strNO As String, ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能;
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    If lng医嘱ID = 0 And lng发送号 = 0 Then Exit Function
    
    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2' AS 报表编号," & _
                       "A.NO," & _
                       "A.记录性质 " & _
                "FROM 病人医嘱发送 A,病历文件列表 C,病人医嘱记录 D,病历单据应用 E " & _
                "Where E.病历文件id = C.ID " & _
                        "AND D.诊疗项目ID=E.诊疗项目ID " & _
                      "AND A.医嘱ID=D.ID AND E.应用场合=Decode(D.病人来源,2,2,4,4,1) " & _
                      " AND D.相关id= [1] "
    If DataMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    On Error GoTo errH
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lng医嘱ID, lng发送号)
                      
    
    If rs.BOF = False Then
        strCode = zlCommFun.NVL(rs("报表编号"))
        strNO = zlCommFun.NVL(rs("NO"))
        bytMode = zlCommFun.NVL(rs("记录性质"), 1)
    End If
    GetReportCode = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReportPrint(ByVal blnPrint As Boolean)
    '单个报告打印
    
    Dim strReportCode As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lng医嘱ID As Long, lng发送号 As Long, lng病人ID As Long
    Dim strSQL As String

    Dim intLoop As Integer
    On Error GoTo errH
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    
    'blnCurrMoved = rptList.SelectedRows(0).Record.Item(mCol.转出).Value = "√"
    Call Open_LIS_Report(Me, mlng医嘱ID, mlng病人ID, blnCurrMoved, blnPrint)

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowOrHideItem(Control As CommandBarControl, DkpID As Integer)
    '功能               '显示或隐藏
    Dim Pane As Pane
    Set Pane = Me.dkpMain.FindPane(DkpID)
    If Control.Checked = True Then
        Pane.Close
    Else
        
        If Pane.Closed Then Me.dkpMain.ShowPane (DkpID)
        Pane.Select
    End If
    If DkpID = Dkp_ID_Image Then ReadImageData mlng标本ID, False
    Me.dkpMain.RecalcLayout
    Me.cbrThis.RecalcLayout
End Sub

Private Sub BackOrNextPatient(Move As Integer)
    '功能                 移动到上一个病人或下一个病人
    '参数                 Move = 1 上一病人 =2 下一病人
    Dim Rerow As ReportRow
    Dim i As Long
    With Me.rptList
        If .Rows.Count <= 0 Then Exit Sub
        i = .SelectedRows(0).Index
        If Move = 1 Then            '向上移动
            If i - 1 >= 0 Then
                i = i - 1
                .FocusedRow = .Rows(i)
            End If
        Else
            If i < .Rows.Count - 1 Then
                i = i + 1
                .FocusedRow = .Rows(i)
            End If
        End If
    End With
End Sub

Private Sub cboPages_Click()
    mlngPageId = Val(Trim(Replace(Replace(cboPages.Text, "第", ""), "次", "")))
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    
    Select Case Control.ID
        
        '''''''''''''''''''''''''''''''''''''''文件''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_File_PrintSet                                                      '打印设置
             PrintSetup
            
        Case conMenu_File_Preview                                                       '报告预览
            ReportPrint False
        
        Case conMenu_File_Print                                                         '报告打印
            ReportPrint True
        Case conMenu_File_Exit                                                          '退出
            Unload Me
        Case conMenu_View_Refresh
            Call zlRefresh
        '''''''''''''''''''''''''''''''''''''''查看'''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_View_ToolBar_Button                                                '标准按钮
            Control.Checked = Not Control.Checked
            Me.cbrThis(2).Visible = Control.Checked
            Me.cbrThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text                                                  '文本标签
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbrThis(2).Controls
                cbrControl.Style = IIF(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbrThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Size                                                  '大图标
            Control.Checked = Not Control.Checked
            Me.cbrThis.Options.LargeIcons = Not Me.cbrThis.Options.LargeIcons
            Me.cbrThis.RecalcLayout
        
        Case conMenu_View_StatusBar                                                     '状态栏
            Control.Checked = Not Control.Checked
            Me.stbThis.Visible = Control.Checked
            Me.cbrThis.RecalcLayout
'''
        Case conMenu_View_Expend_CurCollapse                            '折叠当前组
            If rptList.SelectedRows.Count > 0 Then
                If rptList.SelectedRows(0).GroupRow Then
                    rptList.SelectedRows(0).Expanded = False
                ElseIf Not rptList.SelectedRows(0).ParentRow Is Nothing Then
                    If rptList.SelectedRows(0).ParentRow.GroupRow Then
                        rptList.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '因折叠定位到分组上,不会自动激活该事件
            Call rptList_SelectionChanged
    
        Case conMenu_View_Expend_CurExpend                              '展开当前组
            If rptList.SelectedRows.Count > 0 Then
                rptList.SelectedRows(0).Expanded = True
            End If
        Case conMenu_View_Expend_AllCollapse                            '折叠所有组
            For Each objRow In rptList.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '因折叠定位到分组上,不会自动激活该事件
            Call rptList_SelectionChanged
        Case conMenu_View_Expend_AllExpend                              '展开所有组
            For Each objRow In rptList.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        Case conMenu_View_Location
            rptList.ShowGroupBox = Not rptList.ShowGroupBox             '显示分组框

'''
        Case conMenu_View_Forward                                                       '前一条
            BackOrNextPatient 2
        
        Case conMenu_View_Backward                                                      '后一条
            BackOrNextPatient 1
            
        Case conMenu_LIS_HideList                                                       '隐藏列表
            Control.Checked = Not Control.Checked
            ShowOrHideItem Control, Dkp_ID_Request
        
        Case conMenu_Manage_LeaveMedi                                                   '隐藏检验图形
            Control.Checked = Not Control.Checked
            ShowOrHideItem Control, Dkp_ID_Image
        ''''''''''''''''''''''''''''''''''''''帮助''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Help_Help                                                          '帮助主题
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_Web                                                           'WEB上的
            Call zlHomePage(hWnd)
        
        Case conMenu_Help_Web_Home                                                      '主页
            Call zlHomePage(Me.hWnd)
        
        Case conMenu_Help_Web_Mail                                                      '发送反馈
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_Help_About                                                         '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
            
    End Select
End Sub

Private Sub cbrThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    
    Case conMenu_File_Print, conMenu_File_Preview  '报告打印,预览（预览中也可打印）
        Control.Enabled = InStr(mstrPrivs, "报告打印") > 0
    Case conMenu_File_Exit
        Control.Visible = mblnShowBorder
    End Select
End Sub

Private Sub cmdOK_Click()
    Call zlRefresh
End Sub

Private Sub cmd项目_Click()
    Call ShowSelect
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionClosed Then Cancel = True
    If Pane.ID = Dkp_ID_Append Then Cancel = False

End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    
    Me.cbrThis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    Top = lngTop
    Bottom = Me.ScaleHeight - lngBottom
End Sub

Private Sub dkpMain_Resize()
    Me.cbrThis.RecalcLayout
    
    Call ImageTypeSet(Me.VScroll.Max)
End Sub



Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Dkp_ID_Request
        Item.Handle = Me.PicInfo.hWnd
    Case Dkp_ID_Append
        Item.Handle = Me.PicTab.hWnd
    Case Dkp_ID_Image
        Item.Handle = Me.PicImage.hWnd
    End Select
End Sub

Private Sub cbrthis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub


Private Sub Form_Activate()
'   Call rptList_SelectionChanged '触发选择事件
End Sub

Private Sub Form_Load()

    On Error Resume Next
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    

    
    If Dir(App.Path & "\zlLisPic*.Bmp") <> "" Then
        Kill App.Path & "\zlLisPic*.Bmp"
    End If
    If Dir(App.Path & "\*.cht") <> "" Then Kill App.Path & "\*.cht"
    '=====================================================
    'Call RestoreWinState(Me, App.ProductName)                   '界面恢复

    Set mfrmLisRptGeneral = frmLisRptGeneral                     '普通标本窗体
    Set mfrmLisRptMicrobiology = frmLisRptMicrobiology           '微生物标本窗体
    mfrmLisRptGeneral.mlngMod = mlngMod
    mfrmLisRptMicrobiology.mlngMode = mlngMod
    mstrPrivs = IIF(Right(gMainPrivs, 1) = ";", gMainPrivs, gMainPrivs & ";") & gcolPrivs(glngSys & "_" & mlngMod)
    CreateCbs                           '创建工具条
    CreateDockPane                      '创建浮动窗体
    CreateTableControl                  '创建TAB
    CreateRptListHead
    mstrWhere = ""
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    End If
    
    mlngPageId = GetLastPageId(mlng病人ID)
    Call LoadAllData(mlng病人ID)
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane
    Dim intLoop As Integer
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub

    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Append)
    Pane1.MinTrackSize.SetSize 6954 / Screen.TwipsPerPixelX, 380 / Screen.TwipsPerPixelY
    Pane1.MaxTrackSize.SetSize Pane1.MaxTrackSize.Width, 380 / Screen.TwipsPerPixelY
    
    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Request)
    Pane1.MinTrackSize.SetSize 3080 / Screen.TwipsPerPixelX, 2295 / Screen.TwipsPerPixelY
    Pane1.MaxTrackSize.SetSize 3980 / Screen.TwipsPerPixelX, 2295 / Screen.TwipsPerPixelY
    
    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Image)
    Pane1.MinTrackSize.SetSize 1880 / Screen.TwipsPerPixelX, 500 / Screen.TwipsPerPixelY
    
    Me.dkpMain.RecalcLayout
    Me.dkpMain.NormalizeSplitters
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnCheck As Boolean
    blnCheck = Me.cbrThis.FindControl(, conMenu_Manage_LeaveMedi, , True).Checked
    
    Call zlDatabase.SetPara("隐藏检验图形", IIF(blnCheck, "True", "False"), glngSys, mlngMod)
    Call SaveWinState(Me, App.ProductName)
    Unload mfrmLisRptGeneral
    Unload mfrmLisRptMicrobiology
    Set mfrmLisRptGeneral = Nothing
    Set mfrmLisRptMicrobiology = Nothing
    
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    End If
    mlng病人ID = 0
    mlng主页ID = 0
    mlng医嘱ID = 0
    mlng标本ID = 0
    mlng结果次数 = 0
End Sub


Private Sub picFind_Resize()
    On Error Resume Next
    dtpStart.Width = (picFind.ScaleWidth - dtpStart.Left - 90) / 2
    dtpEnd.Left = dtpStart.Left + dtpStart.Width + 45
    dtpEnd.Width = picFind.ScaleWidth - dtpEnd.Left - 45
    
    cmdOK.Left = picFind.ScaleWidth - cmdOK.Width - 45
    
    cmd项目.Left = cmdOK.Left - 45 - cmd项目.Width
    txt项目.Width = cmd项目.Left - txt项目.Left - 10
    
End Sub

Private Sub PicInfo_Resize()
    On Error Resume Next

    picFind.Top = 0
    picFind.Left = 0
    picFind.Width = PicInfo.ScaleWidth

    Me.rptList.Left = 0
    Me.rptList.Top = picFind.Top + picFind.Height

    Me.rptList.Width = PicInfo.ScaleWidth
    Me.rptList.Height = PicInfo.ScaleHeight - Me.rptList.Top
    

End Sub

Private Sub picList_Resize()

End Sub

Private Sub picTab_Resize()
    
    Me.TabCtlWindow.Top = 0
    Me.TabCtlWindow.Left = 0
    Me.TabCtlWindow.Width = Me.PicTab.ScaleWidth
    Me.TabCtlWindow.Height = Me.PicTab.ScaleHeight
    Call ImageTypeSet(VScroll.Max)
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    On Error Resume Next
    If Button = 2 Then
        If rptList.Records.Count <= 0 Then Exit Sub
        If Not rptList.SelectedRows(0).GroupRow Then
            Set objPopup = cbrThis.ActiveMenuBar.FindControl(, ID_MENU_MOUSE)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub rptList_SelectionChanged()
    On Error GoTo errHandle
    Dim i As Integer
     '---------------改变选择前先清空所有记录----------------------
'    Call mfrmLisRptGeneral.zlRefresh(0)
'    Call mfrmLisRptMicrobiology.zlRefresh(0, 0)
'    Call ReadImageData(0, False)
    '-------------------------------------------------------------
    If rptList.SelectedRows.Count = 0 Then
        If rptList.Rows.Count > 0 Then
            '有记录,取第个非分组行,做当前行
            For i = 0 To rptList.Rows.Count - 1
                If Not rptList.Rows(i).GroupRow Then
                    rptList.Rows(i).Selected = True
                    
                    mlng医嘱ID = Val(rptList.Rows(i).Record(mCol.医嘱ID).value)
                    mlng标本ID = Val(rptList.Rows(i).Record(mCol.ID).value)
                    mlng结果次数 = Val(rptList.Rows(i).Record(mCol.结果次数).value)
                    
                    If rptList.Rows(i).Record(mCol.微生物标本).value = "0" Then
                        Me.TabCtlWindow.Item(0).Visible = True
                        Me.TabCtlWindow.Item(1).Visible = False
                        Me.TabCtlWindow.Item(0).Selected = True
                        mfrmLisRptGeneral.zlRefresh (mlng医嘱ID)
                        ReadImageData mlng标本ID, False
                    Else
                        Me.TabCtlWindow.Item(0).Visible = False
                        Me.TabCtlWindow.Item(1).Visible = True
                        Me.TabCtlWindow.Item(1).Selected = True
                        Call mfrmLisRptMicrobiology.zlRefresh(mlng标本ID, mlng结果次数)
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    If rptList.FocusedRow Is Nothing Then
        mlng医嘱ID = 0
        mlng标本ID = 0
        mlng结果次数 = 0
        Exit Sub
    End If
    If rptList.FocusedRow.GroupRow Then Exit Sub
    
    If Val(rptList.FocusedRow.Record(mCol.医嘱ID).value) <> 0 And _
      (mlng医嘱ID <> Val(rptList.FocusedRow.Record(mCol.医嘱ID).value) Or _
      mlng标本ID <> Val(rptList.FocusedRow.Record(mCol.ID).value)) Then
        mlng医嘱ID = Val(rptList.FocusedRow.Record(mCol.医嘱ID).value)
        mlng标本ID = Val(rptList.FocusedRow.Record(mCol.ID).value)
        mlng结果次数 = Val(rptList.FocusedRow.Record(mCol.结果次数).value)
        If rptList.FocusedRow.Record(mCol.微生物标本).value = "0" Then
            Me.TabCtlWindow.Item(0).Visible = True
            Me.TabCtlWindow.Item(1).Visible = False
            Me.TabCtlWindow.Item(0).Selected = True
            mfrmLisRptGeneral.zlRefresh (mlng医嘱ID)
            
        Else
            Me.TabCtlWindow.Item(0).Visible = False
            Me.TabCtlWindow.Item(1).Visible = True
            Me.TabCtlWindow.Item(1).Selected = True
            Call mfrmLisRptMicrobiology.zlRefresh(mlng标本ID, mlng结果次数)
        End If
    Else
        mlng医嘱ID = 0
        mfrmLisRptGeneral.zlRefresh (0)
        Call mfrmLisRptMicrobiology.zlRefresh(0, 0)
    End If
    ReadImageData mlng标本ID, False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
    Resume
    End If
End Sub

Private Sub zlRefresh()
    '点刷新时调用
    If dtpStart.value > dtpEnd.value Then
        MsgBox "查询开始日期不能大于结束日期！"
        Exit Sub
    End If
    mstrWhere = Format(dtpStart.value, "yyyy-MM-dd") & "|" & Format(dtpEnd.value, "yyyy-MM-dd") & "|" & Trim(txt项目.Text)
    Call LoadAllData(mlng病人ID)
End Sub

Private Function ShowSelect()

    Dim vRect As RECT, strSQL As String, rsTmp As ADODB.Recordset
    Dim str输入 As String, blnCanel As Boolean, strSel项目 As String
    Dim strInput As String
    
    On Error GoTo errHandle
    str输入 = Trim(txt项目.Text)
    If Trim(str输入) <> "" Then
        str输入 = Replace(str输入, "%", "")
        str输入 = Replace(str输入, "'", "")
        str输入 = Replace(UCase(str输入), "AND", "")
        str输入 = Replace(UCase(str输入), "OR", "")
    Else
        str输入 = ""
    End If
    vRect = GetControlRect(txt项目.hWnd)

    
    If str输入 <> "" Then
        strInput = " And (B.简码 Like [1] Or A.编码 Like [1] Or A.名称 Like [1])"
        If IsNumeric(str输入) Then
            '1X.输入全是数字时只匹配编码
            If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.编码 Like [1]"
        ElseIf zlCommFun.IsCharAlpha(str输入) Then
            'X1.输入全是字母时只匹配简码
            If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.简码 Like [1]"
        ElseIf zlCommFun.IsCharChinese(str输入) Then
            '包含汉字,则只匹配名称
            strInput = " And A.名称 Like [1]"
        End If
        
        str输入 = IIF(Len(str输入) < 3, "", mstrLike) & str输入 & "%"
        strSQL = "Select Distinct A.ID, A.操作类型 As 类型, A.编码, A.名称, Decode(A.组合项目, 1, '是', '否') As 组合项目" & vbNewLine & _
                "From 诊疗项目目录 A, 诊疗项目别名 B" & vbNewLine & _
                "Where A.ID = B.诊疗项目id And A.类别 = 'C' And A.单独应用 = 1 And" & vbNewLine & _
                "      (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & strInput
    
    Else
        strSQL = "Select A.ID, A.操作类型 As 类型, A.编码, A.名称, Decode(A.组合项目, 1, '是', '否') As 组合项目" & vbNewLine & _
                "From 诊疗项目目录 A" & vbNewLine & _
                "Where A.类别 = 'C' And A.单独应用 = 1 And" & vbNewLine & _
                "      (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) "
    End If
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "项目", False, "", "选择项目", False, False, True, _
                                         vRect.Left, vRect.Top, txt项目.Height, blnCanel, False, True, str输入)
    
    If Not blnCanel And Not rsTmp Is Nothing Then
        Do Until rsTmp.EOF
            strSel项目 = strSel项目 & "," & rsTmp!名称
            rsTmp.MoveNext
        Loop
        txt项目.Text = ""
        If strSel项目 <> "" Then txt项目.Text = Mid(strSel项目, 2)
    Else
        If txt项目.Enabled Then
            txt项目.SelStart = 0: txt项目.SelLength = Len(txt项目.Text)
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Private Sub txt项目_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ShowSelect
End Sub

Private Sub VScroll_Change()
    Dim intLoop As Integer
    If Me.Visible = False Then Exit Sub
    For intLoop = 0 To Me.VScroll.Max
        If intLoop < Me.VScroll.value Then
            Me.ChartThis(intLoop).Visible = False
        Else
            Me.ChartThis(intLoop).Visible = True
            If intLoop = Me.VScroll.value Then
                Me.ChartThis(intLoop).Top = 0
            Else
                Me.ChartThis(intLoop).Top = Me.ChartThis(intLoop - 1).Top + Me.ChartThis(intLoop - 1).Height + 10
            End If
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/5/25
'功    能:调用API动态设置窗体的border
'入    参:
'           new_Hwnd    窗体的句柄
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub YSystemMenu(ByVal new_Hwnd As Long)
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 'Or WS_SYSMENU Or &H20000
End Sub
