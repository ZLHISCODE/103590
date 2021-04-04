VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPriceGradeManage 
   Caption         =   "价格等级管理"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10530
   Icon            =   "frmPriceGradeManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPriceGrade 
      BorderStyle     =   0  'None
      Height          =   3585
      Left            =   1410
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1830
      Width           =   2505
      Begin MSComctlLib.ListView lvwPriceGrade 
         Height          =   1905
         Left            =   300
         TabIndex        =   5
         Top             =   990
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   3360
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils32"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin XtremeSuiteControls.ShortcutCaption sccPriceGrade 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   540
         Width           =   1335
         _Version        =   589884
         _ExtentX        =   2355
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "价格等级"
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
      End
      Begin VB.Shape shpPriceGrade 
         BorderColor     =   &H80000003&
         Height          =   255
         Left            =   210
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.PictureBox picGradeApply 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   4170
      ScaleHeight     =   3615
      ScaleWidth      =   2325
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2325
      Begin VSFlex8Ctl.VSFlexGrid vsfGradeApply 
         Height          =   2145
         Left            =   270
         TabIndex        =   2
         Top             =   930
         Width           =   1785
         _cx             =   3149
         _cy             =   3784
         Appearance      =   0
         BorderStyle     =   0
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
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
      Begin VB.Shape shpGradeApply 
         BorderColor     =   &H80000003&
         Height          =   345
         Left            =   270
         Top             =   180
         Width           =   495
      End
      Begin XtremeSuiteControls.ShortcutCaption sccGradeApply 
         Height          =   300
         Left            =   210
         TabIndex        =   3
         Top             =   630
         Width           =   1335
         _Version        =   589884
         _ExtentX        =   2355
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "价格等级应用"
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
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7275
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   635
      SimpleText      =   $"frmPriceGradeManage.frx":030A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPriceGradeManage.frx":0351
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13494
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
   Begin MSComctlLib.ImageList ils32 
      Left            =   2625
      Top             =   345
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":0BE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":103D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":1357
            Key             =   "Default"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":1C31
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1980
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":250B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":2963
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":2C7D
            Key             =   "Default"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":3217
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPriceGradeManage.frx":37B1
      Left            =   1080
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPriceGradeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModule As Long
Private mblnUnload As Boolean
Private mblnFirst As Boolean

Private Enum PaneIndex
    Pane_PriceGrade = 1
    Pane_GradeApply = 2
End Enum

Private Enum ColIndex
    'LVW_名称 = 0
    LVW_编码 = 1
    LVW_简码 = 2
    LVW_适用药品 = 3
    LVW_适用卫材 = 4
    LVW_适用普通项目 = 5
    LVW_建档时间 = 6
    LVW_撤档时间 = 7
    LVW_是否停用 = 8
    
    VSF_应用场合 = 0
    VSF_编码 = 1
    VSF_名称 = 2
End Enum

Private mblnShowStopedGrade As Boolean
Private mbytLvwViewType As Byte

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnStop As Boolean
    
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_EditPopup '编辑
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "增删改") _
                        Or zlStr.IsHavePrivs(mstrPrivs, "停用") _
                        Or zlStr.IsHavePrivs(mstrPrivs, "启用")
    Case conMenu_Edit_NewItem '增加
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "增删改")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify '调整
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "增删改")
        Control.Enabled = Control.Visible And Not lvwPriceGrade.SelectedItem Is Nothing
        If Control.Enabled Then
            Control.Enabled = Val(lvwPriceGrade.SelectedItem.SubItems(LVW_是否停用)) = 0
        End If
    Case conMenu_Edit_Delete '删除
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "增删改")
        Control.Enabled = Control.Visible And Not lvwPriceGrade.SelectedItem Is Nothing
        If Control.Enabled Then
            Control.Enabled = Val(lvwPriceGrade.SelectedItem.SubItems(LVW_是否停用)) = 0
        End If
    Case conMenu_Edit_Reuse '停用
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "停用")
        Control.Enabled = Control.Visible And Not lvwPriceGrade.SelectedItem Is Nothing
        If Control.Enabled Then
            Control.Enabled = Val(lvwPriceGrade.SelectedItem.SubItems(LVW_是否停用)) = 0
        End If
    Case conMenu_Edit_Stop '启用
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "启用")
        Control.Enabled = Control.Visible And Not lvwPriceGrade.SelectedItem Is Nothing
        If Control.Enabled Then
            Control.Enabled = Val(lvwPriceGrade.SelectedItem.SubItems(LVW_是否停用)) = 1
        End If
    End Select
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Form_Activate()
    Err = 0: On Error GoTo ErrHandler
    If mblnUnload Then Unload Me: Exit Sub
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Exit Sub
ErrHandler:
    mblnUnload = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim strLvwCols As String
    Err = 0: On Error GoTo ErrHandler
    mblnUnload = False
    mblnFirst = True
    mstrPrivs = gstrPrivs
    mlngModule = glngModul
    
    mblnShowStopedGrade = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示停用", 0)) = 1
    mbytLvwViewType = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "ListView视图", 0))
    
    If DefMainCommandBars() = False Then mblnUnload = True: Exit Sub
    If InitPanel() = False Then mblnUnload = True: Exit Sub
    If InitVsfGrid() = False Then mblnUnload = True: Exit Sub
    
    Call RestoreWinState(Me, App.ProductName)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
    
    '如果ListView的还未被设置，比如第一次使用，那就调用缺省的初始化
    strLvwCols = "名称,1600,0,1;编码,1000,0,1;简码,1200,0,0;" & _
        "适用药品,0,2,0;适用卫材,0,2,0;适用普通项目,0,2,0;建档时间,1900,2,0;撤档时间,1900,2,0;是否停用,0,0,1"
    If lvwPriceGrade.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwPriceGrade, strLvwCols, True
    End If
    Call SetLvwViewType(mbytLvwViewType)
    Call LoadPriceGrade
    Exit Sub
ErrHandler:
    mblnUnload = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitPanel() As Boolean
    '功能:初始化界面布局
    '返回:设置成功,返回true,否则返回False
    Dim objPane As Pane

    Err = 0: On Error GoTo ErrHandler
    Set objPane = dkpMain.CreatePane(Pane_PriceGrade, 200, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.MinTrackSize.Width = 65

    Set objPane = dkpMain.CreatePane(Pane_GradeApply, 700, 400, DockRightOf, objPane)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.MinTrackSize.Width = 65

    With dkpMain
        .SetCommandBars cbsMain
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    InitPanel = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DefMainCommandBars() As Boolean
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    Dim objPopupControl As CommandBarControl

    Err = 0: On Error GoTo ErrHandler
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsMain.EnableCustomization False

    '菜单定义
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "调整(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "停用(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "启用(&T)")
    End With

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With cbrControl.CommandBar.Controls
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False): cbrSubControl.Checked = True
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False): cbrSubControl.Checked = True
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False): cbrSubControl.Checked = True
        End With
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): cbrControl.BeginGroup = True
        cbrControl.Checked = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "大图标(&G)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "小图标(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "列表(&L)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "详细资料(&D)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "显示停用等级(&S)"): cbrControl.BeginGroup = True
        cbrControl.Checked = mblnShowStopedGrade
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的中联")
        With cbrControl.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With

    '工具栏定义
    Set cbrToolBar = cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "调整")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "停用"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "启用")
        
        Set objPopupControl = .Add(xtpControlSplitButtonPopup, conMenu_View_Append, "查看"): objPopupControl.BeginGroup = True
        objPopupControl.IconId = conMenu_View_LargeICO
        With objPopupControl.CommandBar.Controls
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_LargeICO, "大图标")
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_MinICO, "小图标")
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_ListICO, "列表")
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "详细资料")
        End With
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    '快键绑定
    With cbsMain.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("B"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
    End With

    '设置不常用菜单
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
    End With

    DefMainCommandBars = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitVsfGrid() As Boolean
    '功能：初始化网格控件
    Dim strHead As String, varData As Variant
    Dim i As Long

    Err = 0: On Error GoTo ErrHandler
    With vsfGradeApply
        .redraw = flexRDNone
        .Rows = 1
        .FixedCols = 0: .FixedRows = 1
        '
        strHead = ",1,250|编码,4,700|名称,1,2000"
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .ColKey(i) = Split(varData(i), ",")(0)
        Next
        .FixedAlignment(-1) = flexAlignCenterCenter
        .RowHeightMin = 300
        '.ColHidden(VSF_应用场合) = True'不能隐藏"应用场合"列，隐藏后分组不能展开和收缩，缺省设置"应用场合"宽度为10

        .AllowSelection = False
        .AllowBigSelection = False
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow

        .HighLight = flexHighlightAlways
        .AllowUserResizing = flexResizeColumns
        .GridLines = flexGridFlat
        .PicturesOver = True '文字在图片上面

        .BackColorBkg = vbWindowBackground
        .SheetBorder = vbWindowBackground
        
'        '列属性设置,用于用户选择显示列
'        For i = 0 To .Cols - 1
'            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)|列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
'            Select Case i
'            Case LVW_ID
'                 .ColData(i) = "-1|1"
'            Case LVW_号码
'                .ColData(i) = "1|0"
'            End Select
'        Next
        .redraw = flexRDBuffered
    End With
    InitVsfGrid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim frmEdit As frmPriceGradeEdit
    Dim strItem As String

    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_PrintSet '打印设置
        Call zlPrintSet
    Case conMenu_File_Preview '预览
        Call ZlDataPrint(2)
    Case conMenu_File_Print '打印
        Call ZlDataPrint(1)
    Case conMenu_File_Excel '输出到Excel…
        Call ZlDataPrint(3)
    Case conMenu_File_Exit '退出
        Unload Me
    Case conMenu_Edit_NewItem '增加
        Set frmEdit = New frmPriceGradeEdit
        If frmEdit.ShowMe(Me, 0, , strItem) Then Call LoadPriceGrade(strItem)
    Case conMenu_Edit_Modify '调整
        If lvwPriceGrade.SelectedItem Is Nothing Then Exit Sub
        Set frmEdit = New frmPriceGradeEdit
        strItem = lvwPriceGrade.SelectedItem.Text
        If frmEdit.ShowMe(Me, 1, strItem) Then Call LoadPriceGrade
    Case conMenu_Edit_Delete '删除
        If lvwPriceGrade.SelectedItem Is Nothing Then Exit Sub
        Set frmEdit = New frmPriceGradeEdit
        strItem = lvwPriceGrade.SelectedItem.Text
        If frmEdit.ShowMe(Me, 2, strItem) Then Call LoadPriceGrade
    Case conMenu_Edit_Reuse '停用
        If lvwPriceGrade.SelectedItem Is Nothing Then Exit Sub
        strItem = lvwPriceGrade.SelectedItem.Text
        If StopAndStartPriceGrade(strItem, True) Then Call LoadPriceGrade
    Case conMenu_Edit_Stop '启用
        If lvwPriceGrade.SelectedItem Is Nothing Then Exit Sub
        strItem = lvwPriceGrade.SelectedItem.Text
        If StopAndStartPriceGrade(strItem, False) Then Call LoadPriceGrade
    Case conMenu_View_ToolBar_Button '标准按钮
        Control.Checked = Not Control.Checked
        cbsMain(2).Visible = Control.Checked
        Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Text, , True)
        objControl.Enabled = Control.Checked
        Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Size, , True)
        objControl.Enabled = Control.Checked
        cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '文本标签
        Control.Checked = Not Control.Checked
        For Each objControl In cbsMain(2).Controls
            objControl.Style = IIF(Control.Checked, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Not Control.Checked
        cbsMain.Options.LargeIcons = Control.Checked
        cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Not Control.Checked
        stbThis.Visible = Control.Checked
        cbsMain.RecalcLayout
    Case conMenu_View_Append
        mbytLvwViewType = mbytLvwViewType + 1
        If mbytLvwViewType < lvwIcon Or mbytLvwViewType > lvwReport Then
            mbytLvwViewType = lvwIcon
        End If
        Call SetLvwViewType(mbytLvwViewType)
    Case conMenu_View_LargeICO '大图标
        Call SetLvwViewType(lvwIcon)
        mbytLvwViewType = lvwIcon
    Case conMenu_View_MinICO '小图标
        Call SetLvwViewType(lvwSmallIcon)
        mbytLvwViewType = lvwSmallIcon
    Case conMenu_View_ListICO '列表
        Call SetLvwViewType(lvwList)
        mbytLvwViewType = lvwList
    Case conMenu_View_DetailsICO '详细资料
        Call SetLvwViewType(lvwReport)
        mbytLvwViewType = lvwReport
    Case conMenu_View_ShowStoped '显示停用等级
        Control.Checked = Not Control.Checked
        zlDatabase.SetPara "显示停用等级", IIF(Control.Checked, 1, 0), glngSys, mlngModule
        mblnShowStopedGrade = Control.Checked
        cbsMain.RecalcLayout
        Call LoadPriceGrade
    Case conMenu_View_Refresh '刷新
        Call LoadPriceGrade
    Case conMenu_Help_Help '帮助主题
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home '中联主页
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '关于…
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '调用自定义报表
            Call ZlCallCustomReprot(Me, Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
        End If
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetLvwViewType(ByVal bytView As Byte)
    '功能:更改ListView控件查看菜单状态，以及设置ListView控件查看方式
    '入参：
    '   bytView: 0-lvwIcon （缺省）图标
    '            1-lvwSmallIcon  小图标
    '            2-lvwList 列表
    '            3-lvwReport 报表
    Dim objControl As CommandBarControl
    Dim objPopupControl As CommandBarControl
    
    Err = 0: On Error GoTo ErrHandler
    '菜单栏
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Find(, conMenu_View_LargeICO, , True): objControl.Checked = (bytView = lvwIcon)
        Set objControl = .Find(, conMenu_View_MinICO, , True): objControl.Checked = (bytView = lvwSmallIcon)
        Set objControl = .Find(, conMenu_View_ListICO, , True): objControl.Checked = (bytView = lvwList)
        Set objControl = .Find(, conMenu_View_DetailsICO, , True): objControl.Checked = (bytView = lvwReport)
    End With
    '工具栏
    With cbsMain(2).Controls
        Set objPopupControl = .Find(, conMenu_View_Append, , True)
        Set objControl = .Find(, conMenu_View_LargeICO, , True): objControl.Checked = (bytView = lvwIcon)
        Set objControl = .Find(, conMenu_View_MinICO, , True): objControl.Checked = (bytView = lvwSmallIcon)
        Set objControl = .Find(, conMenu_View_ListICO, , True): objControl.Checked = (bytView = lvwList)
        Set objControl = .Find(, conMenu_View_DetailsICO, , True): objControl.Checked = (bytView = lvwReport)
    End With

    Select Case bytView
    Case lvwIcon
        objPopupControl.IconId = conMenu_View_LargeICO
        lvwPriceGrade.View = lvwIcon
    Case lvwSmallIcon
        objPopupControl.IconId = conMenu_View_MinICO
        lvwPriceGrade.View = lvwSmallIcon
    Case lvwList
        objPopupControl.IconId = conMenu_View_ListICO
        lvwPriceGrade.View = lvwList
    Case lvwReport
        objPopupControl.IconId = conMenu_View_DetailsICO
        lvwPriceGrade.View = lvwReport
    End Select
    cbsMain.RecalcLayout
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ZlCallCustomReprot(ByVal frmMain As Form, ByVal lngSys As Long, strReprotName As String)
    '功能:调用相关的自定义报表
    Err = 0: On Error GoTo ErrHandler
    Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Err = 0: On Error GoTo ErrHandler
    Select Case Item.ID
    Case Pane_PriceGrade
        Item.Handle = picPriceGrade.hwnd
    Case Pane_GradeApply
        Item.Handle = picGradeApply.hwnd
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error GoTo ErrHandler
    
    mblnUnload = False
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "显示停用", IIF(mblnShowStopedGrade, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "ListView视图", mbytLvwViewType
    Call SaveWinState(Me, App.ProductName)
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ZlDataPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objVsfPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim objLvwPrint As New zlPrintLvw
    Dim bytR As Byte
    
    Err = 0: On Error GoTo ErrHandler
    If Me.ActiveControl Is vsfGradeApply Then
        'VSFlexGrid
        Set objVsfPrint.Body = vsfGradeApply
        objVsfPrint.Title.Text = "价格等级应用"
        
        objVsfPrint.Title.Font.Name = "楷体_GB2312"
        objVsfPrint.Title.Font.Size = 18
        objVsfPrint.Title.Font.Bold = True
        
        
        Set objRow = New zlTabAppRow
        objRow.Add "打印人：" & gstrUserName
        objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
        objVsfPrint.BelowAppRows.Add objRow

        If bytMode = 1 Then
            bytR = zlPrintAsk(objVsfPrint)
            If bytR <> 0 Then zlPrintOrView1Grd objVsfPrint, bytR
        Else
            zlPrintOrView1Grd objVsfPrint, bytMode
        End If
    Else
        'ListView
        Set objLvwPrint.Body.objData = lvwPriceGrade
        objLvwPrint.Title.Text = "价格等级"

        objLvwPrint.Title.Font.Name = "楷体_GB2312"
        objLvwPrint.Title.Font.Size = 18
        objLvwPrint.Title.Font.Bold = True
        
        objLvwPrint.BelowAppItems.Add "打印人：" & gstrUserName
        objLvwPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
        
        If bytMode = 1 Then
            bytR = zlPrintAsk(objLvwPrint)
            If bytR <> 0 Then zlPrintOrViewLvw objLvwPrint, bytR
        Else
            zlPrintOrViewLvw objLvwPrint, bytMode
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwPriceGrade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Err = 0: On Error Resume Next
    '点击列标头排序
    lvwPriceGrade.Sorted = True
    lvwPriceGrade.SortKey = ColumnHeader.Index - 1
    lvwPriceGrade.SortOrder = IIF(lvwPriceGrade.SortOrder = lvwDescending, lvwAscending, lvwDescending)
End Sub

Private Sub lvwPriceGrade_DblClick()
    Dim strItem As String
    Dim frmEdit As New frmPriceGradeEdit
    
    Err = 0: On Error GoTo ErrHandler
    If lvwPriceGrade.SelectedItem Is Nothing Then Exit Sub
    strItem = lvwPriceGrade.SelectedItem.Text
    
    If zlStr.IsHavePrivs(mstrPrivs, "增删改") _
        And Val(lvwPriceGrade.SelectedItem.SubItems(LVW_是否停用)) = 0 Then '调整
        If frmEdit.ShowMe(Me, 1, strItem) Then Call LoadPriceGrade
    Else '查看
        frmEdit.ShowMe Me, 3, strItem
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwPriceGrade_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Err = 0: On Error GoTo ErrHandler
    If lvwPriceGrade.Tag = Item.Text Then Exit Sub
    lvwPriceGrade.Tag = Item.Text
    Call LoadPriceGradeApply(Item.Text)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwPriceGrade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo ErrHandler
'    If lvwPriceGrade.Visible And lvwPriceGrade.Enabled Then lvwPriceGrade.SetFocus
    If Not (Button = vbRightButton) Then Exit Sub
    If Not Me.ActiveControl Is lvwPriceGrade Then Exit Sub
    
    Set objPopup = cbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picPriceGrade_Resize()
    Err = 0: On Error Resume Next
    shpPriceGrade.Move 0, 0, picPriceGrade.ScaleWidth, picPriceGrade.ScaleHeight
    sccPriceGrade.Move 10, 10, picPriceGrade.ScaleWidth - 30
    With lvwPriceGrade
        .Left = sccPriceGrade.Left
        .Top = sccPriceGrade.Top + sccPriceGrade.Height
        .Width = picPriceGrade.ScaleWidth - .Left - 20
        .Height = picPriceGrade.ScaleHeight - .Top - 20
    End With
End Sub

Private Sub picGradeApply_Resize()
    Err = 0: On Error Resume Next
    shpGradeApply.Move 0, 0, picGradeApply.ScaleWidth, picGradeApply.ScaleHeight
    sccGradeApply.Move 10, 10, picGradeApply.ScaleWidth - 30
    With vsfGradeApply
        .Left = sccGradeApply.Left
        .Top = sccGradeApply.Top + sccGradeApply.Height
        .Width = picGradeApply.ScaleWidth - .Left - 20
        .Height = picGradeApply.ScaleHeight - .Top - 20
    End With
End Sub

Private Sub vsfGradeApply_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = VSF_应用场合 Then Cancel = True
End Sub

Private Sub vsfGradeApply_DblClick()
    Call lvwPriceGrade_DblClick
End Sub

Private Sub vsfGradeApply_GotFocus()
    vsfGradeApply.BackColorSel = vbHighlight
End Sub

Private Sub vsfGradeApply_LostFocus()
    vsfGradeApply.BackColorSel = &HE0E0E0
End Sub

Private Function LoadPriceGrade(Optional ByVal strSelectItem As String) As Boolean
    '加载价格等级
    '入参：
    '   strSelectItem 缺省选择项目,收费价格等级名称
    Dim strSQL As String, strWhere As String
    Dim rsData As ADODB.Recordset
    Dim objListItem As ListItem, i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    If strSelectItem = "" Then
        If Not lvwPriceGrade.SelectedItem Is Nothing Then
            strSelectItem = lvwPriceGrade.SelectedItem.Text
        End If
    End If
    
    lvwPriceGrade.ListItems.Clear
    lvwPriceGrade.Tag = ""
    vsfGradeApply.Clear 1: vsfGradeApply.Rows = vsfGradeApply.FixedRows
    If mblnShowStopedGrade = False Then
        '不显示停用价格等级
        strWhere = " And (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01','yyyy-mm-dd'))"
    End If
    strSQL = "Select 编码, 名称, 简码, 是否适用药品, 是否适用卫材, 是否适用普通项目, 建档时间, 撤档时间," & vbNewLine & _
            "        Decode(Nvl(撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')), To_Date('3000-01-01', 'yyyy-mm-dd'), 0, 1) As 是否停用" & vbNewLine & _
            " From 收费价格等级" & vbNewLine & _
            " Where 1 = 1 " & strWhere & vbNewLine & _
            " Order By 编码"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "收费价格等级")
    If rsData.RecordCount = 0 Then LoadPriceGrade = True: Exit Function
    
    '名称,编码,简码,适用药品,适用卫材,适用普通项目,建档时间,撤档时间
    Do While Not rsData.EOF
        Set objListItem = lvwPriceGrade.ListItems.Add(, "K" & Nvl(rsData!编码), Nvl(rsData!名称), "Default", "Default")
        objListItem.SubItems(LVW_编码) = Nvl(rsData!编码)
        objListItem.SubItems(LVW_简码) = Nvl(rsData!简码)
        objListItem.SubItems(LVW_适用药品) = IIF(Val(Nvl(rsData!是否适用药品)) = 1, "√", "")
        objListItem.SubItems(LVW_适用卫材) = IIF(Val(Nvl(rsData!是否适用卫材)) = 1, "√", "")
        objListItem.SubItems(LVW_适用普通项目) = IIF(Val(Nvl(rsData!是否适用普通项目)) = 1, "√", "")
        objListItem.SubItems(LVW_建档时间) = Format(Nvl(rsData!建档时间), "yyyy-mm-dd hh:mm:ss")
        objListItem.SubItems(LVW_撤档时间) = Format(Nvl(rsData!撤档时间), "yyyy-mm-dd hh:mm:ss")
        objListItem.SubItems(LVW_是否停用) = Val(Nvl(rsData!是否停用))
        If Val(Nvl(rsData!是否停用)) = 1 Then
            '改变停用价格等级的图标和字体颜色
            objListItem.Icon = "Stop"
            objListItem.SmallIcon = "Stop"
            objListItem.ForeColor = vbRed
            For i = 1 To objListItem.ListSubItems.Count
                objListItem.ListSubItems(i).ForeColor = vbRed
            Next
        End If
        If Nvl(rsData!名称) = strSelectItem Then
            objListItem.Selected = True
        End If
        rsData.MoveNext
    Loop
    If Not lvwPriceGrade.SelectedItem Is Nothing Then
        Call lvwPriceGrade_ItemClick(lvwPriceGrade.SelectedItem)
    End If
    LoadPriceGrade = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadPriceGradeApply(ByVal strPriceGrade As String) As Boolean
    '加载价格等级应用
    '入参：
    '   strPriceGrade 收费价格等级名称
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRow As Long
    Dim i  As Long, j  As Long, strTemp As String
    
    Err = 0: On Error GoTo ErrHandler
    With vsfGradeApply
        .redraw = flexRDNone
        .Clear 1
        .Rows = 1
        
        strSQL = "Select Decode(Nvl(a.性质, 0), 0, '院区', '医疗付款方式') As 应用场合," & vbNewLine & _
                "        Decode(Nvl(a.性质, 0), 0, b.编号, c.编码) As 编码," & vbNewLine & _
                "        Decode(Nvl(a.性质, 0), 0, b.名称, c.名称) As 名称" & vbNewLine & _
                " From 收费价格等级应用 A, Zlnodelist B, 医疗付款方式 C" & vbNewLine & _
                " Where a.站点 = b.编号(+) And a.医疗付款方式 = c.名称(+) And 价格等级 = [1]" & vbNewLine & _
                " Order By a.性质, 编码"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "收费价格等级应用", strPriceGrade)
        If rsData.RecordCount = 0 Then
            .redraw = flexRDBuffered
            LoadPriceGradeApply = True
            Exit Function
        End If
        
        .Rows = rsData.RecordCount + 1
        lngRow = 1
        Do While Not rsData.EOF
            .TextMatrix(lngRow, VSF_应用场合) = Nvl(rsData!应用场合)
            .TextMatrix(lngRow, VSF_编码) = Nvl(rsData!编码)
            .TextMatrix(lngRow, VSF_名称) = Nvl(rsData!名称)
            lngRow = lngRow + 1
            rsData.MoveNext
        Loop
        
        '分组显示
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True

        .Subtotal flexSTNone, VSF_应用场合, , , , , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline VSF_应用场合
        .OutlineCol = VSF_应用场合

        .MergeCells = flexMergeRestrictRows
        .MergeRow(-1) = False
        
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) Then
                .Cell(flexcpText, i, 0, i, .Cols - 1) = .TextMatrix(i + 1, VSF_应用场合)
                .MergeRow(i) = True '该行合并
                .IsCollapsed(i) = flexOutlineExpanded  '是否展开状态
            End If
        Next
        .redraw = flexRDBuffered
    End With
    LoadPriceGradeApply = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function StopAndStartPriceGrade(ByVal str价格等级 As String, _
    ByVal blnStop As Boolean) As Boolean
    '停用/启用价格等级
    '入参：
    '   str价格等级 收费价格等级名称
    '   blnStop 是否停用
    Dim strSQL As String, strWhere As String
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select Decode(Nvl(撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')), To_Date('3000-01-01', 'yyyy-mm-dd'), 0, 1) As 是否停用" & vbNewLine & _
            " From 收费价格等级" & vbNewLine & _
            " Where 名称 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询价格等级", str价格等级)
    If rsTemp.EOF Then
        MsgBox "当前价格等级可能已被他人删除，请刷新后查看...", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If Val(Nvl(rsTemp!是否停用)) = 1 Then
        If blnStop Then
            MsgBox "当前价格等级已被停用，无需再次停用。", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    ElseIf blnStop = False Then
        MsgBox "当前价格等级已是启用状态，无需再启用。", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If blnStop = False Then
        '如果该价格等级启用后，会导致一个站点存在多个有效的价格等级或者一个医疗付款方式存在多个有效的价格等级则不允许启用
        strSQL = "Select Decode(Nvl(a.性质, 0), 0, '院区', '医疗付款方式') As 应用场合," & vbNewLine & _
                "        Decode(Nvl(a.性质, 0), 0, c.名称, a.医疗付款方式) As 名称, a.价格等级" & vbNewLine & _
                " From 收费价格等级应用 A, 收费价格等级应用 B, Zlnodelist C, 收费价格等级 D" & vbNewLine & _
                " Where a.性质 = b.性质 And (a.站点 = b.站点 Or a.医疗付款方式 = b.医疗付款方式) And a.站点 = c.编号(+)" & vbNewLine & _
                "       And a.价格等级 = d.名称 And b.价格等级 = [1] And a.价格等级 <> [1]" & vbNewLine & _
                "       And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询价格等级", str价格等级)
        If Not rsTemp.EOF Then
            Do While Not rsTemp.EOF
                strTemp = strTemp & vbCrLf & Nvl(rsTemp!名称) & "：" & Nvl(rsTemp!价格等级)
                rsTemp.MoveNext
            Loop
            If MsgBox("由于一个院区或一个医疗付款方式只能设置一个有效的价格等级，而你正在启用的价格等级应用中的" & _
                "以下院区或医疗付款方式已设置其它有效的价格等级。如果继续操作，将会清除这些院区或医疗付款方式的其它有效价格等级，" & _
                "然后应用当前价格等级，是否继续？" & vbCrLf & strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    If MsgBox("你确认要" & IIF(blnStop, "停用", "启用") & "名称为“" & str价格等级 & "”的价格等级吗？", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    If blnStop Then '停用
        'Zl_收费价格等级_Stop(
        strSQL = "Zl_收费价格等级_Stop("
        '   名称_In 收费价格等级.名称%Type)
        strSQL = strSQL & "'" & str价格等级 & "')"
    Else '启用
        'Zl_收费价格等级_Start(
        strSQL = "Zl_收费价格等级_Start("
        '   名称_In 收费价格等级.名称%Type)
        strSQL = strSQL & "'" & str价格等级 & "')"
    End If
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    StopAndStartPriceGrade = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
