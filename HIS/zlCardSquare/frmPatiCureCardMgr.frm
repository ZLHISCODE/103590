VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmPatiCureCardMgr 
   Caption         =   "医疗卡管理"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12405
   Icon            =   "frmPatiCureCardMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   12405
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   5700
      ScaleHeight     =   1875
      ScaleWidth      =   5070
      TabIndex        =   25
      Top             =   5115
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   26
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
      TabIndex        =   24
      Top             =   8520
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiCureCardMgr.frx":1CFA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12885
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   180
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiCureCardMgr.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiCureCardMgr.frx":28E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   6765
      Left            =   0
      ScaleHeight     =   6765
      ScaleWidth      =   3690
      TabIndex        =   29
      Top             =   1110
      Width           =   3690
      Begin VB.TextBox txtName 
         Height          =   350
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   2445
      End
      Begin VB.TextBox txtEdit 
         Height          =   350
         Index           =   0
         Left            =   960
         TabIndex        =   17
         Top             =   4215
         Width           =   2445
      End
      Begin VB.TextBox txtEdit 
         Height          =   350
         Index           =   1
         Left            =   960
         TabIndex        =   19
         Top             =   4635
         Width           =   2445
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "过滤(&F)"
         Height          =   350
         Left            =   2325
         TabIndex        =   20
         Top             =   5100
         Width           =   1100
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "按挂失时间查找(&G)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   2925
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker dtp开始日期 
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   8
         Top             =   1920
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   81526787
         CurrentDate     =   40722
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "按发卡日期查找(&S)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   1590
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.TextBox txtCard 
         Height          =   350
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   525
         Width           =   2445
      End
      Begin VB.TextBox txtCard 
         Height          =   350
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   80
         Width           =   2445
      End
      Begin MSComCtl2.DTPicker dtp结束日期 
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   10
         Top             =   2310
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   81526787
         CurrentDate     =   40722
      End
      Begin MSComCtl2.DTPicker dtp开始日期 
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   13
         Top             =   3300
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   81526787
         CurrentDate     =   40722
      End
      Begin MSComCtl2.DTPicker dtp结束日期 
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   15
         Top             =   3735
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   81526787
         CurrentDate     =   40722
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "病人姓名"
         Height          =   180
         Left            =   210
         TabIndex        =   4
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "发卡人"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   16
         Top             =   4290
         Width           =   540
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "挂失人"
         Height          =   180
         Index           =   3
         Left            =   390
         TabIndex        =   18
         Top             =   4710
         Width           =   540
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "结束日期(&M)"
         Height          =   180
         Index           =   1
         Left            =   375
         TabIndex        =   14
         Top             =   3780
         Width           =   990
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         Caption         =   "开始日期(&L)"
         Height          =   180
         Index           =   1
         Left            =   375
         TabIndex        =   12
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "结束日期(&E)"
         Height          =   180
         Index           =   0
         Left            =   375
         TabIndex        =   9
         Top             =   2370
         Width           =   990
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         Caption         =   "开始日期(&S)"
         Height          =   180
         Index           =   0
         Left            =   375
         TabIndex        =   7
         Top             =   1980
         Width           =   990
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "结束卡号"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "开始卡号"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   165
         Width           =   720
      End
   End
   Begin VB.PictureBox picCardList 
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   4305
      ScaleHeight     =   2565
      ScaleWidth      =   8175
      TabIndex        =   27
      Top             =   1410
      Width           =   8175
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   22
         Top             =   0
         Width           =   2280
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCardList 
         Height          =   2055
         Left            =   300
         TabIndex        =   23
         Top             =   465
         Width           =   6825
         _cx             =   12039
         _cy             =   3625
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPatiCureCardMgr.frx":2C36
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
         Begin VB.PictureBox picImgList 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   28
            Top             =   60
            Width           =   210
            Begin VB.Image imgCol 
               Height          =   195
               Left            =   0
               Picture         =   "frmPatiCureCardMgr.frx":2C8F
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   330
         Left            =   720
         TabIndex        =   30
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         Appearance      =   2
         IDKindStr       =   $"frmPatiCureCardMgr.frx":31DD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         DefaultCardType =   "0"
         BackColor       =   -2147483633
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "持卡人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   21
         Top             =   75
         Width           =   630
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   495
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmPatiCureCardMgr.frx":32C0
      Left            =   1515
      Top             =   90
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPatiCureCardMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mblnFirst  As Boolean, mstrPrivs As String, mstrTitle As String    '功能标题
Private mlngModule As Long, mstrKey As String
Private Enum mPgIndex
    Pg_变动记录 = 250101
    Pg_帐户入出记录 = 250102
    Pg_家属关系 = 250103
End Enum
Private Enum mPaneID
    Pane_Search = 1     '搜索条件
    Pane_CardLists = 2  '卡列表
    Pane_CardDetails = 3    '详细列表
End Enum
Private Enum mtxtIdx
    idx_发卡人 = 0
    idx_挂失人 = 1
End Enum
Private WithEvents mobjIDCard As zlIDCard.clsIDCard  '身份证接口
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard   'IC卡接口
Attribute mobjICCard.VB_VarHelpID = -1
Private mlngCardTypeID As Long
Private mPanSearch As Pane
Private mobjSubFrm As Collection
Private mArrFilter As Collection
Private mblnInited As Boolean
Private Const mconMenu_Lable = 3999
Private WithEvents mfrmChage As frmPatiCureCardChangeMgr
Attribute mfrmChage.VB_VarHelpID = -1
Private WithEvents mfrmConsume As frmPatiCureCardConsumeMgr
Attribute mfrmConsume.VB_VarHelpID = -1
Private WithEvents mfrmFamily As frmPatiCureCardFamilyMgr
Attribute mfrmFamily.VB_VarHelpID = -1
Private mcolCard As Collection
Private mblnNotRefresh As Boolean  '不刷新数据
Private mblnNotClick As Boolean
Private mstrPrepayPrivs As String '预交款的相关权限
Private mlng卡类别ID As Long
Private mbln自制卡 As Boolean '当前是否自制卡
Private mbln重复使用 As Boolean '57899
Private mbln发卡 As Boolean '当前是否发卡;问题号:56599
Private mstrCurStatus As String '当前状态
Private mblnSeekName As Boolean '姓名模糊查找
Private mintNameDays As Integer '姓名查找天数

Private mlngCurPatient As Long '当前选中的病人ID '状态标记
'-------------------------------------------------------------------------
'卡相关处理
'Private mPatiCard As SquareCard '刷卡卡相关
Private mstrPassWord As String
Private mobjPatiCardObject As clsCardObject
Private mblnDefaultPassInputCardNo As Boolean
'-------------------------------------------------------------------------
Private mstrListReportName As String    '清册名称 问题50122
Private mlngListReportID As Long    '清册 问题50122
Private mstrPrivs_RollingCurtain As String  '收费轧帐管理权限
Private mobjPubPatient As Object
Private mstrPubPatiPrivs As String '公共病人信息权限
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private Type Ty_PrintProperty
    intPrintMode As Integer '打印模式 3-重打 ，4-补打，5-打印凭条
    strUseType As String '使用类别
    strInvoice As String '票据号
    strPrintNo As String '发卡单据号
    lng领用ID As Long '本次票据领用ID
    strBackInvoice As String  '回收票据
    bytPrintPayCard As Byte
    bytPrintBoundCard As Byte
End Type
Private mPrint As Ty_PrintProperty

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2011-06-28 15:22:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsCardList
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("卡号")) = "1|0"
        .ColData(.ColIndex("标志")) = "-1|1"
        If .ColIndex("ID") >= 0 Then
            .ColData(.ColIndex("ID")) = "-1|1"
            .ColHidden(.ColIndex("ID")) = True
        End If
    End With
End Sub
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2011-06-28 15:22:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:
    Set mobjSubFrm = New Collection
    
    Set mfrmChage = New frmPatiCureCardChangeMgr
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_变动记录, "变动情况", mfrmChage.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_变动记录
    mobjSubFrm.Add mfrmChage, CStr(objItem.Tag)
    
    Set mfrmConsume = New frmPatiCureCardConsumeMgr
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_帐户入出记录, "帐户入出信息", mfrmConsume.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_帐户入出记录
    mobjSubFrm.Add mfrmConsume, CStr(objItem.Tag)
    
    Set mfrmFamily = New frmPatiCureCardFamilyMgr
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_家属关系, "家属关系", mfrmFamily.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_家属关系
    mobjSubFrm.Add mfrmFamily, CStr(objItem.Tag)
    
    mblnNotClick = True
     With tbPage
        tbPage.Item(i).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    mblnNotClick = False
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '编制:刘兴洪
    '日期:2009-11-18 16:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
     With dkpMan
        .ImageList = imlPaneIcons
        Set mPanSearch = .CreatePane(mPaneID.Pane_Search, 400, 400, DockLeftOf, Nothing)
        mPanSearch.Title = "条件设置": mPanSearch.Options = PaneNoCloseable
        mPanSearch.MinTrackSize.Width = picFilter.Width / Screen.TwipsPerPixelX
        mPanSearch.MaxTrackSize.Width = picFilter.Width / Screen.TwipsPerPixelX
        
        Set objPane = .CreatePane(mPaneID.Pane_CardLists, 400, 400, DockRightOf, mPanSearch)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        Set objPane = .CreatePane(mPaneID.Pane_CardDetails, 400, 400, DockBottomOf, objPane)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Function
Private Function zlIsHaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否有数据
     '返回:当前控件有数据,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-28 18:17:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String
    zlIsHaveData = False
    If Me.ActiveControl Is vsCardList Then
        zlIsHaveData = vsCardList.TextMatrix(1, vsCardList.ColIndex("卡号")) <> ""
    End If
End Function

Private Function zlIsCardBinding() As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:是否自制卡重复使用进行绑定的特殊情况
'返回:
'编制:
'日期:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String
    zlIsCardBinding = False
    If Me.ActiveControl Is vsCardList Then
        With vsCardList
            If .TextMatrix(.Row, .ColIndex("卡号")) <> "" Then
                zlIsCardBinding = Val(.TextMatrix(.Row, .ColIndex("变动类别"))) = 11 And Val(.TextMatrix(.Row, .ColIndex("是否重复使用"))) = 1 And _
                             mbln自制卡
            End If
        End With
    End If
End Function
 

Private Sub cmdFilter_Click()
    Call InitFilterToVar
    Call LoadDataToGrid
End Sub
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
'    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case mPaneID.Pane_Search    '搜索条件窗体
        Item.Handle = picFilter.hWnd
    Case mPaneID.Pane_CardDetails   '详细卡信息
        Item.Handle = picList.hWnd
    Case mPaneID.Pane_CardLists '卡列表
        Item.Handle = picCardList.hWnd
    End Select
End Sub
Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开指定报表
    '入参:lngSys-系统号
    '     strReportCode报表编号
    '编制:刘兴洪
    '日期:2011-06-28 18:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String
    Dim str发卡人 As String, str发卡日期 As String
    With vsCardList
        If .Row < 0 Then Exit Sub
        str卡号 = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
        If str卡号 = "" Then Exit Sub
        
        str发卡人 = Trim(.TextMatrix(.Row, .ColIndex("发卡人")))
        str发卡日期 = Trim(.TextMatrix(.Row, .ColIndex("发卡日期")))
    End With
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, "卡类别ID=" & mlngCardTypeID, "卡号=" & str卡号, "发卡人=" & str发卡人, "发卡日期=" & str发卡日期)
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
    Err = 0: On Error GoTo Errhand:
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
    mcbrMenuBar.id = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "收费轧帐(&M)")
        mcbrControl.IconId = 227
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSingleBill, "重打缴款单(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_BarcodePrint, "重打发卡凭条(&W)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.id = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "发卡(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBound, "绑定卡(&B)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CancelCardBound, "取消绑定(&C)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "退卡(&T)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Cardtrade, "换卡(&H)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardFill, "补卡(&F)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardLoss, "挂失(&G)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelLoss, "取消挂失(&O)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "缴预存(&J)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "冲预存(&Y)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBackMoney, "余额退款(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_MzToZy, "门诊转住院(&M)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ZyToMz, "住院转门诊(&Z)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ModiyPati, "调整病人信息(&X)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChangePassWord, "密码调整(&P)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Family, "家属登记(&D)"): mcbrControl.BeginGroup = True
        
        '95809:李南春,2016/8/26,退病人历史病历费
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Medical, "退病历费(&N)"): mcbrControl.BeginGroup = True
        '104726:李南春,2017/4/17,收费发票重打补打
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Wham, "重打票据(&W)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Make, "补打票据(&M)")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.id = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Family, "家属信息(&V)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.id = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_CardPay
        .Add FCONTROL, Asc("T"), conMenu_Edit_CardBack
        
        .Add FCONTROL, Asc("J"), conMenu_Edit_CardInFull
        .Add FCONTROL, Asc("B"), conMenu_Edit_CardBackMoney
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F11, conMenu_Edit_RollingCurtain
    End With
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "发卡"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBound, "绑定卡"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CancelCardBound, "取消绑定")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "退卡")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Cardtrade, "换卡"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardFill, "补卡")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "缴预存"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "收费轧帐(&M)")
        mcbrControl.IconId = 227
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
        Set objComBar = .Add(xtpControlComboBox, conMenu_COMBOX_INTERFACE, "医疗卡类别")
        objComBar.Flags = xtpFlagRightAlign
        objComBar.HideFlags = xtpNoHide
        objComBar.Width = (TextWidth("刘") * 16) / Screen.TwipsPerPixelX
         objComBar.Style = xtpComboLabel
    End With
    For Each mcbrControl In mcbrToolBar.Controls
          If mcbrControl.id <> conMenu_COMBOX_INTERFACE Then
            mcbrControl.Style = xtpButtonIconAndCaption
          End If
    Next
    '加载数据
     Call LoadTypeData(objComBar)
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub LoadTypeData(ByVal cbrCmb As CommandBarComboBox)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载下拉列表数据
    '编制:刘兴洪
    '日期:2011-06-29 16:51:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, intIndex As Integer, strSQL As String
    
    On Error GoTo errHandle
    
    intIndex = 1
    '问题号:56599
    strSQL = "Select ID,编码,名称,是否自制,是否发卡,Nvl(是否重复使用,0) as 是否重复使用 From 医疗卡类别 where 是否启用=1 And Nvl(是否证件,0)=0 Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set mcolCard = New Collection
    With rsTemp
        cbrCmb.Clear
        Do While Not .EOF
            cbrCmb.AddItem CStr(Nvl(!编码)) & "-" & CStr(Nvl(!名称))
            cbrCmb.ItemData(intIndex) = Val(Nvl(!id))
            mcolCard.Add Array(Val(Nvl(rsTemp!id)), Val(rsTemp!是否自制) & "-" & Val(rsTemp!是否发卡), Val(Nvl(rsTemp!是否重复使用))), "K" & rsTemp!id
            If mlngCardTypeID = Val(Nvl(!id)) Then
               cbrCmb.ListIndex = intIndex
            End If
            intIndex = intIndex + 1
            .MoveNext
        Loop
    End With
    If intIndex > 1 And cbrCmb.ListIndex <= 0 Then
        cbrCmb.ListIndex = 1:
    End If
    If cbrCmb.ListIndex > 0 Then
        mlngCardTypeID = cbrCmb.ItemData(cbrCmb.ListIndex)
        mbln发卡 = Split(mcolCard(cbrCmb.ListIndex)(1), "-")(1) = 1 '问题号:56599
        mbln自制卡 = Split(mcolCard(cbrCmb.ListIndex)(1), "-")(0) = 1 '问题号:56599
        mbln重复使用 = mcolCard(cbrCmb.ListIndex)(2) = 1 '57899
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngID As Long, strTemp As String, strCardNo As String
    Dim ctrCombox As CommandBarComboBox, lng病人ID As Long
    Dim str操作类型 As String '取消绑定,退卡,返回;问题号:56599
    Dim objfrmPrint As frmPrint
    Dim strSelect As String
    '---------------------------------------------
    Set objfrmPrint = New frmPrint
    Load objfrmPrint
    Select Case Control.id
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_PrintSingleBill       '"重打缴款单(&R)")
        strSelect = zlCommFun.ShowMsgbox("缴款单打印", "请选择你要打印的缴款单", "发卡(&F),预存(&I),取消(&C)", Me, _
                                         vbDefaultButton2)
        If Not (strSelect = "取消" Or strSelect = "") Then
            Call objfrmPrint.PrintReBill(strSelect, Trim(vsCardList.TextMatrix(vsCardList.Row, vsCardList.ColIndex("卡号"))), _
                                         mlngCardTypeID, mPrint.bytPrintPayCard)
        End If
    Case conMenu_Edit_RollingCurtain   '收费员轧帐
        Call zlExecuteChargeRollingCurtain(Me)
    Case conMenu_Edit_CardPay    '发卡(&S)")
            If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_发卡, mlngCardTypeID) = False Then Exit Sub
            If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Call LoadDataToGrid
      Case conMenu_Edit_CardBound
            If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_绑定卡, mlngCardTypeID) = False Then Exit Sub
            If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Call LoadDataToGrid
      Case conMenu_Edit_CardBack    '退卡(&B)")
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
           ' If strCardNo = "" Then Exit Sub
        End With
        '问题号:56599
        str操作类型 = Check退卡(strCardNo)
        Select Case str操作类型
            Case "取消绑定"
                zlExecuteCommandBars cbsThis.FindControl("", conMenu_Edit_CancelCardBound)
            Case "退卡"
                If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_退卡, mlngCardTypeID, strCardNo) = False Then Exit Sub
                If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Call LoadDataToGrid
            Case "返回"
            
        End Select
    Case conMenu_Edit_Cardtrade   '换卡
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
            '90233:李南春,2015/11/5,换卡和补卡传入病人ID
            lng病人ID = Val(Trim(.TextMatrix(.Row, .ColIndex("病人ID"))))
            If strCardNo = "" Then Exit Sub
        End With
        If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_换卡, mlngCardTypeID, strCardNo, lng病人ID) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToGrid
    Case conMenu_Edit_CardFill    '补卡(&B)")
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
            lng病人ID = Val(Trim(.TextMatrix(.Row, .ColIndex("病人ID"))))
          '  If strCardNo = "" Then Exit Sub
        End With
        If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_补卡, mlngCardTypeID, strCardNo, lng病人ID) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToGrid
    Case conMenu_Edit_CardLoss        '挂失
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
            lng病人ID = Val(Trim(.TextMatrix(.Row, .ColIndex("病人ID"))))
        
            If strCardNo = "" Then Exit Sub
        End With
        If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_挂失, mlngCardTypeID, strCardNo, lng病人ID) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToGrid
    Case conMenu_Edit_CardCancelLoss        '取消挂失
        Call SaveCardCancelLose
   Case conMenu_Edit_CancelCardBound
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then Exit Sub
        End With
        '问题号:56599
        str操作类型 = Check取消院外卡绑定(lng病人ID, strCardNo)
        Select Case str操作类型
            Case "取消绑定"
                frmPaticurCardCancelBound.mstrPrepayPrivs = mstrPrepayPrivs
                If frmPaticurCardCancelBound.zlCancelBand(Me, mlngModule, mlngCardTypeID, lng病人ID, strCardNo, False) = False Then Exit Sub
                Call LoadDataToGrid
            Case "退卡"
                zlExecuteCommandBars cbsThis.FindControl("", conMenu_Edit_CardBack)
                Exit Sub
            Case "返回"
                
        End Select
    Case conMenu_Edit_CardInFull    '缴预存(&J)"
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
        End With
        'intFunc-1-缴预存;2-退预款;3-作废,4-门诊转住院;5-住院转门诊;
        Call zlPrepayFunc(1, lng病人ID)
    
    Case conMenu_Edit_CardBackMoney '退款
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
        End With
        'intFunc-1-缴预存;2-退预款;3-作废,4-门诊转住院;5-住院转门诊;
     Call zlPrepayFunc(2, lng病人ID)
    Case conMenu_Edit_CardInFullBack    '冲预存
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
        End With        'intFunc-1-缴预存;2-退预款;3-作废,4-门诊转住院;5-住院转门诊;
        Call zlPrepayFunc(3, lng病人ID)
    Case conMenu_Edit_MzToZy
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
        End With        'intFunc-1-缴预存;2-退预款;3-作废,4-门诊转住院;5-住院转门诊;
        Call zlPrepayFunc(4, lng病人ID)
    Case conMenu_Edit_ZyToMz
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
        End With        'intFunc-1-缴预存;2-退预款;3-作废,4-门诊转住院;5-住院转门诊;
        Call zlPrepayFunc(5, lng病人ID)
    Case conMenu_Edit_ModiyPati  '调整病人信息(&M)
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        End With
        If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_调整病人信息, mlngCardTypeID, "", lng病人ID) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToGrid
    Case conMenu_Edit_ChangePassWord    '密码调整
         If frmModiPatiPass.zlModifyPass(Me, mlngModule, mlngCardTypeID, , , InStr(1, mstrPrivs, ";强制修改密码;") = 0) Then
              Exit Sub
         End If
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToGrid
    Case conMenu_Edit_Family
        '功能:病人家属设置
        If Not CreatePublicPatient Then Exit Sub
        Call mobjPubPatient.MakePatiFamily(Me, 0, 2, mlngModule) '编辑
        zlRefrshListData
    Case conMenu_Edit_Wham
        mPrint.intPrintMode = 3
        Call PrintBill
    Case conMenu_Edit_Make
        mPrint.intPrintMode = 4
        Call PrintBill
    Case conMenu_File_BarcodePrint
        mPrint.intPrintMode = 0
        Call PrintBill
    Case conMenu_View_Family
        If Not CreatePublicPatient Then Exit Sub
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        End With
        If lng病人ID <= 0 Then Exit Sub
        Call mobjPubPatient.MakePatiFamily(Me, lng病人ID, 1, mlngModule) '查看
    Case conMenu_Edit_Medical
        '功能：退病历费 如果有病人ID和卡号则自动定位到发卡时的病历费记录
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
            lng病人ID = Val(Trim(.TextMatrix(.Row, .ColIndex("病人ID"))))
        End With
        Call frmPatiBooks.ShowMe(Me, 1, mlngModule, lng病人ID, strCardNo)
    Case conMenu_COMBOX_INTERFACE   '点击选择
        Set ctrCombox = Control
        mlngCardTypeID = ctrCombox.ItemData(ctrCombox.ListIndex)
        mbln自制卡 = Split(mcolCard(ctrCombox.ListIndex)(1), "-")(0) = 1
        mbln发卡 = Split(mcolCard(ctrCombox.ListIndex)(1), "-")(1) = 1 '问题号:56599
        '115505:李南春,2017/10/23,更新卡属性
        mbln重复使用 = mcolCard(ctrCombox.ListIndex)(2) = 1
        Call LoadDataToGrid
    Case conMenu_View_Refresh   '刷新
        '重新刷新数据
        Call LoadDataToGrid
    Case mlngListReportID    '问题50122
        If vsCardList.TextMatrix(vsCardList.Row, vsCardList.ColIndex("病人ID")) = "" Then
            '问题号:57285
            ShowMsgbox "您还没有选择对应的收支明细,请到帐户入出信息选择对应的收支明细!"
            Exit Sub
        End If
        mfrmConsume.zlShowReport CLng(vsCardList.TextMatrix(vsCardList.Row, vsCardList.ColIndex("病人ID")))
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
 Private Function zlPopuReportMenus() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:弹出收支清册菜单
    '编制:刘兴洪
    '日期:2012-06-12 15:28:03
    '问题:50122
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    If mlngListReportID = 0 Then Exit Function
    Set cbrPopupBar = Me.cbsThis.Add("弹出报表菜单", xtpBarPopup)
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mlngListReportID, mstrListReportName)
    cbrPopupBar.ShowPopup
 End Function
 
 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNo As String, strExpand
    Dim strOutPatiInforXml As String

    If IsCardType(IDKind, "IC卡号") And Not txtPatient.Locked Then
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call txtPatient_KeyPress(vbKeyReturn)
            End If
        End If
        Exit Sub
    End If

    lng卡类别ID = IDKind.GetCurCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
   
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNo, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNo
    If txtPatient.Text <> "" Then
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
    Set gobjSquare.objCurCard = objCard
    '105155:李南春,2017/2/8,卡号密文显示判断不正确
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    If objCard.接口序号 > 0 Then
        txtPatient.MaxLength = objCard.卡号长度
    Else
        txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    End If
    
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = "": mlngCurPatient = 0
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
    If mlngCurPatient <> 0 Then
        txtPatient.PasswordChar = ""  '如果已经通过各种方式获取到了病人,此时显示的是病人姓名,不应该设置掩码
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
    
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean
    '60010
    If txtPatient.Locked Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
    mblnNotClick = True

    intIndex = IDKind.GetKindIndex(objCard.名称)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex

    txtPatient.Text = objPatiInfor.卡号
    Call txtPatient_KeyPress(vbKeyReturn)
'    If mrsInfo Is Nothing Then
'        blnNew = True
'    ElseIf mrsInfo.State <> 1 Then
'        blnNew = True
'    End If
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

 '问题50122
Private Sub mfrmConsume_zlPopupMenus(ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    If vsGrid.Rows = 1 Or vsGrid.Row = 1 Then Exit Sub
    zlPopuReportMenus
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean, bln挂失 As Boolean
    Dim lng病人ID As Long
    Dim blnIsBind As Boolean '自制卡是否通过绑定的方式进行重复使用的
    If Me.Visible = False Then Exit Sub
    
    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = zlIsHaveData
    Case conMenu_Edit_RollingCurtain        '收费员轧帐
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs_RollingCurtain, "轧帐")
        Control.Enabled = Control.Visible
        
    Case conMenu_File_PrintSingleBill           '"重打缴款单(&R)"
        blnIsBind = True
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
            blnIsBind = Trim(.TextMatrix(.Row, .ColIndex("单据号"))) <> "" And .TextMatrix(.Row, .ColIndex("状态")) = "有效卡" And lng病人ID > 0
        End With
        Control.Visible = (zlstr.IsHavePrivs(mstrPrivs, "预交收据") Or zlstr.IsHavePrivs(mstrPrivs, "医疗卡收据")) And blnIsBind And Not gbln收费发票
        Control.Enabled = Control.Visible
    Case conMenu_File_BarcodePrint
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
        End With
        Control.Visible = (zlstr.IsHavePrivs(mstrPrivs, "重打发卡凭条")) And lng病人ID > 0
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardPay
        Control.Visible = (zlstr.IsHavePrivs(mstrPrivs, "发卡") And mbln自制卡) Or (zlstr.IsHavePrivs(mstrPrivs, "发卡") And mbln自制卡 = False And mbln发卡 = True) '问题号:56599
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardBound
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "绑定卡") And ((Not mbln自制卡) Or (mbln自制卡 And mbln重复使用))
        Control.Enabled = Control.Visible
        
    Case conMenu_Edit_CancelCardBound
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
        End With
        blnIsBind = True
        If mbln自制卡 Then
            '自制卡 使用绑定进行重复使用
            blnIsBind = zlIsCardBinding
        End If
        '不是自制卡采用绑定,自制卡 重复利用使用绑定的方式
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "取消绑定卡") And ((Not mbln自制卡) Or (mbln自制卡 And blnIsBind))
        Control.Enabled = Control.Visible And lng病人ID > 0
    Case conMenu_Edit_CardBack
        blnIsBind = False
        If mbln自制卡 Then
            '自制卡 使用绑定进行重复使用
            blnIsBind = zlIsCardBinding
        End If
        '园内卡退卡,绑定则不能退卡,只能取消
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "退卡") And mbln自制卡 And Not blnIsBind
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Cardtrade '换卡
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "换卡") And mbln自制卡
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardFill  '补卡
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "补卡") And mbln自制卡
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardLoss  '挂失
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "挂失") And mstrCurStatus = "有效卡"
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardCancelLoss  '取消挂失
        bln挂失 = mstrCurStatus = "已挂失"
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "取消挂失") ' And mbln自制卡
        Control.Enabled = Control.Visible And zlIsHaveData And bln挂失
    Case conMenu_Edit_CardInFull  '缴预存
        Control.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "门诊预交") Or zlstr.IsHavePrivs(mstrPrepayPrivs, "住院预交") Or zlstr.IsHavePrivs(mstrPrepayPrivs, "共用预交")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardInFullBack  '冲预存
        Control.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "预交退款")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardBackMoney  '退款
        Control.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "预交退款")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_MzToZy    '门诊转住院
        Control.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "门诊预交转住院")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_ZyToMz    '住院转门诊
        Control.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "住院预交转门诊")
        Control.Enabled = Control.Visible
   Case conMenu_Edit_ModiyPati
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "调整病人信息")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_ChangePassWord    '密码调整
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "修改密码")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Family
        Control.Visible = zlstr.IsHavePrivs(mstrPubPatiPrivs, "病人家属")
        Control.Enabled = Control.Visible
    Case conMenu_View_Family
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
        End With
        Control.Visible = zlstr.IsHavePrivs(mstrPubPatiPrivs, "病人家属") And lng病人ID > 0
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Wham
        blnIsBind = True
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
            blnIsBind = Trim(.TextMatrix(.Row, .ColIndex("单据号"))) <> "" And .TextMatrix(.Row, .ColIndex("状态")) = "有效卡"
        End With
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "重打发票") And gbln收费发票 And blnIsBind And lng病人ID > 0
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Make
        blnIsBind = True
        With vsCardList
            lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            If lng病人ID <= 0 Then lng病人ID = 0
            blnIsBind = Trim(.TextMatrix(.Row, .ColIndex("单据号"))) <> "" And .TextMatrix(.Row, .ColIndex("状态")) = "有效卡"
        End With
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "补打发票") And gbln收费发票 And blnIsBind And lng病人ID > 0
        Control.Enabled = Control.Visible
    Case conMenu_View_Refresh   '刷新
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_1" And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1107_2"
        End If
    End Select
End Sub
 
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.id
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
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
        Case conMenu_File_Parameter     '参数调用
             If frmPatiCureCardPara.zlSetPara(Me, mlngModule, mstrPrivs) = False Then Exit Sub
        Case Else   '其他操作功能调用
            Call zlExecuteCommandBars(Control)
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If tbPage.Selected Is Nothing Then Exit Sub
    If Me.Visible = False Then Exit Sub

    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        Call zlUpdateCommandBars(Control)
    End Select
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    zlControl.ControlSetFocus vsCardList
    Call vsCardList_GotFocus
    mblnFirst = False
End Sub

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim strShow As String
    Dim i As Long
    If mblnInited = False Then
        mblnInited = True
    Else
        Exit Sub
    End If
    mblnFirst = True: mstrCurStatus = ""
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    mstrPrepayPrivs = ";" & GetPrivFunc(glngSys, 1103) & ";"
    mstrPubPatiPrivs = ";" & GetPrivFunc(glngSys, 9003) & ";"
    Call InitFace
    mlngCardTypeID = Val(zlDatabase.GetPara("上次医疗类别", glngSys, mlngModule, 0, , InStr(1, mstrPrivs, ";参数设置;") > 0))
    '只有发卡才会有重打补打
    mPrint.bytPrintPayCard = Split(zlDatabase.GetPara("医疗卡收据格式", glngSys, mlngModule, "0|0"), "|")(0)
    mPrint.bytPrintBoundCard = Split(zlDatabase.GetPara("医疗卡收据格式", glngSys, mlngModule, "0|0"), "|")(1)
    RestoreWinState Me, App.ProductName, mstrTitle
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call InitData: Call InitPanel: Call InitPage
    Call zlDefCommandBars '初始菜单及工具栏
    Call InitFilterToVar
    Call LoadDataToGrid(-1)
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    Call InitListReport
End Sub
Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2011-06-21 13:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKindStr As String, blnVisible As Boolean
    Dim intKind As Integer, strKey As String
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    
    Call InitIDKind
     
    
    '取缺省的刷卡方式
    '短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
    '第7位后,就只能用索引,不然取不到数
    
    '89086:李南春,2015/10/9,管理界面姓名模糊查找
    mblnSeekName = zlDatabase.GetPara("姓名模糊查找", glngSys, mlngModule) = "1"
    mintNameDays = Val(zlDatabase.GetPara("姓名查找天数", glngSys, mlngModule))
    
    gobjSquare.bln缺省卡号密文 = IDKind.ShowPassText
    
    Call GetRegInFor(g私有模块, Me.Name, "idkind", strKey)
    intKind = Val(strKey)
    If intKind > 0 And intKind <= IDKind.ListCount Then
        IDKind.IDKind = intKind
    End If
    
 End Sub
Private Function InitIDKind() As Boolean
    Dim lngCardID As Long
    If gobjSquare Is Nothing Then Exit Function
    gobjSquare.objSquareCard.mblnYLMgr = True
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModule, 0))
    On Error GoTo ErrEnd
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, IDKind.IDKindStr, txtPatient)
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    IDKind.ShowPropertySet = InStr(";" & mstrPrivs & ";", "参数设置") > 0
ErrEnd:
End Function
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2011-06-29 18:08:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dtp结束日期(0).MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    dtp结束日期(0).value = Format(dtp结束日期(0).MaxDate, "yyyy-mm-dd 23:59:59")
    dtp开始日期(0).MaxDate = dtp结束日期(0).MaxDate
    dtp开始日期(0).value = Format(DateAdd("d", -7, dtp开始日期(0).MaxDate), "yyyy-mm-dd 00:00:00")
    dtp结束日期(1).MaxDate = dtp结束日期(0).MaxDate
    dtp结束日期(1).value = Format(dtp结束日期(1).MaxDate, "yyyy-mm-dd 23:59:59")
    dtp开始日期(1).MaxDate = dtp结束日期(1).MaxDate
    dtp开始日期(1).value = dtp开始日期(0).value
 
     
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Long, strTemp As String
   If Me.Visible = False Then Exit Sub
   SaveWinState Me, App.ProductName, mstrTitle
   zlDatabase.SetPara "上次医疗类别", mlngCardTypeID, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
   zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlstr.IsHavePrivs(mstrPrivs, "参数设置")
   
   
   zlSaveDockPanceToReg Me, dkpMan, "区域"
   If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled False
        Set mobjIDCard = Nothing
   End If
   If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled False
        Set mobjICCard = Nothing
   End If
   If Not mobjReport Is Nothing Then Set mobjReport = Nothing
   If Not mobjPubPatient Is Nothing Then Set mobjPubPatient = Nothing
    '关闭子窗口
    If Not mobjSubFrm Is Nothing Then
        For i = 1 To mobjSubFrm.count
            If Not mobjSubFrm(i) Is Nothing Then Unload mobjSubFrm(i)
        Next
    End If
    
    Call SaveRegInFor(g私有模块, Me.Name, "idkind", IDKind.IDKind)
End Sub
 Private Function zlPopuMenus(ByVal blnListView As Boolean) As Boolean
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Err = 0: On Error Resume Next
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Function
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
        cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
    Next

    If Me.cbsThis.ActiveMenuBar.Controls(3).Visible Then
        Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(3)
        For Each mcbrControl In mcbrMenuBar.CommandBar.Controls

            Select Case mcbrControl.id
            Case conMenu_View_ShowStoped, conMenu_View_ShowAll, conMenu_View_Refresh
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
                cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
                cbrPopupItem.Checked = mcbrControl.Checked
            End Select
        Next
    End If
    cbrPopupBar.ShowPopup
End Function

Private Function zlCheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据依赖性
    '返回:数据合法,返回true，否则返回False
    '编制:刘兴洪
    '日期:2009-11-19 15:37:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, strSQL As String
    zlCheckDepend = False
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 名称   From 结算方式 Where 性质 = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查现金结算方式", UserInfo.id)
    If rsTemp.EOF Then
        ShowMsgbox "结算方式中不存在一条件有现金性质的结算方式,请在结算方式管理中设置!"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    '76009,冉俊明,2014-7-30
    strSQL = "Select 1 From 医疗卡类别 Where Nvl(是否启用, 0) = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查医疗卡类别")
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "医疗卡类别中不存在任何可用类别，请在“医疗卡类别管理”中进行维护！"
        Exit Function
    End If
    zlCheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ShowList(ByVal lngModule As Long, ByVal strTitle As String, ByVal frmMain As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口,显示相关的项目及分类信息
    '编制:刘兴洪
    '日期:2009-11-19 15:38:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrTitle = strTitle: mstrPrivs = gstrPrivs
    If Not zlCheckDepend Then Exit Sub            '数据依赖性测试
    Me.Caption = strTitle
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        zlCommFun.ShowChildWindow Me.hWnd, frmMain
    End If
    Me.ZOrder 0
End Sub

 
Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long, vRect As RECT
    vRect = zlControl.GetControlRect(picImgList.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCardList, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlstr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub InitFilterToVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件给变量
    '编制:刘兴洪
    '日期:2011-06-28 23:56:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mArrFilter = New Collection
    mArrFilter.Add Array(Trim(txtCard(0).Text), Trim(txtCard(1).Text)), "卡号范围"
    If chkFilter(0).value Then
        mArrFilter.Add Array(Format(dtp开始日期(0).value, "yyyy-mm-dd HH:MM:SS"), Format(dtp结束日期(0).value, "yyyy-mm-dd HH:MM:SS")), "发卡时间"
    Else
        mArrFilter.Add Array("1901-01-01", "1901-01-01"), "发卡时间"
    End If
    If chkFilter(1).value Then
        mArrFilter.Add Array(Format(dtp开始日期(1).value, "yyyy-mm-dd HH:MM:SS"), Format(dtp结束日期(1).value, "yyyy-mm-dd HH:MM:SS")), "挂失时间"
    Else
        mArrFilter.Add Array("1901-01-01", "1901-01-01"), "挂失时间"
    End If
    mArrFilter.Add Trim(txtEdit(mtxtIdx.idx_发卡人)), "发卡人"
    mArrFilter.Add Trim(txtEdit(mtxtIdx.idx_挂失人)), "挂失人"
    mArrFilter.Add Trim(txtName.Text), "姓名"
End Sub
Private Function LoadDataToGrid(Optional lng病人ID As Long = 0, Optional strCardNo As String = "") As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:加载数据给网格
'入参:lng病人ID-找指定的病人ID;
'       strCardNo-找指定的卡号
'返回:加载成功,返回true,否则返回False
'编制:刘兴洪
'日期:2009-11-19 15:43:29
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, lngRow As Long, strPreCardNO As String, lngPreTypeID As Long
    Dim rsTemp As ADODB.Recordset, strSQL As String, i As Long
    Err = 0: On Error GoTo Errhand:
    strWhere = ""
    If lng病人ID <> 0 Then
        strWhere = strWhere & " And A.病人ID=[10] "
    ElseIf strCardNo <> "" Then
        strWhere = strWhere & " And A.卡号=[11] "
    Else
        If mArrFilter("发卡时间")(0) <> "1901-01-01" And mArrFilter("挂失时间")(0) <> "1901-01-01" Then
            strWhere = strWhere & " And (A.发卡日期 Between [1] And [2] Or A.挂失时间 Between [3] And [4])"
        ElseIf mArrFilter("发卡时间")(0) = "1901-01-01" And mArrFilter("挂失时间")(0) <> "1901-01-01" Then
            strWhere = strWhere & " And (A.挂失时间 Between [3] And [4])"
        ElseIf mArrFilter("发卡时间")(0) <> "1901-01-01" And mArrFilter("挂失时间")(0) = "1901-01-01" Then
            strWhere = strWhere & " And (A.发卡日期 Between [1] And [2])"
        End If
        If mArrFilter("卡号范围")(0) <> "" And mArrFilter("卡号范围")(1) <> "" Then
            strWhere = strWhere & " And (A.卡号 Between [5] And [6])"
        ElseIf mArrFilter("卡号范围")(0) = "" And mArrFilter("卡号范围")(1) <> "" Then
            strWhere = strWhere & " And A.卡号=[6]"
        ElseIf mArrFilter("卡号范围")(0) <> "" And mArrFilter("卡号范围")(1) = "" Then
            strWhere = strWhere & " And A.卡号=[5]"
        End If
        If mArrFilter("挂失人") <> "" Then strWhere = strWhere & " and  A.挂失人 like [7]"
        If mArrFilter("发卡人") <> "" Then strWhere = strWhere & " and  A.发卡人 like [8]"
        If mArrFilter("姓名") <> "" Then
            If zlstr.ActualLen(mArrFilter("姓名")) < 4 And _
                   (DateDiff("d", CDate(mArrFilter("发卡时间")(0)), CDate(mArrFilter("发卡时间")(1))) > 7 Or _
                   DateDiff("d", CDate(mArrFilter("挂失时间")(0)), CDate(mArrFilter("挂失时间")(1))) > 7) Then
                    MsgBox "输入的信息太少，病人姓名必须至少输入两个汉字或四个字符!", vbInformation, gstrSysName
                    Exit Function
            Else
                strWhere = strWhere & " and  B.姓名 like [12]"
            End If
        End If
        If mArrFilter("挂失人") = "" And mArrFilter("发卡人") = "" And mArrFilter("姓名") = "" And _
           mArrFilter("卡号范围")(0) = "" And mArrFilter("卡号范围")(1) = "" Then
            If DateDiff("d", CDate(mArrFilter("发卡时间")(0)), CDate(mArrFilter("发卡时间")(1))) > 30 Or _
               DateDiff("d", CDate(mArrFilter("挂失时间")(0)), CDate(mArrFilter("挂失时间")(1))) > 30 Then
                If MsgBox("选择的时间范围超过了30天,是否继续?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
            End If
        End If
    End If
    Call zlCommFun.ShowFlash("正在加载病人医疗卡信息,请稍等...", Me)
    '    strSQL = "" & _
         '         "    Select  A.病人ID,A.卡类别ID,C.编码||'-'|| C.名称 As 医疗卡类别,A.卡号,  " & _
         '         "             case when Nvl(A.状态,0)=0 then '有效卡' " & _
         '         "                     when Nvl(A.状态,0)=2 then '补卡停用' " & _
         '         "                     When Nvl(挂失时间,to_date('3000-01-01','yyyy-mm-dd'))+ Nvl(D.有效天数,0)<=sysdate then '已挂失' " & _
         '         "                     Else ''  end as 状态, " & _
         '         "              A.挂失人,A.挂失方式, to_char(A.挂失时间,'yyyy-mm-dd hh24:mi:ss') as 挂失时间," & _
         '         "             A.发卡人,to_char(A.发卡日期,'yyyy-mm-dd hh24:mi:ss')  as 发卡日期, " & _
         '         "             B.姓名,B.性别,B.年龄,to_char(B.出生日期,'yyyy-mm-dd hh24:mi:ss')  as  出生日期,B.出生地点," & _
         '         "             B.身份证号,B.门诊号,B.住院号,b.国籍,b.家庭地址,B.家庭电话,b.监护人,b.联系人姓名, " & _
         '         "             b.联系人关系,b.联系人地址,b.联系人电话,b.工作单位,b.单位电话,b.家庭地址邮编,decode(b.在院,1,'√','') as 在院 " & _
         '         "     From 病人医疗卡信息 A,病人信息 B,医疗卡类别 C, 医疗卡挂失方式 D " & _
         '         "     Where A.病人ID=B.病人ID And A.卡类别ID=C.Id And A.挂失方式=D.名称(+)  and A.卡类别ID=[9] " & strWhere

    strSQL = " Select * " & vbNewLine & _
           "   From (With 医疗卡变动 As (" & vbNewLine & _
           "                      Select Max(Bd.ID) As 变动id, C.是否严格控制, C.是否重复使用, A.病人id, A.卡类别id," & vbNewLine & _
           "                              C.编码 || '-' || C.名称 As 医疗卡类别, A.卡号," & vbNewLine & _
           "                              Case" & vbNewLine & _
           "                                When Nvl(A.状态, 0) = 0 Then" & vbNewLine & _
           "                                 '有效卡'" & vbNewLine & _
           "                                When Nvl(A.状态, 0) = 2 Then" & vbNewLine & _
           "                                 '补卡停用'" & vbNewLine & _
    "                                When Nvl(挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(D.有效天数, 0) <= Sysdate Then"
    strSQL = strSQL & vbNewLine & _
           "                                 '已挂失'" & vbNewLine & _
           "                                Else" & vbNewLine & _
           "                                 ''" & vbNewLine & _
           "                              End As 状态, A.挂失人, A.挂失方式," & vbNewLine & _
           "                              To_Char(A.挂失时间, 'yyyy-mm-dd hh24:mi:ss') As 挂失时间, A.发卡人," & vbNewLine & _
           "                              To_Char(A.发卡日期, 'yyyy-mm-dd hh24:mi:ss') As 发卡日期, B.姓名, B.性别, B.年龄," & vbNewLine & _
           "                              To_Char(B.出生日期, 'yyyy-mm-dd hh24:mi:ss') As 出生日期, B.出生地点, B.身份证号, B.门诊号," & vbNewLine & _
           "                              B.住院号,B.手机号, B.国籍, B.家庭地址 as 现住址, B.家庭电话, B.监护人, B.联系人姓名, B.联系人关系," & vbNewLine & _
           "                              B.联系人地址, B.联系人电话, B.工作单位, B.单位电话, B.家庭地址邮编," & vbNewLine & _
           "                              Decode(B.在院, 1, '√', '') As 在院" & vbNewLine & _
           "                      From 病人医疗卡信息 A, 病人信息 B, 医疗卡类别 C, 医疗卡挂失方式 D, 病人医疗卡变动 Bd" & vbNewLine & _
           "                      Where A.病人id = B.病人id And A.卡类别id = C.ID And A.挂失方式 = D.名称(+) And" & vbNewLine & _
           "                            A.卡类别id=[9]   " & strWhere & " And A.病人id = Bd.病人id(+) And A.卡类别id = Bd.卡类别id(+) And" & vbNewLine & _
           "                            A.卡号 = Bd.卡号(+)" & vbNewLine & _
           "                      Group By C.是否严格控制, C.是否重复使用, A.病人id, A.卡类别id, C.编码 || '-' || C.名称, A.卡号," & vbNewLine & _
           "                                Case" & vbNewLine & _
           "                                  When Nvl(A.状态, 0) = 0 Then" & vbNewLine & _
           "                                   '有效卡'" & vbNewLine & _
           "                                  When Nvl(A.状态, 0) = 2 Then" & vbNewLine & _
           "                                   '补卡停用'" & vbNewLine & _
           "                                  When Nvl(挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(D.有效天数, 0) <=Sysdate Then" & vbNewLine & _
           "                                   '已挂失'" & vbNewLine & _
           "                                  Else " & vbNewLine & _
    "                                   '' "
    strSQL = strSQL & vbNewLine & _
           "                                End, A.挂失人, A.挂失方式, To_Char(A.挂失时间, 'yyyy-mm-dd hh24:mi:ss'), A.发卡人," & vbNewLine & _
           "                                To_Char(A.发卡日期, 'yyyy-mm-dd hh24:mi:ss'), B.姓名, B.性别, B.年龄," & vbNewLine & _
           "                                To_Char(B.出生日期, 'yyyy-mm-dd hh24:mi:ss'), B.出生地点, B.身份证号, B.门诊号, B.住院号,B.手机号," & vbNewLine & _
           "                                B.国籍, B.家庭地址, B.家庭电话, B.监护人, B.联系人姓名, B.联系人关系, B.联系人地址," & vbNewLine & _
           "                                B.联系人电话, B.工作单位, B.单位电话, B.家庭地址邮编, Decode(B.在院, 1, '√', '')" & vbNewLine & _
           "                      )" & vbNewLine & _
           "   Select T.病人id, T.卡类别id, T.医疗卡类别, T.卡号, T.状态, T.挂失人, T.挂失方式, T.挂失时间, T.发卡人, T.发卡日期," & vbNewLine & _
           "          T.姓名, T.性别, T.年龄, T.出生日期, T.出生地点, T.身份证号, T.门诊号, T.住院号,T.手机号, T.国籍, T.现住址," & vbNewLine & _
           "           T.家庭电话, T.监护人, T.联系人姓名, T.联系人关系, T.联系人地址, T.联系人电话, T.工作单位, T.单位电话," & vbNewLine & _
           "          T.家庭地址邮编, T.在院, Nvl(Bd.变动类别, 0) As 变动类别, T.是否严格控制, T.是否重复使用, Z.NO as 单据号," & vbNewLine & _
           "          LTrim(To_Char(max(Decode(Y.类型, 1, Nvl(Y.预交余额,0),0)),'99999999990.00')) As 门诊预交余额, " & _
           "          LTrim(To_Char(max(Decode(Y.类型, 1, 0, Nvl(Y.预交余额,0))),'99999999990.00')) As 住院预交余额 " & _
           "   From 医疗卡变动 T, 病人医疗卡变动 Bd , 住院费用记录 Z,病人余额 Y" & vbNewLine & _
           "   Where T.变动id = Bd.ID(+) And T.病人ID = Z.病人ID(+) And T.病人ID = Y.病人ID(+) and Y.性质(+)=1" & vbNewLine & _
           "         And Z.记录性质(+) = 5 And z.记录状态(+) = 1 And T.卡号 = Z.实际票号(+) And T.卡类别ID = Nvl(Z.结论(+),0)" & vbNewLine & _
           "   Group by T.病人id, T.卡类别id, T.医疗卡类别, T.卡号, T.状态, T.挂失人, T.挂失方式, T.挂失时间, T.发卡人, T.发卡日期," & vbNewLine & _
           "         T.姓名, T.性别, T.年龄, T.出生日期, T.出生地点, T.身份证号, T.门诊号, T.住院号,T.手机号, T.国籍, T.现住址," & vbNewLine & _
           "         T.家庭电话, T.监护人, T.联系人姓名, T.联系人关系, T.联系人地址, T.联系人电话, T.工作单位, T.单位电话," & vbNewLine & _
           "         T.家庭地址邮编, T.在院, Bd.变动类别, T.是否严格控制, T.是否重复使用, Z.NO) T"

    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                          CDate(mArrFilter("发卡时间")(0)), CDate(mArrFilter("发卡时间")(1)), _
                                          CDate(mArrFilter("挂失时间")(0)), CDate(mArrFilter("挂失时间")(1)), _
                                          CStr(mArrFilter("卡号范围")(0)), CStr(mArrFilter("卡号范围")(1)), _
                                          CStr(mArrFilter("挂失人")), CStr(mArrFilter("发卡人")), mlngCardTypeID, _
                                          lng病人ID, strCardNo, Trim(txtName.Text) & "%")
    With vsCardList
        If .Row > 0 And .ColIndex("卡号") >= 0 Then
            strPreCardNO = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
            If strPreCardNO <> "" And .ColIndex("卡类别ID") >= 0 Then
                lngPreTypeID = Val(.TextMatrix(.Row, .ColIndex("卡类别ID")))
            End If
        End If
        .Redraw = flexRDNone
        .Clear: .Rows = 2: .Cols = 1
        .Cell(flexcpForeColor, 1, .FixedCols - 1, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpText, 0, 0, .Rows - 1, .Cols - 1) = ""
        mblnNotRefresh = True
        Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        .Row = 1
        If lngPreTypeID = mlngCardTypeID Then   '行定位
            i = .FindRow(strPreCardNO, 1, .ColIndex("卡号"))
            If i >= 1 Then .Row = i
        End If
        mblnNotRefresh = False

        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColHidden(i) = True: .ColWidth(i) = True
                .ColData(i) = "-1|1"    '不能选择
            ElseIf .ColKey(i) Like "*时间" Or .ColKey(i) Like "*日期" Or .ColKey(i) = "状态" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf InStr(";变动类别;是否严格控制;是否重复使用;单据号;", ";" & .ColKey(i) & ";") > 0 Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"    '不能选择
            ElseIf .ColKey(i) Like "*余额" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .ColData(.ColIndex("卡号")) = "1|0": .ColData(.ColIndex("标志")) = "-1|1"
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsCardList, Me.Name, "卡信息列表", True, True
        .ColWidth(.ColIndex("标志")) = 285
        .ColAlignment(.ColIndex("标志")) = flexAlignCenterCenter
        Call SetGridRowForeColor     '设置行颜色
        .Redraw = flexRDBuffered
    End With
    Call vsCardList_AfterRowColChange(-1, 0, vsCardList.Row, 0)
    Call zlCommFun.StopFlash
    LoadDataToGrid = True
    Exit Function
Errhand:
    Call zlCommFun.StopFlash
    vsCardList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SetGridRowForeColor(Optional ByVal lngRow As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置行颜色
    '入参:lngRow=0,表示重新设置所有行的颜色
    '编制:刘兴洪
    '日期:2011-06-29 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long, int状态 As Integer, lngRows As Long, i As Long
    With vsCardList
        If lngRow = 0 Then lngRows = .Rows - 1: lngRow = 1
        For i = lngRow To lngRows
            lngColor = IIf(Trim(.TextMatrix(i, .ColIndex("挂失时间"))) <> "", vbRed, &H80000008)
            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = lngColor
        Next
    End With
End Sub

Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant)
    Dim lng领用ID As Long, datDate As Date
    Dim strSQL As String
    
    On Error GoTo errH
    If gblnBill发卡 Then
        lng领用ID = GetInvoiceGroupID(1, TotalPages, mPrint.lng领用ID, glngShareUseID, mPrint.strInvoice, mPrint.strUseType)
        If lng领用ID <= 0 Then
            Select Case lng领用ID
                Case -1
                    MsgBox "单据[" & mPrint.strPrintNo & "]" & "共需要" & TotalPages & "张票据！" & vbCrLf & _
                        "你没有足够的自用和共用的票据，请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "单据[" & mPrint.strPrintNo & "]" & "共需要" & TotalPages & "张票据！" & vbCrLf & _
                        "你没有足够的的共用票据，请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -3
                    MsgBox "单据[" & mPrint.strPrintNo & "]" & "共需要" & TotalPages & "张票据！" & vbCrLf & _
                        "票据号[" & mPrint.strInvoice & "]不在可用领用批次的有效票据号范围内！" & _
                        "请重新输入有效的票据号后重打该单据！", vbInformation, gstrSysName
                Case -4
                    MsgBox "单据[" & mPrint.strPrintNo & "]" & "共需要" & TotalPages & "张票据！" & vbCrLf & _
                        "票据号[" & mPrint.strInvoice & "]所在的领用批次没有足够的票据！" & _
                        "请先打印其它票据,用完当前领用批次后，重打该单据！", vbInformation, gstrSysName
                Case Else
                    MsgBox "票据领用信息访问失败！将来，你可以重打单据[" & mPrint.strInvoice & "]！", vbInformation, gstrSysName
            End Select
            Cancel = True: Exit Sub
        End If
    End If
    datDate = zlDatabase.Currentdate
    strSQL = "Zl_病人发卡票据_Print("
    '  No_In           Varchar2,
    strSQL = strSQL & "'" & Replace(mPrint.strPrintNo, "'", "") & "'" & ","
    '  票据号_In       票据使用明细.号码%Type,
    strSQL = strSQL & "'" & mPrint.strInvoice & "',"
    '  领用id_In       票据使用明细.领用id%Type,
    strSQL = strSQL & "" & ZVal(lng领用ID) & ","
    '  使用人_In       票据使用明细.使用人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  使用时间_In     票据使用明细.使用时间%Type,
    strSQL = strSQL & "To_Date('" & Format(datDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  操作类型_In     Number
    strSQL = strSQL & mPrint.intPrintMode & ","
    '  票据张数_In     Number := 1,
    strSQL = strSQL & "" & TotalPages & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "票据数据生成")
    
    '不严格控制票据时保存到注册表
    '更新本地票据
    If Not gblnBill发卡 Then
        zlDatabase.SetPara "当前收费票据号", mPrint.strInvoice, glngSys, 1121
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub

Private Sub picCardList_Resize()
    Err = 0: On Error Resume Next
    With picCardList
        IDKind.Top = .ScaleTop + 100
        txtPatient.Top = IDKind.Top
        lblPati.Top = IDKind.Top + (IDKind.Height - lblPati.Height) \ 2
        vsCardList.Left = .ScaleLeft
        vsCardList.Width = .ScaleWidth
        vsCardList.Height = .ScaleHeight - vsCardList.Top
    End With
End Sub
Private Sub picFilter_Resize()
    Err = 0: On Error Resume Next
    With picFilter
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

 

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNotClick Then Exit Sub
    Call zlRefrshListData
End Sub

Private Sub txtCard_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = mtxtIdx.idx_发卡人 Or Index = mtxtIdx.idx_挂失人 Then
        zlCommFun.OpenIme True
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
        zlCommFun.OpenIme False
End Sub

 

Private Sub txtName_Change()
    If zlstr.ActualLen(txtName.Text) < 4 Then
        If chkFilter(0).value = 0 And chkFilter(1).value = 0 Then chkFilter(0).value = 1
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

 Private Sub vsCardList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strCardNo  As String, lng病人ID As Long
    If NewRow <= 0 Then Exit Sub
    If mblnNotRefresh Then Exit Sub
    zl_VsGridRowChange vsCardList, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow = NewRow Then Exit Sub
    Call zlRefrshListData
End Sub
Private Sub zlRefrshListData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新明细数据
    '编制:刘兴洪
    '日期:2012-06-12 14:43:06
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, lng病人ID As Long
    
    If tbPage.Selected Is Nothing Then Exit Sub
    On Error GoTo errHandle
    zlCommFun.ShowFlash "正在装载数据,请稍候..."
    With vsCardList
        If .ColIndex("卡号") < 0 Or .Row < 0 Then Exit Sub
        strCardNo = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
        lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        mstrCurStatus = Trim(.TextMatrix(.Row, .ColIndex("状态")))
        If tbPage.Selected.Tag = Pg_变动记录 Then
            Call mfrmChage.zlReLoadData(mlngCardTypeID, strCardNo)
        ElseIf tbPage.Selected.Tag = Pg_帐户入出记录 Then
            Call mfrmConsume.zlReLoadData(lng病人ID, mlngCardTypeID, strCardNo)
        Else
            Call mfrmFamily.zlReLoadData(lng病人ID, mlngCardTypeID, strCardNo)
        End If
    End With
    zlCommFun.StopFlash
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub vsCardList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlstr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsCardList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlstr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-11-20 15:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim blnCardList As Boolean
    blnCardList = Me.ActiveControl Is vsCardList
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "医疗卡清册"
    
    If CStr(mArrFilter("发卡时间")(0)) <> "1901-01-01" Then
        objRow.Add "发卡时间：" & CStr(mArrFilter("发卡时间")(0)) & "至" & CStr(mArrFilter("发卡时间")(1))
    End If
    If CStr(mArrFilter("挂失时间")(0)) <> "1901-01-01" Then
        objRow.Add "挂失时间：" & CStr(mArrFilter("挂失时间")(0)) & "至" & CStr(mArrFilter("挂失时间")(1))
    End If
    
    If objRow.count > 1 Then
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
    End If
    If mArrFilter("卡号范围")(0) <> "" And mArrFilter("卡号范围")(1) <> "" Then
        objRow.Add "卡号范围：" & CStr(mArrFilter("卡号范围")(0)) & "至" & CStr(mArrFilter("卡号范围")(1))
    ElseIf mArrFilter("卡号范围")(0) = "" And mArrFilter("卡号范围")(1) <> "" Then
        objRow.Add "卡号：" & CStr(mArrFilter("卡号范围")(1))
    ElseIf mArrFilter("卡号范围")(0) <> "" And mArrFilter("卡号范围")(1) = "" Then
        objRow.Add "卡号：" & CStr(mArrFilter("卡号范围")(0))
    End If
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    '由于打印控件不能识别列隐藏属性
    With vsCardList
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = 0
            End If
            
        Next
    End With
    
    Err = 0: On Error GoTo Errhand:
    Set objPrint.Body = vsCardList
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    '恢复
    With vsCardList
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Function SaveCardCancelLose() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 取消挂失
    '编制:刘兴洪
    '日期:2011-06-28 22:41:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDate As String, strCardNo As String, lngRow As Long, i As Long
    Dim lng病人ID As Long, strSQL As String
    
    With vsCardList
        strCardNo = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
        If strCardNo = "" Then Exit Function
        If .TextMatrix(.Row, .ColIndex("挂失时间")) = "" Then Exit Function
        If MsgBox("你真的要对卡号为:“" & .TextMatrix(.Row, .ColIndex("卡号")) & "”的记录进行取消挂失操作吗？" & vbCrLf & _
                    "   『是』: 进行取消挂失操作,取消后的卡片将能进行刷卡等操作！" & vbCrLf & _
                    "   『否』:放弃本次取消挂失操作", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
    End With
    
        'Zl_病人医疗卡信息_取消挂失
        strSQL = "Zl_病人医疗卡信息_取消挂失("
        '  病人id_In     In 病人医疗卡信息.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '  卡类别id_In   In 病人医疗卡信息.卡类别id%Type,
        strSQL = strSQL & "" & mlngCardTypeID & ","
        '  卡号_In       In 病人医疗卡信息.卡号%Type,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  操作员姓名_In In 病人变动记录.操作员姓名%Type
        strSQL = strSQL & "'" & UserInfo.姓名 & "')"
    
        Err = 0: On Error GoTo Errhand:
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        With vsCardList
            .TextMatrix(.Row, .ColIndex("挂失人")) = ""
            .TextMatrix(.Row, .ColIndex("挂失方式")) = ""
            .TextMatrix(.Row, .ColIndex("挂失时间")) = ""
            .TextMatrix(.Row, .ColIndex("状态")) = "有效卡"
        End With
        SaveCardCancelLose = True
        Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
   
Private Sub vsCardList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsCardList
        Select Case Col
        Case .ColIndex("标志")
            Cancel = True
        Case Else
        End Select
    End With
End Sub

Private Sub vsCardList_DblClick()
    Dim strCardNo As String, lng病人ID As Long
    '84755:李南春,2015/5/15,查看医疗卡信息时传入病人id
    With vsCardList
        strCardNo = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
        lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        If strCardNo = "" Then Exit Sub
    End With
    If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_查询, mlngCardTypeID, strCardNo, lng病人ID) = False Then Exit Sub
End Sub

Private Sub vsCardList_GotFocus()
    zl_VsGridGotFocus vsCardList, gSysColor.lngGridColorSel
End Sub

Private Sub vsCardList_LostFocus()
    zl_VsGridLostFocus vsCardList, gSysColor.lngGridColorLost
End Sub

Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = ""
End Sub
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txtEdit(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    Select Case Index
    Case mtxtIdx.idx_发卡人
        If Select人员选择器(Me, txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then
            Exit Sub
        End If
    Case mtxtIdx.idx_挂失人
        If Select人员选择器(Me, txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then
            Exit Sub
        End If
    Case Else
        '由于卡号不知长度,所以无法补位
    End Select
End Sub

Private Sub chkFilter_Click(Index As Integer)
    Select Case Index
    Case 0
        If chkFilter(Index).value = 0 And zlstr.ActualLen(Trim(txtName.Text)) < 4 Then
           If chkFilter(1).value = 0 Then chkFilter(1).value = 1
        End If
    Case 1
        If chkFilter(Index).value = 0 And zlstr.ActualLen(Trim(txtName.Text)) < 4 Then
           If chkFilter(0).value = 0 Then chkFilter(0).value = 1
        End If
    End Select
    dtp开始日期(Index).Enabled = chkFilter(Index).value = 1
    dtp结束日期(Index).Enabled = chkFilter(Index).value = 1
End Sub

Private Sub chkFilter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtp结束日期_Change(Index As Integer)
     If dtp结束日期(Index).value > dtp开始日期(Index).MaxDate Then dtp结束日期(Index).value = dtp开始日期(Index).MaxDate
    If dtp结束日期(Index).value < dtp开始日期(Index).value Then
        dtp开始日期(Index).value = dtp结束日期(Index).value
    End If
End Sub
Private Sub dtp结束日期_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtp开始日期_Change(Index As Integer)
    If dtp开始日期(Index).value > dtp结束日期(Index).MaxDate Then dtp开始日期(Index).value = dtp结束日期(Index).MaxDate
    If dtp结束日期(Index).value < dtp开始日期(Index).value Then
        dtp结束日期(Index).value = dtp开始日期(Index).value
    End If
End Sub
Private Sub dtp开始日期_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub zlPrepayFunc(ByVal intFunc As Integer, ByVal lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:缴预存款
    '入参:intFunc-1-缴预存;2-退预款;3-作废,4-门诊转住院;5-住院转门诊;
    '编制:刘兴洪
    '日期:2011-07-24 18:25:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun As Object, int预交类型 As Integer
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Sub
    'byt预交类型: 0-收预交款(缺省,可切换到退),1-浏览单据(1),2-作废状态(1); 3-余额退款(37770), 4-门诊转住院;5-住院转门诊
    Select Case intFunc
    Case 1  '1.缴预存
        int预交类型 = 0
    Case 2 '退款
        int预交类型 = 3
    Case 3: int预交类型 = 2
    Case 4: int预交类型 = 4
    Case 5: int预交类型 = 5
    End Select
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能： 调用预交款收款窗口
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '   strDBUser:当前数据库登录用户名
    '  bytCallObject:刘兴洪加入(0-预交款调用(缺省的);1-病人费用查询调用,2-医疗卡调用)
    '  lng病人ID-缺省的病人ID
    '  lng主页ID-缺省的主页ID
    '  dblDefPrePayMoney-缺省的预付金额
    Set gfrmCardMgr = Me
    If objFun.PlusDeposit(glngSys, gcnOracle, Me, gstrDBUser, 2, lng病人ID, 0, 0, int预交类型) = False Then
        Set gfrmCardMgr = Nothing
        Exit Sub
    End If
    Set gfrmCardMgr = Nothing
End Sub
 

Private Sub txtPatient_Change()
    Call AutoBrushSet(txtPatient.Text = "")
    If Trim(txtPatient.Text) = "" Then Call ClearData
End Sub
Private Sub txtPatient_GotFocus()
    If Not txtPatient.Enabled Or txtPatient.Locked Then Exit Sub
    Call AutoBrushSet(txtPatient.Text = "")
    zlControl.TxtSelAll txtPatient
    If IsCardType(IDKind, "姓名") Then
        Call zlCommFun.OpenIme(True)
    End If
End Sub
Private Sub txtPatient_LostFocus()
    Call AutoBrushSet(False)
    Call zlCommFun.OpenIme(False)
End Sub
Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean
    Dim blnPass As Boolean
    On Error GoTo errH
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IsCardType(IDKind, "姓名") Then
        '105567:李南春,2017/5/25,卡号加密导致第一个汉字拼音不能触发输入法
        blnPass = txtPatient.PasswordChar <> ""
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        txtPatient.IMEMode = 0
        blnPass = txtPatient.PasswordChar = "" And blnPass
        If blnPass Then
            If txtPatient.SelLength = Len(txtPatient.Text) Then
                txtPatient.Text = ""
            End If
            SendKeys Chr(KeyAscii): KeyAscii = 0: Exit Sub
        End If
    ElseIf IsCardType(IDKind, "门诊号") Or IsCardType(IDKind, "住院号") Or IsCardType(IDKind, "手机号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
         txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
    
    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then
        '不是刷卡和回车,则退出
        Exit Sub
    End If
    
    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
        txtPatient.SelStart = Len(txtPatient.Text)
    End If
    KeyAscii = 0
    If Not GetPatient(txtPatient.Text, blnCard) Then
        If blnCard Then
            Call ClearData: txtPatient.Text = ""
        Else
            Call ClearData: zlControl.TxtSelAll txtPatient
        End If
        Exit Sub
    End If
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub AutoBrushSet(blnAutoRefrsh As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:自动刷新设置
    '编制:刘兴洪
    '日期:2011-06-20 13:31:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(blnAutoRefrsh)
   If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(blnAutoRefrsh)
   Call IDKind.SetAutoReadCard(blnAutoRefrsh)
End Sub
Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
   Dim lngPreIDKind As Long
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        lngPreIDKind = IDKind.IDKind
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("IC卡号")
        txtPatient.Text = strCardNo
        Call txtPatient_KeyPress(vbKeyReturn)
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub
Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
   Dim lngPreIDKind As Long
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub

Public Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除数据
    '编制:刘兴洪
    '日期:2011-06-20 09:29:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    vsCardList.Clear 1
    vsCardList.Rows = 2
End Sub

Private Function GetPatient(ByVal strInput As String, Optional ByVal blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=表示是否就诊卡刷卡
    '出参:
    '返回:病人读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-20 16:04:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim vRect As RECT, rsTmp As ADODB.Recordset
    Dim strSQL As String, strPati As String, strWhere As String, blnHavePass As Boolean
    Dim lng病人ID As Long, blnCancel As Boolean, blnICCard As Boolean
    Dim strPassWord As String, blnBrushCurCardType As Boolean '是否刷的当前卡
    Dim strCardNo As String, rsInfor As ADODB.Recordset
    Dim blnIsMobileNO As Boolean
    mlngCurPatient = 0 '清空变量
    txtPatient.ForeColor = &HFF0000
    blnIsMobileNO = IDKind.IsMobileNo(strInput)
    If IsCardType(IDKind, "IC卡号") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If (blnCard Or IDKind.IDKind = IDKindDefaultKind) _
        And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   '刷卡或缺省的卡
        mlng卡类别ID = Val(IDKind.GetCurCard.接口序号)
        If mlng卡类别ID <= 0 Then
            mlng卡类别ID = IDKind.GetDefaultCardTypeID
        End If
        strCardNo = strInput
        '短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If mlng卡类别ID = mlngCardTypeID Then blnBrushCurCardType = True
        If GetPatiID(mlng卡类别ID, strInput, False, lng病人ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then
            If blnIsMobileNO Then
                '手机号查找
                If GetPatiID("手机号", strInput, False, lng病人ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                    Set rsInfor = New ADODB.Recordset
                    txtPatient.Text = "": Exit Function
                End If
            Else
                If lng病人ID = 0 Then GoTo NotFoundPati:
                Set rsInfor = New ADODB.Recordset
                txtPatient.Text = "": Exit Function
            End If
        End If
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
        blnHavePass = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then   '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strWhere = strWhere & " And A.门诊号=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strWhere = strWhere & " And A.病人ID = (Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号=[1])"
    ElseIf IsCardType(IDKind, "姓名") And blnIsMobileNO Then
        '手机号查找
        If GetPatiID("手机号", strInput, False, lng病人ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.病人ID=[1]"
        strInput = "-" & lng病人ID
    Else
        Select Case IDKind.GetCurCard.名称
            Case "姓名", "姓名或就诊卡"
                '问题号:116787,焦博,2017/12/12,根据“姓名”模糊查询至少要输入两个汉字才去查询,
                '                              增加"门诊预交余额"和"住院预交余额"显示病人余额。
                If Not mblnSeekName Or zlstr.ActualLen(strInput) < 4 Then Exit Function
                strPati = _
                "Select /*+Rule */" & vbNewLine & _
                "       a.病人id As ID, a.病人id, max(a.姓名)as 姓名, max(a.性别)as 性别, max(a.年龄)as 年龄, max(a.病人类型) as 病人类型, max(a.险类)as 险类,max(a.门诊号)as 门诊号," & vbNewLine & _
                "       max( a.住院号)as 住院号, max(a.出生日期)as 出生日期, max(a.身份证号)as 身份证号, max(a.家庭地址) As 现住址, max(a.工作单位)as 工作单位," & vbNewLine & _
                "       LTrim(To_Char(max(Decode(b.类型, 1, Nvl(b.预交余额, 0), 0)), '99999999990.00')) As 门诊预交余额," & vbNewLine & _
                "       LTrim(To_Char(max(Decode(b.类型, 1, 0, Nvl(b.预交余额, 0))), '99999999990.00')) As 住院预交余额," & vbNewLine & _
                "       max(c.卡号) as 卡号 " & vbNewLine & _
                "From 病人信息 A, 病人余额 B,病人医疗卡信息 C" & vbNewLine & _
                "Where a.停用时间 Is Null And a.病人id = b.病人id(+) And a.病人id=c.病人id(+) And b.性质(+) = 1 And Rownum < 101 And " & vbNewLine & _
                "      a.姓名 Like [1] And c.卡类别ID(+)=[2] " & vbNewLine & _
                "group by a.病人id" & vbNewLine & _
                "Order by  姓名,卡号"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人选择", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, _
                                                     blnCancel, False, True, strInput & "%", mlngCardTypeID)
                If blnCancel Then
                    Set rsInfor = New ADODB.Recordset: Exit Function
                End If
                If rsTmp Is Nothing Then GoTo NotFoundPati:
                If rsTmp.State <> 1 Then GoTo NotFoundPati:
                If rsTmp.RecordCount = 0 Then GoTo NotFoundPati:
                lng病人ID = Val(Nvl(rsTmp!病人ID))
                mlngCurPatient = lng病人ID
                '84490:李南春,2015/5/15,通过姓名查找病人成功后读取病人姓名
                txtPatient.Text = Nvl(rsTmp!姓名)
                '74309:李南春，2014-7-7，病人姓名显示颜色处理
                Call SetPatiColor(txtPatient, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), txtPatient.ForeColor, vbRed))
                Call LoadDataToGrid(lng病人ID)
                GetPatient = True
                Exit Function
            Case "医保号"
                strInput = UCase(strInput)
                strWhere = strWhere & "  And A.医保号=[2]"
             Case "身份证号", "二代身份证号", "二代身份证", "身份证"
                strInput = UCase(strInput)
                If GetPatiID("身份证", strInput, False, lng病人ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If GetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.门诊号=[2]"
            Case Else
                '其他类别的号码
                If Val(IDKind.GetCurCard.接口序号) > 0 Then
                    mlng卡类别ID = IDKind.GetCurCard.接口序号
                     If mlng卡类别ID = mlngCardTypeID Then blnBrushCurCardType = True
                    If GetPatiID(mlng卡类别ID, strInput, False, lng病人ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then
                        If lng病人ID = 0 Then GoTo NotFoundPati:
                        Set rsInfor = New ADODB.Recordset
                        txtPatient.Text = "": Exit Function
                    End If
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                    strInput = "-" & lng病人ID
                    strWhere = strWhere & " And A.病人ID=[1]"
                    blnHavePass = True
                Else
                    If GetPatiID(IDKind.GetCurCard.名称, strInput, False, lng病人ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
        End Select
    End If
    
    On Error GoTo errH
    '读取病人信息
    strSQL = "" & _
    "   Select  A.病人ID, A.姓名, A.卡验证码,A.病人类型,A.险类" & _
    "   From 病人信息 A" & _
    "   Where A.停用时间 is NULL " & strWhere
    Set rsInfor = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    If rsInfor.EOF Then GoTo NotFoundPati:
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    txtPatient.PasswordChar = ""
    '74309:李南春，2014-7-7，病人姓名显示颜色处理
    Call SetPatiColor(txtPatient, Nvl(rsInfor!病人类型), IIf(IsNull(rsInfor!险类), txtPatient.ForeColor, vbRed))
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    txtPatient.Text = Nvl(rsInfor!姓名)
    lng病人ID = Val(Nvl(rsInfor!病人ID))
    mlngCurPatient = lng病人ID
    '是以当前卡在刷卡
    Call LoadDataToGrid(lng病人ID, IIf(blnBrushCurCardType, strInput, ""))
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
NotFoundPati:
    If blnBrushCurCardType And strInput <> "" Then
        If Not mbln自制卡 Then
            If MsgBox("未找到指定卡的病人信息,是否进行卡绑定?", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbNo Then Exit Function
            If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_绑定卡, mlngCardTypeID, strCardNo) = False Then Exit Function
        Else
            If MsgBox("未找到指定卡的病人信息,是否进行发卡?", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbNo Then Exit Function
            If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_发卡, mlngCardTypeID, strCardNo) = False Then Exit Function
        End If
        Call LoadDataToGrid(lng病人ID, strCardNo)
    End If
    If blnCard Then
        MsgBox "不能确定病人信息，请检查是否正确刷卡！    ", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    Else
        MsgBox "病人信息未找到,请检查是否输入正确!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    End If
    Set rsInfor = New ADODB.Recordset
End Function
Private Sub InitListReport()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化报表数据(收支清册)
    '编制:刘兴洪
    '日期:2012-06-12 15:26:26
    '问题:50122
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, objPop As Object
    Dim i As Long, j As Long
    With cbsThis.ActiveMenuBar
        For i = 1 To .Controls.count
             If .Controls(i).id = conMenu_ReportPopup Or .Controls(i).Caption Like "报表" Then
                 Set objPop = cbsThis.ActiveMenuBar.Controls(i)
                   
                With objPop.CommandBar
                     For j = 1 To .Controls.count
                        varData = Split(.Controls(j).Parameter & ",,", ",")
                        If varData(1) = "ZL" & glngSys \ 100 & "_INSIDE_1107_2" Then
                            mlngListReportID = .Controls(j).id
                            mstrListReportName = .Controls(j).Caption
                            Exit Sub
                        End If
                     Next
                End With
             End If
        Next
    End With
End Sub
'获取idkind的默认kind值
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind的默认Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.名称)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

'控件名称是否匹配
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case "住院号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "住院号"
     Case "手机号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "手机号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
     End Select
End Function

Private Function Check取消院外卡绑定(lng病人ID As Long, strCardNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消院外卡操作检查
    '入参:lng病人ID - 病人ID; strCardNo - 卡号
    '返回:操作类型:取消绑定,退卡,返回
    '编制:王吉
    '日期:2012-12-19 15:26:26
    '问题:56599
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln本院发卡 As Boolean 'True - 本院发卡 False - 三方绑定卡
    Dim strSQL As String, msgBoxResult As String
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandl:
    strSQL = "" & _
    "   Select count(1) as 本院发卡 From 住院费用记录 Where 记录性质=5 And 病人ID=[1] And 实际票号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, strCardNo)
    If rsTemp.EOF = False Then
        bln本院发卡 = Val(Nvl(rsTemp!本院发卡)) > 0
        If bln本院发卡 = True And mbln发卡 = True Then
            msgBoxResult = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "卡号:" & strCardNo & "为本院发卡,你是否真的取消该卡的绑定?", "取消绑定,退卡,返回", Me, vbQuestion)
            Check取消院外卡绑定 = msgBoxResult
            If Check取消院外卡绑定 = "" Then Check取消院外卡绑定 = "返回"
            Exit Function
        End If
    End If
    Check取消院外卡绑定 = "取消绑定"
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check退卡(strCardNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退卡操作检查
    '入参:strCardNo - 卡号
    '返回:操作类型:取消绑定,退卡,返回
    '编制:王吉
    '日期:2012-12-19 15:26:26
    '问题:56599
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln绑定卡 As Boolean
    Dim strSQL As String, msgBoxResult As String
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandl:
     
    bln绑定卡 = zlIsCardBinding
    If bln绑定卡 = True And mbln发卡 = True And mbln自制卡 = False Then
        msgBoxResult = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "卡号:" & strCardNo & "卡为绑定卡,是否取消绑定?", "是,否", Me, vbQuestion)
        Select Case msgBoxResult
            Case "是"
                Check退卡 = "取消绑定"
            Case "否"
                Check退卡 = "返回"
            Case Else
                Check退卡 = "返回"
        End Select
        Exit Function
    End If
    Check退卡 = "退卡"
    Exit Function
ErrHandl:
     If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreatePublicPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建zlPublicPatient部件
    '返回:创建成功,返回True,否则返回False
    '编制:冉俊明
    '日期:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "病人信息公共部件（zlPublicPatient）创建失败！", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) = False Then
            MsgBox "病人信息公共部件（zlPublicPatient）初始化失败！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CreatePublicPatient = True
End Function

Private Sub PrintBill()
'功能：当前收款记录重新打印一张票据
'bytMode=0-重打,1-补打
    Dim strCardNo As String, lng病人ID As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strOperName As String, strDate As String
    Dim blnStartFactUseType  As Boolean, strUseType As String
    Dim blnHaveData As Boolean, strFormat As String
    Dim objfrmPrint As frmPrint
    
    Set objfrmPrint = New frmPrint
    Load objfrmPrint
    With vsCardList
        lng病人ID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        strCardNo = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
        strOperName = Trim(.TextMatrix(.Row, .ColIndex("发卡人")))
        strDate = .TextMatrix(.Row, .ColIndex("发卡日期"))
        If strCardNo = "" Then ShowMsgbox "没选中相关的医疗卡！": Exit Sub
    End With
    
    If mPrint.intPrintMode = 0 Then
        '打印发卡/绑定卡凭条
        strFormat = IIf(mPrint.bytPrintBoundCard = 0, "", "ReportFormat=" & mPrint.bytPrintBoundCard)
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1107", Me, "卡类别ID=" & mlngCardTypeID, "NO=" & strCardNo, "卡号=" & strCardNo, "缴款=" & 0, "找补=" & 0, "PrintEmpty=0", strFormat, 2)
        Exit Sub
    End If
    
    strSQL = "Select A.No,B.ID From 住院费用记录 A,(Select A.ID,A.NO From 票据打印内容 A,票据使用明细 B Where A.ID = B.打印ID And A.数据性质 = 5 And B.票种 = 1) B " & _
            " Where A.no=B.NO(+) And A.实际票号 = [1] And Nvl(A.结论,0) = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "发卡记录", strCardNo, mlngCardTypeID)
    With rsTemp
        If .EOF Then
            MsgBox "当前卡号" & strCardNo & "的费用数据在后备数据表中!" & vbCrLf _
                & "请与系统管理员联系,转入到在线数据表再操作!", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    If mPrint.intPrintMode = 3 Then
        If Not BillOperCheck(8, strOperName, CDate(strDate), "重打") Then Exit Sub
    Else
        If Not IsNull(rsTemp!id) Then
            MsgBox "当前发卡单据已打印过票据,不能进行补打！", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    If gblnStartFactUseType Then
        mPrint.strUseType = zl_GetInvoiceUserType(lng病人ID, 0, 0)
    End If
    mPrint.strPrintNo = Nvl(rsTemp!NO)
    
    If Not objfrmPrint.RePrintBill(Me, strCardNo, mlngCardTypeID, mPrint.strUseType, _
                mPrint.strPrintNo, mPrint.intPrintMode, mPrint.bytPrintPayCard, True) Then Exit Sub
End Sub

