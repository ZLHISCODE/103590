VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareCardManager 
   Caption         =   "消费卡管理"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   Icon            =   "frmSquareCardManager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCardList 
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   240
      ScaleHeight     =   2565
      ScaleWidth      =   11055
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   630
      Width           =   11055
      Begin VB.PictureBox picModify 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   4845
         ScaleHeight     =   465
         ScaleWidth      =   4650
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   -90
         Visible         =   0   'False
         Width           =   4650
         Begin VB.ComboBox cboCardType 
            Height          =   300
            Left            =   1605
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   145
            Width           =   1620
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "完成修改(&O)"
            Height          =   350
            Left            =   3330
            TabIndex        =   13
            Top             =   120
            Width           =   1230
         End
         Begin VB.CheckBox chkModify 
            Caption         =   "修改卡类型(&X)"
            Height          =   350
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Value           =   2  'Grayed
            Width           =   1425
         End
         Begin MSComCtl2.DTPicker dtpValidDate 
            Height          =   300
            Left            =   1605
            TabIndex        =   15
            Top             =   150
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   181469187
            CurrentDate     =   40156.0854282407
         End
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "回收"
         Height          =   405
         Index           =   3
         Left            =   4050
         TabIndex        =   10
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "退卡"
         Height          =   405
         Index           =   2
         Left            =   3210
         TabIndex        =   9
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "失效卡"
         Height          =   405
         Index           =   1
         Left            =   2235
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "有效卡"
         Height          =   405
         Index           =   0
         Left            =   1305
         TabIndex        =   7
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCardList 
         Height          =   2055
         Left            =   105
         TabIndex        =   4
         Top             =   435
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
         Cols            =   24
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSquareCardManager.frx":6852
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
            TabIndex        =   5
            Top             =   60
            Width           =   210
            Begin VB.Image imgCol 
               Height          =   195
               Left            =   0
               Picture         =   "frmSquareCardManager.frx":6B7F
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "当前卡信息"
         Height          =   180
         Left            =   210
         TabIndex        =   6
         Top             =   90
         Width           =   900
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   4590
      ScaleHeight     =   1875
      ScaleWidth      =   5070
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5175
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   1
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
      TabIndex        =   2
      Top             =   8025
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSquareCardManager.frx":70CD
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10372
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "frmSquareCardManager.frx":7961
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   645
      Top             =   3500
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
            Picture         =   "frmSquareCardManager.frx":8CBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareCardManager.frx":900F
            Key             =   ""
         EndProperty
      EndProperty
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
      Bindings        =   "frmSquareCardManager.frx":9363
      Left            =   1005
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSquareCardManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mblnFirst As Boolean, mstrTitle As String     '功能标题

Private Enum mPaneID
    Pane_Search = 1     '搜索条件
    Pane_CardLists = 2  '卡列表
    Pane_CardDetails = 3    '详细列表
End Enum
Private Enum mPgIndex
    Pg_充值记录 = 250101
    Pg_回收记录 = 250102
    Pg_消费记录 = 250103
    Pg_条件过滤 = 250104
End Enum

Private mrs消费卡 As ADODB.Recordset
Private Type Ty_CurrentCardType '当前卡类别信息
    lng编号 As Long
    bln启用 As Boolean
    bln严格控制 As Boolean
    bln特定病人 As Boolean
    bln允许换卡 As Boolean
    bln允许余额退款 As Boolean
    bln允许补卡 As Boolean
End Type
Private mTy_CurCardType As Ty_CurrentCardType

Private Enum m_CardStatus '消费卡状态
    Normal = 1 '正常
    Recycled = 2 '回收
    Refunded = 3 '退卡
    Stoped = 4 '停用
    Invalid = 5 '失效
End Enum

Private Type Ty_CurrentCard '当前卡信息
    blnHaveData As Boolean
    lng卡ID As Long
    str卡号 As String
    str卡类型 As String
    byt卡状态 As m_CardStatus
    bln已消费 As Boolean
    bln允许充值回退 As Boolean
    str发卡人 As String
    str发卡日期 As String 'yyyy-MM-dd HH:mm:ss
    bln充值卡 As Boolean
    lng发卡序号 As Long
End Type
Private mTy_CurCard As Ty_CurrentCard

Private WithEvents mfrmFilter As frmSquareCardFilter
Attribute mfrmFilter.VB_VarHelpID = -1
Private mfrmSquareCardCallBack As frmSquareCardCallBack
Private WithEvents mfrmSquareCardConsume As frmSquareCardConsume
Attribute mfrmSquareCardConsume.VB_VarHelpID = -1
Private WithEvents mfrmSquareCardInFull As frmSquareCardInFul
Attribute mfrmSquareCardInFull.VB_VarHelpID = -1
Private mcllSubFrm As New Collection

Private mArrFilter As Variant
Private mstrPrivs_RollingCurtain As String  '收费轧帐管理权限
Private mblnPrinting As Boolean

Public Sub ShowList(ByVal lngModule As Long, ByVal strTitle As String, ByVal frmMain As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口,显示相关的项目及分类信息
    '编制:刘兴洪
    '日期:2009-11-19 15:38:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrTitle = strTitle
    If Not zlCheckDepend Then Exit Sub            '数据依赖性测试
    
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        zlCommFun.ShowChildWindow Me.hWnd, frmMain
    End If
    Me.ZOrder 0
End Sub

Private Function zlCheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据依赖性
    '返回:数据合法,返回true，否则返回False
    '编制:刘兴洪
    '日期:2009-11-19 15:37:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, strSQL As String
    
    On Error GoTo errHandle:
    Set rsTemp = Get结算方式("消费卡", "1,2,8")
    If rsTemp.EOF Then
        ShowMsgbox "消费卡场合没有可用的结算方式，请先到结算方式管理中设置。"
        Exit Function
    End If
    
    strSQL = "Select 编码,名称, 缺省面额, 缺省折扣, 缺省标志 From 消费卡类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "注意:" & vbCrLf & "   没有设置相关的消费卡类型，请先在[字典管理]中设置！"
        Exit Function
    End If
    Do While Not rsTemp.EOF
        cboCardType.AddItem NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
        rsTemp.MoveNext
    Loop
    If cboCardType.ListCount > 0 Then cboCardType.ListIndex = 0
    zlCheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2009-11-20 16:02:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strHead As String, varHead As Variant
    
    strHead = "标志,1,285|ID,1,0|卡号,1,945|卡类型,4,780|当前状态,4,900|充值卡,4,600|有效期,4,1935|发卡人,1,900|发卡时间,4,1850|" & _
            "领卡人,1,900|领卡部门,1,1600|回收人,1,900|回收时间,4,1850|限制类别,1,1860|充值折扣率,7,990|面额,7,720|销售金额,7,885|" & _
            "当前余额,7,1020|已消费,4,900|停用人,1,900|停用时间,4,1850|备注,1,900|发卡序号,1,0"
    varHead = Split(strHead, "|")
    With vsCardList
        .Cols = UBound(varHead) + 1
        For i = 0 To UBound(varHead)
            .TextMatrix(0, i) = Split(varHead(i), ",")(0)
            .ColKey(i) = Split(varHead(i), ",")(0)
            If .TextMatrix(0, i) = "标志" Then .TextMatrix(0, i) = ""
            .ColAlignment(i) = Split(varHead(i), ",")(1)
            .ColWidth(i) = Split(varHead(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("标志")) = "-1|1"
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("卡号")) = "1|0"
        .ColData(.ColIndex("当前余额")) = "1|0"
        .ColData(.ColIndex("发卡序号")) = "-1|1"
    End With
End Sub

Private Function InitPage() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2009-11-19 15:15:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As TabControlItem
    
    Err = 0: On Error GoTo errHand:
    Set mfrmSquareCardInFull = New frmSquareCardInFul
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_充值记录, "充值信息", mfrmSquareCardInFull.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_充值记录
    mcllSubFrm.Add mfrmSquareCardInFull, CStr(objItem.Tag)
    
    Set mfrmSquareCardCallBack = New frmSquareCardCallBack
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_回收记录, "回收信息", mfrmSquareCardCallBack.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_回收记录
    mcllSubFrm.Add mfrmSquareCardCallBack, CStr(objItem.Tag)


    Set mfrmSquareCardConsume = New frmSquareCardConsume
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_消费记录, "消费信息", mfrmSquareCardConsume.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_消费记录
    mcllSubFrm.Add mfrmSquareCardConsume, CStr(objItem.Tag)

     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    InitPage = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitPanel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '编制:刘兴洪
    '日期:2009-11-18 16:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHand:
    Set mfrmFilter = New frmSquareCardFilter
    Call mfrmFilter.Init条件(mlngModule, mstrPrivs)
    mcllSubFrm.Add mfrmFilter, CStr(mPgIndex.Pg_条件过滤)

    With dkpMan
        Set objPane = .CreatePane(mPaneID.Pane_Search, 260, 400, DockLeftOf, Nothing)
        objPane.Title = "条件设置"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable
        objPane.MinTrackSize.Width = 260: objPane.MaxTrackSize.Width = 260
        
        Set objPane = .CreatePane(mPaneID.Pane_CardLists, 400, 400, DockRightOf, objPane)
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        
        Set objPane = .CreatePane(mPaneID.Pane_CardDetails, 400, 400, DockBottomOf, objPane)
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        
        .SetCommandBars cbsThis
        .ImageList = imlPaneIcons
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    InitPanel = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '入参:
    '出参:
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-11-18 16:53:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim intIndex As Integer
      
    Err = 0: On Error GoTo errHand:
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
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.id = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "打印预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "收费轧帐(&M)"): cbrControl.BeginGroup = True
        cbrControl.IconId = 227
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSingleBill, "打印缴款单(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With


    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.id = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        '83399:李南春,2015/7/19,消费卡类别设置
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加卡类(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改卡类(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除卡类(&D)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "发卡(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardModify, "修改(&M)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Cardtrade, "换卡(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardFill, "补卡(&F)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "退卡(&B)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelBack, "取消退卡(&K)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardCallBack, "回收(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelCallBack, "取消回收(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardStop, "停用(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardResume, "启用(&F)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "充值(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "充值回退(&T)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardBackMoney, "余额退款(&E)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ChangePassWord, "修改密码(&G)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ChangePassWord_Force, "强制修改密码(&O)")
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.id = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.id = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        
        .Add FCONTROL, Asc("S"), conMenu_Edit_CardPay
        .Add FCONTROL, Asc("M"), conMenu_Edit_CardModify
        .Add FCONTROL, Asc("C"), conMenu_Edit_CardInFull
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F11, conMenu_File_FeeCollect
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "发卡"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Cardtrade, "换卡"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardFill, "补卡")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "退卡"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardCallBack, "回收")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardStop, "停用"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardResume, "启用")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "充值"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "充值回退")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "收费轧帐"): cbrControl.BeginGroup = True
        cbrControl.IconId = 227
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
        '消费卡类别
        Set cbrControl = .Add(xtpControlComboBox, conMenu_COMBOX_INTERFACE, "卡类别")
        cbrControl.Flags = xtpFlagRightAlign
        cbrControl.Width = 160
        cbrControl.Style = xtpComboLabel '显示文本标签，注意在批量设置的时候排除
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.type <> xtpControlComboBox And cbrControl.type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
 
    zlDefCommandBars = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetCurrentCardType(ByVal lng接口编号 As Long)
    '功能:设置当前卡类别的信息
    Dim ty_Temp As Ty_CurrentCardType
    
    On Error GoTo errHandle
    mTy_CurCardType = ty_Temp '自定义Type初始化
    
    If lng接口编号 = 0 Then Exit Sub
    If mrs消费卡 Is Nothing Then Exit Sub
    
    mrs消费卡.Filter = "编号=" & lng接口编号
    If mrs消费卡.RecordCount = 0 Then Exit Sub
    
    With mTy_CurCardType
        .lng编号 = Val(NVL(mrs消费卡!编号))
        .bln启用 = Val(NVL(mrs消费卡!启用)) = 1
        .bln严格控制 = Val(NVL(mrs消费卡!是否严格控制)) = 1
        .bln特定病人 = Val(NVL(mrs消费卡!是否特定病人)) = 1
        .bln允许换卡 = Val(NVL(mrs消费卡!是否允许换卡)) = 1
        .bln允许补卡 = Val(NVL(mrs消费卡!是否允许补卡)) = 1
        .bln允许余额退款 = Val(NVL(mrs消费卡!是否允许余额退款)) = 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Function RowHidden() As Boolean
    '当前行是否可见
    On Error GoTo errHandle
    If vsCardList.Row < 0 Then Exit Function
    RowHidden = vsCardList.RowHidden(vsCardList.Row)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnValidCardType As Boolean, blnValidCard As Boolean
    
    If Me.Visible = False Then Exit Sub

    Err = 0: On Error Resume Next
    blnValidCardType = mTy_CurCardType.lng编号 > 0 And mTy_CurCardType.bln启用 '正常启用的卡类别
    blnValidCard = mTy_CurCard.blnHaveData And RowHidden() = False _
                And (mTy_CurCard.byt卡状态 = m_CardStatus.Normal Or mTy_CurCard.byt卡状态 = m_CardStatus.Invalid) '有效卡
    
    Select Case Control.id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = mTy_CurCard.blnHaveData
    Case conMenu_File_PrintSingleBill '打印缴款单
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "消费卡收费收据")
        Control.Enabled = Control.Visible And blnValidCardType And blnValidCard
    Case conMenu_File_FeeCollect '收费轧帐
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs_RollingCurtain, "轧帐")
        Control.Enabled = Control.Visible

    Case conMenu_Edit_NewItem '增加消费卡类别
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "增加")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify  '修改消费卡类别
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "修改")
        Control.Enabled = Control.Visible And mTy_CurCardType.lng编号 > 0
    Case conMenu_Edit_Delete '删除消费卡类别
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "删除")
        Control.Enabled = Control.Visible And mTy_CurCardType.lng编号 > 0
    
    Case conMenu_Edit_CardPay
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "发卡")
        Control.Enabled = Control.Visible And blnValidCardType
    Case conMenu_Edit_CardModify
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "修改卡信息")
        Control.Enabled = Control.Visible And blnValidCardType And blnValidCard
    Case conMenu_Edit_Cardtrade
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "换卡") And mTy_CurCardType.bln允许换卡
        Control.Enabled = Control.Visible And blnValidCardType
    Case conMenu_Edit_CardFill
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "补卡") And mTy_CurCardType.bln特定病人 And mTy_CurCardType.bln允许补卡
        Control.Enabled = Control.Visible And blnValidCardType
    
    Case conMenu_Edit_CardBack
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "退卡")
        Control.Enabled = Control.Visible And blnValidCardType And blnValidCard And Not mTy_CurCard.bln已消费
    Case conMenu_Edit_CardCancelBack
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "退卡")
        Control.Enabled = Control.Visible And blnValidCardType _
                        And mTy_CurCard.blnHaveData And mTy_CurCard.byt卡状态 = Refunded And RowHidden() = False _
                        And Not mTy_CurCard.bln已消费
    
    Case conMenu_Edit_CardCallBack
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "回收")
        Control.Enabled = Control.Visible And blnValidCardType
    Case conMenu_Edit_CardCancelCallBack
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "回收")
        Control.Enabled = Control.Visible And blnValidCardType _
                        And mTy_CurCard.blnHaveData And mTy_CurCard.byt卡状态 = m_CardStatus.Recycled And RowHidden() = False
    
    Case conMenu_Edit_CardStop
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "卡片停用")
        Control.Enabled = Control.Visible And blnValidCardType And blnValidCard
    Case conMenu_Edit_CardResume
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "卡片启用")
        Control.Enabled = Control.Visible And blnValidCardType _
                        And mTy_CurCard.blnHaveData And mTy_CurCard.byt卡状态 = m_CardStatus.Stoped And RowHidden() = False
    
    Case conMenu_Edit_CardInFull
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "充值")
        Control.Enabled = Control.Visible And blnValidCardType
    Case conMenu_Edit_CardInFullBack
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "回退充值")
        Control.Enabled = Control.Visible And blnValidCardType And blnValidCard And mTy_CurCard.bln允许充值回退
    Case conMenu_Edit_CardBackMoney
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "余额退款") And mTy_CurCardType.bln允许余额退款
        Control.Enabled = Control.Visible And blnValidCardType

    Case conMenu_View_ToolBar_Button: Control.Checked = cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text: Control.Checked = Not (cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size: Control.Checked = cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_1" _
                            And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_2"
        End If
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim lngID As Long, lng充值ID As Long
    Dim blnValidCard As Boolean
    
    blnValidCard = mTy_CurCard.blnHaveData And RowHidden() = False _
                And (mTy_CurCard.byt卡状态 = m_CardStatus.Normal Or mTy_CurCard.byt卡状态 = m_CardStatus.Invalid) '有效卡
    
    On Error GoTo errHand
    Select Case Control.id
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_Parameter '参数设置
        Call frmSquareCardParaSet.ShowParaSet(Me, mlngModule, mstrPrivs)
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_PrintSingleBill '打印缴款单
        Call PrintReBill
    Case conMenu_File_FeeCollect '收费轧帐
        Call zlExecuteChargeRollingCurtain(Me)
    
    Case conMenu_Edit_NewItem '新增卡类别
        If frmSquareSendCardTypeEdit.zlEditSendCard(Me, mlngModule, mstrPrivs, gSendCardEdit.Card_增加) = False Then Exit Sub
        Call LoadCardTypeData
    Case conMenu_Edit_Modify  '修改卡类别
        If frmSquareSendCardTypeEdit.zlEditSendCard(Me, mlngModule, mstrPrivs, gSendCardEdit.Card_修改, mTy_CurCardType.lng编号) = False Then Exit Sub
        Call LoadCardTypeData
    Case conMenu_Edit_Delete     '删除卡类别
        If DeleteCardType() Then Call LoadCardTypeData
    
    Case conMenu_Edit_CardPay    '发卡
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_发卡, mTy_CurCardType.lng编号) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardModify    '修改
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_修改, mTy_CurCardType.lng编号, mTy_CurCard.lng卡ID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_Cardtrade '换卡
        If blnValidCard Then lngID = mTy_CurCard.lng卡ID
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_换卡, mTy_CurCardType.lng编号, lngID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardFill '补卡
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_补卡, mTy_CurCardType.lng编号) = False Then Exit Sub
        Call LoadDataToRpt
    
    Case conMenu_Edit_CardBack    '退卡
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_退卡, mTy_CurCardType.lng编号, mTy_CurCard.lng卡ID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardCancelBack   '取消退卡
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_取消退卡, mTy_CurCardType.lng编号, mTy_CurCard.lng卡ID) = False Then Exit Sub
        Call LoadDataToRpt
    
    Case conMenu_Edit_CardCallBack    '回收
        '停用的、回收的和正常卡允许回收
        If blnValidCard Or (mTy_CurCard.blnHaveData And RowHidden() = False And mTy_CurCard.byt卡状态 <> m_CardStatus.Stoped) Then
            lngID = mTy_CurCard.lng卡ID
        End If
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_回收, mTy_CurCardType.lng编号, lngID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardCancelCallBack  '取消回收
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_取消回收, mTy_CurCardType.lng编号, mTy_CurCard.lng卡ID) = False Then Exit Sub
        Call LoadDataToRpt
    
    Case conMenu_Edit_CardResume        '卡片启用
        If CardStopAndStart(False) Then Call LoadDataToRpt
    Case conMenu_Edit_CardStop        '卡片停用
        If CardStopAndStart(True) Then Call LoadDataToRpt
    
    Case conMenu_Edit_CardInFull    '充值
        If blnValidCard And mTy_CurCard.bln充值卡 Then lngID = mTy_CurCard.lng卡ID
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_充值, mTy_CurCardType.lng编号, lngID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardInFullBack    '充值回退
        lng充值ID = mfrmSquareCardInFull.zlGet充值ID()
        If lng充值ID = 0 Then Exit Sub
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_充值回退, _
            mTy_CurCardType.lng编号, mTy_CurCard.lng卡ID, lng充值ID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardBackMoney '余额退款
        If frmSquareRefundBalance.ShowMe(Me, mlngModule, mstrPrivs, mTy_CurCardType.lng编号) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_ChangePassWord    '修改密码
        Call frmModiCardPass.zlModifyPass(Me, mlngModule, mTy_CurCardType.lng编号, True)
    Case conMenu_Edit_ChangePassWord_Force  '强制修改密码
        Call frmModiCardPass.zlModifyPass(Me, mlngModule, mTy_CurCardType.lng编号, False)
    Case conMenu_View_Refresh '刷新
        Call LoadDataToRpt
    Case conMenu_COMBOX_INTERFACE '点击选择卡类别
        If Val(Control.Category) = Control.ItemData(Control.ListIndex) Then Exit Sub
        Call SetCurrentCardType(Control.ItemData(Control.ListIndex))
        Call LoadDataToRpt
        Control.Category = Control.ItemData(Control.ListIndex)
        
    Case conMenu_View_StatusBar
        stbThis.Visible = Not stbThis.Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        cbsThis(2).Visible = Not cbsThis(2).Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In cbsThis(2).Controls
            If cbrControl.type <> xtpControlComboBox And cbrControl.type <> xtpControlLabel Then
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            End If
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
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zlOpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    zlControl.ControlSetFocus vsCardList
    Call vsCardList_GotFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim strShow As String, i As Long
    
    On Error GoTo errHandle
    mblnFirst = True
    mstrPrivs = gstrPrivs
    Me.Caption = mstrTitle
    
    mstrPrivs_RollingCurtain = GetPrivFunc(glngSys, 1506)
    mTy_CurCardType.lng编号 = Val(zlDatabase.GetPara("上次接口号", glngSys, mlngModule, 0))
    Set mrs消费卡 = zlGet消费卡接口(, False)
    
    strShow = Trim(zlDatabase.GetPara("卡显示方式", glngSys, mlngModule, "1011"))
    If Len(strShow) < 4 Then strShow = strShow & "1111"
    For i = 0 To 3
        chkStatus(i).value = IIf(Val(Mid(strShow, i + 1, 1)) = 1, vbChecked, vbUnchecked)
    Next
    dtpValidDate.MinDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    dtpValidDate.value = DateAdd("d", 1, dtpValidDate.MinDate)
    chkModify.value = vbUnchecked
    
    Call InitPanel
    Call InitPage
    Call zlDefCommandBars '初始菜单及工具栏
    Call InitVsGrid
    Set mArrFilter = mfrmFilter.GetFilterCon
    Call LoadCardTypeData
    
    RestoreWinState Me, App.ProductName, mstrTitle
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTemp As String
    
    strTemp = ""
    For i = 0 To 3
        strTemp = strTemp & IIf(chkStatus(i).value = 1, 1, 0)
    Next
    zlDatabase.SetPara "卡显示方式", strTemp, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "上次接口号", mTy_CurCardType.lng编号, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlstr.IsHavePrivs(mstrPrivs, "参数设置")
    SaveWinState Me, App.ProductName, mstrTitle
   
    '关闭子窗口
    For i = 1 To mcllSubFrm.count
        If Not mcllSubFrm(i) Is Nothing Then Unload mcllSubFrm(i)
    Next
End Sub

Private Sub chkStatus_Click(Index As Integer)
    Call SetCardRowColHide
End Sub

Private Sub chkModify_Click()
    Call SetModifyEnabled
End Sub

Private Sub cmdModify_Click()
   Call SaveBatchUpdateCardInfor
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case mPaneID.Pane_Search    '搜索条件窗体
        Item.Handle = mfrmFilter.hWnd
    Case mPaneID.Pane_CardLists '卡列表
        Item.Handle = picCardList.hWnd
    Case mPaneID.Pane_CardDetails   '详细卡信息
        Item.Handle = picList.hWnd
    End Select
End Sub

Private Sub SetCardRowColHide(Optional lngLocalRow As Long = -1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置行的显示和隐藏
    '入参:lngLocalRow -指定行(-1代表全部重新设置)
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-22 21:05:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngRows As Long, i As Long
    Dim lngCurRow As Long
    
    Err = 0: On Error GoTo errHand:
    With vsCardList
        i = 1: lngRows = .Rows - 1
        If lngLocalRow < 0 Then
            .Redraw = flexRDNone
        Else
            i = lngLocalRow: lngRows = lngLocalRow
        End If
        
        lngCurRow = -2
        For lngRow = i To lngRows
            .RowHidden(lngRow) = False
            Select Case Val(.Cell(flexcpData, lngRow, .ColIndex("当前状态")))
            Case m_CardStatus.Normal
                If chkStatus(0).value = 0 Then .RowHidden(lngRow) = True
            Case m_CardStatus.Recycled
                If chkStatus(3).value = 0 Then .RowHidden(lngRow) = True
            Case m_CardStatus.Refunded
                If chkStatus(2).value = 0 Then .RowHidden(lngRow) = True
            Case m_CardStatus.Invalid
                If chkStatus(1).value = 0 Then .RowHidden(lngRow) = True
            End Select
            If .RowHidden(lngRow) = False Then
                If lngCurRow < .Row Then lngCurRow = lngRow
            End If
        Next
        If lngLocalRow < 0 Then
            If lngCurRow > 0 Then
                If .Row > 0 Then
                    If .RowHidden(.Row) Then .Row = lngCurRow
                Else
                    .Row = lngCurRow
                End If
            Else
                .Row = -1
            End If
            .Redraw = flexRDBuffered
        End If
    End With
    Exit Sub
errHand:
    vsCardList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mfrmFilter_zlRefreshCon(ByVal arrFilter As Variant)
    Set mArrFilter = arrFilter
    '重新加载数据
    Call LoadDataToRpt
End Sub

Private Function LoadDataToRpt() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-11-19 15:43:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strSubWhere As String, lngRow As Long, lngPre消费ID As Long
    Dim rsTemp As ADODB.Recordset, strCurDate As String, strSQL As String
    
    Err = 0: On Error GoTo errHand:
    If mTy_CurCardType.lng编号 <= 0 Then Exit Function
    
    If mArrFilter("发卡时间")(0) <> "1901-01-01" And mArrFilter("回收时间")(0) <> "1901-01-01" Then
        strSubWhere = strSubWhere & " And (发卡时间 Between [1] And [2] Or 回收时间 Between [3] And [4])"
    ElseIf mArrFilter("发卡时间")(0) <> "1901-01-01" Then
        strSubWhere = strSubWhere & " And 发卡时间 Between [1] And [2]"
    ElseIf mArrFilter("回收时间")(0) <> "1901-01-01" Then
        strSubWhere = strSubWhere & " And 回收时间 Between [3] And [4]"
    End If
    If mArrFilter("卡号范围")(0) <> "" And mArrFilter("卡号范围")(1) <> "" Then
        strSubWhere = strSubWhere & " And 卡号 Between [5] And [6]"
    ElseIf mArrFilter("卡号范围")(0) <> "" Then
        strWhere = strWhere & " And a.卡号=[5]"
    ElseIf mArrFilter("卡号范围")(1) <> "" Then
        strWhere = strWhere & " And a.卡号=[6]"
    End If
    If strSubWhere = "" Then
        '如果没有结定时间范围,就只能查找当前的领卡人和发卡人
        If mArrFilter("领卡人") <> "" Then strWhere = strWhere & " And a.领卡人 like [7]"
        If mArrFilter("发卡人") <> "" Then strWhere = strWhere & " And a.发卡人 like [8]"
    Else
        If mArrFilter("领卡人") <> "" Then strSubWhere = strSubWhere & " And 领卡人 like [7]"
        If mArrFilter("发卡人") <> "" Then strSubWhere = strSubWhere & " And 发卡人 like [8]"
    End If
    
    If Trim(mArrFilter("卡类型")) <> "所有" Then strWhere = strWhere & " And a.卡类型 = [9]"
    
    If Val(mArrFilter("包含停用卡")) = 1 Then
        strWhere = strWhere & " And a.当前状态 <= 9"   '需要用到索引
    Else
        strWhere = strWhere & " And (a.停用日期 Is Null Or a.停用日期 >= To_Date('3000-01-01', 'yyyy-mm-dd'))"   '需要用到索引
    End If
    
    strSQL = _
    "Select a.Id, a.卡类型, a.卡号, a.序号, a.限制类别, a.可否充值, a.有效期, a.发卡时间, a.发卡人, a.领卡人," & vbNewLine & _
    "       a.回收人, a.回收时间, a.备注, a.卡面金额, a.销售金额," & vbNewLine & _
    "       a.充值折扣率, a.余额, a.停用人, a.停用日期, Decode(b.编码, Null, '', b.编码 || '-' || b.名称) As 领卡部门," & vbNewLine & _
    "       Nvl((Select 1 From 病人卡结算记录 Where 消费卡id = a.Id And 记录性质 = 4 And Rownum < 2), 0) As 消费," & vbNewLine & _
    "       Case" & vbNewLine & _
    "          When a.当前状态 = 1 And Nvl(a.有效期, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 5" & vbNewLine & _
    "          When a.当前状态 = 1 And Nvl(a.停用日期, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 4" & vbNewLine & _
    "          When a.当前状态 = 4 Then 2" & vbNewLine & _
    "          When a.当前状态 = 5 Then 4" & vbNewLine & _
    "          Else Nvl(a.当前状态, 1)" & vbNewLine & _
    "       End As 当前状态, a.发卡序号" & vbNewLine
    If strSubWhere <> "" Then
        strSubWhere = Mid(strSubWhere, 6)
        strSQL = strSQL & _
        "From 消费卡信息 A, 部门表 B, " & vbNewLine & _
        "      (Select 卡号, Max(序号) As 序号 From 消费卡信息 Where " & strSubWhere & " And 接口编号 = [10] Group By 卡号) C" & vbNewLine & _
        "Where a.领卡部门id = b.Id(+) And a.卡号 = c.卡号 And a.序号 = c.序号 And a.接口编号 = [10] " & strWhere
    Else
        strSQL = strSQL & _
        "From 消费卡信息 A,部门表 B" & _
        "Where a.领卡部门id = b.Id(+) and a.接口编号=[10] " & strWhere
    End If
    strSQL = strSQL & vbNewLine & _
            "Order By 发卡时间 Desc,卡号"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(mArrFilter("发卡时间")(0)), CDate(mArrFilter("发卡时间")(1)), _
        CDate(mArrFilter("回收时间")(0)), CDate(mArrFilter("回收时间")(1)), _
        CStr(mArrFilter("卡号范围")(0)), CStr(mArrFilter("卡号范围")(1)), _
        CStr(mArrFilter("领卡人")), CStr(mArrFilter("发卡人")), _
        CStr(mArrFilter("卡类型")), mTy_CurCardType.lng编号)
    
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With vsCardList
        If .Row > 0 Then lngPre消费ID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
        
        .Redraw = flexRDNone
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        .Cell(flexcpForeColor, 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("卡号")) = NVL(rsTemp!卡号)
            .TextMatrix(lngRow, .ColIndex("ID")) = NVL(rsTemp!id)
            .TextMatrix(lngRow, .ColIndex("卡类型")) = NVL(rsTemp!卡类型)
            .TextMatrix(lngRow, .ColIndex("充值卡")) = IIf(Val(NVL(rsTemp!可否充值)) = 1, "√", "")
            .TextMatrix(lngRow, .ColIndex("有效期")) = Format(rsTemp!有效期, "yyyy-MM-dd HH:mm:ss")
            If Trim(.TextMatrix(lngRow, .ColIndex("有效期"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("有效期")) = ""
            
            .TextMatrix(lngRow, .ColIndex("发卡人")) = NVL(rsTemp!发卡人)
            .TextMatrix(lngRow, .ColIndex("发卡时间")) = Format(rsTemp!发卡时间, "yyyy-MM-dd HH:mm:ss")
            
            .TextMatrix(lngRow, .ColIndex("领卡人")) = NVL(rsTemp!领卡人)
            .TextMatrix(lngRow, .ColIndex("领卡部门")) = NVL(rsTemp!领卡部门)
            
            .TextMatrix(lngRow, .ColIndex("停用人")) = NVL(rsTemp!停用人)
            .TextMatrix(lngRow, .ColIndex("停用时间")) = Format(rsTemp!停用日期, "yyyy-MM-dd HH:mm:ss")
            If Trim(.TextMatrix(lngRow, .ColIndex("停用时间"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("停用时间")) = ""
            
            .TextMatrix(lngRow, .ColIndex("回收人")) = NVL(rsTemp!回收人)
            .TextMatrix(lngRow, .ColIndex("回收时间")) = Format(rsTemp!回收时间, "yyyy-MM-dd HH:mm:ss")
            If Trim(.TextMatrix(lngRow, .ColIndex("回收时间"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("回收时间")) = ""
            
            .TextMatrix(lngRow, .ColIndex("限制类别")) = NVL(rsTemp!限制类别)
            .TextMatrix(lngRow, .ColIndex("充值折扣率")) = Format(rsTemp!充值折扣率, "###0.00;-###0.00;;")
            
            .TextMatrix(lngRow, .ColIndex("面额")) = Format(rsTemp!卡面金额, "###0.00;-###0.00;;")
            .TextMatrix(lngRow, .ColIndex("销售金额")) = Format(rsTemp!销售金额, "###0.00;-###0.00;;")
            
            .TextMatrix(lngRow, .ColIndex("当前余额")) = Format(rsTemp!余额, "###0.00;-###0.00;;")
            .TextMatrix(lngRow, .ColIndex("已消费")) = IIf(Val(NVL(rsTemp!消费)) = 1, "√", "")
            .TextMatrix(lngRow, .ColIndex("备注")) = NVL(rsTemp!备注)
            
            .TextMatrix(lngRow, .ColIndex("当前状态")) = _
                decode(Val(NVL(rsTemp!当前状态)), _
                            m_CardStatus.Recycled, "回收", _
                            m_CardStatus.Refunded, "退卡", _
                            m_CardStatus.Invalid, "失效", _
                            m_CardStatus.Stoped, "停用", "有效")
            .Cell(flexcpData, lngRow, .ColIndex("当前状态")) = Val(NVL(rsTemp!当前状态))
            .TextMatrix(lngRow, .ColIndex("发卡序号")) = Val(NVL(rsTemp!发卡序号))
            
            If lngPre消费ID = Val(NVL(rsTemp!id)) Then
                .Row = lngRow
                If .RowIsVisible(.Row) = False Then .TopRow = .Row
            End If
            
            '设置颜色行
            Call SetGridRowForeColor(lngRow)
            SetCardRowColHide lngRow
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Row <= 0 Then .Row = 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModule, vsCardList, Me.Name, "卡信息列表", True, True
    
    Call vsCardList_AfterRowColChange(-1, 0, vsCardList.Row, 0)
    LoadDataToRpt = True
    Exit Function
errHand:
    vsCardList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SetGridRowForeColor(ByVal lngRow As Long)
    '根据状态，设置行颜色
    Dim lngColor As Long, int状态 As Integer
    
    With vsCardList
        Select Case Val(.Cell(flexcpData, lngRow, .ColIndex("当前状态")))
        Case m_CardStatus.Stoped
            lngColor = vbRed
        Case m_CardStatus.Invalid
            lngColor = &HFF00FF
        Case m_CardStatus.Recycled, m_CardStatus.Refunded
            lngColor = vbBlue
        Case Else
            '1-有效, 2-回收,3-退卡,4-失效,8-停用
            lngColor = &H80000008
        End Select
        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = lngColor
    End With
End Sub

Private Sub mfrmSquareCardConsume_zlDblClick(ByVal lng结算ID As Long, ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    If zlstr.IsHavePrivs(mstrPrivs, "卡结算消费明细帐") = False Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_INSIDE_1503_1", Me, "卡结算ID=" & lng结算ID, 1)
End Sub

Private Sub mfrmSquareCardInFull_AfterRowChange(ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    mTy_CurCard.bln允许充值回退 = mfrmSquareCardInFull.zl允许回退
End Sub

Private Sub mfrmSquareCardInFull_zlPopupMenus(ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    '弹出菜单:充值相关
    Call ShowPopuMenus(1)
End Sub

Private Function CardStopAndStart(ByVal blnStop As Boolean) As Boolean
    '功能:卡片停用或启用
    '入参:blnStop-停用卡片
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mTy_CurCard.lng卡ID <= 0 Then Exit Function
    
    If blnStop Then '停用
        If mTy_CurCard.byt卡状态 = m_CardStatus.Stoped Then Exit Function
        If MsgBox("你确定要对卡号为:“" & mTy_CurCard.str卡号 & "”的记录进行停用操作吗？" & vbCrLf & _
                    "   『是』:进行停用操作，停用后的卡片将不能进行刷卡消费，也不能再发出" & vbCrLf & _
                    "   『否』:放弃本次停用操作", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If mTy_CurCard.byt卡状态 <> m_CardStatus.Stoped Then Exit Function
        '补卡停用的卡不能启用
        strSQL = "Select 1 From 消费卡信息 Where ID = [1] And 当前状态 = 5"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mTy_CurCard.lng卡ID)
        If rsTemp.EOF = False Then
            ShowMsgbox "当前卡(卡号为:" & mTy_CurCard.str卡号 & ")为补卡停用的卡，不能启用！"
            Exit Function
        End If
        If MsgBox("你确定要对卡号为:“" & mTy_CurCard.str卡号 & "”的记录进行启用操作吗？" & vbCrLf & _
                    "   『是』:进行启用操作，启用后的卡片将能进行刷卡消费或回收后再发出" & vbCrLf & _
                    "   『否』:放弃本次启用用操作", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If

    'Zl_消费卡信息_Stopandstart
    strSQL = "Zl_消费卡信息_Stopandstart("
    '  Id_In     In 消费卡信息.Id%Type,
    strSQL = strSQL & "" & mTy_CurCard.lng卡ID & ","
    '  停用人_In In 消费卡信息.停用人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  停用_In Number:=0 --停用_In 0-启用,1-停用
    strSQL = strSQL & "" & IIf(blnStop, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    CardStopAndStart = True
    Exit Function
ErrHandler:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Function SaveBatchUpdateCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:批量更新卡片信息
    '返回:更新成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-05 12:02:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strFieldValue As String, blnDate As Boolean
    Dim strSQL As String, strIDs As String, lngRow As Long
    Dim blnTrain As Boolean, cllPro As New Collection
    Dim blnYes As Boolean, strTemp As String
    
    On Error GoTo ErrHandler
    With vsCardList
        Select Case .Col
        Case .ColIndex("有效期")
           strFields = "有效期": blnDate = True
           If IsNull(dtpValidDate.value) Then
                strFieldValue = "3000-01-01"
           Else
                strFieldValue = Format(dtpValidDate.value, "yyyy-MM-dd")
           End If
        Case .ColIndex("卡类型")
           strFields = "卡类型": blnDate = False
           If cboCardType.ListIndex < 0 Then
                ShowMsgbox "卡类型未选择，请选择卡类型！"
                Exit Function
           End If
           strFieldValue = zlstr.NeedName(cboCardType.Text)
        Case Else
           Exit Function
        End Select
        ShowMsgbox "你确定要批量修改当前所有卡的“" & strFields & "”的值吗？", True, blnYes
        If blnYes = False Then Exit Function
    End With
    
    strIDs = ""
    With vsCardList
        For lngRow = 1 To .Rows - 1
            strTemp = Val(.TextMatrix(lngRow, .ColIndex("ID")))
            If zlCommFun.ActualLen(strIDs & "," & strTemp) >= 3980 Then
                ' Zl_消费卡信息_Update_Batch
                strSQL = "Zl_消费卡信息_Update_Batch("
                '  Ids_In    Varchar2,
                strSQL = strSQL & "'" & Mid(strIDs, 2) & "',"
                '  字段_In   Varchar2,
                strSQL = strSQL & "'" & strFields & "',"
                '  字段值_In Varchar2
                strSQL = strSQL & "'" & strFieldValue & "')"
                AddArray cllPro, strSQL
                strIDs = ""
            End If
            If strTemp <> 0 And .RowHidden(lngRow) = False Then
                strIDs = strIDs & "," & strTemp
            End If
        Next
    End With
    If strIDs <> "" Then
        ' Zl_消费卡信息_Update_Batch
        strSQL = "Zl_消费卡信息_Update_Batch("
        '  Ids_In    Varchar2,
        strSQL = strSQL & "'" & Mid(strIDs, 2) & "',"
        '  字段_In   Varchar2,
        strSQL = strSQL & "'" & strFields & "',"
        '  字段值_In Varchar2
        strSQL = strSQL & "'" & strFieldValue & "')"
        AddArray cllPro, strSQL
        strIDs = ""
    End If
    If cllPro.count = 0 Then Exit Function
    
    blnTrain = True
    ExecuteProcedureArrAy cllPro, Me.Caption
    blnTrain = False
    
    SaveBatchUpdateCardInfor = True
    ShowMsgbox "修改成功！"
    
    '刷新数据
    Call LoadDataToRpt
    chkModify.value = Unchecked
    
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadCardTypeData()
    '刷新消费卡类别
    Dim intIndex  As Integer
    Dim cbrControl As CommandBarComboBox
    
    On Error GoTo errHand
    Set cbrControl = cbsThis(2).Controls.Find(xtpControlComboBox, conMenu_COMBOX_INTERFACE)
    If cbrControl Is Nothing Then Exit Sub
    
    cbrControl.Clear
    cbrControl.Category = ""
    
    Set grsStatic.rs消费卡接口 = Nothing
    Set mrs消费卡 = zlGet消费卡接口(, False)
    
    intIndex = 1
    With mrs消费卡
        Do While Not .EOF
            cbrControl.AddItem NVL(!编号) & "-" & NVL(!名称) & IIf(NVL(!启用) = 1, "", "(停用)")
            cbrControl.ItemData(intIndex) = Val(NVL(!编号))
            If Val(NVL(!编号)) = mTy_CurCardType.lng编号 Then
               cbrControl.ListIndex = intIndex
            End If
            intIndex = intIndex + 1
            .MoveNext
        Loop
    End With
    If intIndex > 1 And cbrControl.ListIndex <= 0 Then cbrControl.ListIndex = 1
    
    If cbrControl.ListIndex > 0 Then
        Call SetCurrentCardType(cbrControl.ItemData(cbrControl.ListIndex))
        cbrControl.Category = cbrControl.ItemData(cbrControl.ListIndex)
    Else
        Call SetCurrentCardType(0)
    End If
    
    Call LoadDataToRpt
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function DeleteCardType() As Boolean
    '删除消费卡类别
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHand
    If mTy_CurCardType.lng编号 <= 0 Then Exit Function
    
    mrs消费卡.Filter = "编号 = " & mTy_CurCardType.lng编号
    If Val(NVL(mrs消费卡!系统)) = 1 Then
        ShowMsgbox "系统固定项，不能删除，请检查！"
        Exit Function
    End If
    
    '检查是否存在发卡记录
    strSQL = "Select 1 From 消费卡信息 Where 接口编号=[1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mTy_CurCardType.lng编号)
    If Not rsTemp.EOF Then
        ShowMsgbox "名称为“" & NVL(mrs消费卡!名称) & "”的消费卡存在发卡记录，不能删除！"
        Exit Function
    End If
    
    If MsgBox("你确认要删除名称为“" & NVL(mrs消费卡!名称) & "”的消费卡类别吗？", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
       
    'Zl_消费卡类别目录_Delete(
    strSQL = "Zl_消费卡类别目录_Delete("
    ' 编号_In In 消费卡类别目录.编号%Type
    strSQL = strSQL & "" & mTy_CurCardType.lng编号 & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    DeleteCardType = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetModiyCaption()
    With vsCardList
        Select Case .Col
        Case .ColIndex("有效期")
            chkModify.Caption = "修改“有效期”"
        Case .ColIndex("卡类型")
            chkModify.Caption = "修改“卡类型”"
        Case Else
            chkModify.Visible = False
        End Select
    End With
End Sub

Private Sub SetModifyEnabled()
    Dim blnEnabled As Boolean
    
    blnEnabled = (chkModify.value = vbChecked)
    With vsCardList
        cboCardType.Visible = False
        dtpValidDate.Visible = False
        chkModify.Visible = True
        cmdModify.Visible = blnEnabled
        Select Case .Col
        Case .ColIndex("有效期")
            dtpValidDate.Visible = blnEnabled
            picModify.Visible = True
        Case .ColIndex("卡类型")
            cboCardType.Visible = blnEnabled
            picModify.Visible = True
        Case Else
            picModify.Visible = False
        End Select
    End With
End Sub

Private Sub SetModifyDefaultValue()
    With vsCardList
        If .Row <= 0 Then Exit Sub
        
        Select Case .Col
        Case .ColIndex("有效期")
            If Trim(.TextMatrix(.Row, .Col)) = "" Then
                 dtpValidDate.value = Null
            Else
                If CDate(.TextMatrix(.Row, .Col)) < dtpValidDate.MinDate Then
                    dtpValidDate.value = dtpValidDate.MinDate
                Else
                    dtpValidDate.value = CDate(.TextMatrix(.Row, .Col))
                End If
            End If
        Case .ColIndex("卡类型")
            cbo.SeekIndex cboCardType, .TextMatrix(.Row, .Col)
        End Select
    End With
End Sub

Private Sub SetCurrentCard()
    '功能:设置当前选择卡信息
    Dim ty_Temp As Ty_CurrentCard
    
    On Error GoTo ErrHandler
    mTy_CurCard = ty_Temp '自定义Type初始化
    
    With vsCardList
        If .Rows < 2 Then Exit Sub
        If .Row < 1 Then Exit Sub
    
        mTy_CurCard.blnHaveData = .TextMatrix(1, .ColIndex("卡号")) <> ""
        mTy_CurCard.lng卡ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        mTy_CurCard.byt卡状态 = Val(.Cell(flexcpData, .Row, .ColIndex("当前状态")))
        mTy_CurCard.str卡号 = .TextMatrix(.Row, .ColIndex("卡号"))
        mTy_CurCard.str卡类型 = .TextMatrix(.Row, .ColIndex("卡类型"))
        mTy_CurCard.bln已消费 = .TextMatrix(.Row, .ColIndex("已消费")) = "√"
        mTy_CurCard.str发卡人 = .TextMatrix(.Row, .ColIndex("发卡人"))
        mTy_CurCard.str发卡日期 = .TextMatrix(.Row, .ColIndex("发卡时间"))
        mTy_CurCard.bln充值卡 = .TextMatrix(.Row, .ColIndex("充值卡")) = "√"
        mTy_CurCard.lng发卡序号 = Val(.TextMatrix(.Row, .ColIndex("发卡序号")))
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsCardList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error GoTo ErrHandler
    If mblnPrinting Then Exit Sub
    If OldRow <> NewRow Then
        zl_VsGridRowChange vsCardList, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
        
        zlCommFun.ShowFlash "正在装载数据,请稍候..."
        '设置行的信息
        Call SetCurrentCard
        
        With mTy_CurCard
            Call mfrmSquareCardCallBack.zlReLoadData(mTy_CurCardType.lng编号, .lng卡ID, .str卡类型, .str卡号)    '回收记录
            Call mfrmSquareCardInFull.zlReLoadData(mTy_CurCardType.lng编号, .lng卡ID, .str卡类型, .str卡号)     '充值记录
            Call mfrmSquareCardConsume.zlReLoadData(mTy_CurCardType.lng编号, .lng卡ID, .str卡类型, .str卡号)      '消费记录
        End With
        zlCommFun.StopFlash
    End If
    
    If mTy_CurCard.blnHaveData Then
        If zlstr.IsHavePrivs(mstrPrivs, "修改卡信息") Then
            If OldCol <> NewCol Then chkModify.value = vbUnchecked
            Call SetModiyCaption
            Call SetModifyEnabled
            Call SetModifyDefaultValue
        End If
    Else
        picModify.Visible = False
    End If
    Exit Sub
ErrHandler:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsCardList_DblClick()
    On Error GoTo ErrHandler
    '双击查看
    If vsCardList.MouseRow <= 0 Then Exit Sub
    If mTy_CurCard.lng卡ID = 0 Then Exit Sub
    frmSquareSendCard.zlShowCard Me, mlngModule, mstrPrivs, gEd_查询, mTy_CurCardType.lng编号, mTy_CurCard.lng卡ID
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsCardList_GotFocus()
    zl_VsGridGotFocus vsCardList, gSysColor.lngGridColorSel
End Sub

Private Sub vsCardList_LostFocus()
    zl_VsGridLostFocus vsCardList, gSysColor.lngGridColorLost
End Sub

Private Sub vsCardList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlstr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Sub vsCardList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlstr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long, vRect As RECT
    
    vRect = zlControl.GetControlRect(picImgList.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList.Height
    
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCardList, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlstr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
 
Private Sub picCardList_Resize()
    Err = 0: On Error Resume Next
    With picCardList
        vsCardList.Left = .ScaleLeft
        vsCardList.Width = .ScaleWidth
        vsCardList.Height = .ScaleHeight - vsCardList.Top
        picModify.Width = .ScaleWidth - picModify.Left - 50
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

Private Sub picModify_Click()
    Err = 0: On Error Resume Next
    With picModify
        cmdModify.Left = .ScaleWidth - cmdModify.Width - 50
        cboCardType.Width = cmdModify.Left - cboCardType.Left
        dtpValidDate.Width = cboCardType.Width
    End With
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
    Dim lngPreRow As Long, lngPreCol As Long
    
    blnCardList = Me.ActiveControl Is vsCardList
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "消费卡清册"
    
    If CStr(mArrFilter("发卡时间")(0)) <> "1901-01-01" Then
        objRow.Add "发卡时间：" & CStr(mArrFilter("发卡时间")(0)) & "至" & CStr(mArrFilter("发卡时间")(1))
    End If
    If CStr(mArrFilter("回收时间")(0)) <> "1901-01-01" Then
        objRow.Add "回收时间：" & CStr(mArrFilter("回收时间")(0)) & "至" & CStr(mArrFilter("回收时间")(1))
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
    
    If mArrFilter("领卡人") <> "" Then objRow.Add "领卡人：" & mArrFilter("领卡人")
    If mArrFilter("发卡人") <> "" Then objRow.Add "发卡人：" & mArrFilter("发卡人")
    If mArrFilter("卡类型") <> "" Then objRow.Add "卡类型：" & mArrFilter("卡类型")
    If Val(mArrFilter("包含停用卡")) = 1 Then objRow.Add "包含停用卡"
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    mblnPrinting = True
    '由于打印控件不能识别列隐藏属性
    With vsCardList
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            If .ColHidden(i) Or i = 0 Then
                .Cell(flexcpData, 0, i) = .ColWidth(i)
                .ColWidth(i) = 0
            End If
        Next
        lngPreRow = .Row: lngPreCol = .Col
    End With
    
    Err = 0: On Error GoTo errHand:
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
            If .ColHidden(i) Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
                .Cell(flexcpData, 0, i) = ""
            End If
        Next
        .GridColor = &H8000000F
        .Row = lngPreRow: .Col = lngPreCol
        .Redraw = flexRDBuffered
    End With
    mblnPrinting = False
    Exit Sub
errHand:
    mblnPrinting = False
    vsCardList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PrintReBill()
    '功能:重打票据
    Dim strTemp As String
    Dim blnYes As Boolean
    Dim lngID As Long
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If strTemp = "发卡" Then
        If mTy_CurCard.lng卡ID <= 0 Then
            ShowMsgbox "未选中有效的消费卡记录！": Exit Sub
        End If
    Else
        lngID = mfrmSquareCardInFull.zlGet充值ID
        If lngID <= 0 Then
            ShowMsgbox "未选中有效的充值记录！":  Exit Sub
        End If
    End If
    
    If mTy_CurCard.bln充值卡 = False Then
        ShowMsgbox "你确定要打印缴款单吗？", True, blnYes
        If blnYes = False Then Exit Sub
        strTemp = "发卡"
    Else
        strTemp = zlCommFun.ShowMsgbox("缴款单打印", "请选择你要打印的缴款单", "发卡(&F),充值(&I),取消(&C)", Me, vbDefaultButton2)
    End If
    If strTemp = "取消" Or strTemp = "" Then Exit Sub
    
    If strTemp = "发卡" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, _
            "付款序号=" & mTy_CurCard.lng发卡序号, "缴款=" & 0, "找补=" & 0, "充值ID=0", "ReportFormat=1", 2)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, _
            "充值ID=" & lngID, "缴款=" & 0, "找补=" & 0, "付款序号=0", "ReportFormat=2", 2)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlOpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '功能:打开指定报表
    '入参:lngSys-系统号
    '     strReportCode报表编号
    With mTy_CurCard
        If .lng卡ID = 0 Then Exit Sub
        
        Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, _
            "消费卡ID=" & .lng卡ID, "卡号=" & .str卡号, "卡类型=" & .str卡类型, "发卡人=" & .str发卡人, "发卡日期=" & .str发卡日期)
    End With
End Sub

Private Sub vsCardList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '右键弹出菜单
    If Button <> vbRightButton Then Exit Sub
    Call ShowPopuMenus(0)
End Sub

Private Function ShowPopuMenus(ByVal bytMode As Byte) As Boolean
    '显示弹出菜单
    '入参：
    '   bytMode 0-卡面信息列表，1-充值信息列表
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrControl As CommandBarControl
    
    Err = 0: On Error Resume Next
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls(2)
    If cbrMenuBar.Visible = False Then Exit Function
    
    Set cbrPopupBar = cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Select Case cbrControl.id
        Case conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete
            '不显示菜单
        Case Else
            If bytMode = 0 Or _
                (bytMode = 1 And (cbrControl.id = conMenu_Edit_CardInFull Or cbrControl.id = conMenu_Edit_CardInFullBack)) Then
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.id, cbrControl.Caption)
                cbrPopupItem.BeginGroup = cbrControl.BeginGroup
            End If
        End Select
    Next
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls(3)
    If cbrMenuBar.Visible Then
        For Each cbrControl In cbrMenuBar.CommandBar.Controls
            Select Case cbrControl.id
            Case conMenu_View_Refresh
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.id, cbrControl.Caption)
                cbrPopupItem.BeginGroup = cbrControl.BeginGroup
            End Select
        Next
    End If
    
    cbrPopupBar.ShowPopup
End Function
