VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "CODEJOCK.CALENDAR.V16.3.1.OCX"
Begin VB.Form frmSchManage 
   Caption         =   "检查预约管理"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   Icon            =   "frmSchManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11715
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picTimeTable 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   5040
      ScaleHeight     =   4935
      ScaleWidth      =   6255
      TabIndex        =   1
      Top             =   600
      Width           =   6255
      Begin TabDlg.SSTab sstTimeTable 
         Height          =   4575
         Left            =   600
         TabIndex        =   8
         Top             =   720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   8070
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "预约时间表"
         TabPicture(0)   =   "frmSchManage.frx":0442
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "schTimeTable"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "预约列表"
         TabPicture(1)   =   "frmSchManage.frx":045E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "vsfSchList"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin zl9PACSWork.ucScheduleTimetable schTimeTable 
            Height          =   4935
            Left            =   -74760
            TabIndex        =   9
            Top             =   600
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   8705
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfSchList 
            Height          =   3735
            Left            =   360
            TabIndex        =   10
            Top             =   600
            Width           =   4695
            _cx             =   8281
            _cy             =   6588
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
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
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
      End
   End
   Begin VB.PictureBox picCalendar 
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   120
      ScaleHeight     =   7335
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VSFlex8Ctl.VSFlexGrid vsfSchDevice 
         Height          =   2070
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   2895
         _cx             =   5106
         _cy             =   3651
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
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
      Begin XtremeCalendarControl.DatePicker dpCalendar 
         Height          =   2655
         Left            =   720
         TabIndex        =   5
         Top             =   2040
         Width           =   3135
         _Version        =   1048579
         _ExtentX        =   5530
         _ExtentY        =   4683
         _StockProps     =   64
         AutoSize        =   0   'False
         ShowTodayButton =   0   'False
         ShowNoneButton  =   0   'False
         ShowNonMonthDays=   0   'False
         Show3DBorder    =   2
         AskDayMetrics   =   -1  'True
         TextTodayButton =   "返回今天"
      End
      Begin VB.Frame frmChangeDay 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   4560
         Width           =   3135
         Begin VB.CommandButton cmdChangeDay 
            Caption         =   "前一周"
            Height          =   495
            Index           =   1
            Left            =   10
            TabIndex        =   11
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdChangeDay 
            Caption         =   "后一周"
            Height          =   495
            Index           =   2
            Left            =   2250
            TabIndex        =   4
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdToday 
            Caption         =   "今天"
            Height          =   495
            Left            =   870
            TabIndex        =   3
            Top             =   120
            Width           =   1380
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfWeekView 
         Height          =   1095
         Left            =   240
         TabIndex        =   12
         Top             =   5760
         Width           =   3255
         _cx             =   5741
         _cy             =   1931
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   7575
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4154
            MinWidth        =   4154
            Picture         =   "frmSchManage.frx":047A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13785
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   3480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmSchManage.frx":0D0E
      Left            =   4080
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSchManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngOrderID As Long             '医嘱ID
Private mlngSchDeviceID As Long         '预约设备ID
Private mstrDeptIDs As String           '科室ID
Private mschDate As Date                '当前日期
Private mfrmParent As Object            '父窗口
Private mstrModifiedOrderID As String   '保存过预约信息的医嘱ID串，用“,”连接
Private mstrSchRestDate As String       '当月休息日
Private mstrPrivs As String             '调用者的权限

Private mlngColorLblWaiting As Long     '预约标签，预约等候颜色
Private mlngColorLblDone As Long        '预约标签，完成颜色
Private mlngColorLblPassed As Long      '预约标签，过号颜色
Private mblnAutoPrint As Boolean        '是否自动打印预约单

'检查预约设备
Private Enum constScheduleDeviceList
    col_SchDevice_ID = 0
    col_SchDevice_影像类别 = 1
    col_SchDevice_设备名称 = 2
    col_SchDevice_设备说明 = 3
End Enum

'检查预约列表
Private Enum constScheduleList
    col_SchList_ID = 0
    col_SchList_序号 = 1
    col_SchList_姓名 = 2
    col_SchList_医嘱内容 = 3
    col_SchList_诊室名称 = 4
    col_SchList_预约开始时间 = 5
    col_SchList_预约结束时间 = 6
    col_SchList_执行过程 = 7
    col_SchList_手机号 = 8
    col_SchList_检查注意 = 9
End Enum

'检查预约周视图
Private Enum constWeekView
    col_WeekView_星期 = 0
    col_WeekView_空余 = 1
    col_WeekView_已预约 = 2
    col_WeekView_总容量 = 3
End Enum

Private Sub InitCommandBar()
'------------------------------------------------
'功能：初始化工具栏
'参数： 无
'返回： 无
'------------------------------------------------
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    
    On Error GoTo err
    
    '这部分全局设置，是否必要？
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbrMain.VisualTheme = xtpThemeOffice2003
    Set cbrMain.Icons = zlCommFun.GetPubIcons
        
    With cbrMain.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True         '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    
    cbrMain.EnableCustomization False
    cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '不显示菜单
    cbrMain.ActiveMenuBar.Visible = False
    
    '显示工具栏
    Set cbrToolBar = cbrMain.Add("预约工具栏", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Save, "保存预约")
        cbrControl.iconid = 6823
        cbrControl.ToolTipText = "保存预约信息"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Print, "打印预约单")
        cbrControl.iconid = 103
        cbrControl.ToolTipText = "打印患者的预约通知单"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Modify, "修改预约")
        cbrControl.iconid = 6886
        cbrControl.ToolTipText = "修改检查预约"
        
        
'        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Query, "预约查询")
'        cbrControl.IconId = 3946
'        cbrControl.ToolTipText = "查询预约情况"
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Delete, "删除预约")
        cbrControl.iconid = 6822
        cbrControl.ToolTipText = "删除一个检查预约"
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Refresh, "刷新")
        cbrControl.iconid = 791
        cbrControl.ToolTipText = "刷新数据"
        
        cbrControl.BeginGroup = True
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Quit, "退出")
        cbrControl.iconid = 191
        cbrControl.ToolTipText = "关闭窗口"
        
        cbrControl.BeginGroup = True
    End With
    
    cbrToolBar.Position = xtpBarTop
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_PacsSchdule_Modify     '修改预约
            '打开检查预约窗口
            Call ModifySchedule(mlngOrderID)
            
        Case conMenu_PacsSchdule_Delete     '删除预约
            Call DelSchedule(mlngOrderID)
            
        Case conMenu_PacsSchdule_Refresh    '刷新
            Call RefreshSchedule
            
'        Case conMenu_PacsSchdule_Query      '查询
'            Call frmSchQuery.zlShowMe(mstrDeptIDs, Me)
            
        Case conMenu_PacsSchdule_Print      '打印预约单
            Call PrintSchedule(mlngOrderID)
            
        Case conMenu_PacsSchdule_Save       '保存预约
            Call SaveSchedule
            Call RefreshSchedule
            
        Case conMenu_PacsSchdule_Quit       '退出
            Unload Me
            
    End Select
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_PacsSchdule_Refresh, conMenu_PacsSchdule_Quit   '刷新 ,退出
            '什么都不做
        Case Else
            Control.Enabled = IIf(sstTimeTable.Tab = 1, False, True)
    End Select
End Sub

Private Sub cmdChangeDay_Click(Index As Integer)
    Select Case Index
        Case 1
            mschDate = mschDate - 7
        Case 2
            mschDate = mschDate + 7
    End Select
    
    Call ChangeCalendar(mschDate)
    Call RefreshSchedule
End Sub

Private Sub cmdToday_Click()
    mschDate = Format(Now, "YYYY-MM-DD")
    Call ChangeCalendar(mschDate)
    Call RefreshSchedule
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picCalendar.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picTimeTable.hwnd
    End If
End Sub

Private Sub dpCalendar_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If InStr(mstrSchRestDate, Format(Day, "YYYY-MM-DD")) > 0 Then
        Metrics.ForeColor = vbRed
        Metrics.Font.Bold = True
    End If
End Sub

Private Sub dpCalendar_MonthChanged()
    Call RefreshCalendar
End Sub

Private Sub dpCalendar_SelectionChanged()
    '更换了日期，重新刷新时间表
    
    mschDate = dpCalendar.Selection.Blocks(0).DateBegin
    ChangeCalendar (mschDate)
    Call RefreshSchedule
End Sub

Private Sub InitFaceScheme()
'------------------------------------------------
'功能：初始化界面布局
'参数： 无
'返回： 无
'------------------------------------------------
    Dim Pane1 As Pane, Pane2 As Pane
    
    On Error GoTo err
    
    '设置总体显示策略
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .SetCommandBars cbrMain
        .options.HideClient = True
        .options.UseSplitterTracker = False '实时拖动
        .options.ThemedFloatingFrames = True
        .options.AlphaDockingContext = True
        dkpMain.options.DefaultPaneOptions = PaneNoCaption
    End With
    
    '先从注册表读取预先设置好的窗口布局，然后再逐个设置
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    
    '如果注册表中保存的界面布局Pane数量不对，则加载默认的Pane设置
    If dkpMain.PanesCount <> 2 Then
        dkpMain.DestroyAll
        
        Set Pane1 = dkpMain.CreatePane(1, 350, 150, DockLeftOf, Nothing)
        Pane1.title = "预约信息"
        Pane1.options = PaneNoCaption
        
        Set Pane2 = dkpMain.CreatePane(2, 650, 300, DockRightOf, Pane1)
        Pane2.title = "预约时间表"
        Pane2.options = PaneNoCaption
    End If
    
    '默认显示时间表
    sstTimeTable.Tab = 0
    vsfSchList.Visible = False
    
    mlngColorLblWaiting = zlDatabase.GetPara("检查预约标签已预约颜色", glngSys, 1292, "0")
    mlngColorLblDone = zlDatabase.GetPara("检查预约标签已完成颜色", glngSys, 1292, "12632256")
    mlngColorLblPassed = zlDatabase.GetPara("检查预约标签已过号颜色", glngSys, 1292, "255")
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim strOrders() As String
    
    '合并删除和保存的医嘱ID
    strOrders = Split(schTimeTable.strModifiedOrderID, ",")
    For i = 0 To UBound(strOrders)
        If InStr(mstrModifiedOrderID, CStr(strOrders(i))) = 0 Then
            mstrModifiedOrderID = mstrModifiedOrderID & "," & CStr(strOrders(i))
        End If
    Next i
    
    If Trim(mstrModifiedOrderID) <> "" Then
        mstrModifiedOrderID = Mid(mstrModifiedOrderID, 2)
    End If
    
    '关闭窗体的时候，保存界面布局
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    
    Call SaveWinState(Me, App.ProductName)
    
    Set mfrmParent = Nothing
    
    '释放DockingPane
    For i = 1 To dkpMain.PanesCount
        dkpMain.Panes(i).Handle = 0
    Next i
    dkpMain.CloseAll
End Sub

Private Sub picCalendar_Resize()
    On Error Resume Next
    
    vsfSchDevice.Left = 0
    vsfSchDevice.Top = 0
    vsfSchDevice.Width = picCalendar.ScaleWidth
    
    dpCalendar.Left = 0
    dpCalendar.Top = vsfSchDevice.Height
    dpCalendar.Width = picCalendar.ScaleWidth
    
    frmChangeDay.Left = 0
    frmChangeDay.Top = dpCalendar.Top + dpCalendar.Height - 50
    frmChangeDay.Width = picCalendar.ScaleWidth
    
    cmdToday.Width = frmChangeDay.Width - cmdChangeDay(2).Width * 2 - 20
    
    cmdChangeDay(2).Left = frmChangeDay.Width - cmdChangeDay(2).Width - 10
    
    vsfWeekView.Left = 0
    vsfWeekView.Top = frmChangeDay.Top + frmChangeDay.Height + 40
    vsfWeekView.Width = picCalendar.ScaleWidth
    vsfWeekView.Height = picCalendar.ScaleHeight - vsfWeekView.Top - 50
    
End Sub

Private Sub picTimeTable_Resize()
    On Error Resume Next
    
    sstTimeTable.Left = 0
    sstTimeTable.Top = 0
    sstTimeTable.Width = picTimeTable.ScaleWidth
    sstTimeTable.Height = picTimeTable.ScaleHeight - stbThis.Height
    
    schTimeTable.Left = 0
    schTimeTable.Top = 300
    schTimeTable.Width = sstTimeTable.Width - 20
    schTimeTable.Height = sstTimeTable.Height - 300
'
    vsfSchList.Left = 0
    vsfSchList.Top = 300
    vsfSchList.Width = sstTimeTable.Width - 20
    vsfSchList.Height = sstTimeTable.Height - 300
End Sub

Private Sub ChangeCalendar(dtDate As Date)
'------------------------------------------------
'功能：修改预约日历的日期
'参数：dtDate -- 日历的日期
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    dpCalendar.ClearSelection
    Call dpCalendar.Select(dtDate)
    dpCalendar.EnsureVisibleSelection
    If dpCalendar.Visible = True Then
        dpCalendar.SetFocus
    End If
    
    Call LoadWeekView
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function ZlShowMe(strPrivs As String, strDeptIDs As String, lngOrderID As Long, frmParent As Object) As String
'------------------------------------------------
'功能：打开窗口
'参数： strDeptIDs -- 科室ID串
'       lngOrderID -- 医嘱ID，控制预约管理显示的日期
'       frmParent -- 父窗体
'       strPrivs -- 调用者的权限
'返回：保存过预约信息的医嘱ID串，用“,”连接
'------------------------------------------------
    On Error GoTo err
    
    mlngOrderID = 0
    mlngSchDeviceID = 0
    mstrPrivs = strPrivs
    
    mstrDeptIDs = strDeptIDs
    Set mfrmParent = frmParent
    
    mstrModifiedOrderID = ""
    
    '初始化界面布局
    Call InitFaceScheme
    
    '创建工具栏
    Call InitCommandBar
    
    '读取系统参数
    mblnAutoPrint = IIf(Val(zlDatabase.GetPara("保存预约后自动打印预约单", glngSys, 1292)) = 1, True, False)
    
    '先初始化时间表控件
    Call schTimeTable.Init(2)   '预约管理模式
    
    Call RestoreWinState(Me, App.ProductName)
    
    '设置日历参数
    dpCalendar.ShowNonMonthDays = False
    dpCalendar.AskDayMetrics = True
    
    mschDate = GetOrderSchDate(lngOrderID)
    
    If LoadData = False Then
        Exit Function
    End If
    
    Me.Show 1, mfrmParent
    
    ZlShowMe = mstrModifiedOrderID
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Function GetOrderSchDate(lngOrderID As Long) As Date
'------------------------------------------------
'功能：根据医嘱ID，获取这条医嘱的预约时间，如果没有预约，返回今天
'参数：
'       lngOrderID -- 医嘱ID，控制预约管理显示的日期
'返回：预约管理显示的日期
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strSQL = "select 预约开始时间 from 影像预约记录 where 医嘱ID = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询预约日期", lngOrderID)
    
    If rsTemp.EOF = False Then
        GetOrderSchDate = Format(rsTemp!预约开始时间, "YYYY-MM-DD")
    Else
        GetOrderSchDate = Format(Now, "YYYY-MM-DD")
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function


Private Function LoadData() As Boolean
'------------------------------------------------
'功能：加载窗体的所有数据
'参数：
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    On Error GoTo err
    
    '有先后顺序
    If LoadSchDevice = False Then
        Exit Function
    End If
    
    '设置日历
    Call ChangeCalendar(mschDate)
    Call RefreshCalendar
    
    '刷新时间表
    Call RefreshSchedule
    
    LoadData = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Function LoadSchDevice() As Boolean
'------------------------------------------------
'功能：加载预约设备
'参数：
'返回：True -- 成功 ；False -- 失败
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    strSQL = "select ID,设备名称,影像设备号,影像类别,设备说明 from 影像预约设备 where 科室ID in (" & mstrDeptIDs & ") and 是否启用=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询预约设备")
    
    With vsfSchDevice
        .Clear
        .Cols = 4
        .Rows = rsTemp.RecordCount + 1
        .FixedRows = 1
        .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarVertical
        .CellAlignment = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .ExtendLastCol = True
        .RowHeightMin = 400
        
        .ColWidth(col_SchDevice_ID) = 50
        .ColWidth(col_SchDevice_设备名称) = 2000
        .ColWidth(col_SchDevice_影像类别) = 5000
        .ColWidth(col_SchDevice_设备说明) = 500
        
        '合并第一行
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        For i = 0 To 3
            .TextMatrix(0, i) = "预约设备"
        Next i
        
        '从数据库加载数据
        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, col_SchDevice_ID) = rsTemp!ID
            .TextMatrix(i, col_SchDevice_设备名称) = rsTemp!设备名称
            .TextMatrix(i, col_SchDevice_影像类别) = rsTemp!影像类别
            .TextMatrix(i, col_SchDevice_设备说明) = NVL(rsTemp!设备说明)
            rsTemp.MoveNext
        Next i
        
        '隐藏后台数据列
        .ColHidden(col_SchDevice_ID) = True
        .ColHidden(col_SchDevice_影像类别) = True
        
        '选择第一行
        If .Rows > 1 Then
            Call .Select(1, 1)
            mlngSchDeviceID = Val(.TextMatrix(1, col_SchDevice_ID))
        Else
            mlngSchDeviceID = 0
            Call MsgBoxD(Me, "没有可用于预约的影像设备，请先添加预约设备。", vbOKOnly, "检查预约提示")
            Exit Function
        End If
    End With
    
    LoadSchDevice = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function RefreshSchedule() As Boolean
'------------------------------------------------
'功能：加载刷新时间表内容
'参数：
'返回：True -- 成功 ； False -- 失败
'------------------------------------------------
    On Error GoTo err
    
    '刷新预约列表
    Call LoadSchList
    
    '已经存在预约信息，直接显示即可
    If schTimeTable.RefreshSchedule(mlngSchDeviceID, mschDate, mlngOrderID) = True Then
        mlngOrderID = schTimeTable.LabelOrderID
    Else
        mlngOrderID = 0
    End If
        
    RefreshSchedule = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub DelSchedule(lngOrderID As Long)
'------------------------------------------------
'功能：删除预约
'参数： lngOrderID -- 医嘱ID
'返回：无
'------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo err
    
    If lngOrderID = 0 Then
        Exit Sub
    End If
    
    If InStr(mstrModifiedOrderID, CStr(lngOrderID)) = 0 Then
        mstrModifiedOrderID = mstrModifiedOrderID & "," & CStr(lngOrderID)
    End If
    
    strSQL = "Zl_影像预约记录_删除(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure strSQL, "删除检查预约"
    
    Call RefreshSchedule
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub schTimeTable_OnChangeOrder(ByVal lngOrderID As Long, ByVal strOrderInfo As String)
    mlngOrderID = lngOrderID
    stbThis.Panels(2).Text = strOrderInfo
End Sub

Private Sub schTimeTable_OnMenuScheduleModify()
    '打开检查预约窗口
    Call ModifySchedule(mlngOrderID)
End Sub

Private Sub schTimeTable_OnMenuSchedulePrint()
    Call PrintSchedule(mlngOrderID)
End Sub

Private Sub schTimeTable_OnSchLabelModifed(ByVal iIndex As Integer)
    stbThis.Panels(2).Text = schTimeTable.LabelOrderInfo
End Sub

Private Sub sstTimeTable_Click(PreviousTab As Integer)
    '切换了页面
    If PreviousTab <> sstTimeTable.Tab Then
        If sstTimeTable.Tab = 0 Then
            schTimeTable.Visible = True
            vsfSchList.Visible = False
        Else
            '显示预约列表
            schTimeTable.Visible = False
            vsfSchList.Visible = True
            Call LoadSchList
        End If
    End If
    
End Sub

Private Sub vsfSchDevice_Click()
    If vsfSchDevice.Rows > 1 Then
        '修改当前被选中的预约设备ID
        mlngSchDeviceID = vsfSchDevice.TextMatrix(vsfSchDevice.RowSel, col_SchDevice_ID)
        Call RefreshSchedule
        Call RefreshCalendar
    End If
End Sub

Private Sub SaveSchedule()
'------------------------------------------------
'功能：保存预约
'参数：
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim arrOrderID() As String
    
    On Error GoTo err
        
    Call schTimeTable.SaveAllSchedule
    
    '自动打印预约单
    If mblnAutoPrint = True Then
        arrOrderID = Split(schTimeTable.strModifiedOrderID, ",")
        For i = 0 To UBound(arrOrderID)
            Call PrintSchedule(arrOrderID(i))
        Next i
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
        
End Sub

Private Sub PrintSchedule(ByVal lngOrderID As Long)
'------------------------------------------------
'功能：打印当前预约单
'参数： lngOrderID -- 医嘱ID
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsReports As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnPrinted As Boolean
    Dim lngUniFmt As Long           '通用报表格式序号
    
    On Error GoTo err
    
    '打印预约单
    If lngOrderID <> 0 Then
        '首先检查报表是否只有一个格式
        strSQL = "Select a.ID,a.编号,b.序号,b.说明 From zlreports a,zlrptfmts b Where a.Id=b.报表ID And a.编号=[1] Order By 序号"
        Set rsReports = zlDatabase.OpenSQLRecord(strSQL, "查询预约单报表格式", "ZL1_INSIDE_1290_01")

        If rsReports.EOF = True Then
            Call MsgBox("报表“ZL1_INSIDE_1290_01”不存在，请联系管理员添加此报表。", vbInformation, "检查预约提示")
            Exit Sub
        End If
        '如果有多个格式，按照诊疗项目ID，查找对应的报表格式名称
        If rsReports.RecordCount > 1 Then
            strSQL = "Select a.名称 From 病历文件列表 A, 病历单据应用 B, 病人医嘱记录 C " _
                & " Where c.诊疗项目id = b.诊疗项目id And decode(c.病人来源, 3, 1, c.病人来源) = b.应用场合 " _
                & "And b.病历文件id = a.ID And c.ID = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询病历文件名称", lngOrderID)
            
            If rsTemp.EOF = False Then
            While rsReports.EOF = False And blnPrinted = False
                If NVL(rsReports!说明) = "通用检查预约单" Then
                    lngUniFmt = rsReports!序号
                End If
                
                If NVL(rsReports!说明) = NVL(rsTemp!名称) Then
                    If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "医嘱ID=" & lngOrderID, "ReportFormat=" & rsReports!序号, 2) = False Then
                        Call MsgBox("报表“ZL1_INSIDE_1290_01”中，格式为：" & NVL(rsReports!说明) & "的报表，打开不成功，请联系管理员修正此报表。", vbInformation, "检查预约提示")
                    Else
                        '打印完退出循环
                        blnPrinted = True
                    End If
                Else
                    rsReports.MoveNext
                End If
            Wend
            End If
            '如果没有，则查找“通用检查预约单”报表来打印
            If blnPrinted = False Then
                If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "医嘱ID=" & lngOrderID, "ReportFormat=" & lngUniFmt, 2) = False Then
                    Call MsgBox("报表“ZL1_INSIDE_1290_01”中，格式为：“通用检查预约单”的报表，打开不成功，请联系管理员修正此报表。", vbInformation, "检查预约提示")
                Else
                    blnPrinted = True
                End If
            End If
        Else
            If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "医嘱ID=" & lngOrderID, 2) = False Then
                Call MsgBox("报表“ZL1_INSIDE_1290_01”打开不成功，请联系管理员修正此报表。", vbInformation, "检查预约提示")
            Else
                blnPrinted = True
            End If
        End If
        
        '写入打印记录
        strSQL = "Zl_影像预约记录_打印(" & lngOrderID & ")"
        zlDatabase.ExecuteProcedure strSQL, "检查预约单打印"
        
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifySchedule(lngOrderID As Long)
'------------------------------------------------
'功能：修改预约
'参数：lngOrderID -- 医嘱ID
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    If schTimeTable.LabelOrderID <> 0 Then
        strSQL = "select 执行过程 from 病人医嘱发送 where 医嘱ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询执行过程", lngOrderID)
        
        If rsTemp.EOF = False Then
            If NVL(rsTemp!执行过程, 0) = 0 Or NVL(rsTemp!执行过程, 0) = 1 Then
                Call frmSchSchedule.ZlShowMe(mstrPrivs, lngOrderID, mstrDeptIDs, Me)
                Call RefreshSchedule
            Else
                MsgBox "本次预约已经被执行，不能修改。", vbOKOnly, "检查预约提示"
            End If
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadSchList()
'------------------------------------------------
'功能：加载预约记录列表
'参数：
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim i As Integer
    Dim lngColor As Long
    
    On Error GoTo err
    
    strSQL = "Select d.ID, d.医嘱ID, d.序号, d.诊室名称, d.预约开始时间, d.预约结束时间, " _
        & " d.预约开始时间段, d.预约结束时间段, b.姓名, b.医嘱内容, b.婴儿, c.执行过程,d.检查注意,e.手机号 " _
        & " From 病人医嘱记录 B, 病人医嘱发送 C,影像预约记录 D,病人信息 E Where b.id in " _
        & " (Select  a.医嘱ID From 影像预约记录 A Where a.预约设备ID = [1] And " _
        & " a.预约开始时间 Between [2] And [3] )And c.医嘱id = b.id and d.医嘱id=B.id  And b.病人id = e.病人id Order By cast(d.序号 as int)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询今天的预约记录", mlngSchDeviceID, CDate(Format(mschDate, "yyyy-MM-dd 00:00:00")), CDate(Format(mschDate, "yyyy-MM-dd 23:59:59")))
    
    With vsfSchList
        .Clear
        .Cols = 10
        .Rows = IIf(rsTemp.EOF = True, 1, rsTemp.RecordCount + 1)
        .FixedRows = 1
        .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExSort
        .Cell(flexcpAlignment, 0, 0, 0, 1) = flexAlignCenterCenter
        .ExtendLastCol = True
        .RowHeightMin = 350
        
        .ColWidth(col_SchList_ID) = 50
        .ColWidth(col_SchList_序号) = 450
        .ColWidth(col_SchList_姓名) = 800
        .ColWidth(col_SchList_医嘱内容) = 3000
        .ColWidth(col_SchList_预约开始时间) = 1800
        .ColWidth(col_SchList_预约结束时间) = 1800
        .ColWidth(col_SchList_手机号) = 2000
        
        '显示标题
        .TextMatrix(0, col_SchList_ID) = "ID"
        .TextMatrix(0, col_SchList_姓名) = "姓名"
        .TextMatrix(0, col_SchList_序号) = "序号"
        .TextMatrix(0, col_SchList_医嘱内容) = "医嘱内容"
        .TextMatrix(0, col_SchList_预约开始时间) = "开始时间"
        .TextMatrix(0, col_SchList_预约结束时间) = "结束时间"
        .TextMatrix(0, col_SchList_诊室名称) = "诊室名称"
        .TextMatrix(0, col_SchList_执行过程) = "执行过程"
        .TextMatrix(0, col_SchList_检查注意) = "检查注意"
        .TextMatrix(0, col_SchList_手机号) = "手机号"
        '从数据库加载数据
        If rsTemp.EOF = False Then
            For i = 1 To rsTemp.RecordCount
                If rsTemp!婴儿 <> 0 Then
                    strSQL = "Select A.开嘱时间,Nvl(B.婴儿姓名, A.姓名 || '之子' || Trim(To_Char(B.序号, '9'))) As 婴儿姓名, B.婴儿性别, B.出生时间" & vbNewLine & _
                                 "  From 病人医嘱记录 A, 病人新生儿记录 B " & vbNewLine & _
                                 "  Where a.病人ID = b.病人ID  And b.序号 = [2] And a.ID = [1]"
                    Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "提取婴儿信息", CLng(rsTemp!医嘱ID), CLng(rsTemp!婴儿))
                    .TextMatrix(i, col_SchList_姓名) = rsBaby!婴儿姓名
                Else
                    .TextMatrix(i, col_SchList_姓名) = rsTemp!姓名
                End If
                
                .TextMatrix(i, col_SchList_ID) = rsTemp!ID
                .TextMatrix(i, col_SchList_序号) = rsTemp!序号
                .TextMatrix(i, col_SchList_医嘱内容) = rsTemp!医嘱内容
                .TextMatrix(i, col_SchList_预约开始时间) = Format(rsTemp!预约开始时间, "YYYY-MM-DD HH:MM:SS")
                .TextMatrix(i, col_SchList_预约结束时间) = Format(rsTemp!预约结束时间, "YYYY-MM-DD HH:MM:SS")
                .TextMatrix(i, col_SchList_诊室名称) = NVL(rsTemp!诊室名称)
                .TextMatrix(i, col_SchList_手机号) = NVL(rsTemp!手机号)
                .TextMatrix(i, col_SchList_执行过程) = IIf(NVL(rsTemp!执行过程, 0) = -1, "已驳回", IIf(NVL(rsTemp!执行过程, 0) = 0 Or NVL(rsTemp!执行过程, 0) = 1, "已登记", IIf(NVL(rsTemp!执行过程, 0) = 2, "已报到", IIf(NVL(rsTemp!执行过程, 0) = 3, "已检查", IIf(NVL(rsTemp!执行过程, 0) = 4, "已报告", IIf(NVL(rsTemp!执行过程, 0) = 5, "已审核", IIf(NVL(rsTemp!执行过程, 0) = 6, "已完成", "未检查")))))))
                '执行过程: -1-驳回；0或1-已登记；2-已报到；3-已检查；4-已报告；5-已审核；6-已完成
                .TextMatrix(i, col_SchList_检查注意) = NVL(rsTemp!检查注意)
                
                '设置颜色
                 '设置颜色
                If Not (NVL(rsTemp!执行过程, 0) = 0 Or NVL(rsTemp!执行过程, 0) = 1 Or NVL(rsTemp!执行过程, 0) = 2) Then
                    lngColor = mlngColorLblDone
                ElseIf Format(rsTemp!预约开始时间, "YYYY-MM-DD HH:MM:SS") < Format(Now, "YYYY-MM-DD HH:MM:SS") Then
                    lngColor = mlngColorLblPassed
                Else
                    lngColor = mlngColorLblWaiting
                End If
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = lngColor
                
                rsTemp.MoveNext
            Next i
        End If
        
        '隐藏后台数据列
        .ColHidden(col_SchList_ID) = True
        
    End With
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshCalendar()
'------------------------------------------------
'功能：刷新日历
'参数：
'返回：无
'------------------------------------------------
    
    On Error GoTo err
    
    mstrSchRestDate = RefeshSchRestDay(mlngOrderID, mlngSchDeviceID, dpCalendar.LastVisibleDay)
    
    dpCalendar.RedrawControl
    
    '刷新预约周视图
    Call LoadWeekView
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadWeekView()
'------------------------------------------------
'功能：加载预约的周视图
'参数：
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim dtMonday As Date
    Dim lngCapacity As Long
    Dim lngScheduledCount As Long
    Dim lngVacancy As Long
    
    On Error GoTo err
    
    With vsfWeekView
        .Clear
        .Cols = 4
        .Rows = 8
        .FixedRows = 1
        .FixedCols = 1
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarVertical
        .CellAlignment = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .ExtendLastCol = True
        .ColWidthMin = 300
        
        .RowHeight(0) = 300
        For i = 1 To .Rows - 1
            .RowHeight(i) = 550
        Next i
        
        .ColWidth(0) = 1000
        
        .TextMatrix(0, col_WeekView_星期) = "日期"
        .TextMatrix(0, col_WeekView_空余) = "空余"
        .TextMatrix(0, col_WeekView_已预约) = "已预约"
        .TextMatrix(0, col_WeekView_总容量) = "总容量"
        
        '先找到本周一
        dtMonday = mschDate - Weekday(mschDate, vbMonday) + 1
        
        '从数据库加载数据
        For i = 1 To 7
            .TextMatrix(i, col_WeekView_星期) = "周" & WeekdayChinese(CLng(i)) & vbCrLf & Format(dtMonday + i - 1, "M月D日")
            .RowData(i) = dtMonday + i - 1
            
            If DateScheduleInfo(mlngOrderID, dtMonday + i - 1, mlngSchDeviceID, lngCapacity, lngScheduledCount, lngVacancy) = True Then
                .Cell(flexcpBackColor, i, col_WeekView_空余) = vbGreen
                .TextMatrix(i, col_WeekView_空余) = lngVacancy
            Else
                .Cell(flexcpBackColor, i, col_WeekView_空余) = vbRed
                .TextMatrix(i, col_WeekView_空余) = 0
            End If
            
            .TextMatrix(i, col_WeekView_已预约) = lngScheduledCount
            .TextMatrix(i, col_WeekView_总容量) = lngCapacity
            
            '选中当天
            If mschDate = dtMonday + i - 1 Then
                .Select i, 0, i, .Cols - 1
            End If
        Next i
        
    End With
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function WeekdayChinese(lngWeekday As Long) As String
'------------------------------------------------
'功能：将数字的周序号，解析成中文的周一到周日
'参数： lngWeekday -- 周序号，1-7
'返回：周的中文字符，一、二、三、四、五、六、日
'------------------------------------------------
    On Error GoTo err
    
    Select Case Val(lngWeekday)
        Case 1:
            WeekdayChinese = "一"
        Case 2:
            WeekdayChinese = "二"
        Case 3:
            WeekdayChinese = "三"
        Case 4:
            WeekdayChinese = "四"
        Case 5:
            WeekdayChinese = "五"
        Case 6:
            WeekdayChinese = "六"
        Case 7:
            WeekdayChinese = "日"
        Case Else
            WeekdayChinese = "几"
    End Select
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function DateScheduleInfo(ByVal lngOrderID As Long, ByVal dtDate As Date, _
    ByVal lngDeviceID As Long, ByRef lngDayCapacity As Long, ByRef lngScheduledCount As Long, _
    ByRef lngVacancy As Long) As Boolean
'-----------------------------------------------------------
'功能:获取当天的预约情况
'入参:  lngOrderID -- 医嘱ID
'       dtDate -- 日期
'       lngSchDeviceID -- 预约设备ID
'       lngDayCapacity -- 预约总容量
'       lngScheduledCount -- 已预约数量
'       lngVacancy -- 剩余容量
'返回: True -- 可预约；False -- 不可预约
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngPlanID As Long
    Dim dtStartTime As Date

    On Error GoTo err
    
    lngDayCapacity = 0
    lngScheduledCount = 0
    lngVacancy = 0
    
    lngPlanID = schTimeTable.GetSchPlanID(lngDeviceID, dtDate, False, True)
    
    If lngPlanID <> 0 Then
        strSQL = "Select nvl(Sum(b.预约容量), 0) as thecount From 影像预约时间计划 B Where b.预约方案id =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取预约容量", lngPlanID)
        
        If rsTemp.EOF = False Then
            lngDayCapacity = rsTemp!thecount
        End If
            
        strSQL = "Select Count(A.ID) as thecount From 影像预约记录 A Where a.预约设备id = [1] And " _
            & " a.预约开始时间 Between to_date(to_char([2], 'yyyy-mm-dd') || ' 00:00:01', 'yyyy-mm-dd hh24:mi:ss') And " _
            & " to_date(to_char([2], 'yyyy-mm-dd') || ' 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询已预约数量", lngDeviceID, CDate(Format(dtDate, "yyyy-MM-dd 00:00:00")))
        
        If rsTemp.EOF = False Then
            lngScheduledCount = rsTemp!thecount
        End If
        
        If Format(dtDate, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
            '今天，取当前时间2小时之后的预约时间段
            strSQL = " Select a.id, a.开始时间, a.结束时间,a.预约容量 From 影像预约时间计划 A " _
                    & " Where a.预约方案id = [1] and " _
                    & " to_char(a.结束时间, 'hh24:mi:ss') > to_char(sysdate + 2 / 24, 'hh24:mi:ss') " _
                    & " Order By a.开始时间 desc"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取今天的预约容量", lngPlanID)
            
            While rsTemp.EOF = False
                lngVacancy = lngVacancy + rsTemp!预约容量
                dtStartTime = rsTemp!开始时间
                rsTemp.MoveNext
            Wend
            
            strSQL = "Select Count(A.ID) as thecount From 影像预约记录 A Where a.预约设备id = [1] And " _
                    & "  to_char(a.预约开始时间, 'hh24:mi:ss') > to_char([2], 'hh24:mi:ss') " _
                    & " And trunc(a.预约开始时间) = trunc(sysdate)"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询已预约数量", lngDeviceID, CDate(Format(dtStartTime, "yyyy-MM-dd hh:mm:ss")))
            
            If rsTemp.EOF = False Then
                lngVacancy = lngVacancy - rsTemp!thecount
            End If
        Else
            If InStr(mstrSchRestDate, Format(dtDate, "YYYY-MM-DD")) > 0 Then
                lngVacancy = 0
            Else
                lngVacancy = lngDayCapacity - lngScheduledCount
            End If
        End If
    End If
    
        DateScheduleInfo = (lngVacancy <> 0)
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfWeekView_Click()
    On Error GoTo err
    
    If vsfWeekView.Rows > 1 Then
        mschDate = Format(vsfWeekView.RowData(vsfWeekView.RowSel), "YYYY-MM-DD")
        Call ChangeCalendar(mschDate)
        Call RefreshSchedule
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
