VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmRegistPlanList 
   BorderStyle     =   0  'None
   Caption         =   "当前有效号别"
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTime 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   2295
      ScaleHeight     =   2745
      ScaleWidth      =   4425
      TabIndex        =   7
      Top             =   4485
      Width           =   4425
      Begin MSComctlLib.TabStrip tbWeekTime 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Top             =   60
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   503
         Style           =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vsTime 
         Height          =   2145
         Left            =   0
         TabIndex        =   8
         Top             =   375
         Width           =   4410
         _cx             =   7779
         _cy             =   3784
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
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRegistPlanList.frx":0000
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
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   -2055
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   5
      Top             =   4170
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   2295
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4575
         _Version        =   589884
         _ExtentX        =   8070
         _ExtentY        =   4048
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picStop 
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   4455
      ScaleHeight     =   3285
      ScaleWidth      =   3660
      TabIndex        =   3
      Top             =   3840
      Width           =   3660
      Begin VSFlex8Ctl.VSFlexGrid vsStop 
         Height          =   2415
         Left            =   150
         TabIndex        =   4
         Top             =   180
         Width           =   6345
         _cx             =   11192
         _cy             =   4260
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistPlanList.frx":0075
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
      End
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   75
      ScaleHeight     =   2250
      ScaleWidth      =   5160
      TabIndex        =   0
      Top             =   750
      Width           =   5160
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2145
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3405
         _cx             =   6006
         _cy             =   3784
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
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   26
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistPlanList.frx":0164
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
         AllowUserFreezing=   1
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
            TabIndex        =   2
            Top             =   210
            Width           =   210
            Begin VB.Image imgColList 
               Height          =   195
               Left            =   0
               Picture         =   "frmRegistPlanList.frx":04C5
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmRegistPlanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long, mArrFilter As Variant  '过滤条件
Private mblnNotMoveRow As Boolean '不移动行
Private mblnListSetFocus As Boolean  '当前当标位置
Private mrsRoom As ADODB.Recordset
Private mrsPlanRoom As ADODB.Recordset
Private Const conPane_List = 1
Private Const conPane_Plan = 2
Private Const conPane_Stop = 3
Private mblnHaveDate As Boolean
Private mblnHaveUnit As Boolean
Public Event zlPopuMenu(intType As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Private WithEvents mfrm计划 As frmRegistPlanIntend
Attribute mfrm计划.VB_VarHelpID = -1

Private Enum mPgIndex
    pg_计划 = 1
    pg_时段 = 2
    pg_停用 = 3
    pg_单位 = 4
End Enum
Private mblnNotClick As Boolean
Private mrsTime As ADODB.Recordset
Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区哉设置
    '编制:刘兴洪
    '日期:2009-09-14 18:06:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single
    Dim strReg As String
    Dim panThis As Pane
    Set panThis = dkpMan.CreatePane(conPane_List, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "挂号安排信息"
    panThis.Handle = picList.Hwnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Tag = conPane_List
    Set panThis = dkpMan.CreatePane(conPane_Plan, 250, 580, DockBottomOf, panThis)
    panThis.Title = ""
    panThis.Tag = conPane_Plan
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picPage.Hwnd
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = picList.Hwnd
    Case conPane_Plan
        Item.Handle = picPage.Hwnd
    End Select
End Sub
 
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2009-09-09 15:45:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intType As Integer, i As Integer, objGrid As VSFlexGrid
    i = 0
    With vsList
        .Redraw = flexRDNone
        .Rows = 3: .FixedRows = 2
        .FixedCols = 1
        .Cols = 39:   .Clear
        .FrozenCols = 6
        .TextMatrix(0, i) = "  ": .ColWidth(i) = 285
        .TextMatrix(1, i) = "  ":  .ColKey(i) = "标志": i = i + 1
        
        .TextMatrix(0, i) = "ID": .ColHidden(i) = True: .ColWidth(i) = 0
        .TextMatrix(1, i) = "ID": .ColKey(i) = "ID": i = i + 1
         
        .TextMatrix(0, i) = "状态": .ColWidth(i) = 200
        .TextMatrix(1, i) = "状态": .ColKey(i) = "状态": i = i + 1
         
        .TextMatrix(0, i) = "号类": .ColWidth(i) = 720
        .TextMatrix(1, i) = "号类": .ColKey(i) = "号类": i = i + 1

        .TextMatrix(0, i) = "号别": .ColWidth(i) = 480
        .TextMatrix(1, i) = "号别": .ColKey(i) = "号别": i = i + 1

        .TextMatrix(0, i) = "科室": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "科室": .ColKey(i) = "科室": i = i + 1
        .TextMatrix(0, i) = "项目": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "项目": .ColKey(i) = "项目": i = i + 1
        .TextMatrix(0, i) = "医生":: .ColWidth(i) = 1000
        .TextMatrix(1, i) = "医生": .ColKey(i) = "医生": i = i + 1
        .TextMatrix(0, i) = "已挂数": .ColWidth(i) = 660
        .TextMatrix(1, i) = "已挂数": .ColKey(i) = "已挂数": i = i + 1
        .TextMatrix(0, i) = "已约数": .ColWidth(i) = 660
        .TextMatrix(1, i) = "已约数": .ColKey(i) = "已约数": i = i + 1
        '59611 李南春 2014/04/11 16:40:49 添加一列显示预约已接收数
        .TextMatrix(0, i) = "其中已接收数": .ColWidth(i) = 1305
        .TextMatrix(1, i) = "其中已接收数": .ColKey(i) = "其中已接收数": i = i + 1
        
        .TextMatrix(0, i) = "建档": .ColWidth(i) = 495
        .TextMatrix(1, i) = "建档": .ColKey(i) = "建档": i = i + 1
        .TextMatrix(0, i) = "周日": .ColWidth(i) = 420
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周日-安排": i = i + 1
        .TextMatrix(0, i) = "周日": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周日-限号": i = i + 1
        .TextMatrix(0, i) = "周日": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周日-限约": i = i + 1

        .TextMatrix(0, i) = "周一": .ColWidth(i) = 420
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周一-安排": i = i + 1
        .TextMatrix(0, i) = "周一": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周一-限号": i = i + 1
        .TextMatrix(0, i) = "周一": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周一-限约": i = i + 1

        .TextMatrix(0, i) = "周二": .ColWidth(i) = 420
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周二-安排": i = i + 1
        .TextMatrix(0, i) = "周二": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周二-限号": i = i + 1
        .TextMatrix(0, i) = "周二": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周二-限约": i = i + 1

        .TextMatrix(0, i) = "周三": .ColWidth(i) = 420
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周三-安排": i = i + 1
        .TextMatrix(0, i) = "周三": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周三-限号": i = i + 1
        .TextMatrix(0, i) = "周三": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周三-限约": i = i + 1

        .TextMatrix(0, i) = "周四": .ColWidth(i) = 420
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周四-安排": i = i + 1
        .TextMatrix(0, i) = "周四": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周四-限号": i = i + 1
        .TextMatrix(0, i) = "周四": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周四-限约": i = i + 1

        .TextMatrix(0, i) = "周五": .ColWidth(i) = 420
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周五-安排": i = i + 1
        .TextMatrix(0, i) = "周五": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周五-限号": i = i + 1
        .TextMatrix(0, i) = "周五": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周五-限约": i = i + 1

        .TextMatrix(0, i) = "周六": .ColWidth(i) = 420
        .TextMatrix(1, i) = "安排": .ColKey(i) = "周六-安排": i = i + 1
        .TextMatrix(0, i) = "周六": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限号": .ColKey(i) = "周六-限号": i = i + 1
        .TextMatrix(0, i) = "周六": .ColWidth(i) = 420
        .TextMatrix(1, i) = "限约": .ColKey(i) = "周六-限约": i = i + 1
        .TextMatrix(0, i) = "分诊方式": .ColWidth(i) = 855
        .TextMatrix(1, i) = "分诊方式": .ColKey(i) = "分诊方式": i = i + 1
        .TextMatrix(0, i) = "IDS": .ColWidth(i) = 0: .ColHidden(i) = True
        .TextMatrix(1, i) = "IDS": .ColKey(i) = "IDS": i = i + 1
        .TextMatrix(0, i) = "有效范围": .ColWidth(i) = 2800
        .TextMatrix(1, i) = "有效范围": .ColKey(i) = "有效范围": i = i + 1
        .TextMatrix(0, i) = "序号控制": .ColWidth(i) = 765
        .TextMatrix(1, i) = "序号控制": .ColKey(i) = "序号控制": i = i + 1
        .TextMatrix(0, i) = "停用日期": .ColWidth(i) = 1860
        .TextMatrix(1, i) = "停用日期": .ColKey(i) = "停用日期": i = i + 1
        .TextMatrix(0, i) = "应诊诊室": .ColWidth(i) = 2000
        .TextMatrix(1, i) = "应诊诊室": .ColKey(i) = "应诊诊室": i = i + 1
        .Cell(flexcpText, 0, 0, .Rows - 1) = " "
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        For i = 0 To .Cols - 1
            .MergeCol(i) = True:
            .FixedAlignment(i) = flexAlignCenterCenter
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            Select Case .ColKey(i)
            Case "ID", "标志", "IDS"
                 .ColData(i) = "-1|1"
            Case "号类", "号别"
                .ColData(i) = "1|0"
            End Select
        Next
         .MergeRow(0) = True: .MergeRow(1) = True
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    Call InitPancel
    Call InitUnitData
    Call InitPage
    Call InitVsGrid
    Call mfrm计划.SetGotFocus(True): Call vsList_LostFocus: Call vsStop_LostFocus
    vsList_GotFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "有效号-号别列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
    zl_vsGrid_Para_Save mlngModule, vsStop, Me.Caption, "有效号-停用计划", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
    If Not mfrm计划 Is Nothing Then Unload mfrm计划
    Unload frmUnitReg
    Set mfrm计划 = Nothing
End Sub

Public Sub ReloagUnitRegPlan()
    If mfrm计划 Is Nothing Then Exit Sub
    mfrm计划.ReLoadUnitPlan
End Sub


Private Sub InitUnitData()
    '初始化合作单位信息
    Dim strSQL  As String
    Dim rsTmp   As ADODB.Recordset
    
    strSQL = "Select Count(0) as count From 挂号合作单位 Where Rownum=1"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    mblnHaveUnit = Val(Nvl(rsTmp!Count)) > 0
    Set rsTmp = Nothing
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub imgColList_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgList.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsList, lngLeft, lngTop, imgColList.Height)
    
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "有效号-号别列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub
 
 
Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        vsList.Left = .ScaleLeft
        vsList.Width = .ScaleWidth
        vsList.Top = .ScaleTop
        vsList.Height = .ScaleHeight
    End With
End Sub

 

Private Sub picPage_Resize()
    Err = 0: On Error Resume Next
    With picPage
        tbPage.Left = .ScaleLeft
        tbPage.Width = .ScaleWidth
        tbPage.Top = .ScaleTop
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub picTime_Resize()
    Err = 0: On Error Resume Next
    With picTime
        tbWeekTime.Left = .ScaleLeft + 100
        tbWeekTime.Top = .ScaleTop + 50
        tbWeekTime.Width = .ScaleWidth
        If tbWeekTime.Visible = False Then
            vsTime.Top = .ScaleTop
        Else
            vsTime.Top = tbWeekTime.Top + tbWeekTime.Height + 50
        End If
        vsTime.Left = .ScaleLeft
        vsTime.Width = .ScaleWidth
       ' vsTime.Top = .ScaleTop
        vsTime.Height = .ScaleHeight - vsTime.Top
    End With
End Sub
 

Private Function HaveData() As Boolean
    '功能:是否有数据
    If Me.ActiveControl Is vsList Then
        With Me.vsList
            HaveData = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
        End With
    Else
        HaveData = mfrm计划.zlHaveData
    End If
End Function

Public Function Have计划() As Boolean
    '功能:该安排下是否有计划信息
    '问题号:51429
    Dim i As Long
    Dim blnHasData As Boolean
    
    For i = 0 To mfrm计划.vsPlan.Rows - 1
       With mfrm计划.vsPlan
           If Val(Nvl(.TextMatrix(i, .ColIndex("ID")), "0")) <> 0 Then
                blnHasData = True
                Exit For
           End If
       End With
    Next
    Have计划 = blnHasData
End Function
Public Function 是否选中计划列表() As Boolean
    '功能:该安排下是否选中计划列表
    '问题号:51429
    是否选中计划列表 = mfrm计划.mblnSelected计划
End Function
 
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2009-09-09 11:24:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsGrid As VSFlexGrid, rsTemp As New ADODB.Recordset, strSQL As String
     If Not Me.ActiveControl Is vsList Then
        mfrm计划.zlRptPrint (bytFunc): Exit Sub
     End If
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & IIf(Not Me.ActiveControl Is vsList, "挂号安排表", "挂号计划表")
    
    If CStr(mArrFilter("有效期")(0)) <> "1901-01-01" Then
        objRow.Add "效期范围：" & CStr(mArrFilter("有效期")(0)) & "至" & CStr(mArrFilter("有效期")(1))
    End If
    If Val(mArrFilter("科室ID")) > 0 Then
        strSQL = "Select 名称 From 部门表 where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mArrFilter("科室ID")))
        If rsTemp.EOF Then
            objRow.Add "科室：所有科室"
        Else
            objRow.Add "科室：" & Nvl(rsTemp!名称)
        End If
    ElseIf Val(mArrFilter("科室ID")) = -1 Then
        objRow.Add "科室：操作员所属科室"
    Else
        objRow.Add "科室：所有科室"
    End If
    Select Case mArrFilter("医生ID")(1)
    Case "ID"
        strSQL = "Select 姓名 From 人员表 where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mArrFilter("医生ID")(0)))
        If rsTemp.EOF Then
            objRow.Add "医生：所有"
        Else
            objRow.Add "医生：" & Nvl(rsTemp!姓名)
        End If
    Case "UPR", "NONE"
            objRow.Add "医生：" & CStr(mArrFilter("医生ID")(0))
    End Select
    objPrint.UnderAppRows.Add objRow
    Set vsGrid = vsList
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("标志") Then .ColWidth(intCol) = 0
        Next
    End With
    
    Set objPrint.Body = vsGrid
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
    
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
End Sub
Private Sub picStop_Resize()
    Err = 0: On Error Resume Next
    With picStop
        vsStop.Left = .ScaleLeft
        vsStop.Top = .ScaleTop
        vsStop.Height = .ScaleHeight
        vsStop.Width = .ScaleWidth
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call picTime_Resize
End Sub

Private Sub tbWeekTime_Click()
    Dim lng安排ID  As Long, bln序号控制 As Boolean
    If mblnNotClick = True Then Exit Sub
   With vsList
        lng安排ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        bln序号控制 = Trim(.TextMatrix(.Row, .ColIndex("序号控制"))) <> ""
    End With
    Call LoadTimePlan(lng安排ID, bln序号控制)
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng安排ID As Long, bln序号控制 As Boolean
    Dim str号类 As String, str号别 As String, str科室 As String
    If OldRow > 0 Then
        zl_VsGridRowChange vsList, OldRow, NewRow, OldCol, NewCol
    End If
    
    If OldRow = NewRow Or mblnNotMoveRow Then Exit Sub
    
    With vsList
        lng安排ID = Val(.TextMatrix(NewRow, .ColIndex("ID")))
        bln序号控制 = Trim(.TextMatrix(NewRow, .ColIndex("序号控制"))) <> ""
        str号类 = Trim(.TextMatrix(NewRow, .ColIndex("号类")))
        str号别 = Trim(.TextMatrix(NewRow, .ColIndex("号别")))
        str科室 = Trim(.TextMatrix(NewRow, .ColIndex("科室")))
        
    End With
    Call mfrm计划.zlShowPlan(lng安排ID, str号类, str号别, str科室)
    Call LoadStopPlan(lng安排ID)
    Call LoadTimePlan(lng安排ID, bln序号控制)
    Call LoadUnitReg(lng安排ID)
    If OldRow <> NewRow Then
        zlHaveDatPlanForPlan
    End If
    '问题号:51429
    mfrm计划.mblnSelected计划 = False
    On Error Resume Next
    If vsList.Enabled And vsList.Visible Then vsList.SetFocus
    
    DoEvents
    Call vsList_GotFocus
End Sub

Private Sub LoadUnitReg(ByVal lng安排ID As Long)
     frmUnitReg.ShowUnitReg lng安排ID
End Sub

Public Sub zlHaveDatPlanForPlan()
    Dim lngNum  As Long
    Dim lngRow  As Long
    
     With vsList
        If .Row = 0 Then Exit Sub
        lngRow = .Row
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("周日-限号")))
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("周一-限号"))) + lngNum
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("周二-限号"))) + lngNum
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("周三-限号"))) + lngNum
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("周四-限号"))) + lngNum
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("周五-限号"))) + lngNum
        lngNum = Val(.TextMatrix(lngRow, .ColIndex("周六-限号"))) + lngNum
     End With
     mblnHaveDate = lngNum > 0
End Sub

Public Sub ReloadTimePlan(Optional ByVal blnReloadPlan As Boolean = False)
    '***********************************
    '挂号安排时段更改以后
    '更新挂号安排时段显示列表
    '***********************************
    Dim lng安排ID       As Long
    Dim bln序号控制     As Boolean
    If blnReloadPlan Then
        mfrm计划.ReloadTimePlan
        Exit Sub
    End If
    With vsList
        lng安排ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        bln序号控制 = Trim(.TextMatrix(.Row, .ColIndex("序号控制"))) <> ""
    End With
    Call LoadTimePlan(lng安排ID, bln序号控制, False, True)
    zlControl.ControlSetFocus vsList, True
End Sub
Private Sub LoadStopPlan(ByVal lng安排ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载停用计划安排
    '编制:刘兴洪
    '日期:2010-09-09 11:54:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long, strCurDate As String
    On Error GoTo errHandle
    
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM")
    '加载该挂号项目的的停用时间信息
    strSQL = "Select 安排ID,序号,开始停止时间,结束停止时间,制订人,制订日期,备注 From 挂号安排停用状态 where 安排ID=[1] Order by 开始停止时间,制订日期"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
    With vsStop
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("序号")) = i
            .Cell(flexcpData, i, .ColIndex("序号")) = Val(Nvl(rsTemp!序号))
            .TextMatrix(i, .ColIndex("开始停用时间")) = Format(rsTemp!开始停止时间, "yyyy-mm-dd HH:MM")
            .TextMatrix(i, .ColIndex("结束停用时间")) = Format(rsTemp!结束停止时间, "yyyy-mm-dd HH:MM")
            .TextMatrix(i, .ColIndex("制订人")) = Nvl(rsTemp!制订人)
            .TextMatrix(i, .ColIndex("制订日期")) = Format(rsTemp!制订日期, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(i, .ColIndex("备注")) = Nvl(rsTemp!备注)
            If Format(rsTemp!结束停止时间, "yyyy-mm-dd HH:MM:SS") < strCurDate Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = &H8000000C
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        '恢复列设置
         zl_vsGrid_Para_Restore mlngModule, vsStop, Me.Caption, "有效号-停用计划", True
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub LoadTimePlan(ByVal lng安排ID As Long, ByVal bln序号控制 As Boolean, _
    Optional bln计划 As Boolean = False, Optional blnReload As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载停用计划安排
    '入参:bln计划-是否加载计划的时间段
    '编制:刘兴洪
    '日期:2010-09-09 11:54:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str星期 As String
    Dim i As Long, r As Integer, str时点 As String, strTime As String, strKey As String
    Static lngPre安排ID  As Long
    On Error GoTo errHandle
    '加载该挂号项目的的停用时间信息
    If mrsTime Is Nothing Then
        lngPre安排ID = -1
    ElseIf mrsTime.State <> 1 Then
         lngPre安排ID = -1
    End If
    If lngPre安排ID <> lng安排ID Or blnReload Then
        lngPre安排ID = lng安排ID
        strSQL = "" & _
        "   Select decode(星期,'周日',1,'周一',2,'周二',3,'周三',4,'周四',5,'周五',6,7) as 排序,星期,to_char(开始时间,'HH24')||':00' as 时点,序号,to_char(开始时间,'hh24:mi')||'-' ||to_char(结束时间,'hh24:mi') as 时间范围, " & _
        "               限制数量,是否预约" & _
        "   From  挂号安排时段  " & _
        "   Where 安排ID=[1] And 限制数量>0" & _
        "   Order by 排序,时点,序号"
        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
        tbWeekTime.Tabs.Clear
        With mrsTime
            strTime = ""
            Do While Not .EOF
                If strTime <> Nvl(mrsTime!星期) Then
                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!星期), Nvl(mrsTime!星期)
                    strTime = Nvl(mrsTime!星期)
                End If
                .MoveNext
            Loop
            mblnNotClick = True
            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
            If tbWeekTime.Tabs.Count > 0 Then
                tbWeekTime.Tabs(1).Selected = True
            End If
            mblnNotClick = False
            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
        
            Call picTime_Resize
        End With
    End If
    str星期 = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "星期='" & str星期 & "'"
     str时点 = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 800: .RowHeightMin = 800
        .Rows = 1: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln序号控制 Then
             .Cols = 8: .FixedCols = 0
             r = 0: i = -1
            Do While Not mrsTime.EOF
                i = i + 1
                If i > .Cols - 1 Then r = r + 1: i = 0
                strTime = "预约" & Val(Nvl(mrsTime!限制数量)) & "人" & vbCrLf & vbCrLf
                strTime = strTime & mrsTime!时间范围
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i) = strTime
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
             Exit Sub
        End If
        Do While Not mrsTime.EOF
            If str时点 <> Nvl(mrsTime!时点) Then
                r = r + 1
                str时点 = Nvl(mrsTime!时点)
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, 0) = str时点
                i = 0
            End If
            i = i + 1
            strTime = mrsTime!序号 & vbCrLf & vbCrLf
            strTime = strTime & mrsTime!时间范围
            If i > .Cols - 1 Then .Cols = .Cols + 1
            If r > .Rows - 1 Then .Rows = .Rows + 1
            .TextMatrix(r, i) = strTime
            If Val(Nvl(mrsTime!是否预约)) = 1 Then
                .Cell(flexcpForeColor, r, i, r, i) = vbBlue
                .Cell(flexcpFontBold, r, i, r, i) = True
            End If
            mrsTime.MoveNext
        Loop
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "有效号-号别列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
        With vsList
            If .ColKey(Col) Like "周*" Then
                 Position = Col
            End If
        End With
End Sub

Private Sub vsList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent zlPopuMenu(0, Button, Shift, X, Y)
End Sub
    
Private Sub vsList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsList
        If Col = .ColIndex("标志") Then Cancel = True
    End With
End Sub

Private Sub vsList_DblClick()
    With vsList
        If Trim(.TextMatrix(.Row, .ColIndex("停用日期"))) <> "" Then
            Call frmRegistPlanEdit.ShowEdit(Me, edt_查阅, mlngModule, mstrPrivs, .TextMatrix(.Row, .ColIndex("id")))
            Exit Sub
        End If
    End With
    zlExecuteModifyList Me
End Sub
 
Private Sub vsList_GotFocus()
    zl_VsGridGotFocus vsList
    mblnListSetFocus = True
End Sub
Private Sub vsList_LostFocus()
    zl_VsGridLOSTFOCUS vsList
    With vsList
        .ForeColorSel = .Cell(flexcpForeColor, .Row, .Col)
    End With
End Sub
 
Private Sub LoadDataToList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, lngRow As Long, strSQL As String
    Dim blnHistory As Boolean, strStartDate As String, rs挂号汇总 As ADODB.Recordset, strWhere As String
    Dim lngPreID As Long, strTable As String
    Dim lng限制数量 As Long
    Dim str显示条件 As String
    
    Err = 0: On Error GoTo Errhand:
    strWhere = ""
    If CStr(mArrFilter("有效期")(0)) <> "1901-01-01" Then
        strFilter = "  And Nvl(A.开始时间,To_Date('3000-01-01','YYYY-MM-DD'))>=[4]   And Nvl(A.终止时间,To_Date('1900-01-01','YYYY-MM-DD'))<=[5]"
    End If
    If Val(mArrFilter("科室ID")) > 0 Then
        strFilter = strFilter & " And A.科室ID=[1]"
        strWhere = strWhere & " And A.科室ID=[1]"
    End If
    If Val(mArrFilter("科室ID")) = -1 Then
        strFilter = strFilter & " And  A.科室ID in (Select 部门ID From 部门人员 where 人员id=[6]) "
        strWhere = strWhere & " And  A.科室ID in (Select 部门ID From 部门人员 where 人员id=[6]) "
    End If
    Select Case mArrFilter("医生ID")(1)
    Case "ID"
         strFilter = strFilter & "  And A.医生ID=[2]"
         strWhere = strWhere & "  And A.医生ID=[2]"
    Case "UPR"
         strFilter = strFilter & " And Upper(A.医生姓名)=[3]"
         strWhere = strWhere & " And Upper(A.医生姓名)=[3]"
    Case "NONE"
         strFilter = strFilter & " And A.医生姓名=[3]"
         strWhere = strWhere & " And A.医生姓名=[3]"
    End Select
    strTable = "" & _
    "   Select A.ID, " & _
    "             Sum(Decode(B.限制项目,'周日',B.限号数,0)) as 周日限号, Sum(Decode(B.限制项目,'周日',B.限约数))  as 周日限约," & _
    "             Sum(Decode(B.限制项目,'周一',B.限号数,0)) as 周一限号, Sum(Decode(B.限制项目,'周一',B.限约数))  as 周一限约," & _
    "             Sum(Decode(B.限制项目,'周二',B.限号数,0)) as 周二限号, Sum(Decode(B.限制项目,'周二',B.限约数))  as 周二限约," & _
    "             Sum(Decode(B.限制项目,'周三',B.限号数,0)) as 周三限号, Sum(Decode(B.限制项目,'周三',B.限约数))  as 周三限约," & _
    "             Sum(Decode(B.限制项目,'周四',B.限号数,0)) as 周四限号, Sum(Decode(B.限制项目,'周四',B.限约数))  as 周四限约," & _
    "             Sum(Decode(B.限制项目,'周五',B.限号数,0)) as 周五限号, Sum(Decode(B.限制项目,'周五',B.限约数))  as 周五限约," & _
    "             Sum(Decode(B.限制项目,'周六',B.限号数,0)) as 周六限号, Sum(Decode(B.限制项目,'周六',B.限约数))  as 周六限约" & _
    "   From 挂号安排 A,挂号安排限制 B  " & _
    "   Where A.ID=B.安排ID(+) " & strFilter & _
    "   Group by A.ID"
    '问题:32512，38505
    strSQL = "" & _
    "   Select 科室id, 项目id,  nvl(医生姓名,'刘兴洪测试医生') As 医生姓名, nvl(医生id,0) As 医生ID, A.号码, 已挂数, 已约数, 其中已接收 " & _
    "   From 病人挂号汇总 A " & _
    "   Where 日期 = Trunc(Sysdate) " & strWhere
     '问题 :45525
    If Nvl(mArrFilter("显示停用安排"), 0) = 0 And Nvl(mArrFilter("显示删除安排"), 0) = 0 Then '停用与删除同时不显示的时候
        str显示条件 = "And (Nvl(a.停用日期,to_date('3000-1-1','yyyy-mm-dd')) > sysdate)"
    ElseIf Nvl(mArrFilter("显示停用安排"), 0) = 1 And Nvl(mArrFilter("显示删除安排"), 0) = 0 Then '停用显示与删除不显示的时候
        str显示条件 = "And (A.是否删除<> 1)"
    ElseIf Nvl(mArrFilter("显示停用安排"), 0) = 0 And Nvl(mArrFilter("显示删除安排"), 0) = 1 Then '停用不显示与删除显示的时候
        str显示条件 = "And(A.是否删除=1or Nvl(a.停用日期,to_date('3000-1-1','yyyy-mm-dd')) > sysdate)"
    End If
    
    strSQL = _
      "Select A.ID,A.号码 as 号别,A.号类,A.科室ID,C.名称 as 科室,A.项目ID,B.名称 as 项目," & _
      "         A.医生姓名 as 医生,A.医生ID, Nvl(A.病案必须,0) as 建档,Nvl(A.序号控制,0) as 序号控制," & _
      "         Decode(Nvl(A.分诊方式,0),0,'不分诊',1,'指定诊室',2,'动态分诊',3,'平均分诊') as 分诊方式," & _
      "         A.周日,A1.周日限号,A1.周日限约,A.周一,A1.周一限号,A1.周一限约,A.周二,A1.周二限号,A1.周二限约, " & _
      "         A.周三,A1.周三限号,A1.周三限约,A.周四,A1.周四限号,A1.周四限约,A.周五,A1.周五限号,A1.周五限约, " & _
      "         A.周六,A1.周六限号,A1.周六限约, " & _
      "         A.开始时间,A.终止时间, " & _
      "         D.已挂数,D.已约数,D.其中已接收,to_char(A.停用日期,'yyyy-mm-dd HH24:mi:ss') as 停用日期,a.是否删除" & _
      " From 挂号安排 A,收费项目目录 B,部门表 C," & vbCrLf & _
      "         (" & strTable & ") A1, " & _
      "         (" & strSQL & ") D" & _
      " Where A.ID=A1.ID and a.项目ID = B.ID And a.科室ID = C.ID And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & _
      "            And A.科室ID=D.科室ID(+) and A.项目ID=D.项目ID(+) " & _
      "             And nvl(A.医生姓名,'刘兴洪测试医生')= D.医生姓名(+) And nvl(A.医生id,0)=D.医生ID(+) " & _
      "             And A.号码=D.号码(+)" & str显示条件 & _
      " Order by A.号码"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        Val(mArrFilter("科室ID")), _
        Val(mArrFilter("医生ID")(0)), _
        CStr(mArrFilter("医生ID")(0)), _
        CDate(mArrFilter("有效期")(0)), CDate(mArrFilter("有效期")(1)), UserInfo.ID)
    mblnNotMoveRow = True
    If Not mrsRoom Is Nothing Then
        If mrsRoom.State = 1 Then mrsRoom.Close
    End If
    Set mrsRoom = Nothing
    
    With Me.vsList
        If .Row > 0 Then
            lngPreID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End If
        .Clear 1
        .Rows = 3: lngRow = 2
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 2
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("状态")) = IIf(Val(Nvl(rsTemp!是否删除, 0)) = 1, "已删", IIf(Nvl(rsTemp!停用日期) <> "", "已停", "正常"))
            .TextMatrix(lngRow, .ColIndex("号类")) = Nvl(rsTemp!号类)
            .TextMatrix(lngRow, .ColIndex("号别")) = Nvl(rsTemp!号别)
            .TextMatrix(lngRow, .ColIndex("科室")) = Nvl(rsTemp!科室)
            .TextMatrix(lngRow, .ColIndex("项目")) = Nvl(rsTemp!项目)
            .TextMatrix(lngRow, .ColIndex("医生")) = Nvl(rsTemp!医生)
            .TextMatrix(lngRow, .ColIndex("周日-安排")) = Nvl(rsTemp!周日)
            .TextMatrix(lngRow, .ColIndex("周日-限号")) = Format(Val(Nvl(rsTemp!周日限号)), "###;;")
            If IsNull(rsTemp!周日限约) Then
                .TextMatrix(lngRow, .ColIndex("周日-限约")) = Format(Val(Nvl(rsTemp!周日限约)), "###;;")
            Else
                If Val(rsTemp!周日限约) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("周日-限约")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("周日-限约")) = Format(Val(Nvl(rsTemp!周日限约)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("周一-安排")) = Nvl(rsTemp!周一)
            .TextMatrix(lngRow, .ColIndex("周一-限号")) = Format(Val(Nvl(rsTemp!周一限号)), "###;;")
            If IsNull(rsTemp!周一限约) Then
                .TextMatrix(lngRow, .ColIndex("周一-限约")) = Format(Val(Nvl(rsTemp!周一限约)), "###;;")
            Else
                If Val(rsTemp!周一限约) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("周一-限约")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("周一-限约")) = Format(Val(Nvl(rsTemp!周一限约)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("周二-安排")) = Nvl(rsTemp!周二)
            .TextMatrix(lngRow, .ColIndex("周二-限号")) = Format(Val(Nvl(rsTemp!周二限号)), "###;;")
            If IsNull(rsTemp!周二限约) Then
                .TextMatrix(lngRow, .ColIndex("周二-限约")) = Format(Val(Nvl(rsTemp!周二限约)), "###;;")
            Else
                If Val(rsTemp!周二限约) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("周二-限约")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("周二-限约")) = Format(Val(Nvl(rsTemp!周二限约)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("周三-安排")) = Nvl(rsTemp!周三)
            .TextMatrix(lngRow, .ColIndex("周三-限号")) = Format(Val(Nvl(rsTemp!周三限号)), "###;;")
            If IsNull(rsTemp!周三限约) Then
                .TextMatrix(lngRow, .ColIndex("周三-限约")) = Format(Val(Nvl(rsTemp!周三限约)), "###;;")
            Else
                If Val(rsTemp!周三限约) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("周三-限约")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("周三-限约")) = Format(Val(Nvl(rsTemp!周三限约)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("周四-安排")) = Nvl(rsTemp!周四)
            .TextMatrix(lngRow, .ColIndex("周四-限号")) = Format(Val(Nvl(rsTemp!周四限号)), "###;;")
            If IsNull(rsTemp!周四限约) Then
                .TextMatrix(lngRow, .ColIndex("周四-限约")) = Format(Val(Nvl(rsTemp!周四限约)), "###;;")
            Else
                If Val(rsTemp!周四限约) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("周四-限约")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("周四-限约")) = Format(Val(Nvl(rsTemp!周四限约)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("周五-安排")) = Nvl(rsTemp!周五)
            .TextMatrix(lngRow, .ColIndex("周五-限号")) = Format(Val(Nvl(rsTemp!周五限号)), "###;;")
            If IsNull(rsTemp!周五限约) Then
                .TextMatrix(lngRow, .ColIndex("周五-限约")) = Format(Val(Nvl(rsTemp!周五限约)), "###;;")
            Else
                If Val(rsTemp!周五限约) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("周五-限约")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("周五-限约")) = Format(Val(Nvl(rsTemp!周五限约)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("周六-安排")) = Nvl(rsTemp!周六)
            .TextMatrix(lngRow, .ColIndex("周六-限号")) = Format(Val(Nvl(rsTemp!周六限号)), "###;;")
            If IsNull(rsTemp!周六限约) Then
                .TextMatrix(lngRow, .ColIndex("周六-限约")) = Format(Val(Nvl(rsTemp!周六限约)), "###;;")
            Else
                If Val(rsTemp!周六限约) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("周六-限约")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("周六-限约")) = Format(Val(Nvl(rsTemp!周六限约)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("建档")) = IIf(Val(Nvl(rsTemp!建档)) = 0, "", "√")
            .TextMatrix(lngRow, .ColIndex("分诊方式")) = Nvl(rsTemp!分诊方式)
            .TextMatrix(lngRow, .ColIndex("已挂数")) = Nvl(rsTemp!已挂数)
            .TextMatrix(lngRow, .ColIndex("已约数")) = Nvl(rsTemp!已约数)
            '59611 李南春 2014/04/11 16:40:49 添加一列显示预约已接收数
            .TextMatrix(lngRow, .ColIndex("其中已接收数")) = Nvl(rsTemp!其中已接收)
             
            .TextMatrix(lngRow, .ColIndex("IDS")) = Nvl(rsTemp!科室ID) & "_" & Nvl(rsTemp!项目ID) & "_" & Nvl(rsTemp!医生ID)
            .TextMatrix(lngRow, .ColIndex("应诊诊室")) = Read安排应诊诊室(Val(Nvl(rsTemp!ID)))    ' Nvl(rsTemp!门诊诊室)
            If Not IsNull(rsTemp!开始时间) Then
                .TextMatrix(lngRow, .ColIndex("有效范围")) = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") & _
                    "至" & Format(rsTemp!终止时间, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(lngRow, .ColIndex("有效范围")) = Replace(.TextMatrix(lngRow, .ColIndex("有效范围")), " 00:00:00", "")
            End If
            .TextMatrix(lngRow, .ColIndex("序号控制")) = IIf(Val(Nvl(rsTemp!序号控制)) = 0, "", "√")
            .TextMatrix(lngRow, .ColIndex("停用日期")) = Nvl(rsTemp!停用日期)
            If Trim(.TextMatrix(lngRow, .ColIndex("停用日期"))) <> "" Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
            End If
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "有效号-号别列表", True
        .ColWidth(.ColIndex("标志")) = 285
        If lngPreID <> 0 Then
            lngRow = .FindRow(lngPreID, 0, .ColIndex("ID"), , True)
            If lngRow > 0 And lngRow <= .Rows - 1 Then .Row = lngRow
        End If
        If .Row <= 0 Then
            .Row = 1
        End If
        If .RowIsVisible(.Row) = False Then .TopRow = .Row
        mblnNotMoveRow = False
        '获取明细:
        Call vsList_AfterRowColChange(0, 0, .Row, .Col)
    End With
   Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
     Me.vsList.Redraw = flexRDBuffered
End Sub

Public Property Get zlHaveDatePlan(Optional blnPlanID As Boolean = False) As Boolean
    If Not blnPlanID Then
         zlHaveDatePlan = mblnHaveDate
         Exit Property
    End If
     zlHaveDatePlan = mfrm计划.zlHaveDatPlan
End Property

 

Private Function Read安排应诊诊室(ByVal lngID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定诊室
    '入参:lngID-ID
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-14 22:39:14
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strSQL As String
    
    On Error GoTo errH
    If lngID = 0 Then Exit Function
    
    If mrsRoom Is Nothing Then
        strSQL = "Select 门诊诊室,号表ID From 挂号安排诊室"
        Set mrsRoom = New Recordset
        Call zlDatabase.OpenRecordset(mrsRoom, strSQL, Me.Caption)
    End If
    
    With mrsRoom
        .Filter = "号表ID=" & lngID
        If .RecordCount = 0 Then Exit Function
        
        Do While Not .EOF
            Read安排应诊诊室 = Read安排应诊诊室 & ";" & !门诊诊室
            .MoveNext
        Loop
    End With
    Read安排应诊诊室 = Mid(Read安排应诊诊室, 2)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Read计划应诊诊室(ByVal lng安排ID As Long, ByVal lng计划ID As Long, Optional blnReRead As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定诊室
    '入参:lngID-ID
    '     blnReRead-重新读取
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-14 22:39:14
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strSQL As String
    
    On Error GoTo errH
    If lng安排ID = 0 Then Exit Function
    
    If mrsPlanRoom Is Nothing Or blnReRead Then
        strSQL = "Select 门诊诊室,计划ID From 挂号计划诊室 A,挂号安排计划 B where a.计划id=B.ID and b.安排ID=[1]"
        Set mrsPlanRoom = New Recordset
        Set mrsPlanRoom = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
    End If
    
    With mrsPlanRoom
        .Filter = "计划ID=" & lng计划ID
        If .RecordCount = 0 Then Exit Function
        
        Do While Not .EOF
            Read计划应诊诊室 = Read计划应诊诊室 & ";" & Nvl(!门诊诊室)
            .MoveNext
        Loop
    End With
    Read计划应诊诊室 = Mid(Read计划应诊诊室, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Property Get GetList() As VSFlexGrid
    Set GetList = vsList
End Property
Public Function zlExecuteDeleteList(ByVal bln预约单禁止删除 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除指定有效号别数据
    '入参:bln预约单禁止删除-有预约单,禁止删除
    '返回:删除成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-15 10:38:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intRow As Integer
    Dim strSQL As String, str号别 As String
    Dim intOper As Integer
    With vsList
        Err = 0: On Error GoTo Errhand
        If MsgBox("你确认要删除号别""" & .TextMatrix(.Row, .ColIndex("号别")) & """的安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            
            intRow = .Row
            str号别 = .TextMatrix(.Row, .ColIndex("号别"))
            If CheckExistsBooking(str号别) Then
                '问题:46639
                If bln预约单禁止删除 Then
                    Call MsgBox("该号别存在预约挂号单,不能删除!", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
                If MsgBox("该号别存在预约挂号单,你确实要删除吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
            '修改人王吉
            intOper = CheckExistsRegster(str号别)
            If intOper = 0 Then Exit Function '检查是否该安排下有挂号数据
            If intOper = 1 Then '1为假删除
                strSQL = "zl_挂号安排_Delete(" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & ",0)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            ElseIf intOper = 2 Then '2为真删除
                strSQL = "zl_挂号安排_Delete(" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & ",1)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
            If mArrFilter("显示删除安排") And intOper = 1 Then
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbRed
                .Cell(flexcpText, .Row, .ColIndex("状态")) = "已删"
                .Cell(flexcpText, .Row, .ColIndex("停用日期")) = Format(Date, "yyyy-mm-dd")
                Exit Function
            End If
            If .Rows > 2 Then
                .RemoveItem intRow
            Else
                For i = 0 To .Cols - 1
                    .TextMatrix(intRow, i) = ""
                Next
            End If
            If intRow <= .Rows - 1 Then
                .Row = intRow
            Else
                .Row = .Rows - 1
            End If
            .Col = 0: .ColSel = .Cols - 1
        End If
    End With
    If mblnListSetFocus Then
        zlControl.ControlSetFocus vsList
    Else
        mfrm计划.SetGotFocus (True)
    End If
    zlExecuteDeleteList = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlExecuteModifyList(ByVal frmMain As Form) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改指定有效号别数据
    '返回:修改成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-15 10:38:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str号别 As String, intRow As Integer, lngID As Long
    With vsList
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then Exit Function
        If InStr(1, mstrPrivs, ";安排;") < 0 Then Exit Function
        
        intRow = .Row
        str号别 = .TextMatrix(.Row, .ColIndex("号别"))
        If CheckExistsBooking(str号别) Then
           If MsgBox("该号别存在预约挂号单,你确定要修改吗！", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Function
            End If
        End If
        If frmRegistPlanEdit.ShowEdit(frmMain, edt_修改, mlngModule, mstrPrivs, lngID) = False Then
            Exit Function
        End If
        Call LoadDataToList
        zlControl.ControlSetFocus vsList, True
        zlExecuteModifyList = True
    End With
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CheckExistsBooking(str号别 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定号别是否存在预约挂号单
    '入参:str号别-号别
    '返回:存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-15 10:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Min(发生时间) 时间" & vbNewLine & _
            "From 门诊费用记录" & vbNewLine & _
            "Where 记录性质 = 4 And 记录状态 In (0, 1) And 计算单位 = [1] And 发生时间 > 登记时间"
'    If gint预约天数 = 0 Then
    strSQL = strSQL & " And 发生时间 > Sysdate"
'    Else
'        strSQL = strSQL & " And 发生时间 Between Sysdate And Sysdate+" & gint预约天数
'    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str号别)
    
    CheckExistsBooking = Not IsNull(rsTmp!时间)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Sub zlCallCustomReprot(ByVal frmMain As Form, ByVal lngSys As Long, strReprotName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用相关的自定义报表
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-15 11:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, str号类 As String
    '科室ID_项目ID_医生ID
    With vsList
        varData = Split(.TextMatrix(.Row, .ColIndex("IDS")) & "___", "_")
        str号类 = Trim(.TextMatrix(.Row, .ColIndex("号类")))
        If str号类 <> "" Then
            Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain, _
                "号类=" & str号类, "号别=" & Trim(.TextMatrix(.Row, .ColIndex("号别"))), _
                "科室=" & Val(varData(0)), _
                "项目=" & Val(varData(1)), _
                "医生=" & Val(varData(2)))
        Else
            Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain)
        End If
    End With
End Sub
Public Property Get zlGetListCurrRow() As Variant
    Dim varData As Variant
    Dim cllRow As New Collection
    
    '科室ID_项目ID_医生ID
    With vsList
        varData = Split(.TextMatrix(.Row, .ColIndex("IDS")) & "___", "_")
        cllRow.Add Val(varData(0)), "科室ID"
        cllRow.Add Val(varData(1)), "项目ID"
        cllRow.Add Val(varData(2)), "医生ID"
        cllRow.Add Trim(.TextMatrix(.Row, .ColIndex("号类"))), "号类"
        cllRow.Add Trim(.TextMatrix(.Row, .ColIndex("号别"))), "号别"
    End With
    Set zlGetListCurrRow = cllRow
End Property
Public Property Get zlGet安排ID(Optional blnPlanID As Boolean = False) As Long
    If blnPlanID Then
        zlGet安排ID = mfrm计划.zlGet计划ID: Exit Sub
    End If
    With vsList
        zlGet安排ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
End Property

Public Property Get zlPlanStatus() As Long
   zlPlanStatus = mfrm计划.zlPlanStatus
    '获取计划安排的当前状态
    '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效
End Property

Public Sub zlRefreshData(ByVal ArrFilter As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-15 11:19:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mArrFilter = ArrFilter
    Call LoadDataToList
    Call zlActtion
End Sub

Public Sub zlActtion()
    If mblnListSetFocus = True Then
        On Error Resume Next
        If vsList.Visible And vsList.Enabled Then vsList.SetFocus
    Else
        mfrm计划.zlActtion
    End If
End Sub

Private Sub vsPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent zlPopuMenu(1, Button, Shift, X, Y)
End Sub
Public Sub zlRefreshOlnyPlanData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:仅刷新计划数据
    '编制:刘兴洪
    '日期:2009-09-17 11:28:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng安排ID As Long, str号类 As String, str号别 As String, str科室 As String
    With vsList
        lng安排ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        str号类 = Trim(.TextMatrix(.Row, .ColIndex("号类")))
        str号别 = Trim(.TextMatrix(.Row, .ColIndex("号别")))
        str科室 = Trim(.TextMatrix(.Row, .ColIndex("科室")))
    End With
    Call mfrm计划.zlShowPlan(lng安排ID, str号类, str号别, str科室)
    Call LoadStopPlan(lng安排ID)
End Sub

Private Function CheckIsUserPreRegist(ByVal str号码 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在预约挂号
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-24 09:53:25
    '问题:35959
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select 1 From 门诊费用记录 where 发生时间>=sysdate And 记录状态=0 and  计算单位=[1]  and  记录性质=4 and rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str号码)
    If Not rsTemp.EOF Then
        CheckIsUserPreRegist = True
        Exit Function
    End If

    strSQL = "Select 1 From 病人挂号记录 where 发生时间>=sysdate and 记录性质=1 and 记录状态=1 and  号别=[1]  and rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str号码)
    If Not rsTemp.EOF Then
        CheckIsUserPreRegist = True
        Exit Function
    End If
    CheckIsUserPreRegist = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Public Function zlStopAndResume(ByVal blnStop As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：停用或启用指定有效号别数据
    '返回：停用或启用成功,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-08-17 11:19:18
    '说明：31923
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intRow As Integer
    Dim strSQL As String, str号别 As String
    With vsList
        Err = 0: On Error GoTo Errhand
        If blnStop Then
            '问题:35959
            If CheckIsUserPreRegist(.TextMatrix(.Row, .ColIndex("号别"))) Then
                If MsgBox("注意:" & vbCrLf & "   号别为" & .TextMatrix(.Row, .ColIndex("号别")) & "的安排已经存在预约,是否继续停用? ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        End If
        
        If MsgBox("你确认要" & IIf(blnStop, "停用", "启用") & "号别""" & .TextMatrix(.Row, .ColIndex("号别")) & """的安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            intRow = .Row
            str号别 = .TextMatrix(.Row, .ColIndex("号别"))
            strSQL = "zl_挂号安排_StopAndStart(" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & "," & IIf(blnStop, 1, 0) & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            If blnStop Then
                .TextMatrix(.Row, .ColIndex("停用日期")) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbRed
            Else
                .TextMatrix(.Row, .ColIndex("停用日期")) = ""
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = .ForeColor
            End If
            If mArrFilter("显示停用安排") = 1 Or blnStop = False Then
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = IIf(blnStop, vbRed, vbBlack)
                .Cell(flexcpText, .Row, .ColIndex("状态")) = IIf(blnStop = True, "已停", "正常")
                Exit Function
            End If
            If .Rows > 2 Then
                .RemoveItem intRow
            Else
                For i = 0 To .Cols - 1
                    .TextMatrix(intRow, i) = ""
                Next
            End If
            If intRow <= .Rows - 1 Then
                .Row = intRow
            Else
                .Row = .Rows - 1
            End If
            .Col = 0: .ColSel = .Cols - 1
        End If
    End With
    Call zlActtion
    zlStopAndResume = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlStopPlanTimes() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:停用安排计划
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-08 14:11:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intRow As Integer
    Dim strSQL As String, lngID As Long
    Err = 0: On Error GoTo Errhand
    With vsList
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then Exit Function
    End With
    If frmRegistPlanInvalidation.ShowCard(Me, mlngModule, mstrPrivs, lngID) = False Then
        Call zlActtion: Exit Function
    End If
    Call zlActtion
    zlStopPlanTimes = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlIsStopPlan() As Boolean
    '获取该号别是否已经被停用
    With vsList
        If .Row < 0 Then Exit Function
        zlIsStopPlan = Trim(.TextMatrix(.Row, .ColIndex("停用日期"))) <> ""
    End With
End Function

 

Private Sub vsStop_GotFocus()
    vsStop.BackColorSel = &H8000000D
    mblnListSetFocus = True
End Sub

Private Sub vsStop_LostFocus()
    vsStop.BackColorSel = GRD_LOSTFOCUS_COLORSEL
End Sub
Public Function zlClearStopPlanTimes() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除停用计划安排
    '返回:删除成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-09 15:15:32
    '问题:32504
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, intMsg As VbMsgBoxResult
    Dim int标志 As Integer, lng安排ID As Long
    
    '清除所有停用计划
    intMsg = MsgBox("你是否真的要清除所有挂号安排的停用计划?" & vbCrLf & _
                    "【是】表示删除所有制订好的停用计划。" & vbCrLf & _
                    "【否】表示删除所有已经失效了的停用计划。" & vbCrLf & _
                    "【取消】表示不删除。" & vbCrLf & _
                    " ", vbQuestion + vbYesNoCancel + vbDefaultButton2, gstrSysName)
    If intMsg = vbCancel Then Exit Function
    If intMsg = vbYes Then
        int标志 = 2
    Else
        int标志 = 1
    End If
    On Error GoTo errHandle
    'Zl_挂号安排停用状态_Clearall
    strSQL = "Zl_挂号安排停用状态_Clearall(" & int标志 & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    If int标志 = 2 Then
        vsStop.Clear 1: vsStop.Rows = 2
    Else
        With vsList
            lng安排ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End With
        Call LoadStopPlan(lng安排ID)
    End If
    zlClearStopPlanTimes = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2011-11-14 14:52:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:
    Set mfrm计划 = New frmRegistPlanIntend
    
    Set ObjItem = tbPage.InsertItem(mPgIndex.pg_计划, "安排计划信息", mfrm计划.Hwnd, 0)
    ObjItem.Tag = mPgIndex.pg_计划
    Set ObjItem = tbPage.InsertItem(mPgIndex.pg_时段, "安排时段信息", picTime.Hwnd, 0)
    ObjItem.Tag = mPgIndex.pg_时段
    Set ObjItem = tbPage.InsertItem(mPgIndex.pg_停用, "计划停用信息", picStop.Hwnd, 0)
    ObjItem.Tag = mPgIndex.pg_停用
    If mblnHaveUnit Then
        Set ObjItem = tbPage.InsertItem(mPgIndex.pg_单位, "合作单位安排信息", frmUnitReg.Hwnd, 0)
        ObjItem.Tag = mPgIndex.pg_单位
    End If
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameNone
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub zl_ReLoadUnitReg()
    Dim lng安排ID As Long
    lng安排ID = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("ID")))
    If lng安排ID <= 0 Then Exit Sub
    LoadUnitReg lng安排ID
End Sub

Public Sub UpdatePara(ByVal blnPara As Boolean)
    mfrm计划.blnShowExpired = blnPara
End Sub

Public Property Get zlGet安排停用() As Boolean
    '获取安排是否已经停用
      If vsList.Row >= 0 And vsList.Col >= 0 Then
        zlGet安排停用 = vsList.Cell(flexcpForeColor, vsList.Row, vsList.Col) = vbRed
      End If
End Property

Private Function CheckExistsRegster(str号别 As String) As Byte
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定挂号安排是否存在已挂号数据
    '入参:lng安排ID-安排ID
    '返回:0不操作,1假删除,2真删除
    '编制:王吉
    '日期:2012-03-12 10:10:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
     Dim rsTmp As ADODB.Recordset, strSQL As String, strRS As String, blnRs As Boolean
     strSQL = "Select Nvl(Max(A.已挂号数),0) As 已挂号数 ,Nvl(Max(A.计划数),0) As 计划数 From(" & vbNewLine & _
              "Select 1 As 已挂号数, 0 As 计划数 From 病人挂号汇总 A Where Rownum=1 And A.号码=[1] Having Sum(Nvl(已挂数,0)) > 0 or Sum(Nvl(已约数,0)) > 0" & vbNewLine & _
              "Union All" & vbNewLine & _
              "Select 0 As 已挂号数, 1 As 计划数 From 挂号安排计划 C Where C.号码=[1] Having Count(1)>1) A"
     On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str号别)
     If rsTmp!已挂号数 > 0 And rsTmp!计划数 > 0 Then
        '67824:刘尔旋,2013-11-21,对话框符号不当的问题
        If MsgBox("该安排下已经有挂号数据与计划安排,你确定要删除该安排?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        CheckExistsRegster = 1
        Exit Function
     End If
     If rsTmp!已挂号数 > 0 Then
        If MsgBox("该排班下已经有挂号数据,你确定要删除该安排?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        CheckExistsRegster = 1
        Exit Function
     End If
     If rsTmp!计划数 > 0 Then
        If MsgBox("该排班下已经有计划安排,你确定要删除该安排?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        CheckExistsRegster = 2
        Exit Function
     End If
       If MsgBox("如果删除该排班数据将无法恢复,你确定要删除该安排?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
       CheckExistsRegster = 2
       Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
