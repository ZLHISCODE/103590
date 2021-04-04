VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRegistPlanIntend 
   BorderStyle     =   0  'None
   Caption         =   "安排计划"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTbPage 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   5160
      ScaleHeight     =   2175
      ScaleWidth      =   3255
      TabIndex        =   6
      Top             =   1560
      Width           =   3255
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   2295
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4575
         _Version        =   589884
         _ExtentX        =   8070
         _ExtentY        =   4048
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPlan 
      BorderStyle     =   0  'None
      Height          =   2520
      Left            =   780
      ScaleHeight     =   2520
      ScaleWidth      =   6645
      TabIndex        =   3
      Top             =   3090
      Width           =   6645
      Begin VSFlex8Ctl.VSFlexGrid vsPlan 
         Height          =   2145
         Left            =   -150
         TabIndex        =   4
         Top             =   30
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
         Cols            =   26
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistPlanIntend.frx":0000
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
         Begin VB.PictureBox picImgPlan 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   30
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   5
            Top             =   195
            Width           =   210
            Begin VB.Image imgColPlan 
               Height          =   195
               Left            =   0
               Picture         =   "frmRegistPlanIntend.frx":032F
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin VB.PictureBox picTime 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   1425
      ScaleHeight     =   2745
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      Begin MSComctlLib.TabStrip tbWeekTime 
         Height          =   285
         Left            =   0
         TabIndex        =   1
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
         TabIndex        =   2
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
         FormatString    =   $"frmRegistPlanIntend.frx":087D
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   75
      Top             =   90
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmRegistPlanIntend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mblnNotMoveRow As Boolean '不移动行
Private mrsRoom As ADODB.Recordset
Private mrsPlanRoom As ADODB.Recordset
Private Const conPane_Plan = 1
Private Const conPane_Time = 2
Private Const conPane_Unit = 3
'问题号:51156
Private Const conTbPage = 4
Private mblnShowExpired As Boolean
Private mlng安排ID As Long  '安排ID
Private mstr号类 As String, mstr号别 As String, mstr科室 As String
Public Event zlPopuMenu(intType As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PlanGotFocus(blnPlan As Boolean)
Public Event PlanLostFocus(blnPlan As Boolean)
Private Enum mPgIndex
    pg_计划 = 1
    pg_时段 = 2
    pg_停用 = 3
End Enum
Private mblnNotClick As Boolean
Private mrsTime As ADODB.Recordset
Private mblnHaveDatPlan As Boolean
Private mfrmUnitReg As frmUnitRegPlan
Public mblnSelected计划 As Boolean '问题号:51429

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区哉设置
    '编制:刘兴洪
    '日期:2009-09-14 18:06:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single
    Dim strReg As String
    Dim panThis As Pane
      
    Set panThis = dkpMan.CreatePane(conPane_Plan, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "计划信息"
    panThis.Handle = picPlan.Hwnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Tag = conPane_Plan
     '问题号:51156
    Set panThis = dkpMan.CreatePane(conTbPage, 250, 580, DockRightOf, panThis)
    panThis.Title = "计划时段信息与合作单位计划安排信息"
    panThis.Tag = conTbPage
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption
    panThis.Handle = picTbPage.Hwnd
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    Call InitTbPage
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Sub
'问题号:51156
Private Sub InitTbPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2011-11-14 14:52:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:
    If mfrmUnitReg Is Nothing Then Set mfrmUnitReg = New frmUnitRegPlan
    
    Set ObjItem = tbPage.InsertItem(conPane_Time, "计划时段信息", picTime.Hwnd, 0)
    ObjItem.Tag = conPane_Time
    Set ObjItem = tbPage.InsertItem(conPane_Unit, "合作单位计划安排信息", mfrmUnitReg.Hwnd, 0)
    ObjItem.Tag = conPane_Unit
     With tbPage
         tbPage.Item(0).Selected = True
        .PaintManager.Position = xtpTabPositionBottom
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
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocked Or Action = PaneActionDocking Then Exit Sub
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Plan
        Item.Handle = picPlan.Hwnd
    Case conPane_Time
        Item.Handle = picTime.Hwnd
    Case conTbPage '问题号:51156
        Item.Handle = picTbPage.Hwnd
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
    With vsPlan
        .Redraw = flexRDNone
        .Rows = 3: .FixedRows = 2
        .FixedCols = 1
        .Cols = 40:   .Clear
        .FrozenCols = 2
        .TextMatrix(0, i) = "  ": .ColWidth(i) = 285
        .TextMatrix(1, i) = "  ":  .ColKey(i) = "标志": i = i + 1
        
        .TextMatrix(0, i) = "ID": .ColHidden(i) = True: .ColWidth(i) = 0
        .TextMatrix(1, i) = "ID": .ColKey(i) = "ID": i = i + 1

        .TextMatrix(0, i) = "号类": .ColWidth(i) = 720: .ColHidden(i) = True:
        .TextMatrix(1, i) = "号类": .ColKey(i) = "号类": i = i + 1

        .TextMatrix(0, i) = "号别": .ColWidth(i) = 480: .ColHidden(i) = True:
        .TextMatrix(1, i) = "号别": .ColKey(i) = "号别": i = i + 1

        .TextMatrix(0, i) = "科室": .ColWidth(i) = 1000: .ColHidden(i) = True:
        .TextMatrix(1, i) = "科室": .ColKey(i) = "科室": i = i + 1
        .TextMatrix(0, i) = "项目": .ColWidth(i) = 1000: .ColHidden(i) = True:
        .TextMatrix(1, i) = "项目": .ColKey(i) = "项目": i = i + 1
        .TextMatrix(0, i) = "医生":: .ColWidth(i) = 1000
        .TextMatrix(1, i) = "医生": .ColKey(i) = "医生": i = i + 1
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
        .TextMatrix(0, i) = "生效时间": .ColWidth(i) = 2000
        .TextMatrix(1, i) = "生效时间": .ColKey(i) = "生效时间": i = i + 1
        .TextMatrix(0, i) = "失效时间": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "失效时间": .ColKey(i) = "失效时间": i = i + 1
        .TextMatrix(0, i) = "序号控制": .ColWidth(i) = 765
        .TextMatrix(1, i) = "序号控制": .ColKey(i) = "序号控制": i = i + 1
        
        .TextMatrix(0, i) = "安排人": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "安排人": .ColKey(i) = "安排人": i = i + 1
        .TextMatrix(0, i) = "安排时间": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "安排时间": .ColKey(i) = "安排时间": i = i + 1
        
        .TextMatrix(0, i) = "审核人": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "审核人": .ColKey(i) = "审核人": i = i + 1
        .TextMatrix(0, i) = "审核时间": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "审核时间": .ColKey(i) = "审核时间": i = i + 1
        .TextMatrix(0, i) = "实际执行时间": .ColWidth(i) = 1500
        .TextMatrix(1, i) = "实际执行时间": .ColKey(i) = "实际执行时间": i = i + 1
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
            Case "号类", "号别", "生效时间"
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
    Call InitVsGrid
    Call vsPlan_LostFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "有效号-计划列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub
 

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsPlan, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "有效号-计划列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
    If Not mfrmUnitReg Is Nothing Then Unload mfrmUnitReg: Set mfrmUnitReg = Nothing
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
'问题号:51156
Private Sub picTbPage_Resize()
    tbPage.Move picTbPage.Left, picTbPage.Top, picTbPage.Width, picTbPage.Height
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
        vsTime.Height = .ScaleHeight - vsTime.Top
    End With
End Sub
Private Sub picPlan_Resize()
    Err = 0: On Error Resume Next
    With picPlan
        vsPlan.Left = .ScaleLeft
        vsPlan.Width = .ScaleWidth
        vsPlan.Top = .ScaleTop
        vsPlan.Height = .ScaleHeight
    End With
End Sub

Public Function zlHaveDatPlan() As Boolean
    '*************************************
    '时段是否有数据
    '*************************************
   
    zlHaveDatPlan = mblnHaveDatPlan
End Function

Public Function zlHaveData() As Boolean
    '功能:是否有数据
    With Me.vsPlan
        zlHaveData = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
    End With
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
    If mstr号别 = "" Or mlng安排ID = 0 Then Exit Sub
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "挂号计划表"
    objRow.Add "号类：" & mstr号类
    objRow.Add "号别：" & mstr号别
    objRow.Add "科室：" & mstr科室
 
    objPrint.UnderAppRows.Add objRow
    Set vsGrid = vsPlan
        
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
Public Function zlShowPlan(ByVal lng安排ID As Long, ByVal str号类 As String, ByVal str号别 As String, ByVal str科室 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示计划安排信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-11-15 13:54:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mstr号类 = str号类: mstr号别 = str号别: mstr科室 = str科室: mlng安排ID = lng安排ID
    Call LoadPlan(lng安排ID)
    zlShowPlan = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Sub ReloadTimePlan()
    '***********************************************
    ' 重新加载计划的时间段
    '***********************************************
    Dim lng计划ID As Long, bln序号控制 As Boolean
    With vsPlan
        lng计划ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        bln序号控制 = Trim(.TextMatrix(.Row, .ColIndex("序号控制"))) <> ""
    End With
    Call LoadTimePlan(lng计划ID, bln序号控制, True)
End Sub
Private Sub LoadTimePlan(ByVal lng计划ID As Long, ByVal bln序号控制 As Boolean, Optional blnReload As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载停用计划安排
    '编制:刘兴洪
    '日期:2010-09-09 11:54:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str星期 As String
    Dim i As Long, r As Integer, str时点 As String, strTime As String, strKey As String
    Static lngPre计划Id  As Long
    On Error GoTo errHandle
    '加载该挂号项目的的停用时间信息
    If mrsTime Is Nothing Then
        lngPre计划Id = -1
    ElseIf mrsTime.State <> 1 Then
         lngPre计划Id = -1
    End If
    If lngPre计划Id <> lng计划ID Or blnReload Then
        lngPre计划Id = lng计划ID
        strSQL = "" & _
        "   Select decode(星期,'周日',1,'周一',2,'周二',3,'周三',4,'周四',5,'周五',6,7) as 排序,星期,to_char(开始时间,'HH24')||':00' as 时点,序号,to_char(开始时间,'hh24:mi')||'-' ||to_char(结束时间,'hh24:mi') as 时间范围, " & _
        "               限制数量,是否预约" & _
        "   From  挂号计划时段 " & _
        "   Where 计划ID=[1] And 限制数量>0 " & _
        "   Order by 排序,时点,序号"
        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng计划ID)
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
             r = 0: i = 0
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
 
Private Sub tbWeekTime_Click()
   Dim lng计划ID As Long, bln序号控制 As Boolean
    With vsPlan
        lng计划ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        bln序号控制 = Trim(.TextMatrix(.Row, .ColIndex("序号控制"))) <> ""
    End With
    Call LoadTimePlan(lng计划ID, bln序号控制)
End Sub

Private Sub vsPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng计划ID As Long, bln序号控制 As Boolean
    If OldRow = NewRow Then Exit Sub
    With vsPlan
        lng计划ID = Val(.TextMatrix(NewRow, .ColIndex("ID")))
        bln序号控制 = Trim(.TextMatrix(NewRow, .ColIndex("序号控制"))) <> ""
    End With
    Call LoadTimePlan(lng计划ID, bln序号控制)
    Call LoadUnitPlan(lng计划ID)
    mblnSelected计划 = True '问题号:51429
End Sub

Private Sub LoadUnitPlan(ByVal lng计划ID As Long)
    If mfrmUnitReg Is Nothing Then Exit Sub
    mfrmUnitReg.ShowUnitReg lng计划ID
    
End Sub
Public Sub ReLoadUnitPlan()
    Dim lng计划ID As Long
    
    If mfrmUnitReg Is Nothing Then Exit Sub
    With vsPlan
        lng计划ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
    mfrmUnitReg.ShowUnitReg lng计划ID
    
End Sub
Private Sub vsPlan_BeforeMoveColumn(ByVal Col As Long, Position As Long)
        With vsPlan
            If .ColKey(Col) Like "周*" Then
                 Position = Col
            End If
        End With
End Sub
Private Sub vsPlan_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "有效号-计划列表", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsPlan_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPlan
        If Col = .ColIndex("标志") Then Cancel = True
    End With
End Sub
Public Sub SetGotFocus(blnPlan As Boolean)
    If blnPlan Then vsPlan_GotFocus: Exit Sub
    vsTime_GotFocus
End Sub
Public Sub SetLostFocus(blnPlan As Boolean)
    If blnPlan Then vsPlan_LostFocus: Exit Sub
    vsTime_LostFocus
End Sub
Private Sub vsPlan_GotFocus()
    vsPlan.BackColorSel = &H8000000D
    RaiseEvent PlanGotFocus(True)
End Sub

Private Sub vsPlan_LostFocus()
    vsPlan.BackColorSel = GRD_LOSTFOCUS_COLORSEL
    RaiseEvent PlanLostFocus(True)
End Sub
Private Sub vsTime_GotFocus()
    RaiseEvent PlanGotFocus(False)
End Sub
Private Sub vsTime_LostFocus()
    RaiseEvent PlanLostFocus(False)
End Sub
Private Sub LoadPlan(ByVal lng安排ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载安排计划
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, lngRow As Long
    Dim strSQL As String, lngPreID As Long, strTable As String
    Dim lng限制 As Long
    If lng安排ID = 0 Then
        vsPlan.Clear 1
        vsPlan.Rows = 2: Exit Sub
    End If
    
    Err = 0: On Error GoTo Errhand:
    strTable = "" & _
    "   Select A.ID, " & _
    "             Sum(Decode(B.限制项目,'周日',B.限号数,0)) as 周日限号, Sum(Decode(B.限制项目,'周日',B.限约数))  as 周日限约," & _
    "             Sum(Decode(B.限制项目,'周一',B.限号数,0)) as 周一限号, Sum(Decode(B.限制项目,'周一',B.限约数))  as 周一限约," & _
    "             Sum(Decode(B.限制项目,'周二',B.限号数,0)) as 周二限号, Sum(Decode(B.限制项目,'周二',B.限约数))  as 周二限约," & _
    "             Sum(Decode(B.限制项目,'周三',B.限号数,0)) as 周三限号, Sum(Decode(B.限制项目,'周三',B.限约数))  as 周三限约," & _
    "             Sum(Decode(B.限制项目,'周四',B.限号数,0)) as 周四限号, Sum(Decode(B.限制项目,'周四',B.限约数))  as 周四限约," & _
    "             Sum(Decode(B.限制项目,'周五',B.限号数,0)) as 周五限号, Sum(Decode(B.限制项目,'周五',B.限约数))  as 周五限约," & _
    "             Sum(Decode(B.限制项目,'周六',B.限号数,0)) as 周六限号, Sum(Decode(B.限制项目,'周六',B.限约数))  as 周六限约" & _
    "   From 挂号安排计划 A,挂号计划限制 B  " & _
    "   Where A.ID=B.计划ID(+)  and A.安排ID=[1] And A.生效时间 >= A.安排时间-3/24/60/60" & _
    IIf(mblnShowExpired, "", " And A.失效时间 > Sysdate") & _
    "   Group by A.ID"
    '之所以安排时间减去三秒为了避免,立即生效这种情况
   '38505
    strSQL = " " & _
        "   Select P.*,B.名称 As 项目,D.名称 As 科室 " & _
        "   From ( " & _
        "     Select  row_number()  over (Partition By 计划id Order By 计划id,级数 Desc) As 序号1,M.* " & _
        "     From ( " & _
        "       Select Level As 级数, Sys_Connect_By_Path(门诊诊室, ';') 门诊诊室集, Q.*  " & _
        "       From ( Select  C.Id as 计划ID,C.安排ID ,A.号类,  A.号码,  A.科室id,   Nvl(C.项目id,a.项目ID) as 项目ID, C.医生姓名,  C.医生id,   " & _
        "                              C.周日,C1.周日限号,C1.周日限约,C.周一,C1.周一限号,C1.周一限约,C.周二,C1.周二限号,C1.周二限约, " & _
        "                              C.周三,C1.周三限号,C1.周三限约,C.周四,C1.周四限号,C1.周四限约,C.周五,C1.周五限号,C1.周五限约, " & _
        "                              C.周六,C1.周六限号,C1.周六限约, " & _
        "                              A.病案必须,   Decode(Nvl(C.分诊方式,0),0,'不分诊',1,'指定诊室',2,'动态分诊',3,'平均分诊') as 分诊方式 ,  C.序号控制,             " & _
        "                              to_char(A.开始时间,'yyyy-mm-dd hh24:mi:ss') 开始时间,  to_char(A.终止时间,'yyyy-mm-dd hh24:mi:ss') 终止时间,       " & _
        "                              to_char(C.生效时间,'yyyy-mm-dd hh24:mi:ss') as 生效时间,to_char(C.失效时间,'yyyy-mm-dd hh24:mi:ss') as 失效时间,            " & _
        "                              to_char(C.实际生效,'yyyy-mm-dd hh24:mi:ss') as 实际执行时间, C.安排人,to_char(C.安排时间,'yyyy-mm-dd hh24:mi:ss') as 安排时间,            " & _
        "                              C.审核人,to_char(C.审核时间,'yyyy-mm-dd hh24:mi:ss') as 审核时间 , " & _
        "                               b.门诊诊室,row_number() over (Partition By 计划ID Order By 计划id,门诊诊室) As 序号 " & _
        "           From (" & strTable & ") C1,挂号安排计划 C,挂号安排 A,挂号计划诊室 B " & _
        "           Where  C.ID=C1.ID And C.安排ID =A.Id And C.Id=B.计划ID(+) " & _
        "           Order By 计划ID,门诊诊室 ) Q " & _
        "        Connect By 计划id= Prior 计划id And 序号-1 =Prior 序号 " & _
        "        )  M ) P,收费项目目录 B,部门表 D " & _
        "    Where P.序号1=1 And P.项目id=b.Id And P.科室id =d.Id(+)  " & _
        "    Order By   生效时间 Desc, 计划ID DESC "
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
    With Me.vsPlan
        If .Row > 0 And .Row <= .Rows - 1 Then
            lngPreID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End If
        .Clear 1
        .Rows = 3: .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 2
        lngRow = 2
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!计划Id)
            .Cell(flexcpData, lngRow, .ColIndex("ID")) = Nvl(rsTemp!安排ID)
            .TextMatrix(lngRow, .ColIndex("号类")) = Nvl(rsTemp!号类)
            .TextMatrix(lngRow, .ColIndex("号别")) = Nvl(rsTemp!号码)
            .TextMatrix(lngRow, .ColIndex("科室")) = Nvl(rsTemp!科室)
            .TextMatrix(lngRow, .ColIndex("项目")) = Nvl(rsTemp!项目)
            .TextMatrix(lngRow, .ColIndex("医生")) = Nvl(rsTemp!医生姓名)
            .TextMatrix(lngRow, .ColIndex("周日-安排")) = Nvl(rsTemp!周日)
            .TextMatrix(lngRow, .ColIndex("周日-限号")) = Nvl(rsTemp!周日限号)
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
            .TextMatrix(lngRow, .ColIndex("周一-限号")) = Nvl(rsTemp!周一限号)
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
            .TextMatrix(lngRow, .ColIndex("周二-限号")) = Nvl(rsTemp!周二限号)
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
            .TextMatrix(lngRow, .ColIndex("周三-限号")) = Nvl(rsTemp!周三限号)
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
            .TextMatrix(lngRow, .ColIndex("周四-限号")) = Nvl(rsTemp!周四限号)
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
            .TextMatrix(lngRow, .ColIndex("周五-限号")) = Nvl(rsTemp!周五限号)
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
            .TextMatrix(lngRow, .ColIndex("周六-限号")) = Nvl(rsTemp!周六限号)
            If IsNull(rsTemp!周六限约) Then
                .TextMatrix(lngRow, .ColIndex("周六-限约")) = Format(Val(Nvl(rsTemp!周六限约)), "###;;")
            Else
                If Val(rsTemp!周六限约) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("周六-限约")) = "0"
                Else
                    .TextMatrix(lngRow, .ColIndex("周六-限约")) = Format(Val(Nvl(rsTemp!周六限约)), "###;;")
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("建档")) = IIf(Val(Nvl(rsTemp!病案必须)) = 0, "", "√")
            .TextMatrix(lngRow, .ColIndex("分诊方式")) = Nvl(rsTemp!分诊方式)
            .TextMatrix(lngRow, .ColIndex("IDS")) = Nvl(rsTemp!科室ID) & "_" & Nvl(rsTemp!项目ID) & "_" & Nvl(rsTemp!医生ID)
             
            '*************************************************
            lng限制 = Nvl(Nvl(rsTemp!周日限号, 0)) + lng限制
            lng限制 = Nvl(Nvl(rsTemp!周一限号, 0)) + lng限制
            lng限制 = Nvl(Nvl(rsTemp!周二限号, 0)) + lng限制
            lng限制 = Nvl(Nvl(rsTemp!周三限号, 0)) + lng限制
            lng限制 = Nvl(Nvl(rsTemp!周四限号, 0)) + lng限制
            lng限制 = Nvl(Nvl(rsTemp!周五限号, 0)) + lng限制
            lng限制 = Nvl(Nvl(rsTemp!周六限号, 0)) + lng限制
            '*************************************************
            If Nvl(rsTemp!门诊诊室集) <> "" Then
                .TextMatrix(lngRow, .ColIndex("应诊诊室")) = Mid(Nvl(rsTemp!门诊诊室集), 2)  'Read计划应诊诊室(lng安排ID, Val(Nvl(rsTemp!计划ID)), False) ' Nvl(rsTemp!门诊诊室)
            End If
            
            If Not IsNull(rsTemp!生效时间) Then
                .TextMatrix(lngRow, .ColIndex("生效时间")) = Format(rsTemp!生效时间, "yyyy-MM-dd HH:mm:ss")
                If Format(Nvl(rsTemp!生效时间), "yyyy-MM-dd HH:mm:ss") <= Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And Nvl(rsTemp!审核时间) <> "" Then
                    '已经生效,不能更改
                    .Cell(flexcpData, lngRow, .ColIndex("生效时间")) = 1
                Else
                    '未生效,能更改
                    .Cell(flexcpData, lngRow, .ColIndex("生效时间")) = 0
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("失效时间")) = Nvl(rsTemp!失效时间)
            .TextMatrix(lngRow, .ColIndex("序号控制")) = IIf(Val(Nvl(rsTemp!序号控制)) = 0, "", "√")
            .TextMatrix(lngRow, .ColIndex("安排人")) = Nvl(rsTemp!安排人)
            .TextMatrix(lngRow, .ColIndex("安排时间")) = Nvl(rsTemp!安排时间)
            .TextMatrix(lngRow, .ColIndex("审核人")) = Nvl(rsTemp!审核人)
            .TextMatrix(lngRow, .ColIndex("审核时间")) = Nvl(rsTemp!审核时间)
            If Nvl(rsTemp!实际执行时间) < "3000-01-01" Then
                .TextMatrix(lngRow, .ColIndex("实际执行时间")) = Nvl(rsTemp!实际执行时间)
            End If
            If Val(.Cell(flexcpData, lngRow, .ColIndex("生效时间"))) = 1 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000010
            Else
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
            End If
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        If lngPreID <> 0 Then
            lngRow = .FindRow(lngPreID, 0, .ColIndex("ID"), , True)
            If lngRow > 0 Then .Row = lngRow
        Else
            .Row = 1
        End If
        If .RowIsVisible(.Row) = False Then .TopRow = .Row
        '恢复列设置
         zl_vsGrid_Para_Restore mlngModule, vsPlan, Me.Caption, "有效号-计划列表", True
        .ColWidth(.ColIndex("标志")) = 285
        Call vsPlan_AfterRowColChange(0, 0, .Row, .Col)
        .Redraw = flexRDBuffered
        mblnHaveDatPlan = lng限制 > 0
    End With
   Exit Sub
Errhand:
    vsPlan.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
 
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
Public Property Get zlGet计划ID() As Long
    With vsPlan
        If .Row < 0 Or .Col < 0 Then Exit Sub
        zlGet计划ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
End Property
Public Property Get zlPlanStatus() As Long
    Dim lngID As Long
    '获取计划安排的当前状态
    '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效
    With vsPlan
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then zlPlanStatus = 0: Exit Property
        If .TextMatrix(.Row, .ColIndex("审核时间")) <> "" Then
            zlPlanStatus = 2
            If Val(.Cell(flexcpData, .Row, .ColIndex("生效时间"))) = 1 Then
                zlPlanStatus = 3
            End If
        Else
              zlPlanStatus = 1
        End If
    End With
End Property
Public Sub zlActtion()
    zlControl.ControlSetFocus vsPlan, True
End Sub
Private Sub vsPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent zlPopuMenu(1, Button, Shift, X, Y)
End Sub

Public Property Let blnShowExpired(ByVal vNewValue As Boolean)
    mblnShowExpired = vNewValue
End Property

