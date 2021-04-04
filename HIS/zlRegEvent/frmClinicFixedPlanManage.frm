VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicFixedPlanManage 
   BorderStyle     =   0  'None
   Caption         =   "固定出诊安排管理"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   9720
      MaxLength       =   100
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.PictureBox picPlanDateRange 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   5040
      ScaleHeight     =   285
      ScaleWidth      =   3915
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   3915
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   285
         Left            =   900
         TabIndex        =   8
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   -2147483630
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   169869315
         CurrentDate     =   40777
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   285
         Left            =   2505
         TabIndex        =   9
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   -2147483630
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   169869315
         CurrentDate     =   40777
      End
      Begin VB.Label lblSplit 
         Caption         =   "～"
         Height          =   210
         Left            =   2250
         TabIndex        =   10
         Top             =   30
         Width           =   330
      End
      Begin VB.Label lblDateRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省显示："
         Height          =   180
         Left            =   0
         TabIndex        =   7
         Top             =   45
         Width           =   900
      End
   End
   Begin VB.PictureBox picRegistPlan 
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   5280
      ScaleHeight     =   3645
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   1380
      Width           =   4395
      Begin VSFlex8Ctl.VSFlexGrid vsfRegistPlan 
         Height          =   2445
         Left            =   240
         TabIndex        =   4
         Top             =   150
         Width           =   3495
         _cx             =   6165
         _cy             =   4313
         Appearance      =   2
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
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClinicFixedPlanManage.frx":0000
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
         Begin VB.PictureBox picImgPlan 
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   75
            Picture         =   "frmClinicFixedPlanManage.frx":0075
            ScaleHeight     =   135
            ScaleWidth      =   150
            TabIndex        =   5
            Top             =   90
            Width           =   150
         End
      End
   End
   Begin VB.PictureBox picRegistRule 
      BorderStyle     =   0  'None
      Height          =   4065
      Left            =   180
      ScaleHeight     =   4065
      ScaleWidth      =   4575
      TabIndex        =   2
      Top             =   2220
      Width           =   4575
      Begin VB.Frame fraSplitRule 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   50
         Left            =   2850
         MousePointer    =   7  'Size N S
         TabIndex        =   18
         Top             =   1770
         Width           =   1005
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfRegistRule 
         Height          =   1425
         Left            =   330
         TabIndex        =   14
         Top             =   150
         Width           =   2775
         _cx             =   4895
         _cy             =   2514
         Appearance      =   2
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
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClinicFixedPlanManage.frx":016B
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
         Begin VB.PictureBox picImgRule 
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   75
            Picture         =   "frmClinicFixedPlanManage.frx":01E0
            ScaleHeight     =   135
            ScaleWidth      =   150
            TabIndex        =   15
            Top             =   90
            Width           =   150
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfRegistRuleSub 
         Height          =   1305
         Left            =   0
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
         _cx             =   4048
         _cy             =   2302
         Appearance      =   2
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
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClinicFixedPlanManage.frx":02D6
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
         Begin VB.PictureBox picImgRuleSub 
            BorderStyle     =   0  'None
            Height          =   135
            Left            =   75
            Picture         =   "frmClinicFixedPlanManage.frx":034B
            ScaleHeight     =   135
            ScaleWidth      =   150
            TabIndex        =   17
            Top             =   90
            Width           =   150
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   1575
      Left            =   270
      TabIndex        =   0
      Top             =   570
      Width           =   1305
      _Version        =   589884
      _ExtentX        =   2302
      _ExtentY        =   2778
      _StockProps     =   64
   End
   Begin VB.Label lblValidTimeRange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "有效时间：2016-02-12 00:00:00～3000-01-01 00:00:00"
      Height          =   180
      Left            =   1740
      TabIndex        =   12
      Top             =   660
      Width           =   4500
   End
   Begin VB.Label lblPublishInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发布人：冉俊明  发布时间：2016-01-02 12:32:12"
      Height          =   180
      Left            =   6750
      TabIndex        =   11
      Top             =   210
      Width           =   4050
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   7035
      Left            =   -600
      Top             =   -150
      Width           =   11475
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10845
      _Version        =   589884
      _ExtentX        =   19129
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "出诊安排>固定出诊"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmClinicFixedPlanManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mlngModule As Long
Private mstrPrivs As String

Private Enum mPgIndex '固定出诊TabPage索引
    Pg_出诊规则 = 0
    Pg_出诊安排 = 1
End Enum

Private Enum mPanIndex
    Pan_RegistRuleMain = 0
    Pan_RegistRuleSub = 1
End Enum
Private mlng出诊ID As Long

Private mrsRuleRecords As ADODB.Recordset
Private mrsPlanRecords As ADODB.Recordset
Private mrsRuleRecordsSub As ADODB.Recordset
Private mlngSignalCount As Long '号源总数
Private mintFindType As Integer

Private mlngCopyPlanID As Long, mstrCopyPlanItem As String '用于复制粘贴
Private mdtToday As Date
Private mstrOldSelRangePlan As String '选择网格区域，格式"开始行|结束行|开始列|结束列"
Private mstrOldSelRangeRule As String
Private mstrOldSelRangeRuleSub As String

Private mblnShowInvalidPlan As Boolean '是否显示无效临时安排

Public Sub InitCommVariable(frmParent As Form, cbsMain As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsMain
    
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom

    Err = 0: On Error GoTo errHandler

    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '放在输出到Excel之后
        Set cbrControl = .Find(, conMenu_File_Excel)
'        Set cbrControl = .Add(xtpControlButton,conMenu_File_ExportToXML,"导出为XML文件(&L)…",cbrControl.Index + 1)
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddMonthPlan, "制定月出诊表(&Y)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddWeekPlan, "制定周出诊表(&W)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PlanAdd, "制定临时安排(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddNewSignalSource, "新增号源安排(&A)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyPlanItem, "调整安排(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyUnitRegist, "调整预约挂号控制(&U)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PlanVerify, "审核临时安排(&V)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PlanCancel, "取消临时安排审核(&C)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStartNO, "全部启用序号控制(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStopNO, "全部取消序号控制(&T)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CopyPlan, "复制安排(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PastPlan, "粘贴安排(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearCurPlan, "清除当前安排(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAllPlan, "清除当前号源安排(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearTempPlan, "清除当前临时安排(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAll, "清除所有号源安排(&A)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PublishPlan, "发布安排(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnPublishPlan, "取消发布(&U)")

        '发布后安排调整
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddTempPlan, "临时出诊(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UpdatePlan, "调整安排(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_LockResource, "锁号(&L)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnLockResource, "解锁(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_StopOutCall, "停诊(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnStopOutCall, "取消停诊(&Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_OpenStopPlan, "开放停诊安排(&O)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyDoctor, "替诊(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnModifyDoctor, "取消替诊(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddNumberLimit, "加号(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ReduceNumberLimit, "减号(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyDoctorOffice, "调整分诊诊室(&Z)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UpdateUnitRegist, "调整预约挂号控制(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintPlan, "打印出诊表(&P)"): cbrControl.BeginGroup = True
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowDoctorStopPlan, "显示医生停诊安排(&P)", cbrControl.index)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_PlanChangeInfo, "查询变动信息(&C)", cbrControl.index): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "显示失效临时安排(&S)", cbrControl.index): cbrControl.BeginGroup = True
        cbrControl.Checked = mblnShowInvalidPlan
    End With

    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddMonthPlan, "月出诊表", cbrControl.index + 1)
        cbrControl.ToolTipText = "制定月出诊表"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddWeekPlan, "周出诊表", cbrControl.index + 1)
        cbrControl.ToolTipText = "制定周出诊表"

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyPlanItem, "调整安排", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyUnitRegist, "预约挂号控制", cbrControl.index + 1)
        cbrControl.ToolTipText = "调整预约挂号控制"

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PublishPlan, "发布安排", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnPublishPlan, "取消发布", cbrControl.index + 1)

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_LockResource, "锁号", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnLockResource, "解锁", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_StopOutCall, "停诊", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyDoctor, "替诊", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddNumberLimit, "加号", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ReduceNumberLimit, "减号", cbrControl.index + 1)

        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowDoctorStopPlan, "停诊安排", cbrControl.index + 1): cbrControl.BeginGroup = True
        cbrControl.IconId = conMenu_Edit_StopOutCall
    End With

    Set objPopup = cbrToolBar.Controls.Add(xtpControlButtonPopup, conMenu_View_FindType, "按号码过滤↓")
    objPopup.flags = xtpFlagRightAlign
    '被绑定的控件必须动态加载，因为工具栏一但被删除，被绑定的控件的句柄就会变成0
    Set objCustom = cbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Find, "")
    If txtFind.UBound > 0 Then Unload txtFind(1)
    Load txtFind(1)
    objCustom.Handle = txtFind(1).Hwnd
    objCustom.flags = xtpFlagRightAlign

    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("Y"), conMenu_Edit_AddMonthPlan
        .Add FCONTROL, Asc("W"), conMenu_Edit_AddWeekPlan
        .Add FCONTROL, Asc("M"), conMenu_Edit_ModifyPlanItem

        .Add FCONTROL, Asc("C"), conMenu_Edit_CopyPlan
        .Add FCONTROL, Asc("V"), conMenu_Edit_PastPlan
        .Add 0, VK_DELETE, conMenu_Edit_ClearCurPlan

        .Add FCONTROL, Asc("G"), conMenu_Edit_PublishPlan
        .Add FCONTROL, Asc("I"), conMenu_Edit_UnPublishPlan
    End With

    '设置不常用命令
    '-----------------------------------------------------
    With mcbsMain.Options
'        .AddHiddenCommand conMenu_Edit_Archive
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function PlanIsValid(ByVal vsfGrid As VSFlexGrid) As Boolean
    '判断当前选择安排是否有效
    Dim strCurItem As String
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        strCurItem = .Cell(flexcpData, 0, .Col)
        If IsDate(strCurItem) = False Then Exit Function
        If DateDiff("d", strCurItem, mdtToday) > 0 Then Exit Function
    End With
    PlanIsValid = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsCan临床排班(ByVal vsfGrid As VSFlexGrid) As Boolean
    '判断当前选择号源是否能够临床能够排班
    Dim strCurItem As String
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") Then
            IsCan临床排班 = True
        Else
            IsCan临床排班 = Trim(.TextMatrix(.Row, COL_是否临床排班)) <> ""
        End If
    End With
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim bytActiveGrid As Byte   '当前激活表格
    Dim vsfGrid As VSFlexGrid
    Dim blnEnabled As Boolean
    Dim lng安排ID As Long
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    If Not tbPage.Selected Is Nothing Then
        If tbPage.Selected.index = Pg_出诊规则 Then
            If Me.ActiveControl Is vsfRegistRuleSub Then
                bytActiveGrid = 2
                Set vsfGrid = vsfRegistRuleSub
            Else
                bytActiveGrid = 1
                Set vsfGrid = vsfRegistRule
            End If
        Else
            bytActiveGrid = 3
            Set vsfGrid = vsfRegistPlan
        End If
    End If

    blnEnabled = mlng出诊ID <> 0
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = vsfGrid.Rows > vsfGrid.FixedRows
    Case conMenu_EditPopup
        If mfrmMain.mFunListActived Then
            Control.Visible = HavePrivs(mstrPrivs, "出诊安排;发布安排;取消发布")
        Else
            Control.Visible = ((bytActiveGrid = 1 Or bytActiveGrid = 2) And HavePrivs(mstrPrivs, "出诊安排")) _
                Or (bytActiveGrid = 3 And (HavePrivs(mstrPrivs, "调整安排;临时出诊安排;停诊;替诊;加号;减号;调整分诊诊室;调整预约挂号")))
        End If
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AddMonthPlan, conMenu_Edit_AddWeekPlan '制定月出诊表,制定周出诊表
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible
    Case conMenu_Edit_PublishPlan, conMenu_Edit_UnPublishPlan '发布安排,取消发布
        Control.Visible = mfrmMain.mFunListActived And HavePrivs(mstrPrivs, Decode(Control.ID, _
            conMenu_Edit_PublishPlan, "发布安排", conMenu_Edit_UnPublishPlan, "取消发布"))
        If blnEnabled Then
            If Control.ID = conMenu_Edit_PublishPlan Then
                blnEnabled = Val(lblPublishInfo.Tag) = 0
            Else
                blnEnabled = Val(lblPublishInfo.Tag) = 1
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled

    Case conMenu_Edit_PlanAdd '临时安排
        Control.Visible = HavePrivs(mstrPrivs, "临时出诊安排") And mfrmMain.mFunListActived = False _
            And (bytActiveGrid = 1 Or bytActiveGrid = 2)
        lng安排ID = vsfGrid.TextMatrix(vsfGrid.Row, COL_安排ID)
        If blnEnabled Then blnEnabled = lng安排ID <> 0
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_AddNewSignalSource '新增号源安排
        Control.Visible = HavePrivs(mstrPrivs, "调整安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And (bytActiveGrid = 1 Or bytActiveGrid = 2)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_PlanVerify '临时安排审核
        Control.Visible = HavePrivs(mstrPrivs, "审核临时固定安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And (bytActiveGrid = 1 Or bytActiveGrid = 2)
        lng安排ID = vsfGrid.TextMatrix(vsfGrid.Row, COL_安排ID)
        If blnEnabled Then blnEnabled = lng安排ID <> 0
        If blnEnabled Then blnEnabled = IsVerified(vsfGrid) = False
        Control.Enabled = Control.Visible And blnEnabled
        If bytActiveGrid = 2 Then
            Control.Caption = "审核临时固定安排"
        ElseIf bytActiveGrid = 1 Then
            Control.Caption = "审核新增号源安排"
        End If
    Case conMenu_Edit_PlanCancel '取消临时安排审核
        Control.Visible = HavePrivs(mstrPrivs, "取消临时固定安排审核") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And bytActiveGrid = 2
        lng安排ID = vsfGrid.TextMatrix(vsfGrid.Row, COL_安排ID)
        If blnEnabled Then blnEnabled = lng安排ID <> 0
        If blnEnabled Then blnEnabled = IsTempPlan(vsfGrid)
        If blnEnabled Then blnEnabled = IsVerified(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ModifyPlanItem '调整出诊项
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = vsfGrid.Col >= gPlanGrid_FixedCols
        If blnEnabled Then blnEnabled = (Val(lblPublishInfo.Tag) = 0 Or IsVerified(vsfGrid) = False)
        If blnEnabled Then blnEnabled = IsCan临床排班(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ModifyUnitRegist '调整预约挂号控制
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = Is禁止预约(vsfGrid) = False
        If blnEnabled Then blnEnabled = (Val(lblPublishInfo.Tag) = 0 Or IsVerified(vsfGrid) = False)
        If blnEnabled Then blnEnabled = IsCan临床排班(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_AllStartNO, conMenu_Edit_AllStopNO '全部启用序号控制,全部取消序号控制
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And Val(lblPublishInfo.Tag) = 0 And mfrmMain.mFunListActived = False
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_CopyPlan '复制安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_PastPlan '粘贴安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = (Val(lblPublishInfo.Tag) = 0 Or IsVerified(vsfGrid) = False)
        If blnEnabled Then blnEnabled = IsCan临床排班(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled And mlngCopyPlanID <> 0
    Case conMenu_Edit_ClearCurPlan '清除当前安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = (Val(lblPublishInfo.Tag) = 0 Or IsVerified(vsfGrid) = False)
        If blnEnabled Then blnEnabled = IsCan临床排班(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ClearAllPlan, conMenu_Edit_ClearTempPlan '清除当前号源所有安排,清除当前临时安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived = False
        If Control.Visible Then
            If Control.ID = conMenu_Edit_ClearAllPlan Then
                Control.Visible = bytActiveGrid = 1
            Else
                Control.Visible = bytActiveGrid = 2
            End If
        End If
        If blnEnabled Then blnEnabled = (Val(lblPublishInfo.Tag) = 0 Or IsVerified(vsfGrid) = False)
        If Val(lblPublishInfo.Tag) = 0 And blnEnabled Then blnEnabled = IsCan临床排班(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ClearAll '清除所有号源安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And (bytActiveGrid = 1 Or bytActiveGrid = 2) _
            And mfrmMain.mFunListActived = False And Val(lblPublishInfo.Tag) = 0
        Control.Enabled = Control.Visible And blnEnabled

    '已发布安排调整
    Case conMenu_Edit_LockResource, conMenu_Edit_UnLockResource '锁号,解锁
        Control.Visible = HavePrivs(mstrPrivs, "调整安排") And bytActiveGrid = 3 And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then
            If Control.ID = conMenu_Edit_LockResource Then
                blnEnabled = (PlanIsSelOne(vsfGrid) = False Or PlanIsLocked(vsfGrid) = False)
            Else
                blnEnabled = (PlanIsSelOne(vsfGrid) = False Or PlanIsLocked(vsfGrid))
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 1
    Case conMenu_Edit_AddTempPlan, conMenu_Edit_UpdatePlan '临时出诊,调整发布后的安排
        Control.Visible = HavePrivs(mstrPrivs, "调整安排") And bytActiveGrid = 3 And mfrmMain.mFunListActived = False
        If Control.ID = conMenu_Edit_UpdatePlan And blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfGrid)
        If Control.ID = conMenu_Edit_UpdatePlan And blnEnabled Then blnEnabled = PlanIsValid(vsfGrid)
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 1
    Case conMenu_Edit_StopOutCall, conMenu_Edit_UnStopOutCall, conMenu_Edit_OpenStopPlan '停诊,取消停诊,开放停诊安排
        Control.Visible = HavePrivs(mstrPrivs, "停诊") And bytActiveGrid = 3 And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfGrid)
        If blnEnabled Then
            If Control.ID = conMenu_Edit_StopOutCall Then
                blnEnabled = (PlanIsStopVisit(vsfGrid) = False)
            Else
                blnEnabled = PlanIsStopVisit(vsfGrid)
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 1
    Case conMenu_Edit_ModifyDoctor, conMenu_Edit_UnModifyDoctor '替诊,取消替诊
        Control.Visible = HavePrivs(mstrPrivs, "替诊") And bytActiveGrid = 3 And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsStopVisit(vsfGrid) = False
        If blnEnabled Then
            If Control.ID = conMenu_Edit_ModifyDoctor Then
                blnEnabled = (PlanIsReplaceDoctor(vsfGrid) = False)
            Else
                blnEnabled = PlanIsReplaceDoctor(vsfGrid)
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 1
    Case conMenu_Edit_AddNumberLimit, conMenu_Edit_ReduceNumberLimit, _
        conMenu_Edit_ModifyDoctorOffice, conMenu_Edit_UpdateUnitRegist '加号,减号,调整分诊诊室,调整预约挂号控制
        Control.Visible = HavePrivs(mstrPrivs, Decode(Control.ID, _
            conMenu_Edit_AddNumberLimit, "加号", conMenu_Edit_ReduceNumberLimit, "减号", _
            conMenu_Edit_ModifyDoctorOffice, "调整分诊诊室", conMenu_Edit_UpdateUnitRegist, "调整预约挂号")) _
            And bytActiveGrid = 3 And mfrmMain.mFunListActived = False
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfGrid)
        If blnEnabled Then blnEnabled = PlanIsStopVisit(vsfGrid) = False
        If Control.ID = conMenu_Edit_UpdateUnitRegist And blnEnabled Then blnEnabled = Is禁止预约(vsfGrid) = False
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 1
    Case conMenu_Edit_PrintPlan '    打印出诊表
        Control.Visible = HavePrivs(mstrPrivs, "固定出诊表")
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_View_FindType '查找方式
        Control.Caption = "按" & Decode(mintFindType, 0, "号码", 1, "科室", 2, "医生", "号码") & "过滤↓"
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '查找方式
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
    Case conMenu_View_ShowDoctorStopPlan '显示医生停诊安排
        Control.Visible = mfrmMain.mFunListActived = False
        blnEnabled = False
        If vsfGrid.Row >= vsfGrid.FixedRows Then
            blnEnabled = Trim(vsfGrid.Cell(flexcpData, vsfGrid.Row, COL_医生)) <> ""
        End If
        Control.Enabled = Control.Visible And blnEnabled
    End Select
End Sub

Public Sub InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub

    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "号码(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "科室(&2)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "医生(&3)"
            End If
        End With
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim frmStopVisitAndModifyDoctor As frmClinicPlanStopVisitAndModifyDoctor
    Dim frmOfficeAndUnitRegModify As frmClinicPlanOfficeAndUnitRegModify
    Dim frmNumberLimitModify As frmClinicPlanNumberLimitModify
    Dim frmEdit As frmClinicPlanEdit
    Dim lng记录ID As Long, lng号源Id As Long, lng安排ID As Long, str号码 As String, strItem As String
    Dim obj出诊记录 As 出诊记录, obj出诊号源 As 出诊号源
    Dim bytActiveGrid As Byte '当前激活表格
    Dim strIDs As String, lngRowStart As Long, lngRowEnd As Long, i As Integer
    Dim lngCurCol As Long, str安排IDs As String
    Dim lng出诊ID As Long, strTemp As String
    Dim strDoctorName As String, vsfGrid As VSFlexGrid
    Dim str记录IDs As String

    Err = 0: On Error GoTo errHandler
    If Not tbPage.Selected Is Nothing Then
        If tbPage.Selected.index = Pg_出诊规则 Then
            If Me.ActiveControl Is vsfRegistRuleSub Then
                bytActiveGrid = 2
                Set vsfGrid = vsfRegistRuleSub
            Else
                bytActiveGrid = 1
                Set vsfGrid = vsfRegistRule
            End If
        Else
            bytActiveGrid = 3
            Set vsfGrid = vsfRegistPlan
        End If
        
        With vsfGrid
            lng号源Id = Val(.TextMatrix(.Row, COL_号源ID))
            str号码 = Trim(.TextMatrix(.Row, COL_号码))
            If bytActiveGrid = 3 Then
                '存储了“出诊ID,安排ID”
                strTemp = .Cell(flexcpData, .Row, GetPlanItemNameCol(.Col) + 2)
                If InStr(strTemp, ",") > 0 Then
                    lng出诊ID = Val(Split(strTemp, ",")(0))
                    lng安排ID = Val(Split(strTemp, ",")(1))
                Else
                    lng出诊ID = mlng出诊ID
                    lng安排ID = Val(.TextMatrix(.Row, COL_安排ID))
                End If
            Else
                lng安排ID = Val(.TextMatrix(.Row, COL_安排ID))
            End If
            lng记录ID = Val(.Cell(flexcpData, .Row, GetPlanItemNameCol(.Col)))
            strItem = .Cell(flexcpData, 0, .Col)
            strDoctorName = Trim(.Cell(flexcpData, .Row, COL_医生))
        End With
    End If

    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_ModifyPlanItem '调整安排
        If (bytActiveGrid = 1 Or bytActiveGrid = 2) And (lng号源Id <> 0 Or lng安排ID <> 0) Then
            Set frmEdit = New frmClinicPlanEdit
            If frmEdit.ShowMe(Me, 0, IIf(bytActiveGrid = 1, Fun_Update, Fun_TempPlan), mlng出诊ID, lng号源Id, lng安排ID, strItem, mstrPrivs) Then
                If bytActiveGrid = 1 Then
                    Call RefreshOneData
                Else
                    Call RefreshDataSub
                End If
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ModifyUnitRegist '调整预约挂号控制
        If lng号源Id <> 0 Or lng安排ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            Call frmEdit.ShowMe(Me, 0, Fun_UpdateUnit, mlng出诊ID, lng号源Id, lng安排ID, strItem, mstrPrivs)
        End If
    Case conMenu_Edit_AllStartNO '全部启用序号控制
        If MsgBox("你确定要对当前出诊表的所有限号或限约的号码启用序号控制吗？", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
            Exit Sub
        End If
        Call ZlBatchSNControl(mlng出诊ID, True, IIf(HavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID))
    Case conMenu_Edit_AllStopNO '全部取消序号控制
        If MsgBox("你确定要对当前出诊表的所有限号或限约的号码取消序号控制吗？", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
            Exit Sub
        End If
        Call ZlBatchSNControl(mlng出诊ID, False, IIf(HavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID))
    Case conMenu_Edit_CopyPlan '复制安排
        If lng安排ID <> 0 And strItem <> "" Then
            mlngCopyPlanID = lng安排ID
            mstrCopyPlanItem = strItem
        End If
    Case conMenu_Edit_PastPlan '粘贴安排
        If PastPlan(mlng出诊ID, mlngCopyPlanID, mstrCopyPlanItem) Then
            If bytActiveGrid = 1 Then
                Call RefreshOneData
            Else
                Call RefreshDataSub
            End If
        End If
    Case conMenu_Edit_ClearCurPlan '清除当前安排
        If strItem = "" Then Exit Sub
        If bytActiveGrid = 2 Then '临时安排
            Dim rsFixedRecord As ADODB.Recordset
            With vsfRegistRuleSub
                Set rsFixedRecord = Get预约挂号记录(lng号源Id, _
                    CDate(Format(.TextMatrix(.Row, COL_开始时间), "yyyy-mm-dd")), CDate(Format(.TextMatrix(.Row, COL_终止时间), "yyyy-mm-dd")))
            End With
            If Not rsFixedRecord Is Nothing Then
                Do While Not rsFixedRecord.EOF
                    If Nvl(rsFixedRecord!限制项目) = strItem Then
                        MsgBox "当前号源在该临时安排有效时间范围内的【" & strItem & "】存在预约挂号记录，" & _
                            "【" & strItem & "】的安排在该临时安排中必须固定，不能粘贴！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    rsFixedRecord.MoveNext
                Loop
            End If
        End If
        If MsgBox("你确定要清除号码为【" & str号码 & "】【" & strItem & "】的安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlan(lng安排ID, strItem, False) Then
            If bytActiveGrid = 1 Then
                Call RefreshOneData
            Else
                Call RefreshDataSub
            End If
            If mlngCopyPlanID = lng安排ID And mstrCopyPlanItem = strItem Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ClearAllPlan '清除当前号源所有安排
        If MsgBox("你确定要清除号码为【" & str号码 & "】的所有安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlanBatch(mlng出诊ID, lng号源Id, , , _
            Val(lblPublishInfo.Tag) = 1 And Val(vsfRegistRule.TextMatrix(vsfRegistRule.Row, COL_是否审核)) = 0) Then
            Call RefreshOneData
            If mlngCopyPlanID = lng安排ID Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ClearTempPlan '清除当前临时安排
        If MsgBox("你确定要清除号码为【" & str号码 & "】的当前临时安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlanBatch(mlng出诊ID, lng号源Id, , lng安排ID, True) Then
            If mlngCopyPlanID = lng安排ID Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
            Call RefreshDataSub
        End If
    Case conMenu_Edit_ClearAll '清除所有号源安排
        If MsgBox("你确定要清除所有号源的安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If ZlClearPlanBatch(mlng出诊ID, 0, IIf(HavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID)) Then
            Call RefreshData(mlng出诊ID)
            mlngCopyPlanID = 0: mstrCopyPlanItem = ""
        End If
    Case conMenu_Edit_PublishPlan '发布安排
        If PublishPlan(mlng出诊ID, True) Then
            Call PrintPlan(mlng出诊ID)
            '刷新数据
            Call mfrmMain.NodeChanged("K1_" & mlng出诊ID) '固定出诊表节点：K1_出诊ID
        End If
    Case conMenu_Edit_UnPublishPlan '取消发布
       If PublishPlan(mlng出诊ID, False) Then
            '刷新数据
            Call mfrmMain.NodeChanged("K1_" & mlng出诊ID) '固定出诊表节点：K1_出诊ID
        End If

    Case conMenu_Edit_PlanAdd '临时安排
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, 0, Fun_TempPlan, mlng出诊ID, lng号源Id, , strItem, mstrPrivs) Then
            Call RefreshDataSub '刷新数据
        End If
    '已发布安排调整
    Case conMenu_Edit_AddNewSignalSource '新增号源安排
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, 0, Fun_AddSignalSourcePlan, mlng出诊ID, , , strItem, mstrPrivs) Then
            Call RefreshData(mlng出诊ID)
            With vsfRegistRule
                If .Rows > .FixedRows And .Cols > .FixedCols Then
                    .ShowCell .Rows - 1, .Col '立刻显示到指定单元
                End If
            End With
        End If
    Case conMenu_Edit_PlanVerify '临时安排审核
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, 0, Fun_TempPlanVerify, mlng出诊ID, lng号源Id, lng安排ID, strItem, mstrPrivs) Then
            '刷新出诊规则
            If bytActiveGrid = 1 Then
                Call RefreshOneData
            End If
            Call RefreshDataSub
            '切换页签时重新刷新出诊记录，不使用RefreshOneData()方法刷新是因为可能调整了收费项目，会新多一组数据
            tbPage(Pg_出诊安排).Tag = "0"
        End If
    Case conMenu_Edit_PlanCancel '取消临时安排审核
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, 0, Fun_TempPlanCancel, mlng出诊ID, lng号源Id, lng安排ID, strItem, mstrPrivs) Then
            '刷新出诊规则
            Call RefreshDataSub
            '切换页签时重新刷新出诊记录，不使用RefreshOneData()方法刷新是因为可能调整了收费项目，会少一组数据
            tbPage(Pg_出诊安排).Tag = "0"
        End If
    Case conMenu_Edit_LockResource '锁号
        Call LockPlan(False)
    Case conMenu_Edit_UnLockResource '解锁
        Call LockPlan(True)
    Case conMenu_Edit_AddTempPlan '临时出诊
        Set frmEdit = New frmClinicPlanEdit

        If CheckCanTempVisit(lng安排ID, strItem) = False Then Exit Sub
        If frmEdit.ShowMe(Me, 1, Fun_TempPlanRecord, lng出诊ID, lng号源Id, lng安排ID, strItem, mstrPrivs) Then
            Call RefreshOneData(True)
        End If
    Case conMenu_Edit_UpdatePlan '调整发布后的安排
        Call LockPlanByDay(False, str记录IDs)
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, 1, Fun_UpdatePlan, lng出诊ID, lng号源Id, lng安排ID, strItem, mstrPrivs) Then
            Call LockPlanByDay(True, str记录IDs)
            Call RefreshOneData(True)
        End If
        Call LockPlanByDay(True, str记录IDs)
    Case conMenu_Edit_StopOutCall '停诊
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        If lng记录ID = 0 Then Exit Sub
        'bytFun 功能：1-停诊,2-取消停诊,3-替诊,4-取消替诊
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 1, lng记录ID) Then
            Call RefreshOneData(True)
        End If
    Case conMenu_Edit_UnStopOutCall '取消停诊
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        If lng记录ID = 0 Then Exit Sub
        'bytFun 功能：1-停诊,2-取消停诊,3-替诊,4-取消替诊
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 2, lng记录ID) Then
            Call RefreshOneData(True)
        End If
    Case conMenu_Edit_OpenStopPlan '开放停诊安排
        'zlOpenStopedPlanBySN(ByVal frmMain As Object, ByVal lngModule As Long, _
            Optional ByVal lng记录ID As Long, _
            Optional ByVal lngDeptID As Long, Optional ByVal lngDoctorID As Long) As Boolean
        '功能：对启用了序号控制分时段的已停诊安排按序号开放挂号
        '入参：
        '   frmMain 调用的主窗体
        '   lngModule 调用模块号
        '   lng记录ID 记录ID,1114模块调用时传入
        '   lngDeptID 科室ID
        '   lngDoctorID 医生ID
        '返回：成功返回True，失败返回False
        If lng记录ID <> 0 And Not gobjRegist Is Nothing Then
            gobjRegist.zlOpenStopedPlanBySN Me, mlngModule, lng记录ID
        End If
    Case conMenu_Edit_ModifyDoctor '替诊
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        If lng记录ID = 0 Then Exit Sub
        'bytFun 功能：1-停诊,2-取消停诊,3-替诊,4-取消替诊
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 3, lng记录ID) Then
            Call RefreshOneData(True)
        End If
    Case conMenu_Edit_UnModifyDoctor '取消替诊
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        If lng记录ID = 0 Then Exit Sub
        'bytFun 功能：1-停诊,2-取消停诊,3-替诊,4-取消替诊
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 4, lng记录ID) Then
            Call RefreshOneData(True)
        End If
    Case conMenu_Edit_AddNumberLimit '加号
        If Get出诊记录(lng号源Id, lng记录ID, True, obj出诊号源, obj出诊记录) Then
            Set frmNumberLimitModify = New frmClinicPlanNumberLimitModify
            If frmNumberLimitModify.ShowMe(Me, 1, obj出诊号源, obj出诊记录) Then
                Call RefreshOneData(True)
            End If
        End If
    Case conMenu_Edit_ReduceNumberLimit '减号
        If Get出诊记录(lng号源Id, lng记录ID, True, obj出诊号源, obj出诊记录) Then
            Set frmNumberLimitModify = New frmClinicPlanNumberLimitModify
            If frmNumberLimitModify.ShowMe(Me, 2, obj出诊号源, obj出诊记录) Then
                Call RefreshOneData(True)
            End If
        End If
    Case conMenu_Edit_ModifyDoctorOffice '调整分诊诊室
        If Get出诊记录(lng号源Id, lng记录ID, True, obj出诊号源, obj出诊记录) Then
            Set frmOfficeAndUnitRegModify = New frmClinicPlanOfficeAndUnitRegModify
            Call frmOfficeAndUnitRegModify.ShowMe(Me, 1, obj出诊号源, obj出诊记录, True)
        End If
    Case conMenu_Edit_UpdateUnitRegist '调整预约挂号控制
        If Get出诊记录(lng号源Id, lng记录ID, True, obj出诊号源, obj出诊记录) Then
            Set frmOfficeAndUnitRegModify = New frmClinicPlanOfficeAndUnitRegModify
            Call frmOfficeAndUnitRegModify.ShowMe(Me, 2, obj出诊号源, obj出诊记录, True)
        End If
    Case conMenu_Edit_PrintPlan '    打印出诊表
        Call PrintPlan(mlng出诊ID, 1)
    Case conMenu_View_PlanChangeInfo '查询信息
        Dim frmPlanChangeHistory As New frmClinicPlanChangeHistory
        frmPlanChangeHistory.ShowMe Me, mlngModule
    Case conMenu_View_Refresh
        '104266，可预约天数可能变了，刷新时要重新初始化列
        If tbPage.Selected.index = Pg_出诊安排 Then
            Call RefreshData(mlng出诊ID, True, True)
        Else
            Call RefreshData(mlng出诊ID)
        End If
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 3 '查找方式
        mintFindType = Val(Right(Control.ID, 2)) - 1
        mcbsMain.RecalcLayout
        txtFind(1).Text = ""
        If txtFind(1).Visible And txtFind(1).Enabled Then txtFind(1).SetFocus
    Case conMenu_View_ShowStoped '是否显示无效临时安排
        Control.Checked = Not Control.Checked
        mblnShowInvalidPlan = Control.Checked
        Call zlDatabase.SetPara("显示无效临时安排", IIf(mblnShowInvalidPlan, "1", "0"), glngSys, mlngModule)
        Call RefreshDataSub '刷新数据
    Case conMenu_View_ShowDoctorStopPlan '显示医生停诊安排
        If strDoctorName <> "" Then
            Dim frmDoctorStopVisit As New frmClinicPlanStopVisitManage
            frmDoctorStopVisit.ShowDoctorStopVisit Me, strDoctorName
        End If
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub PrintPlan(ByVal lng出诊ID As Long, Optional ByVal bytMode As Byte)
    '打印出诊表
    '入参：
    '   bytMode 0-发布后打印,1-菜单选择打印
    Err = 0: On Error GoTo errHandler
    If bytMode = 1 Then '防止误操作
        If MsgBox("要打印出诊表吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If

    If gVisitPlan_ModulePara.byt出诊表打印方式 = 1 Or bytMode = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_1", Me, "出诊ID=" & mlng出诊ID, 2)
    ElseIf gVisitPlan_ModulePara.byt出诊表打印方式 = 2 Then
        If MsgBox("要打印出诊表吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_1", Me, "出诊ID=" & mlng出诊ID, 2)
        End If
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecuteFilter(Optional ByVal blnOnlyRefrashRule As Boolean, _
    Optional ByVal blnOnlyRefrashRecord As Boolean)
    '过滤数据
    Dim strKey As String

    Err = 0: On Error GoTo errHandler
    Call zlControl.TxtSelAll(txtFind(1))

    Screen.MousePointer = vbHourglass
    If blnOnlyRefrashRecord = False Then
        If Not mrsRuleRecords Is Nothing Then
            With mrsRuleRecords
                If Trim(txtFind(1).Text) = "" Then
                    .Filter = ""
                Else
                    strKey = Replace(gstrLike, "%", "*") & UCase(txtFind(1).Text) & "*"
                    Select Case mintFindType
                    Case 0   '号码
                        .Filter = "号码 Like '" & strKey & "'"
                    Case 1   '科室(简码)
                        .Filter = "科室 Like '" & strKey & "' Or 科室简码 Like '" & strKey & "'"
                    Case 2   '医生(简码)
                        .Filter = "医生姓名 Like '" & strKey & "' Or 医生简码 Like '" & strKey & "'"
                    Case Else
                        .Filter = ""
                    End Select
                End If
            End With
        End If
        Call LoadPlanDataByRecordset(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecords, 0, mlngSignalCount)
    End If
    
    If blnOnlyRefrashRule = False Then
        If Not mrsPlanRecords Is Nothing Then
            With mrsPlanRecords
                If Trim(txtFind(1).Text) = "" Then
                    .Filter = ""
                Else
                    strKey = Replace(gstrLike, "%", "*") & UCase(txtFind(1).Text) & "*"
                    Select Case mintFindType
                    Case 0   '号码
                        .Filter = "号码 Like '" & strKey & "'"
                    Case 1   '科室(简码)
                        .Filter = "科室 Like '" & strKey & "' Or 科室简码 Like '" & strKey & "'"
                    Case 2   '医生(简码)
                        .Filter = "医生姓名 Like '" & strKey & "' Or 医生简码 Like '" & strKey & "'"
                    Case Else
                        .Filter = ""
                    End Select
                End If
            End With
        End If
        Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, 0, , , Val(lblPublishInfo.Tag) = 1)
    End If
    If mintFindType = 8 Then mintFindType = 0 '清除
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LockPlan(ByVal blnUnlock As Boolean) As Boolean
    '锁定/解锁出诊记录
    '入参：
    '   blnUnlock 是否解锁,True-解锁,False-加锁
    Dim str记录ID As String
    Dim lngRowStart As Long, lngRowEnd As Long '起始行和终止行
    Dim lngColStart As Long, lngColEnd As Long '起始列和终止列
    Dim i As Long, j As Long, lngTemp As Long
    Dim cll记录ID As Collection

    Err = 0: On Error GoTo errHandler
    Set cll记录ID = New Collection
    With vsfRegistPlan
        '选择行范围
        lngRowStart = .Row: lngRowEnd = .RowSel
        If lngRowStart > lngRowEnd Then lngTemp = lngRowStart: lngRowStart = lngRowEnd: lngRowEnd = lngTemp

        '选择列范围
        lngColStart = GetPlanItemNameCol(.Col) '确定"时间段"列
        lngColEnd = GetPlanItemNameCol(.ColSel)
        If lngColStart > lngColEnd Then lngTemp = lngColStart: lngColStart = lngColEnd: lngColEnd = lngTemp

        For i = lngRowStart To lngRowEnd
            For j = lngColStart To lngColEnd Step 3
                If PlanIsLocked(vsfRegistPlan, i, j) = blnUnlock Then
                    If Val(.Cell(flexcpData, i, j)) <> 0 Then
                        If zlStr.ActualLen(str记录ID & "," & Val(.Cell(flexcpData, i, j))) >= 4000 Then
                            cll记录ID.Add Mid(str记录ID, 2)
                            str记录ID = ""
                        End If
                        str记录ID = str记录ID & "," & Val(.Cell(flexcpData, i, j))
                    End If
                End If
            Next
        Next
        If str记录ID <> "" Then
            cll记录ID.Add Mid(str记录ID, 2)
        End If
        If cll记录ID.Count = 0 Then
            MsgBox "当前没有选择需要" & IIf(blnUnlock, "解锁", "锁号") & "的出诊安排！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    gcnOracle.BeginTrans
    For i = 1 To cll记录ID.Count
        If ZlBatchLockPlan(cll记录ID(i), blnUnlock) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    Next
    gcnOracle.CommitTrans
    LockPlan = True

    '刷新界面
    If LockPlan Then
        For i = lngRowStart To lngRowEnd
            Call RefreshOneData(True, i, i = lngRowStart)
        Next
    End If
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LockPlanByDay(ByVal blnUnlock As Boolean, ByRef str记录IDs As String) As Boolean
    '锁定/解锁某一天的出诊记录
    '入参：
    '   blnUnlock 是否解锁,True-解锁,False-加锁
    '说明：
    '   解锁的时候传入加锁时的记录ID
    Dim lngRowStart As Long, lngRowEnd As Long '起始行和终止行
    Dim lngCol  As Long, i As Long

    Err = 0: On Error GoTo errHandler
    If blnUnlock Then
        LockPlanByDay = ZlBatchLockPlan(str记录IDs, True)
        Exit Function
    End If

    With vsfRegistPlan
        '选择行范围
        GetPlanGroupRange vsfRegistPlan, .Row, lngRowStart, lngRowEnd
        lngCol = GetPlanItemNameCol(.Col)  '确定"时间段"列

        For i = lngRowStart To lngRowEnd
            If PlanIsLocked(vsfRegistPlan, i, lngCol) = False Then
                If Val(.Cell(flexcpData, i, lngCol)) <> 0 Then
                    str记录IDs = str记录IDs & "," & Val(.Cell(flexcpData, i, lngCol))
                End If
            End If
        Next
    End With
    If str记录IDs = "" Then Exit Function
    str记录IDs = Mid(str记录IDs, 2)
    
    LockPlanByDay = ZlBatchLockPlan(str记录IDs, False)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckCanTempVisit(ByVal lng安排ID As Long, ByVal strCurDate As String) As Boolean
    '检查当前号源是否可进行临时出诊
    Dim strSQL As String, rsTemp As ADODB.Recordset

    Err = 0: On Error GoTo errHandler
    If lng安排ID = 0 Or IsDate(strCurDate) = False Then Exit Function
    strSQL = "Select 1" & vbNewLine & _
            " From 临床出诊安排 A, 临床出诊号源 B" & vbNewLine & _
            " Where a.号源id = b.Id And a.ID = [1]" & vbNewLine & _
            "       And ([2] Between a.开始时间 And a.终止时间" & vbNewLine & _
            "       Or ([2] > a.终止时间 Or [2] < a.开始时间) And Nvl(b.是否删除, 0) = 0" & vbNewLine & _
            "           And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And b.排班方式 = 0)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查号源是否可临时出诊", lng安排ID, CDate(strCurDate))
    If rsTemp.EOF Then
        MsgBox "该号源可能已被停用或当前日期包含在了其它出诊表中，不能通过当前出诊表进行临时出诊！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckCanTempVisit = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function PublishPlan(ByVal lng出诊ID As Long, ByVal blnPublish As Boolean) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim obj上班时段 As 上班时段
    
    Err = 0: On Error GoTo errHandler
    If MsgBox("你确定要" & IIf(blnPublish, "", "取消") & "发布当前出诊表吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    If blnPublish Then
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊限制 B, 临床出诊表 C" & vbNewLine & _
                " Where a.Id = b.安排id And a.出诊id = c.Id And c.排班方式 = 0 And c.Id = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng出诊ID)
        If rsTemp.EOF Then
            MsgBox "当前出诊表无有效的安排，不能发布！", vbInformation, gstrSysName: Exit Function
        End If
        
        strSQL = "Select Distinct d.号类, d.号码, e.站点, b.限制项目, b.上班时段, To_Char(c.开始时间, 'hh24:mi:ss') As 开始时间, " & vbNewLine & _
                "       To_Char(c.终止时间, 'hh24:mi:ss') As 终止时间" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊限制 B, 临床出诊时段 C, 临床出诊号源 D, 部门表 E" & vbNewLine & _
                " Where a.Id = b.安排id And b.Id = c.限制id And a.号源id = d.Id And d.科室id = e.Id" & vbNewLine & _
                "       And c.序号 = 1 And 出诊id = [1]" & vbNewLine & _
                " Order By " & IIf(gVisitPlan_ModulePara.byt号码比较方式 = 0, "d.号码", "Lpad(d.号码,5,'0')")
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查上班时段与序号分时段时间是否一致", lng出诊ID)
        Do While Not rsTemp.EOF
            Set obj上班时段 = GetWorkTimeRange(Nvl(rsTemp!上班时段), Nvl(rsTemp!站点), Nvl(rsTemp!号类))
            If Format(obj上班时段.开始时间, "hh:mm:00") <> Format(Nvl(rsTemp!开始时间), "hh:mm:00") Then
                If MsgBox("当前出诊表中部分分时段的安排不是根据上班时段的时间进行分段的，如：" & vbCrLf & _
                    "号码为 " & Nvl(rsTemp!号码) & " ，" & Nvl(rsTemp!限制项目) & " " & Nvl(rsTemp!上班时段) & _
                    "[" & Format(obj上班时段.开始时间, "hh:mm") & "-" & Format(obj上班时段.结束时间, "hh:mm") & "]，" & _
                    "第一个序号时段为[" & Format(Nvl(rsTemp!开始时间), "hh:mm") & "-" & Format(Nvl(rsTemp!终止时间), "hh:mm") & "])" & vbCrLf & vbCrLf & _
                    "是否仍要继续发布？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                Exit Do
            End If
            rsTemp.MoveNext
        Loop

        'Zl_临床出诊安排_Publish
        strSQL = "Zl_临床出诊安排_Publish("
        '  Id_In       临床出诊表.Id%Type,
        strSQL = strSQL & "" & lng出诊ID & ","
        '  发布人_In   临床出诊表.发布人%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  发布时间_In 临床出诊表.发布时间%Type := Null,
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
        '  取消发布_In Number:=0
        Screen.MousePointer = vbHourglass
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        PublishPlan = True

        '自动生成临床出诊记录
        'Zl1_Auto_Buildingregisterplan
        '  --功能说明：自动生成临床出诊记录
        '  --          1、根据号源自动生成预约数内的临床出诊记录;
        '  --          2、预约天数的确定:号源预约天数-->预约方式的天数（取最大)-->系统预约天数
        '  --入参:挂号时间_IN:NULL时，自动生成;否则只检查指定日期是否生成了出诊记录没有
        strSQL = "Zl1_Auto_Buildingregisterplan("
        '    挂号时间_In In Date := Null
        strSQL = strSQL & "" & "NULL" & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Screen.MousePointer = vbDefault
    Else
        strSQL = "Select 1" & vbNewLine & _
                " From 病人挂号记录 C, 临床出诊记录 A, 临床出诊安排 B" & vbNewLine & _
                " Where c.出诊记录id = a.Id And a.安排id = b.Id And b.出诊id = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng出诊ID)
        If Not rsTemp.EOF Then
            MsgBox "当前出诊表的安排已被使用，不允许取消发布！", vbInformation, gstrSysName: Exit Function
        End If

        'Zl_临床出诊安排_Publish
        strSQL = "Zl_临床出诊安排_Publish("
        '  Id_In       临床出诊表.Id%Type,
        strSQL = strSQL & "" & lng出诊ID & ","
        '  发布人_In   临床出诊表.发布人%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  发布时间_In 临床出诊表.发布时间%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  取消发布_In Number:=0
        strSQL = strSQL & "" & 1 & ")"
        Screen.MousePointer = vbHourglass
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Screen.MousePointer = vbDefault
    End If
    PublishPlan = True
    Exit Function
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function PastPlan(ByVal lng出诊ID As Long, ByVal lng原安排ID As Long, ByVal str原项目 As String) As Long
    '功能：粘贴安排
    '参数：
    Dim strSQL As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim rsPlan As ADODB.Recordset, rsSignalSource As ADODB.Recordset
    Dim blnTran As Boolean, blnNoPlan As Boolean
    Dim lng安排ID  As Long, lng号源Id As Long, strApplyItem  As String
    Dim dtStart  As Date, dtEnd As Date, str号类 As String
    Dim rsFixedRecord As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lng出诊ID = 0 Then Exit Function
    If lng原安排ID = 0 Then Exit Function
    If str原项目 = "" Then Exit Function

    If tbPage.Selected Is Nothing Then Exit Function
    If tbPage.Selected.index <> Pg_出诊规则 Then Exit Function

    If Me.ActiveControl Is vsfRegistRule Then
        With vsfRegistRule
            lng号源Id = Val(.TextMatrix(.Row, COL_号源ID))
            lng安排ID = Val(.TextMatrix(.Row, COL_安排ID))
            strApplyItem = .Cell(flexcpData, 0, .Col)
            str号类 = Trim(.TextMatrix(.Row, COL_号类))
        End With
    Else
        With vsfRegistRuleSub
            lng号源Id = Val(.TextMatrix(.Row, COL_号源ID))
            lng安排ID = Val(.TextMatrix(.Row, COL_安排ID))
            strApplyItem = .Cell(flexcpData, 0, .Col)
            str号类 = Trim(.TextMatrix(.Row, COL_号类))
            Set rsFixedRecord = Get预约挂号记录(lng号源Id, _
                CDate(Format(.TextMatrix(.Row, COL_开始时间), "yyyy-mm-dd")), CDate(Format(.TextMatrix(.Row, COL_终止时间), "yyyy-mm-dd")))
        End With
    End If

    If lng号源Id = 0 Then Exit Function
    If strApplyItem = "" Then Exit Function
    If lng原安排ID = lng安排ID And str原项目 = strApplyItem Then
        MsgBox "当前安排与复制安排相同，不能粘贴！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not rsFixedRecord Is Nothing Then
        Do While Not rsFixedRecord.EOF
            If Nvl(rsFixedRecord!限制项目) = strApplyItem Then
                MsgBox "当前号源在该临时安排有效时间范围内的【" & strApplyItem & "】存在预约挂号记录，" & _
                    "【" & strApplyItem & "】的安排在该临时安排中必须固定，不能粘贴！", vbInformation, gstrSysName
                Exit Function
            End If
            rsFixedRecord.MoveNext
        Loop
    End If
    
    '检查某个上班时段是否适用于当前号源
    strSQL = "Select c.号类, a.上班时段" & vbNewLine & _
            " From 临床出诊限制 A,临床出诊安排 B, 临床出诊号源 C" & vbNewLine & _
            " Where a.安排ID =b.ID And b.号源ID = c.ID And b.ID = [1] And a.限制项目 = [2] And a.上班时段 Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng原安排ID, str原项目)
    Do While Not rsTemp.EOF
        If GetWorkTimeRange(Nvl(rsTemp!上班时段), gstrNodeNo, str号类) Is Nothing Then
            MsgBox "上班时段“" & Nvl(rsTemp!上班时段) & "”不适用于" & str号类 & "号，不能粘贴！", vbInformation, gstrSysName
            Exit Function
        End If
        rsTemp.MoveNext
    Loop
    
    If Me.ActiveControl Is vsfRegistRule Then
        If Trim(vsfRegistRule.TextMatrix(vsfRegistRule.Row, vsfRegistRule.Col)) <> "" Then
            If MsgBox("被粘贴的星期当前已存在出诊安排，粘贴后这部分安排将会被覆盖！是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    Else
        If Trim(vsfRegistRuleSub.TextMatrix(vsfRegistRuleSub.Row, vsfRegistRuleSub.Col)) <> "" Then
            If MsgBox("被粘贴的星期当前已存在出诊安排，粘贴后这部分安排将会被覆盖！是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If

    If lng安排ID = 0 Then
        '获取安排的时间范围，与原安排一致
        strSQL = "Select 开始时间,终止时间 From 临床出诊安排 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取安排时间范围", lng原安排ID)
        If rsTemp.EOF Then
            dtStart = Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd")
            dtEnd = CDate("3000-01-01")
        Else
            dtStart = Format(Nvl(rsTemp!开始时间, DateAdd("d", 1, zlDatabase.Currentdate)), "yyyy-mm-dd hh:mm:ss")
            dtEnd = Format(Nvl(rsTemp!终止时间, CDate("3000-01-01")), "yyyy-mm-dd hh:mm:ss")
        End If
        blnNoPlan = True

        Set rsSignalSource = GetSignalSource("", lng号源Id)
        If rsSignalSource.EOF Then
            MsgBox "号源信息未找到！", vbInformation, gstrSysName
            Exit Function
        End If

        lng安排ID = zlDatabase.GetNextId("临床出诊安排")
        'Zl_临床出诊安排_Insert(
        strSQL = "Zl_临床出诊安排_Insert("
        'Id_In           临床出诊安排.Id%Type,
        strSQL = strSQL & "" & lng安排ID & ","
        '出诊id_In       临床出诊安排.出诊id%Type,
        strSQL = strSQL & "" & lng出诊ID & ","
        '号源id_In       临床出诊安排.号源id%Type,
        strSQL = strSQL & "" & lng号源Id & ","
        '项目id_In       临床出诊安排.项目id%Type,
        strSQL = strSQL & "" & ZVal(Nvl(rsSignalSource!项目ID)) & ","
        '医生id_In       临床出诊安排.医生id%Type,
        strSQL = strSQL & "" & ZVal(Nvl(rsSignalSource!医生ID)) & ","
        '医生姓名_In     临床出诊安排.医生姓名%Type,
        strTemp = Nvl(rsSignalSource!医生姓名)
        strSQL = strSQL & "" & IIf(strTemp = "", "NULL", "'" & strTemp & "'") & ","
        '排班规则_In     临床出诊安排.排班规则%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '是否周六出诊_In 临床出诊安排.是否周六出诊%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '是否周日出诊_In 临床出诊安排.是否周日出诊%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '开始时间_In     临床出诊安排.开始时间%Type,
        strSQL = strSQL & "" & ZDate(dtStart) & ","
        '终止时间_In     临床出诊安排.终止时间%Type,
        strSQL = strSQL & "" & ZDate(dtEnd) & ","
        '操作员姓名_In   临床出诊安排.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '登记时间_In     临床出诊安排.登记时间%Type
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
    End If

    blnTran = True
    gcnOracle.BeginTrans
        If blnNoPlan And strSQL <> "" Then
            zlDatabase.ExecuteProcedure strSQL, "新增安排"
        End If
        If ZlPlanApplyTo(0, lng原安排ID, str原项目, lng安排ID, strApplyItem) = False Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
    gcnOracle.CommitTrans
    blnTran = False
    PastPlan = True
    Exit Function
errHandler:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub RefreshData(Optional ByVal lng出诊ID As Long, Optional ByVal blnClear As Boolean, _
    Optional ByVal blnLoadRecord As Boolean)
    '功能：刷新安排详情数据
    '入数：
    '   lng出诊ID - 出诊ID
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dtStart As Date, dtEnd As Date
    Dim lngOldRow As Long, lngoldCol As Long

    Err = 0: On Error GoTo errHandler
    If blnLoadRecord Then
        lngOldRow = vsfRegistPlan.Row: lngoldCol = vsfRegistPlan.Col
    End If
    mdtToday = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
    tbPage(Pg_出诊安排).Tag = ""
    If blnClear Then
        mlng出诊ID = lng出诊ID '存储出诊ID
        mlngSignalCount = 0
        Set mrsRuleRecords = Nothing
        Set mrsPlanRecords = Nothing
        Set mrsRuleRecordsSub = Nothing
        mlngCopyPlanID = 0: mstrCopyPlanItem = ""

        sccTitle.Caption = "出诊安排>固定出诊表" & IIf(lng出诊ID = 0, "(无出诊表)", "")
        lblDateRange.Visible = (tbPage.Selected.index = Pg_出诊规则) And lng出诊ID <> 0
        lblDateRange.Caption = "有效时间：" & Format(mdtToday, "yyyy-mm-dd hh:mm:ss") & "～" & "3000-01-01 00:00:00"
        lblPublishInfo.Tag = ""

        '显示时间范围
'        picPlanDateRange.Visible = (tbPage.Selected.index = Pg_出诊安排) And lng出诊ID <> 0
'        dtpStartDate.MaxDate = CDate("3000-01-01")
'        dtpStartDate.MinDate = CDate(Format(mdtToday, "yyyy-mm-dd")): dtpStartDate.MaxDate = CDate(Format(mdtToday, "yyyy-mm-dd")) + 6
'        dtpEndDate.MaxDate = CDate("3000-01-01")
'        dtpEndDate.MinDate = CDate(Format(mdtToday, "yyyy-mm-dd")): dtpEndDate.MaxDate = CDate(Format(mdtToday, "yyyy-mm-dd")) + 6
        dtpStartDate.Value = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
        dtpEndDate.Value = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd")) + Get预约天数(lng出诊ID)

'        dtStart = mdtToday
'        dtEnd = DateAdd("d", mdtToday, 7)

        strSQL = "Select b.出诊表名, b.发布人, b.发布时间 From 临床出诊表 B Where b.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊表信息", mlng出诊ID)
        If Not rsTemp.EOF Then
            sccTitle.Caption = "出诊安排>" & Nvl(rsTemp!出诊表名)
            lblPublishInfo.Tag = IIf(Nvl(rsTemp!发布时间) = "", "", "1") '标记是否发布
            lblPublishInfo.Caption = "发布人：" & IIf(Nvl(rsTemp!发布人) = "", "      ", Nvl(rsTemp!发布人)) & _
                "  发布时间：" & IIf(Nvl(rsTemp!发布时间) = "", "                   ", Format(Nvl(rsTemp!发布时间), "yyyy-mm-dd hh:mm:ss"))
        End If

        '出诊记录
'        If txtPublisher.Caption <> "" Then
'            '缺省显示时间范围
'            strSql = "Select Min(a.出诊日期) As 最小日期, Max(a.出诊日期) As 最大日期" & vbNewLine & _
'                    " From 临床出诊记录 A, 临床出诊安排 B" & vbNewLine & _
'                    " Where a.安排id = b.Id And b.出诊Id = [1]"
'            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取出诊记录时间范围", lng出诊ID)
'            If rsTemp.EOF Then Exit Sub
'
'            dtpStartDate.MaxDate = CDate("3000-01-01")
'            dtpStartDate.MinDate = CDate(Nvl(rsTemp!最小日期, CDate(Format(mdtToday, "yyyy-mm-dd"))))
'            dtpStartDate.MaxDate = CDate(Nvl(rsTemp!最大日期, CDate(Format(mdtToday, "yyyy-mm-dd")) + 6))
'            dtpEndDate.MaxDate = CDate("3000-01-01")
'            dtpEndDate.MinDate = CDate(Nvl(rsTemp!最小日期, CDate(Format(mdtToday, "yyyy-mm-dd"))))
'            dtpEndDate.MaxDate = CDate(Nvl(rsTemp!最大日期, CDate(Format(mdtToday, "yyyy-mm-dd")) + 6))
'
'            dtpStartDate.Value = CDate(Nvl(rsTemp!最小日期, CDate(Format(mdtToday, "yyyy-mm-dd"))))
'            dtpEndDate.Value = CDate(Nvl(rsTemp!最大日期, CDate(Format(mdtToday, "yyyy-mm-dd")) + 6))

'            If DateDiff("d", dtpStartDate.Value, mdtToday) > 0 Then dtpStartDate.Value = CDate(Format(mdtToday, "yyyy-mm-dd"))
'        End If
        
        Call InitPlanGrid(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRule, Me.Name, "规则")
        Call InitPlanGrid(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRuleSub, Me.Name, "规则")
        Call InitPlanGrid(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, dtpStartDate.Value, dtpEndDate.Value, Val(lblPublishInfo.Tag) = 1)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "安排")
        Call ShowHolidayToPlan(vsfRegistPlan, Format(dtpStartDate.Value, "yyyy-mm-dd hh:mm:ss"), Format(dtpEndDate.Value, "yyyy-mm-dd hh:mm:ss"))
    End If
    
    If lng出诊ID <> 0 Then
        '加载数据
        Screen.MousePointer = vbHourglass
        If blnLoadRecord = False Then
            tbPage.Enabled = False
            tbPage(Pg_出诊规则).Selected = True
            tbPage.Enabled = True
        End If
        '规则
        Set mrsRuleRecords = GetPlanRuleData(lng出诊ID, 0, Val(lblPublishInfo.Tag) = 1)
        Call ExecuteFilter(True)
        
        If blnLoadRecord Then
            tbPage(Pg_出诊安排).Tag = "1"
            '出诊记录
            If Val(lblPublishInfo.Tag) = 1 Then
                Set mrsPlanRecords = GetPlanRecords(lng出诊ID, Format(dtpStartDate.Value, "yyyy-mm-dd"), Format(dtpEndDate.Value, "yyyy-mm-dd"))
                Call ExecuteFilter(False, True)
            End If
            '定位上一次行
            With vsfRegistPlan
                If .Rows > .FixedRows And .Cols > .FixedCols Then     '缺省定位行
                    .Row = -1 '保证在选择行不变的情况下也触发RowColChange事件
                    .Row = IIf(lngOldRow < .FixedRows Or lngOldRow > .Rows - 1, IIf(.Rows > .FixedRows, .FixedRows + 1, .FixedRows), lngOldRow)
                    .Col = IIf(lngoldCol = 0 Or lngoldCol > .Cols - 1, .FixedCols, lngoldCol)
                    .ShowCell .Row, .Col  '立刻显示到指定单元
                End If
            End With
        End If
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshOneData(Optional ByVal blnRecord As Boolean, Optional ByVal lngCurRow As Long = -1, _
    Optional ByVal blnReLoadData As Boolean = True)
    '刷新指定行号源数据
    '入参：
    '   blnRecord 是否刷新出诊记录
    Dim lng安排ID As Long, lng号源Id As Long
    Dim strSQL  As String, rsData As ADODB.Recordset
    Dim str收费项目 As String

    Err = 0: On Error GoTo errHandle
    If blnRecord Then
        '1.记录原数据，并获取新数据
        With vsfRegistPlan
            lng号源Id = Val(.TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_号源ID))
            str收费项目 = .TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_项目)
        End With
        
        If blnReLoadData Then
            '跟新本地记录集,主要是更新本次修改的
            Set mrsPlanRecords = GetPlanRecords(mlng出诊ID, Format(dtpStartDate.Value, "yyyy-mm-dd"), Format(dtpEndDate.Value, "yyyy-mm-dd"))
        End If
        
        '2.更新界面
        mrsPlanRecords.Filter = "号源ID=" & lng号源Id & " And 收费项目='" & str收费项目 & "'"
        Call RefreshOnePlanData(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, lngCurRow, _
            Val(lblPublishInfo.Tag) = 1, 0)
        mrsPlanRecords.Filter = ""
    Else
        '1.记录原数据，并获取新数据
        With vsfRegistRule
            lng安排ID = Val(.TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_安排ID))
            lng号源Id = Val(.TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_号源ID))
            str收费项目 = .TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_项目)
        End With
        '如果安排ID为0，表示是生成出诊表后新增的号源，需要重新获取安排ID
        '都要重新获取安排ID，因为在调整安排时，安排ID可能已经变了
    '    If lng安排ID = 0 Then
            strSQL = "Select a.Id, a.开始时间, a.终止时间 From 临床出诊安排 A Where a.出诊id = [1] And a.号源id = [2]"
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "获取安排ID", mlng出诊ID, lng号源Id)
            If Not rsData.EOF Then
                lng安排ID = Val(Nvl(rsData!ID))
            Else
                '若安排ID仍为零，则退出
                '不能退出，要清空当前行
                'Exit Sub
            End If
    '    End If
        
        If blnReLoadData Then
            '更新本地记录集
            Set mrsRuleRecords = GetPlanRuleData(mlng出诊ID, 0, Val(lblPublishInfo.Tag) = 1)
        End If
        
        '2.更新界面
        mrsRuleRecords.Filter = "安排ID=" & lng安排ID & " And 收费项目='" & str收费项目 & "'"
        Call RefreshOnePlanData(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecords, , _
            Val(lblPublishInfo.Tag) = 1, 0)
        mrsRuleRecords.Filter = ""
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim strSQL As String

    Err = 0: On Error GoTo errHandler
'    dtpStartDate.MaxDate = CDate("3000-01-01")
'    dtpStartDate.MinDate = CDate(Format(Now, "yyyy-mm-dd")): dtpStartDate.MaxDate = CDate(Format(Now, "yyyy-mm-dd")) + 6
'    dtpEndDate.MaxDate = CDate("3000-01-01")
'    dtpEndDate.MinDate = CDate(Format(Now, "yyyy-mm-dd")): dtpEndDate.MaxDate = CDate(Format(Now, "yyyy-mm-dd")) + 6
'    dtpStartDate.Value = CDate(Format(Now, "yyyy-mm-dd"))
'    dtpEndDate.Value = CDate(Format(Now, "yyyy-mm-dd")) + 6
    
    
    mblnShowInvalidPlan = Val(zlDatabase.GetPara("显示无效临时安排", glngSys, mlngModule, "0")) = 1
    
    Call InitPage

    Dim strFindType As String
    Call GetRegInFor(g私有模块, Me.Name, "FindType", strFindType)
    mintFindType = Val(strFindType)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    lblPublishInfo.Move sccTitle.Width - lblPublishInfo.Width - 100, sccTitle.Top + sccTitle.Height - lblPublishInfo.Height - 50

    With tbPage
        .Left = sccTitle.Left
        .Top = sccTitle.Top + sccTitle.Height
        .Width = sccTitle.Width
        .Height = shpBorder.Height - .Top - 10
    End With
    lblDateRange.Move 2500, tbPage.Top + 40, Me.ScaleWidth - lblDateRange.Left - 16
End Sub

Private Sub InitPage()
    '功能:初始化页面控件
    Dim i As Long, ObjItem As TabControlItem

    Err = 0: On Error GoTo errHandler
    tbPage.RemoveAll
    tbPage.InsertItem mPgIndex.Pg_出诊规则, "出诊规则", picRegistRule.Hwnd, 0
    tbPage.InsertItem mPgIndex.Pg_出诊安排, "出诊安排", picRegistPlan.Hwnd, 0

     With tbPage.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003 '显示风格
        .BoldSelected = True '显示页标题字体加粗
        .ClientFrame = xtpTabFrameSingleLine '页面边框
        .Layout = xtpTabLayoutAutoSize
    End With
    tbPage.Enabled = False
    tbPage.Item(Pg_出诊安排).Selected = True
    tbPage.Item(Pg_出诊规则).Selected = True
    tbPage.Enabled = True
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsRuleRecords = Nothing
    Set mrsPlanRecords = Nothing
    Set mrsRuleRecordsSub = Nothing

    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistRule, Me.Name, "规则")
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistPlan, Me.Name, "安排")
    Call SaveRegInFor(g私有模块, Me.Name, "FindType", mintFindType)
End Sub

Private Sub picImgRule_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT

    vRect = zlControl.GetControlRect(picImgRule.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgRule.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfRegistRule, lngLeft, lngTop, picImgRule.Height)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistRule, Me.Name, "规则")
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRuleSub, Me.Name, "规则")
End Sub

Private Sub picImgRuleSub_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT

    vRect = zlControl.GetControlRect(picImgRuleSub.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgRuleSub.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfRegistRuleSub, lngLeft, lngTop, picImgRuleSub.Height)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistRuleSub, Me.Name, "规则")
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRule, Me.Name, "规则")
End Sub

Private Sub picRegistPlan_GotFocus()
    On Error Resume Next
    If vsfRegistPlan.Visible And vsfRegistPlan.Enabled Then vsfRegistPlan.SetFocus
End Sub

Private Sub picRegistPlan_Resize()
    On Error Resume Next
    vsfRegistPlan.Move -10, 0, picRegistPlan.ScaleWidth + 20, picRegistPlan.ScaleHeight
End Sub

Private Sub picRegistRule_GotFocus()
    On Error Resume Next
    If vsfRegistRule.Visible And vsfRegistRule.Enabled Then vsfRegistRule.SetFocus
End Sub

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If tbPage.Selected Is Nothing Then
        If vsfRegistRule.Visible And vsfRegistRule.Enabled Then vsfRegistRule.SetFocus
    Else
        If tbPage.Selected.index = Pg_出诊规则 Then
            If vsfRegistRule.Visible And vsfRegistRule.Enabled Then vsfRegistRule.SetFocus
        Else
            If vsfRegistPlan.Visible And vsfRegistPlan.Enabled Then vsfRegistPlan.SetFocus
        End If
    End If
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Err = 0: On Error GoTo errHandler
    If tbPage.ItemCount < 2 Then Exit Sub
    If tbPage.Tag = Item.Caption Then Exit Sub
    tbPage.Tag = Item.Caption
    lblDateRange.Visible = (Item.index = Pg_出诊规则)

    If Item.index = Pg_出诊安排 And Val(tbPage(Pg_出诊安排).Tag) = 0 Then
        '出诊记录
        If Val(lblPublishInfo.Tag) = 1 Then
            Screen.MousePointer = vbHourglass
            Set mrsPlanRecords = GetPlanRecords(mlng出诊ID, Format(dtpStartDate.Value, "yyyy-mm-dd"), Format(dtpEndDate.Value, "yyyy-mm-dd"))
            Call ExecuteFilter(False, True)
            Screen.MousePointer = vbDefault
        End If
        tbPage(Pg_出诊安排).Tag = "1"
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlDataPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrSysName = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    Dim vsfTemp As VSFlexGrid

    Err = 0: On Error GoTo errHandler
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objOut)
        If bytMode = 0 Then Exit Sub
    End If
    
    objOut.Title.Text = Mid(sccTitle.Caption, InStr(sccTitle.Caption, ">") + 1) & IIf(tbPage.Selected.index = Pg_出诊规则, "规则", "记录") & "清单"
    If VSFlexGridCopyTo(IIf(tbPage.Selected.index = Pg_出诊规则, vsfRegistRule, vsfRegistPlan), _
        vsfTemp, bytMode) = False Then Exit Sub
    If vsfTemp Is Nothing Then Exit Sub
    vsfTemp.ColWidth(COL_图标) = 0 '隐藏图标列
    Set objOut.Body = vsfTemp

    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow

    zlPrintOrView1Grd objOut, bytMode
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng记录ID As Long
    Dim strTemp As String

    On Error Resume Next
    '显示停诊信息
    lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, NewRow, GetPlanItemNameCol(NewCol)))
    If lng记录ID = 0 Or mrsPlanRecords Is Nothing Then
        strTemp = ""
    Else
        Dim strFilter As Variant
        strFilter = mrsPlanRecords.Filter
        mrsPlanRecords.Filter = "记录ID=" & lng记录ID
        If mrsPlanRecords.EOF Then
            strTemp = ""
        Else
            If Nvl(mrsPlanRecords!停诊开始时间) = "" Then
                strTemp = ""
            Else
                strTemp = Nvl(mrsPlanRecords!上班时段) & _
                " 停诊时间：" & Format(Nvl(mrsPlanRecords!停诊开始时间), "mm-dd hh:mm") & _
                "～" & Format(Nvl(mrsPlanRecords!停诊终止时间), "mm-dd hh:mm") & "，停诊原因：" & Nvl(mrsPlanRecords!停诊原因)
            End If
        End If
        mrsPlanRecords.Filter = strFilter
    End If
    Call mfrmMain.StatusShowInfoChanged(2, strTemp)
End Sub

Private Sub vsfRegistPlan_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    On Error Resume Next
    Call SetPlanGridRangeColor(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mstrOldSelRangePlan)
    mstrOldSelRangePlan = vsfRegistPlan.Row & "|" & vsfRegistPlan.RowSel & "|" & vsfRegistPlan.Col & "|" & vsfRegistPlan.ColSel
End Sub

Private Sub vsfRegistPlan_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistPlan, Me.Name, "安排")
End Sub

Private Sub vsfRegistPlan_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    If Val(vsfRegistPlan.RowData(NewRow)) = -1 Then Cancel = True
End Sub

Private Sub vsfRegistPlan_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = gPlanGrid_ColIndex.COL_图标 Then Cancel = True: Exit Sub
End Sub

Private Sub vsfRegistPlan_DblClick()
    Dim lng号源Id As Long, lng出诊ID As Long, lng安排ID As Long
    Dim frmEdit As New frmClinicPlanEdit
    Dim strCurItem As String, strTemp As String
    Dim lngCol As Long, lngRow As Long
    Dim strSort As String

    Err = 0: On Error GoTo errHandler
    lngCol = vsfRegistPlan.MouseCol
    lngRow = vsfRegistPlan.MouseRow
    If lngRow = 0 Or lngRow = 1 Then
        '排序
        If mrsPlanRecords Is Nothing Then Exit Sub
        If mrsPlanRecords.RecordCount = 0 Then Exit Sub
        strSort = GetPlanSortCircleStr(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, lngRow, lngCol)
        If strSort <> "" Then
            mrsPlanRecords.Sort = strSort
            Screen.MousePointer = vbHourglass
            Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, 0, , True, Val(lblPublishInfo.Tag) = 1)
            Screen.MousePointer = vbDefault
        End If
    Else
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        lngCol = GetPlanItemNameCol(vsfRegistPlan.Col)
        '存储了“出诊ID,安排ID”
        strTemp = vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, lngCol + 2)
        If InStr(strTemp, ",") > 0 Then
            lng出诊ID = Val(Split(strTemp, ",")(0))
            lng安排ID = Val(Split(strTemp, ",")(1))
        Else
            lng出诊ID = mlng出诊ID
            lng安排ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_安排ID))
        End If
        strCurItem = vsfRegistPlan.Cell(flexcpData, 0, lngCol)
        If lng号源Id = 0 And lng安排ID = 0 Then Exit Sub
        If IsDate(strCurItem) = False Then Exit Sub
        If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) = "" Then Exit Sub

        Call frmEdit.ShowMe(Me, 1, Fun_View, lng出诊ID, lng号源Id, lng安排ID, strCurItem, mstrPrivs)
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistPlan_GotFocus()
    Call SetSelectedBackColor(vsfRegistPlan, True)
End Sub

Private Sub vsfRegistPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RegistPlan_KeyDown(vsfRegistPlan, KeyCode, Shift)
End Sub

Private Sub vsfRegistPlan_LostFocus()
    Call SetSelectedBackColor(vsfRegistPlan, False)
End Sub

Private Sub vsfRegistPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbCommandBar As CommandBar

    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)

    Set cbCommandBar = GetPopupCommandBar(Me, mcbsMain)
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub
    
    cbCommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistRule_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If NewRow < 2 Then
        Call mfrmMain.StatusShowInfoChanged(2, "")
    Else
        Call mfrmMain.StatusShowInfoChanged(2, "当前共" & mlngSignalCount & "条号源，当前号源：" & vsfRegistRule.TextMatrix(NewRow, COL_号码) & _
            "，开始时间：" & vsfRegistRule.TextMatrix(NewRow, COL_开始时间) & "，终止时间：" & vsfRegistRule.TextMatrix(NewRow, COL_终止时间) & "")
    End If
End Sub

Private Sub vsfRegistRule_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    On Error Resume Next
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    vsfRegistRuleSub.LeftCol = NewLeftCol
End Sub

Private Sub vsfRegistRule_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    On Error Resume Next
    Call SetPlanGridRangeColor(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule, mstrOldSelRangeRule)
    mstrOldSelRangeRule = vsfRegistRule.Row & "|" & vsfRegistRule.RowSel & "|" & vsfRegistRule.Col & "|" & vsfRegistRule.ColSel
End Sub

Private Sub vsfRegistRule_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistRule, Me.Name, "规则")
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRuleSub, Me.Name, "规则")
End Sub

Private Sub vsfRegistRule_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    If Val(vsfRegistRule.RowData(NewRow)) = -1 Then Cancel = True
End Sub

Private Sub vsfRegistRule_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = gPlanGrid_ColIndex.COL_图标 Then Cancel = True: Exit Sub
End Sub

Private Sub vsfRegistRule_DblClick()
    Dim lng号源Id As Long, lng安排ID As Long
    Dim frmEdit As New frmClinicPlanEdit
    Dim strCurItem As String, blnUpdate As Boolean
    Dim lngCol As Long, lngRow As Long
    Dim strSort As String

    Err = 0: On Error GoTo errHandler
    lngCol = vsfRegistRule.MouseCol
    lngRow = vsfRegistRule.MouseRow
    If lngRow = 0 Or lngRow = 1 Then
        '排序
        If Not mrsRuleRecords Is Nothing Then
            strSort = GetPlanSortCircleStr(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule, lngRow, lngCol)
            If strSort <> "" Then
                mrsRuleRecords.Sort = strSort
                Screen.MousePointer = vbHourglass
                Call LoadPlanDataByRecordset(vsfRegistRule, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecords, 0, , True)
                Screen.MousePointer = vbDefault
            End If
        End If
    Else
        With vsfRegistRule
            lng号源Id = Val(.TextMatrix(.Row, COL_号源ID))
            lng安排ID = Val(.TextMatrix(.Row, COL_安排ID))
            strCurItem = .Cell(flexcpData, 0, .Col)
            If lng号源Id = 0 And lng安排ID = 0 Then Exit Sub
            If strCurItem = "" Then Exit Sub
            
            blnUpdate = zlStr.IsHavePrivs(mstrPrivs, "出诊安排") _
                And (Val(lblPublishInfo.Tag) = 0 Or Val(.TextMatrix(.Row, COL_是否审核)) = 0)
            If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
                '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
                If Trim(.TextMatrix(.Row, COL_是否临床排班)) = "" Then blnUpdate = False
            End If
    
            If frmEdit.ShowMe(Me, 0, IIf(blnUpdate, Fun_Update, Fun_View), mlng出诊ID, lng号源Id, lng安排ID, strCurItem, mstrPrivs) Then
                If blnUpdate Then Call RefreshOneData
            End If
        End With
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistRule_EnterCell()
    Dim lng安排ID As Long

    Err = 0: On Error GoTo errHandler
    If Val(vsfRegistRule.Tag) = vsfRegistRule.Row Then Exit Sub
    vsfRegistRule.Tag = vsfRegistRule.Row
    lng安排ID = Val(vsfRegistRule.TextMatrix(vsfRegistRule.Row, COL_安排ID))
    LoadPlanDataSub mlng出诊ID, lng安排ID
    If vsfRegistRule.Visible And vsfRegistRule.Enabled Then vsfRegistRule.SetFocus
    Call SetSelectedBackColor(vsfRegistRuleSub, False)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistRule_GotFocus()
    Call SetSelectedBackColor(vsfRegistRule, True)
End Sub

Private Sub vsfRegistRule_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RegistPlan_KeyDown(vsfRegistRule, KeyCode, Shift)
End Sub

Private Sub vsfRegistRule_LostFocus()
    Call SetSelectedBackColor(vsfRegistRule, False)
End Sub

Private Sub vsfRegistRule_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbCommandBar As CommandBar

    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)

    Set cbCommandBar = GetPopupCommandBar(Me, mcbsMain)
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub

    cbCommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistPlan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim varData As Variant, strTemp As String
    Dim lngRow As Long, lngCol As Long

    On Error GoTo errHandler
    With vsfRegistPlan
        If Not .Visible Then Call mfrmMain.StatusShowInfoChanged(3, ""): Exit Sub

        lngRow = .MouseRow: lngCol = .MouseCol
        If .Tag = lngRow & "," & lngCol Then Exit Sub
        .Tag = lngRow & "," & lngCol

        If lngRow < .FixedRows Or lngCol < gPlanGrid_FixedCols Then Call mfrmMain.StatusShowInfoChanged(3, ""): Exit Sub
        If (lngCol - gPlanGrid_FixedCols) Mod 3 = 0 Then Call mfrmMain.StatusShowInfoChanged(3, ""): Exit Sub '"时段"列退出
        strTemp = Trim(.TextMatrix(lngRow, lngCol))
        If (strTemp = "" Or InStr(strTemp, "/") = 0) And strTemp <> "-" Then Call mfrmMain.StatusShowInfoChanged(3, ""): Exit Sub

        '2.显示内容
        If strTemp = "-" Then
            strTemp = "禁止预约！"
        Else
            varData = Split(strTemp, "/")
            If (lngCol - gPlanGrid_FixedCols) Mod 3 = 1 Then
                strTemp = "限号数:" & IIf(Trim(varData(1)) = "∞", "不限制", Trim(varData(1))) & ", 其中已挂数:" & Trim(varData(0))
            ElseIf (lngCol - gPlanGrid_FixedCols) Mod 3 = 2 Then
                strTemp = "限约数:" & IIf(Trim(varData(1)) = "∞", "不限制", Trim(varData(1))) & ", 其中已约数:" & Trim(varData(0))
            End If
        End If
        Call mfrmMain.StatusShowInfoChanged(3, strTemp)
    End With
    Exit Sub
errHandler:
    Err.Clear
    Call mfrmMain.StatusShowInfoChanged(3, "")
End Sub

Private Sub picImgPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT

    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfRegistPlan, lngLeft, lngTop, picImgPlan.Height)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistPlan, Me.Name, "安排")
End Sub

Private Function GetPlanRuleData(ByVal lng出诊ID As Long, Optional ByVal lng安排ID As Long, _
    Optional ByVal blnPublished As Boolean) As ADODB.Recordset
    '功能：获取安排规则
    '入数：周出诊表
    '   lng出诊ID   - 出诊ID
    '   blnPublished- 是否已发布
    Dim strSQL As String, strColSub As String
    Dim strWhere As String, str是否有效 As String
    Dim str排序号码 As String

    Err = 0: On Error GoTo errHandler
    str排序号码 = IIf(gVisitPlan_ModulePara.byt号码比较方式 = 0, "a.号码", "Lpad(a.号码,5,'0')")
    strColSub = "       " & str排序号码 & " As 排序号码, a.Id As 号源id, a.号类, a.号码, Nvl(a.是否建病案, 0) As 是否建病案,a.预约天数, a.出诊频次," & vbNewLine & _
                "       Decode(a.假日控制状态, 1, '开放预约', 2, '禁止预约', 3, '受节假日设置控制', '不上班') As 假日控制状态," & vbNewLine & _
                "       Nvl(a.是否临床排班, 0) As 是否临床排班, Decode(a.排班方式, 1, '按月排班', 2, '按周排班', '固定排班') As 排班方式," & vbNewLine & _
                "       f.名称 As 科室, f.简码 As 科室简码, Nvl(a.是否假日换休, 0) As 是否假日换休," & vbNewLine

    '没有"所有科室"权限的操作员只能操作自己所属科室的号源
    If HavePrivs(mstrPrivs, "所有科室") = False Then
        strWhere = "      And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = [3])"
    End If
    
    '有效安排，以下规则必须同时满足：
    '    --1.已审核
    '    --2.终止时间大于当前时间
    '    --3.后面无其它临时安排或者有其它安排但时间范围没有被覆盖
    '    --4.没有调整为其它排班方式，或者调整为其它排班方式但还没有出诊安排
    '说明：由多个临时安排一起覆盖的安排，判断不了
    str是否有效 = "Nvl((Select 1" & vbNewLine & _
                " From Dual" & vbNewLine & _
                " Where b.审核时间 Is Not Null And b.终止时间 > Sysdate" & vbNewLine & _
                "       And Not Exists(Select 1" & vbNewLine & _
                "        From 临床出诊安排" & vbNewLine & _
                "        Where 审核时间 Is Not Null And 号源id = b.号源id And 登记时间 > b.登记时间" & vbNewLine & _
                "              And (Nvl(b.是否临时安排, 0) = 0 And Nvl(是否临时安排, 0) = 0 Or Nvl(b.是否临时安排, 0) = 1)" & vbNewLine & _
                "              And Decode(Sign(Sysdate - 开始时间), 1, Sysdate, 开始时间) <= Decode(Sign(Sysdate - b.开始时间), 1, Sysdate, b.开始时间)" & vbNewLine & _
                "              And 终止时间 >= b.终止时间)" & vbNewLine & _
                "       And Not Exists(Select 1" & vbNewLine & _
                "           From 临床出诊安排 P, 临床出诊表 Q" & vbNewLine & _
                "           Where p.出诊id = q.Id And p.号源id = b.号源id And Nvl(q.排班方式, 0) In (1, 2)And p.开始时间 < Sysdate)" & vbNewLine & _
                "   ), 0) As 是否有效,"
    
    If lng安排ID = 0 Then
        strSQL = "Select m.Id As 安排ID, m.出诊ID, m.号源id, m.项目ID, m.医生ID, m.医生姓名, m.排班规则," & _
                "        m.开始时间, m.终止时间, m.登记时间,m.审核时间,m.是否临时安排" & vbNewLine & _
                " From 临床出诊安排 M" & vbNewLine & _
                " Where m.出诊id = [1] And Nvl(m.是否临时安排, 0) = 0"
        If blnPublished Then
            strSQL = "Select " & str是否有效 & _
                    "       b.出诊ID, b.安排id, e.名称 As 收费项目, b.医生姓名, g.简码 As 医生简码, g.专业技术职务 as 医生职称,i.标识符," & vbNewLine & _
                    "       b.排班规则, b.开始时间, b.终止时间, b.登记时间, b.是否临时安排 As 临时安排, Decode(b.审核时间,Null,0,1) As 是否审核," & vbNewLine & strColSub & _
                    "       c.Id As 记录id, c.限制项目, c.上班时段, c.限号数, c.限约数, c.预约控制 As 预约控制方式" & vbNewLine & _
                    "From 临床出诊号源 A, (" & strSQL & ") B, 临床出诊限制 C, 收费项目目录 E, 部门表 F, 人员表 G, 临床出诊表 H,专业技术职务 I" & vbNewLine & _
                    "Where a.Id = b.号源id And b.安排ID = c.安排id(+) And a.科室id = f.Id And b.项目id = e.Id And b.医生ID = g.ID(+) And b.出诊ID = h.ID" & strWhere & vbNewLine & _
                    "      And g.专业技术职务=i.名称(+) And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
                    "      And (e.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.撤档时间 Is Null)" & vbNewLine & _
                    "      And (f.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.撤档时间 Is Null)" & vbNewLine & _
                    "      And (g.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.撤档时间 Is Null)" & vbNewLine & _
                    "Order By " & str排序号码 & ", b.登记时间 Desc, c.限制项目, c.上班时段"
        Else
            strSQL = "Select " & str是否有效 & _
                    "       b.出诊ID, b.安排id," & _
                    "       Decode(b.安排ID,Null,e.名称,m.名称) As 收费项目, Decode(b.安排ID,Null,a.医生姓名,b.医生姓名) As 医生姓名," & vbNewLine & _
                    "       Decode(b.安排ID,Null,g.简码,n.简码) As 医生简码, Decode(b.安排ID,Null,g.专业技术职务,n.专业技术职务) as 医生职称, 0 As 临时安排," & vbNewLine & _
                    "       Decode(b.安排ID,Null,i.标识符,j.标识符) as 标识符,b.排班规则, b.开始时间, b.终止时间, b.登记时间, 0 As 是否审核," & vbNewLine & strColSub & _
                    "       c.Id As 记录id, c.限制项目, c.上班时段, c.限号数, c.限约数, c.预约控制 As 预约控制方式" & vbNewLine & _
                    "From 临床出诊号源 A, (" & strSQL & ") B, 临床出诊限制 C, 收费项目目录 E, 部门表 F, 人员表 G, 收费项目目录 M, 人员表 N, 临床出诊表 H,专业技术职务 I,专业技术职务 J" & vbNewLine & _
                    "Where a.Id = b.号源id(+) And b.安排ID = c.安排id(+) And a.科室id = f.Id(+)" & vbNewLine & _
                    "      And a.项目id = e.Id And a.医生ID= g.ID(+) And b.项目id = m.Id(+) And b.医生ID= n.ID(+) And b.出诊ID = h.ID(+)" & strWhere & vbNewLine & _
                    "      And g.专业技术职务=i.名称(+) And n.专业技术职务=j.名称(+)" & vbNewLine & _
                    "      And (b.安排ID Is Not Null Or (b.安排ID Is Null And a.排班方式 = 0))" & vbNewLine & _
                    "      And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
                    "      And (e.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.撤档时间 Is Null)" & vbNewLine & _
                    "      And (f.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.撤档时间 Is Null)" & vbNewLine & _
                    "      And (g.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.撤档时间 Is Null)" & vbNewLine & _
                    "      And Nvl(Nvl(f.站点,[5]),Nvl([4],'-')) = Nvl([4],'-')" & vbNewLine & _
                    "      And Not Exists(Select 1 From 临床出诊安排 P,临床出诊表 Q" & vbNewLine & _
                    "                     Where p.出诊ID = q.ID And p.号源ID = a.ID And q.排班方式 = 0 And q.ID <> Nvl(b.出诊id,0))" & vbNewLine & _
                    "Order By " & str排序号码 & ", b.登记时间 Desc, c.限制项目, c.上班时段"
        End If
    Else
        strSQL = "Select m.Id As 安排id, m.出诊ID, m.号源id, m.项目ID, m.医生ID, m.医生姓名, m.排班规则, m.开始时间, m.终止时间, m.登记时间, m.审核时间, m.是否临时安排" & vbNewLine & _
                " From 临床出诊安排 M, 临床出诊安排 J" & vbNewLine & _
                " Where m.出诊id = j.出诊id And m.号源id = j.号源id And j.id = [2] And Nvl(m.是否临时安排, 0) = 1"
        strSQL = "Select " & str是否有效 & _
                "       b.出诊ID, b.安排id, e.名称 As 收费项目, b.医生姓名, g.简码 As 医生简码, g.专业技术职务 as 医生职称,i.标识符," & _
                "       b.排班规则, b.开始时间, b.终止时间, b.登记时间,Decode(b.审核时间,Null,0,1) As 是否审核," & vbNewLine & _
                "       Case When b.登记时间 > Nvl(h.发布时间, To_date('3000-01-01','yyyy-mm-dd')) Then 1 Else 0 End As 临时安排," & vbNewLine & strColSub & _
                "       c.Id As 记录id, c.限制项目, c.上班时段, c.限号数, c.限约数, c.预约控制 As 预约控制方式" & vbNewLine & _
                "From 临床出诊号源 A, (" & strSQL & ") B, 临床出诊限制 C, 收费项目目录 E, 部门表 F, 人员表 G, 临床出诊表 H,专业技术职务 I" & vbNewLine & _
                "Where a.Id = b.号源id And b.安排ID = c.安排id(+) And a.科室id = f.Id And b.项目id = e.Id And b.医生ID = g.ID(+) And b.出诊ID = h.ID and g.专业技术职务=i.名称(+)" & vbNewLine & _
                "      And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
                "      And (e.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.撤档时间 Is Null)" & vbNewLine & _
                "      And (f.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.撤档时间 Is Null)" & vbNewLine & _
                "      And (g.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.撤档时间 Is Null)" & vbNewLine & _
                "Order By " & str排序号码 & ", b.登记时间 Desc, c.限制项目, c.上班时段"
        If mblnShowInvalidPlan = False Then
            strSQL = "Select 是否有效, 出诊ID, 安排id, 收费项目, 医生姓名, 医生简码, 医生职称,标识符," & vbNewLine & _
                    "        排班规则, 开始时间, 终止时间, 登记时间, 是否审核, 临时安排," & vbNewLine & _
                    "        号源id, 号类, 号码, 是否建病案, 预约天数, 出诊频次," & vbNewLine & _
                    "        假日控制状态, 是否临床排班, 排班方式, 科室, 科室简码, 是否假日换休," & vbNewLine & _
                    "        记录id, 限制项目, 上班时段, 限号数, 限约数, 预约控制方式" & vbNewLine & _
                    " From (" & strSQL & ")" & vbNewLine & _
                    " Where 是否有效=1 Or 是否审核=0"
        End If
    End If
    Set GetPlanRuleData = zlDatabase.OpenSQLRecord(strSQL, "获取排班信息", lng出诊ID, lng安排ID, UserInfo.ID, _
        gstrNodeNo, gVisitPlan_ModulePara.str号源维护站点)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPlanRecords(ByVal lng出诊ID As Long, _
    Optional ByVal dtStart As Date, Optional ByVal dtEnd As Date) As ADODB.Recordset
    '功能：获取安排记录
    Dim strSQL As String, str是否有效 As String
    Dim strPrivsWhere As String
    Dim str排序号码 As String

    Err = 0: On Error GoTo errHandler
    str排序号码 = IIf(gVisitPlan_ModulePara.byt号码比较方式 = 0, "c.号码", "Lpad(c.号码,5,'0')")
    '没有"所有科室"权限的操作员只能操作自己所属科室的号源
    If HavePrivs(mstrPrivs, "所有科室") = False Then
        strPrivsWhere = "      And Exists (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = [2])"
    End If
    
    '有效安排，以下规则必须同时满足：
    '    --1.规则已审核
    '    --2.规则终止时间大于当前时间
    '    --3.后面无其它临时安排或者有其它安排但时间范围没有被覆盖
    '    --4.没有调整为其它排班方式，或者调整为其它排班方式但还没有出诊安排
    '    --5.有有效的出诊记录
    '说明：由多个临时安排一起覆盖的安排，判断不了
    str是否有效 = "Nvl((Select 1" & vbNewLine & _
                " From Dual" & vbNewLine & _
                " Where a.审核时间 Is Not Null And a.终止时间 > Sysdate" & vbNewLine & _
                "       And Not Exists(Select 1" & vbNewLine & _
                "           From 临床出诊安排" & vbNewLine & _
                "           Where 审核时间 Is Not Null And 号源id = a.号源id And 登记时间 > a.登记时间" & vbNewLine & _
                "                 And (Nvl(a.是否临时安排, 0) = 0 And Nvl(是否临时安排, 0) = 0 Or Nvl(a.是否临时安排, 0) = 1)" & vbNewLine & _
                "                 And Decode(Sign(Sysdate - 开始时间), 1, Sysdate, 开始时间) <= Decode(Sign(Sysdate - a.开始时间), 1, Sysdate, a.开始时间)" & vbNewLine & _
                "                 And 终止时间 >= b.终止时间)" & vbNewLine & _
                "       And Not Exists(Select 1" & vbNewLine & _
                "           From 临床出诊安排 P, 临床出诊表 Q" & vbNewLine & _
                "           Where p.出诊id = q.Id And p.号源id = a.号源id And Nvl(q.排班方式, 0) In (1, 2) And p.开始时间 < Sysdate)" & vbNewLine & _
                "       And Exists(Select 1" & vbNewLine & _
                "           From 临床出诊记录 P, 临床出诊安排 Q" & vbNewLine & _
                "           Where p.安排id = q.Id And q.出诊id = [1] And q.号源id = a.号源ID And p.出诊日期 + 1 > Sysdate)" & vbNewLine & _
                "   ), 0) As 是否有效,"
                
    strSQL = "Select " & str是否有效 & vbNewLine & _
            "        " & str排序号码 & " As 排序号码, c.Id As 号源id, c.号类, c.号码, Nvl(c.是否建病案, 0) As 是否建病案, c.预约天数, c.出诊频次," & vbNewLine & _
            "        Decode(c.假日控制状态, 1, '开放预约', 2, '禁止预约', 3, '受节假日设置控制', '不上班') As 假日控制状态," & vbNewLine & _
            "        Decode(c.排班方式, 1, '按月排班', 2, '按周排班', '固定排班') As 排班方式," & vbNewLine & _
            "        Nvl(c.是否假日换休, 0) As 是否假日换休, Nvl(c.是否临床排班, 0) As 是否临床排班," & vbNewLine & _
            "        f.名称 As 科室, f.简码 As 科室简码, e.名称 As 收费项目, a.医生姓名, g.专业技术职务 As 医生职称,h.标识符,g.简码 As 医生简码," & vbNewLine & _
            "        a.出诊id, a.Id As 安排id, a.开始时间, a.终止时间," & vbNewLine & _
            "        b.Id As 记录id, b.出诊日期, b.上班时段, b.限号数, b.限约数, b.已挂数, b.已约数, b.预约控制 As 预约控制方式," & vbNewLine & _
            "        b.停诊开始时间, b.停诊终止时间, b.停诊原因, b.替诊医生姓名, b.是否临时出诊, b.是否锁定" & vbNewLine & _
            " From 临床出诊安排 A, 临床出诊记录 B, 临床出诊号源 C, 收费项目目录 E, 部门表 F, 人员表 G,专业技术职务 H" & vbNewLine & _
            " Where a.Id = b.安排id(+) And a.号源id = c.Id And c.科室id = f.Id And a.项目id = e.Id And a.医生id = g.Id(+) And a.出诊id = [1]" & vbNewLine & _
            "       And a.审核时间 Is Not Null" & vbNewLine & _
            "       And g.专业技术职务=h.名称(+) " & vbNewLine & _
            "       And Nvl(c.是否删除, 0) = 0 And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)" & vbNewLine & _
            "       And (e.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.撤档时间 Is Null)" & vbNewLine & _
            "       And (f.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.撤档时间 Is Null)" & vbNewLine & _
            "       And (g.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.撤档时间 Is Null)" & vbNewLine & _
                    strPrivsWhere & vbNewLine & _
            "       And Nvl(b.出诊日期,[3]) Between [3] And [4]" & vbNewLine & _
            "       And Exists(Select 1 From 临床出诊记录 Where 安排ID = a.ID And b.出诊日期 Between [3] And [4])" & vbNewLine & _
            " Order By " & str排序号码 & ", 科室, 收费项目, 医生姓名, 出诊日期, 上班时段"
    Set GetPlanRecords = zlDatabase.OpenSQLRecord(strSQL, "获取排班信息", lng出诊ID, UserInfo.ID, dtStart, dtEnd, gstrNodeNo)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadPlanDataSub(Optional ByVal lng出诊ID As Long, Optional ByVal lng安排ID As Long)
    '功能：加载同一个号源多个安排的安排数据
    '入数：
    '   lng出诊ID - 出诊ID
    Err = 0: On Error GoTo errHandler
    Set mrsRuleRecordsSub = Nothing
    If lng出诊ID <> 0 And lng安排ID <> 0 Then
        '加载数据
        Screen.MousePointer = vbHourglass
        '规则
        Set mrsRuleRecordsSub = GetPlanRuleData(lng出诊ID, lng安排ID, Val(lblPublishInfo.Tag) = 1)
        Call LoadPlanDataByRecordset(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecordsSub, 0)
        Screen.MousePointer = vbDefault
    Else
        Call LoadPlanDataByRecordset(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecordsSub, 0)
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistRuleSub_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If NewRow < 2 Then
        Call mfrmMain.StatusShowInfoChanged(2, "")
    Else
        Call mfrmMain.StatusShowInfoChanged(2, "当前共" & mlngSignalCount & "条号源，当前号源：" & vsfRegistRuleSub.TextMatrix(NewRow, COL_号码) & _
            "，开始时间：" & vsfRegistRuleSub.TextMatrix(NewRow, COL_开始时间) & "，终止时间：" & vsfRegistRuleSub.TextMatrix(NewRow, COL_终止时间) & "")
    End If
End Sub

Private Sub vsfRegistRuleSub_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    On Error Resume Next
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    vsfRegistRule.LeftCol = NewLeftCol
End Sub

Private Sub vsfRegistRuleSub_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    On Error Resume Next
    Call SetPlanGridRangeColor(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule, mstrOldSelRangeRuleSub)
    mstrOldSelRangeRuleSub = vsfRegistRuleSub.Row & "|" & vsfRegistRuleSub.RowSel & "|" & vsfRegistRuleSub.Col & "|" & vsfRegistRuleSub.ColSel
End Sub

Private Sub vsfRegistRuleSub_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistRuleSub, Me.Name, "规则")
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistRule, Me.Name, "规则")
End Sub

Private Sub vsfRegistRuleSub_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    If Val(vsfRegistRuleSub.RowData(NewRow)) = -1 Then Cancel = True
End Sub

Private Sub vsfRegistRuleSub_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = gPlanGrid_ColIndex.COL_图标 Then Cancel = True: Exit Sub
End Sub

Private Sub vsfRegistRuleSub_DblClick()
    Dim lng号源Id As Long, lng安排ID As Long
    Dim frmEdit As New frmClinicPlanEdit
    Dim strCurItem As String, blnUpdate As Boolean
    Dim lngCol As Long, lngRow As Long
    Dim strSort As String

    Err = 0: On Error GoTo errHandler
    lngCol = vsfRegistRuleSub.MouseCol
    lngRow = vsfRegistRuleSub.MouseRow
    If lngRow = 0 Or lngRow = 1 Then
        '排序
        If mrsRuleRecordsSub Is Nothing Then Exit Sub
        If mrsRuleRecordsSub.RecordCount = 0 Then Exit Sub
        strSort = GetPlanSortCircleStr(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule, lngRow, lngCol)
        If strSort <> "" Then
            mrsRuleRecordsSub.Sort = strSort
            Screen.MousePointer = vbHourglass
            Call LoadPlanDataByRecordset(vsfRegistRuleSub, gPlanGrid_DataStyle.Data_FixedRule, mrsRuleRecordsSub, 0, , True)
            Screen.MousePointer = vbDefault
        End If
    Else
        With vsfRegistRuleSub
            lng号源Id = Val(.TextMatrix(.Row, COL_号源ID))
            lng安排ID = Val(.TextMatrix(.Row, COL_安排ID))
            strCurItem = .Cell(flexcpData, 0, .Col)
            If lng号源Id = 0 And lng安排ID = 0 Then Exit Sub
            If strCurItem = "" Then Exit Sub
            
            blnUpdate = zlStr.IsHavePrivs(mstrPrivs, "出诊安排") _
                And (Val(.TextMatrix(.Row, COL_临时安排)) = 1 And Val(.TextMatrix(.Row, COL_是否审核)) = 0 Or Val(lblPublishInfo.Tag) = 0)
            If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
                '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
                If Trim(.TextMatrix(.Row, COL_是否临床排班)) = "" Then blnUpdate = False
            End If
    
            If frmEdit.ShowMe(Me, 0, IIf(blnUpdate, Fun_TempPlan, Fun_View), mlng出诊ID, lng号源Id, lng安排ID, strCurItem, mstrPrivs) Then
                If blnUpdate Then Call RefreshDataSub
            End If
        End With
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshDataSub()
    '刷新临时安排列表
    Dim lngOldRow As Long, lngoldCol As Long

    Err = 0: On Error GoTo errHandler
    With vsfRegistRuleSub
        lngOldRow = .Row
        lngoldCol = .Col

        vsfRegistRule.Tag = "": Call vsfRegistRule_EnterCell

        If .Rows > .FixedRows And .Cols > .FixedCols Then
            .Row = IIf(lngOldRow = 0 Or lngOldRow > .Rows - 1, .FixedRows, lngOldRow)
            .Col = IIf(lngoldCol = 0 Or lngoldCol > .Cols - 1, .FixedCols, lngoldCol)
            .ShowCell .Row, .Col  '立刻显示到指定单元
            If .Visible And .Enabled Then .SetFocus
        End If
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfRegistRuleSub_GotFocus()
    Call SetSelectedBackColor(vsfRegistRuleSub, True)
End Sub

Private Sub vsfRegistRuleSub_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RegistPlan_KeyDown(vsfRegistRuleSub, KeyCode, Shift)
End Sub

Private Sub vsfRegistRuleSub_LostFocus()
    Call SetSelectedBackColor(vsfRegistRuleSub, False)
End Sub

Private Sub vsfRegistRuleSub_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbCommandBar As CommandBar

    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)

    Set cbCommandBar = GetPopupCommandBar(Me, mcbsMain)
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub

    cbCommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txtFind_KeyPress(index As Integer, KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        Call ExecuteFilter(Val(tbPage(Pg_出诊安排).Tag) = 0)
    End If
End Sub

Private Sub SetSelectedBackColor(vsfGrid As VSFlexGrid, ByVal blnFocus As Boolean)
    '根据网格激活状态设置选择行背景颜色
    Dim lngRowStart As Long, lngColStart As Long, lngRowEnd As Long, lngColEnd As Long
    Dim strOldSelRange As String, dataType As gPlanGrid_DataStyle

    Err = 0: On Error GoTo errHandler
    If vsfGrid Is vsfRegistPlan Then
        strOldSelRange = mstrOldSelRangePlan
        dataType = gPlanGrid_DataStyle.Data_Plan
    ElseIf vsfGrid Is vsfRegistRule Then
        strOldSelRange = mstrOldSelRangeRule
        dataType = gPlanGrid_DataStyle.Data_FixedRule
    ElseIf vsfGrid Is vsfRegistRuleSub Then
        strOldSelRange = mstrOldSelRangeRuleSub
        dataType = gPlanGrid_DataStyle.Data_FixedRule
    Else
        Exit Sub
    End If
    If blnFocus Then
        Call SetPlanGridRangeColor(vsfGrid, dataType, strOldSelRange)
    Else
        If GetSelectRange(vsfGrid, strOldSelRange, lngRowStart, lngRowEnd, lngColStart, lngColEnd) Then
            vsfGrid.Cell(flexcpBackColor, lngRowStart, lngColStart, lngRowEnd, lngColEnd) = G_LostFocusColor
        End If
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub fraSplitRule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Err = 0: On Error Resume Next
    If Button <> vbLeftButton Then Exit Sub
    If vsfRegistRule.Height + Y < 1200 Or vsfRegistRuleSub.Height - Y < 1200 Then Exit Sub

    fraSplitRule.Top = fraSplitRule.Top + Y
    
    vsfRegistRule.Height = vsfRegistRule.Height + Y
    vsfRegistRuleSub.Top = vsfRegistRuleSub.Top + Y
    vsfRegistRuleSub.Height = vsfRegistRuleSub.Height - Y
    Me.Refresh
End Sub

Private Sub picRegistRule_Resize()
    On Error Resume Next
    vsfRegistRule.Move -10, 0, picRegistRule.ScaleWidth + 20, picRegistRule.ScaleHeight * 2 / 3
    fraSplitRule.Move 0, vsfRegistRule.Top + vsfRegistRule.Height, picRegistRule.ScaleWidth + 20
    With vsfRegistRuleSub
        .Left = -10
        .Top = fraSplitRule.Top + fraSplitRule.Height
        .Width = picRegistRule.ScaleWidth + 20
        .Height = picRegistRule.ScaleHeight - .Top + 10
    End With
End Sub

Private Sub txtFind_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then
        '按了右键菜单快捷键，清除粘贴板内容
        If Clipboard.GetText <> "" Then Clipboard.Clear
    End If
End Sub

Private Sub txtFind_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtFind(index).Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtFind(index).Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtFind_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtFind(index).Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


