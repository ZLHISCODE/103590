VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicPlanDaysManage 
   BorderStyle     =   0  'None
   Caption         =   "出诊安排管理"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   9150
      MaxLength       =   100
      TabIndex        =   13
      Top             =   1140
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.PictureBox picSelectWeek 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   90
      ScaleHeight     =   345
      ScaleWidth      =   10845
      TabIndex        =   3
      Top             =   450
      Width           =   10845
      Begin VB.CheckBox chkShowAllPlan 
         Caption         =   "显示本月所有安排"
         Height          =   225
         Left            =   7020
         TabIndex        =   11
         Top             =   75
         Width           =   1755
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "第6周"
         Height          =   195
         Index           =   6
         Left            =   6000
         TabIndex        =   10
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "第5周"
         Height          =   195
         Index           =   5
         Left            =   5040
         TabIndex        =   9
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "第4周"
         Height          =   195
         Index           =   4
         Left            =   4050
         TabIndex        =   8
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "第3周"
         Height          =   195
         Index           =   3
         Left            =   3045
         TabIndex        =   7
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "第2周"
         Height          =   195
         Index           =   2
         Left            =   2055
         TabIndex        =   6
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "第1周"
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   5
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "全部"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   90
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfRegistPlan 
      Height          =   2445
      Left            =   630
      TabIndex        =   1
      Top             =   1110
      Width           =   3495
      _cx             =   6165
      _cy             =   4313
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
      FormatString    =   $"frmClinicPlanDaysManage.frx":0000
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
         Picture         =   "frmClinicPlanDaysManage.frx":0075
         ScaleHeight     =   135
         ScaleWidth      =   150
         TabIndex        =   2
         Top             =   90
         Width           =   150
      End
   End
   Begin VB.Label lblPublishInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发布人：冉俊明  发布时间：2016-01-02 12:32:12"
      Height          =   180
      Left            =   6840
      TabIndex        =   12
      Top             =   150
      Width           =   4050
   End
   Begin VB.Line lineSplit 
      BorderColor     =   &H8000000A&
      X1              =   1020
      X2              =   5010
      Y1              =   960
      Y2              =   960
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   10845
      _Version        =   589884
      _ExtentX        =   19129
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "出诊安排>出诊安排"
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
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1035
      Left            =   30
      Top             =   960
      Width           =   525
   End
End
Attribute VB_Name = "frmClinicPlanDaysManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mlngModule As Long
Private mstrPrivs As String

Private mbytFun As Byte '1-月安排，2-周安排
Private mlng出诊ID As Long
Private mrsPlanRecords As ADODB.Recordset
Private mintFindType As Integer

Private mlngCopyPlanID As Long, mstrCopyPlanItem As String
Private mintYear As Integer, mintMonth As Integer, mintWeek As Integer
Private mlng跨月周出诊ID As Long, mintElseWeek As Integer '周安排跨月时会有两个出诊表，记录另一个出诊表的信息
Private mdtStartDate As Date, mdtEndDate As Date

Private mdtToday As Date
Private mstrOldSelRangePlan As String '选择网格区域，格式"开始行|结束行|开始列|结束列"

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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除出诊表(&D)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddNewSignalSource, "新增号源安排(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyPlanItem, "调整安排(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyUnitRegist, "调整预约挂号控制(&R)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStartNO, "全部启用序号控制(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStopNO, "全部取消序号控制(&T)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CopyPlan, "复制安排(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PastPlan, "粘贴安排(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearCurPlan, "清除当前安排(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAllPlan, "清除当前号源安排(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAll, "清除所有号源安排(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyToDay, "应用于“所有单日”(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyToWeekDay, "应用于“所有星期几”(&W)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PublishPlan, "发布安排(&G)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnPublishPlan, "取消发布(&I)")
        
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UpdateUnitRegist, "调整预约挂号控制(&H)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintPlan, "打印出诊表(&P)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PlanSaveAsTemplet, "另存为模板(&A)..."): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextNewPlan, "生成月出诊表(&N)")
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowDoctorStopPlan, "显示医生停诊安排(&P)", cbrControl.index)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_PlanChangeInfo, "查询变动信息(&C)", cbrControl.index)
        cbrControl.BeginGroup = True
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddMonthPlan, "月出诊表", cbrControl.index + 1): cbrControl.BeginGroup = True
        cbrControl.ToolTipText = "制定月出诊表"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddWeekPlan, "周出诊表", cbrControl.index + 1)
        cbrControl.ToolTipText = "制定周出诊表"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除出诊表", cbrControl.index + 1): cbrControl.BeginGroup = True
        
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

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextNewPlan, "生成月出诊表", cbrControl.index + 1): cbrControl.BeginGroup = True
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
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
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

Private Function IsInCurrentPlan(ByVal vsfGrid As VSFlexGrid) As Boolean
    '判断当前选择日期是否在当前安排中
    Dim strCurItem As String
    On Error GoTo errHandler
    With vsfGrid
        If .Col < gPlanGrid_FixedCols Then Exit Function
        If .Row < .FixedRows Or .Row > .Rows - 1 Then Exit Function
        strCurItem = .Cell(flexcpData, 0, .Col)
        If IsDate(strCurItem) = False Then Exit Function
        If strCurItem < mdtStartDate Or strCurItem > mdtEndDate Then Exit Function
    End With
    IsInCurrentPlan = True
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
    Dim strDoctorName As String
    Dim blnEnabled As Boolean
    
    '说明：显示本月所有安排时，只能双击查看，不能操作任何功能
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    blnEnabled = mlng出诊ID <> 0
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = vsfRegistPlan.Rows > vsfRegistPlan.FixedRows
    Case conMenu_EditPopup
        If mfrmMain.mFunListActived Then
            Control.Visible = HavePrivs(mstrPrivs, "出诊安排;发布安排;取消发布")
        Else
            Control.Visible = chkShowAllPlan.Value = vbUnchecked _
                And HavePrivs(mstrPrivs, "出诊安排;调整安排;临时出诊安排;停诊;替诊;加号;减号;调整分诊诊室;调整预约挂号;模板管理")
        End If
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AddMonthPlan, conMenu_Edit_AddWeekPlan '制定月出诊表,制定周出诊表
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Delete '删除出诊表
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible And blnEnabled And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
    Case conMenu_Edit_PublishPlan, conMenu_Edit_UnPublishPlan '发布安排,取消发布
        Control.Visible = mfrmMain.mFunListActived And HavePrivs(mstrPrivs, Decode(Control.ID, _
            conMenu_Edit_PublishPlan, "发布安排", conMenu_Edit_UnPublishPlan, "取消发布")) And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then
            If Control.ID = conMenu_Edit_PublishPlan Then
                blnEnabled = Val(lblPublishInfo.Tag) = 0
            Else
                blnEnabled = Val(lblPublishInfo.Tag) = 1
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled
    
    Case conMenu_Edit_ModifyPlanItem '调整出诊项
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsCan临床排班(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ModifyUnitRegist '调整预约挂号控制
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsCan临床排班(vsfRegistPlan)
        If blnEnabled Then blnEnabled = Is禁止预约(vsfRegistPlan) = False
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_AllStartNO, conMenu_Edit_AllStopNO '全部启用序号控制,全部取消序号控制
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_CopyPlan '复制安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_PastPlan '粘贴安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsCan临床排班(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled And mlngCopyPlanID <> 0
    Case conMenu_Edit_ClearCurPlan, conMenu_Edit_ClearAllPlan '清除当前安排,清除当前号源所有安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If Control.ID = conMenu_Edit_ClearCurPlan And blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsCan临床排班(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ClearAll '清除所有号源安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ApplyToDay, conMenu_Edit_ApplyToWeekDay '应用于“所有单日”,应用于“所有星期几”
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 0 And chkShowAllPlan.Value = vbUnchecked
        If Control.ID = conMenu_Edit_ApplyToWeekDay And Control.Visible Then Control.Visible = mbytFun = 1
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsCan临床排班(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled

    '已发布安排调整
    Case conMenu_Edit_AddNewSignalSource '新增号源安排
        Control.Visible = HavePrivs(mstrPrivs, "调整安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_LockResource, conMenu_Edit_UnLockResource '锁号,解锁
        Control.Visible = HavePrivs(mstrPrivs, "调整安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then
            If Control.ID = conMenu_Edit_LockResource Then
                blnEnabled = (PlanIsSelOne(vsfRegistPlan) = False Or PlanIsLocked(vsfRegistPlan) = False)
            Else
                blnEnabled = (PlanIsSelOne(vsfRegistPlan) = False Or PlanIsLocked(vsfRegistPlan))
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_AddTempPlan, conMenu_Edit_UpdatePlan '临时出诊,调整发布后的安排
        Control.Visible = HavePrivs(mstrPrivs, "调整安排") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        If Control.ID = conMenu_Edit_UpdatePlan And blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfRegistPlan)
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_StopOutCall, conMenu_Edit_UnStopOutCall, conMenu_Edit_OpenStopPlan '停诊,取消停诊,开放停诊安排
        Control.Visible = HavePrivs(mstrPrivs, "停诊") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfRegistPlan)
        If blnEnabled Then
            If Control.ID = conMenu_Edit_StopOutCall Then
                blnEnabled = (PlanIsStopVisit(vsfRegistPlan) = False)
            Else
                blnEnabled = PlanIsStopVisit(vsfRegistPlan)
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ModifyDoctor, conMenu_Edit_UnModifyDoctor '替诊,取消替诊
        Control.Visible = HavePrivs(mstrPrivs, "替诊") And mfrmMain.mFunListActived = False _
            And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsStopVisit(vsfRegistPlan) = False
        If blnEnabled Then
            If Control.ID = conMenu_Edit_ModifyDoctor Then
                blnEnabled = (PlanIsReplaceDoctor(vsfRegistPlan) = False)
            Else
                blnEnabled = PlanIsReplaceDoctor(vsfRegistPlan)
            End If
        End If
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_AddNumberLimit, conMenu_Edit_ReduceNumberLimit, _
        conMenu_Edit_ModifyDoctorOffice, conMenu_Edit_UpdateUnitRegist '加号,减号,调整分诊诊室,调整预约挂号控制
        Control.Visible = HavePrivs(mstrPrivs, Decode(Control.ID, _
            conMenu_Edit_AddNumberLimit, "加号", conMenu_Edit_ReduceNumberLimit, "减号", _
            conMenu_Edit_ModifyDoctorOffice, "调整分诊诊室", conMenu_Edit_UpdateUnitRegist, "调整预约挂号")) _
            And mfrmMain.mFunListActived = False And Val(lblPublishInfo.Tag) = 1 And chkShowAllPlan.Value = vbUnchecked
        If blnEnabled Then blnEnabled = SelectedIsNotNull(vsfRegistPlan)
        If blnEnabled Then blnEnabled = IsInCurrentPlan(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsValid(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsSelOne(vsfRegistPlan)
        If blnEnabled Then blnEnabled = PlanIsStopVisit(vsfRegistPlan) = False
        If Control.ID = conMenu_Edit_UpdateUnitRegist And blnEnabled Then blnEnabled = Is禁止预约(vsfRegistPlan) = False
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_NextNewPlan, conMenu_Edit_PlanSaveAsTemplet '生成新安排,另存为模板(&A)...
        Control.Visible = chkShowAllPlan.Value = vbUnchecked And HavePrivs(mstrPrivs, Decode(Control.ID, _
            conMenu_Edit_NextNewPlan, "出诊安排", conMenu_Edit_PlanSaveAsTemplet, "模板管理"))
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_PrintPlan '    打印出诊表
        Control.Visible = HavePrivs(mstrPrivs, Decode(mbytFun, 1, "月出诊表", "周出诊表")) And chkShowAllPlan.Value = vbUnchecked
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_View_FindType '查找方式
        Control.Caption = "按" & Decode(mintFindType, 0, "号码", 1, "科室", 2, "医生", "号码") & "过滤↓"
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '查找方式
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
    Case conMenu_View_ShowDoctorStopPlan '显示医生停诊安排
        Control.Visible = mfrmMain.mFunListActived = False
        blnEnabled = False
        If vsfRegistPlan.Row >= vsfRegistPlan.FixedRows Then
            blnEnabled = Trim(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, COL_医生)) <> ""
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
    Dim lng安排ID As Long
    Dim lng记录ID As Long, lng号源Id As Long, str号码 As String, strItem As String
    Dim obj出诊记录 As 出诊记录, obj出诊号源 As 出诊号源
    Dim blnFixedRule As Boolean
    Dim strIDs As String, lngRowStart As Long, lngRowEnd As Long, i As Integer
    Dim lngCurCol As Long
    Dim strKey As String
    Dim strDoctorName As String
    Dim str记录IDs As String
    
    Err = 0: On Error GoTo errHandler
    lng安排ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_安排ID))
    strItem = vsfRegistPlan.Cell(flexcpData, 0, vsfRegistPlan.Col)
    strDoctorName = Trim(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, COL_医生))
    
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_Delete '删除出诊表
        If DeletePlan(mlng出诊ID, mbytFun, mintWeek, mlng跨月周出诊ID, mintElseWeek) Then Call mfrmMain.NodeChanged("")
    Case conMenu_Edit_ModifyPlanItem '调整安排
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        If lng号源Id <> 0 Or lng安排ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            If frmEdit.ShowMe(Me, mbytFun, Fun_Update, mlng出诊ID, lng号源Id, lng安排ID, strItem) Then
                Call RefreshOneData
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ModifyUnitRegist '调整预约挂号控制
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        If lng号源Id <> 0 Or lng安排ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            Call frmEdit.ShowMe(Me, IIf(mbytFun = 1, 1, 2), Fun_UpdateUnit, mlng出诊ID, lng号源Id, lng安排ID, strItem)
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
        If PastPlan(mlng出诊ID, mlngCopyPlanID, mstrCopyPlanItem) Then Call RefreshOneData
    Case conMenu_Edit_ClearCurPlan '清除当前安排
        If strItem = "" Then Exit Sub
        If IsDate(strItem) = False Then Exit Sub
        str号码 = vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号码)
        If MsgBox("你确定要清除号码为【" & str号码 & "】【" & Format(strItem, "mm月dd日") & "】的安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlan(lng安排ID, strItem, True) Then
            Call RefreshOneData
            If mlngCopyPlanID = lng安排ID And mstrCopyPlanItem = strItem Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ClearAllPlan '清除当前号源安排
        str号码 = vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号码)
        If MsgBox("你确定要清除号码为【" & str号码 & "】的所有安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        If ZlClearPlanBatch(mlng出诊ID, lng号源Id) Then
            Call RefreshOneData
            If mlngCopyPlanID = lng安排ID Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ClearAll '清除所有号源安排
        If MsgBox("你确定要清除所有号源的安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If ZlClearPlanBatch(mlng出诊ID, 0, IIf(HavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID)) Then
            Call RefreshData(mbytFun, mlng出诊ID)
            mlngCopyPlanID = 0: mstrCopyPlanItem = ""
        End If
    Case conMenu_Edit_ApplyToDay '应用于“所有单日”
        If ApplyToDay(lng安排ID, strItem) Then Call RefreshOneData
    Case conMenu_Edit_ApplyToWeekDay '应用于“所有星期几”
        If ApplyToWeekDay(lng安排ID, strItem) Then Call RefreshOneData
    Case conMenu_Edit_PublishPlan '发布安排
        If PublishPlan(mlng出诊ID, True, mlng跨月周出诊ID) Then
            Call PrintPlan(mlng出诊ID, mbytFun)
            '刷新数据
            If mbytFun = 1 Then
                strKey = "K2_" & mintYear & "_" & mintMonth 'XX月出诊表节点: K2_年份_月份
            Else
                strKey = "K3_" & mintYear & "_" & mintMonth & "_" & mintWeek 'XX周出诊表节点：K3_年份_月份_周数
            End If
            Call mfrmMain.NodeChanged(strKey)
        End If
    Case conMenu_Edit_UnPublishPlan '取消发布
        If PublishPlan(mlng出诊ID, False, mlng跨月周出诊ID) Then
            '刷新数据
            If mbytFun = 1 Then
                strKey = "K2_" & mintYear & "_" & mintMonth 'XX月出诊表节点: K2_年份_月份
            Else
                strKey = "K3_" & mintYear & "_" & mintMonth & "_" & mintWeek 'XX周出诊表节点：K3_年份_月份_周数
            End If
            Call mfrmMain.NodeChanged(strKey)
        End If
    
    '已发布安排调整
    Case conMenu_Edit_AddNewSignalSource '新增号源安排
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, mbytFun, Fun_AddSignalSourcePlan, mlng出诊ID, , , strItem) Then
            Call RefreshData(mbytFun, mlng出诊ID)
        End If
    Case conMenu_Edit_LockResource '锁号
        Call LockPlan(False)
    Case conMenu_Edit_UnLockResource '解锁
        Call LockPlan(True)
    Case conMenu_Edit_AddTempPlan '临时出诊
        Set frmEdit = New frmClinicPlanEdit
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        If frmEdit.ShowMe(Me, mbytFun, Fun_TempPlanRecord, mlng出诊ID, lng号源Id, lng安排ID, strItem) Then
            Call RefreshOneData
        End If
    Case conMenu_Edit_UpdatePlan '调整发布后的安排
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        If lng号源Id = 0 Or lng安排ID = 0 Then Exit Sub
        
        Call LockPlanByDay(False, str记录IDs)
        Set frmEdit = New frmClinicPlanEdit
        If frmEdit.ShowMe(Me, mbytFun, Fun_UpdatePlan, mlng出诊ID, lng号源Id, lng安排ID, strItem) Then
            Call LockPlanByDay(True, str记录IDs)
            Call RefreshOneData
        End If
        Call LockPlanByDay(True, str记录IDs)
    Case conMenu_Edit_StopOutCall '停诊
        lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If lng记录ID = 0 Then Exit Sub
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        'bytFun 功能：1-停诊,2-取消停诊,3-替诊,4-取消替诊
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 1, lng记录ID) Then
            Call RefreshOneData
        End If
    Case conMenu_Edit_UnStopOutCall '取消停诊
        lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If lng记录ID = 0 Then Exit Sub
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        'bytFun 功能：1-停诊,2-取消停诊,3-替诊,4-取消替诊
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 2, lng记录ID) Then
            Call RefreshOneData
        End If
    Case conMenu_Edit_OpenStopPlan '开放停诊安排
        lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
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
        lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If lng记录ID = 0 Then Exit Sub
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        'bytFun 功能：1-停诊,2-取消停诊,3-替诊,4-取消替诊
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 3, lng记录ID) Then
            Call RefreshOneData
        End If
    Case conMenu_Edit_UnModifyDoctor '取消替诊
        lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If lng记录ID = 0 Then Exit Sub
        Set frmStopVisitAndModifyDoctor = New frmClinicPlanStopVisitAndModifyDoctor
        'bytFun 功能：1-停诊,2-取消停诊,3-替诊,4-取消替诊
        If frmStopVisitAndModifyDoctor.ShowMe(Me, mlngModule, 4, lng记录ID) Then
            Call RefreshOneData
        End If
    Case conMenu_Edit_AddNumberLimit '加号
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If Get出诊记录(lng号源Id, lng记录ID, True, obj出诊号源, obj出诊记录) Then
            Set frmNumberLimitModify = New frmClinicPlanNumberLimitModify
            If frmNumberLimitModify.ShowMe(Me, 1, obj出诊号源, obj出诊记录) Then
                Call RefreshOneData
            End If
        End If
    Case conMenu_Edit_ReduceNumberLimit '减号
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If Get出诊记录(lng号源Id, lng记录ID, True, obj出诊号源, obj出诊记录) Then
            Set frmNumberLimitModify = New frmClinicPlanNumberLimitModify
            If frmNumberLimitModify.ShowMe(Me, 2, obj出诊号源, obj出诊记录) Then
               Call RefreshOneData
            End If
        End If
    Case conMenu_Edit_ModifyDoctorOffice '调整分诊诊室
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If Get出诊记录(lng号源Id, lng记录ID, True, obj出诊号源, obj出诊记录) Then
            Set frmOfficeAndUnitRegModify = New frmClinicPlanOfficeAndUnitRegModify
            Call frmOfficeAndUnitRegModify.ShowMe(Me, 1, obj出诊号源, obj出诊记录, True)
        End If
    Case conMenu_Edit_UpdateUnitRegist '调整预约挂号控制
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
        If Get出诊记录(lng号源Id, lng记录ID, True, obj出诊号源, obj出诊记录) Then
            Set frmOfficeAndUnitRegModify = New frmClinicPlanOfficeAndUnitRegModify
            Call frmOfficeAndUnitRegModify.ShowMe(Me, 2, obj出诊号源, obj出诊记录, True)
        End If
    Case conMenu_Edit_PlanSaveAsTemplet '另存为模板(&A)...
        Call SaveAsTemplet(mlng出诊ID, mbytFun = 1)
    Case conMenu_Edit_NextNewPlan '生成新安排
        Call NextNewPlanByPlan(mlng出诊ID, mbytFun = 1)
    Case conMenu_Edit_PrintPlan '    打印出诊表
        Call PrintPlan(mlng出诊ID, mbytFun, 1)
    Case conMenu_View_Refresh
        Call RefreshData(mbytFun, mlng出诊ID)
    Case conMenu_View_PlanChangeInfo '查询信息
        Dim frmPlanChangeHistory As New frmClinicPlanChangeHistory
        frmPlanChangeHistory.ShowMe Me, mlngModule
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 3 '查找方式
        mintFindType = Val(Right(Control.ID, 2)) - 1
        mcbsMain.RecalcLayout
        txtFind(1).Text = ""
        If txtFind(1).Visible And txtFind(1).Enabled Then txtFind(1).SetFocus
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

Private Sub PrintPlan(ByVal lng出诊ID As Long, ByVal bytFun As Byte, Optional ByVal bytMode As Byte)
    '打印出诊表
    '入参：
    '   mbytFun 1-月安排，2-周安排
    '   bytMode 0-发布后打印,1-菜单选择打印
    Dim str出诊ID As String
    
    Err = 0: On Error GoTo errHandler
    If bytMode = 1 Then '防止误操作
        If MsgBox("要打印出诊表吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    If bytFun = 1 Then
        If gVisitPlan_ModulePara.byt出诊表打印方式 = 1 Or bytMode = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_2", Me, "出诊ID=" & mlng出诊ID, 2)
        ElseIf gVisitPlan_ModulePara.byt出诊表打印方式 = 2 Then
            If MsgBox("要打印出诊表吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_2", Me, "出诊ID=" & mlng出诊ID, 2)
            End If
        End If
    Else
        str出诊ID = mlng出诊ID
        If mlng跨月周出诊ID <> 0 Then str出诊ID = str出诊ID & "," & mlng跨月周出诊ID
        If gVisitPlan_ModulePara.byt出诊表打印方式 = 1 Or bytMode = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_3", Me, "出诊ID=" & str出诊ID, 2)
        ElseIf gVisitPlan_ModulePara.byt出诊表打印方式 = 2 Then
            If MsgBox("要打印出诊表吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_3", Me, "出诊ID=" & str出诊ID, 2)
            End If
        End If
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecuteFilter()
    '过滤数据
    Dim strKey As String
    
    Err = 0: On Error GoTo errHandler
    Call zlControl.TxtSelAll(txtFind(1))
    
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
    If mintFindType = 8 Then mintFindType = 0 '清除
    Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, mbytFun, , , _
        Val(lblPublishInfo.Tag) = 1, Format(mdtStartDate, "yyyy-mm-dd"), Format(mdtEndDate, "yyyy-mm-dd"))
    Exit Sub
errHandler:
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
            Call RefreshOneData(i, i = lngRowStart)
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

Private Function DeletePlan(ByVal lng出诊ID As Long, ByVal byt排班方式 As Byte, ByVal intWeek As Integer, _
    ByVal lng跨月周出诊ID As Long, ByVal intElseWeek As Integer) As Boolean
    '功能：删除出诊表
    '入参：
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String
    Dim strElsePlanName As String
    Dim strToolTip  As String
    Dim lngArray出诊ID(1) As Long, i As Integer
    
    Err = 0: On Error GoTo errHandler
    strSQL = "Select 1 From 临床出诊表 Where ID = [1] And 发布时间 Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng出诊ID)
    If rsTemp.EOF Then
        MsgBox "当前出诊表已被他人删除或已发布，不能删除！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '整周跨月的周出诊表的同步处理
    If byt排班方式 = 2 And lng跨月周出诊ID <> 0 Then
        strSQL = "Select 出诊表名,发布时间 From 临床出诊表 Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng跨月周出诊ID)
        If rsTemp.EOF Then
            lng跨月周出诊ID = 0
        Else
            If Nvl(rsTemp!发布时间) <> "" Then
                MsgBox "当前出诊表所在星期内的另一个出诊表已发布，不能删除！", vbInformation, gstrSysName
                Exit Function
            End If
            strElsePlanName = Nvl(rsTemp!出诊表名)
        End If
    End If
    
    strSQL = "Select ID" & vbNewLine & _
            " From (Select a.Id" & vbNewLine & _
            "       From 临床出诊表 A, 临床出诊安排 B" & vbNewLine & _
            "       Where a.排班方式 = [1] And a.Id = b.出诊id(+) And Nvl(a.站点,'-') = Nvl([2],'-')" & vbNewLine & _
            "       Order By a.年份 Desc, a.月份 Desc, a.周数 Desc)" & vbNewLine & _
            " Where Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, byt排班方式, gstrNodeNo)
    If rsTemp.EOF Then
        MsgBox "当前出诊表不存在，可能已被他人删除，请刷新查看！", vbInformation, gstrSysName
        Exit Function
    Else
        If Val(Nvl(rsTemp!ID)) <> lng出诊ID _
            And (byt排班方式 = 1 Or byt排班方式 = 2 And Val(Nvl(rsTemp!ID)) <> lng跨月周出诊ID) Then
            MsgBox "删除失败，你只能从最后一个未发布的出诊表开始删除！", vbInformation, gstrSysName: Exit Function
            Exit Function
        End If
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
        '如果出诊表中含有其他科室人员的排班就不能删除该出诊表
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊号源 B" & vbNewLine & _
                " Where a.号源id = b.Id And a.出诊id In ([1],[2])" & vbNewLine & _
                "       And Not (Nvl(b.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = b.科室id And 人员id = [3]))"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng出诊ID, lng跨月周出诊ID, UserInfo.ID)
        If Not rsTemp.EOF Then
            MsgBox "当前出诊表中含有其它人员已经制定的安排，不能删除！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strToolTip = "【" & Mid(sccTitle.Caption, InStr(sccTitle.Caption, ">") + 1) & "】"
    If byt排班方式 = 2 And lng跨月周出诊ID <> 0 Then
        If intWeek > intElseWeek Then
            strToolTip = strToolTip & "和【" & strElsePlanName & "】"
            lngArray出诊ID(0) = lng跨月周出诊ID
            lngArray出诊ID(1) = lng出诊ID
        Else
            strToolTip = "【" & strElsePlanName & "】和" & strToolTip
            lngArray出诊ID(0) = lng出诊ID
            lngArray出诊ID(1) = lng跨月周出诊ID
        End If
    End If
    If MsgBox("你确定要删除" & strToolTip & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Function
    End If
    
    '删除出诊表
    If byt排班方式 = 2 And lng跨月周出诊ID <> 0 Then
        On Error GoTo TransErrHandler
        gcnOracle.BeginTrans
            For i = 0 To UBound(lngArray出诊ID)
                If lngArray出诊ID(i) <> 0 Then
                    'Zl_临床出诊表_Delete
                    strSQL = "Zl_临床出诊表_Delete("
                    '  Id_In       临床出诊表.Id%Type
                    strSQL = strSQL & "" & lngArray出诊ID(i) & ","
                    '  人员id_In 人员表.Id%Type := Null
                    strSQL = strSQL & "" & IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID) & ","
                    '  站点_In   部门表.站点%Type
                    strSQL = strSQL & "'" & gstrNodeNo & "')"
                    zlDatabase.ExecuteProcedure strSQL, "删除出诊表"
                End If
            Next
        gcnOracle.CommitTrans
        DeletePlan = True
    Else
        DeletePlan = ZlDeletePlan(lng出诊ID, IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID))
    End If
    Exit Function
TransErrHandler:
    gcnOracle.RollbackTrans
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub RefreshData(ByVal bytFun As Byte, ByVal lng出诊ID As Long, Optional ByVal blnClear As Boolean, _
    Optional ByVal intYear As Integer, Optional ByVal intMonth As Integer, Optional ByVal strTitle As String)
    '功能：刷新安排详情数据
    '入参：
    '   bytFun - 1-月安排，2-周安排
    '   lng出诊ID - 出诊ID
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim i As Integer, varDateRange As Variant
    Dim dtStartDate As Date, dtEndDate As Date
    Dim intWeek As Integer
    
    Err = 0: On Error GoTo errHandler
    
    mdtToday = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
    If blnClear Then
        mbytFun = bytFun: mlng出诊ID = lng出诊ID
        sccTitle.Caption = "出诊安排>" & IIf(strTitle = "", "出诊安排", strTitle) & IIf(lng出诊ID = 0, "(无出诊表)", "")
        
        chkShowAllPlan.Value = vbUnchecked
        chkShowAllPlan.Visible = bytFun = 1
        picSelectWeek.Visible = bytFun = 1
        For i = optWeek.LBound To optWeek.UBound
            optWeek(i).Visible = True
        Next
        optWeek(0).Value = True
        
        lblPublishInfo.Visible = chkShowAllPlan.Value = vbUnchecked
        lblPublishInfo.Tag = ""
        Set mrsPlanRecords = Nothing
        mlngCopyPlanID = 0: mstrCopyPlanItem = ""
        mintYear = 0: mintMonth = 0: mintWeek = 0
        mlng跨月周出诊ID = 0: mintElseWeek = 0
        
        '改变菜单名称
        Call ZlUpdatePlanMenu(Me, mcbsMain, bytFun, IIf(HavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID))
        
        strSQL = "Select b.出诊表名, b.年份, b.月份, b.周数, b.发布人, b.发布时间" & vbNewLine & _
                " From 临床出诊表 B" & vbNewLine & _
                " Where b.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊表信息", lng出诊ID)
        If Not rsTemp.EOF Then
            sccTitle.Caption = "出诊安排>" & Nvl(rsTemp!出诊表名)
            lblPublishInfo.Tag = IIf(Nvl(rsTemp!发布时间) = "", "", "1") '标记是否发布
            lblPublishInfo.Caption = "发布人：" & IIf(Nvl(rsTemp!发布人) = "", "      ", Nvl(rsTemp!发布人)) & _
                "  发布时间：" & IIf(Nvl(rsTemp!发布时间) = "", "                   ", Format(Nvl(rsTemp!发布时间), "yyyy-mm-dd hh:mm:ss"))
            mintYear = Val(Nvl(rsTemp!年份))
            mintMonth = Val(Nvl(rsTemp!月份))
            mintWeek = Val(Nvl(rsTemp!周数))
        End If
        If mintYear = 0 Then mintYear = intYear
        If mintMonth = 0 Then mintMonth = intMonth
        
        If mintYear = 0 Then mintYear = Year(mdtToday)
        If mintMonth = 0 Then mintMonth = Month(mdtToday)
        If mintWeek = 0 Then mintWeek = GetDateWeek(mdtToday)
        
        '周数确定
        For i = GetWeekCount(mintYear, mintMonth) + 1 To optWeek.UBound
            optWeek(i).Visible = False
        Next
        
        '显示日期范围确定
        varDateRange = GetDateRange(mintYear, mintMonth, IIf(bytFun = 2, mintWeek, 0))
        mdtStartDate = varDateRange(0): mdtEndDate = varDateRange(1)
        
        dtStartDate = mdtStartDate: dtEndDate = mdtEndDate
        If bytFun = 2 Then
            If IsDoubleMonthWeekPlan(intYear, intMonth, intWeek, dtStartDate, dtEndDate) Then
                mintElseWeek = intWeek
                mlng跨月周出诊ID = Get周出诊表ID(intYear, intMonth, intWeek)
            End If
        End If
        
        Call InitPlanGrid(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, dtStartDate, dtEndDate, Val(lblPublishInfo.Tag) = 1)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "安排")
        Call ShowHolidayToPlan(vsfRegistPlan, dtStartDate, dtEndDate)
        Call Form_Resize
    End If
    
    Screen.MousePointer = vbHourglass
    If lng出诊ID = 0 And chkShowAllPlan.Value = vbUnchecked Then
        Set mrsPlanRecords = Nothing
    Else
        Set mrsPlanRecords = GetPlanRecords(bytFun = 1, lng出诊ID, Val(lblPublishInfo.Tag) = 1, _
            chkShowAllPlan.Value = vbChecked, mintYear, mintMonth, mdtStartDate, mdtEndDate, mlng跨月周出诊ID)
    End If
    '加载数据
    Call ExecuteFilter
'    Call ShowStopVisitPlan(vsfRegistPlan, mdtStartDate, mdtEndDate)
    '定位到当前日期
    If bytFun = 1 And (mdtToday >= mdtStartDate And mdtStartDate <= mdtEndDate) Then
        For i = gPlanGrid_FixedCols To vsfRegistPlan.Cols - 1 Step 3
            If CDate(vsfRegistPlan.Cell(flexcpData, 0, i)) = mdtToday Then
                vsfRegistPlan.LeftCol = i
                vsfRegistPlan.Col = i
                If vsfRegistPlan.Rows > 2 Then vsfRegistPlan.Row = 3
                Exit For
            End If
        Next
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshOneData(Optional ByVal lngCurRow As Long = -1, _
    Optional ByVal blnReLoadData As Boolean = True)
    '刷新指定行号源数据
    Dim lng号源Id As Long, str收费项目 As String
    
    Err = 0: On Error GoTo errHandle
    '1.记录原数据，并获取新数据
    With vsfRegistPlan
        lng号源Id = Val(.TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_号源ID))
        str收费项目 = .TextMatrix(IIf(lngCurRow = -1, .Row, lngCurRow), COL_项目)
    End With
    
    If blnReLoadData Then
        '更新本地记录集
        Set mrsPlanRecords = GetPlanRecords(mbytFun = 1, mlng出诊ID, Val(lblPublishInfo.Tag) = 1, _
            chkShowAllPlan.Value = vbChecked, mintYear, mintMonth, mdtStartDate, mdtEndDate, mlng跨月周出诊ID)
    End If
    
    '2.更新界面
    mrsPlanRecords.Filter = "号源ID=" & lng号源Id & " And 收费项目='" & str收费项目 & "'"
    Call RefreshOnePlanData(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, lngCurRow, _
        Val(lblPublishInfo.Tag) = 1, mbytFun, Format(mdtStartDate, "yyyy-mm-dd"), Format(mdtEndDate, "yyyy-mm-dd"))
'    Call ShowStopVisitPlan(vsfRegistPlan, mdtStartDate, mdtEndDate, lng号源Id)
    mrsPlanRecords.Filter = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function AddNewPlan(Optional blnMonth As Boolean) As String
    '功能：新增出诊表
    '入参：
    '   blnMonth 是否按月排班
    Dim strSQL As String, rsTemp As ADODB.Recordset, lng出诊ID As Long
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim strName As String, strKey As String, blnDeletePlan As Boolean
    Dim cllPlan As Collection, i As Integer
    Dim dtCurrent As Date
    
    Err = 0: On Error GoTo errHandler
    Set cllPlan = GetNewPlanInfo(Me, mstrPrivs, blnMonth, strKey, blnDeletePlan)
    If cllPlan Is Nothing Then Exit Function
    If cllPlan.Count = 0 Then Exit Function
    
    dtCurrent = zlDatabase.Currentdate
    On Error GoTo TransErrHandler
    If cllPlan.Count > 1 Then gcnOracle.BeginTrans
    For i = 1 To cllPlan.Count
        'Array(年份,月份,周数,开始日期,结束日期)
        intYear = cllPlan(i)(0)
        intMonth = cllPlan(i)(1)
        intWeek = cllPlan(i)(2)
        dtStart = cllPlan(i)(3)
        dtEnd = cllPlan(i)(4)
        
        '确定出诊表ID
        strSQL = "Select ID From 临床出诊表 Where 排班方式 = [1] And 年份 = [2] And 月份 = [3]" & _
            IIf(blnMonth, "", " And 周数 = [4]") & " And Nvl(站点,'-') = Nvl([5],'-')"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(blnMonth, 1, 2), intYear, intMonth, intWeek, gstrNodeNo)
        If rsTemp.EOF Then
            lng出诊ID = zlDatabase.GetNextId("临床出诊表")
        Else
            lng出诊ID = Val(Nvl(rsTemp!ID))
        End If
        
        '出诊表名
        strName = intYear & "年" & intMonth & "月"
        If Not blnMonth Then strName = strName & "第" & intWeek & "周"
        strName = strName & "出诊表"
        
        'Zl_临床出诊表_Add(
        strSQL = "Zl_临床出诊表_Add("
        '  操作类型_In Number,--1-模板，2-固定安排, 3-月安排，4-周安排
        strSQL = strSQL & "" & IIf(blnMonth, 3, 4) & ","
        '  出诊id_In   临床出诊表.Id%Type,
        strSQL = strSQL & "" & lng出诊ID & ","
        '  出诊表名_In 临床出诊表.出诊表名%Type,
        strSQL = strSQL & "'" & strName & "',"
        '  站点_In     部门表.站点%Type,
        strSQL = strSQL & "'" & gstrNodeNo & "',"
        '  全院号源归属站点_In 部门表.站点%Type,
        strSQL = strSQL & "'" & gVisitPlan_ModulePara.str号源维护站点 & "',"
        '  操作员_In   临床出诊安排.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  操作时间_In 临床出诊安排.登记时间%Type := Null
        strSQL = strSQL & "" & ZDate(dtCurrent) & ","
        '  开始时间_In 临床出诊安排.开始时间%Type := Null,
        strSQL = strSQL & "" & ZDate(dtStart) & ","
        '  终止时间_In 临床出诊安排.终止时间%Type := Null,
        strSQL = strSQL & "" & ZDate(dtEnd) & ","
        '  年份_In     临床出诊表.年份%Type := Null,
        strSQL = strSQL & "" & intYear & ","
        '  月份_In     临床出诊表.月份%Type := Null,
        strSQL = strSQL & "" & intMonth & ","
        '  周数_In     临床出诊表.周数%Type := Null,
        strSQL = strSQL & "" & ZVal(intWeek) & ","
        '  应用范围_In 临床出诊表.应用范围%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  科室id_In   临床出诊表.科室id%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  备注_In     临床出诊表.备注%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '  人员id_In   人员表.Id%Type := Null,
        strSQL = strSQL & "" & IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室"), "NULL", UserInfo.ID) & ","
        '  删除安排_In Number:=0
        strSQL = strSQL & "" & IIf(blnDeletePlan, 1, 0) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Next
    If cllPlan.Count > 1 Then gcnOracle.CommitTrans
    
    'XX月出诊表节点：K2_年份_月份
    'XX周出诊表节点：K3_年份_月份_周数
    AddNewPlan = strKey
    Exit Function
TransErrHandler:
    If cllPlan.Count > 1 Then gcnOracle.RollbackTrans
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub chkShowAllPlan_Click()
    Err = 0: On Error GoTo errHandler
    
    lblPublishInfo.Visible = chkShowAllPlan.Value = vbUnchecked
    Screen.MousePointer = vbHourglass
    If mlng出诊ID = 0 And chkShowAllPlan.Value = vbUnchecked Then
        Set mrsPlanRecords = Nothing
    Else
        Set mrsPlanRecords = GetPlanRecords(mbytFun = 1, mlng出诊ID, Val(lblPublishInfo.Tag) = 1, _
            chkShowAllPlan.Value = vbChecked, mintYear, mintMonth, mdtStartDate, mdtEndDate)
    End If
    Call ExecuteFilter
'    Call ShowStopVisitPlan(vsfRegistPlan, mdtStartDate, mdtEndDate)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandler
    
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "安排")
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
    
    picSelectWeek.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, sccTitle.Width
    lineSplit.X1 = sccTitle.Left + 10
    lineSplit.Y1 = IIf(picSelectWeek.Visible, picSelectWeek.Top + picSelectWeek.Height, sccTitle.Top + sccTitle.Height - 10)
    lineSplit.X2 = sccTitle.Width
    lineSplit.Y2 = lineSplit.Y1
    With vsfRegistPlan
        .Left = sccTitle.Left + 10
        .Top = IIf(picSelectWeek.Visible, picSelectWeek.Top + picSelectWeek.Height + 20, sccTitle.Top + sccTitle.Height + 10)
        .Width = sccTitle.Width
        .Height = Me.ScaleHeight - .Top - 20
    End With
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
    
    objOut.Title.Text = Mid(sccTitle.Caption, InStr(sccTitle.Caption, ">") + 1) & "清单"
    If VSFlexGridCopyTo(vsfRegistPlan, vsfTemp, bytMode) = False Then Exit Sub
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

Private Sub Form_Unload(Cancel As Integer)
    Set mrsPlanRecords = Nothing
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistPlan, Me.Name, "安排")
    Call SaveRegInFor(g私有模块, Me.Name, "FindType", mintFindType)
End Sub

Private Sub optWeek_Click(index As Integer)
    Dim varDateRange As Variant, intWeek As Integer
    Dim dtStart As Date, dtEnd As Date
    
    Err = 0: On Error GoTo errHandler
    intWeek = index
    Screen.MousePointer = vbHourglass
    varDateRange = GetDateRange(mintYear, mintMonth, intWeek)
    dtStart = varDateRange(0): dtEnd = varDateRange(1)
    Call InitPlanGrid(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, dtStart, dtEnd, Val(lblPublishInfo.Tag) = 1)
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "安排")
    Call ShowHolidayToPlan(vsfRegistPlan, mdtStartDate, mdtEndDate)
    '使用缓存数据
    Call ExecuteFilter
'    Call ShowStopVisitPlan(vsfRegistPlan, mdtStartDate, mdtEndDate)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If vsfRegistPlan.Visible And vsfRegistPlan.Enabled Then vsfRegistPlan.SetFocus
End Sub

Private Sub vsfRegistPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim cbrControl As CommandBarControl, lngCol As Long, dtCur As Date
    Dim strTemp As String, lng记录ID As Long
    
    On Error Resume Next
    '菜单控制
    strTemp = Trim(vsfRegistPlan.Cell(flexcpData, 0, vsfRegistPlan.Col))
    If strTemp = "" Then Exit Sub
    If IsDate(strTemp) = False Then Exit Sub
    dtCur = CDate(strTemp)
    Set cbrControl = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_ApplyToDay, , True) '应用于单日/双日
    If Not cbrControl Is Nothing Then
        cbrControl.Caption = "应用于“所有" & IIf(Day(dtCur) Mod 2 = 0, "双日", "单日") & "”(&D)"
    End If
    Set cbrControl = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_ApplyToWeekDay, , True) '应用于星期几
    If Not cbrControl Is Nothing Then
        cbrControl.Caption = "应用于“所有" & GetWeekName(Weekday(dtCur, vbMonday) - 1) & "”(&W)"
    End If
    
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
    Dim lng号源Id As Long, lng安排ID As Long, lng出诊ID As Long
    Dim frmEdit As New frmClinicPlanEdit
    Dim strCurItem As String, blnUpdate As Boolean
    Dim lngCol As Long, lngRow As Long
    Dim strSort As String
    Dim strTemp As String
    
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
            Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, mbytFun, , True, _
                Val(lblPublishInfo.Tag) = 1, Format(mdtStartDate, "yyyy-mm-dd"), Format(mdtEndDate, "yyyy-mm-dd"))
'            Call ShowStopVisitPlan(vsfRegistPlan, mdtStartDate, mdtEndDate)
            Screen.MousePointer = vbDefault
        End If
    Else
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        lng安排ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_安排ID))
        lngCol = GetPlanItemNameCol(vsfRegistPlan.Col)
        strCurItem = vsfRegistPlan.Cell(flexcpData, 0, lngCol)
        If lng号源Id = 0 And lng安排ID = 0 Then Exit Sub
        If IsDate(strCurItem) = False Then Exit Sub
        If Not (strCurItem >= mdtStartDate And strCurItem <= mdtEndDate) Then Exit Sub
        
        If chkShowAllPlan.Value = vbChecked Then
            lng出诊ID = 0: lng安排ID = 0
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) = "" Then Exit Sub
            '存储了“出诊ID,安排ID”
            strTemp = vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col) + 2)
            If InStr(strTemp, ",") = 0 Then Exit Sub
            lng出诊ID = Val(Split(strTemp, ",")(0))
            lng安排ID = Val(Split(strTemp, ",")(1))
        Else
            lng出诊ID = mlng出诊ID
        End If
        
        blnUpdate = zlStr.IsHavePrivs(mstrPrivs, "出诊安排") And Val(lblPublishInfo.Tag) = 0
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_是否临床排班)) = "" Then blnUpdate = False
        End If
        '显示所有安排时，只能查看
        If chkShowAllPlan.Value = vbChecked Then blnUpdate = False
    
        If frmEdit.ShowMe(Me, mbytFun, IIf(blnUpdate, Fun_Update, Fun_View), lng出诊ID, lng号源Id, lng安排ID, strCurItem) Then
            If blnUpdate Then Call RefreshOneData
        End If
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

Private Function PublishPlan(ByVal lng出诊ID As Long, ByVal blnPublish As Boolean, _
    ByVal lng跨月周出诊ID As Long) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim obj上班时段 As 上班时段
    Dim dtCurrent As Date, cll出诊ID  As New Collection, i As Integer
    Dim strPlanName As String
    
    Err = 0: On Error GoTo errHandler
    If blnPublish Then
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊记录 B, 临床出诊表 C" & vbNewLine & _
                " Where a.Id = b.安排id And a.出诊id = c.Id And c.排班方式 In (1, 2)" & vbNewLine & _
                "       And c.Id In ([1], [2]) And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng出诊ID, lng跨月周出诊ID)
        If rsTemp.EOF Then
            MsgBox "当前出诊表无有效的安排，不能发布！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select a.Id" & vbNewLine & _
                " From (Select ID, 年份 || LPad(月份, 2, '0') || 周数 As 日期" & vbNewLine & _
                "        From 临床出诊表" & vbNewLine & _
                "        Where Nvl(排班方式, 0) = [2] And 发布人 Is Null And Id Not In ([1], [3])" & vbNewLine & _
                "              And Nvl(站点,'-') = Nvl([4],'-')) A," & vbNewLine & _
                "      (Select ID, 年份 || LPad(月份, 2, '0') || 周数 As 日期 From 临床出诊表 Where ID = [1]) B" & vbNewLine & _
                " Where a.日期 < b.日期 And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng出诊ID, mbytFun, lng跨月周出诊ID, gstrNodeNo)
        If Not rsTemp.EOF Then
            MsgBox "当前出诊表前面还有未发布的" & IIf(mbytFun = 1, "月", "周") & "出诊表，必须先将其发布后才能发布该出诊表！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊记录 A, 临床出诊安排 B" & vbNewLine & _
                " Where a.号源id = b.号源id And a.出诊日期 Between b.开始时间 And b.终止时间" & vbNewLine & _
                "       And a.安排ID <> b.Id And b.出诊id In ([1],[2]) And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng出诊ID, lng跨月周出诊ID)
        If Not rsTemp.EOF Then
            MsgBox "当前出诊表中的部分号源在当前出诊表的生效时间范围内已经存在有效的安排，不能发布！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select Distinct d.号类, d.号码, e.站点, b.出诊日期, b.上班时段, To_Char(c.开始时间, 'hh24:mi:ss') As 开始时间, " & vbNewLine & _
                "       To_Char(c.终止时间, 'hh24:mi:ss') As 终止时间" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊记录 B, 临床出诊序号控制 C, 临床出诊号源 D, 部门表 E" & vbNewLine & _
                " Where a.Id = b.安排id And b.Id = c.记录id And a.号源id = d.Id And d.科室id = e.Id" & vbNewLine & _
                "       And c.序号 = 1 And 出诊id In ([1],[2])" & vbNewLine & _
                " Order By " & IIf(gVisitPlan_ModulePara.byt号码比较方式 = 0, "d.号码", "Lpad(d.号码,5,'0')")
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查上班时段与序号分时段时间是否一致", lng出诊ID, lng跨月周出诊ID)
        Do While Not rsTemp.EOF
            Set obj上班时段 = GetWorkTimeRange(Nvl(rsTemp!上班时段), Nvl(rsTemp!站点), Nvl(rsTemp!号类))
            If Format(obj上班时段.开始时间, "hh:mm:00") <> Format(Nvl(rsTemp!开始时间), "hh:mm:00") Then
                If MsgBox("当前出诊表中部分分时段的安排不是根据上班时段的时间进行分段的，如：" & vbCrLf & _
                    "号码为 " & Nvl(rsTemp!号码) & " ，" & Format(Nvl(rsTemp!出诊日期), "yyyy-mm-dd") & " " & Nvl(rsTemp!上班时段) & _
                    "[" & Format(obj上班时段.开始时间, "hh:mm") & "-" & Format(obj上班时段.结束时间, "hh:mm") & "]，" & _
                    "第一个序号时段为[" & Format(Nvl(rsTemp!开始时间), "hh:mm") & "-" & Format(Nvl(rsTemp!终止时间), "hh:mm") & "])" & vbCrLf & vbCrLf & _
                    "是否仍要继续发布？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                Exit Do
            End If
            rsTemp.MoveNext
        Loop
        
        strSQL = "Select Id, 出诊表名" & vbNewLine & _
                " From 临床出诊表" & vbNewLine & _
                " Where Id In([1],[2]) And 发布时间 Is Null" & _
                " Order By 年份,月份,周数"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊表信息", lng出诊ID, lng跨月周出诊ID)
        If rsTemp.EOF Then
            MsgBox "当前出诊表可能已被他人发布或已删除，请刷新数据后查看！", vbInformation, gstrSysName
            Exit Function
        End If
        Do While Not rsTemp.EOF
            cll出诊ID.Add Val(Nvl(rsTemp!ID))
            strPlanName = strPlanName & IIf(strPlanName <> "", "和", "") & "【" & Nvl(rsTemp!出诊表名) & "】"
            rsTemp.MoveNext
        Loop
        
        If MsgBox("你确定要发布" & strPlanName & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        
        dtCurrent = zlDatabase.Currentdate
        On Error GoTo TransErrHandler
        
        Screen.MousePointer = vbHourglass
        If cll出诊ID.Count > 1 Then gcnOracle.BeginTrans
        For i = 1 To cll出诊ID.Count
            'Zl_临床出诊安排_Publish
            strSQL = "Zl_临床出诊安排_Publish("
            '  Id_In       临床出诊表.Id%Type,
            strSQL = strSQL & "" & cll出诊ID(i) & ","
            '  发布人_In   临床出诊表.发布人%Type := Null,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  发布时间_In 临床出诊表.发布时间%Type := Null,
            strSQL = strSQL & "" & ZDate(dtCurrent) & ")"
            '  取消发布_In Number:=0
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Next
        If cll出诊ID.Count > 1 Then gcnOracle.CommitTrans
        Screen.MousePointer = vbDefault
    Else
        strSQL = "Select a.Id" & vbNewLine & _
                " From (Select ID, 年份 || LPad(月份, 2, '0') || 周数 As 日期" & vbNewLine & _
                "        From 临床出诊表" & vbNewLine & _
                "        Where Nvl(排班方式, 0) = [2] And 发布人 Is Not Null And Id Not In ([1], [3])" & vbNewLine & _
                "              And Nvl(站点,'-') = Nvl([4],'-')) A," & vbNewLine & _
                "      (Select ID, 年份 || LPad(月份, 2, '0') || 周数 As 日期 From 临床出诊表 Where ID = [1]) B" & vbNewLine & _
                " Where a.日期 > b.日期 And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng出诊ID, mbytFun, lng跨月周出诊ID, gstrNodeNo)
        If Not rsTemp.EOF Then
            MsgBox "当前出诊后面还有已发布的" & IIf(mbytFun = 1, "月", "周") & "出诊表，必须先将其取消发布后才能取消发布该出诊表！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1" & vbNewLine & _
                " From 病人挂号记录 C, 临床出诊记录 A, 临床出诊安排 B" & vbNewLine & _
                " Where c.出诊记录id = a.Id And a.安排id = b.Id And b.出诊id In ([1],[2]) And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng出诊ID, lng跨月周出诊ID)
        If Not rsTemp.EOF Then
            MsgBox "当前出诊表所在周的安排已被使用，不允许取消发布！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select Id, 出诊表名" & vbNewLine & _
                " From 临床出诊表" & vbNewLine & _
                " Where Id In([1],[2]) And 发布时间 Is Not Null" & _
                " Order By 年份 Desc,月份 Desc,周数 Desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊表信息", lng出诊ID, lng跨月周出诊ID)
        If rsTemp.EOF Then
            MsgBox "当前出诊表可能已被他人取消发布或已删除，请刷新数据后查看！", vbInformation, gstrSysName
            Exit Function
        End If
        Do While Not rsTemp.EOF
            cll出诊ID.Add Val(Nvl(rsTemp!ID))
            strPlanName = "【" & Nvl(rsTemp!出诊表名) & "】" & IIf(strPlanName <> "", "和", "") & strPlanName
            rsTemp.MoveNext
        Loop
        
        If MsgBox("你确定要取消发布" & strPlanName & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        dtCurrent = zlDatabase.Currentdate
        On Error GoTo TransErrHandler
        
        Screen.MousePointer = vbHourglass
        If cll出诊ID.Count > 1 Then gcnOracle.BeginTrans
        For i = 1 To cll出诊ID.Count
            'Zl_临床出诊安排_Publish
            strSQL = "Zl_临床出诊安排_Publish("
            '  Id_In       临床出诊表.Id%Type,
            strSQL = strSQL & "" & cll出诊ID(i) & ","
            '  发布人_In   临床出诊表.发布人%Type := Null,
            strSQL = strSQL & "" & "NULL" & ","
            '  发布时间_In 临床出诊表.发布时间%Type := Null,
            strSQL = strSQL & "" & "NULL" & ","
            '  取消发布_In Number:=0
            strSQL = strSQL & "" & 1 & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Next
        If cll出诊ID.Count > 1 Then gcnOracle.CommitTrans
        Screen.MousePointer = vbDefault
    End If
    PublishPlan = True
    Exit Function
TransErrHandler:
    If cll出诊ID.Count > 1 Then gcnOracle.RollbackTrans
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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

Public Function PastPlan(ByVal lng出诊ID As Long, ByVal lng原安排ID As Long, ByVal str原项目 As String) As Long
    '功能：粘贴安排
    '参数：
    Dim strSQL As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim rsPlan As ADODB.Recordset, rsSignalSource As ADODB.Recordset
    Dim blnTran As Boolean, blnNoPlan As Boolean
    Dim lng安排ID  As Long, lng号源Id As Long, strApplyItem  As String
    Dim str号类 As String
    
    Err = 0: On Error GoTo errHandler
    If lng出诊ID = 0 Then Exit Function
    If lng原安排ID = 0 Then Exit Function
    If IsDate(str原项目) = False Then Exit Function
    
    With vsfRegistPlan
        lng安排ID = Val(.TextMatrix(.Row, COL_安排ID))
        lng号源Id = Val(.TextMatrix(.Row, COL_号源ID))
        strApplyItem = .Cell(flexcpData, 0, .Col)
        str号类 = Trim(.TextMatrix(.Row, COL_号类))
    End With
    
    If lng号源Id = 0 Then Exit Function
    If IsDate(strApplyItem) = False Then Exit Function
    
    If lng原安排ID = lng安排ID And CDate(str原项目) = CDate(strApplyItem) Then
        MsgBox "当前安排与复制安排相同，不能粘贴！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '检查某个上班时段是否适用于当前号源
    strSQL = "Select b.号类, a.上班时段" & vbNewLine & _
            " From 临床出诊记录 A, 临床出诊号源 B" & vbNewLine & _
            " Where a.号源ID = b.ID And a.安排ID = [1] And a.出诊日期 = [2] And a.上班时段 Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng原安排ID, CDate(str原项目))
    Do While Not rsTemp.EOF
        If GetWorkTimeRange(Nvl(rsTemp!上班时段), gstrNodeNo, str号类) Is Nothing Then
            MsgBox "上班时段“" & Nvl(rsTemp!上班时段) & "”不适用于" & str号类 & "号，不能粘贴！", vbInformation, gstrSysName
            Exit Function
        End If
        rsTemp.MoveNext
    Loop
    
    If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) <> "" Then
        If MsgBox("被粘贴的日期当前已存在出诊安排，粘贴后这部分安排将会被覆盖！是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    If lng安排ID = 0 Then
        blnNoPlan = True
        
        '检查当前日期是否已由其它出诊表生成,一个号源某一天的安排只能由一个出诊表设置
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊记录 A" & vbNewLine & _
                " Where a.出诊日期 = [1] And A.号源Id = [2] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strApplyItem), lng号源Id)
        If Not rsTemp.EOF Then
            MsgBox Format(strApplyItem, "yyyy-mm-dd") & " 已在其它出诊表中进行了安排，不能重复安排！", vbInformation, gstrSysName
            Exit Function
        End If
        
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
        strSQL = strSQL & "" & ZDate(mdtStartDate) & ","
        '终止时间_In     临床出诊安排.终止时间%Type,
        strSQL = strSQL & "" & ZDate(mdtEndDate) & ","
        '操作员姓名_In   临床出诊安排.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '登记时间_In     临床出诊安排.登记时间%Type
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
    Else
        '检查当前日期是否已由其它出诊表生成,一个号源某一天的安排只能由一个出诊表设置
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊记录 A" & vbNewLine & _
                " Where a.出诊日期 = [1] And a.号源Id = [2] And a.安排id <> [3] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strApplyItem), lng号源Id, lng安排ID)
        If Not rsTemp.EOF Then
            MsgBox Format(strApplyItem, "yyyy-mm-dd") & " 已在其它出诊表中进行了安排，不能重复安排！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    blnTran = True
    gcnOracle.BeginTrans
        If blnNoPlan And strSQL <> "" Then
            zlDatabase.ExecuteProcedure strSQL, "新增安排"
        End If
        If ZlPlanApplyTo(1, lng原安排ID, str原项目, lng安排ID, strApplyItem) = False Then
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

Private Function ApplyToDay(ByVal lng安排ID As Long, ByVal strCurDate As String) As Boolean
    '功能：应用于“所有单日/双日”
    '参数：
    '   lng安排ID 被应用的安排ID
    '   dtCurDate 被应用的日期
    Dim strApply As String, dtCur As Date
    Dim intDoubleDay As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim lng号源Id As Long
    
    Err = 0: On Error GoTo errHandler
    If lng安排ID = 0 Or strCurDate = "" Then Exit Function
    If IsDate(strCurDate) = False Then Exit Function
    
    dtStart = CDate(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_开始时间))
    dtEnd = CDate(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_终止时间))
    lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
    
    '查询号源非当前出诊表设置的出诊记录
    strSQL = "Select a.出诊日期" & vbNewLine & _
            " From 临床出诊记录 A,临床出诊安排 B,临床出诊表 C" & vbNewLine & _
            " Where a.安排ID=b.ID And b.出诊ID<>[1] And a.号源ID=[2] And a.出诊日期 Between [3] And [4]" & vbNewLine & _
            "       And c.ID=b.出诊ID And Nvl(c.排班方式,0) In (1,2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取号源已设置了的出诊记录", mlng出诊ID, lng号源Id, dtStart, dtEnd)
    
    intDoubleDay = Day(strCurDate) Mod 2 '单日还是双日
    dtCur = dtStart
    Do While DateDiff("d", dtCur, dtEnd) >= 0
        If DateDiff("d", strCurDate, dtCur) <> 0 And (Day(dtCur) Mod 2) = intDoubleDay Then
            rsTemp.Filter = "出诊日期=#" & Format(dtCur, "yyyy-mm-dd") & "#"
            If rsTemp.RecordCount = 0 Then
                strApply = strApply & "|" & Format(dtCur, "yyyy-mm-dd")
            End If
        End If
        dtCur = DateAdd("d", 1, dtCur)
    Loop
    If strApply <> "" Then strApply = Mid(strApply, 2)
    
    If strApply = "" Then Exit Function
    If CheckExistRecord(lng号源Id, strApply) Then
        If MsgBox("注意：" & vbCrLf & _
                  "      部分被应用的日期当前已存在出诊安排，应用后这部分安排将会被覆盖！是否仍要应用？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    ApplyToDay = ZlPlanApplyTo(1, lng安排ID, Format(strCurDate, "yyyy-mm-dd"), lng安排ID, strApply)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
        
Private Function ApplyToWeekDay(ByVal lng安排ID As Long, ByVal strCurDate As String) As Boolean
    '功能：应用于“所有星期几”
    '参数：
    '   lng安排ID 被应用的安排ID
    '   dtCurDate 被应用的日期
    Dim strApply As String, dtCur As Date
    Dim intWeekDay As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim lng号源Id As Long
    
    Err = 0: On Error GoTo errHandler
    If lng安排ID = 0 Or strCurDate = "" Then Exit Function
    If IsDate(strCurDate) = False Then Exit Function
    
    dtStart = CDate(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_开始时间))
    dtEnd = CDate(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_终止时间))
    lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
    
    '查询号源非当前出诊表设置的出诊记录
    strSQL = "Select a.出诊日期" & vbNewLine & _
            " From 临床出诊记录 A,临床出诊安排 B,临床出诊表 C" & vbNewLine & _
            " Where a.安排ID=b.ID And b.出诊ID<>[1] And a.号源ID=[2] And a.出诊日期 Between [3] And [4]" & vbNewLine & _
            "       And c.ID=b.出诊ID And Nvl(c.排班方式,0) In (1,2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取号源已设置了的出诊记录", mlng出诊ID, lng号源Id, dtStart, dtEnd)
    
    intWeekDay = Weekday(strCurDate, vbMonday) '星期几
    dtCur = dtStart
    Do While DateDiff("d", dtCur, dtEnd) >= 0
        If DateDiff("d", strCurDate, dtCur) <> 0 And Weekday(dtCur, vbMonday) = intWeekDay Then
            rsTemp.Filter = "出诊日期=#" & Format(dtCur, "yyyy-mm-dd") & "#"
            If rsTemp.RecordCount = 0 Then
                strApply = strApply & "|" & Format(dtCur, "yyyy-mm-dd")
            End If
        End If
        dtCur = DateAdd("d", 1, dtCur)
    Loop
    If strApply <> "" Then strApply = Mid(strApply, 2)
    
    If strApply = "" Then Exit Function
    If CheckExistRecord(lng号源Id, strApply) Then
        If MsgBox("注意：" & vbCrLf & _
                  "      部分被应用的日期当前已存在出诊安排，应用后这部分安排将会被覆盖！是否仍要应用？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    ApplyToWeekDay = ZlPlanApplyTo(1, lng安排ID, Format(strCurDate, "yyyy-mm-dd"), lng安排ID, strApply)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function NextNewPlanByPlan(ByVal lng原出诊ID As Long, Optional ByVal blnMonth As Boolean) As Boolean
    '根据现有出诊表生成新安排
    Dim strSQL As String, rsTemp As ADODB.Recordset, lng出诊ID As Long
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim strName As String, strKey As String, blnDeletePlan As Boolean
    Dim cllPlan As Collection, i As Integer
    Dim dtCurrent As Date
    
    Err = 0: On Error GoTo errHandler
    Set cllPlan = GetNewPlanInfo(Me, mstrPrivs, blnMonth, strKey, blnDeletePlan)
    If cllPlan Is Nothing Then Exit Function
    If cllPlan.Count = 0 Then Exit Function
    
    dtCurrent = zlDatabase.Currentdate
    On Error GoTo TransErrHandler
        
    Screen.MousePointer = vbHourglass
    If cllPlan.Count > 1 Then gcnOracle.BeginTrans
    For i = 1 To cllPlan.Count
        'Array(年份,月份,周数,开始日期,结束日期)
        intYear = cllPlan(i)(0)
        intMonth = cllPlan(i)(1)
        intWeek = cllPlan(i)(2)
        dtStart = cllPlan(i)(3)
        dtEnd = cllPlan(i)(4)
    
        '确定出诊表ID
        strSQL = "Select ID From 临床出诊表 Where 排班方式 = [1] And 年份 = [2] And 月份 = [3]" & _
            IIf(blnMonth, "", " And 周数 = [4]") & " And Nvl(站点,'-') = Nvl([5],'-')"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(blnMonth, 1, 2), intYear, intMonth, intWeek, gstrNodeNo)
        If rsTemp.EOF Then
            lng出诊ID = zlDatabase.GetNextId("临床出诊表")
        Else
            lng出诊ID = Val(Nvl(rsTemp!ID))
        End If
        
        strName = intYear & "年" & intMonth & "月"
        If Not blnMonth Then strName = strName & "第" & intWeek & "周"
        strName = strName & "出诊表"
        
        'zl_临床出诊表_Addbyrecord(
        strSQL = "zl_临床出诊表_Addbyrecord("
        '原出诊Id_In         临床出诊表.Id%Type,
        strSQL = strSQL & "" & lng原出诊ID & ","
        '新出诊id_In      临床出诊表.Id%Type,
        strSQL = strSQL & "" & lng出诊ID & ","
        '排班方式_In   临床出诊表.排班方式%Type,
        strSQL = strSQL & "" & IIf(blnMonth, 1, 2) & ","
        '出诊表名_In   临床出诊表.出诊表名%Type,
        strSQL = strSQL & "'" & strName & "',"
        '年份_In     临床出诊表.年份%Type,
        strSQL = strSQL & "" & intYear & ","
        '月份_In     临床出诊表.月份%Type,
        strSQL = strSQL & "" & intMonth & ","
        '周数_In     临床出诊表.周数%Type := Null,
        strSQL = strSQL & "" & ZVal(intWeek) & ","
        '开始时间_In 临床出诊安排.开始时间%Type := Null,
        strSQL = strSQL & "" & ZDate(dtStart) & ","
        '终止时间_In 临床出诊安排.终止时间%Type := Null,
        strSQL = strSQL & "" & ZDate(dtEnd) & ","
        '操作员_In   临床出诊安排.操作员姓名%Type := Null,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '登记时间_In 临床出诊安排.登记时间%Type := Null,
        strSQL = strSQL & "" & ZDate(dtCurrent) & ","
        '站点_In       部门表.站点%Type,
        strSQL = strSQL & "'" & gstrNodeNo & "',"
        '全院号源归属站点_In 部门表.站点%Type,
        strSQL = strSQL & "'" & gVisitPlan_ModulePara.str号源维护站点 & "',"
        '人员id_In     人员表.Id%Type := Null,
        strSQL = strSQL & "" & IIf(HavePrivs(mstrPrivs, "所有科室"), "NULL", UserInfo.ID) & ","
        '删除安排_In Number:=0
        strSQL = strSQL & "" & IIf(blnDeletePlan, 1, 0) & ")"
        
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Next
    If cllPlan.Count > 1 Then gcnOracle.CommitTrans
    
    'XX月出诊表节点：K2_年份_月份
    'XX周出诊表节点：K3_年份_月份_周数
    Call mfrmMain.NodeChanged(strKey)
    NextNewPlanByPlan = True
    
    Screen.MousePointer = vbDefault
    Exit Function
TransErrHandler:
    If cllPlan.Count > 1 Then gcnOracle.RollbackTrans
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPlanRecords(ByVal blnMonth As Boolean, Optional ByVal lng出诊ID As Long, Optional ByVal blnPublished As Boolean, _
    Optional ByVal blnMonthAllPlan As Boolean, Optional intYear As Integer, Optional intMonth As Integer, _
    Optional ByVal dtStartDate As Date, Optional ByVal dtEndDate As Date, Optional ByVal lng跨月周出诊ID As Long) As ADODB.Recordset
    '功能：获取安排记录
    '入数：
    '   blnMonth - 是否月排班
    '   lng出诊ID   - 出诊ID
    '   blnPublished- 是否已发布
    '   blnMonthAllPlan - 是否显示本月所有安排
    '   lng跨月周出诊ID - 完整周跨月的另一个出诊表
    Dim strSQL As String, strSqlSub As String
    Dim strWhere As String, str有效号源 As String
    Dim str是否有效 As String
    Dim str排序号码 As String

    Err = 0: On Error GoTo errHandler
    str排序号码 = IIf(gVisitPlan_ModulePara.byt号码比较方式 = 0, "a.号码", "Lpad(a.号码,5,'0')")
    strSqlSub = "       " & str排序号码 & " As 排序号码, a.Id As 号源id, a.号类, a.号码, Nvl(a.是否建病案, 0) As 是否建病案,a.预约天数, a.出诊频次," & vbNewLine & _
                "       Decode(a.假日控制状态, 1, '开放预约', 2, '禁止预约', 3, '受节假日设置控制', '不上班') As 假日控制状态," & vbNewLine & _
                "       Nvl(a.是否临床排班, 0) As 是否临床排班, Decode(a.排班方式, 1, '按月排班', 2, '按周排班', '固定排班') As 排班方式," & vbNewLine & _
                "       f.名称 As 科室, f.简码 As 科室简码,Nvl(a.是否假日换休, 0) As 是否假日换休," & vbNewLine
    
    '没有"所有科室"权限的操作员只能操作自己所属科室的号源
    If HavePrivs(mstrPrivs, "所有科室") = False Then
        strWhere = "      And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = [3])"
    End If
    
    str是否有效 = "Decode(Sign(b.终止时间- Trunc(Sysdate)), -1, 0, 1)*Decode(b.审核时间, NULL, 0, 1) As 是否有效,"
    
    If blnMonth And blnMonthAllPlan Then
        '查询本月所有安排
        '先查找出诊ID
        strSQL = "Select m.Id" & vbNewLine & _
                " From 临床出诊表 M, 临床出诊表 N" & vbNewLine & _
                " Where m.排班方式 In(1, 2) And m.年份 = [1] And m.月份 = [2] And Nvl(m.站点,'-') = Nvl([4],'-')"
        
        strSQL = "Select " & str是否有效 & vbNewLine & _
                "        b.出诊id, b.Id As 安排id, e.名称 As 收费项目, b.医生姓名, g.简码 As 医生简码, g.专业技术职务 as 医生职称,h.标识符, " & vbNewLine & strSqlSub & _
                "        c.Id As 记录id, c.出诊日期, c.上班时段, c.限号数, c.限约数, c.已挂数, c.已约数, c.预约控制 As 预约控制方式," & vbNewLine & _
                "        c.是否临时出诊, c.停诊开始时间, c.停诊终止时间, c.停诊原因, c.替诊医生姓名, c.是否锁定, b.开始时间, b.终止时间" & vbNewLine & _
                " From 临床出诊号源 A, 临床出诊安排 B, 临床出诊记录 C, 收费项目目录 E, 部门表 F, 人员表 G,专业技术职务 H" & vbNewLine & _
                " Where a.Id = b.号源id And b.Id = c.安排id And a.科室id = f.Id And b.项目id = e.Id And b.医生ID = g.ID(+) and g.专业技术职务=h.名称(+)" & vbNewLine & _
                "       And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
                "       And (e.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.撤档时间 Is Null)" & vbNewLine & _
                "       And (f.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.撤档时间 Is Null)" & vbNewLine & _
                "       And (g.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.撤档时间 Is Null)" & vbNewLine & _
                "       And b.出诊id In (" & strSQL & ")" & strWhere & vbNewLine & _
                " Order By " & str排序号码 & ", 出诊日期, 上班时段"
        Set GetPlanRecords = zlDatabase.OpenSQLRecord(strSQL, "获取排班信息", intYear, intMonth, UserInfo.ID, gstrNodeNo)
        Exit Function
    End If
    
    '查询月安排或周安排
    If blnPublished Then
        strSQL = "Select " & str是否有效 & vbNewLine & _
                "        b.出诊id, b.Id As 安排id, e.名称 As 收费项目, b.医生姓名, g.简码 As 医生简码, g.专业技术职务 as 医生职称,h.标识符, " & vbNewLine & strSqlSub & _
                "        c.Id As 记录id, c.出诊日期, c.上班时段, c.限号数, c.限约数, c.已挂数, c.已约数, c.预约控制 As 预约控制方式," & vbNewLine & _
                "        c.是否临时出诊, c.停诊开始时间, c.停诊终止时间, c.停诊原因, c.替诊医生姓名, c.是否锁定, b.开始时间, b.终止时间" & vbNewLine & _
                " From 临床出诊号源 A, 临床出诊安排 B, 临床出诊记录 C, 收费项目目录 E, 部门表 F, 人员表 G,专业技术职务 H" & vbNewLine & _
                " Where a.Id = b.号源id And b.Id = c.安排id(+) And a.科室id = f.Id And b.项目id = e.Id And b.医生ID = g.Id(+) and g.专业技术职务=h.名称(+)" & vbNewLine & _
                "       And (b.出诊id = [1] Or b.出诊id = [7])" & strWhere & vbNewLine & _
                "       And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
                "       And (e.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.撤档时间 Is Null)" & vbNewLine & _
                "       And (f.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.撤档时间 Is Null)" & vbNewLine & _
                "       And (g.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.撤档时间 Is Null)" & vbNewLine & _
                " Order By " & str排序号码 & ", 出诊日期, 上班时段"
    Else
        If blnMonth Then
            str有效号源 = " And a.排班方式 = [2]"
        Else
            '当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
            '同时,当前出诊表所在时间范围内不能有月排班
            str有效号源 = " And (a.排班方式 = [2] And Not Exists (Select 1" & vbNewLine & _
                        "           From 临床出诊记录 O, 临床出诊安排 P, 临床出诊表 Q" & vbNewLine & _
                        "           Where o.安排id = p.Id And p.出诊id = q.Id And p.号源id+0 = a.Id" & vbNewLine & _
                        "               And o.出诊日期 Between [6] And Last_Day([4]) And q.排班方式 = 1)" & vbNewLine & _
                        "       Or a.排班方式 = 1 And Exists (Select 1" & vbNewLine & _
                        "           From 临床出诊记录 O, 临床出诊安排 P, 临床出诊表 Q" & vbNewLine & _
                        "           Where o.安排id = p.Id And p.出诊id = q.Id And p.号源id+0 = a.Id" & vbNewLine & _
                        "               And o.出诊日期 Between [6] And Last_Day([4]) And q.排班方式 = 2))" & vbNewLine
        End If
        '还没有制作安排的号源需要是无出诊记录 的
        str有效号源 = str有效号源 & vbNewLine & _
                    "       And Not Exists" & vbNewLine & _
                    "           (Select 1" & vbNewLine & _
                    "            From 临床出诊记录 P" & vbNewLine & _
                    "            Where p.号源id+0 = a.Id And p.出诊日期 Between [4] And [5])" & vbNewLine
        '未发布时，将所有号源缺省提取出来
        strSQL = "Select " & str是否有效 & vbNewLine & _
                "        b.出诊id, b.Id As 安排id, " & vbNewLine & strSqlSub & _
                "        Decode(b.ID,Null,e.名称,m.名称) As 收费项目, Decode(b.ID,Null,a.医生姓名,b.医生姓名) As 医生姓名," & vbNewLine & _
                "        Decode(b.ID,Null,g.简码,n.简码) As 医生简码, Decode(b.ID,Null,g.专业技术职务,n.专业技术职务) as 医生职称," & vbNewLine & _
                "        Decode(b.ID,Null,i.标识符,j.标识符) As 标识符 ," & vbNewLine & _
                "        c.Id As 记录id, c.出诊日期, c.上班时段, c.限号数, c.限约数, b.开始时间, b.终止时间, " & vbNewLine & _
                "        c.已挂数, c.已约数, c.预约控制 As 预约控制方式, c.是否临时出诊, c.停诊开始时间, c.停诊终止时间, c.停诊原因, c.替诊医生姓名, c.是否锁定" & vbNewLine & _
                " From 临床出诊号源 A, " & vbNewLine & _
                "      (Select 出诊id, ID, 号源id, 项目id, 医生ID, 医生姓名, 开始时间, 终止时间, 审核时间" & vbNewLine & _
                "        From 临床出诊安排 Where (出诊id = [1] Or 出诊id = [7])) B," & vbNewLine & _
                "      临床出诊记录 C, 收费项目目录 E, 部门表 F, 人员表 G, 收费项目目录 M, 人员表 N,专业技术职务 I,专业技术职务 J" & vbNewLine & _
                " Where a.Id = b.号源id(+) And b.Id = c.安排id(+) And a.科室id = f.Id" & vbNewLine & _
                "       And a.项目id = e.Id And a.医生ID = g.ID(+) And b.项目id = m.Id(+) And b.医生ID = n.ID(+)" & vbNewLine & _
                "       And g.专业技术职务=i.名称(+) And n.专业技术职务=j.名称(+)" & vbNewLine & _
                "       And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
                "       And (e.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.撤档时间 Is Null)" & vbNewLine & _
                "       And (f.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.撤档时间 Is Null)" & vbNewLine & _
                "       And (g.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.撤档时间 Is Null)" & vbNewLine & _
                "       And (b.Id Is Not Null Or (b.Id Is Null " & str有效号源 & "))" & strWhere & vbNewLine & _
                "       And Nvl(Nvl(f.站点,[9]),Nvl([8],'-')) = Nvl([8],'-')" & vbNewLine & _
                " Order By " & str排序号码 & ", c.出诊日期, c.上班时段"
    End If
    Set GetPlanRecords = zlDatabase.OpenSQLRecord(strSQL, "获取排班信息", lng出诊ID, IIf(blnMonth, 1, 2), _
        UserInfo.ID, dtStartDate, dtEndDate, CDate(Format(dtStartDate, "yyyy-mm-01")), lng跨月周出诊ID, _
        gstrNodeNo, gVisitPlan_ModulePara.str号源维护站点)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtFind_KeyPress(index As Integer, KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        Call ExecuteFilter
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

Private Function SaveAsTemplet(ByVal lng出诊ID As Long, Optional ByVal blnMonth As Boolean) As Boolean
    '另存为模板
    Dim obj出诊安排 As New 出诊安排, frmPlanInfoEdit As New frmClinicPlanInfoEdit
    Dim strSQL As String, strAddVisitTableSQL As String
    Dim rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lng出诊ID = 0 Then Exit Function
    '检查是否有有效的安排
    If ExistsPlanOnVisitTable(IIf(blnMonth, 1, 2), lng出诊ID, IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID)) = False Then
        MsgBox "当前出诊表中无有效的安排，不能另存为模板！", vbInformation, gstrSysName
        Exit Function
    End If
    
    obj出诊安排.排班方式 = 3 '排班方式：0-固定排班;1-按月排班;2-按周排班;3-模板
    obj出诊安排.模板类型 = IIf(mbytFun = 1, 2, 0) '0-周排班模板，1-不是按天排班的月排班模板，2-按天排班的月排班模板
    If frmPlanInfoEdit.ShowMe(Me, mlngModule, 1, obj出诊安排, False, True) = False Then Exit Function
    
    obj出诊安排.出诊ID = zlDatabase.GetNextId("临床出诊表")
    'Zl_临床出诊表_Totemplet(
    strSQL = "Zl_临床出诊表_Totemplet("
    '  出诊id_In   临床出诊表.Id%Type,
    strSQL = strSQL & "" & lng出诊ID & ","
    '  模板id_In   临床出诊表.Id%Type,
    strSQL = strSQL & "" & obj出诊安排.出诊ID & ","
    '  出诊表名_In 临床出诊表.出诊表名%Type,
    strSQL = strSQL & "'" & obj出诊安排.出诊表名 & "',"
    '  应用范围_In 临床出诊表.应用范围%Type,
    strSQL = strSQL & "" & obj出诊安排.应用范围 & ","
    '  科室id_In   临床出诊表.科室id%Type,
    strSQL = strSQL & "" & ZVal(obj出诊安排.科室ID) & ","
    '  备注_In     临床出诊表.备注%Type,
    strSQL = strSQL & "'" & obj出诊安排.备注 & "',"
    '  操作员_In   临床出诊安排.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  操作时间_In 临床出诊安排.登记时间%Type,
    strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ","
    '  站点_In     部门表.站点%Type,
    strSQL = strSQL & IIf(gstrNodeNo = "", "NULL", "'" & gstrNodeNo & "'") & ","
    '  人员id_In   人员表.Id%Type := Null
    strSQL = strSQL & "" & IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室"), "NULL", UserInfo.ID) & ")"
    
    Screen.MousePointer = vbHourglass
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    '出诊表模板节点：K0_出诊ID
    Call mfrmMain.NodeChanged("K0_" & obj出诊安排.出诊ID)
    SaveAsTemplet = True
    
    Screen.MousePointer = vbDefault
    Exit Function
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
