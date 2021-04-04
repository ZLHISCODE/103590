VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRegistPlan 
   AutoRedraw      =   -1  'True
   Caption         =   "挂号安排管理"
   ClientHeight    =   9435
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12030
   Icon            =   "frmRegistPlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   4050
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   0
      Top             =   1470
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmRegistPlan.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistPlan.frx":049E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   9075
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRegistPlan.frx":07F2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16140
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   1170
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmRegistPlan.frx":1086
      Left            =   480
      Top             =   1365
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmRegistPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mArrFilter As Variant, mlngModule As Long, mstrPrivs As String, mblnDisStop As Boolean, mblnDisDel As Boolean
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private Enum mPgIndex
    Pg_当前有效号别 = 1
    Pg_计划安排号别 = 2
End Enum
Private Const ID_PANE_SEARCH = 1
Private Const ID_PANE_Page = 2

Private mPanSearch As Pane
Private mblnUnload  As Boolean
Private mfrm安排号别 As frmRegistPlanPlan
Private WithEvents mfrm有效号别 As frmRegistPlanList
Attribute mfrm有效号别.VB_VarHelpID = -1
Private WithEvents mfrmFilter As frmRegistPlanFilter
Attribute mfrmFilter.VB_VarHelpID = -1
Private mfrmUnitReg  As frmCooperateUnitsReg
Private mfrmUnitRegPlan As frmCooperateUnitsRegPlan

Private mblnFirst As Boolean

Private mbln自动默认限约数 As Boolean '45519
Private mbln预约单存在禁止删除 As Boolean
Private Sub zlRptPrint(bytMode As Byte)
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    If Val(tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
        mfrm有效号别.zlRptPrint bytMode
    Else
        mfrm安排号别.zlRptPrint bytMode
    End If
End Sub
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区哉
    '编制:刘兴洪
    '日期:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, strKey As String
    If mfrmFilter Is Nothing Then Set mfrmFilter = New frmRegistPlanFilter
    Load mfrmFilter
    '初始化 参数 是否显示停用安排
    mfrmFilter.ShowStop = mblnDisStop
    mfrmFilter.ShowDel = mblnDisDel
    Set mArrFilter = mfrmFilter.GetCondition
    
    With dkpMan
        .ImageList = imlPaneIcons
        Set mPanSearch = .CreatePane(ID_PANE_SEARCH, 300, 100, DockLeftOf, Nothing)
        mPanSearch.Title = "条件设置": mPanSearch.Options = PaneNoCloseable
         Set objPane = .CreatePane(ID_PANE_Page, 400, 400, DockRightOf, mPanSearch)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.Hwnd
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    dkpMan.RecalcLayout: DoEvents
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
    Call GetRegInFor(g私有模块, Me.Name, "隐藏", strKey)
    If Val(strKey) = 1 Then mPanSearch.Hide
    mPanSearch.MinTrackSize.Width = 230: mPanSearch.MaxTrackSize.Width = 230
       
End Function
Private Sub zlRefreshData()
    zlCommFun.ShowFlash "正在装载数据,请稍后..."
    Call InitData
    Set mArrFilter = mfrmFilter.GetCondition
    Call mfrm有效号别.zlRefreshData(mArrFilter)
    Call mfrm安排号别.zlRefreshData(mArrFilter)
    If Val(tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
        Call mfrm有效号别.zlActtion
    Else
        Call mfrm安排号别.zlActtion
    End If
    zlCommFun.StopFlash
End Sub
Private Sub zlPlanManager(ByVal bytFun As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加/取消/审核/取消审核：计划安排
    '入参:bytFun-(0-增加,1-取消,2-审核,3-取消审核,4-查阅
    '编制:刘兴洪
    '日期:2009-09-15 17:16:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lng计划ID As Long
    frmRegistPlanArrange.自动默认限约数 = mbln自动默认限约数
    Select Case bytFun
    Case 0  '增加计划安排
        If Val(tbPage.Selected.Tag) = mPgIndex.Pg_计划安排号别 Then Exit Sub
        lngID = mfrm有效号别.zlGet安排ID
        If lngID = 0 Then Exit Sub
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, ed_计划安排, lngID, "") = False Then Exit Sub
    Case 5  '修改计划
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_计划安排号别 Then
            lngID = mfrm有效号别.zlGet安排ID
            lng计划ID = mfrm有效号别.zlGet安排ID(True)
        Else
            lngID = mfrm安排号别.zlGet安排ID(False)
            lng计划ID = mfrm安排号别.zlGet安排ID(True)
        End If
        If lng计划ID = 0 Then Exit Sub
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, Ed_安排修改, lngID, lng计划ID) = False Then Exit Sub
         Call mfrm有效号别.ReloadTimePlan(True)
    Case 1  '取消计划安排
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_计划安排号别 Then
            lngID = mfrm有效号别.zlGet安排ID
            lng计划ID = mfrm有效号别.zlGet安排ID(True)
        Else
            lngID = mfrm安排号别.zlGet安排ID(False)
            lng计划ID = mfrm安排号别.zlGet安排ID(True)
        End If
        If lng计划ID = 0 Then Exit Sub
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, Ed_安排删除, lngID, lng计划ID) = False Then Exit Sub
    Case 2   '审核
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_计划安排号别 Then
            lngID = mfrm有效号别.zlGet安排ID
            lng计划ID = mfrm有效号别.zlGet安排ID(True)
        Else
            lngID = mfrm安排号别.zlGet安排ID(False)
            lng计划ID = mfrm安排号别.zlGet安排ID(True)
        End If
        If lng计划ID = 0 Then Exit Sub
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, Ed_安排审核, lngID, lng计划ID) = False Then Exit Sub
    Case 3   '取消
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_计划安排号别 Then
            lngID = mfrm有效号别.zlGet安排ID
            lng计划ID = mfrm有效号别.zlGet安排ID(True)
        Else
            lngID = mfrm安排号别.zlGet安排ID(False)
            lng计划ID = mfrm安排号别.zlGet安排ID(True)
        End If
        If lng计划ID = 0 Then Exit Sub
        If CheckPlanBooking(lng计划ID) Then
            If MsgBox("当前计划已经存在预约挂号单，你确定要取消审核吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, Ed_安排取消, lngID, lng计划ID) = False Then Exit Sub
    Case 4  '查阅
    
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_计划安排号别 Then
            lngID = mfrm有效号别.zlGet安排ID
            lng计划ID = mfrm有效号别.zlGet安排ID(True)
        Else
            lngID = mfrm安排号别.zlGet安排ID(False)
            lng计划ID = mfrm安排号别.zlGet安排ID(True)
        End If
        If lng计划ID = 0 Then Exit Sub
        If frmRegistPlanArrange.ShowCard(Me, mlngModule, mstrPrivs, ed_安排查阅, lngID, lng计划ID) = False Then Exit Sub
        Exit Sub
    Case Else
    End Select
    
    If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_计划安排号别 Then
        zlCommFun.ShowFlash "正在装载数据,请稍后..."
        mfrm有效号别.zlRefreshOlnyPlanData
        mfrm有效号别.Tag = "1"
        mfrm有效号别.zlActtion
        mfrm安排号别.Tag = ""
        zlCommFun.StopFlash
    Else
        zlCommFun.ShowFlash "正在装载数据,请稍后..."
        mfrm有效号别.Tag = ""
        mfrm安排号别.Tag = "1"
        mfrm安排号别.zlRefreshData mArrFilter
        mfrm安排号别.zlActtion
        zlCommFun.StopFlash
    End If
End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_Edit_NewItem   '增加
        Call zlAddItem
    Case conMenu_Edit_Modify    '修改
        Call zlModifyItem
    Case conMenu_Edit_Delete '删除操作
        Call zlDeleteItem
    Case conMenu_View_Refresh   '刷新
        Call zlRefreshData
    Case conMenu_View_ShowStoped '显示停用安排
        mblnDisStop = IIf(mblnDisStop, False, True)
        Call zlDatabase.SetPara("显示停用安排", IIf(mblnDisStop, "1", "0"), glngSys, mlngModule)
        mfrmFilter.ShowStop = mblnDisStop
        Call zlRefreshData
    Case conMenu_View_ShowDel '显示删除安排
        mblnDisDel = IIf(mblnDisDel, False, True)
        Call zlDatabase.SetPara("显示删除安排", IIf(mblnDisDel, "1", "0"), glngSys, mlngModule)
        mfrmFilter.ShowDel = mblnDisDel
        Call zlRefreshData
    Case conMenu_Edit_AllStartNO     '全部启用序号控制
        Call BatchSet(1)
    Case conMenu_Edit_AllStopNO      '全部取消序号控制
        Call BatchSet(0)
    Case conMenu_Edit_Reuse       ' "启用安排(&I)")
        Call zlStopAndResume(False)
    Case conMenu_Edit_Stop          ' "停用安排(&T)"):
        Call zlStopAndResume(True)
    Case conMenu_Edit_StopPlanTimes '设置停用计划
        Call zlStopPlanTimes
    Case conMenu_Edit_ClearStopPlan '清除所有停用计划
        Call zlClearStopPlanTimes
    Case conMenu_Manage_Bespeak '时间段设置
        '*****************
        '上班时间段设置
        '*****************
          frmSplitTime.Show 1, Me
    Case ComMenu_Edit_AutoDefaultLimitAppointment '默认限约数
         Control.Checked = Not Control.Checked
         mbln自动默认限约数 = Control.Checked
         Call zlDatabase.SetPara("自动默认限约数", IIf(mbln自动默认限约数, 1, 0), glngSys, mlngModule)
    Case comMenu_Edit_SetDateSegment
'        '*****************
'        '挂号安排时间段设置
'        '*****************
'        '问题号:51429
'        If Control.Caption = "安排时段设置" Then
'            Call zlSetDateSegment
'        Else
'            Call zlSetPlanDateSegment
'        End If
    Case conMenu_Edit_SetPlanDateSeqment
          '*********************
          '挂号安排计划时间段设置
          '*********************
          Call zlSetPlanDateSegment
    Case comMenu_Edit_UnitRegModify  '合作单位安排控制
        zlExecuteUnitReg
    Case ComMenu_Edit_UnitRegArrangeModify '合作单位计划控制
        Call zlExecuteUnitReg(True)
        
    Case conMenu_Edit_PlanAdd   '计划安排
        Call zlPlanManager(0)  '0-增加,1-取消,2-审核,3-取消审核,4-查阅,5-修改
    Case conMenu_Edit_PlanModify   '修改计划安排
        Call zlPlanManager(5)  '0-增加,1-取消,2-审核,3-取消审核,4-查阅,5-修改
    Case conMenu_Edit_PlanDelete   '取消安排
        Call zlPlanManager(1)  '0-增加,1-取消,2-审核,3-取消审核,4-查阅,5-修改
    Case conMenu_Edit_PlanVerify      '审核
        Call zlPlanManager(2)  '0-增加,1-取消,2-审核,3-取消审核,4-查阅,5-修改
    Case conMenu_Edit_PlanCancel   '取消审核
        Call zlPlanManager(3)  '0-增加,1-取消,2-审核,3-取消审核,4-查阅,5-修改
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
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.Hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.Hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            If Val(Me.tbPage.Selected.Tag) = Pg_当前有效号别 Then
                Call mfrm有效号别.zlCallCustomReprot(Me, Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
            Else
                Call mfrm安排号别.zlCallCustomReprot(Me, Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
            End If
        End If
    End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
            Control.Enabled = mfrm有效号别.zlGet安排ID <> 0
        Else
            Control.Enabled = mfrm安排号别.zlGet安排ID <> 0
        End If
    Case conMenu_Edit_NewItem '增加
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "安排")
        Else
            Control.Visible = False
        End If
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AllStartNO     '全部启用序号控制
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "安排")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AllStopNO      '全部取消序号控制
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "安排")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_StopPlanTimes
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "安排")
            Control.Enabled = Control.Visible And Not mfrm有效号别.zlIsStopPlan And mfrm有效号别.zlGet安排ID <> 0
        Else
            Control.Visible = False
        End If
    Case conMenu_Edit_Stop  '停用安排
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 And zlStr.IsHavePrivs(mstrPrivs, "停用安排") Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "安排")
            Control.Enabled = Control.Visible And Not mfrm有效号别.zlIsStopPlan And mfrm有效号别.zlGet安排ID <> 0
        Else
            Control.Visible = False
        End If
    Case conMenu_Edit_Reuse '启用安排
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 And zlStr.IsHavePrivs(mstrPrivs, "启用安排") Then
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "安排")
                Control.Enabled = Control.Visible And mfrm有效号别.zlIsStopPlan And mfrm有效号别.zlGet安排ID <> 0
        Else
            Control.Visible = False
        End If
    Case conMenu_Edit_PlanAdd   '计划安排
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "增加计划")
            Control.Enabled = mfrm有效号别.zlGet安排ID <> 0
        Else
            Control.Visible = False
        End If
        Control.Enabled = Control.Visible And Control.Enabled
        If Control.Enabled Then
            Control.Enabled = Not mfrm有效号别.zlGet安排停用
        End If
    Case conMenu_Edit_PlanModify    '修改计划
        If zlStr.IsHavePrivs(mstrPrivs, "修改诊室") = True Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "修改计划")
            If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
                'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
                lngID = mfrm有效号别.zlPlanStatus
                Control.Enabled = lngID <> 0
            Else
                'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
                lngID = mfrm安排号别.zlPlanStatus
                Control.Enabled = lngID <> 0
            End If
            Control.Enabled = Control.Visible And Control.Enabled
        Else
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "修改计划")
            If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
                'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
                lngID = mfrm有效号别.zlPlanStatus
                Control.Enabled = lngID = 1
            Else
                'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
                lngID = mfrm安排号别.zlPlanStatus
                Control.Enabled = lngID = 1
            End If
            Control.Enabled = Control.Visible And Control.Enabled
        End If
    Case conMenu_Edit_PlanDelete, conMenu_Edit_SetPlanDateSeqment, ComMenu_Edit_UnitRegArrangeModify  '删除计划
        If Control.ID = conMenu_Edit_PlanModify Or Control.ID = conMenu_Edit_SetPlanDateSeqment Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "修改计划")
        Else
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "删除计划")
        End If
        If Control.ID = ComMenu_Edit_UnitRegArrangeModify Then
            '处理规则
             Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "挂号合作单位控制")
        End If
        
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
            'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
            lngID = mfrm有效号别.zlPlanStatus
            Control.Enabled = lngID = 1
        Else
            'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
            lngID = mfrm安排号别.zlPlanStatus
            Control.Enabled = lngID = 1
        End If
        Control.Enabled = Control.Visible And Control.Enabled
        If Control.ID = conMenu_Edit_SetPlanDateSeqment Then
            Control.Enabled = mfrm有效号别.zlHaveDatePlan(True) And Control.Enabled
        End If
    Case conMenu_Edit_ClearStopPlan '清除所有停用计划
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "清除所有停用计划")
        
    Case conMenu_Edit_PlanVerify      '审核
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "计划审核")
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
            'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
            lngID = mfrm有效号别.zlPlanStatus
            Control.Enabled = lngID = 1
        Else
            'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
            lngID = mfrm安排号别.zlPlanStatus
            Control.Enabled = lngID = 1
        End If
        Control.Enabled = Control.Visible And Control.Enabled
    Case conMenu_Edit_PlanCancel   '取消审核
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "取消审核")
         If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
            'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
            lngID = mfrm有效号别.zlPlanStatus
            Control.Enabled = (lngID <> 3 And lngID = 2)
        Else
            'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
            lngID = mfrm安排号别.zlPlanStatus
            Control.Enabled = (lngID <> 3 And lngID = 2)
        End If
        Control.Enabled = Control.Visible And Control.Enabled
    Case conMenu_Edit_Modify
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "安排")
            'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
            lngID = mfrm有效号别.zlPlanStatus
            Control.Enabled = mfrm有效号别.zlGet安排ID <> 0 And Not mfrm有效号别.zlIsStopPlan
        Else
            Control.Visible = False
        End If
        Control.Enabled = Control.Enabled And Control.Visible
    Case conMenu_Edit_Delete, comMenu_Edit_UnitRegModify  ', comMenu_Edit_SetDateSegment
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "安排")
            'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
            lngID = mfrm有效号别.zlPlanStatus
            Control.Enabled = mfrm有效号别.zlGet安排ID <> 0 And lngID = 0 And Not mfrm有效号别.zlIsStopPlan
        Else
            Control.Visible = False
        End If
        If Control.ID = comMenu_Edit_UnitRegModify Then
             Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "挂号合作单位控制")
        End If
        Control.Enabled = Control.Enabled And Control.Visible
'        If Control.ID = comMenu_Edit_SetDateSegment Then
'             '问题号:51427
'            If mfrm有效号别.是否选中计划列表 = False Then
'                Control.Enabled = Not mfrm有效号别.Have计划
'                Control.Caption = "安排时段设置"
'            Else
'                If Val(Me.tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
'                'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
'                        lngID = mfrm有效号别.zlPlanStatus
'                        Control.Enabled = lngID = 1
'                Else
'                'zlPlanStatus() '0-不存在计划安排,1-未审核,2-已经审核,3-已经生效\
'                        lngID = mfrm安排号别.zlPlanStatus
'                        Control.Enabled = lngID = 1
'                End If
'                Control.Enabled = Control.Visible And Control.Enabled
'                Control.Enabled = mfrm有效号别.zlHaveDatePlan(True) And Control.Enabled
'                Control.Caption = "计划时段设置"
'            End If
'        End If
    Case ComMenu_Edit_AutoDefaultLimitAppointment  '自动默认限约数
         Control.Checked = mbln自动默认限约数
      '显示停用安排
    Case conMenu_View_ShowStoped
         Control.Checked = mblnDisStop
         '显示删除安排
    Case conMenu_View_ShowDel
         Control.Checked = mblnDisDel
    Case conMenu_Manage_Bespeak '时间段设置
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "时间段设置")
        Control.Enabled = Control.Visible
    Case conMenu_View_Refresh
                
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
'    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case ID_PANE_SEARCH
        Item.Handle = mfrmFilter.Hwnd
    Case ID_PANE_Page
        Item.Handle = picList.Hwnd
    End Select
End Sub
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:
    
 
    Set mfrm有效号别 = New frmRegistPlanList
    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_当前有效号别, "当前有效号别", mfrm有效号别.Hwnd, 0)
    ObjItem.Tag = mPgIndex.Pg_当前有效号别
    Call mfrm有效号别.UpdatePara(mfrmFilter.chkShowExpiredPlan.Value = 1)
    
    Set mfrm安排号别 = New frmRegistPlanPlan
    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_计划安排号别, "计划安排号别", mfrm安排号别.Hwnd, 0)
    ObjItem.Tag = mPgIndex.Pg_计划安排号别

     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
    stbThis.Top = Me.ScaleHeight - Me.stbThis.Height
End Sub
'Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
'    Top = IIf(cbr.Visible, cbr.Height, 0)
'    Bottom = IIf(stbThis.Visible, stbThis.Height, 0)
'End Sub
'Private Sub Form_Resize()
'    Dim cbrH As Long '工具条占用高度
'    Dim staH As Long '状态栏占用高度
'    Dim i As Integer, lngW As Long
'
'    On Error Resume Next
'    If WindowState = 1 Then Exit Sub
'    '靠齐控件宽度和高度
'    cbrH = IIf(cbr.Visible, cbr.Height, 0)
'    staH = IIf(stbThis.Visible, stbThis.Height, 0)
'    With mshPlan
'        .Left = Me.ScaleLeft
'        .Top = Me.ScaleTop + cbrH
'        .Width = Me.ScaleWidth
'        .Height = Me.ScaleHeight - cbrH - staH
'    End With
'End Sub


Private Sub Form_Activate()
    Dim strKey As String
    If mblnUnload Then Unload Me: Exit Sub
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
'    Form_Resize
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    mblnFirst = True
    '获取参数 是否显示停用安排
    mblnDisStop = Val(zlDatabase.GetPara("显示停用安排", glngSys, mlngModule, 0)) = 1
    '获取参数 是否显示删除安排
    mblnDisDel = Val(zlDatabase.GetPara("显示删除安排", glngSys, mlngModule, 0)) = 1
    mbln自动默认限约数 = Val(zlDatabase.GetPara("自动默认限约数", glngSys, mlngModule, 0)) = 1
    '46639预约单存在
    mbln预约单存在禁止删除 = Val(zlDatabase.GetPara("预约单存在禁止删除", glngSys, mlngModule)) = 1
    Call InitData
    Call zlCommFun.SetWindowsInTaskBar(Me.Hwnd, False)
    Call zlDefCommandBars
    Call InitPanel
    Call InitPage
    '获取数据
    Call zlRefreshData
    '权限处理
    'Call 权限控制
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
End Sub
Private Sub InitData()
    Dim strSQL As String
    '重新计划:按计划安排
    strSQL = "Zl_挂号安排_Autoupdate"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    Call SaveRegInFor(g私有模块, Me.Name, "隐藏", IIf(mPanSearch.Hidden, 1, 0))
    If Not mfrmFilter Is Nothing Then Unload mfrmFilter
    If Not mfrm安排号别 Is Nothing Then Unload mfrm安排号别
    If Not mfrm有效号别 Is Nothing Then Unload mfrm有效号别
    
    Set mfrmFilter = Nothing
    Set mfrm安排号别 = Nothing
    Set mfrm有效号别 = Nothing
    SaveWinState Me, App.ProductName
End Sub
Private Sub mfrmFilter_zlRefreshCon(ByVal ArrFilter As Variant)
    Set mArrFilter = ArrFilter
    Call mfrm有效号别.UpdatePara(mfrmFilter.chkShowExpiredPlan.Value = 1)
    Call mfrm有效号别.ReloagUnitRegPlan
    '条件发生了改变
    Select Case Val(tbPage.Selected.Tag)
    Case mPgIndex.Pg_当前有效号别
        Call mfrm有效号别.zlRefreshData(ArrFilter)
    Case mPgIndex.Pg_计划安排号别
        Call mfrm安排号别.zlRefreshData(ArrFilter)
    Case Else
    End Select
End Sub
Private Sub mfrm有效号别_zlPopuMenu(intType As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = 2 And intType = 0) Then Exit Sub
    Dim objPopup As CommandBarPopup
    Set objPopup = cbsThis.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
End Sub

Private Sub zlSetDateSegment()
    '************************
    '挂号安排时段设置
    '************************
    Dim lng安排ID       As Long
    lng安排ID = mfrm有效号别.zlGet安排ID(False)
    If ExistsBooking(lng安排ID) Then
         If MsgBox("该号别存在预约挂号单,是否修改时段?", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
            Exit Sub
         End If
    End If
    If frmRegistPlanDatSet.ShowMe(lng安排ID, Edit) Then
         mfrm有效号别.ReloadTimePlan
    End If
End Sub

Private Sub zlSetPlanDateSegment()
     '************************
    '挂号安排计划时段设置
    '************************
    Dim lng计划ID       As Long
    If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_计划安排号别 Then
        lng计划ID = mfrm有效号别.zlGet安排ID(True)
    Else
        lng计划ID = mfrm安排号别.zlGet安排ID(True)
    End If
   ' lng计划ID = mfrm有效号别.zlGet安排ID(True)
    If frmRegistPlanPlanDatSet.ShowMe(lng计划ID, Edit) Then
         mfrm有效号别.ReloadTimePlan (True)
    End If
    
End Sub
Private Sub zlExecuteUnitReg(Optional ByVal blnPlan As Boolean = False)
     '************************
    '
    '************************
    Dim lng安排ID       As Long
    Dim lng计划ID       As Long
    If blnPlan = False Then
        lng安排ID = mfrm有效号别.zlGet安排ID(False)
        If ExistsBooking(lng安排ID) Then
            Call MsgBox("该号别存在预约挂号单,不能修改合作单位预约号分配！", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
            Exit Sub
        End If
        If Not mfrmUnitReg Is Nothing Then Set mfrmUnitReg = Nothing
        Set mfrmUnitReg = New frmCooperateUnitsReg
        If mfrmUnitReg.zlShowMe(lng安排ID, mlngModule, mstrPrivs) Then              '刷新
           mfrm有效号别.zl_ReLoadUnitReg
        End If
        Set mfrmUnitReg = Nothing
    Else
         '************************
        '
        '************************
        
        If Val(tbPage.Selected.Tag) <> mPgIndex.Pg_计划安排号别 Then
            lng计划ID = mfrm有效号别.zlGet安排ID(True)
        Else
            lng计划ID = mfrm安排号别.zlGet安排ID(True)
        End If
        If Not mfrmUnitRegPlan Is Nothing Then Set mfrmUnitRegPlan = Nothing
            Set mfrmUnitRegPlan = New frmCooperateUnitsRegPlan
            If mfrmUnitRegPlan.zlShowMe(lng计划ID, mlngModule, mstrPrivs) Then
                mfrm有效号别.ReloagUnitRegPlan
                  
            End If
            Set mfrmUnitRegPlan = Nothing
    End If
End Sub

Private Sub zlAddItem()
    frmRegistPlanEdit.自动默认限约数 = mbln自动默认限约数
    If frmRegistPlanEdit.ShowEdit(Me, edt_新增, mlngModule, mstrPrivs, 0, mfrmFilter.zlGet科室ID) = False Then Exit Sub
    mfrm有效号别.zlRefreshData (mArrFilter)
End Sub
Private Sub BatchSet(bytFun As Byte)
    Dim strSQL As String
    Dim i As Long
        
    If MsgBox("你确定要对所有限号或限约的号别" & IIf(bytFun = 1, "启用", "取消") & "序号控制吗?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
        Exit Sub
    End If
    
    On Error GoTo errH
    strSQL = "Zl_挂号安排_序号控制(" & bytFun & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mfrm有效号别.zlRefreshData mArrFilter
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function CheckExistsBooking(str号别 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定号别是否存在预约挂号单
    '入参:str号别-号别
    '返回:存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-15 10:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select /*+ Rule*/ Min(发生时间) 时间" & vbNewLine & _
            "From 门诊费用记录" & vbNewLine & _
            "Where 记录性质 = 4 And 记录状态 In (0, 1) And 计算单位 = [1] And 发生时间 > 登记时间"
    If gint预约天数 = 0 Then
        strSQL = strSQL & " And 发生时间 > Sysdate"
    Else
        strSQL = strSQL & " And 发生时间 Between Sysdate And Sysdate+" & gint预约天数
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str号别)
    
    CheckExistsBooking = Not IsNull(rsTmp!时间)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckPlanBooking(lng计划ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定计划是否存在预约挂号单
    '入参:str号别-号别
    '返回:存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-15 10:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1" & vbNewLine & _
            " From 病人挂号记录 A, 挂号安排计划 B" & vbNewLine & _
            " Where b.号码 = a.号别 And a.记录状态 = 1 And a.发生时间 Between b.生效时间 + 0 And b.失效时间 And b.审核时间 Is Not Null And b.Id = [1]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng计划ID)
    
    CheckPlanBooking = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub zlDeleteItem()
    '删除号别
    mfrmFilter.ShowDel = mblnDisDel
    Set mArrFilter = mfrmFilter.GetCondition
    If mfrm有效号别.zlExecuteDeleteList(mbln预约单存在禁止删除) = False Then Exit Sub
    mfrm有效号别.Tag = ""
End Sub
Private Sub zlModifyItem()
    frmRegistPlanEdit.自动默认限约数 = mbln自动默认限约数 '45519
    If mfrm有效号别.zlExecuteModifyList(Me) = False Then Exit Sub
    mfrm有效号别.Tag = ""
    mfrm有效号别.ReloadTimePlan
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
    If Not mfrmFilter Is Nothing Then
        mfrmFilter.zlblnShowPlanCon = Val(Item.Tag) = mPgIndex.Pg_计划安排号别
    End If
    zlCommFun.ShowFlash "正在装载数据,请稍后..."
   If Val(tbPage.Selected.Tag) = mPgIndex.Pg_当前有效号别 Then
        If mfrm有效号别.Tag = "" Then
            mfrm有效号别.zlRefreshOlnyPlanData
            mfrm有效号别.Tag = "1"
        End If
        Call mfrm有效号别.zlActtion
    Else
        If mfrm安排号别.Tag = "" Then
            mfrm安排号别.zlRefreshData mArrFilter
            mfrm安排号别.Tag = "1"
        End If
        Call mfrm安排号别.zlActtion
    End If
    zlCommFun.StopFlash
End Sub

Public Function zlDefCommandBars() As Boolean
    '----------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/1/9
    '----------------------------------------------------------------------------------------
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
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        'Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加安排(&A)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改安排(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除安排(&D)")
'        '问题号:51156
'        Set mcbrControl = .Add(xtpControlButton, comMenu_Edit_SetDateSegment, "安排时段设置"): mcbrControl.IconId = 3063
        Set mcbrControl = .Add(xtpControlButton, comMenu_Edit_UnitRegModify, "合作单位安排控制"): mcbrControl.IconId = 3813
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_AllStartNO, "全部启用序号控制(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_AllStopNO, "全部取消序号控制(&U)"):
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用安排(&I)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停用安排(&T)"):
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StopPlanTimes, "设置停用计划(&Q)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClearStopPlan, "清除所有停用计划(&W)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanAdd, "增加计划(&N)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanModify, "修改计划(&G)")
        'Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SetPlanDateSeqment, "设置时段"): mcbrControl.IconId = 3063:
        Set mcbrControl = .Add(xtpControlButton, ComMenu_Edit_UnitRegArrangeModify, "合作单位计划控制"): mcbrControl.IconId = 3813
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanDelete, "删除计划(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanVerify, "审核计划(&V)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanCancel, "取消审核(&C)")
        '问题号:51156
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "上班时段设置(&T)"): mcbrControl.IconId = 3038: mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, ComMenu_Edit_AutoDefaultLimitAppointment, "自动默认限约数")
        mcbrControl.Checked = mbln自动默认限约数
        
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
         '显示停用安排
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "显示停用(&D)")
        mcbrControl.Checked = mblnDisStop
         '问题: 45525
         '显示删除安排
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ShowDel, "显示删除(&Z)")
        mcbrControl.Checked = mblnDisDel
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("T"), comMenu_Edit_SetDateSegment
        .Add FCONTROL, Asc("V"), conMenu_Edit_PlanVerify
         
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加安排"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改安排")
        '问题号:51156
        'Set mcbrControl = .Add(xtpControlButton, comMenu_Edit_SetDateSegment, "安排时段设置"): mcbrControl.IconId = 3063:
       ' Set mcbrControl = .Add(xtpControlButton, comMenu_Edit_UnitRegModify, "合作单位安排控制"): mcbrControl.IconId = 3813
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除安排")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用安排"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停用安排"):
   
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanAdd, "增加计划"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanModify, "修改计划"): mcbrControl.BeginGroup = True
        '问题号:51156
        'Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SetPlanDateSeqment, "设置时段"): mcbrControl.IconId = 3063:
       ' Set mcbrControl = .Add(xtpControlButton, ComMenu_Edit_UnitRegArrangeModify, "合作单位计划控制"): mcbrControl.IconId = 3813
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanDelete, "删除计划")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanVerify, "计划审核")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_PlanCancel, "取消审核")
        '问题号:51156
        'Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "时间段"): mcbrControl.IconId = 3063: mcbrControl.BeginGroup = True
        

        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub zlStopAndResume(Optional blnStop As Boolean = True)
    '删除号别
    mfrmFilter.ShowStop = mblnDisStop
    Set mArrFilter = mfrmFilter.GetCondition
    If mfrm有效号别.zlStopAndResume(blnStop) = False Then Exit Sub
    mfrm有效号别.Tag = ""
End Sub
Private Sub zlStopPlanTimes()
    If mfrm有效号别.zlStopPlanTimes() = False Then Exit Sub
    mfrm有效号别.Tag = ""
End Sub
Private Sub zlClearStopPlanTimes()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除所有停用计划
    '编制:刘兴洪
    '日期:2010-09-09 14:42:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrm有效号别.zlClearStopPlanTimes() = False Then Exit Sub
    mfrm有效号别.Tag = ""
End Sub
 
 
Private Function ExistsBooking(ByVal lng安排ID As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定号别是否存在预约挂号单
    '入参:str号别-号别
    '返回:存在,返回true,否则返回False
    '编制:
    '日期:2012-04-26 10:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select min(A.发生时间) as 时间  From 病人挂号记录 A, 挂号安排 B "
    strSQL = strSQL & vbCrLf & " Where A.号别 = B.号码 "
    strSQL = strSQL & vbCrLf & "       And 记录状态 = 1 and b.id=[1] And 发生时间 > 登记时间 "
    If gint预约天数 = 0 Then
        strSQL = strSQL & " And A.发生时间 > Sysdate"
    Else
        strSQL = strSQL & " And A.发生时间 Between Sysdate And Sysdate+" & gint预约天数
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
    ExistsBooking = Not IsNull(rsTmp!时间)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

