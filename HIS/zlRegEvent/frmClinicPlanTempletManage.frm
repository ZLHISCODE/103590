VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicPlanTempletManage 
   BorderStyle     =   0  'None
   Caption         =   "出诊模板管理"
   ClientHeight    =   7260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   9360
      MaxLength       =   100
      TabIndex        =   4
      Top             =   930
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfRegistPlan 
      Height          =   2085
      Left            =   390
      TabIndex        =   1
      Top             =   1050
      Width           =   2535
      _cx             =   4471
      _cy             =   3678
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
      FormatString    =   $"frmClinicPlanTempletManage.frx":0000
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
         Picture         =   "frmClinicPlanTempletManage.frx":0075
         ScaleHeight     =   135
         ScaleWidth      =   150
         TabIndex        =   2
         Top             =   90
         Width           =   150
      End
   End
   Begin VB.Label lblPlanInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用范围：所属科室(门诊内科)  备注：发鬼地方个梵蒂冈发鬼地方个法规定功夫功夫"
      Height          =   180
      Left            =   3960
      TabIndex        =   3
      Top             =   150
      Width           =   6840
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   10845
      _Version        =   589884
      _ExtentX        =   19129
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "出诊安排>出诊模板"
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
      Height          =   6915
      Left            =   0
      Top             =   0
      Width           =   11595
   End
End
Attribute VB_Name = "frmClinicPlanTempletManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mlngModule As Long
Private mstrPrivs As String

Private mbytFun As Byte '1-月模板,2-周模板
Private mlng出诊ID As Long
Private mrsPlanRecords As ADODB.Recordset
Private mintFindType As Integer

Private mstrOldSelRangePlan As String '记录选择网格区域，格式"开始行|结束行|开始列|结束列"，用于恢复背景色

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
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)", cbrControl.Index + 1): cbrControl.BeginGroup = True
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddTemplet, "增加模板(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改模板(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除模板(&D)")

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyPlanItem, "调整安排(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyUnitRegist, "调整预约挂号控制(&U)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStartNO, "全部启用序号控制(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AllStopNO, "全部取消序号控制(&T)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAllPlan, "清除当前号源安排(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAll, "清除所有号源安排(&A)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextMonthNewPlan, "生成月出诊表(&N)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextWeekNewPlan, "生成周出诊表(&W)"): cbrControl.BeginGroup = True
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
'        Set cbrControl = .Add(xtpControlButton,conMenu_View_Notify,"刷新提醒(&B)",cbrControl.Index)
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AddTemplet, "增加模板", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改模板", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除模板", cbrControl.index + 1)
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyPlanItem, "调整安排", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyUnitRegist, "预约挂号控制", cbrControl.index + 1)
        cbrControl.ToolTipText = "调整预约挂号控制"

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextMonthNewPlan, "生成月出诊表", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextWeekNewPlan, "生成周出诊表", cbrControl.index + 1): cbrControl.BeginGroup = True
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
        .Add FCONTROL, Asc("T"), conMenu_Edit_AddTemplet
        .Add FCONTROL, Asc("E"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
        .Add FCONTROL, Asc("M"), conMenu_Edit_ModifyPlanItem
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

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnPlanDataCol As Boolean '当前选择是否为安排数据列
    Dim blnHavePlan As Boolean '当前选择是否有安排
    Dim bln禁止预约 As Boolean
    Dim blnEnabled As Boolean
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = vsfRegistPlan.Rows > vsfRegistPlan.FixedRows
    Case conMenu_EditPopup
        Control.Visible = ((mfrmMain.mFunListActived And (HavePrivs(mstrPrivs, "模板管理;出诊安排")) _
                        Or (mfrmMain.mFunListActived = False And HavePrivs(mstrPrivs, "模板管理"))))
        Control.Enabled = Control.Visible
    Case conMenu_Edit_AddTemplet '增加模板
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify '修改模板
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible And mlng出诊ID <> 0
    Case conMenu_Edit_Delete '删除模板
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived
        Control.Enabled = Control.Visible And mlng出诊ID <> 0

    Case conMenu_Edit_ModifyPlanItem '调整安排
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        blnPlanDataCol = vsfRegistPlan.Col >= gPlanGrid_FixedCols
        blnEnabled = mlng出诊ID > 0
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_是否临床排班)) = "" Then blnEnabled = False
        End If
        Control.Enabled = Control.Visible And blnEnabled And blnPlanDataCol
    Case conMenu_Edit_ModifyUnitRegist '调整预约挂号控制
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        blnPlanDataCol = vsfRegistPlan.Col >= gPlanGrid_FixedCols
        blnHavePlan = Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) <> ""
        bln禁止预约 = Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col) + 2)) = "-"
        blnEnabled = mlng出诊ID > 0
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_是否临床排班)) = "" Then blnEnabled = False
        End If
        Control.Enabled = Control.Visible And blnEnabled And blnPlanDataCol And blnHavePlan And Not bln禁止预约
    Case conMenu_Edit_AllStartNO '全部启用序号控制
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        Control.Enabled = Control.Visible And mlng出诊ID <> 0
    Case conMenu_Edit_AllStopNO '全部取消序号控制
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        Control.Enabled = Control.Visible And mlng出诊ID <> 0
    Case conMenu_Edit_ClearAllPlan '清除当前号源所有安排
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        blnEnabled = mlng出诊ID > 0
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_是否临床排班)) = "" Then blnEnabled = False
        End If
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_ClearAll '清除所有号源安排
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        Control.Enabled = Control.Visible And mlng出诊ID <> 0
    Case conMenu_Edit_NextMonthNewPlan '生成下月安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mbytFun = 1
        Control.Enabled = Control.Visible And mlng出诊ID <> 0
    Case conMenu_Edit_NextWeekNewPlan '生成下周安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排") And mbytFun = 2
        Control.Enabled = Control.Visible And mlng出诊ID <> 0
    Case conMenu_View_FindType '查找方式
        Control.Caption = "按" & Decode(mintFindType, 0, "号码", 1, "科室", 2, "医生", "号码") & "过滤↓"
    Case conMenu_View_Find
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '查找方式
        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
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
    Dim frmEdit As frmClinicPlanEdit
    Dim lng记录ID As Long, lng号源Id As Long, lng安排ID As Long, str号码 As String, strItem As String
    Dim obj出诊记录 As 出诊记录, obj出诊号源 As 出诊号源
    Dim blnFixedRule As Boolean
    
    Err = 0: On Error GoTo errHandler
    lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
    lng安排ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_安排ID))
    lng记录ID = Val(vsfRegistPlan.Cell(flexcpData, vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col)))
    strItem = vsfRegistPlan.Cell(flexcpData, 0, vsfRegistPlan.Col)
    str号码 = vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号码)
    
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_Modify '修改模板
        If ModifyPlanInfo(Me, mstrPrivs, mlngModule, mlng出诊ID) Then Call mfrmMain.NodeChanged("K0_" & mlng出诊ID)
    Case conMenu_Edit_Delete '删除模板
        If DeletePlan(mstrPrivs, mlng出诊ID, sccTitle.Tag) Then Call mfrmMain.NodeChanged("")
    Case conMenu_Edit_ModifyPlanItem '调整出诊项
        If lng号源Id <> 0 Or lng安排ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            If strItem = "其他规则" Then strItem = ""
            If frmEdit.ShowMe(Me, 3, Fun_Update, mlng出诊ID, lng号源Id, lng安排ID, strItem) Then
                Call RefreshOneData
            End If
        End If
    Case conMenu_Edit_ModifyUnitRegist '调整合作单位
        If lng号源Id <> 0 Or lng安排ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            Call frmEdit.ShowMe(Me, 3, Fun_UpdateUnit, mlng出诊ID, lng号源Id, lng安排ID, strItem)
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
    Case conMenu_Edit_ClearAllPlan '清除当前号源安排
        If MsgBox("你确定要清除号码为【" & str号码 & "】的所有安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlanBatch(mlng出诊ID, lng号源Id) Then Call RefreshOneData
    Case conMenu_Edit_ClearAll '清除所有号源安排
        If MsgBox("你确定要清除所有号源的安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If ZlClearPlanBatch(mlng出诊ID, 0, IIf(HavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID)) Then
            Call RefreshData(mbytFun, mlng出诊ID)
        End If
    Case conMenu_Edit_NextWeekNewPlan '生成下周安排
        Call NextNewPlanByTemplet(mlng出诊ID, False)
    Case conMenu_Edit_NextMonthNewPlan '生成下月安排
        Call NextNewPlanByTemplet(mlng出诊ID, True)
    Case conMenu_View_Refresh
        Call RefreshData(mbytFun, mlng出诊ID)
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 3 '查找方式
        mintFindType = Val(Right(Control.ID, 2)) - 1
        mcbsMain.RecalcLayout
        txtFind(1).Text = ""
        If txtFind(1).Visible And txtFind(1).Enabled Then txtFind(1).SetFocus
    End Select
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
    Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet, mrsPlanRecords, 3)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function AddNewPlanTemplet() As String
    '功能：新增模板
    Dim obj出诊安排 As New 出诊安排, frmPlanInfoEdit As New frmClinicPlanInfoEdit
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandler
    
    '检查是否有按月或按周排班的有效号源
    If CheckIsHavePlan(3, IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID)) = False Then
        MsgBox "模板只对按月或按周排班的号源有效，但当前无按月或按周排班的号源，请先到“基础设置>临床号源管理”中添加！", vbInformation, gstrSysName
        Exit Function
    End If
    
    obj出诊安排.排班方式 = 3 '排班方式：0-固定排班;1-按月排班;2-按周排班;3-模板
    If frmPlanInfoEdit.ShowMe(Me, mlngModule, 1, obj出诊安排) = False Then Exit Function
    
    obj出诊安排.出诊ID = zlDatabase.GetNextId("临床出诊表")
    'Zl_临床出诊表_Add
    '  --说明：在新增出诊表时，插入一条无号源信息的临床出诊安排
    strSQL = "Zl_临床出诊表_Add("
    '  操作类型_In Number,--1-模板，2-固定安排, 3-月安排，4-周安排
    strSQL = strSQL & "" & 1 & ","
    '  出诊id_In   临床出诊表.Id%Type,
    strSQL = strSQL & "" & obj出诊安排.出诊ID & ","
    '  出诊表名_In 临床出诊表.出诊表名%Type,
    strSQL = strSQL & "'" & obj出诊安排.出诊表名 & "',"
    '  站点_In     部门表.站点%Type,
    strSQL = strSQL & "'" & gstrNodeNo & "',"
    '  全院号源归属站点_In 部门表.站点%Type,
    strSQL = strSQL & "'" & gVisitPlan_ModulePara.str号源维护站点 & "',"
   '  操作员_In   临床出诊安排.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  操作时间_In 临床出诊安排.登记时间%Type := Null
    strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ","
    '  开始时间_In 临床出诊安排.开始时间%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  终止时间_In 临床出诊安排.终止时间%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  年份_In     临床出诊表.年份%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  月份_In     临床出诊表.月份%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  周数_In     临床出诊表.周数%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  应用范围_In 临床出诊表.应用范围%Type := Null,
    strSQL = strSQL & "" & obj出诊安排.应用范围 & ","
    '  科室id_In   临床出诊表.科室id%Type := Null,
    strSQL = strSQL & "" & ZVal(obj出诊安排.科室ID) & ","
    '  备注_In     临床出诊表.备注%Type := Null,
    strSQL = strSQL & "'" & obj出诊安排.备注 & "',"
    '  人员id_In   人员表.Id%Type := Null),
    strSQL = strSQL & "" & IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室"), "NULL", UserInfo.ID) & ","
    '  删除安排_In Number := 0,
    strSQL = strSQL & "" & "0" & ","
    '  模板类型_In 临床出诊表.模板类型%Type := Null
    strSQL = strSQL & "" & obj出诊安排.模板类型 & ")"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    '出诊表模板节点：K0_出诊ID
    AddNewPlanTemplet = "K0_" & obj出诊安排.出诊ID
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ModifyPlanInfo(frmMain As Form, ByVal strPrivs As String, _
    ByVal lngModule As Long, ByVal lng出诊ID As Long) As Boolean
    '修改模板信息
    Dim frmPlanInfoEdit As New frmClinicPlanInfoEdit
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim obj出诊安排 As New 出诊安排
    
    Err = 0: On Error GoTo errHandler
    strSQL = "Select a.ID, a.排班方式, a.出诊表名, a.应用范围, a.科室id, a.备注, a.模板类型, a.发布人" & vbNewLine & _
            "From 临床出诊表 A" & vbNewLine & _
            "Where a.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊表信息", lng出诊ID)
    If rsTemp.EOF Then Exit Function
    If zlStr.IsHavePrivs(strPrivs, "所有科室") = False Then
        If Nvl(rsTemp!发布人) <> UserInfo.姓名 Then
            MsgBox "你没有权限修改他人制定的模板！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With obj出诊安排
        .出诊ID = Val(Nvl(rsTemp!ID))
        .排班方式 = Val(Nvl(rsTemp!排班方式))
        .出诊表名 = Nvl(rsTemp!出诊表名)
        .应用范围 = Val(Nvl(rsTemp!应用范围))
        .科室ID = Val(Nvl(rsTemp!科室ID))
        .备注 = Nvl(rsTemp!备注)
        .模板类型 = Val(Nvl(rsTemp!模板类型))
    End With
    If frmPlanInfoEdit.ShowMe(frmMain, mlngModule, 1, obj出诊安排, True) = False Then Exit Function
    
    '保存数据
    'Zl_临床出诊表_Update
    strSQL = "Zl_临床出诊表_Update("
    '  操作类型_In Number, --1-模板，2-固定安排
    strSQL = strSQL & "" & 1 & ","
    '  Id_In       临床出诊表.Id%Type,
    strSQL = strSQL & "" & lng出诊ID & ","
    '  出诊表名_In 临床出诊表.出诊表名%Type := Null,
    strSQL = strSQL & "'" & obj出诊安排.出诊表名 & "',"
    '  开始时间_In 临床出诊安排.开始时间%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  终止时间_In 临床出诊安排.终止时间%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  应用范围_In 临床出诊表.应用范围%Type := Null,
    strSQL = strSQL & "" & obj出诊安排.应用范围 & ","
    '  科室id_In   临床出诊表.科室id%Type := Null,
    strSQL = strSQL & "" & ZVal(obj出诊安排.科室ID) & ","
    '  备注_In     临床出诊表.备注%Type := Null
    strSQL = strSQL & "'" & obj出诊安排.出诊表名 & "')"
    zlDatabase.ExecuteProcedure strSQL, frmMain.Caption
    
    ModifyPlanInfo = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub RefreshData(ByVal bytFun As Byte, Optional ByVal lng出诊ID As Long, Optional ByVal blnClear As Boolean)
    '功能：刷新安排详情数据
    '入数：
    '   strTitle - 标题显示
    '   lng出诊ID - 出诊ID
    '   bytFun - 1-月安排，2-周安排
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim int应用范围 As Integer
    
    Err = 0: On Error GoTo errHandler
    If blnClear Then
        mbytFun = bytFun
        mlng出诊ID = lng出诊ID
        
        sccTitle.Caption = "出诊安排>出诊模板"
        sccTitle.Tag = ""
        lblPlanInfo.Visible = False
        lblPlanInfo.Caption = "应用范围：         备注：                             "
        lblPlanInfo.ToolTipText = ""
        Set mrsPlanRecords = Nothing
        
        '改变菜单名称
        Call ZlUpdatePlanMenu(Me, mcbsMain, bytFun, IIf(HavePrivs(mstrPrivs, "所有科室"), 0, UserInfo.ID))
    
        strSQL = "Select a.出诊表名, a.应用范围, b.名称 As 应用科室, a.备注" & vbNewLine & _
                " From 临床出诊表 A, 部门表 B" & vbNewLine & _
                " Where a.科室Id = b.Id(+) And a.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊表信息", lng出诊ID)
        If Not rsTemp.EOF Then
            sccTitle.Caption = "出诊安排>" & Nvl(rsTemp!出诊表名)
            sccTitle.Tag = Nvl(rsTemp!出诊表名)
            lblPlanInfo.Visible = True
            int应用范围 = Val(Nvl(rsTemp!应用范围)) '0-本人;1-所属科室;2-全院通用
            lblPlanInfo.Caption = "应用范围：" & _
                Decode(int应用范围, 0, "本人", 1, "所属科室(" & Nvl(rsTemp!应用科室) & ")", "全院") & _
                "  备注：" & Left(Nvl(rsTemp!备注), 20)
            lblPlanInfo.ToolTipText = Nvl(rsTemp!备注)
        End If
        Call InitPlanGrid(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "安排")
    End If
    
    If lng出诊ID <> 0 Then
        '加载数据
        Screen.MousePointer = vbHourglass
        Set mrsPlanRecords = GetPlanRecords(bytFun = 1, lng出诊ID)
        Call ExecuteFilter
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub RefreshOneData(Optional ByVal blnReLoadData As Boolean = True)
    '刷新指定行号源数据
    Dim lng号源Id As Long, str收费项目 As String
    
    Err = 0: On Error GoTo errHandle
    '1.记录原数据，并获取新数据
    With vsfRegistPlan
        lng号源Id = Val(.TextMatrix(.Row, COL_号源ID))
        str收费项目 = .TextMatrix(.Row, COL_项目)
    End With
    
    If blnReLoadData Then
        '更新本地记录集
        Set mrsPlanRecords = GetPlanRecords(mbytFun = 1, mlng出诊ID)
    End If
    
    '2.更新界面
    mrsPlanRecords.Filter = "号源ID=" & lng号源Id & " And 收费项目='" & str收费项目 & "'"
    Call RefreshOnePlanData(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet, mrsPlanRecords, , , 3)
    mrsPlanRecords.Filter = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim strFindType As String
    Call GetRegInFor(g私有模块, Me.Name, "FindType", strFindType)
    mintFindType = Val(strFindType)
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    
    lblPlanInfo.Move sccTitle.Width - lblPlanInfo.Width - 100, sccTitle.Top + sccTitle.Height - lblPlanInfo.Height - 50
    With vsfRegistPlan
        .Left = 10
        .Top = sccTitle.Top + sccTitle.Height
        .Width = sccTitle.Width
        .Height = Me.ScaleHeight - .Top - 10
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

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If vsfRegistPlan.Visible And vsfRegistPlan.Enabled Then vsfRegistPlan.SetFocus
End Sub

Private Sub vsfRegistPlan_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    On Error Resume Next
    Call SetPlanGridRangeColor(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet, mstrOldSelRangePlan)
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
    Dim lng号源Id As Long, lng安排ID As Long
    Dim frmEdit As New frmClinicPlanEdit
    Dim strCurItem As String, blnUpdate As Boolean
    Dim lngCol As Long, lngRow As Long
    Dim strSort As String
    
    Err = 0: On Error GoTo errHandler
    lngCol = vsfRegistPlan.MouseCol
    lngRow = vsfRegistPlan.MouseRow
    If lngRow = 0 Or lngRow = 1 Then
        '排序
        If mrsPlanRecords Is Nothing Then Exit Sub
        If mrsPlanRecords.RecordCount = 0 Then Exit Sub
        strSort = GetPlanSortCircleStr(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet, lngRow, lngCol)
        If strSort <> "" Then
            mrsPlanRecords.Sort = strSort
            Screen.MousePointer = vbHourglass
            Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Templet, mrsPlanRecords, 3, , True)
            Screen.MousePointer = vbDefault
        End If
    Else
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        lng安排ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_安排ID))
        lngCol = GetPlanItemNameCol(vsfRegistPlan.Col)
        strCurItem = vsfRegistPlan.Cell(flexcpData, 0, lngCol)
        If lng号源Id = 0 And lng安排ID = 0 Then Exit Sub
        If strCurItem = "" Then Exit Sub
        
        blnUpdate = zlStr.IsHavePrivs(mstrPrivs, "模板管理")
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_是否临床排班)) = "" Then blnUpdate = False
        End If
        If strCurItem = "其他规则" Then strCurItem = ""
    
        If frmEdit.ShowMe(Me, 3, IIf(blnUpdate, Fun_Update, Fun_View), mlng出诊ID, lng号源Id, lng安排ID, strCurItem) Then
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

    '弹出菜单
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

Private Sub picImgPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    
    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfRegistPlan, lngLeft, lngTop, picImgPlan.Height)
    Call zl_vsGrid_Para_Save(mlngModule, vsfRegistPlan, Me.Name, "安排")
End Sub

Private Function NextNewPlanByTemplet(ByVal lng模板ID As Long, Optional ByVal blnMonth As Boolean) As Boolean
    '根据模板生成新安排
    Dim strSQL As String, rsTemp As ADODB.Recordset, lng出诊ID As Long
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim dtStart As Date, dtEnd As Date
    Dim strName As String, strKey As String, blnDeletePlan As Boolean
    Dim cllPlan As Collection, i As Integer
    Dim dtCurrent As Date
    
    Err = 0: On Error GoTo errHandler
    If lng模板ID = 0 Then Exit Function
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
        
        'zl_临床出诊表_Addbytemplet(
        strSQL = "zl_临床出诊表_Addbytemplet("
        '模板id_In   临床出诊表.Id%Type,
        strSQL = strSQL & "" & lng模板ID & ","
        '人员id_In   人员表.Id%Type,
        strSQL = strSQL & "" & IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室"), "NULL", UserInfo.ID) & ","
        '出诊id_In   临床出诊表.Id%Type,
        strSQL = strSQL & "" & lng出诊ID & ","
        '排班方式_In 临床出诊表.排班方式%Type,
        strSQL = strSQL & "" & IIf(blnMonth, 1, 2) & ","
        '出诊表名_In 临床出诊表.出诊表名%Type,
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
        strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ","
        '站点_In       部门表.站点%Type,
        strSQL = strSQL & "'" & gstrNodeNo & "',"
        '全院号源归属站点_In 部门表.站点%Type,
        strSQL = strSQL & "'" & gVisitPlan_ModulePara.str号源维护站点 & "',"
        '删除安排_In Number:=0
        strSQL = strSQL & "" & IIf(blnDeletePlan, 1, 0) & ")"
    
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Next
    If cllPlan.Count > 1 Then gcnOracle.CommitTrans
    
    'XX月出诊表节点：K2_年份_月份
    'XX周出诊表节点：K3_年份_月份_周数
    Call mfrmMain.NodeChanged(strKey)
    NextNewPlanByTemplet = True
    
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

Private Function GetPlanRecords(ByVal blnMonth As Boolean, ByVal lng出诊ID As Long) As ADODB.Recordset
    '功能：获取安排记录
    Dim strSQL As String, strSqlSub As String
    Dim strWhere As String
    Dim str排序号码 As String

    Err = 0: On Error GoTo errHandler
    str排序号码 = IIf(gVisitPlan_ModulePara.byt号码比较方式 = 0, "a.号码", "Lpad(a.号码,5,'0')")
    strSqlSub = "       " & str排序号码 & " As 排序号码, a.Id As 号源id, a.号类, a.号码, Nvl(a.是否建病案, 0) As 是否建病案,a.预约天数, a.出诊频次," & vbNewLine & _
                "       Decode(a.假日控制状态, 1, '开放预约', 2, '禁止预约', 3, '受节假日设置控制', '不上班') As 假日控制状态," & vbNewLine & _
                "       Nvl(a.是否临床排班, 0) As 是否临床排班, Decode(a.排班方式, 1, '按月排班', 2, '按周排班', '固定排班') As 排班方式," & vbNewLine & _
                "       f.名称 As 科室, f.简码 As 科室简码, Nvl(a.是否假日换休, 0) As 是否假日换休," & vbNewLine
    
    '没有"所有科室"权限的操作员只能操作自己所属科室的号源
    If HavePrivs(mstrPrivs, "所有科室") = False Then
        strWhere = "      And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = [2])"
    End If
    
    strSQL = "Select m.Id, m.出诊id, m.号源id, m.项目ID, m.医生ID, m.医生姓名, m.排班规则, m.开始时间, m.终止时间, n.限制项目, n.上班时段, n.限号数, n.限约数, n.预约控制" & vbNewLine & _
            " From 临床出诊安排 M, 临床出诊限制 N" & vbNewLine & _
            " Where m.Id = n.安排id(+) And m.出诊id = [1] And Nvl(m.排班规则, 0) <> 6"
    '按特定日期排班的单独查询
    strSQL = strSQL & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select ID, 出诊id, 号源id, 项目ID, 医生ID, 医生姓名, 排班规则, 开始时间, 终止时间," & vbNewLine & _
            "        f_List2str(Cast(Collect(限制项目 || '' Order By 限制项目) As t_Strlist)) As 限制项目, 上班时段, 限号数, 限约数, 预约控制" & vbNewLine & _
            " From (Select Count(1) Over(Partition By m.号源id, n.限制项目) As 组号, m.Id, m.出诊id, m.号源id, m.项目ID, m.医生ID, m.医生姓名, m.排班规则," & vbNewLine & _
            "               m.开始时间, m.终止时间, To_Number(RTrim(n.限制项目, '日')) As 限制项目, n.上班时段, n.限号数, n.限约数, n.预约控制" & vbNewLine & _
            "        From 临床出诊安排 M, 临床出诊限制 N" & vbNewLine & _
            "        Where m.Id = n.安排id(+) And m.出诊id = [1] And Nvl(m.排班规则, 0) = 6)" & vbNewLine & _
            " Group By ID, 出诊id, 组号, 号源id, 项目id, 医生id, 医生姓名, 排班规则, 开始时间, 终止时间, 上班时段, 限号数, 限约数, 预约控制"
    
    strSQL = "Select b.出诊id, b.Id As 安排id, " & _
            "       Decode(b.ID,Null,e.名称,m.名称) As 收费项目, Decode(b.ID,Null,a.医生姓名,b.医生姓名) As 医生姓名," & vbNewLine & _
            "       Decode(b.ID,Null,g.简码,n.简码) As 医生简码, Decode(b.ID,Null,g.专业技术职务,n.专业技术职务) as 医生职称," & vbNewLine & _
            "       Decode(b.id,Null,i.标识符,j.标识符) as 标识符," & vbNewLine & _
            "        b.排班规则, b.开始时间, b.终止时间," & strSqlSub & vbNewLine & _
            "        b.Id As 记录id, b.限制项目, b.上班时段, b.限号数, b.限约数, b.预约控制 As 预约控制方式" & vbNewLine & _
            " From 临床出诊号源 A, (" & strSQL & ") B, 收费项目目录 E, 部门表 F, 人员表 G, 收费项目目录 M, 人员表 N,专业技术职务 I,专业技术职务 J" & vbNewLine & _
            " Where a.Id = b.号源id(+) And a.科室id = f.Id" & vbNewLine & _
            "       And a.项目id = e.Id And a.医生ID= g.ID(+) And b.项目id = m.Id(+) And b.医生ID= n.ID(+)" & vbNewLine & _
            "       And g.专业技术职务=i.名称(+) And n.专业技术职务=j.名称(+)" & vbNewLine & _
            "       And Nvl(a.排班方式, 0) = [3]" & strWhere & vbNewLine & _
            "       And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
            "       And (e.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.撤档时间 Is Null)" & vbNewLine & _
            "       And (f.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.撤档时间 Is Null)" & vbNewLine & _
            "       And (g.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.撤档时间 Is Null)" & vbNewLine & _
            "       And Nvl(Nvl(f.站点,[5]),Nvl([4],'-')) = Nvl([4],'-')" & vbNewLine & _
            " Order By " & str排序号码 & ", b.开始时间, b.终止时间, b.限制项目, b.上班时段"
    Set GetPlanRecords = zlDatabase.OpenSQLRecord(strSQL, "获取排班信息", lng出诊ID, UserInfo.ID, IIf(blnMonth, 1, 2), _
        gstrNodeNo, gVisitPlan_ModulePara.str号源维护站点)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeletePlan(ByVal strPrivs As String, ByVal lng出诊ID As Long, _
    strTableName As String) As Boolean
    '功能：删除出诊表
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lng出诊ID = 0 Then Exit Function
    
    If zlStr.IsHavePrivs(strPrivs, "所有科室") = False Then
        strSQL = "Select 1 From 临床出诊表 A" & vbNewLine & _
                " Where a.排班方式 = 3 And a.ID=[1] and a.发布人=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查权限", lng出诊ID, UserInfo.姓名)
        If rsTemp.EOF Then
            MsgBox "你没有权限删除他人制定的模板！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If MsgBox("你确定要删除【" & strTableName & "】吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Function
    End If
    
    DeletePlan = ZlDeletePlan(lng出诊ID)
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
