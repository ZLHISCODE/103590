VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicPlanTempletByDayManage 
   BorderStyle     =   0  'None
   Caption         =   "出诊月模板管理"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   9360
      MaxLength       =   100
      TabIndex        =   7
      Top             =   930
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.PictureBox picSelectWeek 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   90
      ScaleHeight     =   345
      ScaleWidth      =   10845
      TabIndex        =   0
      Top             =   450
      Width           =   10845
      Begin VB.OptionButton optWeek 
         Caption         =   "第5周"
         Height          =   195
         Index           =   5
         Left            =   5040
         TabIndex        =   6
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "第4周"
         Height          =   195
         Index           =   4
         Left            =   4050
         TabIndex        =   5
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "第3周"
         Height          =   195
         Index           =   3
         Left            =   3045
         TabIndex        =   4
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "第2周"
         Height          =   195
         Index           =   2
         Left            =   2055
         TabIndex        =   3
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "第1周"
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   2
         Top             =   90
         Width           =   795
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "全部"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   90
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfRegistPlan 
      Height          =   2085
      Left            =   390
      TabIndex        =   8
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
      FormatString    =   $"frmClinicPlanTempletByDayManage.frx":0000
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
         Picture         =   "frmClinicPlanTempletByDayManage.frx":0075
         ScaleHeight     =   135
         ScaleWidth      =   150
         TabIndex        =   9
         Top             =   90
         Width           =   150
      End
   End
   Begin VB.Line lineSplit 
      BorderColor     =   &H8000000A&
      X1              =   0
      X2              =   3990
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   6915
      Left            =   0
      Top             =   0
      Width           =   11595
   End
   Begin VB.Label lblPlanInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用范围：所属科室(门诊内科)  备注：发鬼地方个梵蒂冈发鬼地方个法规定功夫功夫"
      Height          =   180
      Left            =   3960
      TabIndex        =   10
      Top             =   150
      Width           =   6840
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   90
      TabIndex        =   11
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
End
Attribute VB_Name = "frmClinicPlanTempletByDayManage"
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
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CopyPlan, "复制安排(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PastPlan, "粘贴安排(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearCurPlan, "清除当前安排(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAllPlan, "清除当前号源安排(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearAll, "清除所有号源安排(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyToDay, "应用于“所有单日”(&D)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextMonthNewPlan, "生成月出诊表(&N)"): cbrControl.BeginGroup = True
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
    Dim bln禁止预约 As Boolean, blnSelectedNotNull As Boolean
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
        bln禁止预约 = Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, GetPlanItemNameCol(vsfRegistPlan.Col) + 2)) = "-"
        blnSelectedNotNull = vsfRegistPlan.Col >= gPlanGrid_FixedCols _
            And Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) <> "" '选择行是否有数据
        blnEnabled = mlng出诊ID > 0
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_是否临床排班)) = "" Then blnEnabled = False
        End If
        Control.Enabled = Control.Visible And blnEnabled And blnPlanDataCol And blnSelectedNotNull And Not bln禁止预约
    Case conMenu_Edit_AllStartNO '全部启用序号控制
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        Control.Enabled = Control.Visible And mlng出诊ID <> 0
    Case conMenu_Edit_AllStopNO '全部取消序号控制
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        Control.Enabled = Control.Visible And mlng出诊ID <> 0
        
    Case conMenu_Edit_CopyPlan '复制安排
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        blnSelectedNotNull = vsfRegistPlan.Col >= gPlanGrid_FixedCols _
            And Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) <> "" '选择行是否有数据
        Control.Enabled = Control.Visible And mlng出诊ID <> 0 And blnSelectedNotNull
    Case conMenu_Edit_PastPlan '粘贴安排
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        blnEnabled = mlng出诊ID > 0
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_是否临床排班)) = "" Then blnEnabled = False
        End If
        Control.Enabled = Control.Visible And blnEnabled And mlngCopyPlanID <> 0
    Case conMenu_Edit_ClearCurPlan '清除当前安排
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        blnSelectedNotNull = vsfRegistPlan.Col >= gPlanGrid_FixedCols _
            And Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) <> "" '选择行是否有数据
        blnEnabled = mlng出诊ID > 0
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_是否临床排班)) = "" Then blnEnabled = False
        End If
        Control.Enabled = Control.Visible And blnEnabled And blnSelectedNotNull
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
    Case conMenu_Edit_ApplyToDay '应用于“所有单日”
        Control.Visible = HavePrivs(mstrPrivs, "模板管理") And mfrmMain.mFunListActived = False
        blnSelectedNotNull = vsfRegistPlan.Col >= gPlanGrid_FixedCols _
            And Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, vsfRegistPlan.Col)) <> "" '选择行是否有数据
        blnEnabled = mlng出诊ID > 0
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_是否临床排班)) = "" Then blnEnabled = False
        End If
        Control.Enabled = Control.Visible And blnEnabled And blnSelectedNotNull
        
    Case conMenu_Edit_NextMonthNewPlan '生成下月安排
        Control.Visible = HavePrivs(mstrPrivs, "出诊安排")
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
        If frmClinicPlanTempletManage.ModifyPlanInfo(Me, mstrPrivs, mlngModule, mlng出诊ID) Then Call mfrmMain.NodeChanged("K0_" & mlng出诊ID)
    Case conMenu_Edit_Delete '删除模板
        If frmClinicPlanTempletManage.DeletePlan(mstrPrivs, mlng出诊ID, sccTitle.Tag) Then Call mfrmMain.NodeChanged("")
    Case conMenu_Edit_ModifyPlanItem '调整出诊项
        If lng号源Id <> 0 Or lng安排ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            If frmEdit.ShowMe(Me, 4, Fun_Update, mlng出诊ID, lng号源Id, lng安排ID, strItem) Then
                Call RefreshOneData
            End If
        End If
    Case conMenu_Edit_ModifyUnitRegist '调整合作单位
        If lng号源Id <> 0 Or lng安排ID <> 0 Then
            Set frmEdit = New frmClinicPlanEdit
            Call frmEdit.ShowMe(Me, 4, Fun_UpdateUnit, mlng出诊ID, lng号源Id, lng安排ID, strItem)
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
        If MsgBox("你确定要清除号码为【" & str号码 & "】【" & FormatApplyToStr(strItem) & "】的安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If ZlClearPlan(lng安排ID, FormatApplyToStr(strItem)) Then
            Call RefreshOneData
            If mlngCopyPlanID = lng安排ID And mstrCopyPlanItem = strItem Then
                mlngCopyPlanID = 0: mstrCopyPlanItem = ""
            End If
        End If
    Case conMenu_Edit_ClearAllPlan '清除当前号源安排
        If MsgBox("你确定要清除号码为【" & str号码 & "】的所有安排吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
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
    Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, mbytFun)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub RefreshData(ByVal bytFun As Byte, ByVal lng出诊ID As Long, Optional ByVal blnClear As Boolean, _
    Optional ByVal intYear As Integer, Optional ByVal intMonth As Integer, Optional ByVal strTitle As String)
    '功能：刷新安排详情数据
    '入参：
    '   bytFun - 1-月安排，2-周安排
    '   lng出诊ID - 出诊ID
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim i As Integer, dtStartDate As Date, dtEndDate As Date
    Dim int应用范围 As Integer
    
    Err = 0: On Error GoTo errHandler
    
    If blnClear Then
        mbytFun = bytFun: mlng出诊ID = lng出诊ID
        sccTitle.Caption = "出诊安排>出诊模板"
        sccTitle.Tag = ""
        
        For i = optWeek.LBound To optWeek.UBound
            optWeek(i).Visible = True
        Next
        optWeek(0).Value = True
        
        lblPlanInfo.Visible = False
        lblPlanInfo.Caption = "应用范围：         备注：                             "
        lblPlanInfo.ToolTipText = ""
        Set mrsPlanRecords = Nothing
        mlngCopyPlanID = 0: mstrCopyPlanItem = ""
        
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
        
        '显示日期范围确定,缺省1900年1月
        dtStartDate = CDate("1900-01-01"): dtEndDate = CDate("1900-01-31")
        Call InitPlanGrid(vsfRegistPlan, gPlanGrid_DataStyle.Data_MonthTemplet, dtStartDate, dtEndDate)
        Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "安排")
        Call Form_Resize
    End If
    
    Screen.MousePointer = vbHourglass
    If lng出诊ID = 0 Then
        Set mrsPlanRecords = Nothing
    Else
        Set mrsPlanRecords = GetPlanRecords(bytFun = 1, lng出诊ID)
    End If
    '加载数据
    Call ExecuteFilter
    Screen.MousePointer = vbDefault
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
    Call RefreshOnePlanData(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, , , 3)
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
    lblPlanInfo.Move sccTitle.Width - lblPlanInfo.Width - 100, sccTitle.Top + sccTitle.Height - lblPlanInfo.Height - 50
    
    picSelectWeek.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, sccTitle.Width
    lineSplit.X1 = sccTitle.Left + 10
    lineSplit.Y1 = picSelectWeek.Top + picSelectWeek.Height
    lineSplit.X2 = sccTitle.Width
    lineSplit.Y2 = lineSplit.Y1
    With vsfRegistPlan
        .Left = sccTitle.Left + 10
        .Top = picSelectWeek.Top + picSelectWeek.Height + 20
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
    
    objOut.Title.Text = sccTitle.Tag & "清单"
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
    Dim intWeek As Integer
    Dim dtStart As Date, dtEnd As Date
    
    Err = 0: On Error GoTo errHandler
    intWeek = index
    Screen.MousePointer = vbHourglass
    Select Case index
    Case 0
        dtStart = CDate("1900-01-01"): dtEnd = CDate("1900-01-31")
    Case 5
        dtStart = CDate("1900-01-01") + 7 * (index - 1): dtEnd = CDate("1900-01-31")
    Case Else
        dtStart = CDate("1900-01-01") + 7 * (index - 1): dtEnd = CDate("1900-01-01") + 7 * index - 1
    End Select
    Call InitPlanGrid(vsfRegistPlan, gPlanGrid_DataStyle.Data_MonthTemplet, dtStart, dtEnd, False)
    Call vsGrid_Para_Restore_Plan(mlngModule, vsfRegistPlan, Me.Name, "安排")
    '使用缓存数据
    Call ExecuteFilter
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
    Dim lng号源Id As Long, lng安排ID As Long
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
            Call LoadPlanDataByRecordset(vsfRegistPlan, gPlanGrid_DataStyle.Data_Plan, mrsPlanRecords, mbytFun, , True)
            Screen.MousePointer = vbDefault
        End If
    Else
        lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
        lng安排ID = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_安排ID))
        lngCol = GetPlanItemNameCol(vsfRegistPlan.Col)
        strCurItem = vsfRegistPlan.Cell(flexcpData, 0, lngCol)
        If lng号源Id = 0 And lng安排ID = 0 Then Exit Sub
        If IsDate(strCurItem) = False Then Exit Sub
        
        blnUpdate = zlStr.IsHavePrivs(mstrPrivs, "模板管理")
        If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
            '没有“所有科室”权限时，只能调整“允许临床科室排班”的号源
            If Trim(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_是否临床排班)) = "" Then blnUpdate = False
        End If
        
        If frmEdit.ShowMe(Me, 4, IIf(blnUpdate, Fun_Update, Fun_View), mlng出诊ID, lng号源Id, lng安排ID, strCurItem) Then
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
            " From 临床出诊限制 A,临床出诊安排 C, 临床出诊号源 B" & vbNewLine & _
            " Where a.安排ID = c.ID And c.号源ID = b.ID And c.ID = [1] And a.限制项目 = [2] And a.上班时段 Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng原安排ID, FormatApplyToStr(str原项目))
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
        strSQL = strSQL & "" & "6" & "," '固定"6-特定日期"规则
        '是否周六出诊_In 临床出诊安排.是否周六出诊%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '是否周日出诊_In 临床出诊安排.是否周日出诊%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '开始时间_In     临床出诊安排.开始时间%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '终止时间_In     临床出诊安排.终止时间%Type,
        strSQL = strSQL & "" & "NULL" & ","
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
        If ZlPlanApplyTo(0, lng原安排ID, FormatApplyToStr(str原项目), lng安排ID, FormatApplyToStr(strApplyItem)) = False Then
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
    Dim lng号源Id As Long
    
    Err = 0: On Error GoTo errHandler
    If lng安排ID = 0 Or strCurDate = "" Then Exit Function
    If IsDate(strCurDate) = False Then Exit Function
    
    dtStart = CDate("1900-01-01"): dtEnd = CDate("1900-01-31")
    lng号源Id = Val(vsfRegistPlan.TextMatrix(vsfRegistPlan.Row, COL_号源ID))
    
    intDoubleDay = Day(strCurDate) Mod 2 '单日还是双日
    dtCur = dtStart
    Do While DateDiff("d", dtCur, dtEnd) >= 0
        If DateDiff("d", strCurDate, dtCur) <> 0 And (Day(dtCur) Mod 2) = intDoubleDay Then
            strApply = strApply & "|" & Format(dtCur, "yyyy-mm-dd")
        End If
        dtCur = DateAdd("d", 1, dtCur)
    Loop
    If strApply <> "" Then strApply = Mid(strApply, 2)
    
    If strApply = "" Then Exit Function
    strApply = FormatApplyToStr(strApply)
    If CheckExistRecord(lng号源Id, strApply, , True, lng安排ID) Then
        If MsgBox("注意：" & vbCrLf & _
                  "      被应用的日期当前已存在出诊安排，应用后这部分安排将会被覆盖！是否仍要应用？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    ApplyToDay = ZlPlanApplyTo(0, lng安排ID, FormatApplyToStr(Format(strCurDate, "yyyy-mm-dd")), lng安排ID, strApply)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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

Private Function GetPlanRecords(ByVal blnMonth As Boolean, Optional ByVal lng出诊ID As Long) As ADODB.Recordset
    '功能：获取安排记录
    '入数：
    '   blnMonth - 是否月排班
    '   lng出诊ID   - 出诊ID
    Dim strSQL As String, strSqlSub As String
    Dim strWhere As String
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
    
    '将所有号源缺省提取出来
    strSQL = "Select 1 As 是否有效," & vbNewLine & _
            "        b.出诊id, b.Id As 安排id, " & vbNewLine & strSqlSub & _
            "        Decode(b.ID,Null,e.名称,m.名称) As 收费项目, Decode(b.ID,Null,a.医生姓名,b.医生姓名) As 医生姓名," & vbNewLine & _
            "        Decode(b.ID,Null,g.简码,n.简码) As 医生简码, Decode(b.ID,Null,g.专业技术职务,n.专业技术职务) as 医生职称," & vbNewLine & _
            "        Decode(b.ID,Null,i.标识符,j.标识符) As 标识符 ," & vbNewLine & _
            "        c.Id As 记录id, To_Date(Decode(c.限制项目, Null, '', '1900-01-' || Replace(c.限制项目, '日', '')), 'yyyy-mm-dd') As 出诊日期, " & vbNewLine & _
            "        c.上班时段, c.限号数, c.限约数, b.开始时间, b.终止时间, " & vbNewLine & _
            "        NULL As 已挂数, NULL As 已约数, c.预约控制 As 预约控制方式, NULL As 是否临时出诊, NULL As 停诊开始时间, NULL As 停诊终止时间, NULL As 停诊原因, NULL As 替诊医生姓名,NULL As 是否锁定" & vbNewLine & _
            " From 临床出诊号源 A, (Select 出诊id, ID, 号源id, 项目id, 医生ID, 医生姓名, 开始时间, 终止时间, 审核时间 From 临床出诊安排 Where 出诊id = [1]) B," & vbNewLine & _
            "      临床出诊限制 C, 收费项目目录 E, 部门表 F, 人员表 G, 收费项目目录 M, 人员表 N,专业技术职务 I,专业技术职务 J" & vbNewLine & _
            " Where a.Id = b.号源id(+) And b.Id = c.安排id(+) And a.科室id = f.Id" & vbNewLine & _
            "       And g.专业技术职务=i.名称(+) And n.专业技术职务=j.名称(+)" & vbNewLine & _
            "       And a.项目id = e.Id And a.医生ID = g.ID(+) And b.项目id = m.Id(+) And b.医生ID = n.ID(+)" & vbNewLine & _
            "       And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
            "       And (e.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or e.撤档时间 Is Null)" & vbNewLine & _
            "       And (f.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or f.撤档时间 Is Null)" & vbNewLine & _
            "       And (g.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or g.撤档时间 Is Null)" & vbNewLine & _
            "       And Nvl(a.排班方式,0) = [2]" & strWhere & vbNewLine & _
            "       And Nvl(Nvl(f.站点,[5]),Nvl([4],'-')) = Nvl([4],'-')" & vbNewLine & _
            " Order By " & str排序号码 & ", 出诊日期, 上班时段"
    Set GetPlanRecords = zlDatabase.OpenSQLRecord(strSQL, "获取排班信息", lng出诊ID, IIf(blnMonth, 1, 2), UserInfo.ID, _
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
