VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGradeManage 
   Caption         =   "电子病案评分"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   12240
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   0
      Left            =   5280
      ScaleHeight     =   3585
      ScaleWidth      =   4470
      TabIndex        =   3
      Top             =   1620
      Width           =   4470
      Begin XtremeSuiteControls.TabControl tbcTask 
         Height          =   1830
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   9735
      TabIndex        =   2
      Top             =   105
      Width           =   1125
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   1
      Left            =   450
      ScaleHeight     =   3855
      ScaleWidth      =   3690
      TabIndex        =   0
      Top             =   1320
      Width           =   3690
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2385
         Picture         =   "frmGradeManage.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   60
         Width           =   285
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2520
         Index           =   0
         Left            =   30
         TabIndex        =   1
         Top             =   375
         Width           =   2670
         _cx             =   4710
         _cy             =   4445
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Begin VB.TextBox txt科室 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   450
         TabIndex        =   8
         Top             =   30
         Width           =   2250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   30
         TabIndex        =   6
         Top             =   75
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7620
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15743
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmGradeManage.frx":0049
      Left            =   630
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmGradeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
''窗体级变量定义
''######################################################################################################################
'Private mstrPrivs As String
'Private mblnStartUp As Boolean
'Private mblnAllowClose As Boolean
'Private mstrCondition As String
'Private mstrFindKey As String
'Private mlngTmp As Long
'Private mobjFindKey As CommandBarControl
'Private mclsVsf(0) As clsVsf
'Private mlngModul As Long
'Private mintIndex As Integer
'Private mbytMode As Byte
'Private mfrmChildMedrec As frmChildMedrec
'Private WithEvents mfrmGradeEdit As frmGradeEdit
'
''######################################################################################################################
'
'Public Property Get 模块号() As Long
'    模块号 = mlngModul
'End Property
'
'Private Function InitCommandBar() As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim objMenu As CommandBarPopup
'    Dim objBar As CommandBar
'    Dim objPopup As CommandBarPopup
'    Dim objControl As CommandBarControl
'    Dim cbrCustom As CommandBarControlCustom
'
'    '------------------------------------------------------------------------------------------------------------------
'    '初始设置
'
'    Call CommandBarInit(cbsMain)
'
'    '------------------------------------------------------------------------------------------------------------------
'    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
'
'    cbsMain.ActiveMenuBar.Title = "菜单"
'    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
'
'    '文件
'    '------------------------------------------------------------------------------------------------------------------
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
'    objMenu.ID = conMenu_FilePopup
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "全部打印(&L)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
'
'    '编辑
'    '------------------------------------------------------------------------------------------------------------------
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
'    objMenu.ID = conMenu_EditPopup
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "病案评分(&A)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "修改结果(&M)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Append, "重新评分(&R)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除结果(&D)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "通过审核(&P)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "取消审核(&C)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SelAll, "全部选中(&L)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ClsAll, "取消选择(&S)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Select, "反向选择(&B)")
'
'    '查看
'    '------------------------------------------------------------------------------------------------------------------
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
'    objMenu.ID = conMenu_ViewPopup
'    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
'    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
'    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
'    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Filter, "病案检索(&F)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
'
'
'    '帮助
'    '------------------------------------------------------------------------------------------------------------------
'    Call CreateHelpMenu(cbsMain)
'
'    '主菜单右侧的查找
'    '------------------------------------------------------------------------------------------------------------------
'    cbsMain.ActiveMenuBar.SetIconSize 16, 16
'    mstrFindKey = GetPara("定位依据", mlngModul, True, "No")
'    If mstrFindKey = "" Then mstrFindKey = "No"
'    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
'    mobjFindKey.IconId = conMenu_View_Find
'    mobjFindKey.Flags = xtpFlagRightAlign
'    mobjFindKey.STYLE = xtpButtonIconAndCaption
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.姓名", , , "姓名")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.住院号", , , "住院号")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.床号", , , "床号")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&4.就诊卡号", , , "就诊卡号")
'    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, ""): cbrCustom.Handle = txtLocation.Hwnd: cbrCustom.Flags = xtpFlagRightAlign
'    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "前一条"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon
'    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "后一条"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon
'
'    '工具栏定义:包括公共部份
'    '------------------------------------------------------------------------------------------------------------------
'    Set objBar = cbsMain.Add("标准", xtpBarTop)
'    objBar.ContextMenuPresent = False
'    objBar.ShowTextBelowIcons = False
'    objBar.EnableDocking xtpFlagStretched
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "评分", True)
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Audit, "审核")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "检索", True)
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
'
'    '命令的快键绑定:公共部份主界面已处理
'    '------------------------------------------------------------------------------------------------------------------
'    With cbsMain.KeyBindings
'        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
'        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
'        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
'        .Add FCONTROL, vbKeyV, conMenu_File_Preview
'        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '评分
'        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify          '修改
'        .Add FCONTROL, vbKeyR, conMenu_Edit_Append          '重评
'        .Add FCONTROL, vbKeyF, conMenu_View_Filter          '过滤
'        .Add 0, vbKeyDelete, conMenu_Edit_Delete            '删除
'        .Add 0, vbKeyF3, conMenu_View_Location              '定位
'        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      '前一条
'        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '后一条
'    End With
'
'End Function
'
'Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim intLoop As Integer
'    Dim intRow As Integer
'    Dim rs As New ADODB.Recordset
'    Dim rsSQL As New ADODB.Recordset
'    Dim strTmp As String
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    Call SQLRecord(rsSQL)
'
'    Select Case strCommand
'    '------------------------------------------------------------------------------------------------------------------
'    Case "初始控件"
'
'        Set mclsVsf(0) = New clsVsf
'        With mclsVsf(0)
'            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
'            Call .ClearColumn
'            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
'            Call .AppendColumn("No", 900, flexAlignLeftCenter, flexDTString, "", , True)
'            Call .AppendColumn("申请人", 810, flexAlignLeftCenter, flexDTString, "", , True)
'            Call .AppendColumn("记录状态", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
'            Call .AppendColumn("申请时间", 1440, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd", , True)
'            Call .AppendColumn("申请期限", 1440, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd", , True)
'            Call .AppendColumn("申请理由", 1500, flexAlignLeftCenter, flexDTString, "", , True)
'            .AppendRows = True
'        End With
'
'        '初始菜单及工具栏
'        '--------------------------------------------------------------------------------------------------------------
'        Call InitCommandBar
'
'        '划分停靠区域
'        '--------------------------------------------------------------------------------------------------------------
'        Dim objPane As Pane
'        Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "病人信息": objPane.Options = PaneNoCaption
'        Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, Nothing): objPane.Title = "详细内容": objPane.Options = PaneNoCaption
'
'
'        dkpMain.SetCommandBars cbsMain
'        Call DockPannelInit(dkpMain)
'
'
'        Call TabControlInit(tbcTask)
'        With tbcTask
'            .PaintManager.BoldSelected = True
'
'            Set mfrmGradeEdit = New frmGradeEdit
'            Set mfrmChildMedrec = New frmChildMedrec
'
'            Call mfrmGradeEdit.InitData(Me, True)
'
'            .InsertItem 0, "病案评分", mfrmGradeEdit.Hwnd, 0
'            .InsertItem 1, "首页记录", mfrmChildMedrec.Hwnd, 0
'
'            .Item(0).Selected = True
'
'        End With
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "初始数据"
'
''        '创建过滤条件项目，并进行初始化
''        Call ParamCreate(mrsCondition)
''        Call ParamAdd(mrsCondition, "开始单据号", "")
''        Call ParamAdd(mrsCondition, "结束单据号", "")
''        Call ParamAdd(mrsCondition, "申请人", "")
''        Call ParamAdd(mrsCondition, "批准人", "")
''        Call ParamAdd(mrsCondition, "拒绝人", "")
''
''        Call ParamAdd(mrsCondition, "新登记单据", "1")
''        Call ParamAdd(mrsCondition, "登记开始日期", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
''        Call ParamAdd(mrsCondition, "登记结束日期", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
''        Call ParamAdd(mrsCondition, "已批准单据", "0")
''        Call ParamAdd(mrsCondition, "批准开始日期", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
''        Call ParamAdd(mrsCondition, "批准结束日期", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
''        Call ParamAdd(mrsCondition, "已拒绝单据", "0")
''        Call ParamAdd(mrsCondition, "拒绝开始日期", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
''        Call ParamAdd(mrsCondition, "拒绝结束日期", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
''
''        '读取缺省的借阅申请登记查询时间范围
''        strTmp = GetPara("登记缺省范围", mlngModul, True, "今  天")
''        If strTmp = "" Then strTmp = "今  天"
''        Call ParamWrite(mrsCondition, "登记开始日期", GetDateTime(strTmp, 1))
''        Call ParamWrite(mrsCondition, "登记结束日期", GetDateTime(strTmp, 2))
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "控件状态"
'
''        If vsf(0).Enabled <> Not DataChanged Then
''            vsf(0).Enabled = Not DataChanged
''            vsf(0).ForeColor = IIf(DataChanged, COLOR.深灰色, COLOR.黑色)
''            tbcTask.Enabled = Not DataChanged
''        End If
''        stbThis.Panels(3).Enabled = DataChanged
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "刷新状态"
'
''        If Val(vsf(0).RowData(vsf(0).Row)) = 0 Then
''            strTmp = "当前还没有任何电子病案借阅申请单！"
''        Else
''            strTmp = "共有 " & vsf(0).Rows - 1 & " 个电子病案借阅申请单！"
''        End If
''
''        stbThis.Panels(2).Text = strTmp
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "刷新数据"
'
'        Call ExecuteCommand("读取病人记录")
'        Call ExecuteCommand("读取病案评分")
'        Call ExecuteCommand("读取首页记录")
'
'        Call ExecuteCommand("刷新状态")
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "过滤数据"
'
''        mrsCondition.Filter = ""
''        ExecuteCommand = frmCISBorrowFilter.ShowPara(Me, mrsCondition)
''
''        GoTo endHand
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "读取申请单据"
'
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "读取申请内容"
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "前一条"
'        With vsf(0)
'            If .Row > 1 Then
'                .Row = .Row - 1
'                .ShowCell .Row, .Col
'            End If
'        End With
'    '------------------------------------------------------------------------------------------------------------------
'    Case "后一条"
'        With vsf(0)
'            If .Row < .Rows - 1 Then
'                .Row = .Row + 1
'                .ShowCell .Row, .Col
'            End If
'        End With
'    '------------------------------------------------------------------------------------------------------------------
'    Case "读注册表"
'
'        If Val(GetPara("使用个性化风格", , , True)) = 1 Then
'            '使用个性化设置
'
'            mstrFindKey = Trim(GetPara("定位依据", mlngModul, True, "No"))
'            mclsVsf(0).LoadStateFromString Trim(GetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(0)), ""))
'        End If
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "写注册表"
'        If Val(GetPara("使用个性化风格", , , True)) = 1 Then
'            '使用个性化设置
'            Call SetPara("定位依据", mstrFindKey, mlngModul, True)
'        End If
'        Call SetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(0)), mclsVsf(0).SaveStateToString)
'    End Select
'
'    ExecuteCommand = True
'
'    GoTo endHand
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call SaveErrLog
'
'    '------------------------------------------------------------------------------------------------------------------
'endHand:
'
'
'End Function
'
''######################################################################################################################
'
'Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    Dim objControl As CommandBarControl
'    Dim lngLoop As Long
'
'    Select Case Control.ID
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_File_Parameter
'        Call frmCISBorrowPara.ShowEdit(Me, mstrPrivs)
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_NewItem
'
'        Call ExecuteCommand("增加借阅申请")
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_Delete                '删除借阅申请
'
'        If ExecuteCommand("删除借阅申请") Then
'            Call ExecuteCommand("移除借阅申请")
'        End If
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_Audit                '批准借阅申请
'
'        Call ExecuteCommand("批准借阅申请")
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Manage_Refuse                '拒绝借阅申请
'
'        Call ExecuteCommand("拒绝借阅申请")
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_Transf_Cancle                  '恢复数据
'
'        Call ExecuteCommand("恢复数据")
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_Filter '过滤
'
'        If ExecuteCommand("过滤数据") Then
'            Call ExecuteCommand("刷新数据")
'        End If
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_Forward
'        Call ExecuteCommand("前一条")
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_Backward
'        Call ExecuteCommand("后一条")
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_Option
'        mobjFindKey.Execute
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_LocationItem
'
'        mstrFindKey = Control.Parameter
'        mobjFindKey.Caption = mstrFindKey
'        cbsMain.RecalcLayout
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_Location
'
'        LocationObj txtLocation
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case Else
'
'        If Control.ID > 400 And Control.ID < 500 Then
'            Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me, "ID=" & Val(vsf(0).RowData(vsf(0).Row)))
'        Else
'             '与业务无关的功能，公共的功能
'            Call CommandBarExecutePublic(Control, Me, vsf(0), "电子病案借阅申请单")
'        End If
'
'    End Select
'End Sub
'
'Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
'    If stbThis.Visible Then Bottom = stbThis.Height
'End Sub
'
'Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    On Error GoTo errHand
'
'    With vsf(0)
'        Select Case Control.ID
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
'            Control.Enabled = (Val(.RowData(.Row)) > 0)
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_File_Parameter, conMenu_View_Filter, conMenu_View_Refresh
''
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_EditPopup
''            Control.Visible = (tbcTask.Selected.Index = 0)
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Edit_NewItem
''            Control.Visible = IsPrivs(mstrPrivs, "登记申请") And tbcTask.Selected.Index = 0
''
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Edit_Delete
''            Control.Visible = IsPrivs(mstrPrivs, "登记申请") And tbcTask.Selected.Index = 0
''
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Edit_Audit
''
''            Control.Visible = IsPrivs(mstrPrivs, "审批申请") And tbcTask.Selected.Index = 0
''
'
'
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Manage_Refuse
''            Control.Visible = IsPrivs(mstrPrivs, "审批申请") And tbcTask.Selected.Index = 0
''            With vsf(0)
''                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 1
''            End With
''
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Edit_Untread
''            Control.Visible = IsPrivs(mstrPrivs, "审批申请") And tbcTask.Selected.Index = 0
''            With vsf(0)
''                Control.Visible = Control.Visible And (Val(.TextMatrix(.Row, .ColIndex("记录状态"))) > 1)
''                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("记录状态"))) > 1
''                Control.Caption = IIf(Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 2, "回退批准(&B)", "回退拒绝(&B)")
''            End With
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
''            Control.Visible = IsPrivs(mstrPrivs, "登记申请") And tbcTask.Selected.Index = 0
''            Control.Enabled = Control.Visible And DataChanged = True
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_View_Forward
''
''            Select Case tbcTask.Selected.Index
''            Case 0
''                Control.Enabled = (.Row > 1 And DataChanged = False)
''            Case 1
''                Control.Enabled = (mfrmChildPatientView.VsfBody.Row > 1)
''
''            End Select
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_View_Backward
''            Select Case tbcTask.Selected.Index
''            Case 0
''                Control.Enabled = (.Row < .Rows - 1 And DataChanged = False)
''            Case 1
''                Control.Enabled = (mfrmChildPatientView.VsfBody.Row < mfrmChildPatientView.VsfBody.Rows - 1)
''            End Select
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_View_LocationItem        '
''            Control.Checked = (mstrFindKey = Control.Parameter)
''            Select Case tbcTask.Selected.Index
''            Case 0
''                Control.Enabled = (DataChanged = False)
''            Case 1
''
''            End Select
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_View_Location, conMenu_View_Column
''            Select Case tbcTask.Selected.Index
''            Case 0
''                Control.Enabled = (DataChanged = False)
''            Case 1
''
''            End Select
'        '--------------------------------------------------------------------------------------------------------------
'        Case Else
'            Call CommandBarUpdatePublic(Control, Me)
'        End Select
'    End With
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'End Sub
'
'Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
'
'    Select Case Item.ID
'    Case 1
'        Item.Handle = picPane(1).Hwnd
'    Case 2
'        Item.Handle = picPane(0).Hwnd
'    End Select
'
'End Sub
'
'Private Sub Form_Activate()
'    If mblnStartUp = False Then Exit Sub
'    mblnStartUp = False
'    DoEvents
'
'    If ExecuteCommand("初始数据") = False Then GoTo errHand
'
'    Call ExecuteCommand("刷新数据")
'
'    mblnAllowClose = True
'    Exit Sub
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mblnAllowClose = True
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    mblnStartUp = True
'    mblnAllowClose = False
'
'    picPane(1).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
'
'
'    mstrPrivs = UserInfo.模块权限
'    mlngModul = ParamInfo.模块号
'
'    Call ExecuteCommand("初始控件")
'    Call ExecuteCommand("读注册表")
'
'    Call RestoreWinState(Me, App.ProductName)
'    Call zlCommFun.SetWindowsInTaskBar(Me.Hwnd, gblnShowInTaskBar)
'
'End Sub
'
'Private Sub Form_Resize()
'    On Error Resume Next
'
'    Call SetPaneRange(dkpMain, 1, 100, 100, 300, Me.ScaleHeight)
'    dkpMain.RecalcLayout
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'
'    Cancel = Not mblnAllowClose
'
'    If Cancel = False Then
'
'        If Cancel = False Then
'
'            Call ExecuteCommand("写注册表")
'
'            Call SaveWinState(Me, App.ProductName)
'
'            Set mclsVsf(0) = Nothing
'
'            On Error Resume Next
'
'            Unload mfrmGradeEdit
'            Unload mfrmChildMedrec
'
'        End If
'    End If
'
'End Sub
'
''自定义过程或函数
''######################################################################################################################
'
'Private Sub picPane_Resize(Index As Integer)
'    On Error Resume Next
'
'    Select Case Index
'    Case 0
'        tbcTask.Move 0, 0, picPane(Index).Width, picPane(Index).Height
'    Case 1
'        txt科室.Move txt科室.Left, txt科室.Top, picPane(Index).Width - txt科室.Left - 30
'        cmdSelect.Move txt科室.Left + txt科室.Width - cmdSelect.Width - 30, txt科室.Top + 30
'        vsf(0).Move 0, vsf(0).Top, picPane(Index).Width, picPane(Index).Height - vsf(0).Top
'        mclsVsf(0).AppendRows = True
'    End Select
'End Sub
'
'Private Sub tbcTask_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'    mintIndex = Item.Index
'End Sub
'
'Private Sub txtLocation_GotFocus()
'    Call zlControl.TxtSelAll(txtLocation)
'End Sub
'
'Private Sub txtLocation_KeyPress(KeyAscii As Integer)
'    Dim lngRow As Long
'    Dim intCol As Integer
'    Dim bytMatch As Byte
'
'    If KeyAscii = vbKeyReturn Then
'        lngRow = -1
'        bytMatch = 2
'
'        intCol = mclsVsf(0).ColIndex(mstrFindKey)
'        If intCol >= 0 Then
'            lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch, vsf(0).Row + 1)
'            If lngRow = -1 Then
'                lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch)
'            End If
'            If lngRow > 0 And vsf(0).Row <> lngRow Then
'                vsf(0).Row = lngRow
'                vsf(0).ShowCell vsf(0).Row, vsf(0).Col
'            End If
'        End If
'
'        Call LocationObj(txtLocation)
'    End If
'End Sub
'
'Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
'    mclsVsf(Index).AppendRows = True
'End Sub
'
'Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    mclsVsf(Index).AppendRows = True
'End Sub
'
'Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
'    mclsVsf(Index).AppendRows = True
'End Sub
'
'Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim cbrPopupBar As CommandBar
'
'    If Button = 2 And Index = 0 Then
'        Call SendLMouseButton(vsf(Index).Hwnd, x, y)
'
'        Set cbrPopupBar = CopyMenu(cbsMain, 2)
'        If cbrPopupBar Is Nothing Then Exit Sub
'
'        cbrPopupBar.ShowPopup
'    End If
'
'End Sub
'
'
