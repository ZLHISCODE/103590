VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISBorrow 
   Caption         =   "电子病案借阅"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11595
   Icon            =   "frmCISBorrow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11595
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   1
      Left            =   5745
      ScaleHeight     =   3015
      ScaleWidth      =   4500
      TabIndex        =   3
      Top             =   1005
      Width           =   4500
      Begin VB.Frame fra 
         Height          =   630
         Left            =   0
         TabIndex        =   6
         Top             =   -90
         Width           =   4410
         Begin VB.CheckBox chk 
            Caption         =   "&4.已归还"
            Height          =   180
            Index           =   3
            Left            =   3270
            TabIndex        =   10
            Top             =   255
            Value           =   1  'Checked
            Width           =   1050
         End
         Begin VB.CheckBox chk 
            Caption         =   "&3.已拒绝"
            Height          =   180
            Index           =   2
            Left            =   2220
            TabIndex        =   9
            Top             =   255
            Value           =   1  'Checked
            Width           =   1290
         End
         Begin VB.CheckBox chk 
            Caption         =   "&2.已批准"
            Height          =   180
            Index           =   1
            Left            =   1185
            TabIndex        =   8
            Top             =   255
            Value           =   1  'Checked
            Width           =   1110
         End
         Begin VB.CheckBox chk 
            Caption         =   "&1.新申请"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   7
            Top             =   255
            Value           =   1  'Checked
            Width           =   1020
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   540
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   9720
      TabIndex        =   2
      Top             =   105
      Width           =   1125
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   0
      Left            =   285
      ScaleHeight     =   3585
      ScaleWidth      =   4470
      TabIndex        =   1
      Top             =   900
      Width           =   4470
      Begin XtremeSuiteControls.TabControl tbcTask 
         Height          =   1830
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6720
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
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
            Object.Width           =   16563
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
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
      Left            =   -15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmCISBorrow.frx":6852
      Left            =   615
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmCISBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'窗体级变量定义
'######################################################################################################################
Private mstrPrivs As String
Private mblnStartUp As Boolean
Private mblnAllowClose As Boolean
Private mstrCondition As String
Private mstrFindKey As String
Private mlngTmp As Long
Private mobjFindKey As CommandBarControl
Private mclsVsf(0) As clsVsf
Private mlngModul As Long
Private mintIndex As Integer
Private mbytMode As Byte

Private mobjPrintView As CommandBarControl
Private mobjPrintPatient As CommandBarControl
Private mobjPrint As CommandBarControl

Private mrsCondition As New ADODB.Recordset
Private mfrmChildDocumentView As frmChildDocumentView
Private mblnBorrowReason As Boolean
Private mblnBorrowAccount As Boolean

Private WithEvents mfrmCISBorrowEdit As frmCISBorrowEdit
Attribute mfrmCISBorrowEdit.VB_VarHelpID = -1
Private WithEvents mfrmChildPatientView As frmChildPatient
Attribute mfrmChildPatientView.VB_VarHelpID = -1

'######################################################################################################################

Public Property Get 模块号() As Long
    模块号 = mlngModul
End Property

Private Property Let DataChanged(ByVal blnData As Boolean)
    mfrmCISBorrowEdit.DataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmCISBorrowEdit Is Nothing) Then
        DataChanged = mfrmCISBorrowEdit.DataChanged
    End If
End Property

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    
    Call CommandBarInit(cbsMain)
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
    
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '文件
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
    
    Set mobjPrintPatient = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BatPrint, "打印病人所有档案(&B)", True)
    Set mobjPrintView = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BillPrintView, "预览文档(&E)")
    Set mobjPrint = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BillPrint, "打印文档(&T)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "参数设置(&M)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '编辑
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "增加申请(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除申请(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "批准申请(&Y)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Refuse, "拒绝批准(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "回退批准(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Send, "归还接收(&S)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Save, "保存更改(&S)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消更改(&C)")
       
    '查看
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Column, "选择列项(&H)...", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Filter, "过滤(&F)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
    
            
    '帮助
    '------------------------------------------------------------------------------------------------------------------
    Call CreateHelpMenu(cbsMain)
    
    '主菜单右侧的查找
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    mstrFindKey = GetPara("定位依据", mlngModul, "No")
    If mstrFindKey = "" Then mstrFindKey = "No"
    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.No", , , "No")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.申请人", , , "申请人")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.申请时间", , , "申请时间")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&4.申请理由", , , "申请理由")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&5.姓名", , , "姓名")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&6.住院号", , , "住院号")

    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, ""): cbrCustom.Handle = txtLocation.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "前一条"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "后一条"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon
    
    '工具栏定义:包括公共部份
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("标准", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "增加", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Audit, "批准")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Send, "归还")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "保存", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "过滤", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
    
'    Set objControl = NewToolBar(objBar, xtpControlPopup, conMenu_File_Print, "test")
    
    '命令的快键绑定:公共部份主界面已处理
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save           '保存
        .Add 0, vbKeyF12, conMenu_File_Parameter            '参数设置
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add FCONTROL, vbKeyV, conMenu_File_Preview
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '新增
        .Add FCONTROL, vbKeyF, conMenu_View_Filter         '过滤
        .Add 0, vbKeyF3, conMenu_View_Location              '定位
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      '前一条
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '后一条
    End With

End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String

    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        Set mclsVsf(0) = New clsVsf
        With mclsVsf(0)
            Call .Initialize(Me.Controls, vsf(0), True, False, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("记录状态", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            Call .AppendColumn("No", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("申请人", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("申请时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("申请期限", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("申请理由", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("批准人", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("批准时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("借阅时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("借阅期限", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("拒借人", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("拒借时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("拒借理由", 900, flexAlignLeftCenter, flexDTString, "", , True)
            
            .SysHidden(.ColIndex("记录状态")) = True
            
            .AppendRows = True
        End With
        
        '初始菜单及工具栏
        '--------------------------------------------------------------------------------------------------------------
        Call InitCommandBar
        
        '划分停靠区域
        '--------------------------------------------------------------------------------------------------------------
        Dim objPane As Pane
        Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "申请": objPane.Options = PaneNoCaption
        Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, Nothing): objPane.Title = "编辑": objPane.Options = PaneNoCaption
        Set objPane = dkpMain.CreatePane(3, 100, 100, DockRightOf, Nothing): objPane.Title = "查阅": objPane.Options = PaneNoCaption

        dkpMain.SetCommandBars cbsMain
        Call DockPannelInit(dkpMain)


        Call TabControlInit(tbcTask)
        With tbcTask
            .PaintManager.BoldSelected = True
            
            Set mfrmChildPatientView = New frmChildPatient
            Call mfrmChildPatientView.zlInitData(Me, 4, mstrPrivs)
            
            .InsertItem 0, "借阅申请单", picPane(1).hWnd, 1
            If IsPrivs(mstrPrivs, "查阅病案") Then
                .InsertItem 1, "阅读电子病案", mfrmChildPatientView.hWnd, 2
            End If
            .Item(0).Selected = True
        End With
        
        mlngTmp = Val(GetPara("上次状态", 模块号, "0"))
        If mlngTmp >= 0 And mlngTmp <= 1 And tbcTask.ItemCount > mlngTmp Then tbcTask.Item(mlngTmp).Selected = True
        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        
        '创建过滤条件项目，并进行初始化
        Call ParamCreate(mrsCondition)
        Call ParamAdd(mrsCondition, "开始单据号", "")
        Call ParamAdd(mrsCondition, "结束单据号", "")
        Call ParamAdd(mrsCondition, "申请人", "")
        Call ParamAdd(mrsCondition, "批准人", "")
        Call ParamAdd(mrsCondition, "拒绝人", "")
        
        Call ParamAdd(mrsCondition, "新登记单据", "1")
        Call ParamAdd(mrsCondition, "登记开始日期", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "登记结束日期", Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59")
        Call ParamAdd(mrsCondition, "已批准单据", "0")
        Call ParamAdd(mrsCondition, "批准开始日期", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "批准结束日期", Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59")
        Call ParamAdd(mrsCondition, "已拒绝单据", "0")
        Call ParamAdd(mrsCondition, "拒绝开始日期", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "拒绝结束日期", Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59")
        
        Call ParamAdd(mrsCondition, "住院号", "")
        Call ParamAdd(mrsCondition, "病人姓名", "")
        Call ParamAdd(mrsCondition, "申请理由", "")
        
        '读取缺省的借阅申请登记查询时间范围
        strTmp = GetPara("登记缺省范围", mlngModul, "今  天")
        If strTmp = "" Then strTmp = "今  天"
        Call ParamWrite(mrsCondition, "登记开始日期", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "登记结束日期", GetDateTime(strTmp, 2))
        mblnBorrowReason = zlDatabase.GetPara("必须录入借阅原因", ParamInfo.系统号, ParamInfo.模块号, "0", , IsPrivs(mstrPrivs, "参数设置"))
        mblnBorrowAccount = zlDatabase.GetPara("允许自由录入借阅原因", ParamInfo.系统号, ParamInfo.模块号, "0", , IsPrivs(mstrPrivs, "参数设置"))
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
        
        If vsf(0).Enabled <> Not DataChanged Then
            vsf(0).Enabled = Not DataChanged
            vsf(0).ForeColor = IIf(DataChanged, COLOR.深灰色, COLOR.黑色)
            tbcTask.Enabled = Not DataChanged
        End If
        stbThis.Panels(3).Enabled = DataChanged
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新状态"
        
        If mintIndex = 0 Then
            With vsf(0)
                If Val(.RowData(.Row)) = 0 Then
                    strTmp = "当前还没有任何电子病案借阅申请单！"
                Else
                    strTmp = "共有 " & .Rows - 1 & " 个电子病案借阅申请单！"
                End If
            End With
        Else
            With mfrmChildPatientView.VsfBody
                
                If Val(.RowData(.Row)) = 0 Then
                    strTmp = "当前还没有您可以查阅的电子病案！"
                Else
                    strTmp = "当前共有 " & .Rows - 1 & " 个您可以查阅的电子病案！"
                End If
            
            End With
        End If

        stbThis.Panels(2).Text = strTmp
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        
        Call ExecuteCommand("读取申请单据")
        Call ExecuteCommand("读取申请内容")
        Call ExecuteCommand("读取借阅病案")
        Call ExecuteCommand("刷新状态")
        
    '------------------------------------------------------------------------------------------------------------------
    Case "过滤数据"
        
        mrsCondition.Filter = ""
        ExecuteCommand = frmCISBorrowFilter.ShowPara(Me, mrsCondition)

        GoTo endHand
            
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新指定单据"
        
        Set rs = gclsPackage.GetBorrow(1, mlngTmp)
        If rs.BOF = True Then Exit Function
        
        intRow = mclsVsf(0).FindRow(mlngTmp, -1)
        With vsf(0)
            If intRow > 0 Then
                '已加载
                .Row = intRow
            Else
                '未加载
                If Val(.RowData(.Rows - 1)) > 0 Then
                    .Rows = .Rows + 1
                    mclsVsf(0).AppendRows = True
                End If
                .Row = .Rows - 1
            End If
            Call mclsVsf(0).LoadGridRow(.Row, rs)
        End With
        
        Call ExecuteCommand("读取申请内容")
        Call ExecuteCommand("刷新状态")
            
    '------------------------------------------------------------------------------------------------------------------
    Case "读取申请单据"
                
        mclsVsf(0).ClearGrid
        
        Set rs = gclsPackage.GetBorrow(2, 0, ParamRead(mrsCondition, "开始单据号"), _
                                                ParamRead(mrsCondition, "结束单据号"), _
                                                ParamRead(mrsCondition, "申请人"), _
                                                ParamRead(mrsCondition, "批准人"), _
                                                ParamRead(mrsCondition, "拒绝人"), _
                                                IIf(Val(ParamRead(mrsCondition, "新登记单据")) = 1, ParamRead(mrsCondition, "登记开始日期"), ""), _
                                                IIf(Val(ParamRead(mrsCondition, "新登记单据")) = 1, ParamRead(mrsCondition, "登记结束日期"), ""), _
                                                IIf(Val(ParamRead(mrsCondition, "已批准单据")) = 1, ParamRead(mrsCondition, "批准开始日期"), ""), _
                                                IIf(Val(ParamRead(mrsCondition, "已批准单据")) = 1, ParamRead(mrsCondition, "批准结束日期"), ""), _
                                                IIf(Val(ParamRead(mrsCondition, "已拒绝单据")) = 1, ParamRead(mrsCondition, "拒绝开始日期"), ""), _
                                                IIf(Val(ParamRead(mrsCondition, "已拒绝单据")) = 1, ParamRead(mrsCondition, "拒绝结束日期"), ""), _
                                                (chk(0).Value = 1), (chk(1).Value = 1), (chk(2).Value = 1), (chk(3).Value = 1), _
                                                ParamRead(mrsCondition, "住院号"), _
                                                ParamRead(mrsCondition, "病人姓名"), _
                                                ParamRead(mrsCondition, "申请理由"))
        If rs.BOF = False Then
            Call mclsVsf(0).LoadGrid(rs)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取申请内容"
        
        With vsf(0)
            
            If IsPrivs(mstrPrivs, "修改他人申请") Then
                Call mfrmCISBorrowEdit.RefreshData(.RowData(.Row), True, mblnBorrowAccount)
            Else
                Call mfrmCISBorrowEdit.RefreshData(.RowData(.Row), (Trim(.TextMatrix(.Row, .ColIndex("申请人"))) = UserInfo.姓名), mblnBorrowAccount)
            End If
            
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "读取借阅病案"
        
        Call mfrmChildPatientView.zlRefreshData(mrsCondition)
        Call mfrmChildPatientView.zlShowDocument
        
    '------------------------------------------------------------------------------------------------------------------
    Case "增加借阅申请"
        
        mbytMode = 1
        
        With vsf(0)
            If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
            .Row = .Rows - 1
            If .Col = -1 Then .Col = 1
            .ShowCell .Row, .Col
        End With
        
        Call ExecuteCommand("读取申请内容")
        
        Call mfrmCISBorrowEdit.NewData
        
        GoTo endHand
            
    '------------------------------------------------------------------------------------------------------------------
    Case "删除借阅申请"
    
        With vsf(0)
            If Val(.RowData(.Row)) = 0 Then GoTo endHand
            
            If MsgBox("您是否真的要删除申请单号为“" & .TextMatrix(.Row, .ColIndex("No")) & "”的电子病案借阅申请吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                strSQL = "zl_病案借阅记录_Delete(" & Val(.RowData(.Row)) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                If SQLRecordExecute(rsSQL, Me.Caption) Then
                    ExecuteCommand = True
                    Call ExecuteCommand("移除借阅申请")
                    If .Rows = 2 Then
                        Call mfrmCISBorrowEdit.ClearData
                    End If
                End If
            End If
            GoTo endHand
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "批准借阅申请"
        
        mbytMode = 3
        Call mfrmCISBorrowEdit.Aduit
        GoTo endHand
    
    Case "归还接收"
        
        mbytMode = 5
        Call mfrmCISBorrowEdit.Revert
        GoTo endHand
        
    '------------------------------------------------------------------------------------------------------------------
    Case "拒绝借阅申请"
    
        mbytMode = 4
        Call mfrmCISBorrowEdit.Refuse
        GoTo endHand
    '------------------------------------------------------------------------------------------------------------------
    Case "回退批准申请"
        With vsf(0)
            mlngTmp = Val(.RowData(.Row))
            If mlngTmp = 0 Then GoTo endHand

            If MsgBox("您是否真的要回退申请单号为“" & .TextMatrix(.Row, .ColIndex("No")) & "”的电子病案借阅申请的批准吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                strSQL = "zl_病案借阅记录_Rollback(" & Val(.RowData(.Row)) & ",1)"
                Call SQLRecordAdd(rsSQL, strSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "回退拒绝申请"
        With vsf(0)
            mlngTmp = Val(.RowData(.Row))
            If mlngTmp = 0 Then GoTo endHand

            If MsgBox("您是否真的要回退申请单号为“" & .TextMatrix(.Row, .ColIndex("No")) & "”的电子病案借阅申请的拒绝吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                strSQL = "zl_病案借阅记录_Rollback(" & Val(.RowData(.Row)) & ",2)"
                Call SQLRecordAdd(rsSQL, strSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "移除借阅申请"
        
        With vsf(0)
            If .Rows > 2 Then
                .RemoveItem .Row
                mclsVsf(0).AppendRows = True
            Else
                Call mclsVsf(0).ClearGrid
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "恢复数据"
            
        If mfrmCISBorrowEdit.DataChanged Then
            With vsf(0)
                If Val(.RowData(.Row)) = 0 And .Rows > 2 Then
                    .Rows = .Rows - 1
                    .Row = .Rows - 1
                End If
            End With
            Call ExecuteCommand("读取申请内容")
            mfrmCISBorrowEdit.DataChanged = False
        End If
        
        mbytMode = 2
    '------------------------------------------------------------------------------------------------------------------
    Case "校验数据"
    
        '1.
        '--------------------------------------------------
        If mfrmCISBorrowEdit.DataChanged Then
            If mfrmCISBorrowEdit.ValidData(mblnBorrowReason) = False Then GoTo endHand
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "保存数据"
            
        mlngTmp = Val(vsf(0).RowData(vsf(0).Row))
        
        '1.保存详细资料
        '--------------------------------------------------
        If mfrmCISBorrowEdit.DataChanged Then
            If mfrmCISBorrowEdit.SaveData(rsSQL, mlngTmp) = False Then GoTo endHand
        End If
        
        ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        
        GoTo endHand
            
    '------------------------------------------------------------------------------------------------------------------
    Case "前一条"
        With vsf(0)
            If .Row > 1 Then
                .Row = .Row - 1
                .ShowCell .Row, .Col
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "后一条"
        With vsf(0)
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "读注册表"
        
        If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
            '使用个性化设置
            
            mstrFindKey = Trim(GetPara("定位依据", mlngModul, "No"))
            mclsVsf(0).LoadStateFromString Trim(GetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(0)), ""))
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "写注册表"
        If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
            '使用个性化设置
            Call SetPara("定位依据", mstrFindKey, mlngModul)
        End If
        Call SetPara("上次状态", mintIndex, 模块号)
        Call SetRegister(私有模块, Me.Name & "\界面设置\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
        Call SetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(0)), mclsVsf(0).SaveStateToString)
        
    End Select

    ExecuteCommand = True
  
    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    '------------------------------------------------------------------------------------------------------------------
endHand:
    

End Function

'######################################################################################################################

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim lngLoop As Long
    
    Select Case Control.ID
    Case conMenu_File_Parameter
        Call frmCISBorrowPara.ShowEdit(Me, mstrPrivs)

    Case conMenu_File_BillPrintView                    '预览当前文档
        If Not mfrmChildDocumentView Is Nothing Then
        
            Call mfrmChildDocumentView.zlPrintDocument(cbsMain, 1)
            
        End If

    Case conMenu_File_BillPrint                    '打印当前文档
        If Not mfrmChildDocumentView Is Nothing Then
            Call mfrmChildDocumentView.zlPrintDocument(cbsMain, 2)
        End If

    Case conMenu_File_BatPrint
        Dim blnDoctorAdvice As Boolean
        If zlDatabase.GetPara("住院医嘱打印", ParamInfo.系统号, ParamInfo.模块号, "病人医嘱本", , IsPrivs(mstrPrivs, "参数设置")) = "病人医嘱本" Then
            blnDoctorAdvice = False
        Else
            blnDoctorAdvice = True
        End If
        Call frmCISAduitPDF.ShowMe(Me, mfrmChildPatientView.VsfBody, 0, blnDoctorAdvice, False)
        
    Case conMenu_Edit_NewItem
        Call ExecuteCommand("增加借阅申请")

    Case conMenu_Edit_Delete                '删除借阅申请
        Call ExecuteCommand("删除借阅申请")

    Case conMenu_Edit_Audit                      '批准借阅申请
        Call ExecuteCommand("批准借阅申请")
    
    Case conMenu_Edit_Send                      '归还接收
        Call ExecuteCommand("归还接收")
        
    Case conMenu_Manage_Refuse                  '拒绝借阅申请
        Call ExecuteCommand("拒绝借阅申请")
        
    Case conMenu_Edit_Untread                   '回退批准/拒绝
        With vsf(0)
            Select Case Val(.TextMatrix(.Row, .ColIndex("记录状态")))
            Case 2
                If ExecuteCommand("回退批准申请") Then
                    Call ExecuteCommand("刷新指定单据")
                End If
            Case 3
                If ExecuteCommand("回退拒绝申请") Then
                    Call ExecuteCommand("刷新指定单据")
                End If
            End Select
        End With
    Case conMenu_Edit_Transf_Save                  '保存数据
    
        If ExecuteCommand("校验数据") And DataChanged Then
            If ExecuteCommand("保存数据") Then
                
                DataChanged = False
                
                Call ExecuteCommand("刷新指定单据")
                
            End If
        End If
    Case conMenu_Edit_Transf_Cancle                  '恢复数据
        Call ExecuteCommand("恢复数据")
        
    Case conMenu_View_Filter '过滤
        If ExecuteCommand("过滤数据") Then
            Call ExecuteCommand("刷新数据")
        End If

    Case conMenu_View_Column
        If mintIndex = 0 Then
            If frmTemplateColumn.ShowColumn(Me, mclsVsf(0)) Then
                mclsVsf(0).AppendRows = True
            End If
        Else
            Call mfrmChildPatientView.zlColumnSelect
        End If

    Case conMenu_View_Refresh
        Call ExecuteCommand("刷新数据")
        
    Case conMenu_View_Forward
        Call ExecuteCommand("前一条")

    Case conMenu_View_Backward
        Call ExecuteCommand("后一条")

    Case conMenu_View_Option
        mobjFindKey.Execute

    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    Case conMenu_View_Location
        LocationObj txtLocation
        
    Case Else
        If Control.ID > 400 And Control.ID < 500 Then
            Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me, "ID=" & Val(vsf(0).RowData(vsf(0).Row)))
        Else
             '与业务无关的功能，公共的功能
            Call CommandBarExecutePublic(Control, Me, vsf(0), "电子病案借阅申请单")
        End If
        
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHand
    
    With vsf(0)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Enabled = (Val(.RowData(.Row)) > 0) And tbcTask.Selected.Index = 0
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_BillPrintView, conMenu_File_BillPrint, conMenu_File_BatPrint

            Control.Visible = IsPrivs(mstrPrivs, "打印预览文档") And tbcTask.Selected.Index = 1
            Control.Enabled = (Control.Visible And tbcTask.Selected.Index = 1)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Parameter, conMenu_View_Filter, conMenu_View_Refresh, conMenu_View_Column
            Control.Enabled = DataChanged = False
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_EditPopup
            Control.Visible = (tbcTask.Selected.Index = 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem
            Control.Visible = IsPrivs(mstrPrivs, "登记申请") And tbcTask.Selected.Index = 0
            Control.Enabled = Control.Visible And DataChanged = False
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            Control.Visible = IsPrivs(mstrPrivs, "登记申请") And tbcTask.Selected.Index = 0
            With vsf(0)
                
                If IsPrivs(mstrPrivs, "修改他人申请") Then
                    Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 1
                Else
                    Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 1 And Val(.TextMatrix(.Row, .ColIndex("申请人"))) = UserInfo.姓名
                End If
                
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Audit
                        
            Control.Visible = IsPrivs(mstrPrivs, "审批申请") And tbcTask.Selected.Index = 0
            
            With vsf(0)
                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 1
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Send
            
            Control.Visible = IsPrivs(mstrPrivs, "归还接收") And tbcTask.Selected.Index = 0
            
            With vsf(0)
                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 2
            End With
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Refuse
            Control.Visible = IsPrivs(mstrPrivs, "审批申请") And tbcTask.Selected.Index = 0
            With vsf(0)
                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 1
            End With
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Untread
            Control.Visible = IsPrivs(mstrPrivs, "审批申请") And tbcTask.Selected.Index = 0
            With vsf(0)
                Control.Visible = Control.Visible And (Val(.TextMatrix(.Row, .ColIndex("记录状态"))) > 1) And (Val(.TextMatrix(.Row, .ColIndex("记录状态"))) < 4)
                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("记录状态"))) > 1
                Control.Caption = IIf(Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 2, "回退批准(&B)", "回退拒绝(&B)")
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
            
            Control.Visible = (IsPrivs(mstrPrivs, "登记申请") Or IsPrivs(mstrPrivs, "修改他人申请") Or Trim(vsf(0).TextMatrix(.Row, vsf(0).ColIndex("申请人"))) = UserInfo.姓名) And tbcTask.Selected.Index = 0
            
            Control.Enabled = Control.Visible And DataChanged = True
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Forward
            
            Select Case tbcTask.Selected.Index
            Case 0
                Control.Enabled = (.Row > 1 And DataChanged = False)
            Case 1
                Control.Enabled = (mfrmChildPatientView.VsfBody.Row > 1)
                            
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Backward
            Select Case tbcTask.Selected.Index
            Case 0
                Control.Enabled = (.Row < .Rows - 1 And DataChanged = False)
            Case 1
                Control.Enabled = (mfrmChildPatientView.VsfBody.Row < mfrmChildPatientView.VsfBody.Rows - 1)
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationItem        '
            Control.Checked = (mstrFindKey = Control.Parameter)
            Select Case tbcTask.Selected.Index
            Case 0
                Control.Enabled = (DataChanged = False)
            Case 1
                
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Location
            Select Case tbcTask.Selected.Index
            Case 0
                Control.Enabled = (DataChanged = False)
            Case 1
                
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case Else
            Call CommandBarUpdatePublic(Control, Me)
        End Select
    End With
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
End Sub

Private Sub chk_Click(Index As Integer)
        
    Call ExecuteCommand("刷新数据")
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)

    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Set mfrmCISBorrowEdit = New frmCISBorrowEdit
        Call mfrmCISBorrowEdit.InitData(Me, mlngModul, True, mstrPrivs, mblnBorrowAccount)
        Item.Handle = mfrmCISBorrowEdit.hWnd
    Case 3
        Set mfrmChildDocumentView = New frmChildDocumentView
        Call mfrmChildDocumentView.zlInitData(Me)
        Item.Handle = mfrmChildDocumentView.hWnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    DoEvents

    If ExecuteCommand("初始数据") = False Then GoTo errHand
    
    Call ExecuteCommand("刷新数据")
        
    mblnAllowClose = True
    Exit Sub

    '------------------------------------------------------------------------------------------------------------------
errHand:
    mblnAllowClose = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mblnAllowClose = False

    mstrPrivs = UserInfo.模块权限
    mlngModul = ParamInfo.模块号

    Call ExecuteCommand("初始控件")
    Call ExecuteCommand("读注册表")
    
    Call RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 1, 100, 100, 300, Me.ScaleHeight)
    dkpMain.RecalcLayout
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = Not mblnAllowClose
    
    If Cancel = False Then
    
        If DataChanged Then
            Cancel = (MsgBox("修改后的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbNo)
        End If
        
        If Cancel = False Then
        
            Call ExecuteCommand("写注册表")
            
            Call SaveWinState(Me, App.ProductName)
            
            Set mclsVsf(0) = Nothing
            
            On Error Resume Next

            Unload mfrmCISBorrowEdit
            Unload mfrmChildPatientView
            Unload mfrmChildDocumentView
        End If
    End If

End Sub

Private Sub mfrmChildPatientView_AfterDocumentChanged(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strObject As String, ByVal strParam As String, ByVal strCaption As String, ByVal lng提交Id As Long, ByVal blnDataMove As Boolean, ByVal blnScale As Boolean)
    Call mfrmChildDocumentView.zlRefresh(lng病人ID, lng主页ID, strObject, strParam, strCaption, blnDataMove)
    
    mobjPrintView.Caption = "预览""" & mfrmChildPatientView.Title & """(&E)"
    mobjPrint.Caption = "打印""" & mfrmChildPatientView.Title & """(&T)"
    With mfrmChildPatientView.VsfBody
        mobjPrintPatient.Caption = "打印""" & .TextMatrix(.Row, .ColIndex("姓名")) & """的档案(&B)"
    End With
    cbsMain.RecalcLayout
    
End Sub

Private Sub mfrmChildPatientView_DbClick()
'    Dim intRow As Integer
'    Dim strNo As String
'
'
'    With mfrmChildPatientView.VsfBody
'
'        strNo = .TextMatrix(.Row, .ColIndex("No"))
'
'        If strNo <> "" And DataChanged = False Then
'            tbcTask.Item(0).Selected = True
'            With vsf(0)
'                For intRow = 1 To .Rows - 1
'                    If strNo = .TextMatrix(intRow, .ColIndex("No")) Then
'                        .Row = intRow
'                        .ShowCell .Row, .Col
'                        Exit Sub
'                    End If
'                Next
'            End With
'        End If
'
'    End With
    '不在定位 20120326去掉定位功能
    
End Sub

Private Sub mfrmChildPatientView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cbrPopupBar As CommandBar
    
    If Button = 2 Then
        Set cbrPopupBar = CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        
        cbrPopupBar.ShowPopup
    End If
End Sub

Private Sub mfrmCISBorrowEdit_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmCISBorrowEdit_ViewDocument(ByVal strNo As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    
    '切换到查阅病人的电子病案状态
    If strNo <> "" And lng病人ID > 0 And lng主页ID > 0 Then
        tbcTask.Item(1).Selected = True
        Call mfrmChildPatientView.zlLocationPatient(1, , , strNo, lng病人ID, lng主页ID)
    End If
    
End Sub

'自定义过程或函数
'######################################################################################################################

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        tbcTask.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 1
        fra.Move fra.Left, fra.Top, picPane(Index).Width - fra.Left
        vsf(0).Move 0, vsf(0).Top, picPane(Index).Width, picPane(Index).Height - vsf(0).Top
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Private Sub tbcTask_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    mintIndex = Item.Index
    
    Select Case Item.Index
    Case 0
        If dkpMain.Panes(2).Selected = False Then dkpMain.Panes(2).Select
        If dkpMain.Panes(3).Closed = False Then dkpMain.Panes(3).Close
    Case 1
        If dkpMain.Panes(3).Selected = False Then dkpMain.Panes(3).Select
        If dkpMain.Panes(2).Closed = False Then dkpMain.Panes(2).Close
    End Select
    
    Call ExecuteCommand("刷新状态")
End Sub

Private Sub txtLocation_GotFocus()
    Call zlControl.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim intCol As Integer
    Dim bytMatch As Byte
    
    If KeyAscii = vbKeyReturn Then
        If txtLocation.Text = "" Then Exit Sub
        lngRow = -1
        bytMatch = 2
        
        If tbcTask.Item(0).Selected Then
            intCol = mclsVsf(0).ColIndex(mstrFindKey)
            If intCol >= 0 Then
                lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch, vsf(0).Row + 1)
                If lngRow = -1 Then
                    lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch)
                End If
                If lngRow > 0 And vsf(0).Row <> lngRow Then
                    vsf(0).Row = lngRow
                    vsf(0).ShowCell vsf(0).Row, vsf(0).Col
                End If
            End If
        Else
            Call mfrmChildPatientView.zlLocationPatient(2, mstrFindKey, txtLocation.Text)
        End If
        
        Call LocationObj(txtLocation)
    Else
        If InStr(":：;；?？''||", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        
        Call ExecuteCommand("读取申请内容")

    End If
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(Index).BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cbrPopupBar As CommandBar
    
    If Button = 2 And Index = 0 Then
        Call SendLMouseButton(vsf(Index).hWnd, x, y)
        
        Set cbrPopupBar = CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        
        cbrPopupBar.ShowPopup
    End If
    
End Sub

