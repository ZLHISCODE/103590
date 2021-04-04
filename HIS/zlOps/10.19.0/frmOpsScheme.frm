VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~4.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmOpsScheme 
   Caption         =   "手术方案设置"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11295
   Icon            =   "frmOpsScheme.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   6555
      TabIndex        =   5
      ToolTipText     =   "快捷键：F3"
      Top             =   1095
      Width           =   1320
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2715
      Index           =   0
      Left            =   240
      ScaleHeight     =   2715
      ScaleWidth      =   2970
      TabIndex        =   2
      Top             =   1200
      Width           =   2970
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2145
         Index           =   0
         Left            =   45
         TabIndex        =   3
         Top             =   75
         Width           =   2520
         _cx             =   4445
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2865
      Index           =   2
      Left            =   4425
      ScaleHeight     =   2865
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   2370
      Width           =   5265
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   2025
         Left            =   690
         TabIndex        =   1
         Top             =   240
         Width           =   2700
         _Version        =   589884
         _ExtentX        =   4762
         _ExtentY        =   3572
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   6690
      Width           =   11295
      _ExtentX        =   19923
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
            Object.Width           =   14076
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
      Bindings        =   "frmOpsScheme.frx":6852
      Left            =   1170
      Top             =   195
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmOpsScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义
'######################################################################################################################

'常量定义


'变量定义
Private mstrPrivs As String
Private mblnStartUp As Boolean
Private mblnAllowClose As Boolean
Private mclsVsf(0) As New clsVsf
Private mlngTmp As Long
Private mobjFindKey As CommandBarPopup
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnNew As Boolean
Private mlng模块号 As Long
Private WithEvents mfrmChildSchemeEdit As frmChildSchemeEdit
Attribute mfrmChildSchemeEdit.VB_VarHelpID = -1
Private WithEvents mfrmChildSchemeDrug As frmChildSchemeDrug
Attribute mfrmChildSchemeDrug.VB_VarHelpID = -1
Private WithEvents mfrmChildSchemeCharge As frmChildSchemeCharge
Attribute mfrmChildSchemeCharge.VB_VarHelpID = -1
Private WithEvents mfrmChildSchemeMaterial As frmChildSchemeMaterial
Attribute mfrmChildSchemeMaterial.VB_VarHelpID = -1
Private WithEvents mfrmChildSchemeOps As frmChildSchemeOps
Attribute mfrmChildSchemeOps.VB_VarHelpID = -1

'######################################################################################################################

Private Property Let DataChanged(ByVal blnData As Boolean)
    mfrmChildSchemeEdit.DataChanged = blnData
    mfrmChildSchemeDrug.DataChanged = blnData
    mfrmChildSchemeCharge.DataChanged = blnData
    mfrmChildSchemeMaterial.DataChanged = blnData
    mfrmChildSchemeOps.DataChanged = blnData

    If mfrmChildSchemeEdit.DataChanged Or mfrmChildSchemeDrug.DataChanged Or mfrmChildSchemeCharge.DataChanged Or mfrmChildSchemeMaterial.DataChanged Or mfrmChildSchemeOps.DataChanged Then
        stbThis.Panels(3).Enabled = True
    Else
        stbThis.Panels(3).Enabled = False
    End If
End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmChildSchemeEdit Is Nothing) And Not (mfrmChildSchemeDrug Is Nothing) And Not (mfrmChildSchemeCharge Is Nothing) And Not (mfrmChildSchemeMaterial Is Nothing) And Not (mfrmChildSchemeCharge Is Nothing) And Not (mfrmChildSchemeOps Is Nothing) Then
        DataChanged = mfrmChildSchemeEdit.DataChanged Or mfrmChildSchemeDrug.DataChanged Or mfrmChildSchemeCharge.DataChanged Or mfrmChildSchemeMaterial.DataChanged Or mfrmChildSchemeCharge.DataChanged Or mfrmChildSchemeOps.DataChanged
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
    Dim objExtendedBar As CommandBar
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
    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)

    '------------------------------------------------------------------------------------------------------------------
    '编辑
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "增加方案(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_CopyNewItem, "复制增加(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除方案(&D)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Save, "保存更改(&S)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消更改(&R)")
    
    
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")

    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Jump, "窗格跳转(&J)")

    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & ParamInfo.产品名称)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.产品名称 & "主页(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.产品名称 & "论坛(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)
    
    '主菜单右侧的查找
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    
    mstrFindKey = Trim(GetRegister(私有模块, Me.Name, "定位依据", "名称"))
    If mstrFindKey = "" Then mstrFindKey = "名称"

    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.名称", , , "名称")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.编码", , , "编码")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.简码", , , "简码")
    
    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = txtLocation.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "前一条")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon

    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "后一条")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "增加", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "保存", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF6, conMenu_View_Jump                  '跳转
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save                  '保存
        
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '新增
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save            '保存
        
        .Add 0, vbKeyF3, conMenu_View_Location              '定位
        .Add 0, vbKeyF4, conMenu_View_Option                '选择定位依据
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      '前一条
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '后一条
        
    End With
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 200, 100, DockLeftOf, Nothing)
    objPane.Title = "手术方案列表"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 350, 300, DockRightOf, Nothing)
    objPane.Title = "详细资料"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(3, 350, 150, DockBottomOf, objPane)
    objPane.Title = "方案内容"
    objPane.Options = PaneNoCaption
        
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)

End Sub

Private Function InitTabControl() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
        End With

        Set .Icons = frmPubIcons.imgPublic.Icons


        Set mfrmChildSchemeDrug = New frmChildSchemeDrug
        Call mfrmChildSchemeDrug.InitData(Me, IsPrivs(mstrPrivs, "增删改"))

        Set mfrmChildSchemeMaterial = New frmChildSchemeMaterial
        Call mfrmChildSchemeMaterial.InitData(Me, IsPrivs(mstrPrivs, "增删改"))
        
        Set mfrmChildSchemeOps = New frmChildSchemeOps
        Call mfrmChildSchemeOps.InitData(Me, IsPrivs(mstrPrivs, "增删改"))
        
        Set mfrmChildSchemeCharge = New frmChildSchemeCharge
        Call mfrmChildSchemeCharge.InitData(Me, IsPrivs(mstrPrivs, "增删改"))

        .InsertItem 0, "用药方案", mfrmChildSchemeDrug.hWnd, 0
        .InsertItem 1, "材料方案", mfrmChildSchemeMaterial.hWnd, 0
        .InsertItem 2, "治疗方案", mfrmChildSchemeCharge.hWnd, 0
        .InsertItem 3, "适用手术", mfrmChildSchemeOps.hWnd, 0

        .Item(0).Selected = True
        
    End With
    
    InitTabControl = True
    
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
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            Call .AppendColumn("名称", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("编码", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("简码", 750, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("说明", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With

        
        '初始菜单及工具栏
        Call InitCommandBar
        
        '初始窗体分割区域
        Call InitDockPannel
        Call InitTabControl
        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        
        
        
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
        
        If vsf(0).Enabled <> Not DataChanged Then
            vsf(0).Enabled = Not DataChanged
            vsf(0).ForeColor = IIf(DataChanged, COLOR.深灰色, COLOR.黑色)
        End If
        stbThis.Panels(3).Enabled = DataChanged

    '------------------------------------------------------------------------------------------------------------------
    Case "刷新状态"
    
        If Val(vsf(0).RowData(vsf(0).Row)) = 0 Then
            strTmp = "当前还没有定义手术方案！"
        Else
            strTmp = "共定义了 " & vsf(0).Rows - 1 & " 个手术方案！"
        End If

        stbThis.Panels(2).Text = strTmp
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
    
        Call ExecuteCommand("读取手术方案")
        Call ExecuteCommand("读取基本资料")
        Call ExecuteCommand("刷新方案内容")
        Call ExecuteCommand("刷新状态")
            
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新指定数据"

        strSQL = "SELECT '方案' As 图标,A.ID,A.编码,A.名称,A.简码,A.说明 FROM 手术方案参考 A Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngTmp)
        If rs.BOF = True Then Exit Function
                
        intRow = mclsVsf(0).FindRow(mlngTmp, -1)
        If intRow > 0 Then
            '已加载
            vsf(0).Row = intRow
        Else
            '未加载
            If Val(vsf(0).RowData(vsf(0).Rows - 1)) > 0 Then vsf(0).Rows = vsf(0).Rows + 1
            vsf(0).Row = vsf(0).Rows - 1
        End If
        
        Call mclsVsf(0).LoadGridRow(vsf(0).Row, rs)
        Call ExecuteCommand("读取基本资料")
        Call ExecuteCommand("刷新方案内容")
        Call ExecuteCommand("刷新状态")
    
    '------------------------------------------------------------------------------------------------------------------
    Case "读取手术方案"
        
        '清空原有数据
        Call mclsVsf(0).ClearGrid

        '读取现有数据
        strSQL = "SELECT '方案' As 图标,A.ID,A.编码,A.名称,A.简码,A.说明 FROM 手术方案参考 A "
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rs.BOF = False Then Call mclsVsf(0).LoadGrid(rs)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取基本资料"
    
        Call mfrmChildSchemeEdit.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))
            
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新方案内容"
        
        Call ExecuteCommand("读取用药方案")
        Call ExecuteCommand("读取材料方案")
        Call ExecuteCommand("读取适用手术")
        Call ExecuteCommand("读取费用方案")
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取用药方案"
        
        Call mfrmChildSchemeDrug.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))
                    
    '------------------------------------------------------------------------------------------------------------------
    Case "读取材料方案"
        
        Call mfrmChildSchemeMaterial.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取适用手术"
        
        Call mfrmChildSchemeOps.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))
    
    '------------------------------------------------------------------------------------------------------------------
    Case "读取费用方案"
        
        Call mfrmChildSchemeCharge.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))

    '------------------------------------------------------------------------------------------------------------------
    Case "增加手术方案"
        
        mblnNew = True

        If Val(vsf(0).RowData(vsf(0).Rows - 1)) > 0 Then vsf(0).Rows = vsf(0).Rows + 1
        vsf(0).Row = vsf(0).Rows - 1
        vsf(0).ShowCell vsf(0).Row, vsf(0).Col
        
        Call ExecuteCommand("刷新附加数据")

        Call mfrmChildSchemeEdit.NewData(0, mlngTmp)

        Call mfrmChildSchemeDrug.NewData(mlngTmp)
        Call mfrmChildSchemeMaterial.NewData(mlngTmp)
        Call mfrmChildSchemeOps.NewData(mlngTmp)
        Call mfrmChildSchemeCharge.NewData(mlngTmp)
        
        Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "删除手术方案"
        If Val(vsf(0).RowData(vsf(0).Row)) = 0 Then Exit Function
        
        If MsgBox("您是否真的要删除“" & vsf(0).TextMatrix(vsf(0).Row, mclsVsf(0).ColIndex("名称")) & "”手术方案吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
            strSQL = "ZL_手术方案参考_DELETE(" & Val(vsf(0).RowData(vsf(0).Row)) & ")"
            Call SQLRecordAdd(rsSQL, strSQL)
            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End If
        Exit Function

    '------------------------------------------------------------------------------------------------------------------
    Case "移除手术方案"
    
        If vsf(0).Rows > 2 Then
            vsf(0).RemoveItem vsf(0).Row
            mclsVsf(0).AppendRows = True
        Else
            Call mclsVsf(0).ClearGrid
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "恢复数据"
    
        '1.恢复基本资料
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeEdit.DataChanged Then
            If Val(vsf(0).RowData(vsf(0).Row)) = 0 And vsf(0).Rows > 2 Then
                vsf(0).Rows = vsf(0).Rows - 1
                vsf(0).Row = vsf(0).Rows - 1
            End If

            Call ExecuteCommand("读取基本资料")
            mfrmChildSchemeEdit.DataChanged = False
        End If

        '2.恢复用药方案
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeDrug.DataChanged Then
            Call ExecuteCommand("读取用药方案")
            mfrmChildSchemeDrug.DataChanged = False
        End If

        '3.恢复材料方案
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeMaterial.DataChanged Then
            Call ExecuteCommand("读取材料方案")
            mfrmChildSchemeMaterial.DataChanged = False
        End If

        '4.恢复适用手术
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeOps.DataChanged Then
            Call ExecuteCommand("读取适用手术")
            mfrmChildSchemeOps.DataChanged = False
        End If

        '5.恢复费用方案
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeCharge.DataChanged Then
            Call ExecuteCommand("读取费用方案")
            mfrmChildSchemeCharge.DataChanged = False
        End If

        mblnNew = False
    '------------------------------------------------------------------------------------------------------------------
    Case "校验数据"
    
        '1.校验详细资料
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeEdit.DataChanged Then
            If mfrmChildSchemeEdit.ValidData = False Then Exit Function
        End If

        '2.校验用药方案
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeDrug.DataChanged Then
            If mfrmChildSchemeDrug.ValidData = False Then Exit Function
        End If

        '3.校验材料方案
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeMaterial.DataChanged Then
            If mfrmChildSchemeMaterial.ValidData = False Then Exit Function
        End If

        '4.校验适用手术
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeOps.DataChanged Then
            If mfrmChildSchemeOps.ValidData = False Then Exit Function
        End If

        '5.校验费用方案
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeCharge.DataChanged Then
            If mfrmChildSchemeCharge.ValidData = False Then Exit Function
        End If

        ExecuteCommand = True
        
        Exit Function
    '------------------------------------------------------------------------------------------------------------------
    Case "保存数据"
        
        mlngTmp = Val(vsf(0).RowData(vsf(0).Row))

        '1.保存详细资料
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeEdit.DataChanged Then

            If mfrmChildSchemeEdit.SaveData(rsSQL, mlngTmp) = False Then Exit Function

        End If

        '2.保存用药方案
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeDrug.DataChanged Then
            If mfrmChildSchemeDrug.SaveData(rsSQL, mlngTmp) = False Then Exit Function
        End If

        '3.保存材料方案
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeMaterial.DataChanged Then
            If mfrmChildSchemeMaterial.SaveData(rsSQL, mlngTmp) = False Then Exit Function
        End If

        '4.保存适用手术
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeOps.DataChanged Then
            If mfrmChildSchemeOps.SaveData(rsSQL, mlngTmp) = False Then Exit Function
        End If

        '5.保存费用方案
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeCharge.DataChanged Then
            If mfrmChildSchemeCharge.SaveData(rsSQL, mlngTmp) = False Then Exit Function
        End If

        ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)

        Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "前一条"
        If vsf(0).Row > 1 Then
            vsf(0).Row = vsf(0).Row - 1
            vsf(0).ShowCell vsf(0).Row, vsf(0).Col
            Call ExecuteCommand("读取基本资料")
            Call ExecuteCommand("刷新方案内容")
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "后一条"
        If vsf(0).Row < vsf(0).Rows - 1 Then
            vsf(0).Row = vsf(0).Row + 1
            vsf(0).ShowCell vsf(0).Row, vsf(0).Col
            Call ExecuteCommand("读取基本资料")
            Call ExecuteCommand("刷新方案内容")
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读注册表"
        
        If Val(GetRegister(私有全局, "", "使用个性化风格", "0")) = 1 Then
            '使用个性化设置
            
'            dkpMain.LoadStateFromString GetRegister(私有模块, Me.Name & "\界面设置\" & TypeName(dkpMain), dkpMain.Name, "")
            
            mstrFindKey = Trim(GetRegister(私有模块, Me.Name, "定位依据", "名称"))
            mclsVsf(0).LoadStateFromString Trim(GetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(0)), ""))
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "写注册表"
        If Val(GetRegister(私有全局, "", "使用个性化风格", "0")) = 1 Then
            '使用个性化设置
            Call SetRegister(私有模块, Me.Name, "定位依据", mstrFindKey)
        End If
        Call SetRegister(私有模块, Me.Name & "\界面设置\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
        Call SetRegister(私有模块, Me.Name, "表格参数_" & TypeName(vsf(0)), mclsVsf(0).SaveStateToString)
    End Select

    ExecuteCommand = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

'######################################################################################################################

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem               '增加手术方案
        mlngTmp = 0
        Call ExecuteCommand("增加手术方案")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_CopyNewItem           '复制增加手术方案
    
        mlngTmp = Val(vsf(0).RowData(vsf(0).Row))
        Call ExecuteCommand("增加手术方案")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                '删除手术方案

        If ExecuteCommand("删除手术方案") Then
            Call ExecuteCommand("移除手术方案")
            Call ExecuteCommand("读取基本资料")
            Call ExecuteCommand("刷新方案内容")
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save                  '保存手术方案
    
        If ExecuteCommand("校验数据") And DataChanged Then
            If ExecuteCommand("保存数据") Then
                DataChanged = False
                Call ExecuteCommand("刷新指定数据")
                mblnNew = False
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle                  '恢复手术方案
    
        Call ExecuteCommand("恢复数据")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Jump
        
        If tbcPage.Selected.Index + 1 <= tbcPage.ItemCount - 1 Then
            tbcPage.Item(tbcPage.Selected.Index + 1).Selected = True
        Else
            tbcPage.Item(0).Selected = True
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward
        Call ExecuteCommand("前一条")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward
        Call ExecuteCommand("后一条")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        
        mobjFindKey.Execute
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
    
        LocationObj txtLocation
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               '刷新

        Call ExecuteCommand("刷新数据")
        
    '------------------------------------------------------------------------------------------------------------------
    Case Else

        If Control.ID > 400 And Control.ID < 500 Then
            Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
        Else
             '与业务无关的功能，公共的功能
            Call CommandBarExecutePublic(Control, Me, vsf(0), "手术方案清单")
        End If

    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    On Error GoTo errHand

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel
    
        Control.Enabled = (Val(vsf(0).RowData(vsf(0).Row)) > 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup
        
        Control.Visible = IsPrivs(mstrPrivs, "增删改")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem, conMenu_Edit_CopyNewItem
    
        Control.Visible = IsPrivs(mstrPrivs, "增删改")
        Control.Enabled = (DataChanged = False And Control.Visible)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                   '修改,删除
    
        Control.Visible = IsPrivs(mstrPrivs, "增删改")
        Control.Enabled = (Val(vsf(0).RowData(vsf(0).Row)) > 0 And DataChanged = False And Control.Visible)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
        
        Control.Visible = IsPrivs(mstrPrivs, "增删改")
        Control.Enabled = (DataChanged And Control.Visible)
                
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward
        Control.Enabled = (vsf(0).Row > 1 And DataChanged = False)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward
        Control.Enabled = (vsf(0).Row < vsf(0).Rows - 1 And DataChanged = False)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem        '
        Control.Checked = (mstrFindKey = Control.Parameter)
        Control.Enabled = (DataChanged = False)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
         Control.Enabled = (DataChanged = False)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Find, conMenu_View_Refresh
        Control.Enabled = (DataChanged = False And Control.Visible)
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        Call CommandBarUpdatePublic(Control, Me)
    End Select
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

End Sub

Private Sub cbsSearch_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsMain_Execute(Control)
End Sub

Private Sub cbsSearch_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsMain_Update(Control)
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Set mfrmChildSchemeEdit = New frmChildSchemeEdit
        Item.Handle = mfrmChildSchemeEdit.hWnd
        Call mfrmChildSchemeEdit.InitData(Me, IsPrivs(mstrPrivs, "增删改"))
    Case 3
        Item.Handle = picPane(2).hWnd
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

errHand:
    mblnAllowClose = True
    Unload Me
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    mblnAllowClose = False
    
    mstrPrivs = UserInfo.模块权限
    mlng模块号 = ParamInfo.模块号
    
    Call ExecuteCommand("初始控件")
    Call ExecuteCommand("读注册表")

    Call RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDatabase.ShowReportMenu(Me, ParamInfo.系统号, ParamInfo.模块号, UserInfo.模块权限)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Call SetPaneRange(dkpMain, 1, 100, 60, 200, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 2, 15, 150, Me.ScaleWidth, 150)

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
            
            Unload mfrmChildSchemeEdit
            Unload mfrmChildSchemeDrug
            Unload mfrmChildSchemeMaterial
            Unload mfrmChildSchemeOps
            Unload mfrmChildSchemeCharge
        End If
    End If

End Sub

Private Sub mfrmChildSchemeCharge_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmChildSchemeDrug_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmChildSchemeEdit_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmChildSchemeMaterial_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmChildSchemeOps_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(0).AppendRows = True
    Case 2
        tbcPage.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    End Select
End Sub

Private Sub txtLocation_GotFocus()
    Call zlControl.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim intCol As Integer
    Dim bytMatch As Byte
    
    If KeyAscii = vbKeyReturn Then
    
        lngRow = -1
        bytMatch = 0
        Select Case mstrFindKey
        Case "名称"
            bytMatch = 2
            intCol = mclsVsf(0).ColIndex("名称")
        Case "简码"
            bytMatch = 2
            intCol = mclsVsf(0).ColIndex("简码")
        Case "编码"
            bytMatch = 2
            intCol = mclsVsf(0).ColIndex("编码")
        Case Else
            Exit Sub
        End Select
        
        lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch, vsf(0).Row + 1)
        If lngRow = -1 Then
            lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch)
        End If
        If lngRow > 0 And vsf(0).Row <> lngRow Then
            vsf(0).Row = lngRow
            vsf(0).ShowCell vsf(0).Row, vsf(0).Col
        End If
        
        Call LocationObj(txtLocation)
    End If
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Call mclsVsf(Index).AfterMoveColumn(Col, Position)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Index = 0 Then
        If OldRow = NewRow Then Exit Sub
        Call mclsVsf(Index).SelectRow(OldRow, NewRow)
        Call ExecuteCommand("读取基本资料")
        Call ExecuteCommand("刷新方案内容")
    End If
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Call mclsVsf(Index).RestoreRow(mclsVsf(Index).SaveKey)
    vsf(Index).ShowCell vsf(Index).Row, vsf(Index).Col
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    mclsVsf(Index).SaveKey = Val(vsf(Index).RowData(vsf(Index).Row))
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col = mclsVsf(Index).ColIndex("图标"))
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)
        Select Case Index
        Case 0
            If mclsVsf(Index).MoveColumn = False Then
                
                Set cbrPopupBar = CopyMenu(cbsMain, 2)
                If cbrPopupBar Is Nothing Then Exit Sub
                cbrPopupBar.ShowPopup
            End If
        End Select
        
    End Select
End Sub


