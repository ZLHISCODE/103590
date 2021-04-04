VERSION 5.00
Begin VB.Form frmGradeStandard 
   Caption         =   "评分标准维护"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11295
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "frmGradeStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''///////////////////////////////////////////////////////////////////////////////
''
''       模块：评分标准维护
''       功能：病案评分标准的录入、修改、删除、打印、选用等。
''       编写：吴庆伟
''       日期：2005年1月5日
''
''///////////////////////////////////////////////////////////////////////////////
'
'
'Option Explicit
'
'
'Private mstrPrivs As String
'Private mblnStartUp As Boolean
'Private mblnAllowClose As Boolean
'Private mlngModul As Long
'
'Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
'
'Private m_lngOldRow As Long
'Private m_lngCurRow As Long
'Private m_lngCurID As Long
'Private m_lngCurFAID As Long
'Private m_lngCurSJID As Long     '记录当前记录ID,方案ID,上级ID
'Private m_strTreeKey As String
'Private m_lngOldSJID As Long
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
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
'
'    '编辑
'    '------------------------------------------------------------------------------------------------------------------
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
'    objMenu.ID = conMenu_EditPopup
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewKind, "增加方案(&N)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyKind, "修改方案(&F)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteKind, "删除方案(&L)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Import, "导入方案(&P)...")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Select, "选用方案(&S)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewParent, "增加项目(&X)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Insert, "插入项目(&R)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyParent, "修改项目(&G)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteParent, "删除项目(&C)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "增加标准(&A)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Append, "插入标准(&I)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "修改标准(&M)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除标准(&D)")
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
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
'
'
'    '帮助
'    '------------------------------------------------------------------------------------------------------------------
'    Call CreateHelpMenu(cbsMain)
'
'    '工具栏定义:包括公共部份
'    '------------------------------------------------------------------------------------------------------------------
'    Set objBar = cbsMain.Add("标准", xtpBarTop)
'    objBar.ContextMenuPresent = False
'    objBar.ShowTextBelowIcons = False
'    objBar.EnableDocking xtpFlagStretched
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
'
'    Set objPopup = NewToolBar(objBar, xtpControlPopup, conMenu_Edit_NewKind * 10# + 1, "方案", True, , , , objControl.Index + 1)
'    objPopup.ID = conMenu_Edit_NewKind
'    objPopup.IconId = conMenu_Edit_NewParent
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_NewKind, "增加方案(&A)")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_ModifyKind, "修改方案(&D)")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_DeleteKind, "删除方案(&D)")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Select, "选用方案(&S)")
'
'    Set objPopup = NewToolBar(objBar, xtpControlPopup, conMenu_Edit_NewParent * 10# + 1, "项目", True, , , , objControl.Index + 1)
'    objPopup.ID = conMenu_Edit_NewParent
'    objPopup.IconId = conMenu_Edit_NewParent
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_NewParent, "增加项目(&A)")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改项目(&D)")
'    Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_DeleteParent, "删除项目(&D)")
'
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "增加", True)
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
'
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")
'
'    '命令的快键绑定:公共部份主界面已处理
'    '------------------------------------------------------------------------------------------------------------------
'    With cbsMain.KeyBindings
'        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
'        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
'        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
'    End With
'
'End Function
'
'
'Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim intLoop As Integer
'    Dim intRow As Integer
'    Dim Rs As New ADODB.Recordset
'    Dim rsSQL As New ADODB.Recordset
'    Dim strTmp As String
'    Dim strSql As String
'
'    On Error GoTo errHand
'
'    Call SQLRecord(rsSQL)
'
'    Select Case strCommand
'    '------------------------------------------------------------------------------------------------------------------
'    Case "初始控件"
'
'
'        '初始菜单及工具栏
'        '--------------------------------------------------------------------------------------------------------------
'        Call InitCommandBar
'
'        '划分停靠区域
'        '--------------------------------------------------------------------------------------------------------------
'        Dim objPane As Pane
'        Set objPane = dkpMain.CreatePane(1, 100, 200, DockLeftOf, Nothing): objPane.Title = "评分方案": objPane.Options = PaneNoCaption
'        Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, Nothing): objPane.Title = "评分标准": objPane.Options = PaneNoCaption
'        Set objPane = dkpMain.CreatePane(3, 100, 100, DockBottomOf, objPane): objPane.Title = "项目信息": objPane.Options = PaneNoCaption
'
'        dkpMain.SetCommandBars cbsMain
'        Call DockPannelInit(dkpMain)
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "初始数据"
'
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "控件状态"
'
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "刷新状态"
'
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "刷新数据"
'
'        '填充Tree
'        Call FillTree
'
'        '填充列表
'        Call Fill结果
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "读注册表"
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "写注册表"
'
'    End Select
'
'    ExecuteCommand = True
'
'    Exit Function
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call SaveErrLog
'
'End Function
'
'Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    Dim objControl As CommandBarControl
'    Dim lngLoop As Long
'
'    Select Case Control.ID
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_File_Parameter
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case Else
'
'        If Control.ID > 400 And Control.ID < 500 Then
'            Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me, "评分方案=" & m_lngCurFAID)
'        Else
'             '与业务无关的功能，公共的功能
'            Call CommandBarExecutePublic(Control, Me, fgMain, "评分方案标准内容")
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
'    With fgMain
'        Select Case Control.ID
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel
'            Control.Enabled = .Row > 0
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_EditPopup, conMenu_Edit_NewKind, conMenu_Edit_NewParent, conMenu_Edit_Insert, conMenu_Edit_NewItem, conMenu_Edit_Append
'            Control.Visible = IsPrivs(mstrPrivs, "增删改")
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_Edit_ModifyKind, conMenu_Edit_DeleteKind, conMenu_Edit_Select
'            Control.Visible = IsPrivs(mstrPrivs, "增删改")
'            Control.Enabled = Control.Visible And Not (tvw方案.SelectedItem Is Nothing)
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_Edit_Import
'            Control.Visible = IsPrivs(mstrPrivs, "增删改")
'            If tvw方案.SelectedItem Is Nothing Then
'                Control.Enabled = False
'            Else
'                Control.Enabled = Control.Visible And tvw方案.Nodes.Count > 1
'            End If
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_Edit_ModifyParent, conMenu_Edit_DeleteParent
'            Control.Visible = IsPrivs(mstrPrivs, "增删改")
'            Control.Enabled = Control.Visible And .Row > 0
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_Edit_Modify, conMenu_Edit_Delete
'            Control.Visible = IsPrivs(mstrPrivs, "增删改")
'            Control.Enabled = Control.Visible And .Row > 0
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
'    Select Case Item.ID
'    Case 1
'        Item.Handle = picPane(0).Hwnd
'    Case 2
'        Item.Handle = picPane(1).Hwnd
'    Case 3
'        Item.Handle = picPane(2).Hwnd
'    End Select
'End Sub
'
'Private Sub fgMain_Click()
'    fgMain_SelChange
'End Sub
'
'Private Sub fgMain_DblClick()
'    If fgMain.MouseRow = 0 Then Exit Sub
'    Call fgMain_KeyPress(13)
'End Sub
'
'Private Sub fgMain_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'
'        If IsPrivs(mstrPrivs, "增删改") Then
'            If fgMain.Row > 0 And fgMain.TextMatrix(fgMain.Row, 2) <> "" Then
'                Call mnuEditModBZ_Click
'            End If
'        End If
'    End If
'End Sub
'
''Private Sub fgMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    If InStr(gstrPrivs, "增删改") = 0 Then Exit Sub
''    If Button = vbRightButton Then
''        If fgMain.MouseRow = -1 And fgMain.Rows >= 1 Then
''            fgMain.Row = fgMain.Rows - 1
''        ElseIf fgMain.MouseRow = 0 And fgMain.Rows > 1 Then
''            fgMain.Row = 1
''        Else
''            fgMain.Row = fgMain.MouseRow
''        End If
''        fgMain.Col = fgMain.MouseCol
''
''        m_lngCurSJID = IIf(Len(fgMain.Cell(flexcpText, fgMain.Row, 5)) = 0, 0, Val(fgMain.Cell(flexcpText, fgMain.Row, 5)))      '获取ID
''
''        PopupMenu mnuShortEdit
''    End If
''End Sub
'
'Private Sub fgMain_SelChange()
'    Dim lngID As Long
'    m_lngCurRow = fgMain.Row
'    If m_lngCurRow < 0 Then m_lngCurSJID = 0: m_lngCurID = 0: Exit Sub
'    m_lngCurID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 4)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 4)))    '获取ID
'    m_lngCurSJID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 5)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 5)))     '获取ID
'    m_lngCurFAID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 6)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 6)))     '获取ID
'    If m_lngCurSJID = 0 Then
'        lngID = m_lngCurID
'    Else
'        lngID = m_lngCurSJID
'    End If
'
'    Show基本要求 lngID, fgMain.Cell(flexcpText, m_lngCurRow, 0), fgMain.Cell(flexcpText, m_lngCurRow, 1)
'    m_lngOldRow = m_lngCurRow
'    SetMenu
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
'Private Sub Form_Initialize()
'    Call InitCommonControls
'End Sub
'
'Private Sub Form_Load()
'
'    mblnStartUp = True
'    mblnAllowClose = False
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
'
'    Me.KeyPreview = True
'    m_lngOldRow = -1
'    m_lngCurRow = -1
'    m_lngCurID = -1
'    m_lngOldSJID = -1
'
'    '权限控制
''    Call 权限控制
''    '填充Tree
''    Call FillTree
''
''    '填充列表
''    Call Fill结果
'
'    '恢复界面位置
'
''    RestoreWinState Me, App.ProductName
''    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
''    Call SetMenu
'
'    picFAXX.Picture = imgClose.Picture
'End Sub
'
'Private Sub Form_Resize()
'    On Error Resume Next
'
'    Call SetPaneRange(dkpMain, 1, 100, 100, 250, Me.ScaleHeight)
'    Call SetPaneRange(dkpMain, 3, 100, 100, Me.ScaleWidth, 200)
'    dkpMain.RecalcLayout
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    m_strTreeKey = ""
'    SaveWinState Me, App.ProductName
'End Sub
'
'Private Sub picFAXX_Click()
'    If picFAXX.Tag = "" Then
'        picFAXX.Tag = "Opened"
'        picFAXX.Picture = imgOpen.Picture
'        pic方案信息.Height = 340
'    Else
'        picFAXX.Tag = ""
'        picFAXX.Picture = imgClose.Picture
'        pic方案信息.Height = 1695
'    End If
'    picFAXX.Refresh
'    Call picPane_Resize(0)
'End Sub
'
'Private Sub mnuEditDelBZ_Click()
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'
'    On Error GoTo errHandle
'
'    Dim intIndex As Long
'
'    If m_lngCurID < 1 Then Exit Sub
'    If MsgBox("你确认要删除该条评分标准吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'    gstrSQL = "ZL_病案评分标准_Delete(" & CStr(m_lngCurID) & ",1)"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'    Call Fill结果
'    Call SetMenu
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub
'
'Private Sub mnuEditDelFA_Click()
'    '删除评分方案
'    On Error GoTo errHandle
'    Dim intIndex As Long
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'
'    If m_lngCurFAID < 1 Then Exit Sub
'
'    If MsgBox("你确认要删除该条方案吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'
'    gstrSQL = "ZL_病案评分方案_Delete(" & CStr(m_lngCurFAID) & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'    Call FillTree
'    Call Fill结果
'    Call SetMenu
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub
'
'Private Sub mnuEditDelXM_Click()
'     m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'
'    On Error GoTo errHandle
'    Dim intIndex As Long
'
'    If m_lngCurID < 1 Then Exit Sub
'
'    If MsgBox("你确认要删除该条评分项目吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'
'    If m_lngCurSJID = 0 Then
'        gstrSQL = "ZL_病案评分标准_Delete(" & CStr(m_lngCurID) & ",0)"
'    Else
'        gstrSQL = "ZL_病案评分标准_Delete(" & CStr(m_lngCurSJID) & ",0)"
'    End If
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'    Call Fill结果
'    Call SetMenu
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub
'
'Private Sub mnuEditEmportFA_Click()
'    '导入已有方案
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'    If m_lngCurFAID <= 0 Then Exit Sub
'
'    Dim lID As Long     '选中的方案ID
'    Dim lNewID As Long
'
'    Dim f As New frm选择评分方案
'    f.FillCmbSelFA m_lngCurFAID
'    f.Show 1
'    lID = f.ID_From
'
'    '执行导入操作！！！
'    '源ID为：lID   目的ID为： m_lngCurFAID
'    Dim Rs As New ADODB.Recordset, lng总分 As Double, rsTmp As New ADODB.Recordset
'    Dim strT As String
'    gstrSQL = "select * from 病案评分标准 where 上级ID is null and 方案ID=" & lID & " order by 上级序号,序号,ID"
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    zlCommFun.ShowFlash "请稍候，系统正在导入评分方案……", Me
'    DoEvents
'    On Error GoTo LL
'    gcnOracle.BeginTrans
'    Do While Not Rs.EOF
'        '找到了项目，添加项目
'        lNewID = zlDatabase.GetNextId("病案评分标准")
'        gstrSQL = "ZL_病案评分标准_Insert" & _
'            "(" & lNewID & _
'            "," & NVL(Rs("上级ID"), "NULL") & _
'            "," & m_lngCurFAID & _
'            ",'" & NVL(Rs("名称")) & _
'            "','" & NVL(Rs("描述")) & _
'            "'," & NVL(Rs("标准分值"), "NULL") & _
'            ",'" & NVL(Rs("缺陷等级")) & _
'            "','" & NVL(Rs("评分单位")) & "',0)"
'        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'        '进一步查找下级项目，循环添加之！
'        gstrSQL = "select * from 病案评分标准 where 上级ID=" & Rs("ID") & " and 方案ID=" & lID & " order by 上级序号,序号,ID"
'        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
'        Do While Not rsTmp.EOF
'            gstrSQL = "ZL_病案评分标准_Insert" & _
'                "(" & zlDatabase.GetNextId("病案评分标准") & _
'                "," & lNewID & _
'                "," & m_lngCurFAID & _
'                ",'" & NVL(rsTmp("名称")) & _
'                "','" & NVL(rsTmp("描述")) & _
'                "'," & NVL(rsTmp("标准分值"), "NULL") & _
'                ",'" & NVL(rsTmp("缺陷等级")) & _
'                "','" & NVL(rsTmp("评分单位")) & "',0)"
'            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'            rsTmp.MoveNext
'        Loop
'        Rs.MoveNext
'    Loop
'
'    '刷新结果！
'    gcnOracle.CommitTrans
'
'    Call Fill结果
'    zlCommFun.StopFlash
'    Exit Sub
'LL:
'    gcnOracle.RollbackTrans
'    zlCommFun.StopFlash
'End Sub
'
'Private Sub mnuEditInsBZ_Click()
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'    Dim f As New frm评分标准编辑
'    If m_lngCurSJID < 1 Then '为独立评分项
'        f.ShowForm "插入", m_lngCurFAID, m_lngCurID, m_lngCurSJID
'    Else
'        f.ShowForm "插入", m_lngCurFAID, m_lngCurSJID, m_lngCurID
'    End If
'
'    Call 刷新方案信息
'    If f.Moded Then
'        Call Fill结果
'    End If
'End Sub
'
'Private Sub mnuEditInsXM_Click()
'    '新增评分项目
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'    Dim f As New frm评分标准编辑
'    If m_lngCurSJID < 1 Then '为独立评分项
'        f.ShowForm "插入", m_lngCurFAID, 0, m_lngCurID
'    Else
'        f.ShowForm "插入", m_lngCurFAID, 0, m_lngCurSJID
'    End If
'    Call 刷新方案信息
'    If f.Moded Then
'        Call Fill结果
'    End If
'End Sub
'
'Private Sub mnuEditModBZ_Click()
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'    '修改评分标准
'    If m_lngCurID < 1 Then Exit Sub
'    Dim f As New frm评分标准编辑
'    If fgMain.Col < 2 Then  '一级项目
'        If m_lngCurSJID < 1 Then
'            f.ShowForm "修改", m_lngCurFAID, , m_lngCurID
'        Else
'            f.ShowForm "修改", m_lngCurFAID, , m_lngCurSJID
'        End If
'    Else                    '子项目
'        f.ShowForm "修改", m_lngCurFAID, m_lngCurSJID, m_lngCurID
'    End If
'    Call 刷新方案信息
'    If f.Moded Then
'        Call Fill结果
'    End If
'End Sub
'
'Private Sub mnuEditModFA_Click()
'    '修改评分方案
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'    If m_lngCurFAID < 1 Then Exit Sub
'    Dim f As New frm评分方案编辑, lng总分 As Double
'    f.ShowForm m_lngCurFAID   '修改，传入ID
'    Call 刷新方案信息
'    If f.Moded Then
'        Call FillTree
'        '填充列表
'        Call Fill结果
'    End If
'End Sub
'
'Private Sub mnuEditModXM_Click()
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'    If m_lngCurID < 1 Then Exit Sub
'    Dim f As New frm评分标准编辑
'    If m_lngCurSJID < 1 Then
'        f.ShowForm "修改", m_lngCurFAID, , m_lngCurID
'    Else
'        f.ShowForm "修改", m_lngCurFAID, , m_lngCurSJID
'    End If
'    Call 刷新方案信息
'    If f.Moded Then
'        Call Fill结果
'    End If
'End Sub
'
'Private Sub mnuEditNewBZ_Click()
'    '新增下级评分标准
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'    Dim f As New frm评分标准编辑
'    If m_lngCurSJID < 1 Then '为独立评分项
'        f.ShowForm "新增", m_lngCurFAID, m_lngCurID
'    Else
'        f.ShowForm "新增", m_lngCurFAID, m_lngCurSJID
'    End If
'
'    Call 刷新方案信息
'    If f.Moded Then
'        Call Fill结果
'    End If
'
'End Sub
'
'Private Sub mnuEditNewFA_Click()
'    '增加评分方案
'    Dim f As New frm评分方案编辑
'    f.ShowForm   '新增
'    Call 刷新方案信息
'    If f.Moded Then
'        Call FillTree
'        '填充列表
'        Call Fill结果
'    End If
'End Sub
'
'Private Sub 刷新方案信息()
'    Dim Rs As New ADODB.Recordset, lng总分 As Double
'    gstrSQL = "select * from 病案评分方案 where ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    If Not Rs.EOF Then
'        lbl方案名称 = Rs("名称")
'        lbl分制 = "分制:" & Rs("分制")
'        lbl上值 = "上值:" & Rs("上值")
'        lbl下值 = "下值:" & Rs("下值")
'        lbl总分 = "总分:" & Rs("总分")
'        lng总分 = Rs("总分")
'    Else
'        lbl方案名称 = ""
'        lbl分制 = ""
'        lbl上值 = ""
'        lbl下值 = ""
'        lbl总分 = ""
'    End If
'
'    Rs.Close
'    gstrSQL = "select sum(标准分值) from 病案评分标准 where 上级ID is null and 方案ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    If Not Rs.EOF Then
'        If Abs(lng总分 - Rs.Fields(0)) > 0.01 Then
'            lbl总分 = lbl总分 + "，项目分数和为:" & Rs.Fields(0)
'            lbl总分.ForeColor = vbRed
'        Else
'            lbl总分.ForeColor = vbBlack
'        End If
'    Else
'        lbl总分.ForeColor = vbRed
'    End If
'End Sub
'
'Private Sub mnuEditNewXM_Click()
'    '新增评分项目
'    m_lngCurFAID = Mid(tvw方案.SelectedItem.Key, 2)
'    Dim f As New frm评分标准编辑
'    f.ShowForm "新增", m_lngCurFAID
'    Call 刷新方案信息
'    If f.Moded Then
'        Call Fill结果
'    End If
'End Sub
'
'Private Sub mnuEditSelFA_Click()
'    On Error GoTo errHandle
'    Dim intIndex As Long, bln已使用 As Boolean
'
'    If m_lngCurFAID < 1 Then Exit Sub
'    If MsgBox("注意：评分分案的选用是一件非常慎重的事情，通常不要随意更改！" & vbCrLf & "请确认选用本评分方案吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'
'    Dim rsTemp As New ADODB.Recordset
'    gstrSQL = "select count(*) from 病案评分结果 where 方案ID=(select ID from 病案评分方案 where 类型='住院' and 选用=1)"
'    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
'    If rsTemp(0).Value > 0 Then
'        '默认住院方案已经使用
'        If MsgBox("注意：系统默认评分分案正在使用当中，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'    End If
'    rsTemp.Close
'
'    gstrSQL = "ZL_病案评分方案_选用(" & CStr(m_lngCurFAID) & ",1)"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'    Call FillTree
'    Call SetMenu
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub
'
'Private Sub mnuFileSetup_Click()
'    frm评分标准参数设置.Show 1
'End Sub
'
'Private Sub mnuHelpAbout_Click()
'    '关于对话框
'    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
'End Sub
'
'Private Sub mnuShortMenuBZ_Click(Index As Integer)
'    '弹出菜单处理
'    Select Case Index
'        Case 1
'            mnuEditNewBZ_Click
'        Case 2
'            mnuEditInsBZ_Click
'        Case 3
'            mnuEditModBZ_Click
'        Case 4
'            mnuEditDelBZ_Click
'    End Select
'End Sub
'
'Private Sub mnuShortMenuFA_Click(Index As Integer)
'    '弹出菜单处理
'    Select Case Index
'        Case 1
'            mnuEditNewFA_Click
'        Case 2
'            mnuEditModFA_Click
'        Case 3
'            mnuEditDelFA_Click
'        Case 4
'            mnuEditSelFA_Click
'        Case 5
'            mnuEditEmportFA_Click
'    End Select
'End Sub
'
'Private Sub mnuShortMenuXM_Click(Index As Integer)
'    Select Case Index
'        Case 1
'            mnuEditNewXM_Click
'        Case 2
'            mnuEditInsXM_Click
'        Case 3
'            mnuEditModXM_Click
'        Case 4
'            mnuEditDelXM_Click
'    End Select
'End Sub
'
'Private Sub mnuShortMnuXM_Click(Index As Integer)
'    Select Case Index
'        Case 1
'            mnuEditNewXM_Click
'        Case 2
'            mnuEditInsXM_Click
'        Case 3
'            mnuEditModXM_Click
'        Case 4
'            mnuEditDelXM_Click
'    End Select
'End Sub
'
'Private Sub mnuViewRefresh_Click()
'    '刷新TreeView
'    Call FillTree
'End Sub
'
'Private Sub mnuFileExit_Click()
'    '关闭窗体
'    Unload Me
'End Sub
'
'Private Sub mnuFileExcel_Click()
'    '输出到Excel
'    '1 打印;2 预览;3 输出到EXCEL
'    subPrint 3
'End Sub
'
'Private Sub mnufilepre_Click()
'    '预览
'    '1 打印;2 预览;3 输出到EXCEL
'    subPrint 2
'End Sub
'
'Private Sub mnuFilePrint_Click()
'    '打印
'    '1 打印;2 预览;3 输出到EXCEL
'    subPrint 1
'End Sub
'
'Private Sub mnufileset_Click()
'    '打印设置 （zlPrintMethod）
'    zlPrintSet
'End Sub
'
'Private Sub picFAXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If X >= 0 And X <= picFAXX.ScaleWidth And Y >= 0 And Y <= picFAXX.ScaleHeight Then
'        SetCapture picFAXX.Hwnd
'        '鼠标移入！！！
'        picFAXX.Line (0, 0)-(picFAXX.ScaleWidth - Screen.TwipsPerPixelX, picFAXX.ScaleHeight - Screen.TwipsPerPixelY), vbBlue, B
'    Else
'        '鼠标移出！！！
'        picFAXX.Cls
'        ReleaseCapture
'    End If
'End Sub
'
'Private Sub picPane_Resize(Index As Integer)
'    Select Case Index
'    Case 0
'
'        pic方案信息.Move 135, picPane(Index).ScaleHeight - pic方案信息.Height - 270, picPane(Index).ScaleWidth - 270
'        picTree.Move 135, 135, pic方案信息.Width, Abs(picPane(Index).ScaleHeight - pic方案信息.Height - 270 * 2)
'        picTree.Cls
'        picTree.PaintPicture imgBGBlue.Picture, 0, 0, picTree.Width, 360, 0, 0, imgBGBlue.Width, 360
'        picTree.PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, picTree.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
'        picTree.PaintPicture imgBGBlue.Picture, picTree.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picTree.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
'        picTree.PaintPicture imgBGBlue.Picture, 0, picTree.ScaleHeight - Screen.TwipsPerPixelY, picTree.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
'
'        tvw方案.Move Screen.TwipsPerPixelX * 4, 390, Abs(picTree.ScaleWidth - 8 * Screen.TwipsPerPixelX), Abs(picTree.ScaleHeight - 390 - Screen.TwipsPerPixelY * 4)
'
'        pic方案信息.Cls
'        pic方案信息.PaintPicture imgBG.Picture, 0, 0, pic方案信息.Width, 360, 0, 0, imgBG.Width, 360
'        pic方案信息.PaintPicture imgBG.Picture, 0, 360, Screen.TwipsPerPixelX, pic方案信息.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
'        pic方案信息.PaintPicture imgBG.Picture, pic方案信息.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic方案信息.Height - 360, imgBG.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBG.Height - 360
'        pic方案信息.PaintPicture imgBG.Picture, 0, pic方案信息.ScaleHeight - Screen.TwipsPerPixelY, pic方案信息.Width, Screen.TwipsPerPixelY, 0, imgBG.Height - Screen.TwipsPerPixelY, imgBG.Width, Screen.TwipsPerPixelY
'        picFAXX.Move pic方案信息.ScaleWidth - picFAXX.Width - 100
'
'        Refresh
'
'    Case 1
'        fgMain.Move 0, 0, picPane(Index).Width, picPane(Index).Height
'    Case 2
'        lblInfo.Move lblInfo.Left, lblInfo.Top, Abs(picPane(Index).ScaleWidth - 2 * lblInfo.Left), Abs(picPane(Index).ScaleHeight - lblInfo.Top)
'    End Select
'End Sub
'
''Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''    '工具条按钮事件
''    Select Case Button.Key
''        Case "FA"
''            PopupMenu mnuShortFA, , Button.Left + 45, Button.Top + Button.Height + 45
''        Case "XM"
''            PopupMenu mnuShortXM, , Button.Left + 45, Button.Top + Button.Height + 45
''        Case "NewBZ"
''            mnuEditNewBZ_Click
''        Case "ModBZ"
''            mnuEditModBZ_Click
''        Case "DelBZ"
''            mnuEditDelBZ_Click
''        Case "Quit"
''            mnuFileExit_Click
''        Case "Print"
''            mnuFilePrint_Click
''        Case "Preview"
''            mnufilepre_Click
''        Case "Help"
''            mnuHelpTitle_Click
''    End Select
''
''End Sub
'
''Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
''    If ButtonMenu.Parent.Key = "FA" Then
''        Select Case ButtonMenu.Index
''        Case 1
''            mnuEditNewFA_Click
''        Case 2
''            mnuEditModFA_Click
''        Case 3
''            mnuEditDelFA_Click
''        Case 4
''            mnuEditSelFA_Click
''        End Select
''    Else
''        Select Case ButtonMenu.Index
''        Case 1
''            mnuEditNewXM_Click
''        Case 2
''            mnuEditModXM_Click
''        Case 3
''            mnuEditDelXM_Click
''        End Select
''    End If
''End Sub
'
'Private Sub picTree_DblClick()
'    If Left(tvw方案.SelectedItem.Key, 4) = "Root" Then Exit Sub
'    mnuEditModFA_Click
'End Sub
'
'Private Sub picTree_KeyPress(KeyAscii As Integer)
'    If IsNumeric(Mid(m_strTreeKey, 2)) Then
'        If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call mnuEditModFA_Click
'    End If
'End Sub
'
''Private Sub tvw方案_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    If InStr(gstrPrivs, "增删改") = 0 Then Exit Sub
''    If Button = vbRightButton Then
''        PopupMenu mnuShortFA
''    End If
''End Sub
'
'Private Sub tvw方案_NodeClick(ByVal Node As MSComctlLib.Node)
'On Error Resume Next
'    If m_strTreeKey = Node.Key Then Exit Sub     '避免重复刷新
'    m_strTreeKey = Node.Key
'    m_lngCurFAID = Val(Mid(m_strTreeKey, 2))
'    Dim Rs As New ADODB.Recordset, lng总分 As Double
'    gstrSQL = "select * from 病案评分方案 where ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    If Not Rs.EOF Then
'        lbl方案名称 = Rs("名称")
'        lbl分制 = "分制:" & Rs("分制")
'        lbl上值 = "上值:" & Rs("上值")
'        lbl下值 = "下值:" & Rs("下值")
'        lbl总分 = "总分:" & Rs("总分")
'        lng总分 = Rs("总分")
'    Else
'        lbl方案名称 = ""
'        lbl分制 = ""
'        lbl上值 = ""
'        lbl下值 = ""
'        lbl总分 = ""
'    End If
'
'    Rs.Close
'    gstrSQL = "select sum(标准分值) from 病案评分标准 where 上级ID is null and 方案ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    If Not Rs.EOF Then
'        If Abs(lng总分 - Rs.Fields(0)) > 0.01 Then
'            lbl总分 = lbl总分 + "，项目分数和为:" & Rs.Fields(0)
'            lbl总分.ForeColor = vbRed
'        Else
'            lbl总分.ForeColor = vbBlack
'        End If
'    Else
'        lbl总分.ForeColor = vbRed
'    End If
'    '填充列表
'    Call Fill结果
'End Sub
'
'Private Sub subPrint(bytMode As Byte)
'    '-------------------------------------------------
'    '功能:将数据表进行打印,预览和输出到EXCEL
'    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    '-------------------------------------------------
'    Dim objPrint As New zlPrint1Grd
'    Dim objAppRow As zlTabAppRow
'    Dim bytR As Byte
'
'    Set objPrint.Body = fgMain
'    objPrint.Title.Text = tvw方案.SelectedItem.Text
'    objPrint.Title.Font.Name = "楷体_GB2312"
'    objPrint.Title.Font.Size = 18
'    objPrint.Title.Font.Bold = True
'
'    Set objAppRow = New zlTabAppRow
'    Dim Rs As New ADODB.Recordset
'    gstrSQL = "select * from 病案评分方案 where ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'    If Not Rs.EOF Then
'        objAppRow.Add "总分:" & NVL(Rs("总分"), 0)
'        objAppRow.Add "甲级分数线:" & NVL(Rs("上值"), 0)
'        objAppRow.Add "乙级分数线:" & NVL(Rs("下值"), 0)
'
'        objPrint.UnderAppRows.Add objAppRow
'    End If
'
'    Set objAppRow = New zlTabAppRow
'    objAppRow.Add "打印人：" & gstrUserName
'    objAppRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
'    objPrint.BelowAppRows.Add objAppRow
'
'    If bytMode = 1 Then
'        bytR = zlPrintAsk(objPrint)
'        If bytR <> 0 Then zlPrintOrView1Grd objPrint, bytR
'    Else
'        zlPrintOrView1Grd objPrint, bytMode
'    End If
'
'End Sub
'
'
'Private Sub FillTree()
'    '功能:装入评分方案 目前只考虑住院病案
'    Dim rsTemp As New ADODB.Recordset
'    Dim nod As Node, i As Long, FirstKey As String
'    rsTemp.CursorLocation = adUseClient
'
'    fgMain.Tag = ""
'    'Tree的初始化
'    tvw方案.Nodes.Clear
'    '添加根节点
''    Set nod = tvw方案.Nodes.Add(, , "Root", "评分方案列表", "Root", "Root")
''    nod.Expanded = True
'
'    '注意调用格式：先赋值gstrSQL,然后打开数据集
'    gstrSQL = "select ID,名称,选用 from 病案评分方案 where 类型='住院' Order by 选用 desc,名称,启用时间"
'    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
'
'    i = 1
'    Do Until rsTemp.EOF
'        '添加子节点
''        Set nod = tvw方案.Nodes.Add("Root", tvwChild, "A" & rsTemp("ID"), rsTemp("名称"), "Child", "Child")
'        Set nod = tvw方案.Nodes.Add(, , "A" & rsTemp("ID"), rsTemp("名称"), IIf(rsTemp("选用") = 1, "RootSel", "Root"), IIf(rsTemp("选用") = 1, "RootSel", "Root"))
'        If rsTemp("选用") = 1 Then
'            nod.Bold = True
'        Else
'            nod.Bold = False
'        End If
'        If i = 1 Then FirstKey = nod.Key
'        If FirstKey = nod.Key Then i = 2
'        If FirstKey = "" And i = 1 Then FirstKey = nod.Key: i = 2
'        rsTemp.MoveNext
'    Loop
''    '添加根节点
''    Set nod = tvw方案.Nodes.Add(, tvwNext, "RootMZ", "门诊评分方案", "B", "B")
''    nod.Expanded = True
''
''    '注意调用格式：先赋值gstrSQL,然后打开数据集
''    gstrSQL = "select ID,名称,选用 from 病案评分方案 where 类型='门诊' Order by 选用 desc,名称,启用时间"
''    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
''    i = 1
''
''    Do Until rsTemp.EOF
''        '添加子节点
''        Set nod = tvw方案.Nodes.Add("RootMZ", tvwChild, "B" & rsTemp("ID"), IIf(rsTemp("选用") = 1, "√", "") + rsTemp("名称"), "C", "C")
''        rsTemp.MoveNext
''    Loop
'    If i = 1 Then m_strTreeKey = FirstKey   'm_strTreeKey不为空，但是又没有找到。
'    Dim v As Variant
'    For Each v In tvw方案.Nodes
'        If v.Key = FirstKey Then
'            '设置选中
'            v.Selected = True
'            v.EnsureVisible
'            If picTree.Visible = True Then picTree.SetFocus
'        End If
'    Next
'    tvw方案_NodeClick tvw方案.SelectedItem
'
'End Sub
'
'Public Sub Fill结果()
'    '功能:装入对应方案的评分标准
'    Dim rsTemp As New ADODB.Recordset
'    With fgMain
'        .Redraw = flexRDNone
'        .Rows = 1
'        .Clear
'        Dim i As Long
'        .Cell(flexcpText, 0, 0) = "项目"
'        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
'        .Cell(flexcpText, 0, 1) = "标准分值"
'        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
'        .Cell(flexcpText, 0, 2) = "缺陷内容"
'        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
'        .Cell(flexcpText, 0, 3) = "评分标准"
'        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
'        .Cell(flexcpText, 0, 4) = "ID"
'        .Cell(flexcpText, 0, 5) = "上级ID"
'        .Cell(flexcpText, 0, 6) = "方案ID"
'        .Cell(flexcpText, 0, 7) = "序号"
'        rsTemp.CursorLocation = adUseClient
'
'        '确定方案名称
'        If tvw方案.SelectedItem Is Nothing Then .Redraw = flexRDDirect: Exit Sub
'        With tvw方案.SelectedItem
'            Select Case Left(.Key, 1)
'                Case "A", "B"
'                    m_lngCurFAID = Val(Mid(.Key, 2))
'                    gstrSQL = "select * from 病案评分标准视图 Where 隐藏='否' and 方案ID=" & CStr(Mid(.Key, 2))
'                Case Else
'                    Call SetMenu
'                    fgMain.Redraw = flexRDDirect
'                    Exit Sub
'            End Select
'        End With
'        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
'
'        .FocusRect = flexFocusSolid
'        '数据填入
'        .Cols = 8
'        .Rows = rsTemp.RecordCount + 1
'        i = 1
'        Do Until rsTemp.EOF
'            .Cell(flexcpText, i, 0) = NVL(rsTemp.Fields("项目"))
'            .Cell(flexcpAlignment, i, 0) = flexAlignCenterCenter
'            .Cell(flexcpText, i, 1) = IIf(IsNull(rsTemp.Fields("标准分值")), " ", Format(rsTemp.Fields("标准分值"), "####分"))
'            .Cell(flexcpAlignment, i, 1) = flexAlignCenterCenter
'            .Cell(flexcpText, i, 2) = NVL(rsTemp.Fields("缺陷内容"))
'            .Cell(flexcpAlignment, i, 2) = flexAlignLeftTop
'            .Cell(flexcpText, i, 3) = IIf(IsNull(rsTemp.Fields("扣分标准")), "", IIf(rsTemp.Fields("扣分标准") = "甲", "甲级", IIf(rsTemp.Fields("扣分标准") = "乙", "乙级", IIf(rsTemp.Fields("扣分标准") = "丙", "丙级", IIf(rsTemp.Fields("扣分标准") = "否", "单项否决", rsTemp.Fields("扣分标准"))))))
'            .Cell(flexcpAlignment, i, 3) = flexAlignCenterCenter
'            .Cell(flexcpText, i, 4) = NVL(rsTemp.Fields("ID"), 0)
'            .Cell(flexcpText, i, 5) = NVL(rsTemp.Fields("上级ID"), 0)
'            .Cell(flexcpText, i, 6) = NVL(rsTemp.Fields("方案ID"), 0)
'            .Cell(flexcpText, i, 7) = NVL(rsTemp.Fields("序号"), 0)
'            rsTemp.MoveNext
'            i = i + 1
'        Loop
'
'
'        '自动换行
'        .WordWrap = True
'        '合并单元格
'        .MergeCells = 2
'        .MergeCol(.ColIndex("项目")) = True
'        .MergeCol(.ColIndex("标准分值")) = True
'        '对齐设置
'        .ColAlignment(.ColIndex("项目")) = flexAlignLeftCenter
'        .ColAlignment(.ColIndex("标准分值")) = flexAlignCenterCenter
'        .ColAlignment(.ColIndex("评分标准")) = flexAlignCenterCenter
'        '隐藏单元格
'        .ColWidth(.ColIndex("ID")) = 0
'        .ColWidth(.ColIndex("上级ID")) = 0
'        .ColWidth(.ColIndex("方案ID")) = 0
'        .ColWidth(.ColIndex("序号")) = 0
'        '宽度设置
'        .ColWidth(.ColIndex("项目")) = 1500
'        .ColWidth(.ColIndex("标准分值")) = 850
'        .ColWidth(.ColIndex("缺陷内容")) = 3700
'        .ColWidth(.ColIndex("评分标准")) = 1100
'        '行高设置
''        .RowHeightMin = 300
'        '最大宽度设置
''        .ColWidthMax = 7000
'        '自动适应行高、列宽
'        .AutoSizeMode = flexAutoSizeRowHeight
'        .AutoSize .ColIndex("缺陷内容")
'        .SelectionMode = flexSelectionListBox
'        .AllowBigSelection = False
'        .Redraw = flexRDBuffered
'        '选中先前的行
'        If m_lngOldRow > 0 And m_lngOldRow < i Then
'            .Row = m_lngOldRow
'            .Col = 2
'            .ShowCell m_lngOldRow, 2
'            On Error Resume Next
'            If .Visible = True Then .SetFocus
'            fgMain_SelChange
'        ElseIf fgMain.Tag = "" And i > 1 And .Rows > 1 Then
'            m_lngOldRow = 1
'            fgMain.Tag = "选中第一行"
'            .Row = 1
'            .Col = 2
'            .ShowCell m_lngOldRow, 2
'            On Error Resume Next
'            If .Visible = True Then .SetFocus
'            fgMain_SelChange
'        Else
'            lblInfo = "无内容"
'        End If
'
'    End With
'
'    Call SetMenu
'    Call 刷新方案信息
'End Sub
'
'Private Sub SetMenu()
'    '功能:设置修改和删除按钮的有效值
'    '如果没有选择则屏蔽相应按钮
'
'    Dim blnModBZ As Boolean, blnModFA As Boolean
'    If fgMain.Rows <= 1 Then    '无数据
'        fgMain.WallPaper = imgBG_fg(0).Picture
'    Else
'        fgMain.WallPaper = LoadPicture("")
'    End If
'    If IsNumeric(Mid(m_strTreeKey, 2)) Then '方案为空
'        blnModFA = True
'    Else
'        blnModFA = False
'    End If
'    If m_lngCurRow < 1 Or fgMain.Rows <= 1 Then  '标准为空
'        blnModBZ = False
'    Else
'        blnModBZ = True
'    End If
'
'    Dim rsTemp As New ADODB.Recordset
'    gstrSQL = "select count(*) from 病案评分结果 where 方案ID=" & m_lngCurFAID
'    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
'    If rsTemp(0).Value > 0 Then
'        '该方案已经使用
'        fgMain.WallPaper = imgBG_fg(1).Picture
'        blnModBZ = False
'        blnModFA = False
'    End If
'    rsTemp.Close
'
''    mnuShortMenuFA(5).Enabled = IIf(fgMain.Rows > 1, False, blnModFA)    ' 必须为空白方案时才能导入！
''    mnuEditEmportFA.Enabled = IIf(fgMain.Rows > 1, False, blnModFA)
''
''    Toolbar1.Buttons("FA").ButtonMenus(2).Enabled = blnModFA
''    Toolbar1.Buttons("FA").ButtonMenus(3).Enabled = blnModFA
''    mnuEditDelFA.Enabled = blnModFA
''    mnuEditModFA.Enabled = blnModFA
''    mnuShortMenuFA(2).Enabled = blnModFA
''    mnuShortMenuFA(3).Enabled = blnModFA
''    Toolbar1.Buttons("NewBZ").Enabled = blnModFA
''    Toolbar1.Buttons("XM").ButtonMenus("NewXM").Enabled = blnModFA
''    mnuEditNewXM.Enabled = blnModFA
''    mnuEditInsXM.Enabled = blnModFA
''    mnuShortMenuXM(1).Enabled = blnModFA
''    mnuShortMenuXM(2).Enabled = blnModFA
''    mnuShortMnuXM(1).Enabled = blnModFA
''    mnuShortMnuXM(2).Enabled = blnModFA
''    mnuEditNewBZ.Enabled = blnModFA
''    mnuEditInsBZ.Enabled = blnModFA
''    mnuShortMenuBZ(1).Enabled = blnModFA
''    mnuShortMenuBZ(2).Enabled = blnModFA
''    If fgMain.Rows > 1 Then
''        mnuEditInsXM.Enabled = blnModFA
''        mnuEditInsBZ.Enabled = blnModFA
''        mnuShortMenuXM(2).Enabled = blnModFA
''        mnuShortMnuXM(2).Enabled = blnModFA
''        mnuShortMenuBZ(4).Enabled = blnModFA
''        mnuShortMenuBZ(3).Enabled = blnModFA
''        mnuShortMenuXM(3).Enabled = blnModFA
''        mnuShortMnuXM(3).Enabled = blnModFA
''        mnuShortMnuXM(2).Enabled = blnModFA
''    Else
''        mnuEditInsXM.Enabled = False
''        mnuEditInsBZ.Enabled = False
''        mnuShortMenuXM(2).Enabled = False
''        mnuShortMnuXM(2).Enabled = False
''        mnuEditNewBZ.Enabled = False
''        mnuEditInsBZ.Enabled = False
''        Toolbar1.Buttons("NewBZ").Enabled = False
''        mnuShortMenuBZ(1).Enabled = False
''        mnuShortMenuBZ(2).Enabled = False
''        mnuShortMenuBZ(4).Enabled = False
''        mnuShortMenuBZ(3).Enabled = False
''        mnuShortMenuXM(3).Enabled = False
''        mnuShortMnuXM(3).Enabled = False
''        mnuShortMnuXM(2).Enabled = False
''    End If
''
''    Toolbar1.Buttons("ModBZ").Enabled = blnModBZ
''    Toolbar1.Buttons("DelBZ").Enabled = blnModBZ
''    Toolbar1.Buttons("XM").ButtonMenus("ModXM").Enabled = blnModBZ
''    Toolbar1.Buttons("XM").ButtonMenus("DelXM").Enabled = blnModBZ
''
''    mnuEditDelBZ.Enabled = blnModBZ
''    mnuEditModBZ.Enabled = blnModBZ
''    mnuEditDelXM.Enabled = blnModBZ
''    mnuEditModXM.Enabled = blnModBZ
''
''    mnuShortMenuBZ(4).Enabled = blnModBZ
''    mnuShortMenuXM(4).Enabled = blnModBZ
''    mnuShortMnuXM(4).Enabled = blnModBZ
''
''    If m_lngCurSJID <= 0 And fgMain.Rows > 1 Then  '无上级，表示独立评分项
''        mnuEditModBZ.Enabled = False
''        mnuEditDelBZ.Enabled = False
''        mnuShortMenuBZ(3).Enabled = False
''        mnuShortMenuBZ(4).Enabled = False
''        Toolbar1.Buttons("SplitBZ").Enabled = False
''        Toolbar1.Buttons("ModBZ").Enabled = False
''        Toolbar1.Buttons("DelBZ").Enabled = False
''    Else
''        mnuEditModBZ.Enabled = blnModBZ
''        mnuEditDelBZ.Enabled = blnModBZ
''        mnuShortMenuBZ(3).Enabled = blnModBZ
''        mnuShortMenuBZ(4).Enabled = blnModBZ
''        Toolbar1.Buttons("SplitBZ").Enabled = blnModBZ
''        Toolbar1.Buttons("ModBZ").Enabled = blnModBZ
''        Toolbar1.Buttons("DelBZ").Enabled = blnModBZ
''    End If
''
''    If mnuEditNewXM.Enabled = False And mnuEditInsXM.Enabled = False And mnuEditModXM.Enabled = False And mnuEditDelXM.Enabled = False Then
''        Toolbar1.Buttons("XM").Enabled = False
''    Else
''        Toolbar1.Buttons("XM").Enabled = True
''    End If
''    If fgMain.Rows <= 1 Then
''        mnuShortMenuFA(4).Enabled = False     '只要标准存在就允许选用！
''        mnuEditSelFA.Enabled = False
''        Toolbar1.Buttons("FA").ButtonMenus(4).Enabled = False
''    Else
''        If tvw方案.Nodes(tvw方案.SelectedItem.Index).Image = "RootSel" Then
''            mnuShortMenuFA(4).Enabled = False
''            mnuEditSelFA.Enabled = False
''            Toolbar1.Buttons("FA").ButtonMenus(4).Enabled = False
''        Else
''            mnuShortMenuFA(4).Enabled = True
''            mnuEditSelFA.Enabled = True
''            Toolbar1.Buttons("FA").ButtonMenus(4).Enabled = True
''        End If
''    End If
''
''    '如果列表值大于1，则允许打印
''    EnablePrint fgMain.Rows > 1
'
'    '显示记录数信息
'    stbThis.Panels(2).Text = "列表中共显示有" & fgMain.Rows - 1 & "行数据。"
'End Sub
'
'Private Sub 权限控制()
'    '功能:由于有的用户权限不够,故使一些菜单项或按钮不可见
''    If InStr(gstrPrivs, "增删改") = 0 Then
''        mnuEdit.Visible = False
''        'mnusplit1.Visible = False
''        'mnuFileSetup.Visible = False
''        mnuShortMenuBZ(1).Visible = False
''        mnuShortMenuBZ(2).Visible = False
''        mnuShortMenuBZ(3).Visible = False
''        mnuShortMenuBZ(4).Visible = False
''        mnuShortMenuSplit.Visible = False
''        mnuShortMenuXM(1).Visible = False
''        mnuShortMenuXM(2).Visible = False
''        mnuShortMenuXM(3).Visible = False
''        mnuShortMnuXM(1).Visible = False
''        mnuShortMnuXM(2).Visible = False
''        mnuShortMnuXM(3).Visible = False
''        mnuShortMenuFA(1).Visible = False
''        mnuShortMenuFA(2).Visible = False
''        mnuShortMenuFA(3).Visible = False
''        Toolbar1.Buttons("SplitFA").Visible = False
''        Toolbar1.Buttons("FA").Visible = False
''        Toolbar1.Buttons("FA").ButtonMenus(1) = False
''        Toolbar1.Buttons("FA").ButtonMenus(2) = False
''        Toolbar1.Buttons("FA").ButtonMenus(3) = False
''        Toolbar1.Buttons("FA").ButtonMenus(4) = False
''        Toolbar1.Buttons("SplitXM").Visible = False
''        Toolbar1.Buttons("XM").Visible = False
''        Toolbar1.Buttons("XM").ButtonMenus(1) = False
''        Toolbar1.Buttons("XM").ButtonMenus(2) = False
''        Toolbar1.Buttons("XM").ButtonMenus(3) = False
''        Toolbar1.Buttons("SplitBZ").Visible = False
''        Toolbar1.Buttons("NewBZ").Visible = False
''        Toolbar1.Buttons("ModBZ").Visible = False
''        Toolbar1.Buttons("DelBZ").Visible = False
''    End If
'End Sub
'
''Private Sub EnablePrint(ByVal blnEnabled As Boolean)
''    '功能:设置打印和预鉴按钮的有效值
''    '参数:blnEnabled 有效值
''
''    Toolbar1.Buttons("Print").Enabled = blnEnabled
''    Toolbar1.Buttons("Preview").Enabled = blnEnabled
''    mnuFilePre.Enabled = blnEnabled
''    mnuFilePrint.Enabled = blnEnabled
''    mnuFileExcel.Enabled = blnEnabled
''End Sub
'
'Private Sub Show基本要求(lngID As Long, 项目 As String, 标准分值 As String)
'    '根据项目ID显示基本要求
'    Dim Rs As New ADODB.Recordset
'    gstrSQL = "select ID,描述 as 基本要求,上级ID from 病案评分标准 Where ID=" & CStr(lngID)
'    Call zlDatabase.OpenRecordset(Rs, gstrSQL, Me.Caption)
'
'    If Not Rs.EOF Then
'        If m_lngOldSJID > 0 And m_lngOldSJID = lngID Then Exit Sub
'        If IsNull(Rs.Fields("基本要求")) Then
'                lblInfo = "名称：" + 项目 + "  " + IIf(Len(Trim(标准分值)) = 0, "", "(" + 标准分值 + ")")
'                lblInfo = lblInfo + vbCrLf
'        Else
'            If Len(Rs.Fields("基本要求")) > 0 Then
'                lblInfo = "名称：" + 项目 + "  " + IIf(Len(Trim(标准分值)) = 0, "", "(" + 标准分值 + ")")
'                lblInfo = lblInfo + vbCrLf + Rs.Fields("基本要求")
'            End If
'        End If
'    Else
'        lblInfo.Caption = "无内容":
'    End If
'    m_lngOldSJID = m_lngCurSJID
'End Sub
'
'
