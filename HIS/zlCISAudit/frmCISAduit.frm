VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISAduit 
   Caption         =   "电子病案审查"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14370
   Icon            =   "frmCISAduit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   14370
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   0
      Left            =   270
      ScaleHeight     =   3585
      ScaleWidth      =   3570
      TabIndex        =   1
      Top             =   705
      Width           =   3570
      Begin XtremeSuiteControls.TabControl tbcTask 
         Height          =   1830
         Left            =   345
         TabIndex        =   3
         Top             =   615
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
      Left            =   9720
      TabIndex        =   0
      Top             =   105
      Width           =   1125
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8745
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
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
            Object.Width           =   22437
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
      Bindings        =   "frmCISAduit.frx":6852
      Left            =   4020
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmCISAduit"
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
Private mblnAllowModify As Boolean
Private mstrCondition As String
Private mstrFindKey As String
Private mlngTmp As Long
Private mobjFindKey As CommandBarControl
Private mobjPrintView As CommandBarControl
Private mobjPrintPatient As CommandBarControl
Private mobjPrint As CommandBarControl
Private mobjPrint1 As CommandBarControl
Private mintIndex As Integer
Private mlngModul As Long
Private mrsCondition As ADODB.Recordset
Private mblnMuliSelect As Boolean
Private mblnMediAudit  As Boolean
Private mblnMediAuditPass As Boolean
Private mblnTrans As Boolean
Private Type SELECTEDPERSON
    未接收人数 As Long
    未归档人数 As Long
    已拒绝人数 As Long
    已归档人数 As Long
End Type
Private mblnAudit       As Boolean              '是否需要审核后才能归档
Private mblnAuditEnter  As Boolean              '是否允许录入审查意见
Private mblnDoctorAdvice As Boolean             '病人医嘱单、病人医嘱本
Private mSelectedPerson As SELECTEDPERSON
Private mstrFindDeal As String
Private WithEvents mfrmChildPatientAduit As frmChildPatient '出院病人信息
Attribute mfrmChildPatientAduit.VB_VarHelpID = -1
Private WithEvents mfrmChildPatientIn As frmChildPatient    '在院病人信息
Attribute mfrmChildPatientIn.VB_VarHelpID = -1
Private WithEvents mfrmChildQuestion As frmChildQuestion
Attribute mfrmChildQuestion.VB_VarHelpID = -1
Private WithEvents mfrmChildDocumentView As frmChildDocumentView
Attribute mfrmChildDocumentView.VB_VarHelpID = -1
Private WithEvents mfrmChildDocumentScaleView As frmChildDocumentView
Attribute mfrmChildDocumentScaleView.VB_VarHelpID = -1

Private Const conMenu_Manage_magnify = 3946 '放大查阅
Private mblnShowDept As Boolean             '是否显示停用部门
Public Property Let ShowDept(ByVal blnData As Boolean)
    mblnShowDept = blnData
End Property

Public Property Get ShowDept() As Boolean
    ShowDept = mblnShowDept
End Property
Public Property Let AllowModify(ByVal blnData As Boolean)
    mblnAllowModify = blnData
    If blnData = False Then mfrmChildQuestion.AllowModify = blnData
End Property

Public Property Get AllowModify() As Boolean
 AllowModify = mblnAllowModify
End Property
Public Property Get 模块号() As Long
    模块号 = mlngModul
End Property

Private Property Let DataChanged(ByVal blnData As Boolean)
    mfrmChildQuestion.DataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmChildQuestion Is Nothing) Then
        DataChanged = mfrmChildQuestion.DataChanged
    End If
End Property

Private Function GetChildPatient(ByVal intIndex As Integer) As frmChildPatient
    Select Case intIndex
        Case 0
            Set GetChildPatient = mfrmChildPatientAduit
        Case 1
            Set GetChildPatient = mfrmChildPatientIn
    End Select
End Function

Private Function CountSelected() As Boolean
    '******************************************************************************************************************
    '功能：统计选中的个数
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim lngCount As Long

    mSelectedPerson.未接收人数 = 0
    mSelectedPerson.未归档人数 = 0
    mSelectedPerson.已归档人数 = 0
    mSelectedPerson.已拒绝人数 = 0
     
    Select Case mintIndex
        Case 0
            GetChildPatient(mintIndex).labSelect.Caption = "出院"
            With GetChildPatient(mintIndex).VsfBody
                If GetChildPatient(mintIndex).VsfBody.Rows = 2 And GetChildPatient(mintIndex).VsfBody.RowData(1) = "" Then Exit Function
                If .ColIndex("选择") = -1 Then Exit Function
                If .ColIndex("病案状态值") = -1 Then Exit Function
                For lngLoop = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(lngLoop, .ColIndex("选择")))) = 1 Then
                        If Val(.TextMatrix(lngLoop, .ColIndex("病案状态值"))) = "10" Then
                            mSelectedPerson.未接收人数 = mSelectedPerson.未接收人数 + 1
                        ElseIf Val(.TextMatrix(lngLoop, .ColIndex("病案状态值"))) = "2" Then
                            mSelectedPerson.已拒绝人数 = mSelectedPerson.已拒绝人数 + 1
                        End If
                        
                        If .TextMatrix(lngLoop, .ColIndex("病案状态值")) = 3 Then
                            mSelectedPerson.未归档人数 = mSelectedPerson.未归档人数 + 1
                        End If
                        
                        If .TextMatrix(lngLoop, .ColIndex("病案状态值")) = 5 Then
                            mSelectedPerson.已归档人数 = mSelectedPerson.已归档人数 + 1
                        End If
                    End If
                Next
            End With
            GetChildPatient(mintIndex).labNum.Caption = mSelectedPerson.未接收人数 + mSelectedPerson.已拒绝人数 + mSelectedPerson.未归档人数 + mSelectedPerson.已归档人数
        Case 1
            GetChildPatient(mintIndex).labSelect.Visible = False
            GetChildPatient(mintIndex).labNum.Visible = False
            With GetChildPatient(mintIndex).VsfBody
                If .Rows = 1 Then
                    GetChildPatient(mintIndex).LabStatus.Caption = ""
                Else
                    If .ColIndex("姓名") <> -1 Then
                        GetChildPatient(mintIndex).LabStatus.Caption = "姓名：" & .TextMatrix(.Row, .ColIndex("姓名")) & "    住院号：" & .TextMatrix(.Row, .ColIndex("住院号"))
                    End If
                End If
            End With
    End Select
    mblnMuliSelect = (mSelectedPerson.未接收人数 > 0 Or mSelectedPerson.未归档人数 > 0 Or mSelectedPerson.已归档人数 > 0 Or mSelectedPerson.已拒绝人数 > 0)
    CountSelected = True

End Function

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
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_magnify, "打开(&O)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_File_Print * 100, "清单输出(&P)")
        Call NewCommandBar(objControl, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
        Call NewCommandBar(objControl, xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Call NewCommandBar(objControl, xtpControlButton, conMenu_File_Print, "打印(&P)")
        Call NewCommandBar(objControl, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_File_BillPrint * 100, "输出文档")
        Set mobjPrintView = NewCommandBar(objControl, xtpControlButton, conMenu_File_BillPrintView, "预览文档(&E)")
        Set mobjPrint = NewCommandBar(objControl, xtpControlButton, conMenu_File_BillPrint, "打印文档(&T)")
        Set mobjPrint1 = NewCommandBar(objControl, xtpControlButton, conMenu_File_MedRecSetup, "打印设置")
    
    Set mobjPrintPatient = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BatPrint, "输出档案(&B)", True)
    Call NewCommandBar(objMenu, xtpControlButton, conMenu_File_BatPrint * 100, "输出到PDF(&E)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "参数设置(&M)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '编辑
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Plan, "病案抽审(&B)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "自动审查(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit * 10, "批量审查(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Audit, "审查归档(&D)", True)
    
'    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Edit_Untread, "回退操作(&U)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "回退抽审(&1)", , , "回退抽审")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "回退归档(&2)", , , "回退归档")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Pause, "病案封存(&P)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Reuse, "病案解封(&R)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SelAll, "全部选中(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ClsAll, "取消选择(&C)")
        
    '查看
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_ShowStoped, "显示停用部门(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SafeKeep, "封存查看(&F)...", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Column, "选择列项(&H)...", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_ReportView, "查阅病案(&V)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Filter, "过滤(&F)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
    
'     Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "测试", -1, False)
'    objMenu.ID = conMenu_Edit_MediAudit
'    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Edit_MediAudit, "药嘱审查")
    '帮助
    '------------------------------------------------------------------------------------------------------------------
    Call CreateHelpMenu(cbsMain)
    
    '主菜单右侧的查找
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    
    mstrFindKey = GetPara("定位依据", mlngModul, "姓名")
    If mstrFindKey = "" Then mstrFindKey = "姓名"
    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.姓名", , , "姓名")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.住院号", , , "住院号")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.床号", , , "床号")
'''    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&4.就诊卡号", , , "就诊卡号")
    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, ""): cbrCustom.Handle = txtLocation.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "前一条"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "后一条"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon

    Set objPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationMethod, "选项")
    objPopup.IconId = conMenu_View_LocationMethod
    objPopup.Flags = xtpFlagRightAlign
  
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_LocationMethod, "&1.仅查找定位", , , "仅查找定位")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_LocationMethod, "&2.查找并选中", , , "查找并选中")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_LocationMethod, "&3.查找并不选", , , "查找并不选")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_LocationMethod, "&4.查找并反选", , , "查找并反选")

    
    '工具栏定义:包括公共部份
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("标准", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "打印")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "预览")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_magnify, "打开")
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Plan, "抽审", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Audit, "自动")

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_Audit, "归档")
        
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Pause, "封存", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Reuse, "解封")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SafeKeep, "查看封存", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_ReportView, "查阅", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "过滤", True)
    
    Set objControl = NewToolBar(objBar, xtpControlButtonPopup, conMenu_Edit_MediAudit, "药嘱审查", True)
'    objControl.ID = 0
'    Set objPopup = objControl.Add(xtpControlPopup, conMenu_Edit_MediAudit, "药嘱审查1", , False)
'    Set objPopup = NewToolBar(objControl, xtpControlSplitButtonPopup, conMenu_Edit_MediAudit, "药嘱审查")
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出")

    
    '命令的快键绑定:公共部份主界面已处理
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF6, conMenu_Edit_Audit                 '自动
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save           '保存
        .Add 0, vbKeyF3, conMenu_Manage_Plan                '病案抽审
        .Add 0, vbKeyF12, conMenu_File_Parameter            '参数设置
        .Add 0, vbKeyF11, conMenu_Manage_ReportView         '查阅
        .Add 0, vbKeyF10, conMenu_File_BatPrint            '打印病人所有档案
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add FCONTROL, vbKeyE, conMenu_File_BillPrintView   '预览文档
        .Add FCONTROL, vbKeyT, conMenu_File_BillPrint       '打印文档
        .Add FCONTROL, vbKeyZ, conMenu_File_MedRecSetup     '打印设置
        .Add FCONTROL, vbKeyF, conMenu_View_Filter          '过滤
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll          '全选
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_ClsAll       '全清
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
    Dim intLoop         As Integer
    Dim intRow          As Integer
    Dim rs              As New ADODB.Recordset
    Dim rsSQL           As New ADODB.Recordset
    Dim strTmp          As String
    Dim strSQL          As String
    Dim strNow          As String
    Dim strNote         As String
    Dim strDept         As String
    Dim strMsg          As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        
        Set mfrmChildPatientAduit = New frmChildPatient
        Set mfrmChildPatientIn = New frmChildPatient
        Set mfrmChildDocumentView = New frmChildDocumentView
        Set mfrmChildQuestion = New frmChildQuestion
        Set mfrmChildDocumentScaleView = New frmChildDocumentView
        
        Call mfrmChildPatientAduit.zlInitData(Me, 1, mstrPrivs)
        Call mfrmChildPatientIn.zlInitData(Me, 3, mstrPrivs)
        Call mfrmChildDocumentView.zlInitData(Me)
        Call mfrmChildDocumentScaleView.zlInitData(Me)

        Call mfrmChildQuestion.InitData(Me, mlngModul, IsPrivs(mstrPrivs, "审查病案"), mblnAuditEnter, mstrPrivs)
        
        '初始菜单及工具栏
        '--------------------------------------------------------------------------------------------------------------
        Call InitCommandBar
        
        '划分停靠区域
        '--------------------------------------------------------------------------------------------------------------
        Dim objPane As Pane
        Set objPane = dkpMain.CreatePane(1, 100, 200, DockLeftOf, Nothing): objPane.Title = "病人列表": objPane.Options = PaneNoCaption
        Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, Nothing): objPane.Title = "病案审阅": objPane.Options = PaneNoCaption
        Set objPane = dkpMain.CreatePane(3, 300, 100, DockRightOf, objPane): objPane.Title = "反馈问题": objPane.Options = PaneNoCloseable  'Or PaneNoFloatable
        
        dkpMain.SetCommandBars cbsMain
        Call DockPannelInit(dkpMain)
        Call TabControlInit(tbcTask)
        With tbcTask
            .PaintManager.BoldSelected = True
            
            .InsertItem 0, "出院病人", mfrmChildPatientAduit.hWnd, 4
            .InsertItem 1, "在院病人", mfrmChildPatientIn.hWnd, 3
                                                
            If IsPrivs(mstrPrivs, "查阅审查病案") = False Then .Item(0).Visible = False
            If IsPrivs(mstrPrivs, "查阅抽查病案") = False Then .Item(1).Visible = False
            
            If .Item(0).Visible Then
                .Item(0).Selected = True
            ElseIf .Item(1).Visible Then
                .Item(1).Selected = True
            End If
            
        End With
            
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
                                
        '创建过滤条件项目，并进行初始化
        Call ParamCreate(mrsCondition)
        Call ParamAdd(mrsCondition, "提交待收", 1)
        Call ParamAdd(mrsCondition, "接收待审", 1)
        Call ParamAdd(mrsCondition, "拒绝接收", 1)
        Call ParamAdd(mrsCondition, "正在审查", 1)
        Call ParamAdd(mrsCondition, "审查反馈", 1)
        Call ParamAdd(mrsCondition, "审查整改", 1)
        
        Call ParamAdd(mrsCondition, "当前病况", "")
        Call ParamAdd(mrsCondition, "出院情况", "")
        
        Call ParamAdd(mrsCondition, "病人类型", 0)
        Call ParamAdd(mrsCondition, "医保种类", "")
        
        Call ParamAdd(mrsCondition, "审查开始时间", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "审查结束时间", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "归档开始时间", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "归档结束时间", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    
        Call ParamAdd(mrsCondition, "出院病人", 0)
        Call ParamAdd(mrsCondition, "出院开始时间", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "出院结束时间", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        
        Call ParamAdd(mrsCondition, "医嘱开始时间", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "医嘱结束时间", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "住院医师", "")
        Call ParamAdd(mrsCondition, "疾病名称", "")
        Call ParamAdd(mrsCondition, "检查类型", "")
        Call ParamAdd(mrsCondition, "药品信息", "")
                
        '读取缺省时间范围
        strTmp = GetPara("审查缺省范围", mlngModul, "今  天")
        If strTmp = "" Then strTmp = "今  天"
        Call ParamWrite(mrsCondition, "审查开始时间", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "审查结束时间", GetDateTime(strTmp, 2))
        
        strTmp = GetPara("归档缺省范围", mlngModul, "今  天")
        If strTmp = "" Then strTmp = "今  天"
        Call ParamWrite(mrsCondition, "归档开始时间", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "归档结束时间", GetDateTime(strTmp, 2))
        
        strTmp = GetPara("出院缺省范围", mlngModul, "今  天")
        If strTmp = "" Then strTmp = "今  天"
        Call ParamWrite(mrsCondition, "出院开始时间", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "出院结束时间", GetDateTime(strTmp, 2))
        
        '新加条件
        strTmp = GetPara("医嘱缺省范围", mlngModul, "今  天")
        If strTmp = "" Then strTmp = "今  天"
        Call ParamWrite(mrsCondition, "医嘱开始时间", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "医嘱结束时间", GetDateTime(strTmp, 2))
        
        
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
        If tbcTask.Enabled <> Not DataChanged Then
            
            tbcTask.Enabled = Not DataChanged
            
            mfrmChildPatientAduit.Enabled = Not DataChanged
            mfrmChildPatientIn.Enabled = Not DataChanged
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新状态"
        
        With GetChildPatient(mintIndex).VsfBody
            If Val(.TextMatrix(.Row, .ColIndex("病人id"))) = 0 Then
                strTmp = "当前状态下还没有任何病人！"
            Else
                strTmp = "当前状态下共有 " & .Rows - 1 & " 个病人！"
            End If
        End With
        
        stbThis.Panels(2).Text = strTmp
        
        With GetChildPatient(0).VsfBody
            If tbcTask.ItemCount > 0 Then
                If Val(.TextMatrix(.Row, .ColIndex("病人id"))) > 0 Then
                    tbcTask.Item(0).Caption = "出院病人(" & .Rows - 1 & ")"
                Else
                    tbcTask.Item(0).Caption = "出院病人"
                End If
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        Call ExecuteCommand("读取出院病人")
        Call ExecuteCommand("读取在院病人")
        
        Call ExecuteCommand("读取病人病案")
        Call ExecuteCommand("读取反馈记录")
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新指定病人"
        
        Select Case tbcTask.Selected.Index
            Case 0
                If UBound(varParam) >= 1 Then
                    Call mfrmChildPatientAduit.zlRefreshData(mrsCondition, Val(varParam(0)), Val(varParam(1)))
                Else
                    Call mfrmChildPatientAduit.zlRefreshData(mrsCondition)
                End If
            Case 1
                If UBound(varParam) >= 1 Then
                    Call mfrmChildPatientIn.zlRefreshData(mrsCondition, Val(varParam(0)), Val(varParam(1)))
                Else
                    Call mfrmChildPatientIn.zlRefreshData(mrsCondition)
                End If
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取出院病人"
    
        Call GetChildPatient(0).zlRefreshData(mrsCondition)
    '------------------------------------------------------------------------------------------------------------------
    Case "读取在院病人"
    
        Call GetChildPatient(1).zlRefreshData(mrsCondition)
    '------------------------------------------------------------------------------------------------------------------

    Case "读取病人病案"
        
        Select Case mintIndex
            Case 0
                Call mfrmChildPatientAduit.zlShowDocument
            Case 1
                Call mfrmChildPatientIn.zlShowDocument
        End Select
        
        If mfrmChildQuestion.CurrentPatient Then
            Call mfrmChildQuestion.RefreshData(GetChildPatient(mintIndex).Depts, mrsCondition, mblnAuditEnter)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取反馈记录"
        strDept = GetChildPatient(mintIndex).Depts
        Call mfrmChildQuestion.RefreshData(strDept, mrsCondition, mblnAuditEnter)
    '------------------------------------------------------------------------------------------------------------------
    Case "审查报到"
        
        With GetChildPatient(mintIndex).VsfBody
            If Not mblnMuliSelect Then
                strMsg = "确认抽审如下病案吗?" & vbCrLf & vbCrLf
                strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(.Row, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(.Row, .ColIndex("住院号"))
            ElseIf mSelectedPerson.未接收人数 <= 5 Then
                strMsg = "确认抽审【" & mSelectedPerson.未接收人数 & "】份病案吗?" & vbCrLf & vbCrLf
                For intRow = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And Val(.TextMatrix(intRow, .ColIndex("病案状态值"))) = "10" And .TextMatrix(intRow, .ColIndex("封存时间")) = "" Then
                         strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(intRow, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(intRow, .ColIndex("住院号")) & vbCrLf
                    End If
                Next
            Else
                strMsg = "确认抽审【" & mSelectedPerson.未接收人数 & "】份病案？"
            End If
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                If strNow = "" Then strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                If mblnMuliSelect Then
                    For intRow = 1 To .Rows - 1
                        If Val(.TextMatrix(intRow, .ColIndex("ID"))) > 0 And Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And (Val(.TextMatrix(intRow, .ColIndex("病案状态值"))) = 10 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 1) And .TextMatrix(intRow, .ColIndex("封存时间")) = "" Then
                            strTmp = strTmp & "," & Val(.TextMatrix(intRow, .ColIndex("ID")))
                            If Len(strTmp) > (40000 - 20) Then
                                strTmp = Mid(strTmp, 2)
                                strSQL = "zl_病案提交记录_SeReceive('" & strTmp & "','" & UserInfo.姓名 & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                                Call SQLRecordAdd(rsSQL, strSQL)
                                strTmp = ""
                            End If
                        End If
                    Next
                    
                    If Len(strTmp) > 0 Then
                        strTmp = Mid(strTmp, 2)
                        strSQL = "zl_病案提交记录_SeReceive('" & strTmp & "','" & UserInfo.姓名 & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                        Call SQLRecordAdd(rsSQL, strSQL)
                        strTmp = ""
                    End If
                Else
                    
                    If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And (Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 10 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 1) And .TextMatrix(.Row, .ColIndex("封存时间")) = "" Then
                        strSQL = "zl_病案提交记录_SeReceive('" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & "','" & UserInfo.姓名 & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                        Call SQLRecordAdd(rsSQL, strSQL)
                    End If
                        
                End If
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            
            GoTo endHand
        End With
                        
    '------------------------------------------------------------------------------------------------------------------
    Case "审查归档"
    
        With GetChildPatient(mintIndex).VsfBody
            If mintIndex = 0 Then '不用接收归档
                If Not mblnMuliSelect Then
                    strMsg = "确认归档如下病案吗?" & vbCrLf & vbCrLf
                    strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(.Row, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(.Row, .ColIndex("住院号"))
                ElseIf mSelectedPerson.未接收人数 + mSelectedPerson.未归档人数 <= 5 Then
                    strMsg = "确认归档【" & mSelectedPerson.未接收人数 + mSelectedPerson.未归档人数 & "】份病案吗?" & vbCrLf & vbCrLf
                    For intRow = 1 To .Rows - 1
                        If Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And (Val(.TextMatrix(intRow, .ColIndex("病案状态值"))) = "10" Or Val(.TextMatrix(intRow, .ColIndex("病案状态值"))) = "3") And .TextMatrix(intRow, .ColIndex("封存时间")) = "" Then
                             strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(intRow, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(intRow, .ColIndex("住院号")) & vbCrLf
                        End If
                    Next
                Else
                    strMsg = "确认归档【" & mSelectedPerson.未接收人数 + mSelectedPerson.未归档人数 & "】份病案？"
                End If
            End If
            
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                
                If strNow = "" Then strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                
                If mblnMuliSelect Then
                    For intRow = 1 To .Rows - 1
                        If Val(.TextMatrix(intRow, .ColIndex("ID"))) > 0 And Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And (Val(.TextMatrix(intRow, .ColIndex("病案状态值"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 10) And .TextMatrix(.Row, .ColIndex("封存时间")) = "" Then
                            strTmp = strTmp & "," & Val(.TextMatrix(intRow, .ColIndex("ID")))
                            If Len(strTmp) > (4000 - 20) Then
                                strTmp = Mid(strTmp, 2)
                                strSQL = "zl_病案提交记录_Archive('" & strTmp & "','" & UserInfo.姓名 & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                                Call SQLRecordAdd(rsSQL, strSQL)
                                strTmp = ""
                            End If
                        End If
                    Next
                    
                    If Len(strTmp) > 0 Then
                        strTmp = Mid(strTmp, 2)
                        strSQL = "zl_病案提交记录_Archive('" & strTmp & "','" & UserInfo.姓名 & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                        Call SQLRecordAdd(rsSQL, strSQL)
                        strTmp = ""
                    End If
                Else
                    
                    If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And (Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 10) And .TextMatrix(.Row, .ColIndex("封存时间")) = "" Then
                        strSQL = "zl_病案提交记录_Archive('" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & "','" & UserInfo.姓名 & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
                        Call SQLRecordAdd(rsSQL, strSQL)
                    End If
                        
                End If
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "回退接收"
        
        With GetChildPatient(mintIndex).VsfBody
            If Not mblnMuliSelect Then
                strMsg = "确认回退抽审如下病案吗?" & vbCrLf & vbCrLf
                strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(.Row, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(.Row, .ColIndex("住院号"))
            ElseIf mSelectedPerson.未归档人数 <= 5 Then
                strMsg = "确认回退抽审【" & mSelectedPerson.未归档人数 & "】份病案吗?" & vbCrLf & vbCrLf
                For intRow = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And Val(.TextMatrix(intRow, .ColIndex("病案状态值"))) = "3" And .TextMatrix(intRow, .ColIndex("封存时间")) = "" Then
                         strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(intRow, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(intRow, .ColIndex("住院号")) & vbCrLf
                    End If
                Next
            Else
                strMsg = "确认回退抽审【" & mSelectedPerson.未归档人数 & "】份病案？"
            End If
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                If mblnMuliSelect Then
                    For intRow = 1 To .Rows - 1
                        If Val(.TextMatrix(intRow, .ColIndex("ID"))) > 0 And Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And Val(.TextMatrix(intRow, .ColIndex("病案状态值"))) = 3 And .TextMatrix(intRow, .ColIndex("封存时间")) = "" Then
                            If Val(.TextMatrix(intRow, .ColIndex("反馈条数"))) > 0 Then
                                '检查是否有反馈记录
                                strMsg = "当前病案已有反馈记录不能回退!" & vbCrLf & vbCrLf
                                strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(intRow, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(intRow, .ColIndex("住院号"))
                                
                                Call MsgBox(strMsg, vbQuestion, ParamInfo.系统名称)
                                ExecuteCommand = False
                                GoTo endHand
                            Else
                                strTmp = strTmp & "," & Val(.TextMatrix(intRow, .ColIndex("ID")))
                                If Len(strTmp) > (40000 - 20) Then
                                    strTmp = Mid(strTmp, 2)
                                    strSQL = "zl_病案提交记录_SeUnReceive('" & strTmp & "')"
                                    Call SQLRecordAdd(rsSQL, strSQL)
                                    strTmp = ""
                                End If
                            End If
                        End If
                    Next
                    
                    If Len(strTmp) > 0 Then
                        strTmp = Mid(strTmp, 2)
                        strSQL = "zl_病案提交记录_SeUnReceive('" & strTmp & "')"
                        Call SQLRecordAdd(rsSQL, strSQL)
                        strTmp = ""
                    End If
                Else
                    If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 3 And .TextMatrix(.Row, .ColIndex("封存时间")) = "" Then
                        If Val(.TextMatrix(.Row, .ColIndex("反馈条数"))) > 0 Then
                            strMsg = "当前病案已有反馈记录不能回退!" & vbCrLf & vbCrLf
                            strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(.Row, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(.Row, .ColIndex("住院号"))
                            
                            Call MsgBox(strMsg, vbQuestion, ParamInfo.系统名称)
                            ExecuteCommand = False
                            GoTo endHand
                        Else
                            strSQL = "zl_病案提交记录_SeUnReceive('" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & "')"
                            Call SQLRecordAdd(rsSQL, strSQL)
                        End If
                    End If
                End If
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "回退归档"
    
        With GetChildPatient(mintIndex).VsfBody
            If Not mblnMuliSelect Then
                strMsg = "确认回退归档如下病案吗?" & vbCrLf & vbCrLf
                strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(.Row, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(.Row, .ColIndex("住院号"))
            ElseIf mSelectedPerson.已归档人数 <= 5 Then
                strMsg = "确认回退归档【" & mSelectedPerson.已归档人数 & "】份病案？" & vbCrLf & vbCrLf
                For intRow = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And Val(.TextMatrix(intRow, .ColIndex("病案状态值"))) = "5" And .TextMatrix(intRow, .ColIndex("封存时间")) = "" Then
                         strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(intRow, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(intRow, .ColIndex("住院号")) & vbCrLf
                    End If
                Next
            Else
                strMsg = "确认回退归档【" & mSelectedPerson.已归档人数 & "】份病案？"
            End If
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                If mblnMuliSelect Then
                    For intRow = 1 To .Rows - 1
                        If Val(.TextMatrix(intRow, .ColIndex("ID"))) > 0 And Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And Val(.TextMatrix(intRow, .ColIndex("病案状态值"))) = 5 And .TextMatrix(intRow, .ColIndex("封存时间")) = "" Then
                            strTmp = strTmp & "," & Val(.TextMatrix(intRow, .ColIndex("ID")))
                            If Len(strTmp) > (40000 - 20) Then
                                strTmp = Mid(strTmp, 2)
                                strSQL = "zl_病案提交记录_UnArchive('" & strTmp & "')"
                                Call SQLRecordAdd(rsSQL, strSQL)
                                strTmp = ""
                            End If
                        End If
                    Next
                    
                    If Len(strTmp) > 0 Then
                        strTmp = Mid(strTmp, 2)
                        strSQL = "zl_病案提交记录_UnArchive('" & strTmp & "')"
                        Call SQLRecordAdd(rsSQL, strSQL)
                        strTmp = ""
                    End If
                Else
                    If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 5 And .TextMatrix(.Row, .ColIndex("封存时间")) = "" Then
                        strSQL = "zl_病案提交记录_UnArchive('" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & "')"
                        Call SQLRecordAdd(rsSQL, strSQL)
                    End If
                End If
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "封存病案"
    
        With GetChildPatient(mintIndex).VsfBody
            If Val(.TextMatrix(.Row, .ColIndex("病人id"))) = 0 Then GoTo endHand
            
            If MsgBox("您是否真的要封存当前选中病人的电子病案吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                
                If frmPubNoteEdit.ShowNoteEdit(Me, "输入封存理由", strNote) Then
                    strSQL = "zl_病案封存记录_Lock(" & Val(.TextMatrix(.Row, .ColIndex("病人id"))) & "," & Val(.TextMatrix(.Row, .ColIndex("主页id"))) & ",'" & UserInfo.姓名 & "',Sysdate,'" & strNote & "')"
                    
                    Call SQLRecordAdd(rsSQL, strSQL)
                    ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
                End If
            End If
            GoTo endHand
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "解封病案"
    
        With GetChildPatient(mintIndex).VsfBody
            If Val(.TextMatrix(.Row, .ColIndex("病人id"))) = 0 Then GoTo endHand
            
            If MsgBox("您是否真的要解封当前选中病人的电子病案吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then

                strSQL = "zl_病案封存记录_UnLock(" & Val(.TextMatrix(.Row, .ColIndex("病人id"))) & "," & Val(.TextMatrix(.Row, .ColIndex("主页id"))) & ")"
                
                Call SQLRecordAdd(rsSQL, strSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "封存查看"
        
        Call frmCISAuditSafeKeep.zlInitData(Me, mlngModul)
        frmCISAuditSafeKeep.Show vbModal, Me
        
    '------------------------------------------------------------------------------------------------------------------
    Case "过滤数据"
        
        mrsCondition.Filter = ""
        ExecuteCommand = frmCISAduitFilter.ShowPara(Me, mrsCondition)

        GoTo endHand
        
    '--------------------------------------------------------------------------------------------------------------
    Case "全部选中"
        With GetChildPatient(mintIndex).VsfBody
            If .ColIndex("选择") = -1 Then Exit Function
            .Cell(flexcpText, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 1
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "全部取消"
        With GetChildPatient(mintIndex).VsfBody
            If .ColIndex("选择") = -1 Then Exit Function
            .Cell(flexcpText, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 0
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "前一条"
        
        With GetChildPatient(mintIndex).VsfBody
            If .Row > 1 Then
                .Row = .Row - 1
                .ShowCell .Row, .Col
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "后一条"
        
        With GetChildPatient(mintIndex).VsfBody
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            End If
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case "列项信息设置"
        
        GetChildPatient(mintIndex).zlColumnSelect
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读注册表"
        
        mstrFindDeal = Trim(zlDatabase.GetPara("查到后处理", ParamInfo.系统号, mlngModul, "仅查找定位"))
        
        If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
            '使用个性化设置
            mstrFindKey = Trim(GetPara("定位依据", mlngModul, "姓名"))
            
            Call RestoreWinState(Me, App.ProductName)
            
            dkpMain.LoadStateFromString (GetRegister(私有模块, Me.Name & "\界面设置\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString))
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "写注册表"
        
        Call zlDatabase.SetPara("查到后处理", mstrFindDeal, ParamInfo.系统号, mlngModul)
        
        If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
            '使用个性化设置
            Call SaveWinState(Me, App.ProductName)
            Call SetPara("定位依据", mstrFindKey, mlngModul)

        End If
        
        With GetChildPatient(mintIndex)
            If .cboDept.ListIndex >= 0 Then
                Call SetPara("上次状态", mintIndex & ";" & .cboDept.ItemData(.cboDept.ListIndex) & ";" & .VsfBody.TextMatrix(.VsfBody.Row, .VsfBody.ColIndex("病人id")) & ";" & .VsfBody.TextMatrix(.VsfBody.Row, .VsfBody.ColIndex("主页id")) & ";" & .tvw.SelectedItem.Key, 模块号)
            Else
                Call SetPara("上次状态", mintIndex & ";0;" & .VsfBody.TextMatrix(.VsfBody.Row, .VsfBody.ColIndex("病人id")) & ";" & .VsfBody.TextMatrix(.VsfBody.Row, .VsfBody.ColIndex("主页id")) & ";" & .tvw.SelectedItem.Key, 模块号)
            End If
        End With
        
        Call SetRegister(私有模块, Me.Name & "\界面设置\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    End Select

    ExecuteCommand = True

    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:

End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BillPrintView                    '预览当前文档
        
        If Not mfrmChildDocumentView Is Nothing Then
        
            Call mfrmChildDocumentView.zlPrintDocument(cbsMain, 1, , , mblnDoctorAdvice)
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BillPrint                    '打印当前文档
        
        If Not mfrmChildDocumentView Is Nothing Then
        
            Call mfrmChildDocumentView.zlPrintDocument(cbsMain, 2, , , mblnDoctorAdvice)
            
        End If
    Case conMenu_File_MedRecSetup '打印设置当前文档
        If Not mfrmChildDocumentView Is Nothing Then
        
            Call mfrmChildDocumentView.zlPrintSet(Control, mblnDoctorAdvice)
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_magnify '放大
    
        Call GetChildPatient(mintIndex).FileBatPrint
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BatPrint
        Call frmCISAduitPDF.ShowMe(Me, GetChildPatient(mintIndex).VsfBody, mintIndex, mblnDoctorAdvice, False)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_BatPrint * 100
        Call frmCISAduitPDF.ShowMe(Me, GetChildPatient(mintIndex).VsfBody, mintIndex, mblnDoctorAdvice, True)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Parameter
        If frmCISAduitPara.ShowEdit(Me, mstrPrivs) Then
            If zlDatabase.GetPara("住院医嘱打印", ParamInfo.系统号, ParamInfo.模块号, "病人医嘱本", , IsPrivs(mstrPrivs, "参数设置")) = "病人医嘱本" Then
                mblnDoctorAdvice = False
            Else
                mblnDoctorAdvice = True
            End If
        
            Call GetChildPatient(mintIndex).zlRefreshStruct
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Plan
        
        If ExecuteCommand("审查报到") Then
            '刷新当前病人的显示信息
            Call ExecuteCommand("读取出院病人")
            Call ExecuteCommand("读取病人病案")
            Call ExecuteCommand("刷新状态")
            If mblnMuliSelect Then
                MsgBox "本次成功【抽审】病案【" & mSelectedPerson.未接收人数 & "】份", vbInformation, ParamInfo.产品名称
            End If
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh
        
        Call ExecuteCommand("刷新数据")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_ReportView      '查阅首页
        Call RecordLook
    Case conMenu_Edit_Audit '自动审查
        With frmCISAduitAuto
            If mintIndex = 0 Then
                .lng提交Id = mfrmChildQuestion.提交Id
            Else
                .lng提交Id = -1
            End If
            .lng病人ID = GetChildPatient(mintIndex).VsfBody.TextMatrix(GetChildPatient(mintIndex).VsfBody.Row, GetChildPatient(mintIndex).VsfBody.ColIndex("病人ID"))
            .lng主页ID = GetChildPatient(mintIndex).VsfBody.TextMatrix(GetChildPatient(mintIndex).VsfBody.Row, GetChildPatient(mintIndex).VsfBody.ColIndex("主页ID"))
            .lng科室ID = GetChildPatient(mintIndex).VsfBody.TextMatrix(GetChildPatient(mintIndex).VsfBody.Row, GetChildPatient(mintIndex).VsfBody.ColIndex("出院科室ID"))
            
            .strLink = IIf(mintIndex = 1, "1", "2") '1为审查 2为抽查
            .strTreeSelect = mfrmChildPatientAduit.tvw.SelectedItem.Key
            .Show vbModal
            If .blnCancel Then
                Set frmCISAduitAuto = Nothing
                Exit Sub
            End If
        End With
        Set frmCISAduitAuto = Nothing
        Call ExecuteCommand("读取反馈记录")
    Case conMenu_Edit_Audit * 10
        If frmCISAduitAutos.ShowMe(Me, IIf(mintIndex = 1, 1, 2), GetChildPatient(mintIndex).VsfBody) Then
        Call ExecuteCommand("读取反馈记录")
        End If
    Case conMenu_Manage_Audit
    
        If ExecuteCommand("审查归档") Then
            Call ExecuteCommand("读取出院病人")
            Call ExecuteCommand("读取病人病案")
            Call ExecuteCommand("刷新状态")
            If mblnMuliSelect Then
                If mintIndex = 0 Then
                    MsgBox "本次成功【归档】病案【" & mSelectedPerson.未归档人数 & "】份", vbInformation, ParamInfo.产品名称
                ElseIf mintIndex = 1 Then
                    MsgBox "本次成功【归档】病案【" & mSelectedPerson.未归档人数 & "】份", vbInformation, ParamInfo.产品名称
                End If
            End If
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Untread
        
        Select Case Control.Caption
            Case "回退抽审(&1)"
            
                If ExecuteCommand("回退接收") Then
                    Call ExecuteCommand("读取出院病人")
                    Call ExecuteCommand("读取病人病案")
                    Call ExecuteCommand("刷新状态")
                    If mblnMuliSelect Then
                        MsgBox "本次成功【回退接收】病案【" & mSelectedPerson.未归档人数 & "】份", vbInformation, ParamInfo.产品名称
                    End If
                End If
            Case "回退归档(&2)"
                If ExecuteCommand("回退归档") Then
                    Call ExecuteCommand("读取出院病人")
                    Call ExecuteCommand("读取病人病案")
                    Call ExecuteCommand("刷新状态")
                    If mblnMuliSelect Then
                        MsgBox "本次成功【回退归档】病案【" & mSelectedPerson.已归档人数 & "】份", vbInformation, ParamInfo.产品名称
                    End If
                End If
        End Select
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Pause
    
        If ExecuteCommand("封存病案") Then
            '刷新当前病人的显示信息
            Call ExecuteCommand("刷新指定病人")
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Reuse
    
        If ExecuteCommand("解封病案") Then
            '刷新当前病人的显示信息
            Call ExecuteCommand("刷新指定病人")
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SafeKeep
        
        Call ExecuteCommand("封存查看")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SelAll
        
        Call ExecuteCommand("全部选中")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ClsAll
    
        Call ExecuteCommand("全部取消")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationMethod
    
        mstrFindDeal = Control.Parameter
        cbsMain.RecalcLayout
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter '过滤
        
        If ExecuteCommand("过滤数据") Then
            Call ExecuteCommand("读取出院病人")
            Call ExecuteCommand("读取在院病人")
            Call ExecuteCommand("读取病人病案")
            Call ExecuteCommand("读取反馈记录")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Column
        
        Call ExecuteCommand("列项信息设置")
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward
        Call ExecuteCommand("前一条")
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward
        Call ExecuteCommand("后一条")
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        
        mobjFindKey.Execute
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
    
        Call LocationObj(txtLocation)
        
    '-----------------------------------------------------------------------------------------------------------------
    
    Case conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 22
        '合理用药审查
        Call mfrmChildDocumentView.zlMediAuditShell(Control)
    Case conMenu_View_ShowStoped
        Control.Checked = Not Control.Checked
        Me.ShowDept = Control.Checked
        If mintIndex = 0 Then
        Call GetChildPatient(mintIndex).InitData(Me.ShowDept)
        With GetChildPatient(mintIndex).VsfBody
            mfrmChildQuestion.AllowModify = Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 4 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 6
        End With
        
        Else
            mfrmChildQuestion.AllowModify = True
             Call GetChildPatient(mintIndex).InitData(False)
        End If
    Call ExecuteCommand("读取病人病案")
    Call ExecuteCommand("刷新状态")
     '-----------------------------------------------------------------------------------------------------------------
    Case Else
    
        If Control.ID > 400 And Control.ID < 500 Then
            With GetChildPatient(mintIndex).VsfBody
                Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me, "ID=" & Val(.TextMatrix(.Row, .ColIndex("ID"))), "病人id=" & Val(.TextMatrix(.Row, .ColIndex("病人id"))), "主页id=" & Val(.TextMatrix(.Row, .ColIndex("主页id"))))
            End With
        Else
             '与业务无关的功能，公共的功能
            Select Case mintIndex
                Case 0
                    Call CommandBarExecutePublic(Control, Me, GetChildPatient(mintIndex).VsfBody, "出院病人列表清单")
                Case 1
                    Call CommandBarExecutePublic(Control, Me, GetChildPatient(mintIndex).VsfBody, "在院病人列表清单")
            End Select
        End If
        
    End Select
    Call CountSelected

End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    If CommandBar.Parent.ID = conMenu_Edit_MediAudit Then '加载药嘱审查
        Call mfrmChildDocumentView.zlMediAudit(CommandBar)
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHand

    With GetChildPatient(mintIndex).VsfBody
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel  '预览,打印,输出到Excel
            Control.Enabled = (Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_BillPrintView, conMenu_File_BillPrint, conMenu_File_BatPrint, conMenu_File_MedRecSetup, conMenu_File_BatPrint * 100

            Control.Visible = IsPrivs(mstrPrivs, "打印预览文档")
            Control.Enabled = Control.Visible

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_MediAudit                         '药嘱审查
            
            Control.Visible = mfrmChildDocumentView.GetTbcStatus And mblnMediAudit And mblnMediAuditPass
        Case conMenu_Manage_Plan                            '病案抽审

            Control.Visible = IsPrivs(mstrPrivs, "病案抽审")
            If mblnMuliSelect Then
                Control.Enabled = (Control.Visible And mintIndex = 0 And mSelectedPerson.未接收人数 > 0) And .TextMatrix(.Row, .ColIndex("封存时间")) = ""
            Else
                If mintIndex <> 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = Control.Visible And mintIndex = 0 And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And (Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 10 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 1) And .TextMatrix(.Row, .ColIndex("封存时间")) = ""
                End If
            End If
            If Me.AllowModify = False Then Control.Enabled = Me.AllowModify
        '--------------------------------   ------------------------------------------------------------------------------
        Case conMenu_Edit_Audit                 '自动审查
            Control.Visible = IsPrivs(mstrPrivs, "归档病案")
            If mintIndex = 1 Then
                Control.Enabled = (Control.Visible And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0)
            Else
                Control.Enabled = (Control.Visible And (Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) > 10) And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0)
            End If
            If Me.AllowModify = False Then Control.Enabled = Me.AllowModify
        Case conMenu_Edit_Audit * 10            '批量审查
            Control.Enabled = (Control.Visible And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0)
        Case conMenu_Manage_Audit                           '归档病案

            Control.Visible = IsPrivs(mstrPrivs, "归档病案")

            If mblnMuliSelect Then
                Control.Enabled = (Control.Visible And mintIndex = 0 And Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) <> 4 And Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) <> 5 And (mSelectedPerson.未接收人数 + mSelectedPerson.未归档人数) > 0) And IIf(Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 10, Not mblnAudit, True) And .TextMatrix(.Row, .ColIndex("封存时间")) = ""
            Else
                If Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) > 10 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 5 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 6 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (Control.Visible And mintIndex = 0 And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) <> 2 And Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) <> 4 And Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) <> 5) And IIf(Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 10, Not mblnAudit, True) And .TextMatrix(.Row, .ColIndex("封存时间")) = ""
                End If
            End If
            If Me.AllowModify = False Then Control.Enabled = Me.AllowModify
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Untread                           '回退接收/拒绝/归档

            Select Case Control.Caption
                Case "回退抽审(&1)"
                    Control.Visible = IsPrivs(mstrPrivs, "回退抽审")
                Case "回退归档(&2)"
                    Control.Visible = IsPrivs(mstrPrivs, "回退归档")
            End Select

            If mblnMuliSelect Then
                Select Case Control.Caption
                    Case "回退抽审(&1)"
                        Control.Enabled = (Control.Visible And mintIndex = 0 And mSelectedPerson.未归档人数 > 0)
                    Case "回退归档(&2)"
                        Control.Enabled = (Control.Visible And mintIndex = 0 And mSelectedPerson.已归档人数 > 0)
                End Select
            Else
                Select Case Control.Caption
                    Case "回退抽审(&1)"
                    
                        If .ColIndex("病案状态值") = -1 Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = (Control.Visible And mintIndex = 0 And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 3) And .TextMatrix(.Row, .ColIndex("封存时间")) = ""
                        End If
    
                    Case "回退归档(&2)"
                        If .ColIndex("病案状态值") = -1 Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = (Control.Visible And mintIndex = 0 And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 5) And .TextMatrix(.Row, .ColIndex("封存时间")) = ""
                        End If
                End Select
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Pause                             '病案封存

            Control.Visible = IsPrivs(mstrPrivs, "封存病案")
            Control.Enabled = (DataChanged = False And Control.Visible And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And .TextMatrix(.Row, .ColIndex("封存时间")) = "")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Reuse                             '病案解封

            Control.Visible = IsPrivs(mstrPrivs, "解封病案")
            Control.Enabled = (DataChanged = False And Control.Visible And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 And .TextMatrix(.Row, .ColIndex("封存时间")) <> "")
        
        '------------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SafeKeep
            
            Control.Visible = IsPrivs(mstrPrivs, "封存查看")
            
        '------------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SelAll, conMenu_Edit_ClsAll                       '全选、全清

            Control.Enabled = (DataChanged = False And Control.Visible And mintIndex <> 3 And Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0)
        Case conMenu_Manage_ReportView
            If Val(.TextMatrix(.Row, .ColIndex("ID"))) = 0 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Filter, conMenu_View_Refresh

            Control.Enabled = (DataChanged = False And Control.Visible)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Forward

            Control.Enabled = (.Row > 1 And DataChanged = False)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Backward

            Control.Enabled = (.Row < .Rows - 1 And DataChanged = False)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationMethod               '
            Control.Checked = (mstrFindDeal = Control.Parameter)
            Control.Enabled = (DataChanged = False)

        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationItem        '
            Control.Checked = (mstrFindKey = Control.Parameter)
            Control.Enabled = (DataChanged = False)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Location, conMenu_View_Column
             Control.Enabled = (DataChanged = False)
        '--------------------------------------------------------------------------------------------------------------
        Case Else
            Call CommandBarUpdatePublic(Control, Me)
        End Select
    End With

    '------------------------------------------------------------------------------------------------------------------
errHand:
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = picPane(0).hWnd
        Case 2
            Item.Handle = mfrmChildDocumentView.hWnd
        Case 3
            Item.Handle = mfrmChildQuestion.hWnd
    End Select
End Sub

Private Sub Form_Activate()
    Dim varTmp As Variant
    Dim intPos As Integer
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    DoEvents
    mblnAudit = (zlDatabase.GetPara("接收才能归档", ParamInfo.系统号, ParamInfo.模块号, "0", , IsPrivs(mstrPrivs, "参数设置")) = 1)
    mblnAuditEnter = (zlDatabase.GetPara("允许自由录入审查意见", ParamInfo.系统号, ParamInfo.模块号, "0", , IsPrivs(mstrPrivs, "参数设置")) = 1)
    If zlDatabase.GetPara("住院医嘱打印", ParamInfo.系统号, ParamInfo.模块号, "病人医嘱本", , IsPrivs(mstrPrivs, "参数设置")) = "病人医嘱本" Then
        mblnDoctorAdvice = False
    Else
        mblnDoctorAdvice = True
    End If
    
    If ExecuteCommand("初始数据") = False Then GoTo errHand
    
    Call ExecuteCommand("刷新数据")
    
    mblnAllowClose = True
    
    varTmp = Split(GetPara("上次状态", 模块号, "0"), ";")
    If Val(varTmp(0)) >= 0 And Val(varTmp(0)) <= 2 And tbcTask.ItemCount > Val(varTmp(0)) Then
        
        If tbcTask.Item(Val(varTmp(0))).Visible Then
            tbcTask.Item(Val(varTmp(0))).Selected = True
        
            If UBound(varTmp) > 2 Then
                Call GetChildPatient(mintIndex).zlLocationPatient(3, , , , Val(varTmp(2)), Val(varTmp(3)), Val(varTmp(1)))
            End If
            If UBound(varTmp) > 3 Then
                If varTmp(4) <> "" Then
                    intPos = InStr(varTmp(4), "K")
                    If InStr(varTmp(4), "K") > 0 Then
                        Call GetChildPatient(mintIndex).zlLocationDocument(Val(varTmp(2)), Val(varTmp(3)), Val(Mid(varTmp(4), 2, intPos - 2)), Mid(varTmp(4), intPos + 1))
                    ElseIf InStr(varTmp(4), "R") Then
                        Call GetChildPatient(mintIndex).zlLocationDocument(Val(varTmp(2)), Val(varTmp(3)), Val(Mid(varTmp(4), 2)), "")
                    Else
                        Call GetChildPatient(mintIndex).zlLocationDocument(Val(varTmp(2)), Val(varTmp(3)), 2, varTmp(4))
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub

    '------------------------------------------------------------------------------------------------------------------
errHand:
    mblnAllowClose = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mblnAllowClose = False
    mblnShowDept = False
    mstrPrivs = UserInfo.模块权限
    mlngModul = ParamInfo.模块号
    
    '检查是否具有合理用药权限
    If InStrRev(GetPrivFunc(100, 1253), "合理用药监测") > 0 Then
        mblnMediAudit = True
    Else
        mblnMediAudit = False
    End If
    
    '检查是否启用了合理用药接口 0-未启用 1-美通接口 2-大通接口
    '这里只判断美通接口是否开启
    If Val(zlDatabase.GetPara(30, glngSys)) = 1 Then
        mblnMediAuditPass = True
    Else
        mblnMediAuditPass = False
    End If
    
    Call ExecuteCommand("初始控件")
    Call ExecuteCommand("读注册表")

    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1261_1", "ZL1_INSIDE_1254_1", "ZL1_INSIDE_1254_2")
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 1, 250, 100, 500, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 3, 250, 100, 400, Me.ScaleHeight)
    dkpMain.RecalcLayout
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = Not mblnAllowClose
    
    If Cancel = False Then

        
        Call ExecuteCommand("写注册表")
        

        If Not (mfrmChildDocumentView Is Nothing) Then Unload mfrmChildDocumentView
        If Not (mfrmChildDocumentScaleView Is Nothing) Then Unload mfrmChildDocumentScaleView
        If Not (mfrmChildQuestion Is Nothing) Then Unload mfrmChildQuestion
        If Not (mfrmChildPatientAduit Is Nothing) Then Unload mfrmChildPatientAduit
        If Not (mfrmChildPatientIn Is Nothing) Then Unload mfrmChildPatientIn
        If Not (frmChildScale Is Nothing) Then Unload frmChildScale
        Set frmChildScale = Nothing
        
        Set mrsCondition = Nothing
    End If

End Sub

Private Sub mfrmChildPatientAduit_AfterDeptChanged()
     Call ExecuteCommand("刷新状态")
End Sub

Private Sub mfrmChildPatientAduit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call CountSelected
End Sub

Private Sub mfrmChildPatientAduit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    If Button = 2 Then
        Set cbrPopupBar = CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        
        cbrPopupBar.ShowPopup
    End If
End Sub

Private Sub mfrmChildPatientAduit_AfterDocumentChanged(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strObject As String, ByVal strParam As String, ByVal strCaption As String, ByVal lng提交Id As Long, ByVal blnDataMove As Boolean, ByVal blnScale As Boolean)
     If blnScale Then
        
        Call mfrmChildDocumentScaleView.zlRefresh(lng病人ID, lng主页ID, strObject, strParam, strCaption, blnDataMove)
        Call frmChildScale.zlInitData(mfrmChildDocumentScaleView)
        frmChildScale.Show
        
'        Set mfrmChildDocumentScaleView = New frmChildDocumentView
'        Call mfrmChildDocumentScaleView.zlInitData(Me)
    Else
        Call mfrmChildDocumentView.zlRefresh(lng病人ID, lng主页ID, strObject, strParam, strCaption, blnDataMove)
        
        mobjPrintView.Caption = "预览""" & mfrmChildPatientAduit.Title & """(&E)"
        mobjPrint.Caption = "打印""" & mfrmChildPatientAduit.Title & """(&T)"
        mobjPrint1.Caption = "打印设置""" & mfrmChildPatientAduit.Title & """(&T)"
        
        With GetChildPatient(mintIndex).VsfBody
            mobjPrintPatient.Caption = "打印""" & .TextMatrix(.Row, .ColIndex("姓名")) & """的档案(&B)"
        End With
        cbsMain.RecalcLayout
        
        If Not (mfrmChildQuestion Is Nothing) Then
            Call mfrmChildQuestion.SetParamter(lng病人ID, lng主页ID, strObject, strParam, lng提交Id)
            If mfrmChildQuestion.CurrentPatient Then
                If mintIndex = 0 Then
                    With GetChildPatient(mintIndex).VsfBody
                        mfrmChildQuestion.AllowModify = (Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 4 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 6) And zlCommFun.NVL(.TextMatrix(.Row, .ColIndex("封存时间")) = "" And Me.AllowModify)
                    End With
                Else
                    mfrmChildQuestion.AllowModify = True
                End If
                
                If strParam = "" And (strObject = "住院病历" Or strObject = "护理病历" Or strObject = "护理记录" Or strObject = "医嘱报告" Or strObject = "疾病证明" Or strObject = "知情文件") Then
                    '不刷新详细数据
                Else
                    Call mfrmChildQuestion.RefreshData(GetChildPatient(mintIndex).Depts, mrsCondition, mblnAuditEnter)
                End If
            End If
        End If
    End If
End Sub

Private Sub mfrmChildPatientAduit_StatusChanged()
     Call ExecuteCommand("刷新状态")
End Sub

Private Sub mfrmChildPatientIn_AfterDeptChanged()
     Call ExecuteCommand("刷新状态")
End Sub

Private Sub mfrmChildPatientIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    If Button = 2 Then
        Set cbrPopupBar = CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        cbrPopupBar.ShowPopup
    End If
End Sub

Private Sub mfrmChildPatientIn_AfterDocumentChanged(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strObject As String, ByVal strParam As String, ByVal strCaption As String, ByVal lng提交Id As Long, ByVal blnDataMove As Boolean, ByVal blnScale As Boolean)
    If blnScale Then
        '抽查病案缩放 --ZQ
'        Set mfrmChildDocumentScaleView = New frmChildDocumentView
'        Call mfrmChildDocumentScaleView.zlInitData(Me)
  
        Call mfrmChildDocumentScaleView.zlRefresh(lng病人ID, lng主页ID, strObject, strParam, strCaption, blnDataMove)
        Call frmChildScale.zlInitData(mfrmChildDocumentScaleView)
        frmChildScale.Show
        
'        Unload frmChildScale
'        Set frmChildScale = Nothing
'        Set mfrmChildDocumentScaleView = Nothing
    Else
        Call mfrmChildDocumentView.zlRefresh(lng病人ID, lng主页ID, strObject, strParam, strCaption, blnDataMove)
        
        mobjPrintView.Caption = "预览""" & mfrmChildPatientIn.Title & """(&E)"
        mobjPrint.Caption = "打印""" & mfrmChildPatientIn.Title & """(&T)"
        mobjPrint1.Caption = "打印设置""" & mfrmChildPatientIn.Title & """(&T)"
        With GetChildPatient(mintIndex).VsfBody
            mobjPrintPatient.Caption = "打印""" & .TextMatrix(.Row, .ColIndex("姓名")) & """的档案(&B)"
        End With
        cbsMain.RecalcLayout
        
        If Not (mfrmChildQuestion Is Nothing) Then
            Call mfrmChildQuestion.SetParamter(lng病人ID, lng主页ID, strObject, strParam)
            If mfrmChildQuestion.CurrentPatient Then
                Call mfrmChildQuestion.RefreshData(GetChildPatient(mintIndex).Depts, mrsCondition, mblnAuditEnter)
            End If
        End If
    End If
End Sub

Private Sub mfrmChildPatientIn_StatusChanged()
     Call ExecuteCommand("刷新状态")
End Sub

Private Sub mfrmChildQuestion_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmChildQuestion_AfterDeleteQuestion(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    '
    Call ExecuteCommand("刷新指定病人", lng病人ID, lng主页ID)
End Sub

Private Sub mfrmChildQuestion_AfterQuestionType(ByVal blnQuestionType As Boolean)
    'blnQuestionType=True 院级反馈 =Flase 科级反馈
    If blnQuestionType Then
        If ObjPtr(dkpMain.Panes(2)) > 0 Then
            dkpMain.Panes(3).Title = "院级问题反馈"
        End If
    Else
        If ObjPtr(dkpMain.Panes(2)) > 0 Then
            dkpMain.Panes(3).Title = "科级问题反馈"
        End If
    End If
End Sub

Private Sub mfrmChildQuestion_AfterSaveQuestion(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    '
    Call ExecuteCommand("刷新指定病人", lng病人ID, lng主页ID)
End Sub

Private Sub mfrmChildQuestion_LocationDocument(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal byt反馈对象 As Byte, ByVal lng文件id As Long, ByVal lng医嘱id As Long, ByVal lng科室ID As Long)
    
    '根据信息定位到指定病人的指定病案资料上去
    Dim rs          As ADODB.Recordset
    Dim rsTmp       As ADODB.Recordset
'    lng病人id , lng主页ID
    On Error GoTo errHand
    Set rs = gclsPackage.GetDocumentLocation(lng病人ID, lng主页ID)
    If Not rs.BOF Then
        '加载数据
        gstrSQL = "select b.编码 || '-' || b.名称 from 病案主页 a, 部门表 b where a.出院科室id=b.id And 病人ID = [1] and 主页ID = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, lng主页ID)
        If Not (rsTmp.EOF Or rsTmp.BOF) Then
            If zlCommFun.NVL(rs("状态").Value, 0) = 0 Then
                mintIndex = 1
            ElseIf zlCommFun.NVL(rs("病案状态值").Value, 0) = 5 Then
                mintIndex = 0
            ElseIf zlCommFun.NVL(rs("病案状态值").Value, 0) = 1 Or zlCommFun.NVL(rs("病案状态值").Value, 0) = 2 Then
                mintIndex = 0
            Else
                If zlCommFun.NVL(rs("病案状态值").Value, 0) > 10 Then
                    mintIndex = 1
                Else
                    mintIndex = 0
                End If
            End If
            If GetChildPatient(mintIndex).cboDept.Text <> rsTmp.Fields(0) Then
                Call GetChildPatient(mintIndex).cboDeptRefresh(rsTmp.Fields(0))
            End If
        End If
        
        If zlCommFun.NVL(rs("状态").Value, 0) = 0 Then
                
            If tbcTask.Item(1).Selected = False And tbcTask.Item(1).Visible Then tbcTask.Item(1).Selected = True
            Call mfrmChildPatientIn.zlLocationDocument(lng病人ID, lng主页ID, byt反馈对象, lng文件id & "," & lng医嘱id & "," & lng科室ID)
        
        Else
        
            Select Case zlCommFun.NVL(rs("病案状态值").Value, 0)
                Case 1              '未提交病人状态
                
                Case 10, 2           '接收状态
                    If tbcTask.Item(0).Selected = False And tbcTask.Item(0).Visible Then tbcTask.Item(0).Selected = True
                    Call mfrmChildPatientAduit.zlLocationDocument(lng病人ID, lng主页ID, byt反馈对象, lng文件id & "," & lng医嘱id & "," & lng科室ID)
                Case 3, 4       '审查状态
                    
                    If tbcTask.Item(0).Selected = False And tbcTask.Item(0).Visible Then tbcTask.Item(0).Selected = True
                    Call mfrmChildPatientAduit.zlLocationDocument(lng病人ID, lng主页ID, byt反馈对象, lng文件id & "," & lng医嘱id & "," & lng科室ID)
                    
                Case 5              '归档状态
                    
                    If tbcTask.Item(0).Selected = False And tbcTask.Item(0).Visible Then tbcTask.Item(0).Selected = True
                    Call mfrmChildPatientAduit.zlLocationDocument(lng病人ID, lng主页ID, byt反馈对象, lng文件id & "," & lng医嘱id & "," & lng科室ID)
                
            End Select
            
        End If
                
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'自定义过程或函数
'######################################################################################################################
Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
        Case 0
            tbcTask.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    End Select
End Sub

Private Sub tbcTask_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

    mintIndex = Item.Index
    Call CountSelected
    
    If mintIndex = 0 Then
        Call GetChildPatient(mintIndex).InitData(Me.ShowDept)
        With GetChildPatient(mintIndex).VsfBody
            mfrmChildQuestion.AllowModify = Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 3 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 4 Or Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 6
        End With
        
    Else
        mfrmChildQuestion.AllowModify = True
         Call GetChildPatient(mintIndex).InitData(False)
    End If
    Call ExecuteCommand("读取病人病案")
    Call ExecuteCommand("刷新状态")
    
End Sub

Private Sub txtLocation_GotFocus()
    Call zlControl.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtLocation.Text = "" Then Exit Sub
        Call GetChildPatient(mintIndex).zlLocationPatient(2, mstrFindKey, CheckSpecialSign(txtLocation.Text), , , , , mstrFindDeal)
        
        Call LocationObj(txtLocation)
        Call CountSelected
    Else
        If InStr(":：;；?？''||", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub


'检测长度是否超过长度(字节数)
Private Function ChkStrUniCode(mStr As String, mLen As Long) As String
    Dim strL        As String
On Error GoTo ErrH
    mStr = ConvertString(mStr)
    If mLen <= 0 Then
        ChkStrUniCode = mStr
        Exit Function
    Else
        strL = StrConv(mStr, vbFromUnicode)
        strL = LeftB(strL, mLen)
        ChkStrUniCode = StrConv(strL, vbUnicode)
    End If
    Exit Function
ErrH:
    Err.Clear
    ChkStrUniCode = ""
    Exit Function
End Function

'检查是否存在特殊符号(放回处理值)
Private Function CheckSpecialSign(ByVal mStr As String) As String
    Dim i As Integer
    If mStr = "" Then Exit Function
    If InStrRev(mStr, "'", -1) > 0 Then
        CheckSpecialSign = Replace(mStr, "'", "''")
    Else
        CheckSpecialSign = mStr
    End If
End Function

'==============================================================================
'=功能： 查看首页
'==============================================================================
Private Sub RecordLook()
    
    On Error GoTo ErrH
    With GetChildPatient(mintIndex).VsfBody
        If .Row < 1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("病人id"))) = 0 Then GoTo ErrH
        Call frmArchiveView.ShowArchive(Me, Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))), False)
        
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'''Private Sub 调用_Click() '测试独立电子病案审查评分模块
'''    '初始化 zlRichEPR.clsChildQuestion 接口类
'''    Dim clsChildQuestion As zlRichEPR.clsChildQuestion
'''    Set clsChildQuestion = New zlRichEPR.clsChildQuestion
'''
'''    Call clsChildQuestion.zlOpenQuestion(Me, 727848, 1)
'''    Set clsChildQuestion = Nothing
'''End Sub
