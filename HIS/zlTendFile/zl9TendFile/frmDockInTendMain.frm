VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmDockInTendMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeCommandBars.CommandBars cbsTools 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "frmDockInTendMain.frx":0000
      Left            =   390
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDockInTendMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mblnFirst As Boolean
Private mstrPrivs As String                             '当前使用者对本程序(1255)的权限串
Private mblnSearch As Boolean                           '当前使用者是否具备病历检索(1273)权
Private mlngPatiID As Long                              '病人id
Private mlngPageId As Long                              '主页id
Private mintBaby As Integer
Private mlngDeptId As Long                              '当前操作科室id，如病人科室和当前科室不一致，则不能操作归档外的功能
Private mblnEdit As Boolean                             '是否允许操作，通常由上级程序根据当前操作科室是否当前病人病区决定。
Private mblnDoctorStation As Boolean

Private WithEvents mfrmDockInTend_TendList As frmDockInTend_TendList
Attribute mfrmDockInTend_TendList.VB_VarHelpID = -1
Private WithEvents mfrmDockInTend_Data As frmDockInTend_Data
Attribute mfrmDockInTend_Data.VB_VarHelpID = -1

Private mcbsThis As Object          'CommandBar控件
Private cbrControl As CommandBarControl
Private cbrMenuBar As CommandBarPopup
Private cbrToolBar As CommandBar
Private rsTemp As New ADODB.Recordset
Private mintPageSel As Integer
Private mbytFontSize As Byte

Private Enum enmSEL
    护理
    病历
End Enum

Public Event Activate()
Public Event RefreshPrompt(ByVal strInfo As String, ByVal blnImportant As Boolean)
Public Event StartTimer(ByVal blnStart As Boolean)

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-19 15:16
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize(False)
End Sub

Private Sub ReSetFontSize(Optional ByVal blnStart As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小(对于模块已经加载调用)
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-19 15:16
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim objCtrl As Control
    Dim bytSize As Byte
    
    If mlngPatiID = 0 Then Exit Sub
    
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    Call mfrmDockInTend_Data.SetFontSize(bytSize)
    Call mfrmDockInTend_TendList.SetFontSize(bytSize)

    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        End Select
    Next
End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = mfrmDockInTend_TendList.hWnd
    Case 2
        Item.Handle = mfrmDockInTend_Data.hWnd
    End Select
End Sub

Private Sub cbsTools_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbsTools_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlUpdateCommandBars(Control)
End Sub

Private Sub Form_Activate()
    If mblnFirst Then
        '53330,刘鹏飞,2012-09-04,取消该段代码，如果医生站存在外挂Exe执行本段代码程序报错并卡死
'        mfrmDockInTend_TendList.Show
'        mfrmDockInTend_Data.Show
        mblnFirst = False
    End If
    
    RaiseEvent Activate
End Sub

Private Sub InitDOCK()
    Dim objPane As Pane
    With DkpMain
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.CloseGroupOnButtonClick = True
        .Options.HideClient = True
        .SetCommandBars cbsTools
        
        Set objPane = .CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "文件列表": objPane.Options = PaneNoCaption
        Set objPane = .CreatePane(2, 500, 500, DockRightOf, objPane): objPane.Title = "数据页面": objPane.Options = PaneNoCaption
    End With
End Sub

Private Sub InitCommandBar()
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsTools.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsTools.VisualTheme = xtpThemeOffice2003
    cbsTools.EnableCustomization False
    Set cbsTools.Icons = ZLCommFun.GetPubIcons
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "基本") > 0)
    mstrPrivs = GetPrivFunc(glngSys, 1255)
    '加载窗体
    Set mfrmDockInTend_TendList = New frmDockInTend_TendList
    Call mfrmDockInTend_TendList.InitData(Me, mstrPrivs)
    Load mfrmDockInTend_TendList
    Set mfrmDockInTend_Data = New frmDockInTend_Data
    Call mfrmDockInTend_Data.InitData(Me, mstrPrivs)
    Load mfrmDockInTend_Data
    
    Call InitDOCK
    Call InitCommandBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmDockInTend_TendList Is Nothing Then Unload mfrmDockInTend_TendList
    If Not mfrmDockInTend_Data Is Nothing Then Unload mfrmDockInTend_Data
End Sub


'------------------------------------------------------------
'以下为公共方法
'------------------------------------------------------------
Public Sub zlDefCommandBars(ByVal cbsThis As Object, Optional ByVal blnChildToolBar As Boolean = False)
    '-----------------------------------------------------
    Set mcbsThis = cbsThis
    Set cbsThis.Icons = ZLCommFun.GetPubIcons
    
    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '特殊情况:放在第一个
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开(&O)…", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        '放在输出到Excel之后
        '51588,2012-12-12,刘鹏飞,护理文件添加批量打印
        Set cbrControl = .Find(, conMenu_File_Excel)
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print * 100# + 1, "批量打印(&L)…", cbrControl.Index + 1)
        cbrControl.IconId = conMenu_File_Print
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Find(, conMenu_File_Excel)
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出为XML文件(&L)…", cbrControl.Index + 1)
        
        '放在导出为XML文件之后
        Set cbrControl = .Add(xtpControlButton, conMenu_File_RowPrint, "列表打印(&T)", cbrControl.Index + 1): cbrControl.BeginGroup = True
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FileMan, "文件管理(&N)")
    
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintDayDetail, "批量录入(&B)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve, "曲线编辑(&Q)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CurveTable, "表格编辑(&T)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_Show, "曲线显示(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surgery_Edit, "手术/分娩设置(&F)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "记录签名(&S)"): cbrControl.BeginGroup = True
        '51589:刘鹏飞,2013-03-01,添加交班签名
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignShiftExchange, "交班签名(&K)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消签名(&E)"): cbrControl.IconId = 229
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditAffirm, "上级审签(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditCancel, "取消审签(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Billing, "数据录入(&E)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10, "护理归档(&R)"): cbrControl.Parameter = "护理归档": cbrControl.BeginGroup = True
        cbrControl.IconId = conMenu_Edit_Archive
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnArchive, "护理撤档(&U)")
    End With
    
    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", cbrMenuBar.Index, False)
        cbrControl.ID = conMenu_ToolPopup
    End If
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Option, "护理选项(&O)"): cbrControl.BeginGroup = True
        cbrControl.IconId = conMenu_File_Parameter
    End With
    
    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    If blnChildToolBar Then
        cbsTools.DeleteAll
        Set cbrToolBar = cbsTools.Add("护理工具栏", xtpBarTop)
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    Else
        Set cbrToolBar = cbsThis(2)
        For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
            If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
                Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
            End If
        Next
    End If
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FileMan, "文件", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve, "曲线", cbrControl.Index + 1): cbrControl.ToolTipText = "体温曲线编辑": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CurveTable, "表格", cbrControl.Index + 1): cbrControl.ToolTipText = "体温表格编辑"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_Show, "显示", cbrControl.Index + 1): cbrControl.ToolTipText = "设置曲线显示"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surgery_Edit, "手术", cbrControl.Index + 1): cbrControl.ToolTipText = "设置手术/分娩"
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "签名", cbrControl.Index + 1): cbrControl.BeginGroup = True
        '51589:刘鹏飞,2013-03-01,添加交班签名
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignShiftExchange, "交班签名", cbrControl.Index + 1): cbrControl.ToolTipText = "交接班签名"
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignAuditAffirm, "审签", cbrControl.Index + 1)
       ' Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Billing, "数据", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10, "归档", cbrControl.Index + 1): cbrControl.BeginGroup = True
        cbrControl.IconId = conMenu_Edit_Archive
        
        '特殊情况:放在第一个
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        For Each cbrControl In cbrToolBar.Controls
            cbrControl.Style = xtpButtonIconAndCaption
        Next
    End With
    Call cbsTools.RecalcLayout
    
    '命令的快键绑定
    '-----------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("B"), conMenu_File_PrintDayDetail
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("E"), conMenu_Edit_Billing
        .Add FCONTROL, Asc("L"), conMenu_File_Print * 100# + 1
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F11, conMenu_Tool_Option
    End With
    
    '设置不常用命令
    '-----------------------------------------------------
    With cbsThis.Options
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngFormat As Long, lng序号 As Long
    
    Select Case Control.ID
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Tool_Option '护理选项
        Call mfrmDockInTend_TendList.zlExecuteCommandBars(Control)
    Case conMenu_Edit_FileMan
        '得到新编辑的文件的格式ID,后面根据格式ID定位最后一份文件
        If frmNurseFileMan.ShowEditor(mlngPatiID, mlngPageId, mintBaby, mstrPrivs, False, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)), lngFormat, lng序号) Then
            Call mfrmDockInTend_TendList.RefreshData(mlngPatiID, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, lngFormat, lng序号)
        End If
    Case Else
'        If mfrmDockInTend_Data.tbcData.Selected.Index = 2 Then   '数据页面
'            Call mfrmDockInTend_Data.zlExecuteCommandBars(Control)
'        Else
            Call mfrmDockInTend_TendList.zlExecuteCommandBars(Control)
'        End If
    End Select
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub

    Select Case Control.ID
    Case conMenu_Help_Help
    
    Case conMenu_Tool_Option
        Call mfrmDockInTend_TendList.zlUpdateCommandBars(Control)
    Case Else
        Call mfrmDockInTend_TendList.zlUpdateCommandBars(Control)
    End Select
End Sub

'------------------------------------------------------------
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnForce As Boolean, Optional ByVal blnDoctorStation As Boolean, Optional ByVal blnSeekCase As Boolean, Optional ByVal intCurveReSize As Integer = 0) As Long
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rs As New ADODB.Recordset
    mlngDeptId = lngDeptID: mblnEdit = blnEdit
    mlngPatiID = lngPatiID: mlngPageId = lngPageId
    mblnDoctorStation = blnDoctorStation
    Call mfrmDockInTend_TendList.RefreshData(mlngPatiID, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, , , intCurveReSize, False)
End Function
    
Public Sub zlLocateData(ByVal intType As Integer)
'    tbcData.Item(intType).Selected = True
End Sub

Private Sub mfrmDockInTend_Data_Activate()
    On Error Resume Next
    Me.SetFocus
End Sub

Private Sub mfrmDockInTend_Data_AfterDataChanged(ByVal blnChange As Boolean)
    Call mfrmDockInTend_TendList.SetChange(blnChange)
End Sub

Private Sub mfrmDockInTend_Data_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    RaiseEvent RefreshPrompt(strInfo, blnImportant)
    Call mfrmDockInTend_TendList.SetState(blnSign, blnArchive)
End Sub

Private Sub mfrmDockInTend_Data_ISChartArchive(ByVal blnArchive As Boolean)
    Call mfrmDockInTend_TendList.SetState(True, blnArchive)
End Sub

Private Sub mfrmDockInTend_Data_StartTimer(ByVal blnStart As Boolean)
    Call mfrmDockInTend_TendList.StartTimer(blnStart)
End Sub

Private Sub mfrmDockInTend_Data_zlRefreshViewFile()
    Call mfrmDockInTend_TendList.zlRefreshViewFile
End Sub

Private Sub mfrmDockInTend_TendList_Activate()
'    On Error Resume Next
'    Me.SetFocus
End Sub

Private Sub mfrmDockInTend_TendList_ArchiveDocument(blnOK As Boolean)
    Call mfrmDockInTend_Data.zlArchiveDocument(blnOK)
End Sub


Private Sub mfrmDockInTend_TendList_BulkPrintDocument(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal intBaby As Integer)
    Call mfrmDockInTend_Data.BulkPrintDocument(lngPatiID, lngPageId, lngDeptID, intBaby)
End Sub

Private Sub mfrmDockInTend_TendList_PrintTendFile(ByVal bytKind As Byte, ByVal bytMode As Byte)
    Call mfrmDockInTend_Data.zlPrintTendFile(bytKind, bytMode)
End Sub

Private Sub mfrmDockInTend_TendList_SaveDocument(blnSave As Boolean)
    Call mfrmDockInTend_Data.zlSaveDocument(blnSave)
End Sub

Private Sub mfrmDockInTend_TendList_ShowData(intBaby As Integer, lngFile As Long, lngDept As Long, bytSel As Byte, ByVal intCurveReSize As Integer)
    mintBaby = intBaby
    Call mfrmDockInTend_Data.zlRefreshTend(mlngPatiID, mlngPageId, intBaby, lngDept, mblnEdit, mblnDoctorStation, lngFile, bytSel, intCurveReSize)
End Sub

Private Sub mfrmDockInTend_TendList_SignDocument(blnOK As Boolean, blnVerify As Boolean, blnExchange As Boolean)
    Call mfrmDockInTend_Data.zlSignDocument(blnOK, blnVerify, blnExchange)
End Sub

Private Sub mfrmDockInTend_TendList_SignMarker()
    Call mfrmDockInTend_Data.SignMarker
End Sub

Private Sub mfrmDockInTend_TendList_ViewAnimalHeat(strPara As String, bytMode As Byte, strPrivs As String, ByVal bytSize As Byte)
    Call mfrmDockInTend_Data.zlViewAnimalHeat(strPara, bytMode, strPrivs, bytSize)
End Sub

Private Sub mfrmDockInTend_TendList_ViewCaveData(ByVal intDataEditor As Integer)
    Call mfrmDockInTend_Data.zlViewCaveData(intDataEditor)
End Sub

Private Sub mfrmDockInTend_TendList_ViewFile(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, ByVal blnEdit As Boolean, ByVal bytSize As Byte)
    Call mfrmDockInTend_Data.zlViewFile(lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, blnChildForm, strPrivs, blnEdit, bytSize)
End Sub


Private Sub mfrmDockInTend_TendList_Viewpartogram(strPara As String, bytMode As Byte, strPrivs As String, ByVal bytSize As Byte)
    Call mfrmDockInTend_Data.zlViewpartogram(strPara, bytMode, strPrivs, bytSize)
End Sub

Private Sub mfrmDockInTend_TendList_ViewpartogramEditor(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal intBaby As Integer, ByVal strPrivs As String, ByVal bytSize As Byte)
    Call mfrmDockInTend_Data.zlViewpartogramEditor(lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, strPrivs, bytSize)
End Sub

Private Sub mfrmDockInTend_TendList_ViewReSetFontSize(ByVal intSEL As Integer, ByVal bytSize As Byte)
     Call mfrmDockInTend_Data.ViewReSetFontSize(intSEL, bytSize)
End Sub
