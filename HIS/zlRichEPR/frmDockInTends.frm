VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDockInTends 
   BorderStyle     =   0  'None
   Caption         =   "护理记录管理"
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeSuiteControls.TabControl tbcThis 
      Height          =   5115
      Left            =   150
      TabIndex        =   0
      Top             =   750
      Width           =   7335
      _Version        =   589884
      _ExtentX        =   12938
      _ExtentY        =   9022
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsTools 
      Left            =   150
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDockInTends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mstrPrivs As String                             '当前使用者对本程序(1255)的权限串
Private mblnSearch As Boolean                           '当前使用者是否具备病历检索(1273)权
Private mlngPatiId As Long                              '病人id
Private mlngPageId As Long                              '主页id
Private mlngDeptId As Long                              '当前操作科室id，如病人科室和当前科室不一致，则不能操作归档外的功能
Private mblnEdit As Boolean                             '是否允许操作，通常由上级程序根据当前操作科室是否当前病人病区决定。
Private mblnDoctorStation As Boolean
Private mbytFontSize As Byte                            '字体大小0-9号字体,1-12号字体
Private WithEvents mfrmDockInTendFile As frmDockInTendFile
Attribute mfrmDockInTendFile.VB_VarHelpID = -1
Private WithEvents mfrmDockInTendData As frmDockInTendData
Attribute mfrmDockInTendData.VB_VarHelpID = -1
Private WithEvents mfrmDockInTendEPR As frmDockInTendEPR
Attribute mfrmDockInTendEPR.VB_VarHelpID = -1

Private mcbsThis As Object          'CommandBar控件
Private eMySignLevel As EPRSignLevelEnum
Public Event Activate()

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-18 15:16
    '问题:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小(对于模块已经加载调用)
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-18 15:16
    '问题:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim objCtrl As Control
    Dim bytSize As Byte
    
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    '护理文件
    Call mfrmDockInTendFile.SetFontSize(bytSize)
    '护理数据
    Call mfrmDockInTendData.SetFontSize(bytSize)
    '护理病历
    Call mfrmDockInTendEPR.SetFontSize(bytSize)
    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
            objCtrl.PaintManager.Layout = xtpTabLayoutAutoSize
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


Private Sub cbsTools_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbsTools_Resize()
    Call Form_Resize
End Sub

Private Sub cbsTools_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlUpdateCommandBars(Control)
End Sub

'######################################################################################################################
Private Sub Form_Activate()
    RaiseEvent Activate
End Sub

Private Sub Form_Load()

    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "基本") > 0)
    mstrPrivs = GetPrivFunc(glngSys, 1255)

    mlngPatiId = -1
    mlngPageId = -1
    
    '------------------------------------------
    '数据选卡设置
    With Me.tbcThis
        
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
            .Position = xtpTabPositionRight
            
        End With
        
        Set mfrmDockInTendFile = New frmDockInTendFile
        Call mfrmDockInTendFile.InitData(Me, mstrPrivs)
        
        Set mfrmDockInTendData = New frmDockInTendData
        Call mfrmDockInTendData.InitData(Me, mstrPrivs)
        
        Set mfrmDockInTendEPR = New frmDockInTendEPR
        Call mfrmDockInTendEPR.InitData(mstrPrivs)
        
        .InsertItem(0, "护理文件", mfrmDockInTendFile.hWnd, 0).Tag = "_护理文件"
        .InsertItem(1, "护理记录", mfrmDockInTendData.hWnd, 0).Tag = "_护理数据"
        .InsertItem(2, "护理病历", mfrmDockInTendEPR.hWnd, 0).Tag = "_护理病历"
        
        .Item(0).Selected = True
    End With
    
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
    Set cbsTools.Icons = zlCommFun.GetPubIcons
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Call cbsTools.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    If cbsTools(1).Controls.Count = 0 Then lngTop = 0
    tbcThis.Move lngLeft, lngTop, Me.ScaleWidth, Me.ScaleHeight - lngTop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmDockInTendFile Is Nothing Then Unload mfrmDockInTendFile
    If Not mfrmDockInTendData Is Nothing Then Unload mfrmDockInTendData
    If Not mfrmDockInTendEPR Is Nothing Then Unload mfrmDockInTendEPR
    Set mfrmDockInTendFile = Nothing
    Set mfrmDockInTendData = Nothing
    Set mfrmDockInTendEPR = Nothing
    Set mcbsThis = Nothing
End Sub


'------------------------------------------------------------
'以下为公共方法
'------------------------------------------------------------
Public Sub zlDefCommandBars(ByVal cbsThis As Object, Optional ByVal blnChildToolBar As Boolean = False)
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    eMySignLevel = GetUserSignLevel(glngUserId, , mlngPatiId, mlngPageId) '获取当前用户签名级别
    Set mcbsThis = cbsThis
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '特殊情况:放在第一个
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开(&O)…", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        '放在输出到Excel之后
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增记录(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintDayDetail, "批量录入(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改记录(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除记录(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审阅病历(&U)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10, "护理归档(&R)"): cbrControl.Parameter = "护理归档": cbrControl.BeginGroup = True
        cbrControl.IconId = conMenu_Edit_Archive
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnArchive, "护理撤档(&U)")
        cbrControl.IconId = conMenu_Edit_Archive
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "病历归档(&R)"): cbrControl.Parameter = "病历归档"
        cbrControl.IconId = conMenu_Edit_Archive
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "体温作图(&G)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "记录签名(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消签名(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "取消打印(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "历史版本(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Sort, "病历排序(&S)"): cbrControl.BeginGroup = True
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
    
    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", cbrMenuBar.Index, False)
        cbrMenuBar.ID = conMenu_ToolPopup
    End If
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "病历质量监测(&M)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "病人病历检索(&S)")
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
        'Set cbrControl = .Find(, conMenu_File_Preview) '从预览按钮之后开始加入
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "体温", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "取消打印(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "签名", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审阅", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10, "归档", cbrControl.Index + 1)
        cbrControl.IconId = conMenu_Edit_Archive
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_UnArchive, "撤档", cbrControl.Index + 1)
        cbrControl.IconId = conMenu_Edit_Archive
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "归档", cbrControl.Index + 1)
        cbrControl.IconId = conMenu_Edit_Archive
        
        '特殊情况:放在第一个
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        For Each cbrControl In cbrToolBar.Controls
            cbrControl.STYLE = xtpButtonIconAndCaption
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
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit
        .Add FCONTROL, Asc("G"), conMenu_Edit_MarkMap
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F11, conMenu_Tool_Option
    End With
    
    '设置不常用命令
    '-----------------------------------------------------
    With cbsThis.Options
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet
        Call zlPrintSet
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Option                            '护理选项
        If Not CreateBodyEditor Then Exit Sub
        
        If gobjBodyEditor.GetCaseTendBodyPara.ShowPara(Me, mstrPrivs) Then
            Call mfrmDockInTendFile.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlExecuteCommandBars(Control)
        Case "_护理数据"
            Call mfrmDockInTendData.zlExecuteCommandBars(Control)
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlExecuteCommandBars(Control)
        End Select
    End Select
    
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open
            
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print
    
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
    Case conMenu_Edit_NoPrint
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Enabled = False
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Control.Enabled = False
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_ExportToXML
    
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_RowPrint
    
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem

        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintDayDetail    '批量录入

        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
        
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
                
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
                
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search

        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Control.Visible = False
            Control.Enabled = False
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Sign

        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
        
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
            
        Case "_护理病历"
            Control.Visible = False
            Control.Enabled = False
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_SignEarse
    
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
        
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
            
        Case "_护理病历"
            Control.Visible = False
            Control.Enabled = False
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit

        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
        
            Control.Visible = False
            Control.Enabled = False
            
        Case "_护理病历"
            
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10

        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
            
        Case "_护理病历"
            Control.Visible = False
            Control.Enabled = False
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10 + 1

        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_UnArchive

        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Control.Visible = False
            Control.Enabled = False
        End Select
    Case conMenu_Tool_SignVerify
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理病历"
            Control.Visible = False
            Control.Enabled = False
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup

        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MarkMap
    
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Call mfrmDockInTendData.zlUpdateCommandBars(Control)
        Case "_护理病历"
            Control.Visible = False
            Control.Enabled = False
        End Select
                
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Monitor
    
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search
    
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Sort
    
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理数据"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理病历"
            Call mfrmDockInTendEPR.zlUpdateCommandBars(Control)
        End Select

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Save
    
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理病历"
            Control.Visible = False
            Control.Enabled = False
        End Select
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle
    
        Select Case tbcThis.Selected.Tag
        Case "_护理文件"
            Call mfrmDockInTendFile.zlUpdateCommandBars(Control)
        Case "_护理数据"
            Control.Visible = False
            Control.Enabled = False
        Case "_护理病历"
            Control.Visible = False
            Control.Enabled = False
        End Select
    End Select
End Sub

'------------------------------------------------------------
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnForce As Boolean, Optional ByVal blnDoctorStation As Boolean, Optional ByVal blnSeekCase As Boolean) As Long
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rs As New ADODB.Recordset
    mlngDeptId = lngDeptId: mblnEdit = blnEdit

    mlngPatiId = lngPatiID: mlngPageId = lngPageId
    
    mblnDoctorStation = blnDoctorStation
    mblnMoved_HL = False
        
    If mlngPatiId <> 0 Then
        gstrSQL = "Select 数据转出 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "判断数据是否转出", mlngPatiId, mlngPageId)
        mblnMoved_HL = NVL(rs!数据转出, 0) <> 0
    End If
    Call mfrmDockInTendFile.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
    Call mfrmDockInTendData.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
    Call mfrmDockInTendEPR.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, , mblnMoved_HL)
    
    
End Function
    
Public Sub zlLocateData(ByVal intType As Integer)
    tbcThis.Item(intType).Selected = True
End Sub

Private Sub mfrmDockInTendData_AfterArchiveChanged(ByVal blnArchived As Boolean)
    mfrmDockInTendFile.TendArchive = blnArchived
End Sub

Private Sub mfrmDockInTendData_AfterDataChanged()
    Call mfrmDockInTendFile.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
End Sub

Private Sub mfrmDockInTendData_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    Select Case Button
    Case 2

        If mcbsThis.ActiveMenuBar.Controls(3).Visible = False Then Exit Sub

        Set cbrMenuBar = mcbsThis.ActiveMenuBar.Controls(3)
        Set cbrPopupBar = mcbsThis.Add("弹出菜单", xtpBarPopup)
        For Each cbrControl In cbrMenuBar.CommandBar.Controls
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
            cbrPopupItem.Parameter = cbrControl.Parameter
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
            cbrPopupItem.IconId = cbrControl.IconId
        Next
        cbrPopupBar.ShowPopup

    End Select
End Sub

Private Sub mfrmDockInTendData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objTmp As CommandBarControl

    Set objTmp = mcbsThis.FindControl(, conMenu_Edit_Modify)
    If Not (objTmp Is Nothing) Then
        If objTmp.Enabled And objTmp.Visible Then
            Call zlExecuteCommandBars(objTmp)
        End If
    End If

End Sub

Private Sub mfrmDockInTendEPR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    Select Case Button
    Case 2

        If mcbsThis.ActiveMenuBar.Controls(3).Visible = False Then Exit Sub

        Set cbrMenuBar = mcbsThis.ActiveMenuBar.Controls(3)
        Set cbrPopupBar = mcbsThis.Add("弹出菜单", xtpBarPopup)
        For Each cbrControl In cbrMenuBar.CommandBar.Controls
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
            cbrPopupItem.Parameter = cbrControl.Parameter
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
            cbrPopupItem.IconId = cbrControl.IconId
        Next
        cbrPopupBar.ShowPopup

    End Select
End Sub

Private Sub mfrmDockInTendFile_AfterDataChanged()
    Call mfrmDockInTendData.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
    Call mfrmDockInTendFile.RefreshData(mlngPatiId, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit)
End Sub
