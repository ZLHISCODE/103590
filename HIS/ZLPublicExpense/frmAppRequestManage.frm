VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "codejock.dockingpane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppRequestManage 
   Caption         =   "预约登记管理"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "frmAppRequestManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   11745
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7995
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   635
      SimpleText      =   $"frmAppRequestManage.frx":058A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppRequestManage.frx":05D1
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15637
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
      Left            =   435
      Top             =   525
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmAppRequestManage.frx":0E65
      Left            =   975
      Top             =   585
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAppRequestManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As frmAppRequestMain
Private mfrmFilter As frmAppRequestFilter
Private mlngFaceBackColor As Long

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnable As Boolean
    Select Case Control.ID
    Case conMenu_Edit_CancelRequest
        blnEnable = True
        If mfrmMain.rptMain.SelectedRows.Count = 0 Then
            blnEnable = False
        Else
            If mfrmMain.rptMain.SelectedRows.Row(0).Record Is Nothing Then blnEnable = False
        End If
        Control.Enabled = blnEnable
    Case conMenu_Edit_ViewRequest
        blnEnable = True
        If mfrmMain.rptMain.SelectedRows.Count = 0 Then
            blnEnable = False
        Else
            If mfrmMain.rptMain.SelectedRows.Row(0).Record Is Nothing Then blnEnable = False
        End If
        Control.Enabled = blnEnable
    End Select
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandle
    Call DefMainCommandBars
    Call InitPanel '初始化dkpMain
    
    mlngFaceBackColor = cbsThis.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
    Me.BackColor = mlngFaceBackColor
    RestoreWinState mfrmMain, "frmAppRequestMain"
    RestoreWinState Me, "frmAppRequestManage"
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub zlDataPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    Dim objVsf As VSFlexGrid
    
    Err = 0: On Error GoTo errHandle

    objOut.Title.Text = "预约登记记录清册"
    Set objVsf = gobjControl.RPTCopyToVSF(mfrmMain.rptMain, objVsf)
    Set objOut.Body = objVsf
    
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Now 'Format(sys.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    
    Err = 0: On Error GoTo errHandle
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_View_StatusBar
        Control.Checked = Not Control.Checked
        stbThis.Visible = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        Control.Checked = Not Control.Checked
        cbsThis(2).Visible = Control.Checked
        Set objControl = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Text, , True)
        objControl.Enabled = Control.Checked
        Set objControl = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Size, , True)
        objControl.Enabled = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not Control.Checked
        For Each objControl In cbsThis(2).Controls
            objControl.Style = IIf(Control.Checked, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Control.Checked = Not Control.Checked
        cbsThis.Options.LargeIcons = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_Help_Help: Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call gobjComlib.zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call gobjComlib.zlMailTo(Me.hWnd)
    Case conMenu_Help_About: Call gobjComlib.ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Edit_AppRequest
        frmAppRequestEdit.ShowMe Me
        Call mfrmMain.RefreshData
    Case conMenu_View_Refresh
        Call mfrmMain.RefreshData
    Case conMenu_Edit_CancelRequest
        If mfrmMain.rptMain.SelectedRows.Count = 0 Then Exit Sub
        If mfrmMain.rptMain.SelectedRows.Row(0).Record Is Nothing Then Exit Sub
        Call CancelRequest
        Call mfrmMain.RefreshData
    Case conMenu_Edit_ViewRequest
        If mfrmMain.rptMain.SelectedRows.Count = 0 Then Exit Sub
        If mfrmMain.rptMain.SelectedRows.Row(0).Record Is Nothing Then Exit Sub
        Call frmAppRequestEdit.ReadBill(Me, Val(mfrmMain.rptMain.SelectedRows.Row(0).Record.Tag))
    Case Else
    End Select
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CancelRequest()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = "Select 1 From 病人服务信息记录 Where ID=[1] And 处理时间 Is Null"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mfrmMain.rptMain.SelectedRows.Row(0).Record.Tag))
    If rsTemp.EOF Then
        MsgBox "当前预约登记记录已经被处理,无法取消登记!", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("是否确定取消该条预约登记记录?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
    strSQL = "zl_患者服务中心_更新("
    strSQL = strSQL & mfrmMain.rptMain.SelectedRows.Row(0).Record.Tag & ",'"
    strSQL = strSQL & "取消登记','"
    strSQL = strSQL & UserInfo.姓名 & "','"
    strSQL = strSQL & UserInfo.编号 & "',"
    strSQL = strSQL & "Null,"
    strSQL = strSQL & 1 & ")"
    Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_SpecialColorChanged()
    Me.BackColor = cbsThis.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
End Sub

Public Sub RefreshRecord()
    With mfrmMain
        .mbln登记时间 = mfrmFilter.chkDate(0).Value
        .mbln处理时间 = mfrmFilter.chkDate(1).Value
        .mbln显示处理 = mfrmFilter.chkShowSet.Value
        .mdat处理开始 = mfrmFilter.dtpBegin(1).Value
        .mdat处理结束 = mfrmFilter.dtpEnd(1).Value
        .mdat开始时间 = mfrmFilter.dtpBegin(0).Value
        .mdat结束时间 = mfrmFilter.dtpEnd(0).Value
        .mstr处理人 = NeedName(mfrmFilter.cbo处理人.Text)
        .mstr登记人 = NeedName(mfrmFilter.cbo登记人.Text)
        .mbyt复诊方式 = mfrmFilter.cbo复诊方式.ListIndex
    End With
    Call mfrmMain.RefreshData
End Sub


Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    
    Err = 0: On Error GoTo errHandle
    Set cbsThis.Icons = gobjCommFun.GetPubIcons
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
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 '去除菜单栏前缀
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Edit, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_Edit
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AppRequest, "预约登记(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CancelRequest, "取消登记(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ViewRequest, "查看登记(&V)"):  cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
        cbrSubControl.Checked = True
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
        cbrSubControl.Checked = True
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
        cbrSubControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ModifyStyle &H400000, 0 '去除菜单栏前缀
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AppRequest, "预约登记(&A)"): cbrControl.BeginGroup = True
        cbrControl.IconId = 3003
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CancelRequest, "取消登记(&C)")
        cbrControl.IconId = 3004
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll
        .Add FCONTROL, vbKeyC, conMenu_Edit_ClsAll
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    DefMainCommandBars = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitPanel()
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHandle
    Set objPane = dkpMain.CreatePane(1, 230, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    Set mfrmFilter = New frmAppRequestFilter
    objPane.Handle = mfrmFilter.hWnd
    objPane.MaxTrackSize.Width = 265
    objPane.MinTrackSize.Width = 265
    mfrmFilter.SetForm Me
    
    Set objPane = dkpMain.CreatePane(2, 230, 300, DockRightOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    Set mfrmMain = New frmAppRequestMain
    objPane.Handle = mfrmMain.hWnd
    
    With dkpMain
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState mfrmMain, "frmAppRequestMain"
    SaveWinState Me, "frmAppRequestManage"
End Sub
