VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "ZLSQL Trace"
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16320
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   2460
      Top             =   405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   1425
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":0E42
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   960
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mWorkSpace As TabWorkspace
Attribute mWorkSpace.VB_VarHelpID = -1
Private WithEvents mfrmSession As frmSession
Attribute mfrmSession.VB_VarHelpID = -1
Private WithEvents mfrmTrace As frmTrace
Attribute mfrmTrace.VB_VarHelpID = -1

Public Property Let StatusText(ByVal vNewValue As String)
    Dim i As Integer, arrText As Variant
    
    arrText = Split(vNewValue, "|")
    For i = 0 To UBound(arrText)
        Me.cbsMain.StatusBar.SetPaneText i, Split(vNewValue, "|")(i)
    Next
End Property

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    cbsMain.Item(2).Controls.Find(, conMenu_Edit_TraceOff).Enabled = False
    Select Case Control.Id
    Case conMenu_File_Open
        Call OpenLogFile
    Case conMenu_File_CompareExe
        Call SetCompareExe
'    Case conmenu_File_Logout
'        If CheckTraceState = False Then Exit Sub
'        If gcnOracle.State = 1 Then
'            If MsgBox("确实要注销吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
'            gcnOracle.Close
'        End If
'        Set gcnOracle = Nothing
'        Unload Me
'        Call Main
    Case conMenu_File_Exit
        If CheckTraceState = False Then Exit Sub
        Unload Me
    Case conMenu_Help_About
        ShellAbout Me.hWnd, "SQL Trace", vbCrLf & "SQL Trace 跟踪分析工具", Me.Icon.Handle
    Case Else
        If Not Me.ActiveForm Is Nothing Then
            Call Me.ActiveForm.DoCommand(Control.Id)
        End If
    End Select
End Sub

Private Function CheckTraceState() As Boolean
    If Me.ActiveForm.mlngCount > 0 Then
        If MsgBox("存在未停止跟踪的会话，如果不停止，直到他们退出会话才会停止跟踪。" & vbCrLf & "你确定要继续吗？", vbYesNo, "警告") = vbYes Then
            CheckTraceState = True
        End If
    Else
        CheckTraceState = True
    End If
End Function

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '会话子窗体不允许关闭
    If Not cbsMain.TabWorkspace.Selected Is Nothing Then
        If cbsMain.TabWorkspace.Selected.Caption = "会话" Then
            cbsMain.TabWorkspace.flags = xtpWorkspaceHideClose
        Else
            cbsMain.TabWorkspace.flags = 0
        End If
    End If
    If Me.ActiveForm Is Nothing Then Exit Sub

    Select Case Control.Id
'    Case conmenu_File_Logout
'        If gcnOracle.State = 0 Then
'            Control.Caption = "登录(&L)"
'        Else
'            Control.Caption = "注销(&L)"
'        End If
    Case conMenu_File_CompareExe
        Control.Caption = IIf(gstrCompareExe <> "", "对比工具(&E):" & gstrCompareExe, "设置对比工具位置(&E)")
    Case conMenu_Edit_CompareLeft
        Control.Enabled = Me.ActiveForm.GetCommand(Control.Id) And gstrCompareExe <> ""
    Case conMenu_Edit_Compare
        Control.Caption = "当前窗口与 " & IIf(gstrLeft <> "", gobjFile.GetFileName(gstrLeft), "...") & " 对比(&R)"
        Control.Enabled = gstrLeft <> "" And gstrCompareExe <> ""
    Case conMenu_Edit_TraceOff, conMenu_Edit_Trace, conMenu_Edit_Trace_1, conMenu_Edit_Trace_4, conMenu_Edit_Trace_8, conMenu_Edit_Trace_12

        Control.Enabled = Me.ActiveForm.GetCommand(Control.Id)
 
        If Control.Enabled = False Then
            Control.ToolTipText = "当前用户缺少SYS.DBMS_System包的执行权限，无法使用跟踪功能！" & vbCrLf & "请先进行授权,或使用dba用户登录！"
        End If
    Case conMenu_View_Style
        Control.Enabled = Me.ActiveForm.GetCommand(Control.Id)
        If Control.Enabled Then
            Control.IconId = Me.ActiveForm.ViewStyle
        End If
    Case conMenu_View_Style_Report, conMenu_View_Style_Table
        Control.Checked = Me.ActiveForm.ViewStyle = Control.Id
    Case conMenu_View_Find, conMenu_View_FindNext
        Control.Enabled = Me.ActiveForm.GetCommand(Control.Id)
    Case conMenu_View_Filter
        Control.Enabled = Me.ActiveForm.GetCommand(Control.Id)
        If Control.Enabled Then
            Control.Checked = Me.ActiveForm.Filtering
        End If
    Case conMenu_View_SQLPrev, conMenu_View_SQLNext
        Control.Enabled = Me.ActiveForm.GetCommand(Control.Id)
    Case conMenu_View_Refresh
        Control.Enabled = Me.ActiveForm.GetCommand(Control.Id)
    End Select
End Sub

Private Sub SetCompareExe()
    With Me.cdgFile
        .DialogTitle = "设置对比器位置"
        .Filter = "Compare It!  (wincmp3.exe)|wincmp3.exe|Beyond Compare 2  (BC2.exe)|BC2.exe|其他对比工具  (*.exe)|*.exe"
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .InitDir = ""
        .FileName = ""
        .CancelError = True
        On Error GoTo errh
        .ShowOpen
        gstrCompareExe = .FileName
        SaveSetting "ZLSOFT\公共模块\ZLDBATools", "Setting", "CompareExe", .FileName
    End With
errh:
End Sub

Private Sub OpenLogFile()
    Dim frmNew As New frmTrace
    
    With Me.cdgFile
        .DialogTitle = "打开已经解析的Trace文件"
        .Filter = "SQL Trace(*.log)|*.log"
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .InitDir = GetSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "Output", "")
        .FileName = ""
        .CancelError = True
        On Error GoTo errh
        .ShowOpen
        
        SaveSetting "ZLSOFT\公共模块\ZLDBATools", "Setting", "Output", Left(.FileName, Len(.FileName) - Len(.FileTitle))
        frmNew.ShowMe Me, .FileName
    End With
errh:
End Sub

Private Sub MDIForm_Load()
    Dim strVal As String
    
    strVal = GetSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "MainFormState", 0)
    If Val(strVal) = 2 Then
        Me.WindowState = 2
    ElseIf Val(strVal) = 0 Then
        Me.WindowState = 0
        strVal = GetSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "MainFormSize", "")
        If strVal = "" Then
            Me.Left = (Screen.Width - Me.Width) / 2
            Me.Top = (Screen.Height - Me.Height) / 2 - 1000
        Else
            Me.Left = Split(strVal, ",")(0)
            Me.Top = Split(strVal, ",")(1)
            Me.Width = Split(strVal, ",")(2)
            Me.Height = Split(strVal, ",")(3)
        End If
    End If
    
    gstrCompareExe = GetSetting("ZLSOFT\公共模块\ZLDBATools", "Setting", "CompareExe", "")
    If gstrCompareExe = "" Then
        If gobjFile.FileExists(App.path & "\wincmp3.exe") Then
            gstrCompareExe = App.path & "\wincmp3.exe"
        End If
    ElseIf Not gobjFile.FileExists(gstrCompareExe) Then
        gstrCompareExe = ""
    End If
    
    Call InitCommandBar
    
    Set mfrmSession = New frmSession
    mfrmSession.ShowMe Me
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = imgMain.Icons
    
    '菜单定义:包括公共部份
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.Id = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Open, "打开已解析的文件(&O)...")
        Set objControl = .Add(xtpControlButton, conMenu_File_CompareExe, IIf(gstrCompareExe <> "", "对比工具(&E):" & gstrCompareExe, "设置对比工具位置(&E)")): objControl.BeginGroup = True
        'Set objControl = .Add(xtpControlButton, conmenu_File_Logout, "注销(&L)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.Id = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButtonPopup, conMenu_Edit_Trace, "跟踪(&T)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_TraceOff, "停止(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_CompareLeft, "当前窗口作为左侧对比窗口(&L)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compare, "当前窗口与 ... 对比(&R)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.Id = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_View_Style, "格式(&S)")
        objControl.IconId = conMenu_View_Style_Report
        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "查找(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找一下个(&N)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Filter, "筛选(&I)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_SQLPrev, "前条SQL(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_View_SQLNext, "后条SQL(&N)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, XTP_ID_WINDOW_LIST, "窗体(&W)"): objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.Id = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlPopup, conMenu_Edit_Trace, "跟踪")
        objControl.ToolTipText = "跟踪当前选中的会话"
        objControl.Id = conMenu_Edit_Trace
        objControl.IconId = conMenu_Edit_Trace
        Set objControl = .Add(xtpControlButton, conMenu_Edit_TraceOff, "停止")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Open, "打开"): objControl.BeginGroup = True
        objControl.ToolTipText = "直接打开已经解析的Trace文件"
        
        Set objControl = .Add(xtpControlPopup, conMenu_ComparePopup, "对比")
        objControl.ToolTipText = "对比文件的文本内容，对比前需要设置对比工具，如：Compare It等"
        objControl.Id = conMenu_ComparePopup
        objControl.IconId = conMenu_ComparePopup
        Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_View_Style, "格式")
        objControl.ToolTipText = "Trace文件在当前窗口中的查看格式"
        objControl.IconId = conMenu_View_Style_Report
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "查找"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Filter, "筛选"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_SQLPrev, "前条")
        Set objControl = .Add(xtpControlButton, conMenu_View_SQLNext, "后条")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.BeginGroup = True
    End With
    For Each objControl In objBar.Controls
        If objControl.Id <> conMenu_View_SQLPrev And objControl.Id <> conMenu_View_SQLNext Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyO, conMenu_File_Open
        .Add FCONTROL, vbKeyB, conMenu_Edit_Trace_1
        .Add FCONTROL, vbKeyE, conMenu_Edit_TraceOff
        .Add FALT, vbKey1, conMenu_View_Style_Report
        .Add FALT, vbKey2, conMenu_View_Style_Table
        .Add FCONTROL, vbKeyF, conMenu_View_Find
        .Add 0, vbKeyF3, conMenu_View_FindNext
        .Add FCONTROL, vbKeyI, conMenu_View_Filter
        .Add FCONTROL, vbKeyLeft, conMenu_View_SQLPrev
        .Add FCONTROL, vbKeyRight, conMenu_View_SQLNext
        .Add 0, vbKeyF5, conMenu_View_Refresh
    End With
    
    'MDI Tab
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.SetFlags xtpFlagHideMDIButtons, 0
    Set mWorkSpace = cbsMain.ShowTabWorkspace(True)
    cbsMain.TabWorkspace.AutoTheme = False
    cbsMain.TabWorkspace.PaintManager.Appearance = xtpTabAppearanceVisualStudio
    cbsMain.TabWorkspace.PaintManager.Color = xtpTabColorOffice2003
    cbsMain.TabWorkspace.PaintManager.ClientFrame = xtpTabFrameSingleLine
    
    '状态栏
    '-----------------------------------------------------
    cbsMain.StatusBar.Visible = True
    cbsMain.StatusBar.AddPane 1
    cbsMain.StatusBar.SetPaneStyle 1, SBPS_STRETCH
    cbsMain.StatusBar.SetPaneText 1, ""
    cbsMain.StatusBar.AddPane 2
    cbsMain.StatusBar.SetPaneWidth 2, 100
    cbsMain.StatusBar.SetPaneText 2, ""
    cbsMain.StatusBar.AddPane 3
    cbsMain.StatusBar.SetPaneWidth 3, 60
    cbsMain.StatusBar.SetPaneText 3, ""
    cbsMain.StatusBar.IdleText = ""
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    If CommandBar.Parent.Id = conMenu_View_Style Then
        With CommandBar.Controls
            .DeleteAll
            .Add xtpControlButton, conMenu_View_Style_Report, "报告方式(&R)"
            .Add xtpControlButton, conMenu_View_Style_Table, "表格方式(&T)"
        End With
    ElseIf CommandBar.Parent.Id = conMenu_Edit_Trace Then
        With CommandBar.Controls
            .DeleteAll
            .Add xtpControlButton, conMenu_Edit_Trace_1, "跟踪 - 标准(&1)"
            .Add xtpControlButton, conMenu_Edit_Trace_4, "跟踪 - 绑定值(&2)"
            .Add xtpControlButton, conMenu_Edit_Trace_8, "跟踪 - 等待事件(&3)"
            .Add xtpControlButton, conMenu_Edit_Trace_12, "跟踪 - 所有(&4)"
            .Add xtpControlButton, conMenu_Edit_ChangeReg, "修改Trace文件存储路径"
        End With
    ElseIf CommandBar.Parent.Id = conMenu_ComparePopup Then
        With CommandBar.Controls
            .DeleteAll
            .Add xtpControlButton, conMenu_File_CompareExe, IIf(gstrCompareExe <> "", "对比工具(&E):" & gstrCompareExe, "设置对比工具位置(&E)")
            .Add xtpControlButton, conMenu_Edit_CompareLeft, "当前窗口作为左侧对比窗口"
            .Add xtpControlButton, conMenu_Edit_Compare, "当前窗口与 " & IIf(gstrLeft <> "", gobjFile.GetFileName(gstrLeft), "...") & " 对比"
        End With
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mWorkSpace.ItemCount > 1 And Not gcnOracle Is Nothing Then
        If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload gfrmFind
    Set gfrmFind = Nothing
    
    Unload mfrmSession
    Set mfrmSession = Nothing
    
    Set mWorkSpace = Nothing
    
    Set gobjFile = Nothing
    
    If Me.WindowState <> 1 Then
        SaveSetting "ZLSOFT\公共模块\ZLDBATools", "Setting", "MainFormState", Me.WindowState
    End If
    If Me.WindowState = 0 Then
        SaveSetting "ZLSOFT\公共模块\ZLDBATools", "Setting", "MainFormSize", Me.Left & "," & Me.Top & "," & Me.Width & "," & Me.Height
    End If
End Sub

Private Sub mfrmSession_OpenNewFile(ByVal File As String)
    Dim frmNew As New frmTrace
    frmNew.ShowMe Me, File
End Sub

Private Sub mfrmSession_PopSessionMenu()
    Dim objPopup As CommandBar
    
    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        .Add xtpControlButton, conMenu_Edit_Trace_1, "跟踪 - 标准(&1)"
        .Add xtpControlButton, conMenu_Edit_Trace_4, "跟踪 - 绑定值(&2)"
        .Add xtpControlButton, conMenu_Edit_Trace_8, "跟踪 - 等待事件(&3)"
        .Add xtpControlButton, conMenu_Edit_Trace_12, "跟踪 - 所有(&4)"
        .Add(xtpControlButton, conMenu_Edit_TraceOff, "停止(&S)").BeginGroup = True
    End With
    objPopup.ShowPopup
End Sub

Private Sub mfrmSession_UpdateStatus(ByVal strStatus As String)
    Me.StatusText = strStatus
End Sub

Private Sub mfrmTrace_UpdateStatus(ByVal strStatus As String)
    Me.StatusText = strStatus
End Sub

Private Sub mWorkSpace_RClick(ByVal Item As XtremeCommandBars.ITabControlItem)
    Dim objPopup As CommandBar
    
    If Item Is Nothing Then Exit Sub
    Item.Selected = True
    If Item.Caption = "会话" Then Exit Sub
    
    'mWorkSpace.Refresh
    
    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        .Add xtpControlButton, conMenu_View_Style_Report, "报告方式(&R)"
        .Add xtpControlButton, conMenu_View_Style_Table, "表格方式(&T)"
        .Add(xtpControlButton, conMenu_Edit_CompareLeft, "当前窗口作为左侧对比窗口(&L)").BeginGroup = True
        .Add xtpControlButton, conMenu_Edit_Compare, "当前窗口与 " & IIf(gstrLeft <> "", gobjFile.GetFileName(gstrLeft), "...") & " 对比(&R)"
        .Add(xtpControlButton, conMenu_View_Close, "关闭(&C)").BeginGroup = True
    End With
    objPopup.ShowPopup
    
    'mWorkSpace.Refresh
End Sub

