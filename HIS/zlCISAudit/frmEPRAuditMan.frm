VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEPRAuditMan 
   Caption         =   "病历质量查阅"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   Icon            =   "frmEPRAuditMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   1
      Left            =   6615
      ScaleHeight     =   3585
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   1065
      Width           =   3135
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   1830
         Left            =   150
         TabIndex        =   3
         Top             =   255
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   5790
      Index           =   0
      Left            =   135
      ScaleHeight     =   5790
      ScaleWidth      =   2760
      TabIndex        =   1
      Top             =   900
      Width           =   2760
      Begin VB.PictureBox picDate 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1680
         Left            =   300
         ScaleHeight     =   1680
         ScaleWidth      =   2325
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1830
         Width           =   2325
         Begin VB.CommandButton cmdSearch 
            Caption         =   "重新统计(&R)"
            Height          =   350
            Left            =   450
            TabIndex        =   10
            Top             =   900
            Width           =   1605
         End
         Begin MSComCtl2.DTPicker dtpDateTo 
            Height          =   300
            Left            =   450
            TabIndex        =   11
            Top             =   465
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   72548355
            CurrentDate     =   38683
         End
         Begin MSComCtl2.DTPicker dtpDateFrom 
            Height          =   300
            Left            =   450
            TabIndex        =   12
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   72548355
            CurrentDate     =   38683
         End
         Begin VB.Label lblDateFrom 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "从"
            Height          =   180
            Left            =   240
            TabIndex        =   14
            Top             =   180
            Width           =   180
         End
         Begin VB.Label lblDateTo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Height          =   180
            Left            =   240
            TabIndex        =   13
            Top             =   525
            Width           =   180
         End
      End
      Begin VB.PictureBox picKind 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   225
         ScaleHeight     =   1020
         ScaleWidth      =   2325
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   630
         Width           =   2325
         Begin VB.OptionButton optKind 
            Caption         =   "护理病历(&3)"
            Height          =   180
            Index           =   2
            Left            =   420
            TabIndex        =   8
            Top             =   720
            Width           =   1380
         End
         Begin VB.OptionButton optKind 
            Caption         =   "住院病历(&2)"
            Height          =   180
            Index           =   1
            Left            =   420
            TabIndex        =   7
            Top             =   390
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton optKind 
            Caption         =   "门诊病历(&1)"
            Height          =   180
            Index           =   0
            Left            =   420
            TabIndex        =   6
            Top             =   75
            Width           =   1380
         End
      End
      Begin XtremeSuiteControls.TaskPanel tplThis 
         Height          =   4410
         Left            =   180
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   345
         Width           =   2550
         _Version        =   589884
         _ExtentX        =   4498
         _ExtentY        =   7779
         _StockProps     =   64
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRAuditMan.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18203
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
      Left            =   165
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRAuditMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

'变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串

Private mintKind As Integer     '病历种类
Private mstrDateFrom As String  '开始日期
Private mstrDateTo As String    '结束日期
Private mlngMoual As Long


'临时变量
'-----------------------------------------------------
Private cbrControl As CommandBarControl
Private cbrMenuBar As CommandBarPopup
Private cbrToolBar As CommandBar
Private mblnStartUp As Boolean
Private rsTemp As New ADODB.Recordset
Private strSQL As String
Private lngCount As Long
Private lngRow As Long
Private lngCol As Long

Private mfrmEPRAuditPati As frmEPRAuditPati
Private mfrmEPRAuditFile As frmEPRAuditFile
Private mfrmEPRAuditOutline As frmEPRAuditOutline

'######################################################################################################################
Public Function ShowDialog(ByVal frmMain As Object, ByVal lngMoual As Long) As Boolean
    mlngMoual = lngMoual
    Me.Show 1, frmMain
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Select Case tbcPage.Selected.Index
        Case 0
            If Not (mfrmEPRAuditOutline Is Nothing) Then Call mfrmEPRAuditOutline.zlExecuteCommandBars(Control)
        Case 1
            If Not (mfrmEPRAuditPati Is Nothing) Then Call mfrmEPRAuditPati.zlExecuteCommandBars(Control)
        Case 2
            If Not (mfrmEPRAuditFile Is Nothing) Then Call mfrmEPRAuditFile.zlExecuteCommandBars(Control)
        End Select

    Case conMenu_File_Exit
        Unload Me
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
    
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
        
    Case conMenu_View_Refresh
    
        Call cmdSearch_Click
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Jump
        
        If tbcPage.Selected.Index + 1 <= tbcPage.ItemCount - 1 Then
            tbcPage.Item(tbcPage.Selected.Index + 1).Selected = True
        Else
            tbcPage.Item(0).Selected = True
        End If
        
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '中联论坛
    
        Call zlWebForum(Me.hWnd)
        
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        '执行发布到当前模块的报表
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                "开始日期=" & Format(dtpDateFrom.Value, "yyyy-MM-dd"), "结束日期=" & Format(dtpDateTo.Value, "yyyy-MM-dd"), _
                "病历种类=" & IIf(optKind(0).Value, "门诊病历", IIf(optKind(1).Value, "住院病历", "护理病历")) & "|" & IIf(optKind(0).Value, 1, IIf(optKind(1).Value, 2, 4)))
        End If
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()

    Dim lngScaleLeft As Long, lngScaleTop  As Long, lngScaleRight  As Long, lngScaleBottom  As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    With picPane(0)
        .Left = lngScaleLeft
        .Top = lngScaleTop: .Height = lngScaleBottom - .Top
    End With
    
    With picPane(1)
        .Left = picPane(0).Left + picPane(0).Width + 30:
        .Top = lngScaleTop
        .Width = lngScaleRight - .Left
        .Height = lngScaleBottom - .Top
    End With

End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Open
'        With Me.vfgThis
'            Control.Enabled = (Val(.TextMatrix(.Row, 1)) <> 0)
'            If Control.Enabled = False Then Exit Sub
'            For lngCol = 3 To .Cols - 1
'                Control.Enabled = (Val(.TextMatrix(.Row, lngCol)) <> 0)
'                If Control.Enabled Then Exit Sub
'            Next
'        End With
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Select Case tbcPage.Selected.Index
        Case 0
            If Not (mfrmEPRAuditOutline Is Nothing) Then Call mfrmEPRAuditOutline.zlUpdateCommandBars(Control)
        Case 1
            If Not (mfrmEPRAuditPati Is Nothing) Then Call mfrmEPRAuditPati.zlUpdateCommandBars(Control)
        Case 2
            If Not (mfrmEPRAuditFile Is Nothing) Then Call mfrmEPRAuditFile.zlUpdateCommandBars(Control)
        End Select
        
'        Control.Enabled = (Me.vfgThis.Rows > Me.vfgThis.FixedRows + 1)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

'Private Sub chkNoData_Click()
'    Dim blnData As Boolean
'    With Me.vfgThis
'        If Me.chkNoData.Value = vbChecked Then
'            For lngRow = .FixedRows To .Rows - 2
'                .RowHeight(lngRow) = .RowHeightMin
'                .RowHidden(lngRow) = False
'            Next
'        Else
'            For lngRow = .FixedRows To .Rows - 2
'                blnData = False
'                For lngCol = 3 To .Cols - 1
'                    If Val(.TextMatrix(lngRow, lngCol)) <> 0 Then blnData = True: Exit For
'                Next
'                If blnData = False Then
'                    .RowHeight(lngRow) = 0
'                    .RowHidden(lngRow) = True
'                End If
'            Next
'        End If
'    End With
'End Sub

Private Sub cmdSearch_Click()
    Dim strInfo As String
    
    If dtpDateTo.Value - dtpDateFrom.Value < 0 Then
        MsgBox "开始时间不能大于结束时间！", vbExclamation, ParamInfo.系统名称
        dtpDateTo.SetFocus
        Exit Sub
    End If
    If dtpDateTo.Value - dtpDateFrom.Value > 60 Then
        MsgBox "审查时间范围太长(不能超过60天)！", vbExclamation, ParamInfo.系统名称
        dtpDateTo.SetFocus
        Exit Sub
    End If
    If dtpDateTo.Value - dtpDateFrom.Value > 31 Then
        If MsgBox("审查时间范围超过1个月，是否查询？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbNo Then
            dtpDateTo.SetFocus
            Exit Sub
        End If
    End If
    
    
    cmdSearch.Enabled = False
    
    mstrDateFrom = Format(Me.dtpDateFrom.Value, "yyyy-mm-dd")
    mstrDateTo = Format(Me.dtpDateTo.Value, "yyyy-mm-dd")
    
    If optKind(0).Value Then
        mintKind = 1
        strInfo = mstrDateFrom & " 至 " & mstrDateTo & " 内的门诊病历情况"
    ElseIf optKind(1).Value Then
        mintKind = 2
        strInfo = mstrDateFrom & " 至 " & mstrDateTo & " 内的住院病历情况"
    ElseIf optKind(2).Value Then
        mintKind = 4
        strInfo = mstrDateFrom & " 至 " & mstrDateTo & " 内的护理病历情况"
    Else
        optKind(1).Value = True
        mintKind = 2
        strInfo = mstrDateFrom & " 至 " & mstrDateTo & " 内的住院病历情况"
    End If
    
    stbThis.Panels(2).Text = "当前条件：" & strInfo
    
    tbcPage.Item(0).Tag = ""
    tbcPage.Item(1).Tag = ""
    tbcPage.Item(2).Tag = ""
    
    Call tbcPage_SelectedChanged(tbcPage.Selected)
    
End Sub


Private Sub dtpDateFrom_Change()
    cmdSearch.Enabled = True
End Sub

Private Sub dtpDateFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdSearch.Enabled = True
End Sub

Private Sub dtpDateFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpDateTo_Change()
    cmdSearch.Enabled = True
End Sub

Private Sub dtpDateTo_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdSearch.Enabled = True
End Sub

Private Sub dtpDateTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    
    
    
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    
    '-----------------------------------------------------
    '初始条件
    Call InitTerm
        
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Jump, "窗格跳转(&J)"): cbrControl.BeginGroup = True
    End With
    
    '帮助
    '------------------------------------------------------------------------------------------------------------------
    Call CreateHelpMenu(cbsThis)
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
        
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "展开"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
            
    '
    '------------------------------------------------------------------------------------------------------------------
    Call TabControlInit(tbcPage)
    With tbcPage
        .PaintManager.BoldSelected = True
        
        Set mfrmEPRAuditPati = New frmEPRAuditPati
        Call mfrmEPRAuditPati.zlInitData(Me)
        
        Set mfrmEPRAuditFile = New frmEPRAuditFile
        Call mfrmEPRAuditFile.zlInitData(Me)
        
        Set mfrmEPRAuditOutline = New frmEPRAuditOutline
        Call mfrmEPRAuditOutline.zlInitData(Me, mlngMoual)
        
        Call .InsertItem(0, "书写情况", mfrmEPRAuditOutline.hWnd, 0)
        Call .InsertItem(1, "按病人查阅", mfrmEPRAuditPati.hWnd, 0)
        Call .InsertItem(2, "按文件查阅", mfrmEPRAuditFile.hWnd, 0)
        
        .Item(0).Selected = True
        
    End With
        
            
    '读取发布到该模块的报表:因为是一次性读取,全局变量可用
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    mblnStartUp = False
    
    '刷新数据
    '------------------------------------------------------------------------------------------------------------------
    Call cmdSearch_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call SaveWinState(Me, App.ProductName)
    
    If Not (mfrmEPRAuditOutline Is Nothing) Then Unload mfrmEPRAuditOutline
    If Not (mfrmEPRAuditPati Is Nothing) Then Unload mfrmEPRAuditPati
    If Not (mfrmEPRAuditFile Is Nothing) Then Unload mfrmEPRAuditFile
    
End Sub

Private Sub optKind_Click(Index As Integer)
    Dim lngSelectIndex As Long
    cmdSearch.Enabled = True
    lngSelectIndex = tbcPage.Selected.Index
    Select Case Index
    Case 0 '门诊
        Select Case lngSelectIndex
        Case 0
            '不需要刷新
        Case 1
            Call mfrmEPRAuditPati.zlRefreshData(1, mstrDateFrom, mstrDateTo, True)
        Case 2
            Call mfrmEPRAuditFile.zlRefreshData(1, mstrDateFrom, mstrDateTo, True)
        End Select
    Case 1, 2 '住院护理
        Select Case lngSelectIndex
        Case 0
           '不需要刷新
        Case 1
            Call mfrmEPRAuditPati.zlRefreshData(Index * 2, mstrDateFrom, mstrDateTo, True)
        Case 2
            Call mfrmEPRAuditFile.zlRefreshData(Index * 2, mstrDateFrom, mstrDateTo, True)
        End Select
    End Select
End Sub

Private Sub optKind_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        tplThis.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 1
        tbcPage.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    End Select
End Sub

Private Sub vfgThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbRightButton Then Exit Sub

    Set cbrControl = Me.cbsThis.FindControl(, conMenu_File_Open)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub InitTerm()
    '-------------------------------------------------
    '--功能：初始化条件和布局
    '-------------------------------------------------
    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    
    '-----------------------------------------------------
    '初始数据:
    On Error GoTo errHand
    strSQL = "Select Sysdate From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With Me.dtpDateTo
        .Value = Format(rsTemp.Fields(0).Value, "yyyy-MM-dd")
        .MaxDate = .Value: .MinDate = Format("1990-01-01", "yyyy-MM-dd")
    End With
    With Me.dtpDateFrom
        .Value = Me.dtpDateTo.Value - 7
        .MaxDate = Me.dtpDateTo.MaxDate: .MinDate = Me.dtpDateTo.MinDate
    End With
    
    '-----------------------------------------------------
    '显示形态
    Set tplGroup = Me.tplThis.Groups.Add(0, "审查范围:")
    tplGroup.Expandable = False
    Set tplItem = tplGroup.Items.Add(0, "病历种类:", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set tplItem.Control = Me.picKind
    
    picKind.BackColor = tplItem.BackColor
    
    For lngCount = 0 To Me.optKind.count - 1
        optKind(lngCount).BackColor = tplItem.BackColor
    Next
    
    Set tplItem = tplGroup.Items.Add(0, "书写日期范围:", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set tplItem.Control = Me.picDate
    picDate.BackColor = tplItem.BackColor
'    Me.chkNoData.BackColor = tplItem.BackColor
    
    Set tplGroup = Me.tplThis.Groups.Add(0, "查阅说明:"): tplGroup.Expandable = True
    Set tplItem = tplGroup.Items.Add(0, "  ①、对应各事件的完成病历数，仅包含要求书写一次的病历，不包含要求循环书写的病历；", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "  ②、由于某些事件对应要求书写一次的病历为多种，因此其完成病历数可能超过发生人次数；", xtpTaskItemTypeText)
    
    '-----------------------------------------------------
    Me.tplThis.Reposition
    Me.BackColor = tplItem.BackColor
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '
    If mblnStartUp Then Exit Sub
    
    zlCommFun.ShowFlash "正在读取数据，请稍候...", Me
    DoEvents
    
    Select Case Item.Index
    Case 0
        If tbcPage.Item(0).Tag <> "已读" Then
            tbcPage.Item(0).Tag = "已读"
            If Not (mfrmEPRAuditOutline Is Nothing) Then Call mfrmEPRAuditOutline.zlRefreshData(mintKind, mstrDateFrom, mstrDateTo)
        End If
    Case 1
        If tbcPage.Item(1).Tag <> "已读" Then
            tbcPage.Item(1).Tag = "已读"
            If Not (mfrmEPRAuditPati Is Nothing) Then Call mfrmEPRAuditPati.zlRefreshData(mintKind, mstrDateFrom, mstrDateTo)
        End If
    Case 2
        If tbcPage.Item(2).Tag <> "已读" Then
            tbcPage.Item(2).Tag = "已读"
            If Not (mfrmEPRAuditFile Is Nothing) Then Call mfrmEPRAuditFile.zlRefreshData(mintKind, mstrDateFrom, mstrDateTo)
        End If
    End Select

    zlCommFun.StopFlash
End Sub
