VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEInvoiceManager 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "电子票据自动开具工具"
   ClientHeight    =   4860
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8025
   Icon            =   "frmEInvoiceManager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraCons 
      Height          =   1845
      Left            =   90
      TabIndex        =   1
      Top             =   1080
      Width           =   7605
      Begin VB.TextBox txtSplit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1290
         TabIndex        =   3
         Text            =   "60"
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lbl轮询间隔 
         AutoSize        =   -1  'True
         Caption         =   "执行间隔：           （分钟）"
         Height          =   180
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2610
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4500
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEInvoiceManager.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "14:53:01"
            TextSave        =   "17:41"
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
      Left            =   240
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmEInvoiceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngTimerID As Long, mblnStart As Boolean

Public Function ShowMe(ByVal strPrivs As String) As Boolean
    '程序入口
    mstrPrivs = strPrivs
    Me.Show
End Function

Private Sub Form_Load()
    Call InitCommandBar
    Call RestoreWinState(Me, App.ProductName)
    
    '任务栏显示图标
    With nfIconData
        .hwnd = Me.hwnd
        .uID = Me.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon.Handle
        '定义鼠标移动到托盘上时显示的Tip
        .szTip = Me.Caption + "(版本 " & App.Major & "." & App.Minor & "." & App.Revision & ")" & vbNullChar
        .cbSize = Len(nfIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
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
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '菜单定义
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        'Set objControl = .Add(xtpControlButton, conMenu_File_ViewLog, "运行日志(&L)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Disuse, "停用(&P)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With

    '工具栏定义
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Disuse, "停用")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '命令的快键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
    '设置一些公共的不常用命令
    With cbsMain.Options
        '.AddHiddenCommand conMenu_File_PrintSet '打印设置
        '.AddHiddenCommand conMenu_File_Excel '输出到Excel
    End With
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    '根据Left,Top,Right,Bottom所返回的客户端区域位置对窗体中的其他控件进行排放
    fraCons.Move lngLeft + 10, lngTop, lngRight - lngLeft - 10, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean

    Select Case Control.ID
    Case conMenu_Edit_Reuse '启用
        Control.Enabled = Not mblnStart
    Case conMenu_Edit_Disuse '停用
        Control.Enabled = mblnStart And Not gblnExecuting
    
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub StartTimer()
    '开启定时器
    If Len(txtSplit.Text) > 4 Then
        MsgBox "执行间隔时间无效，请重新设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(txtSplit.Text) <= 0 Then
        MsgBox "执行间隔时间无效，请重新设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    zlWritLog glngModul, "启用自动开具电子票据任务", "StartTimer", "执行间隔：" & glngSplitTime
    
    gblnExecuting = False
    Set gfrmMain = Me
    glngSplitTime = Val(txtSplit.Text) * 60
     
    mlngTimerID = SetTimer(0, 0, glngSplitTime * 1000, AddressOf TimerProc)
    If mlngTimerID = 0 Then
        MsgBox "定时器启用失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnStart = True
    txtSplit.Locked = True
End Sub

Private Sub StopTimer()
    '停止定时器
    If gblnExecuting Then
        MsgBox "当前正在执行电子票据的开具，请稍后！", vbInformation, gstrSysName
        Exit Sub
    End If
    If mlngTimerID = 0 Then Exit Sub
     
    If KillTimer(0, mlngTimerID) = 0 Then
        MsgBox "定时器停用失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnStart = False
    txtSplit.Locked = False
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    
    Select Case Control.ID
    Case conMenu_File_ViewLog '日志查看
    
    Case conMenu_Edit_Reuse '启用
        Call StartTimer
        
    Case conMenu_Edit_Disuse '停用
        Call StopTimer
    
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '退出
        Unload Me
    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lMsg As Single
    lMsg = X / Screen.TwipsPerPixelX
    Select Case lMsg
    Case WM_LBUTTONUP
        '单击左键/右键，显示窗体
        ShowWindow Me.hwnd, SW_RESTORE
        '下面两句的目的是把窗口显示在窗口最顶层
        Me.Show
        Me.SetFocus
    Case WM_RBUTTONUP
        'PopupMenu MenuTray '如果是在系统Tray图标上点右键，则弹出菜单MenuTray
    Case WM_MOUSEMOVE
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN
    Case WM_RBUTTONDBLCLK
    Case Else
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If gblnExecuting Then
        MsgBox "当前正在执行电子票据自动开具任务，请稍后！", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
    
    If mblnStart Then
        If MsgBox("电子票据自动开具任务已启用，退出后将自定停止，你确定要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True: Exit Sub
        Call StopTimer
        Exit Sub
    End If
    
    If MsgBox("当前未启用电子票据自动开具任务，你确定要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True: Exit Sub
    
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Me.Hide '最小化时隐藏窗口
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    mlngTimerID = 0
    mblnStart = False
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub txtSplit_GotFocus()
    zlControl.TxtSelAll txtSplit
End Sub

Private Sub txtSplit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub
