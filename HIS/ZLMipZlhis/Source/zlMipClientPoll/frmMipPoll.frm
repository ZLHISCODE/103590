VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMipPoll 
   Caption         =   "消息轮询服务"
   ClientHeight    =   8880
   ClientLeft      =   180
   ClientTop       =   -120
   ClientWidth     =   13260
   Icon            =   "frmMipPoll.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   13260
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picNotify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   10215
      ScaleHeight     =   720
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   1110
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8520
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMipPoll.frx":0A02
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18521
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin VB.Image imgService 
      Height          =   240
      Index           =   1
      Left            =   7635
      Picture         =   "frmMipPoll.frx":1296
      Top             =   2355
      Width           =   240
   End
   Begin VB.Image imgService 
      Height          =   240
      Index           =   0
      Left            =   7260
      Picture         =   "frmMipPoll.frx":1C98
      Top             =   2370
      Width           =   240
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
      Bindings        =   "frmMipPoll.frx":269A
      Left            =   375
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMipPoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义

Private mblnStartUp As Boolean
Private mblnServiceStart As Boolean
Private mcbrFileHide As CommandBarControl
Private mblnHided As Boolean
Private mintState As Integer
Private mstrTitle As String
Private mblnExitProgram As Boolean

Private mfrmMipPollLog As frmMipPollLog
Private WithEvents mfrmMipPollService As frmMipPollService
Attribute mfrmMipPollService.VB_VarHelpID = -1
Private mfrmMipPollConfig As frmMipPollConfig

'######################################################################################################################

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
    Dim objFindKey As CommandBarControl
    Dim intPostion As Integer
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    
    Call zlCommFun.CommandBarInit(cbsMain)
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    cbsMain.Options.LargeIcons = True
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
'    cbsMain.ActiveMenuBar.Visible = False
    
    '文件
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.Id = conMenu_FilePopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "配置(&C)...")
    
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Service_Start, "启动(&S)", True)
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Service_Stop, "停止(&T)")
    
    Set mcbrFileHide = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_View_Hide, "隐藏(&H)", True)
    Set mcbrFileHide = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_View_Show, "显示(&D)", True)
    
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)")
    
    '查看
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.Id = conMenu_ViewPopup
    Set objPopup = zlCommFun.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    
    
    '帮助
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.Id = conMenu_HelpPopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
'    Set objPopup = zlCommFun.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & zlComLib.zlRegInfo("产品简名"))
'    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, zlComLib.zlRegInfo("产品简名") & "主页(&H)")
'    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, zlComLib.zlRegInfo("产品简名") & "论坛(&F)")
'    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = True
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
            
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Parameter, "配置", , , xtpButtonIconAndCaption)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Service_Start, "启动", True, , xtpButtonIconAndCaption)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Service_Stop, "停止", False, , xtpButtonIconAndCaption)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出", True, , xtpButtonIconAndCaption)
    
'    cbsMain.StatusBar.Visible = True
'    cbsMain.StatusBar.IdleText = "准备"
'    Call cbsMain.StatusBar.AddPane(0)
'    Call cbsMain.StatusBar.SetPaneText(0, cbsMain.StatusBar.IdleText)
'    Call cbsMain.StatusBar.SetPaneStyle(0, SBPS_STRETCH)
'    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_CAPS)
'    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_NUM)
'    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_SCRL)
    
    '命令的快键绑定:公共部份主界面已处理
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain.KeyBindings
        
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        .Add 0, vbKeyF12, conMenu_File_Parameter            '参数

    End With

End Function


Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 300, DockLeftOf, Nothing)
    objPane.Title = "导航"
    objPane.Options = PaneNoCaption
        
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)
    
End Sub

Private Function StopService() As Boolean
    
    Dim blnSucced As Boolean
    
    If Not (mfrmMipPollService Is Nothing) Then
        
        blnSucced = mfrmMipPollService.StopService
        
        Unload mfrmMipPollService
        Set mfrmMipPollService = Nothing
        
        If blnSucced Then
            mblnServiceStart = False
            Call ModifyIcon(picNotify.hwnd, imgService(0).Picture, mstrTitle & "（已停止）")
        End If
    
    End If
    
    StopService = blnSucced
    
End Function

Private Sub HideShow(ByVal blnShow As Boolean)
    If blnShow Then
        Me.Show
        mblnHided = False
    Else
        Me.Hide
        mblnHided = True
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnSucced As Boolean
    
    Select Case Control.Id
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Parameter
        If mfrmMipPollConfig Is Nothing Then Set mfrmMipPollConfig = New frmMipPollConfig
        If mfrmMipPollConfig.ShowConfigDialog(Me) Then
    
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Service_Start
        If mfrmMipPollService Is Nothing Then Set mfrmMipPollService = New frmMipPollService
        If mfrmMipPollService.InitService = False Then Exit Sub
        mblnServiceStart = mfrmMipPollService.StartService
        If mblnServiceStart Then Call ModifyIcon(picNotify.hwnd, imgService(1).Picture, mstrTitle & "（运行中）")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Service_Stop

        Call StopService

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Hide
        
        Call HideShow(False)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Show
        
        Call HideShow(True)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Exit
        If Not (mfrmMipPollService Is Nothing) Then
            If mfrmMipPollService.ServerRunState Then
                If mfrmMipPollService.StopService = False Then Exit Sub
            End If
        End If
        mblnExitProgram = True
        Unload Me
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        Call CommandBarExecutePublic(Control, Me, 100)
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    Case conMenu_Service_Start
        Control.Enabled = Not mblnServiceStart
    Case conMenu_Service_Stop
        Control.Enabled = mblnServiceStart And (mintState = 0)
    Case conMenu_View_Hide
        Control.Visible = (mblnHided = False)
    Case conMenu_View_Show
        Control.Visible = (mblnHided = True)
    Case Else
        Call CommandBarUpdatePublic(Control, Me)
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
    Case 1
        Set mfrmMipPollLog = New frmMipPollLog
        Item.Handle = mfrmMipPollLog.hwnd
    End Select
End Sub

Private Sub Form_Activate()
    '
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    Call mfrmMipPollLog.ShowForm
End Sub

Private Sub Form_Load()
            
    mstrTitle = "消息轮询服务"
    mintState = 0
    mblnStartUp = True
    Call InitCommandBar
    Call InitDockPannel
    
    Call AddIcon(picNotify.hwnd, imgService(0).Picture, mstrTitle & "（未启动）")
    
    Call zlComLib.RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frmThis As Form
    
    If mblnExitProgram Then
        '是退出整个程序
        If mblnServiceStart And (mintState = 0) Then
            If MsgBox("服务已经启动，如果强行退出，自动停止服务，需要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, mstrTitle) = vbNo Then
                Cancel = True
                Exit Sub
            Else
                If StopService = False Then
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
        
        On Error Resume Next
        
        Call zlComLib.SaveWinState(Me, App.ProductName)
        
        Call RemoveIcon(picNotify.hwnd)
        
        If Not (mfrmMipPollService Is Nothing) Then
            Unload mfrmMipPollService
            Set mfrmMipPollService = Nothing
        End If
        
        If Not (mfrmMipPollConfig Is Nothing) Then
            Unload mfrmMipPollConfig
            Set mfrmMipPollConfig = Nothing
        End If
        
        If Not (mfrmMipPollLog Is Nothing) Then
            Unload mfrmMipPollLog
            Set mfrmMipPollLog = Nothing
        End If
            
        '关闭本部件窗体
        For Each frmThis In Forms
            If frmThis.Caption <> Me.Caption Then
                Unload frmThis
            End If
        Next
        
        Set gclsBusiness = Nothing
    Else
        '不是退出程序，只进行显示/隐藏处理
        Call HideShow(False)
        Cancel = True
    End If
End Sub

Private Sub mfrmMipPollService_AfterStateInfoChange(ByVal intState As Integer, ByVal strInfo As String)
    Me.stbThis.Panels(2).Text = strInfo
    mintState = intState
    DoEvents
End Sub

Private Sub picNotify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '--------------------------------------------------------------------------------------------------
    '功能:  处理picNotify的各种处理事件,主要用于自动提醒相关功能(陈渝编写)
    '--------------------------------------------------------------------------------------------------

    Select Case Hex(X) '
        Case "1E3C"     'Right-Button-Down
        Case "1E4B"     'Right-Button-Up
            Call ShowConetneMenu(1).ShowPopup
        Case "1830"     'Right-Button-Down LARGE FONTS '
        Case "1E1E"     'Left-Button-up
'            Call mnuFileOpen_Click
        Case "1E0F"     'Left-Button-Down '
        Case "1E2D"     'Left-Button-Double-Click '
            
            If mcbrFileHide.Enabled Then Call cbsMain_Execute(mcbrFileHide)
                        
        Case "1824"     'Left-Button-Double-Click
            
        Case "1E5A"     'Right-Button-Double-Click '
    End Select '
End Sub

Public Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '弹出菜单处理
    
    On Error GoTo errHand
    
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    
    Select Case bytPlace
    Case 1  '
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_File_Parameter, "配置(&C)...")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Service_Start, "启动(&S)")
        cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Service_Stop, "停止(&T)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Hide, "隐藏(&H)")
        cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Show, "显示(&D)")
        cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
