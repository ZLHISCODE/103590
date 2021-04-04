VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "消息集成平台ZLHIS客户端辅助工具"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   14910
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2640
      Index           =   0
      Left            =   540
      ScaleHeight     =   2640
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   1785
      Width           =   3315
      Begin XtremeSuiteControls.TaskPanel tkpMain 
         Height          =   4770
         Left            =   210
         TabIndex        =   2
         Top             =   615
         Width           =   3210
         _Version        =   589884
         _ExtentX        =   5662
         _ExtentY        =   8414
         _StockProps     =   64
         VisualTheme     =   9
         Animation       =   1
         Behaviour       =   1
         ItemLayout      =   2
         HotTrackStyle   =   3
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   1
      Left            =   5190
      ScaleHeight     =   2535
      ScaleWidth      =   4845
      TabIndex        =   0
      Top             =   1140
      Width           =   4845
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   3195
      Top             =   1005
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Bindings        =   "frmMain.frx":6852
      Left            =   1620
      Top             =   345
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":6866
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   675
      Top             =   330
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mobjFso As New FileSystemObject
Private mobjTextStream As TextStream

Private mobjProgObject As Object
Private mstrProgObject As String
Private mstrDataPath As String

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 200, 100, DockLeftOf, Nothing)
    objPane.Title = "导航"
    objPane.Options = PaneNoCloseable
    
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, objPane)
    objPane.Title = "程序"
    objPane.Options = PaneNoCaption
        
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)

End Sub

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
        
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)
    Set cbsMain.Icons = ImageManager1.Icons
    cbsMain.Options.LargeIcons = True
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
        
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.id = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助(&H)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)...")
        
    Set objControl = NewCommandBar(objMenu, xtpControlButton, 8, "退出(&X)", True)
    
    
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.id = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    
    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.id = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = True
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助", False, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_About, "关于", False, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Close, "退出", True, , xtpButtonIconAndCaption)
    
    cbsMain.StatusBar.Visible = True
    cbsMain.StatusBar.IdleText = "准备"
    Call cbsMain.StatusBar.AddPane(0)
    Call cbsMain.StatusBar.SetPaneText(0, cbsMain.StatusBar.IdleText)
    Call cbsMain.StatusBar.SetPaneStyle(0, SBPS_STRETCH)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_CAPS)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_NUM)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_SCRL)
    
    Exit Function
    
errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'
End Function


Private Sub InitTaskPane()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
'    Dim objPane As Object
    Dim objItem As TaskPanelGroupItem
    Dim objGroup As TaskPanelGroup
'    Dim intGroup As Integer
'    Dim intItem As Integer
'    Dim strSQL As String
'    Dim rsGroup As cRecordset
'    Dim rsItem As cRecordset
'
    With tkpMain
        .SetIconSize 24, 24
        Call .Icons.AddIcons(ImageManager1.Icons)
        .VisualTheme = xtpTaskPanelThemeNativeWinXP
        .Behaviour = xtpTaskPanelBehaviourToolbox
'        .VisualTheme = xtpTaskPanelThemeToolbox
'        .ItemLayout = xtpTaskItemLayoutImagesWithTextBelow
        .SetMargins 6, 6, 6, 6, 6
        .SetGroupInnerMargins 5, 5, 5, 5
        
        .SelectItemOnFocus = True
        
        Set objGroup = .Groups.Add(1, "基本")
                
        Set objItem = objGroup.Items.Add(1, "校验XML", xtpTaskItemTypeLink, 2)
        objItem.Tag = "1001"
        
        Set objItem = objGroup.Items.Add(4, "收发测试", xtpTaskItemTypeLink, 6)
        objItem.Tag = "1004"
        
        Set objItem = objGroup.Items.Add(2, "生成XML", xtpTaskItemTypeLink, 3)
        objItem.Tag = "1002"
        
        Set objItem = objGroup.Items.Add(3, "生成脚本", xtpTaskItemTypeLink, 4)
        objItem.Tag = "1003"
        

        
        objGroup.Expanded = True
        objGroup.Expandable = True
        objGroup.CaptionVisible = False
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    Case conMenu_File_Close
        Unload Me
    End Select
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.Handle = picBack(0).hWnd
    Case 2
        Item.Handle = picBack(1).hWnd
    End Select
End Sub

Private Sub Form_Load()
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    
    mstrDataPath = App.Path & "\ProgData"
    If mobjFso.FolderExists(mstrDataPath) = False Then
        Call mobjFso.CreateFolder(mstrDataPath)
    End If
    
    Call InitCommandBar
    Call InitDockPannel
    Call InitTaskPane

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call SetPaneRange(dkpMain, 1, 150, 15, 150, Me.ScaleHeight)
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frmThis As Form
        
    On Error Resume Next
    '关闭本部件窗体
    For Each frmThis In Forms
        If frmThis.Caption <> Me.Caption Then Unload frmThis
    Next
    
    dkpMain.Panes(2).Handle = picBack(1).hWnd
    
    If Not (mobjProgObject Is Nothing) Then
        Set mobjProgObject = Nothing
    End If
            
    Set mobjFso = Nothing
    Set mobjTextStream = Nothing
    
    If Not (gobjComLib Is Nothing) Then
        gobjComLib.CloseWindows
        Set gobjComLib = Nothing
    End If
    
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = 1 Then gcnOracle.Close
        Set gcnOracle = Nothing
    End If
    
End Sub

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        tkpMain.Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    End Select
End Sub

Private Sub tkpMain_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    
    Dim strDataPath As String
    
    If mstrProgObject <> Item.Tag Then
        
        If Not (mobjProgObject Is Nothing) Then mobjProgObject.CloseFormObject
        
        mstrProgObject = Item.Tag
        
        If mstrProgObject <> "" Then
            
            On Error Resume Next
            Err = 0
            Select Case mstrProgObject
            Case "1001"
                Set mobjProgObject = New clsXMLValidator
            Case "1002"
                Set mobjProgObject = New clsXMLSchema
            Case "1003"
                Set mobjProgObject = New clsMessageScripter
            Case "1004"
                Set mobjProgObject = New clsTest
            End Select
            
            If Err = 0 Then
            
                strDataPath = mstrDataPath & "\" & mstrProgObject
                If mobjFso.FolderExists(strDataPath) = False Then Call mobjFso.CreateFolder(strDataPath)
                If mobjProgObject.Initialize(strDataPath) Then
                    dkpMain.Panes(2).Handle = mobjProgObject.GetFormObject.hWnd
                    DoEvents
                    mobjProgObject.ActiveFormObject
                Else
                    dkpMain.Panes(2).Handle = picBack(1).hWnd
                    Set mobjProgObject = Nothing
                End If
                
            Else
                dkpMain.Panes(2).Handle = picBack(1).hWnd
                Set mobjProgObject = Nothing
            End If
        Else
            dkpMain.Panes(2).Handle = picBack(1).hWnd
            Set mobjProgObject = Nothing
        End If
    End If
    
End Sub


