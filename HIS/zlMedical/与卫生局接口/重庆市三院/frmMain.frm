VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~4.OCX"
Begin VB.Form frmMain 
   Caption         =   "重庆市卫生局干保体检接口平台"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11880
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   2100
      Left            =   1170
      ScaleHeight     =   2040
      ScaleWidth      =   3885
      TabIndex        =   0
      Top             =   2175
      Width           =   3945
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   8040
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6852
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A72
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C92
            Key             =   "Accept"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D4F4
            Key             =   "Send"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DC6E
            Key             =   "Login"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DF88
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E702
            Key             =   "Diag"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14F64
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B7C6
            Key             =   "combo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7455
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":22028
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15875
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
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6975
      Top             =   2985
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   720
      Top             =   1035
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   150
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      DesignerControls=   "frmMain.frx":228BC
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mlngLoop As Long


Private Function InitMenuBar() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化菜单、工具栏
    '------------------------------------------------------------------------------------------------------------------
    Dim cbrMenuBar As Object
    Dim obj As CommandBarControl
    Dim cbrControl As Object
    Dim cbrToolBar As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Login, "注销用户(&L)...")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "本地参数(&S)")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出系统(&X)")
        cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "基础(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        'Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Dept, "体检部门对照(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Diagnose, "诊断建议(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Check, "检查项目(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Verify, "组合项目(&D)")
        
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "转换数据")
        
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "任务(&T)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Task_Accept, "接受任务(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Task_Send, "发送结果(&M)")
    End With
    

    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): cbrControl.BeginGroup = True
    End With
    
     '快键绑定
    With cbsThis.KeyBindings
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    

    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Login, "注销")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Diagnose, "诊断")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Check, "项目")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Verify, "组合")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Task_Accept, "接受")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Task_Send, "发送")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As Object

    On Error GoTo errHand
    
    Select Case Control.ID
        
        Case conMenu_File_Login
            
            If MsgBox("你确定真的要注销吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If CloseChildWindows(Me, Me) = False Then
                MsgBox "无法关闭部分窗体，注销操作中止！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            Unload Me
            
            Call Main
        
        Case conMenu_File_Parameter
            
            frmTaskAcceptParam.Show 1, Me
            
        Case conMenu_Edit_Dept
            
            frmDept.Show , Me
            
        Case conMenu_Edit_Combo
            
        Case conMenu_Edit_Check
            
            frmItems.Show , Me
            
        Case conMenu_Edit_Diagnose
            
            frm对码诊断.Show , Me
            
        Case conMenu_Edit_Verify
            
            frmCombo.Show , Me
        
        Case conMenu_Edit_Import
            
            dlg.Flags = &H4 Or &H200000 Or &H800 & &H1000
            dlg.Filter = "体检项目|体检项目.mdb"
            dlg.FilterIndex = 0
            
            dlg.DialogTitle = "导入体检项目"
            dlg.FileName = ""
            dlg.ShowOpen
            If Dir(dlg.FileName) <> "" Then
                
                If ImportData(Me, dlg.FileName, 782, "19", 383, "05", 563) Then
                    ShowSimpleMsg "转换成功！"
                End If
                
            End If
                        
        Case conMenu_Task_Accept
            
            dlg.Flags = &H4 Or &H200000 Or &H800 & &H1000
            dlg.Filter = "体检任务包|任务包.mdb"
            dlg.FilterIndex = 0
            
            dlg.DialogTitle = "体检任务包"
            dlg.FileName = ""
            dlg.ShowOpen
            If dlg.FileName <> "" Then
                If AcceptPackage(Me, dlg.FileName) Then
                    ShowSimpleMsg "接受任务包成功！"
                End If
            End If
            
        Case conMenu_Task_Send
            
            frmTaskSend.Show , Me
            
        Case conMenu_View_ToolBar_Button
        
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
        
            For Each cbrControl In cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            
            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
            
        Case conMenu_Help_Help
        
            Call ShowHelp(Me.hWnd, Me.Name)
        
        Case conMenu_Help_About
            
            frmAbout.Show 1, Me
            
        Case conMenu_File_Exit
        
            Unload Me
            Exit Sub
            
    End Select
    
    
    cbsThis.RecalcLayout
    
    Exit Sub
    
errHand:
    
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With picContainer
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Top = lngTop
        .Height = lngBottom - lngTop
    End With

End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
            
        Case conMenu_Edit_Dept, conMenu_Edit_Check, conMenu_Edit_Diagnose, conMenu_Edit_Verify
            
            Control.Visible = (InStr(1, gstrPrive, ";数据对码;") > 0)
                
                    
        Case conMenu_Edit_Import
            
            '
                        
        Case conMenu_Task_Accept
            
            Control.Visible = (InStr(1, gstrPrive, ";接受任务;") > 0)
            
        Case conMenu_Task_Send
            
            Control.Visible = (InStr(1, gstrPrive, ";发送结果;") > 0)
            
        Case conMenu_View_ToolBar_Button
            Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_StatusBar
            Control.Checked = Me.stbThis.Visible
            
    End Select
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    stbThis.Panels(2).Text = "用户：" & gstrDBUser
    If GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "") <> "" Then
         stbThis.Panels(2).Text = stbThis.Panels(2).Text & "@" & GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    End If
    
    stbThis.Panels(2).Text = stbThis.Panels(2).Text & "  姓名：" & UserInfo.姓名
    stbThis.Panels(2).Text = stbThis.Panels(2).Text & "  部门：" & UserInfo.部门
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    Call InitMenuBar
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frmThis As Form
    
    On Error Resume Next
    '关闭本部件窗体
    For Each frmThis In Forms
        If frmThis.Caption <> Me.Caption Then Unload frmThis
    Next

End Sub





