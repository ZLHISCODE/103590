VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmViewComm 
   Caption         =   "通讯监控"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   Icon            =   "frmViewComm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11340
   StartUpPosition =   2  '屏幕中心
   Begin RichTextLib.RichTextBox Rtxt_解析结果 
      Height          =   2460
      Left            =   6495
      TabIndex        =   0
      Top             =   3750
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   4339
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmViewComm.frx":000C
   End
   Begin RichTextLib.RichTextBox Rtxt_未知项 
      Height          =   2460
      Left            =   3795
      TabIndex        =   1
      Top             =   1020
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   4339
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmViewComm.frx":00A9
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7005
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmViewComm.frx":0146
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11509
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1905
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "2009-10-15"
            Key             =   "DATE"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "16:21"
            Key             =   "TIME"
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
      Left            =   780
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpCom 
      Bindings        =   "frmViewComm.frx":09DA
      Left            =   180
      Top             =   90
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmViewComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrDev As String '当前监控的设备
Private mstrCom As String
Private mstrCLSName As String
Private mLasterTime As Date

Public Event CloseWindow()

Public Sub ShowDecode(ByVal intType As Integer, ByVal str_In As String)
    '显示收到的解码结果和未知项
    'inttype: 0-结果 1-未知项
    If intType = 0 Then
        Rtxt_解析结果.Text = Rtxt_解析结果.Text & IIf(Rtxt_解析结果.Text = "", "", vbNewLine)
        Rtxt_解析结果.Text = Rtxt_解析结果.Text & "======解码结果======" & vbNewLine
        Rtxt_解析结果.Text = Rtxt_解析结果.Text & str_In
        Rtxt_解析结果.SelStart = Len(Rtxt_解析结果.Text)
    Else
        Rtxt_未知项.Text = Rtxt_未知项.Text & str_In
        Rtxt_未知项.SelStart = Len(Rtxt_未知项.Text)
    End If
End Sub

'Public Sub ShowInOut(ByVal intType As Integer, ByVal str_In As String, Optional ByVal lngEven As Long)
'    '功能：显示收到和发送的信息
'    'intType：0-接收 1-发送 2-错误
'    Me.stbThis.Panels(2).Text = "仪器：" & mstrDev & " " & mstrCom
'    Dim str秒 As String
''    str秒 = IIf(RTxt_Hex.Text = "", "", "(+ " & DateDiff("s", mLasterTime, Now) & " 秒)")
'    If intType = 0 Then
'        Rtxt_In.Text = Rtxt_In.Text & str_In
'        Rtxt_In.SelStart = Len(Rtxt_In.Text)
''        Select Case lngEven
''        Case -1 '是IP方式,其他的是跟onComm定义一样
''        End Select
'        RTxt_Hex.Text = RTxt_Hex.Text & IIf(RTxt_Hex.Text = "", "", vbNewLine)
'        RTxt_Hex.Text = RTxt_Hex.Text & "接收：" & Format(Now, "yyyy-MM-dd hh:MM:ss") & str秒
'        RTxt_Hex.Text = RTxt_Hex.Text & vbNewLine & str_In
'    ElseIf intType = 1 Then
'        Rtxt_Out.Text = Rtxt_Out.Text & str_In
'        Rtxt_Out.SelStart = Len(Rtxt_Out.Text)
'
'        RTxt_Hex.Text = RTxt_Hex.Text & IIf(RTxt_Hex.Text = "", "", vbNewLine)
'        RTxt_Hex.Text = RTxt_Hex.Text & "发送：" & Format(Now, "yyyy-MM-dd hh:MM:ss") & str秒
'        RTxt_Hex.Text = RTxt_Hex.Text & vbNewLine & str_In
'    Else
'        RTxt_Hex.Text = RTxt_Hex.Text & IIf(RTxt_Hex.Text = "", "", vbNewLine)
'        RTxt_Hex.Text = RTxt_Hex.Text & "错误：" & Format(Now, "yyyy-MM-dd hh:MM:ss") & str秒
'        RTxt_Hex.Text = RTxt_Hex.Text & vbNewLine & str_In
'    End If
'
'
'    RTxt_Hex.SelStart = Len(RTxt_Hex.Text)
'    'richtxtHEX 显示接收，发送的信息
'
'    mLasterTime = Now
'End Sub

Public Function ShowMe(ByVal strDev As String, ByVal strCOM As String, ByVal strCLSName As String)
    'strDev 设备名
    'strCom 串口及通讯参数
    'strCLSName 通讯程序名
    mstrDev = strDev
    mstrCom = ""
    If strCOM <> "" Then
        If UBound(Split(strCOM, "|")) > 1 Then
            
            mstrCom = Split(strCOM, "|")(0)
            Select Case Split(strCOM, "|")(1)
                Case "0": mstrCom = mstrCom & " 没有握手"
                Case "1": mstrCom = mstrCom & " (XON/XOFF) 握手"
                Case "2": mstrCom = mstrCom & " RTS/CTS 握手"
                Case "3": mstrCom = mstrCom & " RTS/CTS 和 XON/XOFF 握手皆可 "
            End Select
            Select Case Split(strCOM, "|")(2)
                Case "0": mstrCom = mstrCom & " 接收文本"
                Case "1": mstrCom = mstrCom & " 接收二进制"
            End Select
        ElseIf UBound(Split(strCOM, "|")) > 0 Then
            'mstrCom = Split(strCOM, "|")(1)
            Select Case Split(strCOM, "|")(0)
                Case "0": mstrCom = mstrCom & " 作为终端 端口：" & Mid(Split(strCOM, "|")(1), InStr(Split(strCOM, "|")(1), ":") + 1)
                Case "1": mstrCom = mstrCom & " 作为主机 端口：" & Mid(Split(strCOM, "|")(1), InStr(Split(strCOM, "|")(1), ":") + 1)
            End Select
        End If
    End If
    mstrCLSName = strCLSName
    
    Me.Show
    
End Function

Private Sub initCbsThis(cbsMain As CommandBars)
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
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
    cbsMain.Icons = frmPubIcons.imgPublic.Icons
    cbsMain.Options.LargeIcons = False
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)  '固有
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&T)…")  '固有
'        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
'        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "窗口(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_View & 2, "接收窗口(&J)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_View & 3, "发送窗口(&F)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_View & 4, "解码窗口(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_View & 5, "未知项窗口(&W)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "清空(&R)"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "清空所有(&C)")
    End With

'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False) '固有
'    objMenu.Id = conMenu_HelpPopup
'    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)") '固有
'
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrSysName)  '固有
'        With objPopup.CommandBar.Controls
'            .Add xtpControlButton, conMenu_Help_Web_Home, gstrSysName & "主页(&H)", -1, False '固有
'            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrSysName & "论坛(&F)", -1, False '固有
'            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False '固有
'        End With
'        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True '固有
'    End With

    '查找项特殊处理
    '-----------------------------------------------------


    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
'        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "清空窗口"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存窗口")

'        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.BeginGroup = True '固有
        

    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings

'        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印

'        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
    End With

    '设置一些公共的不常用命令
    '-----------------------------------------------------
'    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet         '打印设置
'        .AddHiddenCommand conMenu_File_Excel            '输出到Excel
'    End With

    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
'    Call gobjDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)

End Sub

Private Sub dkpComInit()
    Dim objPaneA As Pane, objPaneB As Pane, objPaneC As Pane, objPaneD As Pane, objPaneE As Pane
    Dim lngX As Long
    Dim lngY As Long
    
    'DockingPane 初始化
    '-----------------------------------------------------
    Me.dkpCom.SetCommandBars Me.cbsThis
     
'    Set objPaneA = Me.dkpCom.CreatePane(1, 200, 135, DockTopOf)
'    objPaneA.Title = "通讯信息"
'    objPaneA.Options = PaneNoCloseable Or PaneNoFloatable
'
'    Set objPaneB = Me.dkpCom.CreatePane(2, 380, 335, DockRightOf, objPaneA)
'    objPaneB.Title = "接收信息"
'
'    Set objPaneC = Me.dkpCom.CreatePane(3, 380, 335, DockBottomOf, objPaneB)
'    objPaneC.Title = "发送信息"
    
    Set objPaneD = Me.dkpCom.CreatePane(4, 100, 135, DockTopOf)
    objPaneD.Title = "解码结果"
    
    Set objPaneE = Me.dkpCom.CreatePane(5, 100, 135, DockRightOf, objPaneD)
    objPaneE.Title = "未知项目"
    
    Me.dkpCom.Options.UseSplitterTracker = False '实时拖动
    Me.dkpCom.Options.ThemedFloatingFrames = True
    Me.dkpCom.Options.AlphaDockingContext = False
    Me.dkpCom.Options.HideClient = True
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Dim objControl As CommandBarControl
    Dim strFileName As String
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button '工具栏
            For i = 2 To cbsThis.Count
                Me.cbsThis(i).Visible = Not Me.cbsThis(i).Visible
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text '按钮文字
            For i = 2 To cbsThis.Count
                For Each objControl In Me.cbsThis(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size '大图标
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
        Case conMenu_View_StatusBar '状态栏
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsThis.RecalcLayout
    
        Case conMenu_Edit_ClsAll
'            RTxt_Hex.Text = ""
'            Rtxt_In.Text = ""
'            Rtxt_Out.Text = ""
            Rtxt_解析结果.Text = ""
            Rtxt_未知项.Text = ""
        Case conMenu_Edit_Delete
           If TypeName(Me.ActiveControl) = "RichTextBox" Then
                Me.ActiveControl.Text = ""
           End If
        Case conMenu_Edit_Save
           If TypeName(Me.ActiveControl) = "RichTextBox" Then
                If Me.ActiveControl.Text <> "" Then
                    strFileName = Me.ActiveControl.Name
                    strFileName = Replace(strFileName, "in", "接收")
                    strFileName = Replace(strFileName, "Out", "发送")
                    strFileName = Replace(strFileName, "Hex", "通讯")
                    strFileName = Mid(strFileName, InStr(strFileName, "_") + 1) & Format(Now, "yyMMdd_hhMMss") & ".txt"
                    Me.ActiveControl.SaveFile App.Path & "\" & strFileName, 1
                    Me.stbThis.Panels(2).Text = "已保存为" & App.Path & "\" & strFileName
                End If
           End If
        Case conMenu_Edit_Seat_View & 2
            If dkpCom.Panes(2).Closed Then
                dkpCom.ShowPane 2
            Else
                dkpCom.Panes(2).Close
            End If
        Case conMenu_Edit_Seat_View & 3
            If dkpCom.Panes(3).Closed Then
                dkpCom.ShowPane 3
            Else
                dkpCom.Panes(3).Close
            End If
        Case conMenu_Edit_Seat_View & 4
            If dkpCom.Panes(4).Closed Then
                dkpCom.ShowPane 4
            Else
                dkpCom.Panes(4).Close
            End If
        Case conMenu_Edit_Seat_View & 5
            If dkpCom.Panes(5).Closed Then
                dkpCom.ShowPane 5
            Else
                dkpCom.Panes(5).Close
            End If
        Case conMenu_File_Exit        '退出
            Unload Me
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsThis.Count >= 2 Then
            Control.Checked = Me.cbsThis(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsThis.Count >= 2 Then
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
    
    Case conMenu_Edit_Seat_View & 2: Control.Checked = Not dkpCom.Panes(2).Closed
    Case conMenu_Edit_Seat_View & 3: Control.Checked = Not dkpCom.Panes(3).Closed
    Case conMenu_Edit_Seat_View & 4: Control.Checked = Not dkpCom.Panes(4).Closed
    Case conMenu_Edit_Seat_View & 5: Control.Checked = Not dkpCom.Panes(5).Closed
        
    End Select
End Sub

'--------------------------------------------------------------------

Private Sub dkpCom_AttachPane(ByVal Item As XtremeDockingPane.IPane)
'    If Item.ID = 1 Then
'        Item.Handle = RTxt_Hex.hwnd
'    ElseIf Item.ID = 2 Then
'        Item.Handle = Rtxt_In.hwnd
'    ElseIf Item.ID = 3 Then
'        Item.Handle = Rtxt_Out.hwnd
'    Else
    If Item.ID = 4 Then
        Item.Handle = Rtxt_解析结果.hwnd
    ElseIf Item.ID = 5 Then
        Item.Handle = Rtxt_未知项.hwnd
    End If
End Sub

Private Sub dkpCom_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    Me.cbsThis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    Top = lngTop
    Bottom = Me.ScaleHeight - lngBottom
End Sub

Private Sub dkpCom_Resize()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub Form_Load()
    Call initCbsThis(cbsThis)
    Call dkpComInit
    Me.stbThis.Panels(2).Text = "仪器：" & mstrDev & " " & mstrCom
    Me.Caption = Me.Caption & " 通讯程序:" & mstrCLSName
    
'    RTxt_Hex.Locked = True
'    Rtxt_In.Locked = True
'    Rtxt_Out.Locked = True
    Rtxt_解析结果.Locked = True
    Rtxt_未知项.Locked = True
End Sub

Private Sub Form_Resize()
    Me.dkpCom.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent CloseWindow
End Sub
