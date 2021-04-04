VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomQueryCfg 
   Caption         =   "查询方案配置"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomQueryCfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   5400
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSScriptControlCtl.ScriptControl sctExecute 
      Left            =   4035
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   7560
      Left            =   255
      ScaleHeight     =   7560
      ScaleWidth      =   11415
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   285
      Width           =   11415
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   7560
         Left            =   6885
         TabIndex        =   3
         Top             =   0
         Width           =   85
         _ExtentX        =   159
         _ExtentY        =   13335
         SplitWidth      =   85
         SplitLevel      =   3
         Con1MinSize     =   3000
         Con2MinSize     =   5000
         Control1Name    =   "picScheme"
         Control2Name    =   "picSchemeCfg"
      End
      Begin VB.PictureBox picScheme 
         BorderStyle     =   0  'None
         Height          =   7560
         Left            =   0
         ScaleHeight     =   7560
         ScaleWidth      =   6885
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   6885
         Begin VB.ComboBox cbxDepart 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   120
            Width           =   2385
         End
         Begin zl9PACSWork.ucFlexGrid ufgScheme 
            Height          =   7290
            Left            =   105
            TabIndex        =   19
            Top             =   510
            Width           =   6660
            _ExtentX        =   11748
            _ExtentY        =   12859
            DefaultCols     =   ""
            DisCellColor    =   16777215
            HeadCheckValue  =   1
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontSize    =   10.5
            HeadFontCharset =   134
            HeadFontWeight  =   400
            HeadColor       =   0
            DataFontSize    =   10.5
            DataFontCharset =   134
            DataFontWeight  =   400
            DataColor       =   -2147483640
         End
         Begin VB.Label labDepart 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "所属科室:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   21
            Top             =   180
            Width           =   885
         End
      End
      Begin VB.PictureBox picSchemeCfg 
         BorderStyle     =   0  'None
         Height          =   7560
         Left            =   6970
         ScaleHeight     =   7560
         ScaleWidth      =   4440
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   4445
         Begin VB.Frame framSql 
            Caption         =   "过滤语句构造"
            Height          =   2415
            Left            =   255
            TabIndex        =   6
            Top             =   4980
            Width           =   6480
            Begin VB.CommandButton cmdInsertPar 
               Caption         =   "插入参数(&I)"
               Height          =   375
               Left            =   5085
               TabIndex        =   12
               Top             =   1935
               Width           =   1305
            End
            Begin VB.TextBox txtFilterSql 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1650
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   11
               Top             =   240
               Width           =   6345
            End
         End
         Begin VB.Frame framInput 
            Caption         =   "录入项配置"
            Height          =   3225
            Left            =   165
            TabIndex        =   5
            Top             =   1305
            Width           =   6585
            Begin VB.CommandButton cmdMoveNext 
               Caption         =   "下移行(&E)"
               Height          =   375
               Left            =   5415
               TabIndex        =   17
               Top             =   2745
               Width           =   1095
            End
            Begin VB.CommandButton cmdMoveLast 
               Caption         =   "上移行(&L)"
               Height          =   375
               Left            =   4335
               TabIndex        =   16
               Top             =   2745
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelInput 
               Caption         =   "删除行(&D)"
               Height          =   375
               Left            =   1185
               TabIndex        =   15
               Top             =   2745
               Width           =   1095
            End
            Begin VB.CommandButton cmdNewInput 
               Caption         =   "新增行(&N)"
               Height          =   375
               Left            =   105
               TabIndex        =   14
               Top             =   2745
               Width           =   1095
            End
            Begin zl9PACSWork.ucFlexGrid ufgInputCfg 
               Height          =   2430
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   6390
               _ExtentX        =   11271
               _ExtentY        =   4286
               DefaultCols     =   ""
               DisCellColor    =   16777215
               HeadCheckValue  =   1
               IsBtnNextCell   =   0   'False
               IsCopyAdoMode   =   0   'False
               IsEjectConfig   =   -1  'True
               HeadFontSize    =   10.5
               HeadFontCharset =   134
               HeadFontWeight  =   400
               HeadColor       =   0
               DataFontSize    =   10.5
               DataFontCharset =   134
               DataFontWeight  =   400
               DataColor       =   -2147483640
            End
         End
         Begin VB.Frame framBase 
            Caption         =   "基本信息配置"
            Height          =   795
            Left            =   165
            TabIndex        =   4
            Top             =   165
            Width           =   6570
            Begin VB.TextBox txtSchemeMemo 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4155
               TabIndex        =   9
               Top             =   270
               Width           =   2175
            End
            Begin VB.TextBox txtSchemeName 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1200
               TabIndex        =   7
               Top             =   270
               Width           =   1935
            End
            Begin VB.Label labSchemeMemo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "方案说明:"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3195
               TabIndex        =   10
               Top             =   360
               Width           =   975
            End
            Begin VB.Label labObj 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "方案名称:"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   8
               Top             =   360
               Width           =   975
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8175
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCustomQueryCfg.frx":0AE2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7938
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   480
      Top             =   300
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCustomQueryCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'菜单类型枚举定义
Private Enum TMenuType
    mtFile = 1
    mtSave = 2
    mtImport = 3
    mtExport = 4
    mtQuit = 5
    
    mtEdit = 6
    mtNewScheme = 7
    mtModifyScheme = 8
    mtDelScheme = 9
    mtSetDefault = 10
    mtUseScheme = 11
    mtMoveLastScheme = 12
    mtMoveNextScheme = 13
    mtCheckScheme = 14
    mtCancel = 15
    
End Enum

'查询方案列定义
Private Const M_STR_SCHEME_COLS As String = "|Id,hide,key|方案序号,hide|查询语句,hide|所属科室,hide|方案名称,read,w2100|是否默认,read,w1000|使用状态,read,w1000|方案说明,read,w2400|"
Private Const M_STR_SCHEME_CONVERT As String = "|是否默认:0-,1-默认|使用状态:0-禁用,1-启用|"

'录入配置列定义
'cbx<,[当前日期],[当前时间],[当前用户ID],[当前科室ID],[当前系统编号],[当前模块编号]>
Private Const M_STR_INPUT_COLS As String = "|ID,hide,key|方案ID,hide|录入顺序,hide|录入项目,w1400|录入类型,cbx<0-文本框,1-日期框,2-时间框,3-长日期框,4-下拉框,5-多选框>|" & _
                                            "默认值,btn,w1200|数据来源,w2400,btn|"
Private Const M_STR_INPUT_CONVERT As String = "|录入类型:0-文本框,1-日期框,2-时间框,3-长日期框,4-下拉框,5-多选框|"


Private mblnCurModifyState As Boolean
Public mblnIsChange As Boolean




Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '设置菜单栏和工具栏风格
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '总是在工具栏右侧显示选项按钮,即使窗体宽度足够。
        .ToolBarAccelTips = True                                '显示按钮提示
        .AlwaysShowFullMenus = False                            '不常用的菜单项先隐藏
        .UseFadedIcons = False                                  '图标显示为褪色效果
        .IconsWithShadow = True                                 '鼠标指向的命令图标显示阴影效果
        .UseDisabledIcons = True                                '工具栏按钮禁用时图标显示为禁用样式
        .LargeIcons = True                                      '工具栏显示为大图标
        .SetIconSize True, 24, 24                               '设置大图标的尺寸
        .SetIconSize False, 16, 16                              '设置小图标的尺寸
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                      '设置控件显示风格
        .EnableCustomization False                             '是否允许自定义设置
        Set .Icons = zlCommFun.GetPubIcons                     '设置关联的图标控件
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '菜单定义
'Begin------------------------编辑菜单--------------------------------------默认可见
    cbrMain.ActiveMenuBar.Title = "菜单"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtFile, "文件(&F)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "保存(&S)"): cbrControl.IconId = 3091
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "取消(&C)"): cbrControl.IconId = 3565
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtImport, "导入(&I)"): cbrControl.IconId = 0: cbrControl.BeginGroup = True
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtExport, "导出(&E)"): cbrControl.IconId = 0
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "退出(&Q)"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtEdit, "编辑(&E)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtNewScheme, "新增(&N)"): cbrControl.IconId = 4010
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtModifyScheme, "修改(&M)"): cbrControl.IconId = 3003
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtDelScheme, "删除(&D)"): cbrControl.IconId = 4008
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtSetDefault, "设置默认(&F)"): cbrControl.IconId = 3002: cbrControl.BeginGroup = True
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtUseScheme, "禁用(&A)"): cbrControl.IconId = 3006
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtMoveLastScheme, "上移(&L)"): cbrControl.IconId = 3082: cbrControl.BeginGroup = True
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtMoveNextScheme, "下移(&X)"): cbrControl.IconId = 21903
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtCheckScheme, "验证(&V)"): cbrControl.IconId = 3823: cbrControl.BeginGroup = True

    
    
    'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)
    
    'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(H)")
    Call CreateViewAndHelpMenu(Nothing, cbrMenuBar)
    
    
    
    '---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "保存", "保存方案"): cbrControl.IconId = 3091
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "取消", "取消修改"): cbrControl.IconId = 3565
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtNewScheme, "新增", "新增方案"): cbrControl.IconId = 4010: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtModifyScheme, "修改", "修改方案"): cbrControl.IconId = 3003
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtDelScheme, "删除", "删除方案"): cbrControl.IconId = 4008
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtSetDefault, "设置默认", "设置默认方案"): cbrControl.IconId = 3002: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtUseScheme, "禁用", "禁用方案"): cbrControl.IconId = 3006
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtMoveLastScheme, "上移", "上移方案"): cbrControl.IconId = 3082: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtMoveNextScheme, "下移", "下移方案"): cbrControl.IconId = 21903
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtCheckScheme, "验证", "验证方案"): cbrControl.IconId = 3823: cbrControl.BeginGroup = True

    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "退出", "退出"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    cbrControl.BeginGroup = True
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
'执行界面功能
Dim strResult As String

On Error GoTo errHandle
    Select Case control.ID
    
        Case TMenuType.mtCancel
            Call CancelScheme   '撤销方案
            
        Case TMenuType.mtNewScheme
            Call NewScheme      '新增方案
            
        Case TMenuType.mtModifyScheme
            Call ModifyScheme   '修改方案
            
        Case TMenuType.mtDelScheme
            Call DelScheme      '删除方案
            
        Case TMenuType.mtCheckScheme
            strResult = VerifyScheme   '验证方案
            If strResult <> "" Then
                MsgBoxD Me, strResult, vbOKOnly, Me.Caption
            Else
                MsgBoxD Me, "通过验证。", vbOKOnly, Me.Caption
            End If
            
        Case TMenuType.mtSave
            Call SaveScheme     '保存方案
            
        Case TMenuType.mtSetDefault
            Call DefaultScheme  '设置默认方案
            
        Case TMenuType.mtUseScheme
            Call UseScheme      '设置方案使用状态
            
        Case TMenuType.mtMoveLastScheme
            Call MoveLastScheme '上移方案
            
        Case TMenuType.mtMoveNextScheme
            Call MoveNextScheme '下移方案
            
        Case TMenuType.mtQuit
            Call Unload(Me)
            
        Case TMenuType.mtImport
            Call ImportScheme   '导入方案
            
        Case TMenuType.mtExport
            Call ExportScheme   '导出方案
            
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button '工具栏
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '按钮文字
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '大图标
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '状态栏
            Call Menu_View_StatusBar_click(control)
            
'--------------------------帮助-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Menu_Help_Web_Mail_click()
On Error GoTo errHandle
    zlMailTo hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_About_click()
On Error GoTo errHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'功能：调用帮助主题
On Error GoTo errHandle
    ShowHelp App.ProductName, Me.hWnd, Me.Name
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo errHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo errHandle
    zlHomePage hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    picBack.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ImportScheme()
'导入方案
    Dim rsData As ADODB.Recordset
    Dim lngOldSchemeId As Long
    Dim lngNewSchemeId As Long
    Dim strSql As String
    Dim strExeSql() As String
    Dim strResult As String
    Dim rsTemp As ADODB.Recordset
    
    dlgFile.Filter = "(*.XML)|*.XML|(*.*)|*.*"
    dlgFile.DefaultExt = "*.XML"
    
    dlgFile.Filename = ""
    dlgFile.ShowOpen
    
    If dlgFile.Filename = "" Then Exit Sub
    
    Set rsData = New ADODB.Recordset
    Call rsData.Open(dlgFile.Filename)
    
    If rsData.RecordCount <= 0 Then
        MsgBoxD Me, "没有可用于导入的数据，请检查文件是否正确。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    rsData.Sort = "id"
    
    lngOldSchemeId = 0
    rsData.MoveFirst
    ReDim Preserve strExeSql(1)
    
    While Not rsData.EOF
        If lngOldSchemeId <> Val(Nvl(rsData!ID)) Then
            '增加影像查询方案记录
            
            strSql = "select 影像查询方案_ID.NextVal as ID from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取新方案ID")
            If rsTemp.RecordCount <= 0 Then
                MsgBoxD Me, "不能获取新增的方案ID，系统将退出保存。", vbOKOnly, Me.Caption
                Exit Sub
            End If
            
            lngNewSchemeId = Val(Nvl(rsTemp!ID))
            
            ReDim Preserve strExeSql(UBound(strExeSql) + 1)
            strExeSql(UBound(strExeSql) - 1) = "zl_影像查询方案_新增方案(" & lngNewSchemeId & ",'" & _
                                                        Nvl(rsData!方案名称) & "','" & _
                                                        Nvl(rsData!方案说明) & "','" & _
                                                        Nvl(rsData!查询语句) & "'," & _
                                                        ufgScheme.ShowingDataRowCount + Val(Nvl(Nvl(rsData!方案序号))) & "," & _
                                                        0 & ")"
            
            lngOldSchemeId = Val(Nvl(rsData!ID))
        End If
        
        '增加影像方案配置记录
        ReDim Preserve strExeSql(UBound(strExeSql) + 1)
        strExeSql(UBound(strExeSql) - 1) = "zl_影像查询方案_新增配置(" & lngNewSchemeId & ",'" & _
                                                                Nvl(rsData!录入项目) & "'," & _
                                                                Val(Nvl(rsData!录入类型)) & ",'" & _
                                                                Nvl(rsData!默认值) & "','" & _
                                                                Nvl(rsData!数据来源) & "'," & _
                                                                Val(Nvl(rsData!录入顺序)) & ")"
        
        Call rsData.MoveNext
    Wend
    
    
    '写入方案配置相关数据
    strResult = ExeSqlTrans(strExeSql())
    If strResult <> "" Then
        MsgBoxD Me, "方案导入失败，原因为：" & strResult, vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Call LoadSchemeData
    
    MsgBoxD Me, "已成功导入" & rsData.RecordCount & "条数据。"
End Sub

Private Sub ExportScheme()
'导出方案
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    dlgFile.Filter = "(*.XML)|*.XML|(*.*)|*.*"
    dlgFile.DefaultExt = "*.XML"
    
    dlgFile.Filename = ""
    dlgFile.ShowSave
    
    If dlgFile.Filename = "" Then Exit Sub
    
    strSql = "select a.id, 方案名称,方案说明,查询语句,方案序号,是否默认,使用状态,b.id as 配置id,录入项目,录入类型,录入顺序,默认值,数据来源 " & _
            " from 影像查询方案 a, 影像查询配置 b where a.id=b.方案id and 使用状态=1 order by id, 录入顺序"
            
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "导出方案")
    If rsData.RecordCount <= 0 Then
        MsgBoxD Me, "没有可用于导出的数据，请检查方案设置。", vbOKOnly, Me.Caption
        Exit Sub
    End If
            
    Call rsData.Save(dlgFile.Filename, adPersistXML)
    
    MsgBoxD Me, "已成功导出" & rsData.RecordCount & "条数据。"
    
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单和按钮显示
On Error Resume Next
    Dim blnHasRecord As Boolean
    
    '如果没有记录或者没有选中行，菜单和工具栏则不可用
    blnHasRecord = ufgScheme.IsSelectionRow
    
    Select Case control.ID
    
        Case TMenuType.mtSave, TMenuType.mtCancel
            control.Enabled = mblnCurModifyState
            
        Case TMenuType.mtDelScheme, TMenuType.mtModifyScheme, _
            TMenuType.mtMoveLastScheme, TMenuType.mtMoveNextScheme
            control.Enabled = Not mblnCurModifyState And blnHasRecord
            
        Case TMenuType.mtNewScheme
            control.Enabled = Not mblnCurModifyState
            
        Case TMenuType.mtSetDefault
            control.Enabled = Not mblnCurModifyState And blnHasRecord
            
            If blnHasRecord Then
                If control.Parent.type = xtpControlPopup Then
                    control.Caption = IIf(ufgScheme.CurText("是否默认") = "默认", "取消默认(&F)", "设置默认(&F)")
                    control.IconId = IIf(ufgScheme.CurText("是否默认") = "默认", 2616, 3002)
                Else
                    control.Caption = IIf(ufgScheme.CurText("是否默认") = "默认", "取消默认", "设置默认")
                    control.IconId = IIf(ufgScheme.CurText("是否默认") = "默认", 2616, 3002)
                End If
                
                control.Enabled = Not mblnCurModifyState And IIf(ufgScheme.CurText("使用状态") = "启用", True, False)
                
                control.Enabled = Not control.Enabled
                control.Enabled = Not control.Enabled
            End If
            
        Case TMenuType.mtUseScheme
            control.Enabled = Not mblnCurModifyState And blnHasRecord
            
            If blnHasRecord Then
                If control.Parent.type = xtpControlPopup Then
                    control.Caption = IIf(ufgScheme.CurText("使用状态") = "启用", "禁用(&A)", "启用(&A)")
                    control.IconId = IIf(ufgScheme.CurText("使用状态") = "启用", 3006, 3009)
                Else
                    control.Caption = IIf(ufgScheme.CurText("使用状态") = "启用", "禁用", "启用")
                    control.IconId = IIf(ufgScheme.CurText("使用状态") = "启用", 3006, 3009)
                End If
                
                control.Enabled = Not control.Enabled
                control.Enabled = Not control.Enabled
            End If
            
        Case TMenuType.mtCheckScheme
            control.Enabled = blnHasRecord
            
    End Select
End Sub

Private Sub MoveLastScheme()
'上移方案
    Dim strSql As String
    
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Call ufgScheme.MoveUp(ufgScheme.SelectionRow)
    
    strSql = "zl_影像查询方案_移动方案(" & ufgScheme.CurKeyValue & "," & ufgScheme.SelectionRow & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "上移方案")
    
    mblnIsChange = True
End Sub


Private Sub MoveNextScheme()
'下移方案
    Dim strSql As String
    
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Call ufgScheme.MoveDown(ufgScheme.SelectionRow)
    
    strSql = "zl_影像查询方案_移动方案(" & ufgScheme.CurKeyValue & "," & ufgScheme.SelectionRow & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "下移方案")
    
    mblnIsChange = True
End Sub


Private Sub UseScheme()
'启用方案
    Dim strSql As String
    Dim strCurUseState As String
    
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    strCurUseState = ufgScheme.CurText("使用状态")
    
    strSql = "zl_影像查询方案_使用状态(" & ufgScheme.CurKeyValue & "," & IIf(strCurUseState = "启用", 0, 1) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "使用状态配置")
    
    ufgScheme.CurText("使用状态") = IIf(strCurUseState = "启用", "禁用", "启用")
    
    mblnIsChange = True
End Sub

Private Sub DefaultScheme()
'设置默认方案
    Dim strSql As String
    Dim strCurDefaultState As String
    Dim i As Long
    
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    strCurDefaultState = ufgScheme.CurText("是否默认")
    
    strSql = "zl_影像查询方案_配置默认(" & ufgScheme.CurKeyValue & "," & IIf(strCurDefaultState = "默认", 0, 1) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "配置默认")
    
    For i = 1 To ufgScheme.GridRows - 1
        ufgScheme.Text(i, "是否默认") = ""
    Next i
    
    ufgScheme.CurText("是否默认") = IIf(strCurDefaultState = "默认", "", "默认")
    
    mblnIsChange = True
End Sub

Private Sub NewScheme()
'新增方案
    ufgScheme.NewRow
    
    Call ConfigFaceEditState(True)
    
    ufgInputCfg.NewRow
    
    ufgInputCfg.CurText("录入项目") = "开始日期"
    ufgInputCfg.CurText("录入类型") = "1-日期框"
    
    
    ufgInputCfg.NewRow
    
    ufgInputCfg.CurText("录入项目") = "结束日期"
    ufgInputCfg.CurText("录入类型") = "1-日期框"
End Sub

Private Sub DelScheme()
'删除方案
    Dim strSql As String
    
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    strSql = "zl_影像查询方案_删除方案(" & ufgScheme.CurKeyValue & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, "删除查询方案")
    
    Call ufgScheme.DelCurRow(False)
    
    mblnIsChange = True
End Sub

Private Sub ModifyScheme()
'修改方案
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "请选择需要操作的查询方案。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Call ConfigFaceEditState(True)
End Sub

Private Sub SaveScheme()
'保存方案
'需要判断是修改原来的方案还是新增的方案
    Dim lngSchemeId As Long
    Dim rsData As ADODB.Recordset
    Dim strSql As String
    Dim i As Long
    Dim strExeSql() As String
    Dim strResult As String
    
    strResult = VerifyScheme
    If strResult <> "" Then
        If MsgBoxD(Me, strResult & vbCrLf & "需要强制保存吗？？？", vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    
    lngSchemeId = Val(ufgScheme.CurKeyValue)
    If lngSchemeId <= 0 Then
        '小于或等于0表示新增的方案
        strSql = "select 影像查询方案_ID.NextVal as ID from dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "获取新方案ID")
        If rsData.RecordCount <= 0 Then
            MsgBoxD Me, "不能获取新增的方案ID，系统将退出保存。", vbOKOnly, Me.Caption
            Exit Sub
        End If
        
        lngSchemeId = Val(Nvl(rsData!ID))
        
        ReDim Preserve strExeSql(1)
        strExeSql(0) = "zl_影像查询方案_新增方案(" & lngSchemeId & ",'" & _
                                                    txtSchemeName.Text & "','" & _
                                                    txtSchemeMemo.Text & "','" & _
                                                    txtFilterSql.Text & "'," & _
                                                    ufgScheme.SelectionRow & "," & _
                                                    cbxDepart.ItemData(cbxDepart.ListIndex) & ")"

    Else
        '修改的方案保存
        ReDim Preserve strExeSql(1)
        strExeSql(0) = "zl_影像查询方案_清除配置(" & lngSchemeId & ")"
        
        ReDim Preserve strExeSql(2)
        strExeSql(1) = "zl_影像查询方案_更新方案(" & lngSchemeId & ",'" & _
                                                    txtSchemeName.Text & "','" & _
                                                    txtSchemeMemo.Text & "','" & _
                                                    txtFilterSql.Text & "')"
        
    End If
    
    For i = 1 To ufgInputCfg.GridRows - 1
        If Not ufgInputCfg.RowHidden(i) Then
            ReDim Preserve strExeSql(UBound(strExeSql) + 1)
            strExeSql(UBound(strExeSql) - 1) = "zl_影像查询方案_新增配置(" & lngSchemeId & ",'" & _
                                                                    ufgInputCfg.Text(i, "录入项目") & "'," & _
                                                                    Val(ufgInputCfg.Text(i, "录入类型")) & ",'" & _
                                                                    ufgInputCfg.Text(i, "默认值") & "','" & _
                                                                    ufgInputCfg.Text(i, "数据来源") & "'," & _
                                                                    i & ")"
        End If
    Next i
        
    '写入方案配置相关数据
    strResult = ExeSqlTrans(strExeSql())
    If strResult <> "" Then
        MsgBoxD Me, "方案保存失败，原因为：" & strResult, vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    '更新方案列表显示
    With ufgScheme
        .CurKeyValue = lngSchemeId
        .CurText("方案名称") = txtSchemeName.Text
        .CurText("方案说明") = txtSchemeMemo.Text
        .CurText("查询语句") = Replace(txtFilterSql.Text, "''", "'")
        .CurText("是否默认") = ""
        .CurText("使用状态") = "启用"
        .CurText("方案序号") = ufgScheme.SelectionRow
    End With


    Call ConfigFaceEditState(False)
    
    mblnIsChange = True
End Sub


Private Function ExeSqlTrans(ByVal strSql As Variant) As String
    Dim i As Long
    Dim strExeSql As String
    
    ExeSqlTrans = ""
    
    gcnOracle.BeginTrans
    
On Error GoTo errRollback
    For i = 0 To UBound(strSql)
        strExeSql = strSql(i)
        If strExeSql <> "" Then
            Call zlDatabase.ExecuteProcedure(strExeSql, "保存方案配置")
        End If
    Next i
    
    gcnOracle.CommitTrans
Exit Function
errRollback:
    gcnOracle.RollbackTrans
    ExeSqlTrans = err.Description
End Function

Private Sub CancelScheme()
'取消方案的新增或者修改
    If Not mblnCurModifyState Then Exit Sub
    
    If ufgScheme.CurKeyValue = "" Then
        Call ufgInputCfg.ClearListData
        Call ufgScheme.DelCurRow(False)
        
        
        txtSchemeName.Text = ""
        txtSchemeMemo.Text = ""
        txtFilterSql.Text = ""
    Else
        Call LoadSchemeCfgData(ufgScheme.SelectionRow)
    End If
    
    
    Call ConfigFaceEditState(False)
End Sub


Private Function VerifyScheme() As String
'验证当前配置的方案
    Dim i As Long
    Dim j As Long
    Dim strInputProNames As String
    Dim strParName As String
    Dim strResult As String
    Dim rsTemp As ADODB.Recordset
    Dim strSqlFrom As String
    Dim strPars(20) As String
    
    VerifyScheme = ""
    
    If Trim(txtSchemeName.Text) = "" Then
        VerifyScheme = "未能通过验证，原因是方案名称为空。"
        Exit Function
    End If
    
    If Trim(txtFilterSql.Text) = "" Then
        VerifyScheme = "未能通过验证，原因是过滤语句构造为空。"
        Exit Function
    End If
        
    For i = 1 To ufgInputCfg.GridRows - 1
        strInputProNames = strInputProNames & "[" & ufgInputCfg.Text(i, "录入项目") & "]"
        
        '验证默认值配置
        strParName = ufgInputCfg.Text(i, "默认值")
        If IsParameterFormat(strParName) Then
            strResult = TestParameter(strParName, strInputProNames)
            
            If strResult <> "" Then
                VerifyScheme = "第" & i & "行'默认值'列的数据 " & vbCrLf & strParName & vbCrLf & " 未通过验证，原因如下：" & strResult
                Exit Function
            End If
        End If
        
        '验证数据来源配置
        strSqlFrom = ufgInputCfg.Text(i, "数据来源")
        strResult = TestSql(strSqlFrom, strInputProNames)
        If strResult <> "" Then
            VerifyScheme = "第" & i & "行'数据来源'列的数据 " & vbCrLf & strSqlFrom & vbCrLf & " 未通过验证，原因如下：" & strResult
            Exit Function
        End If
        
        
    Next i
    
    '验证过滤语句
    strResult = TestSql(Replace(txtFilterSql.Text, "''", "'"), strInputProNames)
    If strResult <> "" Then
        VerifyScheme = "过滤语句构造 " & vbCrLf & txtFilterSql.Text & vbCrLf & " 未通过验证，原因如下：" & strResult
        Exit Function
    End If

End Function


Private Function TestSql(ByVal strSqlFrom As String, ByVal strInputProNames As String) As String
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim strParName As String
    Dim strPars(20) As String
    Dim strTestSql As String
    
    TestSql = ""
        
    If strSqlFrom = "" Then Exit Function
    
    strTestSql = strSqlFrom
    
    Call GetParameterNames(strTestSql, strPars)
    
    For i = 1 To 20
        strParName = strPars(i)
        If strParName <> "" Then
            TestSql = TestParameter(strParName, strInputProNames)
            If TestSql <> "" Then Exit Function
            
            strTestSql = Replace(strTestSql, strParName, "Null")
            
            '恢复参数数组设置共后面继续使用
            strPars(i) = ""
        End If
    Next i
    
    'sql查询测试验证
    If Not IsParameterFormat(strTestSql) Then
        Set rsTemp = zlDatabase.OpenSQLRecord(strTestSql, "验证Sql查询")
    End If
    
Exit Function
errHandle:
    TestSql = err.Description
End Function

Private Function GetParameterNames(ByVal strSqlFrom As String, ByRef strParameters() As String) As Boolean
'判断数据源sql语句是否包含参数
    Dim strTemp As String
    Dim lngStart As Long, lngEnd As Long
    Dim lngParCount As Long
    
    strTemp = strSqlFrom
    lngStart = InStr(strTemp, "[")
    lngEnd = InStr(strTemp, "]")
    
    GetParameterNames = False
    
    If lngStart <= 0 Or lngEnd <= 0 Then Exit Function
    
    lngParCount = 0
    
    '循环获取所有的参数名称
    While lngStart > 0 And lngEnd > 0
        
        lngParCount = lngParCount + 1
        
        strTemp = Mid(strTemp, lngStart, 1024)
        lngEnd = InStr(strTemp, "]")
        
        strParameters(lngParCount) = Mid(strTemp, 1, lngEnd)
        
        strTemp = Mid(strTemp, lngEnd + 1, 1024)
        
        lngStart = InStr(strTemp, "[")
        lngEnd = InStr(strTemp, "]")
    Wend
       
    GetParameterNames = IIf(lngParCount > 0, True, False)
End Function

Private Function TestParameter(ByVal strParameterName As String, ByVal strInputProNames As String) As String
On Error GoTo errHandle

    TestParameter = ""
        
    If strParameterName = "" Then Exit Function
    If Not IsParameterFormat(strParameterName) Then
        '如果不是参数格式，则可能是直接由默认值配置传入的数据值，比如默认值配置的是“2012-03-05”，并没有采用“[当前时间]”方式
        TestParameter = ""
        Exit Function
    End If
    
    Select Case strParameterName
        Case "[当前日期]", "[当前时间]", "[当前用户ID]", "[当前科室ID]", "[当前系统编号]", "[当前模块编号]"
            Exit Function
        Case Else
            '获取文本框对应的值
             If InStr(strInputProNames, strParameterName) > 0 Then
                Exit Function
            End If
    End Select
    
    '在前面的代码中，如果找到对应的参数，就会直接将值覆盖函数并返回，如果执行到这里，说明没有找到对应参数，即可能是自定义脚本如“[now-1]”
    
    '执行脚本代码
    Call RunScripting(strParameterName)
Exit Function
errHandle:
    TestParameter = err.Description
End Function


Private Function IsParameterFormat(ByVal strData As String) As Boolean
'判断数据是否为参数数据
    IsParameterFormat = False
    
    If strData = "" Then Exit Function
    If Left(strData, 1) <> "[" Or Right(strData, 1) <> "]" Then Exit Function
    
    IsParameterFormat = True
End Function

Private Function RunScripting(ByVal strScript As String) As String
'执行vbs脚本
    Dim strFormatScript As String

    strFormatScript = Replace(Replace(strScript, "[", ""), "]", "")

On Error GoTo errHandle
    RunScripting = sctExecute.Eval(strFormatScript)
    Exit Function
errHandle:
    strFormatScript = "function return()" & vbCrLf & strFormatScript & " end function"
    Call sctExecute.AddCode(strFormatScript)
    
    RunScripting = sctExecute.Run("return")
End Function


Private Sub LoadSchemeData()
'载入方案
    Dim strSql As String
    
    strSql = "select id, 方案名称,方案说明,查询语句,方案序号,是否默认,使用状态,所属科室 from 影像查询方案 where 所属科室=[1] order by 方案序号"
    Set ufgScheme.AdoData = zlDatabase.OpenSQLRecord(strSql, "查询过滤方案", cbxDepart.ItemData(cbxDepart.ListIndex))
    
    Call ufgScheme.RefreshData
    
    If ufgScheme.ShowingDataRowCount > 1 Then
        Call ufgScheme.LocateRow(1)
    End If
End Sub

Private Sub LoadSchemeCfgData(ByVal lngSchemeRowIndex As Long)
'载入方案配置数据
    Dim strSql As String
    
    txtFilterSql.Text = Replace(ufgScheme.Text(lngSchemeRowIndex, "查询语句"), "'", "''")
    
    txtSchemeName.Text = ufgScheme.Text(lngSchemeRowIndex, "方案名称")
    txtSchemeMemo.Text = ufgScheme.Text(lngSchemeRowIndex, "方案说明")
    
    strSql = "select id, 方案ID,录入项目,录入类型,录入顺序,默认值,数据来源 from 影像查询配置 where 方案Id=[1] order by 录入顺序"
    Set ufgInputCfg.AdoData = zlDatabase.OpenSQLRecord(strSql, "查询方案配置", ufgScheme.KeyValue(lngSchemeRowIndex))
    
    Call ufgInputCfg.RefreshData
    
End Sub


Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
    picBack.Left = Left
    picBack.Top = Top
    picBack.Width = Right - Left
    picBack.Height = Bottom - Top - IIf(stbThis.Visible, stbThis.Height, 0)
End Sub



Private Sub cbxDepart_Click()
    Call LoadSchemeData
End Sub

Private Sub cmdDelInput_Click()
'删除录入项目
On Error GoTo errHandle
    Call ufgInputCfg.DelCurRow(False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdInsertPar_Click()
'插入参数
On Error GoTo errHandle
    Dim strPar As String
    Dim frmPar As New frmCustomInsertPar
    
    strPar = frmPar.ShowParameterWindow(ufgInputCfg, True, Me)
    If strPar <> "" Then
        txtFilterSql.SelText = strPar
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdMoveLast_Click()
'上移一行
On Error GoTo errHandle
    Call ufgInputCfg.MoveUp(ufgInputCfg.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdMoveNext_Click()
'下移一行
On Error GoTo errHandle
    Call ufgInputCfg.MoveDown(ufgInputCfg.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdNewInput_Click()
'新增录入项目
On Error GoTo errHandle
    Call ufgInputCfg.NewRow
    Call ufgInputCfg.DataGrid.EditCell
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    
    
'    InitDebugObject 1290, Me, "zlhis", "HIS"
    
    Call RestoreWinState(Me, App.ProductName)
    
    mblnCurModifyState = False
    mblnIsChange = False
    
    Call InitCommandBars
    Call InitFaceList

    Call ConfigFaceEditState(False)
    
    Call InitDepts
'    Call LoadSchemeData
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim str科室IDs As String, str来源 As String
    
    str来源 = "1,2,3"
    If InStr(gstrPrivs, "所有科室") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
            " And instr([1],','||B.服务对象||',')> 0 And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    Else
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=" & UserInfo.ID & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
            " And instr([1],','||B.服务对象||',')>0  And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    End If
   
    cbxDepart.Clear
    cbxDepart.AddItem "公共"
    cbxDepart.ItemData(cbxDepart.ListCount - 1) = 0
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "加载科室信息", CStr("," & str来源 & ","))
    
    If rsTmp.EOF Then
        MsgBoxD Me, "没有发现医技科室信息,请先到部门管理中设置。", vbInformation, gstrSysName
        cbxDepart.ListIndex = 0
        
        Exit Function
    End If
    
    While Not rsTmp.EOF
        cbxDepart.AddItem Nvl(rsTmp!名称)
        cbxDepart.ItemData(cbxDepart.ListCount - 1) = Nvl(rsTmp!ID)
        
        Call rsTmp.MoveNext
    Wend
    
    cbxDepart.ListIndex = 0
End Function


Private Sub InitFaceList()
'初始化界面相关的配置列表

    ufgScheme.IsKeepRows = False
    ufgScheme.IsEjectConfig = False
    ufgScheme.ColNames = M_STR_SCHEME_COLS
    ufgScheme.ColConvertFormat = M_STR_SCHEME_CONVERT
    ufgScheme.ExtendLastCol = True
    ufgScheme.RowHeightMin = 320
    
    ufgInputCfg.IsKeepRows = False
    ufgInputCfg.IsEjectConfig = False
    ufgInputCfg.ColNames = M_STR_INPUT_COLS
    ufgInputCfg.ColConvertFormat = M_STR_INPUT_CONVERT
    ufgInputCfg.ExtendLastCol = True
    ufgInputCfg.RowHeightMin = 320
End Sub


Private Sub ConfigFaceEditState(ByVal blnIsEdit As Boolean)
    mblnCurModifyState = blnIsEdit
    
    txtSchemeName.Locked = Not blnIsEdit
    txtSchemeMemo.Locked = Not blnIsEdit
    
    ufgInputCfg.ReadOnly = Not blnIsEdit
    
    cmdNewInput.Enabled = blnIsEdit
    cmdDelInput.Enabled = blnIsEdit
    cmdMoveLast.Enabled = blnIsEdit
    cmdMoveNext.Enabled = blnIsEdit
    
    txtFilterSql.Locked = Not blnIsEdit
    cmdInsertPar.Enabled = blnIsEdit
    
    cbxDepart.Enabled = Not blnIsEdit
    labDepart.Enabled = Not blnIsEdit
    
    ufgScheme.DataGrid.Enabled = Not blnIsEdit
    
    If blnIsEdit Then
        txtSchemeName.BackColor = &H80000005
        txtSchemeMemo.BackColor = &H80000005
        txtFilterSql.BackColor = &H80000005
        ufgInputCfg.BackColor = &H80000005
        ufgScheme.BackColor = &H8000000F
    Else
        txtSchemeName.BackColor = &H8000000F
        txtSchemeMemo.BackColor = &H8000000F
        txtFilterSql.BackColor = &H8000000F
        ufgInputCfg.BackColor = &H8000000F
        ufgScheme.BackColor = &H80000005
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'对象卸载
On Error GoTo errHandle
'    Unload frmCustomQueryFrom

    Call SaveWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picBack_Resize()
On Error Resume Next
    Call ucSplitter1.RePaint(False)
End Sub


Private Sub picScheme_Resize()
On Error Resume Next
    cbxDepart.Width = picScheme.Width - cbxDepart.Left - 60
    
    ufgScheme.Left = 60
    ufgScheme.Top = cbxDepart.Top + cbxDepart.Height + 60
    ufgScheme.Height = picScheme.Height - ufgScheme.Top
    ufgScheme.Width = picScheme.Width - 60
End Sub

Private Sub picSchemeCfg_Resize()
On Error Resume Next
    framBase.Left = 0
    framBase.Top = 0
    framBase.Width = picSchemeCfg.ScaleWidth

        txtSchemeMemo.Width = framBase.Width - txtSchemeMemo.Left - 120
        
    framInput.Left = 0
    framInput.Top = framBase.Top + framBase.Height + 60
    framInput.Width = picSchemeCfg.ScaleWidth
    framInput.Height = picSchemeCfg.ScaleHeight - framBase.Height - framSql.Height - 120
    
        ufgInputCfg.Left = 60
        ufgInputCfg.Top = 240
        ufgInputCfg.Width = framInput.Width - 120
        ufgInputCfg.Height = framInput.Height - cmdMoveLast.Height - 360
        
        cmdNewInput.Left = ufgInputCfg.Left
        cmdNewInput.Top = ufgInputCfg.Top + ufgInputCfg.Height + 60
        
        cmdDelInput.Left = cmdNewInput.Left + cmdNewInput.Width + 60
        cmdDelInput.Top = cmdNewInput.Top
        
        cmdMoveNext.Left = framInput.Width - 60 - cmdMoveNext.Width
        cmdMoveNext.Top = cmdNewInput.Top
        
        cmdMoveLast.Left = cmdMoveNext.Left - 60 - cmdMoveLast.Width
        cmdMoveLast.Top = cmdNewInput.Top
        
    framSql.Left = 0
    framSql.Top = framInput.Top + framInput.Height + 60
    framSql.Width = picSchemeCfg.ScaleWidth
    
        txtFilterSql.Left = 60
        txtFilterSql.Top = 240
        txtFilterSql.Width = framSql.Width - 120
        
        cmdInsertPar.Left = framSql.Width - 60 - cmdInsertPar.Width
    
End Sub

Private Sub ufgInputCfg_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
'配置录入项目的数据来源或者默认值
On Error GoTo errHandle
    Dim frmSqlFrom As New frmCustomQueryFrom
    Dim strSql As String
    Dim lngCurCol As Long
    
    lngCurCol = ufgInputCfg.GetColIndex("默认值")
    
    If Col > lngCurCol Then lngCurCol = Col


    strSql = ufgInputCfg.Text(Row, ufgInputCfg.GetColName(lngCurCol))
    
    strSql = frmSqlFrom.ShowSqlFromWindow(strSql, ufgInputCfg, Me)
    
    ufgInputCfg.Text(Row, ufgInputCfg.GetColName(lngCurCol)) = strSql
    
    Unload frmSqlFrom
Exit Sub
errHandle:
    Unload frmSqlFrom
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgInputCfg_OnDblClick()
'    If ufgInputCfg.ShowingDataRowCount < 1 Then Exit Sub
'
'    Call ufgInputCfg_OnCellButtonClick(ufgInputCfg.SelectionRow, 0)
End Sub

Private Sub ufgScheme_OnSelChange()
On Error GoTo errHandle
    If ufgScheme.IsSelectionRow Then
        Call LoadSchemeCfgData(ufgScheme.SelectionRow)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
