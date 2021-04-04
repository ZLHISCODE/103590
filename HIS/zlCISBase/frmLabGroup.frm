VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmLabGroup 
   Caption         =   "检验小组设置"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13050
   Icon            =   "frmLabGroup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   13050
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox pic仪器 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5640
      Left            =   4215
      ScaleHeight     =   5640
      ScaleWidth      =   4095
      TabIndex        =   12
      Top             =   1245
      Width           =   4095
      Begin VB.PictureBox picEdit仪器 
         BorderStyle     =   0  'None
         Height          =   2715
         Left            =   -45
         ScaleHeight     =   2715
         ScaleWidth      =   3855
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3270
         Width           =   3855
         Begin VB.CommandButton cmd仪器 
            Caption         =   "∨"
            Height          =   350
            Index           =   1
            Left            =   2160
            TabIndex        =   15
            Top             =   105
            Width           =   1635
         End
         Begin VB.CommandButton cmd仪器 
            Caption         =   "∧"
            Height          =   350
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   75
            Width           =   1740
         End
         Begin MSComctlLib.ListView lvw仪器 
            Height          =   1980
            Left            =   195
            TabIndex        =   16
            Top             =   660
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   3493
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label lblEdit仪器 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   30
            Left            =   165
            TabIndex        =   17
            Top             =   540
            Width           =   90
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg仪器 
         Height          =   3195
         Left            =   15
         TabIndex        =   18
         Top             =   60
         Width           =   3825
         _cx             =   6747
         _cy             =   5636
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox PicList 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6525
      Left            =   15
      ScaleHeight     =   6525
      ScaleWidth      =   3945
      TabIndex        =   0
      Top             =   825
      Width           =   3945
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   6210
         Left            =   105
         TabIndex        =   1
         Top             =   165
         Width           =   3720
         _Version        =   589884
         _ExtentX        =   6562
         _ExtentY        =   10954
         _StockProps     =   0
         FocusSubItems   =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7575
      Width           =   13050
      _ExtentX        =   23019
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabGroup.frx":000C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17939
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin VB.PictureBox pic人员 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   8550
      ScaleHeight     =   5895
      ScaleWidth      =   4080
      TabIndex        =   3
      Top             =   1125
      Width           =   4080
      Begin VB.PictureBox picEdit人员 
         BorderStyle     =   0  'None
         Height          =   3675
         Left            =   90
         ScaleHeight     =   3675
         ScaleWidth      =   3855
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3390
         Width           =   3855
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1350
            TabIndex        =   8
            Top             =   240
            Width           =   2115
         End
         Begin VB.CommandButton cmdFind 
            Height          =   300
            Left            =   3480
            Picture         =   "frmLabGroup.frx":08A0
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "查找符合条件的项目"
            Top             =   225
            Width           =   360
         End
         Begin VB.CommandButton cmd人员 
            Caption         =   "∧"
            Height          =   350
            Index           =   0
            Left            =   45
            TabIndex        =   6
            Top             =   60
            Width           =   390
         End
         Begin VB.CommandButton cmd人员 
            Caption         =   "∨"
            Height          =   350
            Index           =   1
            Left            =   465
            TabIndex        =   5
            Top             =   30
            Width           =   555
         End
         Begin MSComctlLib.ListView lvw人员 
            Height          =   1830
            Left            =   435
            TabIndex        =   9
            Top             =   600
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   3228
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label lblEdit人员 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "查找人员:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   450
            TabIndex        =   10
            Top             =   360
            Width           =   810
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg人员 
         Height          =   3105
         Left            =   15
         TabIndex        =   11
         Top             =   30
         Width           =   3705
         _cx             =   6535
         _cy             =   5477
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   930
      Top             =   330
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabGroup.frx":0E2A
      Left            =   2520
      Top             =   195
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLabGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conPane_List = 201
Const conPane_Edit = 202
Const conPane_仪器 = 203
Const conPane_人员 = 204
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串
Private mLngEditWidth As Long       '为适应大字体情况下窗体变大.先读入窗体大小.
Private mintEditState As Integer    '当前编辑状态：0-非编辑状态,1-编辑状态
Private mlngGroupID As Long         '当前检验小组ID
Private mstr已有人员 As String      '当前小组的人员
Private mstr已有仪器 As String      '当前小组的仪器

'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long

Private Enum mCol
    ID = 0:  小组编码: 小组
    仪器Id = 0: 仪器编码: 名称: 查看: 更改: 只允许扫条码输入
    人员ID = 0: 人员编码: 姓名: 默认: 备注
End Enum
Private mblnInit As Boolean
Private mstr编码名称 As String

Private Sub initMenu()
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlcommfun.GetPubIcons
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "加入仪器(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "加入人员(&M)")
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
        
        
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread

        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_Edit_Untread
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "仪器"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "人员")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
        
End Sub

Private Sub initPic()
    '-----------------------------------------------------
    '设置停靠窗格
    Dim panThis As Pane, panSub1 As Pane, panSub2 As Pane
    
    mblnInit = True
    Set panThis = dkpMan.CreatePane(conPane_List, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "检验小组列表"
    panThis.Options = PaneNoCaption
    
'    Set panThis = dkpMan.CreatePane(conPane_Edit, 350, 50, DockRightOf, Nothing)
'    panThis.Title = "小组基本属性"
'    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption

    Set panSub1 = dkpMan.CreatePane(conPane_仪器, 550, 600, DockRightOf, panThis)
    panSub1.Title = "小组仪器"
    panSub1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panSub2 = dkpMan.CreatePane(conPane_人员, 550, 600, DockRightOf, panThis)
    panSub2.Title = "小组成员"
    panSub2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panSub2.AttachTo panSub1
    
    panSub2.Select

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    Me.dkpMan.VisualTheme = ThemeOffice2003
End Sub

Private Function zlEditStart(ByVal intAdd As Integer) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    
    Select Case mintEditState
    Case 1  '编码名称
        If mlngGroupID = 0 And intAdd <> 1 Then Exit Function
        If frmLabGroupEdit.ShowMe(mlngGroupID, intAdd, mstr编码名称, Me) Then
            Call zlSaveData
        End If
        mintEditState = 0

    Case 2  '仪器
        If mlngGroupID = 0 Then Exit Function
        strSQL = "Select ID,编码,名称 From 检验仪器"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        Me.lvw仪器.ListItems.Clear
        Do Until rsTmp.EOF
            If InStr(mstr已有仪器 & ",", "," & rsTmp.Fields("id") & ",") <= 0 Then
                Me.lvw仪器.ListItems.Add , "_" & rsTmp.Fields("id"), "" & rsTmp.Fields("编码") & " " & rsTmp.Fields("名称")
            End If
            rsTmp.MoveNext
        Loop
        Call pic仪器_Resize
        
        zlEditStart = True
    Case 3  '人员
        If mlngGroupID = 0 Then Exit Function
        Me.lvw人员.ListItems.Clear
        Call pic人员_Resize
        zlEditStart = True
    End Select
    Exit Function
ErrHandle:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub zlEditCancle()
    mintEditState = 0: Me.PicList.Enabled = True
    Me.vfg仪器.Enabled = True
    Me.vfg人员.Enabled = True
    Call pic人员_Resize
    Call pic仪器_Resize
    Call zlLoadData
End Sub

Private Sub zlSaveData()
    '保存数据
    Dim strSQL As String, lngID As Long
    Dim str编码 As String, str名称 As String, intRow As Integer
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    Select Case mintEditState
    Case 1  '编码名称

        
    Case 2  '仪器
        strSQL = ""
        With vfg仪器
            For intRow = .FixedRows To .Rows - 1
                strSQL = strSQL & "|" & .TextMatrix(intRow, mCol.仪器Id) & "," & IIf(.Cell(flexcpChecked, intRow, mCol.查看) = flexChecked, 1, 0) & _
                "," & IIf(.Cell(flexcpChecked, intRow, mCol.更改) = flexChecked, 1, 0) & "," & IIf(.Cell(flexcpChecked, intRow, mCol.只允许扫条码输入) = flexChecked, 1, 0)
            Next
        End With
        If strSQL <> "" Then
            strSQL = Mid(strSQL, 2)
            strSQL = "zl_检验小组仪器_Edit(" & mlngGroupID & ",'" & strSQL & "')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
    Case 3  '人员
        strSQL = ""
        With vfg人员
            For intRow = .FixedRows To .Rows - 1
                strSQL = strSQL & "|" & .TextMatrix(intRow, mCol.人员ID) & "," & IIf(.Cell(flexcpChecked, intRow, mCol.默认) = flexChecked, 1, 0)
            Next
        End With
        If strSQL <> "" Then
            strSQL = Mid(strSQL, 2)
            strSQL = "zl_检验小组成员_Edit(" & mlngGroupID & ",'" & strSQL & "')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
    End Select
    Call zlLoadData
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub zlEditDelete()
    '#
    Dim strSQL As String
    On Error GoTo ErrHandle
    
    strSQL = "zl_检验小组_Edit(3," & mlngGroupID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Call zlLoadData
    Exit Sub
ErrHandle:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CreateRptListHead()
    Dim Column As ReportColumn
    Dim i As Integer

    With Me.rptList.Columns

        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        
        End With
        'rptList.SetImageList Imglist

        Set Column = .Add(mCol.ID, "ID", 30, True): Column.Visible = False
        Set Column = .Add(mCol.小组编码, "编码", 100, True)
        Column.Sortable = True: Column.SortAscending = False: Me.rptList.SortOrder.Add Column
        
        Set Column = .Add(mCol.小组, "小组", 120, True)
    End With
End Sub

Private Sub zlLoadData()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim Record As ReportRecord
    Dim intLoop As Integer

    On Error GoTo ErrHandle
    rptList.Records.DeleteAll
    strSQL = "select id,编码,名称 from 检验小组"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Do Until rsTmp.EOF
        Set Record = Me.rptList.Records.Add
        For intLoop = 0 To Me.rptList.Columns.count + 1
            Record.AddItem ""
        Next
        Record.Item(mCol.ID).Value = Val("" & rsTmp!ID)
        Record.Item(mCol.小组编码).Value = Trim("" & rsTmp!编码)
        Record.Item(mCol.小组).Value = Trim("" & rsTmp!名称)
        rsTmp.MoveNext
    Loop
    rptList.Populate
    
    Dim rptParent As ReportRow
    If mlngGroupID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = mlngGroupID Then
                    Set rptParent = rptRow.ParentRow
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow Then
                If Not (rptRow Is rptParent) Then rptRow.Expanded = False
            End If
        Next
        Set Me.rptList.FocusedRow = Me.rptList.FocusedRow
    Else
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow Then rptRow.Expanded = False
        Next
    End If
    
    
    Call rptList_SelectionChanged
    
    Exit Sub
ErrHandle:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub


Private Sub initVfg仪器()
    With vfg仪器
        .BackColor = &H80000005
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .BackColorFixed = &HFDD6C6
        .GridLinesFixed = flexGridFlat
        .RowHeightMin = 300
        .Editable = flexEDNone
        
        .Rows = 2: .FixedRows = 1
        .Cols = 6: .FixedCols = 0
        
        .TextMatrix(0, mCol.仪器Id) = "": .ColWidth(mCol.仪器Id) = 0: .ColAlignment(mCol.仪器Id) = flexAlignRightCenter
        .ColHidden(mCol.仪器Id) = True
        .TextMatrix(0, mCol.仪器编码) = "编码": .ColWidth(mCol.仪器编码) = 1000: .ColAlignment(mCol.仪器编码) = flexAlignLeftCenter
        .TextMatrix(0, mCol.名称) = "名称": .ColWidth(mCol.名称) = 2000: .ColAlignment(mCol.名称) = flexAlignLeftCenter
        .TextMatrix(0, mCol.查看) = "查看": .ColWidth(mCol.查看) = 600: .ColAlignment(mCol.查看) = flexAlignLeftCenter
        .TextMatrix(0, mCol.更改) = "更改 ": .ColWidth(mCol.更改) = 600: .ColAlignment(mCol.更改) = flexAlignLeftCenter
        .TextMatrix(0, mCol.只允许扫条码输入) = "只允许扫条码输入": .ColWidth(mCol.只允许扫条码输入) = 600: .ColAlignment(mCol.只允许扫条码输入) = flexAlignLeftCenter
    End With
End Sub

Private Sub initVfg人员()
    With vfg人员
        .BackColor = &H80000005
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .BackColorFixed = &HFDD6C6
        .GridLinesFixed = flexGridFlat
        .RowHeightMin = 300
        .Editable = flexEDNone
        
        .Rows = 2: .FixedRows = 1
        .Cols = 5: .FixedCols = 0
        
        .TextMatrix(0, mCol.人员ID) = "": .ColWidth(mCol.人员ID) = 0: .ColAlignment(mCol.人员ID) = flexAlignRightCenter
        .ColHidden(mCol.人员ID) = True
        .TextMatrix(0, mCol.人员编码) = "编号": .ColWidth(mCol.人员编码) = 1000: .ColAlignment(mCol.人员编码) = flexAlignLeftCenter
        .TextMatrix(0, mCol.姓名) = "姓名": .ColWidth(mCol.姓名) = 2000: .ColAlignment(mCol.姓名) = flexAlignLeftCenter
        .TextMatrix(0, mCol.默认) = "默认小组": .ColWidth(mCol.默认) = 1000: .ColAlignment(mCol.默认) = flexAlignLeftCenter
        .TextMatrix(0, mCol.备注) = "备注": .ColWidth(mCol.默认) = 1000: .ColAlignment(mCol.备注) = flexAlignLeftCenter
    End With
End Sub

Private Sub zlRefresh()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
        
    vfg仪器.Clear
    mstr已有仪器 = ""
    Call initVfg仪器
    strSQL = "Select A.仪器id, B.编码, B.名称, A.查看, A.更改,a.条码输入" & vbNewLine & _
            "From 检验小组仪器 A, 检验仪器 B" & vbNewLine & _
            "Where A.仪器id = B.ID And A.小组id = [1] Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
    With vfg仪器
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, mCol.仪器Id) = Val("" & rsTmp.Fields("仪器id"))
            .TextMatrix(.Rows - 1, mCol.仪器编码) = "" & rsTmp.Fields("编码")
            .TextMatrix(.Rows - 1, mCol.名称) = "" & rsTmp.Fields("名称")
            
            mstr已有仪器 = mstr已有仪器 & "," & Val("" & rsTmp.Fields("仪器id"))
            If Val("" & rsTmp.Fields("查看")) = 1 Then
                .Cell(flexcpChecked, .Rows - 1, mCol.查看) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, mCol.查看) = flexUnchecked
            End If
            If Val("" & rsTmp.Fields("更改")) = 1 Then
                .Cell(flexcpChecked, .Rows - 1, mCol.更改) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, mCol.更改) = flexUnchecked
            End If
            If Val("" & rsTmp.Fields("条码输入")) = 1 Then
                .Cell(flexcpChecked, .Rows - 1, mCol.只允许扫条码输入) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, mCol.只允许扫条码输入) = flexUnchecked
            End If
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If Val(.TextMatrix(.Rows - 1, mCol.仪器Id)) = 0 Then .Rows = .Rows - 1
    End With
    
    vfg人员.Clear
    mstr已有人员 = ""
    Call initVfg人员
    strSQL = "Select A.人员id, B.编号, B.姓名, A.默认小组, A.备注" & vbNewLine & _
            "From 检验小组成员 A, 人员表 B" & vbNewLine & _
            "Where A.人员id = B.ID And A.小组id = [1] order by B.编号"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
    With vfg人员
        Do Until rsTmp.EOF
            mstr已有人员 = mstr已有人员 & "," & Val("" & rsTmp.Fields("人员id"))
            .TextMatrix(.Rows - 1, mCol.人员ID) = Val("" & rsTmp.Fields("人员id"))
            .TextMatrix(.Rows - 1, mCol.人员编码) = "" & rsTmp.Fields("编号")
            .TextMatrix(.Rows - 1, mCol.姓名) = "" & rsTmp.Fields("姓名")
            If Val("" & rsTmp.Fields("默认小组")) = 1 Then
                .Cell(flexcpChecked, .Rows - 1, mCol.默认) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, mCol.默认) = flexUnchecked
            End If
            .TextMatrix(.Rows - 1, mCol.备注) = "" & rsTmp.Fields("备注")
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If Val(.TextMatrix(.Rows - 1, mCol.人员ID)) = 0 Then .Rows = .Rows - 1
    End With
    Exit Sub
ErrHandle:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lng
    Select Case Control.ID
        Case conMenu_View_Refresh: Call zlLoadData
        Case conMenu_Help_Help:     Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_View_ToolBar_Button
            Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each cbrControl In Me.cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
        Case conMenu_View_StatusBar
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsThis.RecalcLayout

        Case conMenu_Edit_Save      '保存
            Call zlSaveData
            mintEditState = 0
            Me.PicList.Enabled = True
            Me.vfg仪器.Enabled = True
            Me.vfg人员.Enabled = True
            Call pic人员_Resize
            Call pic仪器_Resize
        Case conMenu_Edit_Untread   '放弃
            Call zlEditCancle

        Case conMenu_File_Exit      '退出
            If mintEditState <> 0 Then
                If MsgBox("在编辑状态中，如果现在退出，未保存的修改将丢失，是否继续？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                    Unload Me
                End If
            Else
                Unload Me
            End If
            
        Case conMenu_Edit_NewItem   '新增
            mintEditState = 1
            If zlEditStart(1) Then
                Me.PicList.Enabled = False
                Me.vfg人员.Enabled = False
                Me.vfg仪器.Enabled = False
            End If
        Case conMenu_Edit_Modify    '修改
            mintEditState = 1
            If zlEditStart(0) Then
                Me.PicList.Enabled = False
                Me.vfg人员.Enabled = False
                Me.vfg仪器.Enabled = False
            End If
        Case conMenu_Edit_Delete    '删除
            If MsgBox("将删除该小组下所有的人员和仪器设置，是否继续？", vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                Call zlEditDelete
            End If
        Case conMenu_Edit_Append    '仪器
            mintEditState = 2
            If zlEditStart(0) Then
                Me.dkpMan.FindPane(conPane_仪器).Select
                Me.PicList.Enabled = False
                Me.vfg人员.Enabled = False
                Me.vfg仪器.Enabled = True
            End If
            
        Case conMenu_Edit_Compend   '人员
            mintEditState = 3
            If zlEditStart(1) Then
                Me.dkpMan.FindPane(conPane_人员).Select
                Me.PicList.Enabled = False
                Me.vfg仪器.Enabled = False
                Me.vfg人员.Enabled = True
            End If
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Edit_Save, conMenu_Edit_Untread     '保存
            Control.Enabled = mintEditState <> 0
        Case conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Append, _
             conMenu_Edit_Compend, conMenu_View_Refresh
            Control.Enabled = mintEditState = 0
    End Select
End Sub

Private Sub cmdFind_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strFind As String
    
    On Error GoTo ErrHandle
    strFind = DelInvalidChar(UCase(Trim(Me.txtFind)))
    If strFind <> "" Then
        strSQL = "Select Distinct /*+Rule */" & vbNewLine & _
                " D.ID, D.编号, D.姓名, D.性别" & vbNewLine & _
                "From 部门性质说明 B, 部门人员 C, 人员表 D, 部门表 A" & vbNewLine & _
                "Where A.ID = B.部门id And B.工作性质 = '检验' And A.ID = C.部门id And C.人员id = D.ID And (" & _
                zlcommfun.GetLike("D", "编号", strFind) & " or " & zlcommfun.GetLike("D", "姓名", strFind) & " or " & zlcommfun.GetLike("D", "简码", strFind) & ")"
    Else
        strSQL = "Select Distinct /*+Rule */" & vbNewLine & _
                " D.ID, D.编号, D.姓名, D.性别" & vbNewLine & _
                "From 部门性质说明 B, 部门人员 C, 人员表 D, 部门表 A" & vbNewLine & _
                "Where A.ID = B.部门id And B.工作性质 = '检验' And A.ID = C.部门id And C.人员id = D.ID"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFind)
    Me.lvw人员.ListItems.Clear
    Do Until rsTmp.EOF
        If InStr(mstr已有人员 & ",", "," & rsTmp.Fields("id") & ",") <= 0 Then
            Me.lvw人员.ListItems.Add , "_" & rsTmp.Fields("id"), "" & rsTmp.Fields("编号") & " " & rsTmp.Fields("姓名")
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd人员_Click(Index As Integer)
    Dim ObjItem  As ListItem
    With vfg人员
    If Index = 0 Then

        If Me.lvw人员.SelectedItem Is Nothing Then Exit Sub
        Set ObjItem = Me.lvw人员.SelectedItem
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.人员ID) = Mid(ObjItem.Key, 2)
        .TextMatrix(.Rows - 1, mCol.人员编码) = Split(ObjItem.Text, " ")(0)
        .TextMatrix(.Rows - 1, mCol.姓名) = Split(ObjItem.Text, " ")(1)
        
        .Cell(flexcpChecked, .Rows - 1, mCol.默认) = flexChecked
        
        If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
        Me.lvw人员.ListItems.Remove ObjItem.Key: Me.lvw人员.SetFocus
    Else
        If .Row < .FixedRows Then Exit Sub
        Set ObjItem = Me.lvw人员.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.人员ID), .TextMatrix(.Row, mCol.人员编码) & " " & .TextMatrix(.Row, mCol.姓名))
        ObjItem.Selected = True
        .RemoveItem .Row
    End If
    End With
End Sub

Private Sub cmd仪器_Click(Index As Integer)
    Dim ObjItem  As ListItem
    With vfg仪器
    If Index = 0 Then

        If Me.lvw仪器.SelectedItem Is Nothing Then Exit Sub
        Set ObjItem = Me.lvw仪器.SelectedItem
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.仪器Id) = Mid(ObjItem.Key, 2)
        .TextMatrix(.Rows - 1, mCol.仪器编码) = Split(ObjItem.Text, " ")(0)
        .TextMatrix(.Rows - 1, mCol.名称) = Split(ObjItem.Text, " ")(1)
        
        .Cell(flexcpChecked, .Rows - 1, mCol.查看) = flexChecked
        .Cell(flexcpChecked, .Rows - 1, mCol.更改) = flexChecked
        .Cell(flexcpChecked, .Rows - 1, mCol.只允许扫条码输入) = flexChecked
        
        If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
        Me.lvw仪器.ListItems.Remove ObjItem.Key: Me.lvw仪器.SetFocus
    Else
        If .Row < .FixedRows Then Exit Sub
        Set ObjItem = Me.lvw仪器.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.仪器Id), .TextMatrix(.Row, mCol.仪器编码) & " " & .TextMatrix(.Row, mCol.名称))
        ObjItem.Selected = True
        .RemoveItem .Row
    End If
    End With
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
    If Action = PaneActionFloating Then Cancel = True
    If Action = PaneActionClosing Then Cancel = True
    
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.PicList.hWnd
    Case conPane_仪器
        Item.Handle = Me.pic仪器.hWnd
    Case conPane_人员
        Item.Handle = Me.pic人员.hWnd
    End Select

End Sub

Private Sub Form_Load()
    
    mstrPrivs = gstrPrivs
    
    mintEditState = 0
    Call zlcommfun.SetWindowsInTaskBar(Me.hWnd, False)

    Me.lvw仪器.ListItems.Clear
    With Me.lvw仪器.ColumnHeaders
        .Clear
        .Add , "_ID", "仪器列表", 3000
    End With
    With Me.lvw仪器
        .SortKey = .ColumnHeaders("_ID").Index - 1
        .SortOrder = lvwAscending
    End With

    Me.lvw人员.ListItems.Clear
    With Me.lvw人员.ColumnHeaders
        .Clear
        .Add , "_ID", "成员列表", 3000
    End With
    With Me.lvw人员
        .SortKey = .ColumnHeaders("_ID").Index - 1
        .SortOrder = lvwAscending
    End With
    
    Call initMenu
    Call initPic
    Call CreateRptListHead
    Call zlLoadData
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
End Sub

Private Sub Form_Resize()
'    Dim panBase As Pane
'    If Me.WindowState = vbMinimized Then Exit Sub
'    Set panBase = Me.dkpMan.FindPane(conPane_Edit)
'    mLngEditWidth = picGroupBase.ScaleHeight
'    panBase.MinTrackSize.SetSize 350, mLngEditWidth / Screen.TwipsPerPixelX
'    panBase.MaxTrackSize.SetSize 350, mLngEditWidth / Screen.TwipsPerPixelX
'    Me.dkpMan.RecalcLayout
'    Me.dkpMan.NormalizeSplitters
'
'    panBase.MinTrackSize.SetSize 0, 0
'    panBase.MaxTrackSize.SetSize 350, mLngEditWidth / Screen.TwipsPerPixelX
'
'    Me.dkpMan.RecalcLayout
'    Me.dkpMan.NormalizeSplitters

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvw人员_DblClick()
    Call cmd人员_Click(0)
End Sub

Private Sub lvw仪器_DblClick()
    Call cmd仪器_Click(0)
End Sub

Private Sub picEdit人员_Resize()
    Err = 0: On Error Resume Next
    Me.cmd人员(0).Left = Me.picEdit人员.ScaleLeft
    Me.cmd人员(0).Top = Me.picEdit人员.ScaleTop
    Me.cmd人员(0).Width = Me.picEdit人员.ScaleWidth / 2
    
    Me.cmd人员(1).Left = Me.cmd人员(0).Left + Me.cmd人员(0).Width
    Me.cmd人员(1).Top = Me.picEdit人员.ScaleTop
    Me.cmd人员(1).Width = Me.picEdit人员.ScaleWidth / 2
    
    
    Me.txtFind.Top = Me.cmd人员(0).Top + Me.cmd人员(0).Height + 15
    Me.txtFind.Left = Me.picEdit人员.ScaleLeft + Me.lblEdit人员.Width + 15
    Me.txtFind.Width = Me.picEdit人员.ScaleWidth - Me.txtFind.Left - Me.cmdFind.Width - 10
    
    Me.cmdFind.Left = Me.txtFind.Left + Me.txtFind.Width + 10
    Me.cmdFind.Top = Me.txtFind.Top
    
    
    Me.lblEdit人员.Top = Me.txtFind.Top + 25
    Me.lblEdit人员.Left = Me.picEdit人员.ScaleLeft
    
    
    Me.lvw人员.Left = Me.picEdit人员.ScaleLeft
    Me.lvw人员.Top = Me.txtFind.Top + Me.txtFind.Height + 15
    Me.lvw人员.Width = Me.picEdit人员.ScaleWidth
    Me.lvw人员.Height = Me.picEdit人员.ScaleHeight - Me.lvw人员.Top

End Sub

Private Sub picEdit仪器_Resize()
    Err = 0: On Error Resume Next
    Me.cmd仪器(0).Left = Me.picEdit仪器.ScaleLeft
    Me.cmd仪器(0).Top = Me.picEdit仪器.ScaleTop
    Me.cmd仪器(0).Width = Me.picEdit仪器.ScaleWidth / 2
    
    Me.cmd仪器(1).Left = Me.cmd仪器(0).Left + Me.cmd仪器(0).Width
    Me.cmd仪器(1).Top = Me.picEdit仪器.ScaleTop
    Me.cmd仪器(1).Width = Me.picEdit仪器.ScaleWidth / 2
    
    Me.lblEdit仪器.Top = Me.cmd仪器(0).Top + Me.cmd仪器(0).Height + 15
    Me.lblEdit仪器.Left = Me.picEdit仪器.ScaleLeft
    
    Me.lvw仪器.Left = Me.picEdit仪器.ScaleLeft
    Me.lvw仪器.Top = Me.lblEdit仪器.Top + Me.lblEdit仪器.Height + 15
    Me.lvw仪器.Width = Me.picEdit仪器.ScaleWidth
    Me.lvw仪器.Height = Me.picEdit仪器.ScaleHeight - Me.lvw仪器.Top
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = Me.PicList.ScaleLeft: .Width = Me.PicList.ScaleWidth - .Left
        .Height = Me.PicList.ScaleHeight - .Top
    End With
End Sub

Private Sub pic人员_Resize()
    Err = 0: On Error Resume Next
    With Me.vfg人员
        .Left = Me.pic人员.ScaleLeft
        .Top = Me.pic人员.ScaleTop
        .Width = Me.pic人员.ScaleWidth
        If mintEditState = 3 Then
            .Height = Me.pic人员.ScaleHeight - Me.picEdit人员.Height
            With Me.picEdit人员
                .Left = Me.pic人员.ScaleLeft
                .Top = Me.vfg人员.Top + Me.vfg人员.Height
                .Width = Me.pic人员.ScaleWidth
                .Visible = True
            End With
            
        Else
            .Height = Me.pic人员.ScaleHeight
            Me.picEdit人员.Visible = False
        End If
    End With
End Sub

Private Sub pic仪器_Resize()
    Err = 0: On Error Resume Next
    With Me.vfg仪器
        .Left = Me.pic仪器.ScaleLeft
        .Top = Me.pic仪器.ScaleTop
        .Width = Me.pic仪器.ScaleWidth
        If mintEditState = 2 Then
            .Height = Me.pic仪器.ScaleHeight - Me.picEdit仪器.Height
            With Me.picEdit仪器
                .Left = Me.pic仪器.ScaleLeft
                .Top = Me.vfg仪器.Top + Me.vfg仪器.Height
                .Width = Me.pic仪器.ScaleWidth
                .Visible = True
            End With
            
        Else
            .Height = Me.pic仪器.ScaleHeight
            Me.picEdit仪器.Visible = False
        End If
    End With
End Sub

Private Sub rptList_SelectionChanged()
    Dim i As Integer
    
    mstr编码名称 = ""
    If rptList.SelectedRows.count = 0 Then
        If rptList.Rows.count > 0 Then
            '有记录,取第个非分组行,做当前行
            For i = 0 To rptList.Rows.count - 1
                If Not rptList.Rows(i).GroupRow Then
                    rptList.Rows(i).Selected = True
                    mlngGroupID = Val(Me.rptList.Rows(i).Record(mCol.ID).Value)
                    mstr编码名称 = Me.rptList.Rows(i).Record(mCol.小组编码).Value & "|" & Me.rptList.Rows(i).Record(mCol.小组).Value
                    Exit For
                End If
            Next
        End If
    End If
        
    If Not Me.rptList.FocusedRow Is Nothing Then
        If Me.rptList.FocusedRow.GroupRow = True Then
            mlngGroupID = 0
        Else
            mlngGroupID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
            mstr编码名称 = Me.rptList.FocusedRow.Record(mCol.小组编码).Value & "|" & Me.rptList.FocusedRow.Record(mCol.小组).Value
        End If
    End If
    
    Dim panThis As Pane, panSub As Pane
    If mblnInit Then
        Set panThis = Me.dkpMan.FindPane(conPane_仪器)
        Set panSub = Me.dkpMan.FindPane(conPane_人员)
        panSub.AttachTo panThis
        panThis.Select
        mblnInit = False
    End If
    
    Me.dkpMan.RecalcLayout
    
    Call zlRefresh
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click
    End If
End Sub

Private Sub vfg人员_DblClick()
    Dim blnCheck As Boolean
'    If mintEditState = 3 Then
'        If InStr("," & mCol.默认 & ",", "," & vfg仪器.MouseCol & ",") <= 0 Then
'            Call cmd人员_Click(1)
'        Else
             blnCheck = Me.vfg人员.Cell(flexcpChecked, vfg人员.MouseRow, vfg人员.MouseCol) = flexChecked
             If blnCheck Then
                Me.vfg人员.Cell(flexcpChecked, vfg人员.MouseRow, vfg人员.MouseCol) = flexUnchecked
             Else
                Me.vfg人员.Cell(flexcpChecked, vfg人员.MouseRow, vfg人员.MouseCol) = flexChecked
             End If
'        End If
'    End If

End Sub

Private Sub vfg仪器_DblClick()
    Dim blnCheck As Boolean
'    If mintEditState = 2 Then
'        If InStr("," & mCol.更改 & ",", "," & vfg仪器.MouseCol & ",") <= 0 Then
'            Call cmd仪器_Click(1)
'        Else
             blnCheck = Me.vfg仪器.Cell(flexcpChecked, vfg仪器.MouseRow, vfg仪器.MouseCol) = flexChecked
             If blnCheck Then
                Me.vfg仪器.Cell(flexcpChecked, vfg仪器.MouseRow, vfg仪器.MouseCol) = flexUnchecked
             Else
                Me.vfg仪器.Cell(flexcpChecked, vfg仪器.MouseRow, vfg仪器.MouseCol) = flexChecked
             End If
'        End If
'    End If
End Sub
