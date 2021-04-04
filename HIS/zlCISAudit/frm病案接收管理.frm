VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm病案接收管理 
   Caption         =   "病案接收管理"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11250
   Icon            =   "frm病案接收管理.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   11250
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   360
      ScaleHeight     =   6855
      ScaleWidth      =   4935
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame fraPatient 
         Caption         =   "病案基本信息"
         Height          =   2295
         Left            =   360
         TabIndex        =   26
         Top             =   120
         Width           =   3855
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Index           =   0
            Left            =   1200
            TabIndex        =   4
            Top             =   1440
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin VB.TextBox txtEndNo 
            Height          =   300
            Left            =   1200
            TabIndex        =   1
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkOutHp 
            Caption         =   "出院日期"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.ComboBox cboOutDept 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   2580
         End
         Begin VB.TextBox txtBeginNo 
            Height          =   300
            Left            =   1200
            TabIndex        =   0
            Top             =   240
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Index           =   0
            Left            =   1200
            TabIndex        =   5
            Top             =   1800
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   30
            Top             =   1845
            Width           =   180
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Height          =   180
            Index           =   0
            Left            =   960
            TabIndex        =   29
            Top             =   645
            Width           =   180
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "出院科室"
            Height          =   180
            Left            =   420
            TabIndex        =   28
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label lblNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院号"
            Height          =   180
            Left            =   600
            TabIndex        =   27
            Top             =   285
            Width           =   540
         End
      End
      Begin VB.Frame fraSong 
         Caption         =   "病案接收信息"
         Height          =   2775
         Left            =   360
         TabIndex        =   21
         Top             =   2520
         Width           =   3855
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   10
            Top             =   1200
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   7
            Top             =   345
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin VB.TextBox txtAuditingMan 
            Height          =   300
            Left            =   1200
            TabIndex        =   13
            Top             =   2280
            Width           =   2175
         End
         Begin VB.TextBox txtApplyman 
            Height          =   300
            Left            =   1200
            TabIndex        =   12
            Top             =   1920
            Width           =   2175
         End
         Begin VB.CheckBox chkIncept 
            Caption         =   "接收日期"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkRecord 
            Caption         =   "编目日期"
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   1245
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   8
            Top             =   720
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   11
            Top             =   1560
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin VB.Label lblAuditingMan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "接收人"
            Height          =   180
            Left            =   600
            TabIndex        =   25
            Top             =   2355
            Width           =   660
         End
         Begin VB.Label lblApplyman 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "运送人"
            Height          =   180
            Left            =   600
            TabIndex        =   24
            Top             =   1965
            Width           =   660
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Enabled         =   0   'False
            Height          =   180
            Index           =   2
            Left            =   960
            TabIndex        =   23
            Top             =   765
            Width           =   300
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Enabled         =   0   'False
            Height          =   180
            Index           =   3
            Left            =   960
            TabIndex        =   22
            Top             =   1620
            Width           =   300
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         Height          =   300
         Left            =   3360
         TabIndex        =   14
         Top             =   5520
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全选(&A)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   15
         Top             =   5520
         Width           =   1100
      End
      Begin VB.CommandButton cmdNoAll 
         Caption         =   "全不选(&N)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   16
         Top             =   5520
         Width           =   1100
      End
      Begin VB.Image imgCelNo 
         Height          =   240
         Index           =   3
         Left            =   3120
         Picture         =   "frm病案接收管理.frx":0442
         Top             =   5970
         Width           =   240
      End
      Begin VB.Image imgCelNo 
         Height          =   240
         Index           =   2
         Left            =   2115
         Picture         =   "frm病案接收管理.frx":6C94
         Top             =   5955
         Width           =   240
      End
      Begin VB.Image imgCelNo 
         Height          =   240
         Index           =   1
         Left            =   1155
         Picture         =   "frm病案接收管理.frx":6FD6
         Top             =   5955
         Width           =   240
      End
      Begin VB.Image imgCelNo 
         Height          =   240
         Index           =   0
         Left            =   75
         Picture         =   "frm病案接收管理.frx":7318
         Top             =   5955
         Width           =   240
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "　 未接收　　  已接收     已编目     再接收"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   135
         TabIndex        =   31
         Top             =   6000
         Width           =   3870
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   7290
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm病案接收管理.frx":765A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14764
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
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   6000
      ScaleHeight     =   4815
      ScaleWidth      =   3855
      TabIndex        =   18
      Top             =   120
      Width           =   3855
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   4575
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Width           =   3615
         _cx             =   6376
         _cy             =   8070
         Appearance      =   1
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
      Left            =   120
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frm病案接收管理.frx":7EEE
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frm病案接收管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objToolBar As CommandBar
Private objMenu As CommandBarPopup
Private objPopup As CommandBarPopup
Private objControl As CommandBarControl
Private objCombox As CommandBarComboBox
Private objExtendedBar As CommandBar

Private Const conMenu_Edit_Display = 209            '查看单据(&C)
Private Const conMenu_View_ToolBar_Visible = 6014    '隐藏工具栏(&V)
Private Const conMenu_View_Choose = 7900

Private mstrPrivs As String            '当前使用者权限串
Private mlngApplyId As Long            '病历科室
Private mlngTempId As Long
Private mstrApply As String
Private mblnBootUp As Boolean
Private mstrMsg As String
Private mstrListTitle As String
Private mbln病案系统 As Boolean

'日期设置
Private mdtInBeginDate As Date           '病案接收时间开始
Private mdtInEndDate As Date             '病案接收时间结束
Private mdtRecBeginDate As Date        '病案记录时间开始
Private mdtRecEndDate As Date          '病案记录时间结束
Private mdtOutBeginDate As Date        '病人出院时间开始
Private mdtOutEndDate As Date          '病人出院时间结束

Private mstrNoShowDate As String          '不显示历史已编目未登记出院日期

Private mblnShow As Boolean
Private mintDblClick As Integer
Private mintDelete As Integer
Private mstrFind As String
Private mcllTemp As New Collection

Private mintEditState As Integer
Private mlngModule As Long   '模块号

Private Function GetInitDept() As Boolean
    '----------------------------------------------------------------------------
    '功能:获取科室
    '返回:如有科室,则返回True,否则返回False
    '----------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If InStr(mstrPrivs, "所有科室") > 0 Then
        gstrSQL = "Select a.id,a.编码, a.名称" & vbNewLine & _
                "From 部门表 A, 部门性质说明 B" & vbNewLine & _
                "Where a.Id = b.部门id And b.工作性质 = '临床' And (b.服务对象 = 2 Or b.服务对象 = 3) And " & vbNewLine & _
                  Where撤档时间("A") & zl_获取站点限制(True, "a")
    Else
        gstrSQL = "Select a.id,a.编码, a.名称,c.缺省" & vbNewLine & _
                "From 部门表 A, 部门性质说明 B ,部门人员 C" & vbNewLine & _
                "Where a.Id = b.部门id And b.工作性质 = '临床' And (b.服务对象 = 2 Or b.服务对象 = 3) And A.ID=c.部门id And C.人员ID=" & UserInfo.ID & vbNewLine & _
                " And " & Where撤档时间("A") & zl_获取站点限制(True, "a")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If rsTemp.EOF Then
        mstrMsg = "当前无可用临床科室,请检查部门设置！"
        GetInitDept = False
        Exit Function
    End If
    
    With cboOutDept
        .Clear
        '装入数据
        If InStr(mstrPrivs, "所有科室") > 0 Then '有权限才增加所有科室
            .AddItem "所有部门"
             .ItemData(.NewIndex) = 1
             If mlngApplyId = 0 Then '首次刷新才选中所有部门
                 .ListIndex = .NewIndex
                 mstrApply = "所有部门"
             End If
        End If
            
        Do Until rsTemp.EOF
            .AddItem rsTemp!编码 & "-" & rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!ID
            
            If mlngApplyId = 0 And InStr(mstrPrivs, "所有科室") = 0 Then '首次打开窗口
                If UserInfo.部门ID = rsTemp!ID Then
                    .ListIndex = .NewIndex
                    mstrApply = rsTemp!编码 & "-" & rsTemp!名称
                End If
            ElseIf mlngApplyId > 0 Then '重复刷新
                If rsTemp!ID = mlngApplyId Then
                    .ListIndex = .NewIndex
                    mstrApply = rsTemp!编码 & "-" & rsTemp!名称
                End If
            End If
            
            rsTemp.MoveNext
        Loop
        cmdAll.Enabled = True
        cmdNoAll.Enabled = True
    End With
    GetInitDept = True
    mblnBootUp = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitMenus()
'初始化菜单及工具栏
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    cbsThis.EnableCustomization False
    
    '工具栏定义
    Set cbsThis.Icons = zlCommFun.GetPubIcons 'imgIcon.Icons
    With cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    '定义菜单
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup  '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "打印预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel...")
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "接收(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Display, "查看(&C)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "小图标(&D)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Visible, "隐藏工具栏(&V)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Choose, "选择列(&C)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的中联") '
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): objControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("X"), conMenu_File_Exit
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
        .AddHiddenCommand conMenu_Help_About
    End With

    '工具栏定义
    Set objToolBar = cbsThis.Add("工具栏", xtpBarTop)
    objToolBar.ShowTextBelowIcons = False
    objToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
    With objToolBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "接收"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each objControl In objToolBar.Controls
        objControl.STYLE = xtpButtonIconAndCaption
    Next
       
    '定义下拉菜单
    Set objExtendedBar = cbsThis.Add("Popup", xtpBarPopup)
    With objExtendedBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "接收")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Display, "查看(&C)"): objControl.BeginGroup = True
    End With

End Sub

Private Sub InitMenusElectron()
'初始化菜单及工具栏-电子病案接收
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    cbsThis.EnableCustomization False
    
    '工具栏定义
    Set cbsThis.Icons = zlCommFun.GetPubIcons 'imgIcon.Icons
    With cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    '定义菜单
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup  '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "打印预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel...")
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Plan, "审查接收(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "拒绝审查(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退接收(&1)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退拒绝(&2)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Display, "查看(&C)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "小图标(&D)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Visible, "隐藏工具栏(&V)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Choose, "选择列(&C)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "查阅病案(&V)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的中联") '
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): objControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("X"), conMenu_File_Exit
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
        .AddHiddenCommand conMenu_Help_About
    End With

    '工具栏定义
    Set objToolBar = cbsThis.Add("工具栏", xtpBarTop)
    objToolBar.ShowTextBelowIcons = False
    objToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
    With objToolBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Plan, "审查接收"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "拒绝审查")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退接收"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退拒绝")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "查阅"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each objControl In objToolBar.Controls
        objControl.STYLE = xtpButtonIconAndCaption
    Next
       
    '定义下拉菜单
    Set objExtendedBar = cbsThis.Add("Popup", xtpBarPopup)
    With objExtendedBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Plan, "审查接收")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "拒绝审查")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退接收"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退拒绝")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Display, "查看(&C)"): objControl.BeginGroup = True
    End With

End Sub

Private Sub InitDkpMain()
    Dim panThis As Pane
    Dim panfilter As Pane
    Set panfilter = dkpMain.CreatePane(1, 300, 200, DockLeftOf, Nothing)
    panfilter.Title = "病案运送信息查询设置"
'
    panfilter.Options = PaneNoFloatable Or PaneNoCloseable
    
    Set panThis = dkpMain.CreatePane(2, 750, 500, DockRightOf, Nothing)
    panThis.Title = "病案运送登记信息"
    panThis.Options = PaneNoFloatable Or PaneNoCloseable Or PaneNoHideable ' Or PaneNoCaption
    
    Call GetdkpMain(Me.Caption & "-" & mlngModule, "dkpMain")
    
    Me.dkpMain.SetCommandBars Me.cbsThis
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.LunaColors = True
    Me.dkpMain.Options.HideClient = True
End Sub

Private Function SavedkpMain(ByVal strCaption As String, ByVal strKey As String) As Boolean
    Dim strValue As String
    If dkpMain.FindPane(1).Hidden Then
        strValue = "0"
    Else
        strValue = "1"
    End If
     If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
     zlDatabase.SetPara "区域条件是否显示", strValue, glngSys, mlngModule
End Function

Private Function GetdkpMain(ByVal strCaption As String, ByVal strKey As String) As Boolean
    Dim strReg As String
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
'        dkpMain.FindPane(1).Hide
        Exit Function
    End If
    strReg = zlDatabase.GetPara("区域条件是否显示", glngSys, mlngModule, "")
    If strReg = "" Then
        dkpMain.FindPane(1).Hide 'HidePane panfilter
        Exit Function
    End If
    Err = 0: On Error GoTo errHand:
    If strReg = "0" Then
        dkpMain.FindPane(1).Hide
    End If
    GetdkpMain = True
    Exit Function
errHand:
End Function

Private Sub cboOutDept_Click()
    
    If Me.cboOutDept.ListCount = 0 Then Exit Sub
    If Me.cboOutDept.ListIndex = -1 Then Exit Sub
    
    If cboOutDept.ItemData(cboOutDept.ListIndex) = 1 And cboOutDept.Text = "所有部门" Then
        mlngApplyId = 0
        cmdAll.Enabled = True
        cmdNoAll.Enabled = True
    Else
        mlngApplyId = cboOutDept.ItemData(cboOutDept.ListIndex)
        cmdAll.Enabled = True
        cmdNoAll.Enabled = True
    End If
    
End Sub
Private Sub cboOutDept_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Call zlControl.CboSetIndex(cboOutDept.hWnd, zlControl.CboMatchIndex(cboOutDept.hWnd, KeyAscii))
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCurrIndex As Long
    Dim strNo As String
    Dim strMsg As String
    Dim rsSQL As ADODB.Recordset
    Dim strSQL As String
    Dim strNote As String
    Dim strNow As String
    
    On Error GoTo errHand
    Call SQLRecord(rsSQL)

    
'   处理按钮不可见其快捷方式还能执行
    If Control.ID <> 0 Then
        If cbsThis.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_File_Preview
            Call zlRptPrint(0)
            If Me.ActiveControl Is vfgList Then
                vfgList.Redraw = False
                zlRptPrint 0
                vfgList.Redraw = True
                vfgList.Col = 0
                vfgList.ColSel = vfgList.Cols - 1
            End If
        Case conMenu_File_Print
            Call zlRptPrint(1)
            If Me.ActiveControl Is vfgList Then
                vfgList.Redraw = False
                zlRptPrint 1
                vfgList.Redraw = True
                vfgList.Col = 0
                vfgList.ColSel = vfgList.Cols - 1
            End If
        Case conMenu_File_Excel
            If Me.ActiveControl Is vfgList Then
                vfgList.Redraw = False
                zlRptPrint 3
                vfgList.Redraw = True
                vfgList.Col = 0
                vfgList.ColSel = vfgList.Cols - 1
            End If
        Case conMenu_File_Parameter
            frm病案接收参数.参数设置 Me, mlngModule, mstrPrivs
            mblnShow = IIf(Val(zlDatabase.GetPara("不显示历史已编目未登记", glngSys, mlngModule)) = 1, 1, 0) = 1
            If mblnShow Then Call GetNoSHowData
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
        Case conMenu_Edit_NewItem
           Call SetAdd
        Case conMenu_Edit_Modify
            Call SetModify
        Case conMenu_Edit_Delete
            Call SetDelete
        Case conMenu_Edit_Display
            Call SetDisplay
        Case conMenu_View_ToolBar_Button
            For Each objControl In Me.cbsThis(2).Controls
                If objControl.Type <> xtpControlLabel Then
                    objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each objControl In Me.cbsThis(2).Controls
                If objControl.Type <> xtpControlLabel Then
                    objControl.STYLE = IIf(objControl.STYLE = xtpButtonCaption, xtpButtonIconAndCaption, xtpButtonCaption)
                End If
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            Call cbsThis_ChangeCaption(Control, "大图标(&D)", "小图标(&D)")
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Visible
            Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
            Call cbsThis_ChangeCaption(Control, "隐藏工具栏(&V)", "显示工具栏(&V)")
            Me.cbsThis.RecalcLayout
        Case conMenu_View_StatusBar
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsThis.RecalcLayout
        Case conMenu_View_Refresh
            Call SetRefresh
        Case conMenu_View_Choose
            Call SetVfgSelect
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_ReportView      '查阅首页
            Call RecordLook
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Plan   '审查接收
            Call SetAdd
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Refuse '拒绝审查
            With vfgList
                strMsg = "确认拒收如下病案吗?" & vbCrLf & vbCrLf
                strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(.Row, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(.Row, .ColIndex("住院号"))
                If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                    If frmPubNoteEdit.ShowNoteEdit(Me, "输入拒绝理由", strNote) Then
                         If strNow = "" Then strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                         
                         If Val(.TextMatrix(.Row, .ColIndex("提交ID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("病案状态值"))) = 1 Then
                                strSQL = "zl_病案提交记录_Refuse('" & Val(.TextMatrix(.Row, .ColIndex("提交ID"))) & "','" & UserInfo.姓名 & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),'" & strNote & "')"
                                Call SQLRecordAdd(rsSQL, strSQL)
                         End If
                         Call SQLRecordExecute(rsSQL, Me.Caption)
                         Call SetRefresh
                    End If
                End If
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Untread
            Select Case Control.Caption
            Case "回退接收", "回退接收(&1)"
                With vfgList
                    strMsg = "确认回退接收如下病案吗?" & vbCrLf & vbCrLf
                    strMsg = strMsg & ChkStrUniCode("姓名：" & .TextMatrix(.Row, .ColIndex("姓名")) & "                    ", 20) & "住院号：" & .TextMatrix(.Row, .ColIndex("住院号"))
                    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                        strSQL = "zl_病案提交记录_UnReceive('" & Val(.TextMatrix(.Row, .ColIndex("提交ID"))) & "')"
                        Call SQLRecordAdd(rsSQL, strSQL)
                        
                        '已经安装了病案,处理病案接收记录
                        If mbln病案系统 Then
                            strSQL = "Zl_电子病案接收记录_Delete(" & Val(.TextMatrix(.Row, .ColIndex("病人ID"))) & "," & Val(.TextMatrix(.Row, .ColIndex("主页ID"))) & ")"
                            Call SQLRecordAdd(rsSQL, strSQL)
                        End If
                        
                        Call SQLRecordExecute(rsSQL, Me.Caption)
                        
                        Call SetRefresh
                    End If
                    
                End With
            Case "回退拒绝", "回退拒绝(&2)"
                
                
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home:  Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Forum: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_Web_Mail:  Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:     Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else
            If Control.ID > 401 And Control.ID < 499 Then
            '相关报表执行
            Call OpenRpt(Control)
        End If
    End Select
    
    GoTo endHand
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:
End Sub

Private Function OpenRpt(ByVal Control As XtremeCommandBars.ICommandBarControl) As Boolean
    '------------------------------------------------------------------------------
    '功能:打开报表
    '参数:Control-执行报表的控件
    '返回:
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim arrData As Variant
    Dim strDept As String
    Dim strTemp As String
    
    arrData = Split(Control.Parameter, ",")
    strTemp = cboOutDept.Text
    If strTemp = "所有部门" Then
'        strDept = "所有科室"
        Call ReportOpen(gcnOracle, Val(arrData(0)), arrData(1), Me, "出院开始日期=" & CDate(Format(dtpBegin(0).Value, "yyyy-MM-dd")), _
                "出院结束日期=" & CDate(Format(dtpEnd(0).Value, "yyyy-MM-dd")), "出院科室=" & "is not null")
    Else
'        strDept = Split(strTemp, "-")(1)
        Call ReportOpen(gcnOracle, Val(arrData(0)), arrData(1), Me, "出院开始日期=" & CDate(Format(dtpBegin(0).Value, "yyyy-MM-dd")), _
                "出院结束日期=" & CDate(Format(dtpEnd(0).Value, "yyyy-MM-dd")), "出院科室=" & "=" & mlngApplyId)
    End If

End Function

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then
        Bottom = stbThis.Height
    Else
        Bottom = 0
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '进行权限控制
    Dim strAuditingMan As String
    Dim strNo As String
    
    Select Case Control.ID
        Case conMenu_View_ToolBar_Visible
            Call cbsThis_ChangeCaption(Control, "隐藏工具栏(&V)", "显示工具栏(&V)")
    End Select
        
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If

    On Error Resume Next
    If mlngModule = 201 Then '病案接收
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel: Control.Enabled = (Me.vfgList.Rows <> 0)
            Case conMenu_Edit_NewItem
                If (InStr(1, mstrPrivs, ";增加;") <> 0) Then
                    Control.Visible = True
                Else
                    Control.Enabled = False
                    Control.Visible = False
                End If
            Case conMenu_Edit_Modify
                If InStr(mstrPrivs, ";修改;") = 0 Then
                    Control.Enabled = False
                    Control.Visible = False
                    Exit Sub
                End If
                Call SetVerify_Update(Control)
                If Control.Enabled And Control.Visible Then
                    mintDblClick = 1
                Else
                    mintDblClick = 0
                End If
            Case conMenu_Edit_Delete
                If InStr(mstrPrivs, ";删除;") = 0 Then
                    Control.Enabled = False
                    Control.Visible = False
                    Exit Sub
                End If
                Call SetVerify_Update(Control)
            Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
            Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonCaption)
            Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
            Case conMenu_View_ToolBar_Visible: Control.Checked = Me.cbsThis(2).Visible
            Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
        End Select
    Else '电子病案接收
        Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel: Control.Enabled = (Me.vfgList.Rows <> 0)
        Case conMenu_Edit_NewItem
            If (InStr(1, mstrPrivs, "基本;") <> 0) Then
                Control.Visible = True
            Else
                Control.Enabled = False
                Control.Visible = False
            End If
        Case conMenu_Edit_Modify
            If InStr(mstrPrivs, "基本;") = 0 Then
                Control.Enabled = False
                Control.Visible = False
                Exit Sub
            End If
            Call SetVerify_Update(Control)
            If Control.Enabled And Control.Visible Then
                mintDblClick = 1
            Else
                mintDblClick = 0
            End If
        Case conMenu_Edit_Delete
            If InStr(mstrPrivs, "基本;") = 0 Then
                Control.Enabled = False
                Control.Visible = False
                Exit Sub
            End If
            Call SetVerify_Update(Control)
        Case conMenu_File_Parameter '参数
            Control.Enabled = (InStr(1, mstrPrivs, "参数设置;") <> 0)
        Case conMenu_Manage_ReportView '查阅病案
            Control.Enabled = (InStr(1, mstrPrivs, "查阅电子病案;") <> 0)
        Case conMenu_Manage_ReportView      '查阅首页
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Plan   '审查接收
            Control.Visible = IsPrivs(mstrPrivs, "审查接收")
            Control.Enabled = Control.Visible
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Refuse '拒绝审查
            Control.Visible = IsPrivs(mstrPrivs, "拒绝审查")
            
            If Me.vfgList.Rows = 0 Then
                Control.Enabled = False
            Else
                Control.Enabled = (Control.Visible And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("提交ID"))) > 0 And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("病案状态值"))) = 1)
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Untread
            Select Case Control.Caption
            Case "回退接收", "回退接收(&1)"
                 Control.Visible = IsPrivs(mstrPrivs, "回退接收")
            Case "回退拒绝", "回退拒绝(&2)"
                 Control.Visible = False ' IsPrivs(mstrPrivs, "回退拒绝")
            End Select
            

            Select Case Control.Caption
                Case "回退接收", "回退接收(&1)"
                
                    If vfgList.ColIndex("病案状态值") = -1 Then
                        Control.Enabled = False
                    Else
                        Control.Enabled = (Control.Visible And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("提交ID"))) > 0 And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("病案状态值"))) = 10)
                    End If

                Case "回退拒绝", "回退拒绝(&2)"
                
                    If vfgList.ColIndex("病案状态值") = -1 Then
                        Control.Enabled = False
                    Else
                        Control.Enabled = (Control.Visible And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("提交ID"))) > 0 And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("病案状态值"))) = 2)
                    End If
            End Select

        Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonCaption)
        Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
        Case conMenu_View_ToolBar_Visible: Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
    End If
End Sub

Private Sub cbsThis_ChangeCaption(ByVal Control As XtremeCommandBars.ICommandBarControl, OldCaption As String, NewCaption As String)
    Select Case Control.ID
        Case conMenu_View_ToolBar_Visible
            If Me.cbsThis(2).Visible Then
                If Control.Caption = NewCaption Then Control.Caption = OldCaption
            Else
                If Control.Caption = OldCaption Then Control.Caption = NewCaption
            End If
        Case Else
            If Control.Caption = OldCaption Then
                Control.Caption = NewCaption
            Else
                If Control.Caption = NewCaption Then Control.Caption = OldCaption
            End If
    End Select
End Sub

Private Sub chkIncept_Click()
    dtpBegin(1).Enabled = IIf(chkIncept.Value = 1, True, False)
    dtpEnd(1).Enabled = IIf(chkIncept.Value = 1, True, False)
'    lbldate(1).Enabled = IIf(chkIncept.Value = 1, True, False)
    lblTo(2).Enabled = IIf(chkIncept.Value = 1, True, False)
    If chkIncept.Value = 1 Then
        chkRecord.Value = 0
        dtpBegin(2).Enabled = IIf(chkRecord.Value = 1, True, False)
        dtpEnd(2).Enabled = IIf(chkRecord.Value = 1, True, False)
        lblTo(3).Enabled = IIf(chkRecord.Value = 1, True, False)
    End If
End Sub

Private Sub chkIncept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkOutHp_Click()
    dtpBegin(0).Enabled = IIf(chkOutHp.Value = 1, True, False)
    dtpEnd(0).Enabled = IIf(chkOutHp.Value = 1, True, False)
    lblTo(1).Enabled = IIf(chkOutHp.Value = 1, True, False)
End Sub

Private Sub chkOutHp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkRecord_Click()
    dtpBegin(2).Enabled = IIf(chkRecord.Value = 1, True, False)
    dtpEnd(2).Enabled = IIf(chkRecord.Value = 1, True, False)
'    lbldate(2).Enabled = IIf(chkRecord.Value = 1, True, False)
    lblTo(3).Enabled = IIf(chkRecord.Value = 1, True, False)
    If chkRecord.Value = 1 Then
        chkIncept.Value = 0
        dtpBegin(1).Enabled = IIf(chkIncept.Value = 1, True, False)
        dtpEnd(1).Enabled = IIf(chkIncept.Value = 1, True, False)
        lblTo(2).Enabled = IIf(chkIncept.Value = 1, True, False)
    End If
End Sub

Private Sub chkRecord_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdAll_Click()
    Dim lngRows As Long
    Dim i As Long
    Dim lngApplyId As Long
    Dim strStatus As String
    Dim strTemp As String
    
    If mlngModule = 201 Then '电子病案
        strTemp = ";增加;"
    Else
        strTemp = "基本;"
    End If
    
    With vfgList
        If .Rows > 1 Then
            lngRows = .Rows - 1
            For i = 1 To lngRows
                lngApplyId = Val(.TextMatrix(i, .ColIndex("出院科室id")))
                strStatus = Trim(.TextMatrix(i, .ColIndex("状态")))
                If InStr(mstrPrivs, strTemp) <> 0 And strStatus = "未接收" Then
                    mlngTempId = cboOutDept.ItemData(cboOutDept.ListIndex)
                    If lngApplyId = mlngTempId Or mlngTempId = 1 Then
                        .TextMatrix(i, .ColIndex("选择")) = -1
                    Else
                        .TextMatrix(i, .ColIndex("选择")) = 0
                    End If
                Else
                    .TextMatrix(i, .ColIndex("选择")) = 0
                End If
            Next
        End If
    End With
End Sub

Private Sub cmdNoAll_Click()
    Dim lngRows As Long
    Dim i As Long
    Dim strStatus As String
    Dim strTemp As String
    
    If mlngModule = 201 Then '电子病案
        strTemp = ";增加;"
    Else
        strTemp = "基本;"
    End If
    
    With vfgList
         If .Rows > 1 Then
            lngRows = .Rows - 1
                For i = 1 To lngRows
            
                strStatus = Trim(.TextMatrix(i, .ColIndex("状态")))
                If InStr(mstrPrivs, strTemp) <> 0 And strStatus = "未接收" Then
                    .TextMatrix(i, .ColIndex("选择")) = 0
                    mlngTempId = 0
'                Else
'                    .TextMatrix(i, .ColIndex("选择")) = 0
                End If
            Next
           
        End If
    End With
End Sub

Private Sub cmdRefresh_Click()
    Call SetRefresh
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = picFilter.hWnd
        Case 2
            Item.Handle = picList.hWnd
    End Select
End Sub

Private Sub dtpBegin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtpEnd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Activate()
    If mblnBootUp = False Then
        If mstrMsg = "" Then Unload Me: Exit Sub
        MsgBox mstrMsg, vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    mstrPrivs = gstrPrivs
   ' mstrPrivs = "基本;参数设置;审查接收;回退拒绝;回退接收;拒绝审查;所有科室;查阅电子病案;" '调试使用
    If ParamInfo.模块号 = 201 Then
        mlngModule = 201
        Me.Caption = "病案接收管理"
    Else
        mlngModule = ParamInfo.模块号
        Me.Caption = "电子病案接收"
        
        Set rs = GetMedicalExits
        If Not rs.EOF Then
            mbln病案系统 = True
        Else
            mbln病案系统 = False
        End If
        
    End If
    
    
    mblnBootUp = False
    mlngApplyId = 0
    mlngTempId = 0
    mintDblClick = 0
    mintDelete = 0
    
    If mlngModule = 201 Then   '病案接收菜单
        Call InitMenus
    Else
        Call InitMenusElectron '电子病案接收菜单
    End If
    Call InitDkpMain
       
    If Not GetInitDept Then Exit Sub
    
    mblnShow = IIf(Val(zlDatabase.GetPara("不显示历史已编目未登记", glngSys, mlngModule)) = 1, 1, 0) = 1
    If mblnShow Then Call GetNoSHowData
    
    mdtInBeginDate = Format(DateAdd("d", -7, zlDatabase.Currentdate), "yyyy-MM-dd")
    mdtInEndDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mdtRecBeginDate = "1901-01-01"
    mdtRecEndDate = "1901-01-01"
    mdtOutBeginDate = "1901-01-01"
    mdtOutEndDate = "1901-01-01"

    dtpBegin(0).Value = mdtInBeginDate
    dtpBegin(1).Value = mdtInBeginDate
    dtpBegin(2).Value = mdtInBeginDate
    dtpEnd(0).Value = mdtInEndDate
    dtpEnd(1).Value = mdtInEndDate
    dtpEnd(2).Value = mdtInEndDate
    
    mstrListTitle = "病历科室运送病人病案情况"
    Call SetRefresh
    
    If mlngModule = 201 Then   '病案接收菜单
        If InStr(mstrPrivs, ";增加;") <> 0 Then
            With vfgList
                .Editable = flexEDKbdMouse
            End With
            cmdAll.Visible = True
            cmdNoAll.Visible = True
        Else
            cmdAll.Visible = False
            cmdNoAll.Visible = False
        End If
    Else
        If InStr(mstrPrivs, "基本;") <> 0 Then
            With vfgList
                .Editable = flexEDKbdMouse
            End With
            cmdAll.Visible = True
            cmdNoAll.Visible = True
        Else
            cmdAll.Visible = False
            cmdNoAll.Visible = False
        End If

    End If
    
    Call zlDatabase.ShowReportMenu(Me, 300, mlngModule, gstrPrivs)
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Err = 0
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 13840 Then Me.Width = 13840
    If Me.Height < 8660 Then Me.Height = 8660  '9020 -360
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SavedkpMain(Me.Caption & "-" & mlngModule, "dkpMain")
    Call SaveHead(vfgList, 1)
    SaveWinState Me, App.ProductName
    Unload Me
End Sub

Private Sub picFilter_Resize()
    Dim lngWidth As Long
    Dim lngTemp As Long
    On Error Resume Next
    lngWidth = picFilter.ScaleWidth
    
    With fraPatient
        lngTemp = lngWidth - .Left
        If lngTemp > 0 Then
            .Width = lngWidth - .Left
        Else
            .Width = 0
        End If
    End With
    
    With fraSong
        lngTemp = lngWidth - .Left
        If lngTemp > 0 Then
            .Width = lngWidth - .Left
        Else
            .Width = 0
        End If
    End With
    
    With cboOutDept
        lngTemp = fraPatient.Width - .Left - 200
        If lngTemp > 0 Then
            .Width = lngTemp
        Else
            .Width = 0
        End If
        .SelLength = 0
    End With
    
    With dtpBegin(0)
        lngTemp = fraPatient.Width - .Left - 200
        If lngTemp > 0 Then
            .Width = lngTemp
        Else
            .Width = 0
        End If
        dtpEnd(0).Width = .Width
        dtpBegin(1).Width = .Width
        dtpEnd(1).Width = .Width
        dtpBegin(2).Width = .Width
        dtpEnd(2).Width = .Width
        txtBeginNo.Width = .Width
        txtEndNo.Width = .Width
    End With
    
'    lblApplyman.Top = 4845
    With txtApplyman
        .Top = lblApplyman.Top - 45
        lngTemp = fraPatient.Width - .Left - 200 'lngWidth - .Left
        If lngTemp > 0 Then
            .Width = lngTemp 'lngWidth - .Left
        Else
            .Width = 0
        End If
    End With
    
    With txtAuditingMan
        lngTemp = fraPatient.Width - .Left - 200 'lngWidth - .Left
        If lngTemp > 0 Then
            .Width = lngTemp 'lngWidth - .Left
        Else
            .Width = 0
        End If
'        If lngWidth - .Left > 0 Then
'            .Width = lngWidth - .Left
'        Else
'            .Width = 0
'        End If
    End With
    
    With cmdRefresh
        .Top = 5520 '6120
        lngTemp = lngWidth - .Width - 100
        If lngTemp > 0 Then
            .Left = lngTemp
            If .Left - 100 - cmdNoAll.Width > 0 Then
                cmdNoAll.Left = .Left - 100 - cmdNoAll.Width
                If cmdNoAll.Left - 100 - cmdAll.Width > 0 Then
                    cmdAll.Left = cmdNoAll.Left - 100 - cmdAll.Width
                Else
                    cmdAll.Left = 0
                End If
            Else
                cmdNoAll.Left = 0
                cmdAll.Left = 0
            End If
        Else
            .Left = 0
            cmdNoAll.Left = 0
            cmdAll.Left = 0
        End If
    End With
    
End Sub

Private Sub picList_Resize()
    vfgList.Top = picList.ScaleTop
    vfgList.Left = picList.ScaleLeft
    vfgList.Width = picList.ScaleWidth
    vfgList.Height = picList.ScaleHeight
End Sub

Private Sub initVfgList()
    Dim strHead As String
    
    strHead = "序号,500,4,1;选择,500,4,1;住院号,1500,1,1;姓名,900,1,0;性别,500,4,0;年龄,500,7,0;住院次数,900,7,0;入院科室,1200,1,0;" & _
              "入院时间,1100,1,0;出院科室,1200,1,0;出院时间,1100,1,0;病案状态,1100,1,0;运送人,900,1,0;接收人,900,1,0;接收时间,1100,1,0;编目时间,1100,1,0;记录时间,1100,1,0;" & _
              "状态,800,4,0;出生日期,1100,1,0;家庭地址,1350,1,0;出院科室id,0,7,-1;病人ID,0,7,-1;主页id,0,7,-1;病案状态值,0,7,-1;提交ID,0,7,-1;提交次数,0,7,-1"
    Call SetVsFlexGridChangeHead(strHead, vfgList, 1)
End Sub

Private Sub SetInitVfgListFormat(ByVal vsGrid As VSFlexGrid)
    Dim i As Long
    With vsGrid
        .ColDataType(.ColIndex("选择")) = flexDTBoolean
        .ForeColorSel = .CellForeColor
        .ExplorerBar = flexExSortShowAndMove
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub GetListData()
'病案接收查询
    Dim strSQL As String
    Dim strBillHead As String
    Dim rsTemp As New ADODB.Recordset
    Dim strBeginNo As String
    Dim strEndNo As String
    Dim strApplyMan As String
    Dim strAuditingMan As String
    Dim strNo As String
    Dim strStatus As String
    Dim i As Long
    Dim strFind As String
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngApplyId As Long
    Dim cllTemp As New Collection
    
    strBeginNo = Trim(txtBeginNo.Text)
    strEndNo = Trim(txtEndNo.Text)
    strApplyMan = Trim(txtApplyman.Text)
    strAuditingMan = Trim(txtAuditingMan.Text)
            
    If gstrMatchMethod = "0" Then
        LfPBF = "%"
        RgPbf = "%"
    Else
        LfPBF = ""
        RgPbf = "%"
    End If
        
    If Trim(strBeginNo) <> "" Then
        If InStr(1, strBeginNo, "'") <> 0 Then
            MsgBox "开始住院号中含有非法字符！", vbInformation, gstrSysName
            If Me.txtBeginNo.Enabled Then Me.txtBeginNo.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(strBeginNo)) > 18 Then
            MsgBox "开始住院号超长,最多能输入9个汉字或18个字符!", vbInformation + vbOKOnly, gstrSysName
            If txtBeginNo.Enabled Then txtBeginNo.SetFocus
            Exit Sub
        End If
    End If
    If Trim(strEndNo) <> "" Then
        If InStr(1, strEndNo, "'") <> 0 Then
            MsgBox "结束住院号中含有非法字符！", vbInformation, gstrSysName
            If Me.txtEndNo.Enabled Then Me.txtEndNo.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(txtEndNo)) > 18 Then
            MsgBox "结束住院号超长,最多能输入9个汉字或18个字符!", vbInformation + vbOKOnly, gstrSysName
            If txtEndNo.Enabled Then txtEndNo.SetFocus
            Exit Sub
        End If
    End If
    
     If Trim(strApplyMan) <> "" Then
        If InStr(1, strApplyMan, "'") <> 0 Then
            MsgBox "运送人中含有非法字符！", vbInformation, gstrSysName
            If Me.txtApplyman.Enabled Then Me.txtApplyman.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(strApplyMan)) > 20 Then
            MsgBox "运送人超长,最多能输入10个汉字或20个字符!", vbInformation + vbOKOnly, gstrSysName
            If txtApplyman.Enabled Then txtApplyman.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(strAuditingMan) <> "" Then
       If InStr(1, strAuditingMan, "'") <> 0 Then
           MsgBox "接收人中含有非法字符！", vbInformation, gstrSysName
           If Me.txtAuditingMan.Enabled Then Me.txtAuditingMan.SetFocus
           Exit Sub
       End If
       If zlCommFun.ActualLen(Trim(txtEndNo)) > 20 Then
            MsgBox "接收人超长,最多能输入10个汉字或20个字符!", vbInformation + vbOKOnly, gstrSysName
            If txtAuditingMan.Enabled Then txtAuditingMan.SetFocus
            Exit Sub
        End If
    End If
    
    mdtInBeginDate = "1901-01-01"
    mdtInEndDate = "1901-01-01"
    mdtRecBeginDate = "1901-01-01"
    mdtRecEndDate = "1901-01-01"
    mdtOutBeginDate = "1901-01-01"
    mdtOutEndDate = "1901-01-01"
        
    If strEndNo <> "" And strBeginNo <> "" Then
'        strFind = strFind & " and A.住院号 Between " & strBeginNo & " And " & strEndNo
        strFind = strFind & " and A.住院号 Between  [1]  And [2] "
    End If
    
    If strEndNo <> "" And strBeginNo = "" Then
'        strFind = strFind & " and A.住院号 = " & strEndNo
        strFind = strFind & " and A.住院号 = [2] "
    End If
    
    If strEndNo = "" And strBeginNo <> "" Then
'        strFind = strFind & " and A.住院号 = " & strBeginNo
        strFind = strFind & " and A.住院号 = [1] "
    End If
        
    If strApplyMan <> "" Then
'        strFind = strFind & " and A.运送人 like '" & LfPBF & strApplyMan & RgPbf & "'"
        strFind = strFind & " and A.运送人 like [3]"
    End If
    If strAuditingMan <> "" Then
'       strFind = strFind & " and A.接收人 like '" & LfPBF & strAuditingMan & RgPbf & "'"
       strFind = strFind & " and A.接收人 like [4]"
    End If
    If strBeginNo <> "" Then
        AddArray cllTemp, strBeginNo
    Else
        AddArray cllTemp, "0"
    End If
    If strEndNo <> "" Then
        AddArray cllTemp, strEndNo
    Else
        AddArray cllTemp, "0"
    End If
'    AddArray cllTemp, strBeginNo
'    AddArray cllTemp, strEndNo
    AddArray cllTemp, LfPBF & strApplyMan & RgPbf
    AddArray cllTemp, LfPBF & strAuditingMan & RgPbf
    
    
    If chkIncept.Value = 1 And chkRecord.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.接收时间 Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'            & " or (A.编目日期 Between To_Date('" & Format(dtpBegin(2), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(2), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS'))"
'            strFind = strFind & " Or (A.出院日期 Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
'
            strFind = strFind & " And ((A.接收时间 Between [7] And [8]) " _
            & " or (A.编目日期 Between To_Date([9]) And [10])"
            strFind = strFind & " Or (A.出院日期 Between [5] And [6])) "
            
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And ((A.接收时间 Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'            & " or (A.编目日期 Between To_Date('" & Format(dtpBegin(2), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(2), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')))"
            strFind = strFind & " And ((A.接收时间 Between [7] And [8]) " _
            & " or (A.编目日期 Between [9] And [10]))"
        End If
        mdtInBeginDate = Format(dtpBegin(1), "yyyy-mm-dd")
        mdtInEndDate = Format(dtpEnd(1), "yyyy-mm-dd")
        mdtRecBeginDate = Format(dtpBegin(2), "yyyy-mm-dd")
        mdtRecEndDate = Format(dtpEnd(2), "yyyy-mm-dd")
    ElseIf chkIncept.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.接收时间 Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','yyyy-MM-dd HH24:MI:SS')) "
'            strFind = strFind & " Or (A.出院日期 Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
            strFind = strFind & " And ((A.接收时间 Between [7] And [8]) "
            strFind = strFind & " Or (A.出院日期 Between [5] And [6])) "
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And A.接收时间 Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','yyyy-MM-dd HH24:MI:SS') "
            strFind = strFind & " And A.接收时间 Between [7] And [8] "
        End If
        mdtInBeginDate = Format(dtpBegin(1), "yyyy-mm-dd")
        mdtInEndDate = Format(dtpEnd(1), "yyyy-mm-dd")
    ElseIf chkRecord.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.编目日期 Between To_Date('" & Format(dtpBegin(2).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(2).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
'            strFind = strFind & " Or (A.出院日期 Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
            strFind = strFind & " And ((A.编目日期 Between [9] And [10]) "
            strFind = strFind & " Or (A.出院日期 Between [5] And [6])) "
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And (A.编目日期 Between To_Date('" & Format(dtpBegin(2).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(2).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
            strFind = strFind & " And (A.编目日期 Between [9] And [10]) "
        End If
        mdtRecBeginDate = Format(dtpBegin(2), "yyyy-mm-dd")
        mdtRecEndDate = Format(dtpEnd(2), "yyyy-mm-dd")
'    Else
'        strFind = strFind & " And A.接收时间 is null and A.记录时间 is null "
    End If
    
    If chkOutHp.Value = 1 And chkIncept.Value = 0 And chkRecord.Value = 0 Then
'        strFind = strFind & " And (A.出院日期 Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
        strFind = strFind & " And (A.出院日期 Between [5] And [6]) "
        mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
        mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
    End If
    
    AddArray cllTemp, Format(mdtOutBeginDate, "yyyy-mm-dd") & " 00:00:00" ' ,yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtOutEndDate, "yyyy-mm-dd") & " 23:59:59 " ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtInBeginDate, "yyyy-mm-dd") & " 00:00:00" ' ,yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtInEndDate, "yyyy-mm-dd") & " 23:59:59" ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtRecBeginDate, "yyyy-mm-dd") & " 00:00:00" ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtRecEndDate, "yyyy-mm-dd") & " 23:59:59" ',yyyy-MM-dd HH24:MI:SS"

    
    If mintDelete = 1 Then
        strFind = mstrFind
        Set cllTemp = mcllTemp
    Else
        mstrFind = strFind
        Set mcllTemp = cllTemp
        If chkIncept.Value = 0 And chkRecord.Value = 0 And chkOutHp.Value = 0 Then
            MsgBox "必须选择一个出院日期或者接收日期或者记录日期!", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    vfgList.Redraw = False
    Call zlCommFun.ShowFlash("正在搜索病案相应的记录,请稍候 ...", Me)
'    DoEvents
    Screen.MousePointer = vbHourglass
    
'     And (U.编目日期 is null Or (U.编目日期 is not null and U.出院日期 >= to_date('" & strNoShowDate & "','yyyy-mm-dd')))
    mlngTempId = 0
    
    If mlngApplyId = 0 Then
        strBillHead = "" & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期, Null As 运送人, Null As 接收人, Null As 接收时间, Null As 记录时间," & _
        "      Decode(Nvl(to_char(U.编目日期), '0'), '0', '未接收', '已编目未登记') As 状态 " & _
        " From 病案主页 U, 病人信息 X " & _
        " Where U.病人ID = X.病人ID And Not Exists" & _
        "      (Select 1 From 病案接收记录 A Where A.病人id = U.病人id And A.主页id = U.主页id) And U.病人性质 = 0 And U.主页ID <> 0 And U.出院日期 is not null " & _
               IIf(mblnShow, " And (U.编目日期 is null )", "")
        strBillHead = strBillHead & " Union All "
        strBillHead = strBillHead & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间," & _
        "      Decode(Nvl(to_char(U.编目日期), '0'), '0', '已接收', '已编目') As 状态" & _
        " From 病案主页 U, 病人信息 X,病案接收记录 A" & _
        " Where U.病人id = X.病人id And U.病人性质 = 0 And U.主页ID <> 0 And " & _
        "       A.病人id = U.病人id And A.主页ID = U.主页ID"
        strSQL = "" & _
        "   Select distinct A.病人id, A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
        "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
        "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室,出院科室id," & _
        "    A.住院天数, A.费用和, A.是否随诊,A.编目员, A.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间,A.状态" & _
        "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
        "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & strFind & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
        "    Order by A.出院日期 desc "
    Else
        strBillHead = "" & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期, Null As 运送人, Null As 接收人, Null As 接收时间, Null As 记录时间," & _
        "      Decode(Nvl(to_char(U.编目日期), '0'), '0', '未接收', '已编目未登记') As 状态 " & _
        " From 病案主页 U, 病人信息 X " & _
        " Where U.病人ID = X.病人ID And U.出院科室id =[11] And Not Exists" & _
        "      (Select 1 From 病案接收记录 A Where A.病人id = U.病人id And A.主页id = U.主页id) And U.病人性质 = 0 And U.主页ID <> 0 And U.出院日期 is not null " & _
               IIf(mblnShow, " And (U.编目日期 is null )", "")
        strBillHead = strBillHead & " Union All "
        strBillHead = strBillHead & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间," & _
        "      Decode(Nvl(to_char(U.编目日期), '0'), '0', '已接收', '已编目') As 状态" & _
        " From 病案主页 U, 病人信息 X,病案接收记录 A" & _
        " Where U.病人id = X.病人id And U.病人性质 = 0 And U.主页ID <> 0 And U.出院科室id =[11] And " & _
        "       A.病人id = U.病人id And A.主页ID = U.主页ID"
        strSQL = "" & _
        "   Select distinct A.病人id, A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
        "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
        "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室,出院科室id," & _
        "    A.住院天数, A.费用和, A.是否随诊,A.编目员, A.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间,A.状态" & _
        "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
        "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & strFind & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
        "    Order by A.出院日期 desc "
    End If
            
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(cllTemp(1)), CLng(cllTemp(2)), cllTemp(3), cllTemp(4), CDate(cllTemp(5)), _
                        CDate(cllTemp(6)), CDate(cllTemp(7)), CDate(cllTemp(8)), CDate(cllTemp(9)), CDate(cllTemp(10)), mlngApplyId)
    With vfgList
        Call initVfgList
        .Rows = IIf(rsTemp.EOF, 0, rsTemp.RecordCount) + 1
        If Not rsTemp.EOF Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("序号")) = i
                .TextMatrix(i, .ColIndex("住院号")) = IIf(IsNull(rsTemp!住院号), 0, rsTemp!住院号)
                .TextMatrix(i, .ColIndex("姓名")) = IIf(IsNull(rsTemp!姓名), "", rsTemp!姓名)
                .TextMatrix(i, .ColIndex("性别")) = IIf(IsNull(rsTemp!性别), "", rsTemp!性别)
                .TextMatrix(i, .ColIndex("年龄")) = IIf(IsNull(rsTemp!年龄), "", rsTemp!年龄)
                .TextMatrix(i, .ColIndex("住院次数")) = IIf(IsNull(rsTemp!总住院次数), "", rsTemp!总住院次数)
                .TextMatrix(i, .ColIndex("入院时间")) = IIf(IsNull(rsTemp!入院日期), "", Format(rsTemp!入院日期, "yyyy-MM-dd hh:mm:ss"))
                .TextMatrix(i, .ColIndex("入院科室")) = IIf(IsNull(rsTemp!入院科室), "", rsTemp!入院科室)
                .TextMatrix(i, .ColIndex("出院科室")) = IIf(IsNull(rsTemp!出院科室), "", rsTemp!出院科室)
                .TextMatrix(i, .ColIndex("运送人")) = IIf(IsNull(rsTemp!运送人), "", rsTemp!运送人)
                .TextMatrix(i, .ColIndex("接收人")) = IIf(IsNull(rsTemp!接收人), "", rsTemp!接收人)
                .TextMatrix(i, .ColIndex("接收时间")) = IIf(IsNull(rsTemp!接收时间), "", Format(rsTemp!接收时间, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("编目时间")) = IIf(IsNull(rsTemp!编目日期), "", Format(rsTemp!编目日期, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("记录时间")) = IIf(IsNull(rsTemp!记录时间), "", Format(rsTemp!记录时间, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("状态")) = IIf(IsNull(rsTemp!状态), "", rsTemp!状态)
                .TextMatrix(i, .ColIndex("出院时间")) = IIf(IsNull(rsTemp!出院日期), "", Format(rsTemp!出院日期, "yyyy-MM-dd hh:mm:ss"))
                .TextMatrix(i, .ColIndex("出生日期")) = IIf(IsNull(rsTemp!出生日期), "", Format(rsTemp!出生日期, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("家庭地址")) = IIf(IsNull(rsTemp!家庭地址), "", rsTemp!家庭地址)
                .TextMatrix(i, .ColIndex("病人ID")) = IIf(IsNull(rsTemp!病人ID), 0, rsTemp!病人ID)
                .TextMatrix(i, .ColIndex("主页id")) = IIf(IsNull(rsTemp!主页ID), 0, rsTemp!主页ID)
                .TextMatrix(i, .ColIndex("出院科室id")) = IIf(IsNull(rsTemp!出院科室ID), 0, rsTemp!出院科室ID)
                
                rsTemp.MoveNext
                strStatus = Trim(.TextMatrix(i, .ColIndex("状态")))
                Select Case strStatus
                Case "未接收"
                    .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(0)
                Case "已接收"
                    .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(1)
                Case "已编目", "已编目未登记"
                    .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(2)
                End Select

                If InStr(mstrPrivs, ";增加;") <> 0 And strStatus = "未接收" Then
'                    lngApplyId = Val(.TextMatrix(i, .ColIndex("出院科室id")))
'                    If mlngTempId = 0 Then mlngTempId = lngApplyId
'                    If lngApplyId <> mlngTempId Then
                        .TextMatrix(i, .ColIndex("选择")) = 0
'                    Else
'                        .TextMatrix(i, .ColIndex("选择")) = -1
'                    End If
                    
                Else
                    .TextMatrix(i, .ColIndex("选择")) = 0
                End If
            Next
        End If


    End With
    
    Call zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    Call GetStatusCount
   
    Call SetInitVfgListFormat(vfgList)
    Call RestoreHead(vfgList, 1)
    rsTemp.Close
    vfgList.Redraw = True
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetStatusCount()
    Dim lng未接收 As Long
    Dim lng已接收 As Long
    Dim lng已编目 As Long
    Dim lngCount As Long
    
    With vfgList
        For lngCount = 1 To .Rows - 1
            Select Case Trim(.TextMatrix(lngCount, .ColIndex("状态")))
            Case "未接收"
                lng未接收 = lng未接收 + 1
            Case "已接收"
                lng已接收 = lng已接收 + 1
            Case "已编目", "已编目未登记"
                lng已编目 = lng已编目 + 1
            End Select
        Next
    End With
    
    stbThis.Panels(2).Text = "只有状态为未接收及具体出院科室时才能进行选择操作！当前共有" & vfgList.Rows - 1 & "条病人病案！其中未接收:" & lng未接收 & "条、已接收:" & lng已接收 & "条、已编目:" & lng已编目 & "条。"
End Sub

Private Sub GetListDataElectron()
'电子病案接收查询
    Dim strSQL As String
    Dim strBillHead As String
    Dim rsTemp As New ADODB.Recordset
    Dim strBeginNo As String
    Dim strEndNo As String
    Dim strApplyMan As String
    Dim strAuditingMan As String
    Dim strNo As String
    Dim strStatus As String
    Dim i As Long
    Dim strFind As String
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngApplyId As Long
    Dim cllTemp As New Collection
    
    strBeginNo = Trim(txtBeginNo.Text)
    strEndNo = Trim(txtEndNo.Text)
    strApplyMan = Trim(txtApplyman.Text)
    strAuditingMan = Trim(txtAuditingMan.Text)
            
    If gstrMatchMethod = "0" Then
        LfPBF = "%"
        RgPbf = "%"
    Else
        LfPBF = ""
        RgPbf = "%"
    End If
        
    If Trim(strBeginNo) <> "" Then
        If InStr(1, strBeginNo, "'") <> 0 Then
            MsgBox "开始住院号中含有非法字符！", vbInformation, gstrSysName
            If Me.txtBeginNo.Enabled Then Me.txtBeginNo.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(strBeginNo)) > 18 Then
            MsgBox "开始住院号超长,最多能输入9个汉字或18个字符!", vbInformation + vbOKOnly, gstrSysName
            If txtBeginNo.Enabled Then txtBeginNo.SetFocus
            Exit Sub
        End If
    End If
    If Trim(strEndNo) <> "" Then
        If InStr(1, strEndNo, "'") <> 0 Then
            MsgBox "结束住院号中含有非法字符！", vbInformation, gstrSysName
            If Me.txtEndNo.Enabled Then Me.txtEndNo.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(txtEndNo)) > 18 Then
            MsgBox "结束住院号超长,最多能输入9个汉字或18个字符!", vbInformation + vbOKOnly, gstrSysName
            If txtEndNo.Enabled Then txtEndNo.SetFocus
            Exit Sub
        End If
    End If
    
     If Trim(strApplyMan) <> "" Then
        If InStr(1, strApplyMan, "'") <> 0 Then
            MsgBox "运送人中含有非法字符！", vbInformation, gstrSysName
            If Me.txtApplyman.Enabled Then Me.txtApplyman.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(strApplyMan)) > 20 Then
            MsgBox "运送人超长,最多能输入10个汉字或20个字符!", vbInformation + vbOKOnly, gstrSysName
            If txtApplyman.Enabled Then txtApplyman.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(strAuditingMan) <> "" Then
       If InStr(1, strAuditingMan, "'") <> 0 Then
           MsgBox "接收人中含有非法字符！", vbInformation, gstrSysName
           If Me.txtAuditingMan.Enabled Then Me.txtAuditingMan.SetFocus
           Exit Sub
       End If
       If zlCommFun.ActualLen(Trim(txtEndNo)) > 20 Then
            MsgBox "接收人超长,最多能输入10个汉字或20个字符!", vbInformation + vbOKOnly, gstrSysName
            If txtAuditingMan.Enabled Then txtAuditingMan.SetFocus
            Exit Sub
        End If
    End If
    
    mdtInBeginDate = "1901-01-01"
    mdtInEndDate = "1901-01-01"
    mdtRecBeginDate = "1901-01-01"
    mdtRecEndDate = "1901-01-01"
    mdtOutBeginDate = "1901-01-01"
    mdtOutEndDate = "1901-01-01"
        
    If strEndNo <> "" And strBeginNo <> "" Then
'        strFind = strFind & " and A.住院号 Between " & strBeginNo & " And " & strEndNo
        strFind = strFind & " and A.住院号 Between  [1]  And [2] "
    End If
    
    If strEndNo <> "" And strBeginNo = "" Then
'        strFind = strFind & " and A.住院号 = " & strEndNo
        strFind = strFind & " and A.住院号 = [2] "
    End If
    
    If strEndNo = "" And strBeginNo <> "" Then
'        strFind = strFind & " and A.住院号 = " & strBeginNo
        strFind = strFind & " and A.住院号 = [1] "
    End If
        
    If strApplyMan <> "" Then
'        strFind = strFind & " and A.运送人 like '" & LfPBF & strApplyMan & RgPbf & "'"
        strFind = strFind & " and A.运送人 like [3]"
    End If
    If strAuditingMan <> "" Then
'       strFind = strFind & " and A.接收人 like '" & LfPBF & strAuditingMan & RgPbf & "'"
       strFind = strFind & " and A.接收人 like [4]"
    End If
    If strBeginNo <> "" Then
        AddArray cllTemp, strBeginNo
    Else
        AddArray cllTemp, "0"
    End If
    If strEndNo <> "" Then
        AddArray cllTemp, strEndNo
    Else
        AddArray cllTemp, "0"
    End If
'    AddArray cllTemp, strBeginNo
'    AddArray cllTemp, strEndNo
    AddArray cllTemp, LfPBF & strApplyMan & RgPbf
    AddArray cllTemp, LfPBF & strAuditingMan & RgPbf
    
    
    If chkIncept.Value = 1 And chkRecord.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.接收时间 Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'            & " or (A.编目日期 Between To_Date('" & Format(dtpBegin(2), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(2), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS'))"
'            strFind = strFind & " Or (A.出院日期 Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
'
            strFind = strFind & " And ((A.接收时间 Between [7] And [8]) " _
            & " or (A.编目日期 Between To_Date([9]) And [10])"
            strFind = strFind & " Or (A.出院日期 Between [5] And [6])) "
            
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And ((A.接收时间 Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'            & " or (A.编目日期 Between To_Date('" & Format(dtpBegin(2), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(2), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')))"
            strFind = strFind & " And ((A.接收时间 Between [7] And [8]) " _
            & " or (A.编目日期 Between [9] And [10]))"
        End If
        mdtInBeginDate = Format(dtpBegin(1), "yyyy-mm-dd")
        mdtInEndDate = Format(dtpEnd(1), "yyyy-mm-dd")
        mdtRecBeginDate = Format(dtpBegin(2), "yyyy-mm-dd")
        mdtRecEndDate = Format(dtpEnd(2), "yyyy-mm-dd")
    ElseIf chkIncept.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.接收时间 Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','yyyy-MM-dd HH24:MI:SS')) "
'            strFind = strFind & " Or (A.出院日期 Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
            strFind = strFind & " And ((A.接收时间 Between [7] And [8]) "
            strFind = strFind & " Or (A.出院日期 Between [5] And [6])) "
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And A.接收时间 Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','yyyy-MM-dd HH24:MI:SS') "
            strFind = strFind & " And A.接收时间 Between [7] And [8] "
        End If
        mdtInBeginDate = Format(dtpBegin(1), "yyyy-mm-dd")
        mdtInEndDate = Format(dtpEnd(1), "yyyy-mm-dd")
    ElseIf chkRecord.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.编目日期 Between To_Date('" & Format(dtpBegin(2).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(2).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
'            strFind = strFind & " Or (A.出院日期 Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
            strFind = strFind & " And ((A.编目日期 Between [9] And [10]) "
            strFind = strFind & " Or (A.出院日期 Between [5] And [6])) "
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And (A.编目日期 Between To_Date('" & Format(dtpBegin(2).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(2).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
            strFind = strFind & " And (A.编目日期 Between [9] And [10]) "
        End If
        mdtRecBeginDate = Format(dtpBegin(2), "yyyy-mm-dd")
        mdtRecEndDate = Format(dtpEnd(2), "yyyy-mm-dd")
'    Else
'        strFind = strFind & " And A.接收时间 is null and A.记录时间 is null "
    End If
    
    If chkOutHp.Value = 1 And chkIncept.Value = 0 And chkRecord.Value = 0 Then
'        strFind = strFind & " And (A.出院日期 Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
        strFind = strFind & " And (A.出院日期 Between [5] And [6]) "
        mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
        mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
    End If
    
    AddArray cllTemp, Format(mdtOutBeginDate, "yyyy-mm-dd") & " 00:00:00" ' ,yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtOutEndDate, "yyyy-mm-dd") & " 23:59:59 " ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtInBeginDate, "yyyy-mm-dd") & " 00:00:00" ' ,yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtInEndDate, "yyyy-mm-dd") & " 23:59:59" ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtRecBeginDate, "yyyy-mm-dd") & " 00:00:00" ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtRecEndDate, "yyyy-mm-dd") & " 23:59:59" ',yyyy-MM-dd HH24:MI:SS"

    
    If mintDelete = 1 Then
        strFind = mstrFind
        Set cllTemp = mcllTemp
    Else
        mstrFind = strFind
        Set mcllTemp = cllTemp
        If chkIncept.Value = 0 And chkRecord.Value = 0 And chkOutHp.Value = 0 Then
            MsgBox "必须选择一个出院日期或者接收日期或者记录日期!", vbInformation, gstrSysName
            If chkOutHp.Enabled And chkOutHp.Visible Then
                chkOutHp.SetFocus
            End If
            Exit Sub
        End If
    End If
    vfgList.Redraw = False
    Call zlCommFun.ShowFlash("正在搜索病案相应的记录,请稍候 ...", Me)
    Screen.MousePointer = vbHourglass
    mlngTempId = 0
    
    If mlngApplyId = 0 Then
        strBillHead = "" & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期, Null As 运送人, Null As 接收人, Null As 接收时间, Null As 记录时间,Decode(Nvl(U.病案状态,1),1,'提交待收',10,'接收待审',2,'拒绝接收',3,'正在审查',4,'审查反馈',5,'审查归档',6,'审查整改',13,'正在抽查',14,'抽查反馈',16,'抽查整改') as 病案状态,Nvl(U.病案状态,1) as 病案状态值,C.ID as 提交ID,(Select Count(1) From 病案提交记录 F Where F.病人ID=U.病人ID and F.主页ID=U.主页ID And F.记录状态=2) as 提交次数," & _
        "      Decode(Nvl(to_char(U.编目日期), '0'), '0', '未接收', '已编目未登记') As 状态 " & _
        " From 病案主页 U, 病人信息 X,病案提交记录 C " & _
        " Where U.病人ID = X.病人ID And U.病人ID = C.病人ID And U.主页ID = C.主页ID And Not Exists" & _
        "      (Select 1 From 病案接收记录 A Where A.病人id = U.病人id And A.主页id = U.主页id) And U.主页ID <> 0 And U.出院日期 is not null And C.记录状态=1 " & _
               IIf(mblnShow, " And (U.编目日期 is null )", "")
        strBillHead = strBillHead & " Union All "
        strBillHead = strBillHead & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间,Decode(Nvl(U.病案状态,1),1,'提交待收',10,'接收待审',2,'拒绝接收',3,'正在审查',4,'审查反馈',5,'审查归档',6,'审查整改',13,'正在抽查',14,'抽查反馈',16,'抽查整改') as 病案状态,Nvl(U.病案状态,1) as 病案状态值,C.ID as 提交ID,0 as 提交次数," & _
        "      Decode(Nvl(to_char(U.编目日期), '0'), '0', '已接收', '已编目') As 状态" & _
        " From 病案主页 U, 病人信息 X,病案接收记录 A,病案提交记录 C" & _
        " Where U.病人id = X.病人id And U.主页ID <> 0  And A.病人ID = C.病人ID And A.主页ID = C.主页ID And  C.记录状态<>2 And " & _
        "       A.病人id = U.病人id And A.主页ID = U.主页ID"
        
        strSQL = "" & _
        "   Select distinct A.病人id, A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
        "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
        "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室,出院科室id," & _
        "    A.住院天数, A.费用和, A.是否随诊,A.编目员, A.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间,A.状态,A.病案状态,A.病案状态值,A.提交ID,A.提交次数" & _
        "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
        "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & strFind & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
        "    Order by A.出院日期 desc "
    Else
        strBillHead = "" & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期, Null As 运送人, Null As 接收人, Null As 接收时间, Null As 记录时间, Decode(Nvl(U.病案状态,1),1,'提交待收',10,'接收待审',2,'拒绝接收',3,'正在审查',4,'审查反馈',5,'审查归档',6,'审查整改',13,'正在抽查',14,'抽查反馈',16,'抽查整改') as 病案状态,Nvl(U.病案状态,1) as 病案状态值,C.ID as 提交ID,(Select Count(1) From 病案提交记录 F Where F.病人ID=U.病人ID and F.主页ID=U.主页ID And F.记录状态=2) as 提交次数," & _
        "      Decode(Nvl(to_char(U.编目日期), '0'), '0', '未接收', '已编目未登记') As 状态 " & _
        " From 病案主页 U, 病人信息 X,病案提交记录 C " & _
        " Where U.病人ID = X.病人ID And U.出院科室id =[11] And U.病人ID = C.病人ID And U.主页ID = C.主页ID And Not Exists" & _
        "      (Select 1 From 病案接收记录 A Where A.病人id = U.病人id And A.主页id = U.主页id) And U.主页ID <> 0 And U.出院日期 is not null And C.记录状态=1 " & _
               IIf(mblnShow, " And (U.编目日期 is null )", "")
        strBillHead = strBillHead & " Union All "
        strBillHead = strBillHead & _
        " Select Distinct X.病人id, U.主页id, U.住院号, X.姓名, X.性别, X.年龄, U.主页id As 总住院次数, X.出生日期, X.出生地点," & _
        "      X.身份证号, X.职业, X.婚姻状况, X.家庭地址, X.家庭电话, X.联系人姓名, X.联系人关系, X.联系人地址," & _
        "      X.联系人电话, X.工作单位, X.单位电话, U.出院日期, U.入院日期, U.入院科室id, U.出院科室id," & _
        "      U.住院天数, U.费用和, Decode(U.随诊标志, 1, '是', 2, '是', 3, '是', '') As 是否随诊," & _
        "      U.编目员姓名 As 编目员, U.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间, Decode(Nvl(U.病案状态,1),1,'提交待收',10,'接收待审',2,'拒绝接收',3,'正在审查',4,'审查反馈',5,'审查归档',6,'审查整改',13,'正在抽查',14,'抽查反馈',16,'抽查整改') as 病案状态,Nvl(U.病案状态,1) as 病案状态值,C.ID as 提交ID,0 as 提交次数," & _
        "      Decode(Nvl(to_char(U.编目日期), '0'), '0', '已接收', '已编目') As 状态" & _
        " From 病案主页 U, 病人信息 X,病案接收记录 A,病案提交记录 C" & _
        " Where U.病人id = X.病人id And U.主页ID <> 0 And U.出院科室id =[11] And A.病人ID = C.病人ID And A.主页ID = C.主页ID And C.记录状态<>2 And  " & _
        "       A.病人id = U.病人id And A.主页ID = U.主页ID"
        strSQL = "" & _
        "   Select distinct A.病人id, A.主页id, A.住院号, A.姓名, A.性别, A.年龄, A.总住院次数, A.出生日期, A.出生地点," & _
        "    A.身份证号, A.职业, A.婚姻状况, A.家庭地址, A.家庭电话, A.联系人姓名, A.联系人关系, A.联系人地址," & _
        "    A.联系人电话, A.工作单位, A.单位电话, A.出院日期, A.入院日期, B1.名称 As 入院科室, B2.名称 As 出院科室,出院科室id," & _
        "    A.住院天数, A.费用和, A.是否随诊,A.编目员, A.编目日期, A.运送人, A.接收人, A.接收时间, A.记录时间,A.状态,A.病案状态,A.病案状态值,A.提交ID,A.提交次数" & _
        "    From (" & strBillHead & ") A,部门表 B1, 部门表 B2" & _
        "    Where A.入院科室id=B1.id And A.出院科室id=B2.id " & strFind & zl_获取站点限制(True, "B1") & zl_获取站点限制(True, "B2") & _
        "    Order by A.出院日期 desc "
    End If
            
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(cllTemp(1)), CLng(cllTemp(2)), cllTemp(3), cllTemp(4), CDate(cllTemp(5)), _
                        CDate(cllTemp(6)), CDate(cllTemp(7)), CDate(cllTemp(8)), CDate(cllTemp(9)), CDate(cllTemp(10)), mlngApplyId)
    With vfgList
        Call initVfgList
        .Rows = IIf(rsTemp.EOF, 0, rsTemp.RecordCount) + 1
        If Not rsTemp.EOF Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("序号")) = i
                .TextMatrix(i, .ColIndex("住院号")) = IIf(IsNull(rsTemp!住院号), 0, rsTemp!住院号)
                .TextMatrix(i, .ColIndex("姓名")) = IIf(IsNull(rsTemp!姓名), "", rsTemp!姓名)
                .TextMatrix(i, .ColIndex("性别")) = IIf(IsNull(rsTemp!性别), "", rsTemp!性别)
                .TextMatrix(i, .ColIndex("年龄")) = IIf(IsNull(rsTemp!年龄), "", rsTemp!年龄)
                .TextMatrix(i, .ColIndex("住院次数")) = IIf(IsNull(rsTemp!总住院次数), "", rsTemp!总住院次数)
                .TextMatrix(i, .ColIndex("入院时间")) = IIf(IsNull(rsTemp!入院日期), "", Format(rsTemp!入院日期, "yyyy-MM-dd hh:mm:ss"))
                .TextMatrix(i, .ColIndex("入院科室")) = IIf(IsNull(rsTemp!入院科室), "", rsTemp!入院科室)
                .TextMatrix(i, .ColIndex("出院科室")) = IIf(IsNull(rsTemp!出院科室), "", rsTemp!出院科室)
                .TextMatrix(i, .ColIndex("运送人")) = IIf(IsNull(rsTemp!运送人), "", rsTemp!运送人)
                .TextMatrix(i, .ColIndex("接收人")) = IIf(IsNull(rsTemp!接收人), "", rsTemp!接收人)
                .TextMatrix(i, .ColIndex("接收时间")) = IIf(IsNull(rsTemp!接收时间), "", Format(rsTemp!接收时间, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("编目时间")) = IIf(IsNull(rsTemp!编目日期), "", Format(rsTemp!编目日期, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("记录时间")) = IIf(IsNull(rsTemp!记录时间), "", Format(rsTemp!记录时间, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("状态")) = IIf(IsNull(rsTemp!状态), "", rsTemp!状态)
                .TextMatrix(i, .ColIndex("出院时间")) = IIf(IsNull(rsTemp!出院日期), "", Format(rsTemp!出院日期, "yyyy-MM-dd hh:mm:ss"))
                .TextMatrix(i, .ColIndex("病案状态")) = IIf(IsNull(rsTemp!病案状态), "", rsTemp!病案状态)
                .TextMatrix(i, .ColIndex("出生日期")) = IIf(IsNull(rsTemp!出生日期), "", Format(rsTemp!出生日期, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("家庭地址")) = IIf(IsNull(rsTemp!家庭地址), "", rsTemp!家庭地址)
                .TextMatrix(i, .ColIndex("病人ID")) = IIf(IsNull(rsTemp!病人ID), 0, rsTemp!病人ID)
                .TextMatrix(i, .ColIndex("主页id")) = IIf(IsNull(rsTemp!主页ID), 0, rsTemp!主页ID)
                .TextMatrix(i, .ColIndex("出院科室id")) = IIf(IsNull(rsTemp!出院科室ID), 0, rsTemp!出院科室ID)
                .TextMatrix(i, .ColIndex("病案状态值")) = IIf(IsNull(rsTemp!病案状态值), 0, rsTemp!病案状态值)
                .TextMatrix(i, .ColIndex("提交ID")) = IIf(IsNull(rsTemp!提交Id), 0, rsTemp!提交Id)
                .TextMatrix(i, .ColIndex("提交次数")) = IIf(IsNull(rsTemp!提交次数), 0, rsTemp!提交次数)
                
                rsTemp.MoveNext
                strStatus = Trim(.TextMatrix(i, .ColIndex("状态")))
                Select Case strStatus
                Case "未接收"
                    If Val(.TextMatrix(i, .ColIndex("提交次数"))) = 0 Then
                        .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(0)
                    Else
                        .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(3)
                    End If
                Case "已接收"
                    .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(1)
                Case "已编目", "已编目未登记"
                    .Cell(flexcpPicture, i, .ColIndex("住院号")) = imgCelNo(2)
                End Select

                If InStr(mstrPrivs, "基本;") <> 0 And strStatus = "未接收" Then
'                    lngApplyId = Val(.TextMatrix(i, .ColIndex("出院科室id")))
'                    If mlngTempId = 0 Then mlngTempId = lngApplyId
'                    If lngApplyId <> mlngTempId Then
                        .TextMatrix(i, .ColIndex("选择")) = 0
'                    Else
'                        .TextMatrix(i, .ColIndex("选择")) = -1
'                    End If
                    
                Else
                    .TextMatrix(i, .ColIndex("选择")) = 0
                End If
            Next
        End If


    End With
    
    Call zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
'    stbThis.Panels(2).Text = "只有状态为未接收及具体出院科室时才能进行选择操作！当前共有" & rsTemp.RecordCount & "条病人病案数据信息！"
    Call GetStatusCount
    Call SetInitVfgListFormat(vfgList)
    Call RestoreHead(vfgList, 1)
    rsTemp.Close
    vfgList.Redraw = True
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetNoSHowData()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "" & _
        " Select min(U.出院日期) as 出院日期" & _
        " From 病案主页 U, 病案接收记录 X " & _
        " Where U.病人ID = X.病人ID And U.主页id = X.主页id And X.接收时间 = " & _
        "      (Select min(A.接收时间) From 病案接收记录 A)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        mstrNoShowDate = IIf(IsNull(rsTemp!出院日期), Format(zlDatabase.Currentdate, "yyyy-MM-dd"), Format(rsTemp!出院日期, "yyyy-MM-dd"))
    Else
        mstrNoShowDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End If
    rsTemp.Close
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
   
Private Sub SetRefresh()
    If mlngModule = 201 Then
        Call GetListData
    Else
        Call GetListDataElectron
    End If
End Sub

Private Sub SetVfgSelect()
    frm病案选择列.ShowColSet Me, "病人信息列设置", vfgList
End Sub


Private Sub RightHead(ByVal vsGrid As VSFlexGrid, intListOrDetail As Integer)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim strHeadInfo As String
    Dim vRect  As RECT
    vRect = GetControlRect(vsGrid.hWnd)
    lngLeft = vRect.Left + vsGrid.Left
    lngTop = vRect.Top + vsGrid.RowHeight(0) 'vsGrid.CellTop ' + vsGrid.CellHeight '
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, vsGrid.RowHeight(0))
    Call SaveHead(vsGrid, intListOrDetail)
End Sub

Private Sub SaveHead(ByVal vsGrid As VSFlexGrid, intListOrDetail As Integer)
    Dim strHeadInfo As String
     If intListOrDetail = 1 Then
        strHeadInfo = "病案接收列头信息"
    End If
    zl_VsGrid_SaveToPara vsGrid, Me.Caption, mlngModule, strHeadInfo, True, True
End Sub

Private Sub RestoreHead(ByVal vsGrid As VSFlexGrid, intListOrDetail As Integer)
    Dim strHeadInfo As String
    If intListOrDetail = 1 Then
        strHeadInfo = "病案接收列头信息"
    End If
    zl_VsGrid_FromParaRestore vsGrid, Me.Caption, mlngModule, strHeadInfo, True, True
End Sub

Private Sub txtApplyman_GotFocus()
    With txtApplyman
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtApplyman_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtApplyman.Text) = "" Then
            zlCommFun.PressKey (vbKeyTab)
            Exit Sub
        End If
        txtApplyman.Text = Replace(UCase(txtApplyman.Text), "'", "")
        vRect = GetControlRect(txtApplyman.hWnd)
        
        strSQL = "" & _
            "   Select 编号,简码,姓名,id " & _
            "   From 人员表 " & _
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) " & zl_获取站点限制(True) & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
            "   order by 编号"
            
        strTemp = Trim(txtApplyman.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        strTemp = LfPBF & strTemp & RgPbf
        lngHeigth = txtApplyman.Height

        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "人员选择", False, txtApplyman.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strTemp)
               
        If rsTemp Is Nothing Then
            If txtApplyman.Enabled Then
                zlCommFun.PressKey (vbKeyTab)
                Exit Sub
            End If
        End If
       
        With rsTemp
            If UCase(TypeName(txtApplyman)) = "TEXTBOX" Then
                txtApplyman = IIf(IsNull(!姓名), "", !姓名)
                zlCommFun.PressKey (vbKeyTab)
            Else
                txtApplyman.SetFocus
                txtApplyman.SelStart = 0
                txtApplyman.SelLength = Len(txtApplyman.Text)
                zlCommFun.PressKey vbKeyTab
            End If
            .Close
        End With
    End If
End Sub

Private Sub txtApplyman_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtApplyman_LostFocus()
    If zlCommFun.ActualLen(Trim(txtApplyman.Text)) > 20 Then
        MsgBox "运送人超长,最多能输入10个汉字或20个字符!", vbInformation + vbOKOnly, gstrSysName
        txtApplyman.SetFocus
        txtApplyman.SelStart = 0
        txtApplyman.SelLength = Len(txtApplyman.Text)
        If txtApplyman.Enabled Then txtApplyman.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtAuditingMan_GotFocus()
    With txtAuditingMan
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAuditingMan_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtAuditingMan.Text) = "" Then
            zlCommFun.PressKey (vbKeyTab)
            Exit Sub
        End If
        txtAuditingMan.Text = Replace(UCase(txtAuditingMan.Text), "'", "")
        vRect = GetControlRect(txtAuditingMan.hWnd)
        
        strSQL = "" & _
            "   Select 编号,简码,姓名,id " & _
            "   From 人员表 " & _
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) " & zl_获取站点限制(True) & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
            "   order by 编号"
            
        strTemp = Trim(txtAuditingMan.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        strTemp = LfPBF & strTemp & RgPbf
        lngHeigth = txtAuditingMan.Height

        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "人员选择", False, txtAuditingMan.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strTemp)
               
        If rsTemp Is Nothing Then
            If Not blnCancel Then MsgBox "没有满足条件的姓名,请检查[人员信息]!", vbInformation, gstrSysName
            If txtAuditingMan.Enabled Then
                txtAuditingMan.SetFocus
                txtAuditingMan.SelStart = 0
                txtAuditingMan.Text = ""
                Exit Sub
            End If
        End If
       
        With rsTemp
            If UCase(TypeName(txtAuditingMan)) = "TEXTBOX" Then
                txtAuditingMan = IIf(IsNull(!姓名), "", !姓名)
                zlCommFun.PressKey (vbKeyTab)
            Else
                txtAuditingMan.SetFocus
                txtAuditingMan.SelStart = 0
                txtAuditingMan.SelLength = Len(txtAuditingMan.Text)
                zlCommFun.PressKey vbKeyTab
            End If
            .Close
        End With
    End If
End Sub

Private Sub txtAuditingMan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtAuditingMan_LostFocus()
    If zlCommFun.ActualLen(Trim(txtAuditingMan.Text)) > 20 Then
        MsgBox "接收人超长,最多能输入10个汉字或20个字符!", vbInformation + vbOKOnly, gstrSysName
        txtAuditingMan.SetFocus
        txtAuditingMan.SelStart = 0
        txtAuditingMan.SelLength = Len(txtAuditingMan.Text)
        If txtAuditingMan.Enabled Then txtAuditingMan.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtBeginNo_GotFocus()
    With txtBeginNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    zlCommFun.OpenIme False
End Sub

Private Sub txtBeginNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtBeginNo.Text = Replace(UCase(txtBeginNo.Text), "'", "")
'        If Len(txtBeginNo) < 20 And Len(txtBeginNo) > 0 Then
'            strNo = txtBeginNo.Text
'            Call MakeNO(117, lng库房ID, strNo)
'            txtBeginNo.Text = strNo
'        End If
        zlCommFun.PressKey (vbKeyTab)
'        txtEndNo.SetFocus
    End If
End Sub

Private Sub txtBeginNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> 8 Then
            If KeyAscii <> 46 Then
                KeyAscii = 0
            ElseIf InStr(txtBeginNo.Text, ".") = 0 Then
                KeyAscii = 0
'            Else
'                KeyAscii = 0
            End If
        End If
    End If
    
    If KeyAscii = 39 Then KeyAscii = 0

End Sub

Private Sub txtBeginNo_LostFocus()
    If zlCommFun.ActualLen(Trim(txtBeginNo.Text)) > 18 Then
        MsgBox "开始住院号超长,最多能输入9个汉字或18个字符!", vbInformation + vbOKOnly, gstrSysName
        txtBeginNo.SetFocus
        txtBeginNo.SelStart = 0
        txtBeginNo.SelLength = Len(txtBeginNo.Text)
        If txtBeginNo.Enabled Then txtBeginNo.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtEndNo_GotFocus()
    With txtEndNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    zlCommFun.OpenIme False
End Sub

Private Sub txtEndNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtEndNo.Text = Replace(UCase(txtEndNo.Text), "'", "")
'        If Len(txtEndNo) < 20 And Len(txtEndNo) > 0 Then
'            strNo = txtEndNo.Text
'            Call MakeNO(117, lng库房ID, strNo)
'            txtEndNo.Text = strNo
'        End If
        zlCommFun.PressKey (vbKeyTab)
'        txtEndNo.SetFocus
    End If
End Sub

Private Sub txtEndNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> 8 Then
            If KeyAscii <> 46 Then
                KeyAscii = 0
            ElseIf InStr(txtBeginNo.Text, ".") = 0 Then
                KeyAscii = 0
'            Else
'                KeyAscii = 0
            End If
        End If
    End If
    
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtEndNo_LostFocus()
    If zlCommFun.ActualLen(Trim(txtEndNo.Text)) > 18 Then
        MsgBox "结束住院号超长,最多能输入9个汉字或18个字符!", vbInformation + vbOKOnly, gstrSysName
        txtEndNo.SetFocus
        txtEndNo.SelStart = 0
        txtEndNo.SelLength = Len(txtEndNo.Text)
        If txtEndNo.Enabled Then txtEndNo.SetFocus
        Exit Sub
    End If
End Sub

Private Sub vfgList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Call SaveHead(vfgList, 1)
End Sub

'Private Sub vfgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'''    If OldRow > 0 Then
''        Call zl_VsGridRowChange(vfgList, OldRow, NewRow, OldCol, NewCol)
'''    End If
'End Sub

Private Sub vfgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strStatus As String
    Dim strTemp As String
    If mlngModule = 201 Then '电子病案
        strTemp = ";增加;"
    Else
        strTemp = "基本;"
    End If
    
    If vfgList.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    
    strStatus = Trim(vfgList.TextMatrix(Row, vfgList.ColIndex("状态")))
                
    Select Case Col
        Case vfgList.ColIndex("选择")
            If InStr(mstrPrivs, strTemp) <> 0 And strStatus = "未接收" Then  'And mlngApplyId <> 0
                Cancel = False
            Else
                Cancel = True
            End If
            
            Exit Sub
        Case Else
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub vfgList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Select Case Col
        Case vfgList.ColIndex("序号"), vfgList.ColIndex("选择")
            Position = -1
            Exit Sub
    End Select
    If Position = 0 Or Position = 1 Then
        Position = Col
    End If
End Sub

Private Sub vfgList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call SaveHead(vfgList, 1)
End Sub

Private Sub vfgList_Click()
    Dim lngApplyId As String
    Dim strStatus As String
    Dim strStuffMan As String
    Dim lngRows As Long
    Dim lngFlag As Long
    Dim i As Long
    Dim strTemp As String
    
    If mlngModule = 201 Then '电子病案
        strTemp = ";增加;"
    Else
        strTemp = "基本;"
    End If
    
    lngFlag = 0
    With vfgList
        If .Rows > 1 Then
            lngRows = .Rows - 1
            For i = 1 To lngRows
                strStatus = Trim(.TextMatrix(i, .ColIndex("状态")))
                If InStr(mstrPrivs, strTemp) <> 0 And strStatus = "未接收" And .TextMatrix(i, .ColIndex("选择")) = -1 Then
                    lngFlag = 1
                    i = lngRows
                End If
            Next
        End If
        
        If lngFlag = 0 Then
            mlngTempId = 0
        End If
        If .Row > 0 Then
            lngApplyId = Val(.TextMatrix(.Row, .ColIndex("出院科室id")))
            strStatus = Trim(.TextMatrix(.Row, .ColIndex("状态")))
            If InStr(mstrPrivs, strTemp) <> 0 And strStatus = "未接收" And .TextMatrix(.Row, .ColIndex("选择")) = -1 Then
                If mlngTempId = 0 Then mlngTempId = lngApplyId
                If lngApplyId <> mlngTempId And cboOutDept.ListIndex <> 0 Then
                    MsgBox "该病案出院科室与已经选择的出院科室不一致，系统自动取消选中!", vbInformation, gstrSysName
                    .TextMatrix(.Row, .ColIndex("选择")) = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vfgList_DblClick()
    '双击编辑
    Call SetModify_DblClick
End Sub

Private Sub vfgList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If KeyCode = vbKeyReturn Then
        Call zlVsMoveGridCell(vfgList, vfgList.ColIndex("序号"), vfgList.ColIndex("家庭地址"), False, lngRow)
    End If
End Sub

Private Sub vfgList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intGetHeight As Integer
    Dim intGetWidth As Integer
    
    intGetWidth = vfgList.ColWidth(0)
    intGetHeight = vfgList.RowHeight(0)
    If (Button = 2) Then
        If x < intGetWidth And y < intGetHeight Then
            Call RightHead(vfgList, 1)
        Else
            objExtendedBar.ShowPopup
        End If
    End If
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    Dim intCol As Long
    With vfgList
        '清除选择行的颜色
''        For intCol = 0 To .Cols - 1
''            .Col = intCol
''            .CellBackColor = glngGetFocus_Font
'''            .CellForeColor = glngLostFocus_Font
''        Next
        .GridLines = flexGridInset
    End With
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = mstrListTitle
        
    Set objRow = New zlTabAppRow

    If cboOutDept.Visible Then
        objRow.Add "出院科室:" & cboOutDept.Text
    End If
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vfgList
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    With vfgList
        .GridLines = flexGridNone
    End With
End Sub

Private Sub SetAdd()
    Dim strNo As String
    Dim intEditState As Integer
    Dim strPatientSum As String
    Dim lngApplyId As Long
    Dim blnReturn As Boolean
    
    intEditState = 1
    strPatientSum = GetChoiceData()
'    If mlngTempId = 0 Then
'        lngApplyId = mlngApplyId
'    Else
'        lngApplyId = mlngTempId
'    End If
    lngApplyId = cboOutDept.ItemData(cboOutDept.ListIndex)
    
    frm病案接收编辑.ShowCard Me, intEditState, strPatientSum, lngApplyId, blnReturn, mlngModule
    If blnReturn Then
        mintDelete = 1
        mintEditState = intEditState
        Call SetRefresh
        mintDelete = 0
    End If
End Sub

Private Sub SetModify()
    Dim intEditState As Integer
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strStatus As String
    Dim blnReturn As Boolean
    Dim strPatientSum As String
    Dim lngApplyId As Long
    Dim lngRow As Long
    
    intEditState = 2
    
    With vfgList
        If .Row > 0 Then
            lngRow = .Row
            lngPatientlId = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            lngMtyId = Val(.TextMatrix(.Row, .ColIndex("主页ID")))
            lngApplyId = Val(.TextMatrix(.Row, .ColIndex("出院科室id")))
            strStatus = Trim(.TextMatrix(.Row, .ColIndex("状态")))
            
            If strStatus = "已接收" Then
                strPatientSum = lngPatientlId & "_" & lngMtyId
                frm病案接收编辑.ShowCard Me, intEditState, strPatientSum, lngApplyId, blnReturn
            End If
        End If
    End With
    If blnReturn Then
        mintDelete = 1
        mintEditState = intEditState
        Call SetRefresh
        mintDelete = 0
        If lngRow > 0 And lngRow < vfgList.Rows Then vfgList.Select lngRow, 2
    End If
End Sub

Private Sub SetDelete()
    Dim strSQL As String
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strStatus As String
    Dim strPname As String
    Dim lngRow As Long
    
    strSQL = ""
    
    With vfgList
        If .Row > 0 Then
            lngRow = .Row
            lngPatientlId = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            lngMtyId = Val(.TextMatrix(.Row, .ColIndex("主页ID")))
            strPname = Trim(.TextMatrix(.Row, .ColIndex("姓名")))
            strStatus = Trim(.TextMatrix(.Row, .ColIndex("状态")))
            
            If strStatus = "已接收" Then
                strSQL = "Zl_病案接收记录_Delete(" & lngPatientlId & "," & lngMtyId & ")"
            End If
        End If
    End With
    If Trim(strSQL) <> "" Then
        If MsgBox("你确定删除病人为【" & strPname & "】的病案接收登记吗，删除后不能恢复？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            Err = 0: On Error GoTo errHand:
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            mintDelete = 1
            Call SetRefresh
            If lngRow + 1 > vfgList.Rows Then
                lngRow = lngRow - 1
            End If
            If lngRow > 0 And lngRow < vfgList.Rows Then vfgList.Select lngRow, 2
            mintDelete = 0
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDisplay()
    Dim intEditState As Integer
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strStatus As String
    Dim blnReturn As Boolean
    Dim strPatientSum As String
    Dim lngApplyId As Long
    
    intEditState = 3
    
    With vfgList
        If .Row > 0 Then
            lngPatientlId = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
            lngMtyId = Val(.TextMatrix(.Row, .ColIndex("主页ID")))
            lngApplyId = Val(.TextMatrix(.Row, .ColIndex("出院科室id")))
            strStatus = Trim(.TextMatrix(.Row, .ColIndex("状态")))
            
            If strStatus <> "未接收" Then
                strPatientSum = lngPatientlId & "_" & lngMtyId
                frm病案接收编辑.ShowCard Me, intEditState, strPatientSum, lngApplyId
            End If
        End If
    End With
    mintEditState = intEditState
End Sub

Private Sub SetModify_DblClick()
    If mintDblClick = 1 Then
        Call SetModify
'    Else
'        Call SetDisplay
    End If
End Sub

Private Sub SetVerify_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strStatus As String
   
    Control.Visible = True
    Control.Enabled = False
  
    If vfgList.Row > 0 Then
        strStatus = Trim(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("状态")))
        If strStatus = "已接收" Then
            Control.Enabled = True
        End If
    End If
End Sub

Private Function GetChoiceData() As String
    Dim lngApplyId As String
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strStatus As String
    Dim lngRows As Long
    Dim strTemp As String
    Dim i As Long
    Dim j As Long
    Dim intCount As Integer
    Dim strTempEle As String
    If mlngModule = 201 Then '电子病案
        strTempEle = ";增加;"
    Else
        strTempEle = "基本;"
    End If
    
    intCount = 0
    strTemp = ""
    GetChoiceData = ""
    With vfgList
        If .Rows > 1 Then
            lngRows = .Rows - 1
            For i = 1 To lngRows
                lngPatientlId = Val(.TextMatrix(i, .ColIndex("病人ID")))
                lngMtyId = Val(.TextMatrix(i, .ColIndex("主页ID")))
                strStatus = Trim(.TextMatrix(i, .ColIndex("状态")))
                If InStr(mstrPrivs, strTempEle) <> 0 And strStatus = "未接收" And .TextMatrix(i, .ColIndex("选择")) = -1 Then
                    intCount = intCount + 1
                    If intCount > 100 Then
                        GetChoiceData = strTemp
                        MsgBox "你所选的病案数太多了，只处理前面选中的100份。", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If strTemp = "" Then
                        strTemp = lngPatientlId & "_" & lngMtyId
                    Else
                        strTemp = strTemp & "," & lngPatientlId & "_" & lngMtyId
                    End If
                End If
            Next
        End If
    End With
    GetChoiceData = strTemp
End Function

'==============================================================================
'=功能： 查看首页
'==============================================================================
Private Sub RecordLook()
    
    On Error GoTo ErrH
    With vfgList
        If .Row < 1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("病人id"))) = 0 Then GoTo ErrH
        Call frmArchiveView.ShowArchive(Me, Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))), False)
        
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetMedicalExits() As ADODB.Recordset
    '******************************************************************************************************************
    '功能:检查是否安装了病案系统
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    Dim strSQL As String
    strSQL = "Select 编号 From zlsystems where 编号=300"
    
    Set GetMedicalExits = zlDatabase.OpenSQLRecord(strSQL, "病案系统")
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

