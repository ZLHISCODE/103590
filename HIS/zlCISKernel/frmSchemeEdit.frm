VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSchemeEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "成套方案"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   Icon            =   "frmSchemeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4065
      Left            =   60
      TabIndex        =   0
      Top             =   555
      Width           =   10770
      _cx             =   18997
      _cy             =   7170
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   18
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSchemeEdit.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Frame fraAdvice 
      Height          =   2070
      Left            =   45
      TabIndex        =   19
      Top             =   4680
      Width           =   10800
      Begin VB.CommandButton cmd适用证候 
         Caption         =   "…"
         Height          =   285
         Left            =   10285
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   203
         Width           =   285
      End
      Begin VB.CheckBox chkMedicineVariety 
         Caption         =   "按品种输入医嘱"
         Height          =   300
         Left            =   3360
         TabIndex        =   2
         Top             =   195
         Width           =   1575
      End
      Begin VB.TextBox txt天数 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2380
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1635
         Visible         =   0   'False
         Width           =   360
      End
      Begin MSComctlLib.Toolbar tbrFree 
         Height          =   330
         Left            =   300
         TabIndex        =   33
         Top             =   810
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "自由录入医嘱(F3)"
               ImageIndex      =   1
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cbo附加执行 
         Height          =   300
         Left            =   6255
         TabIndex        =   18
         Text            =   "cbo附加执行"
         Top             =   1635
         Width           =   1725
      End
      Begin VB.CommandButton cmd频率 
         Height          =   240
         Left            =   4860
         Picture         =   "frmSchemeEdit.frx":065F
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(F4)"
         Top             =   1305
         Width           =   270
      End
      Begin VB.TextBox txt频率 
         Height          =   300
         Left            =   3495
         TabIndex        =   8
         Top             =   1275
         Width           =   1665
      End
      Begin VB.TextBox txt单量 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3495
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1635
         Width           =   1380
      End
      Begin VB.TextBox txt总量 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   930
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1635
         Width           =   1530
      End
      Begin VB.CommandButton cmd用法 
         Height          =   240
         Left            =   2445
         Picture         =   "frmSchemeEdit.frx":0755
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(F4)"
         Top             =   1305
         Width           =   270
      End
      Begin VB.TextBox txt用法 
         Height          =   300
         Left            =   930
         TabIndex        =   6
         Top             =   1275
         Width           =   1815
      End
      Begin VB.ComboBox cbo期效 
         Height          =   300
         ItemData        =   "frmSchemeEdit.frx":084B
         Left            =   930
         List            =   "frmSchemeEdit.frx":0855
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   195
         Width           =   2160
      End
      Begin VB.CommandButton cmdExt 
         Height          =   285
         Left            =   4890
         Picture         =   "frmSchemeEdit.frx":0869
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "编辑(F4)"
         Top             =   600
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   285
         Left            =   4890
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   900
         Width           =   285
      End
      Begin VB.ComboBox cbo执行科室 
         Height          =   300
         Left            =   6255
         TabIndex        =   16
         Text            =   "cbo执行科室"
         Top             =   1275
         Width           =   1725
      End
      Begin VB.TextBox txt医嘱内容 
         Height          =   675
         Left            =   930
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "按 ~ 键切换快捷浮动面板"
         Top             =   555
         Width           =   3945
      End
      Begin VB.ComboBox cbo执行时间 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6255
         TabIndex        =   15
         Top             =   915
         Width           =   4350
      End
      Begin VB.ComboBox cbo执行性质 
         Height          =   300
         ItemData        =   "frmSchemeEdit.frx":095F
         Left            =   8805
         List            =   "frmSchemeEdit.frx":096C
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1275
         Width           =   1800
      End
      Begin VB.ComboBox cbo医生嘱托 
         Height          =   300
         Left            =   6255
         TabIndex        =   14
         Top             =   555
         Width           =   4350
      End
      Begin VB.ComboBox cbo滴速 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6255
         TabIndex        =   37
         Top             =   195
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox txt适用证候 
         Height          =   300
         Left            =   6255
         MaxLength       =   100
         TabIndex        =   13
         Top             =   195
         Width           =   4335
      End
      Begin VB.Label lbl滴速单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "滴/分钟"
         Height          =   180
         Left            =   7320
         TabIndex        =   38
         Top             =   255
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lbl适用证候 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "适用证候"
         Height          =   180
         Left            =   5490
         TabIndex        =   36
         Top             =   255
         Width           =   720
      End
      Begin VB.Label lbl天数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用    天"
         Height          =   180
         Left            =   2190
         TabIndex        =   34
         Top             =   1695
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl附加执行 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "附加执行"
         Height          =   180
         Left            =   5490
         TabIndex        =   32
         Top             =   1695
         Width           =   720
      End
      Begin VB.Label lbl频率 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "频率"
         Height          =   180
         Left            =   3105
         TabIndex        =   26
         Top             =   1335
         Width           =   360
      End
      Begin VB.Label lbl单量单位 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "单位"
         Height          =   180
         Left            =   4905
         TabIndex        =   22
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label lbl单量 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单量"
         Height          =   180
         Left            =   3105
         TabIndex        =   21
         Top             =   1695
         Width           =   360
      End
      Begin VB.Label lbl总量单位 
         BackStyle       =   0  'Transparent
         Caption         =   "单位"
         Height          =   180
         Left            =   2490
         TabIndex        =   24
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label lbl总量 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总量"
         Height          =   180
         Left            =   525
         TabIndex        =   23
         Top             =   1695
         Width           =   360
      End
      Begin VB.Label lbl期效 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医嘱期效"
         Height          =   180
         Left            =   165
         TabIndex        =   31
         Top             =   255
         Width           =   720
      End
      Begin VB.Label lbl医生嘱托 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医生嘱托"
         Height          =   180
         Left            =   5490
         TabIndex        =   30
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lbl执行科室 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室"
         Height          =   180
         Left            =   5490
         TabIndex        =   28
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lbl用法 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用法"
         Height          =   180
         Left            =   525
         TabIndex        =   25
         Top             =   1335
         Width           =   360
      End
      Begin VB.Label lbl医嘱内容 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医嘱内容"
         Height          =   180
         Left            =   165
         TabIndex        =   20
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl执行时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行时间"
         Height          =   180
         Left            =   5490
         TabIndex        =   27
         Top             =   975
         Width           =   720
      End
      Begin VB.Label lbl执行性质 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行性质"
         Height          =   180
         Left            =   8055
         TabIndex        =   29
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lbl滴速 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "↓滴速"
         Height          =   180
         Left            =   5640
         TabIndex        =   39
         Top             =   255
         Visible         =   0   'False
         Width           =   570
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   435
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSchemeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOK As Boolean
        
'入口参数
Private mint范围 As Integer '1-门诊使用,2-住院使用,3-门诊和住院都可以使用
Private mrsScheme As ADODB.Recordset '与表"诊疗项目组合"相同结构的动态记录集
Private mbln显示缺省列 As Boolean        '临床路径项目定义调用"选择使用"时显示缺省列
Private mbyt场合 As Byte            'byt场合=1-临床路径项目定义调用，0-成套方案调用,2-临床路径项目批量调整

Private mstr诊疗分类 As String
Private mstr操作类型 As String
Private mstr执行分类 As String

'程序变量
Private mobjVBA As Object
Private mobjScript As clsScript
Private mrsDefine As ADODB.Recordset
Private mlngNextID As Long
Private mblnView As Boolean
Private msng天数 As Single
Private mstr使用科室 As String

'本地参数
Private mint简码 As Integer
Private mstrLike As String
Private mbln一次性 As Boolean '临嘱缺省为一次性
Private mblnNewLIS As Boolean

'事件状态控制变量
Private mblnNoSave As Boolean
Private mblnRowMerge As Boolean
Private mblnRunFirst As Boolean
Private mblnRowChange As Boolean

'工具栏命令
Private Const conMenu_New = 100
Private Const conMenu_Insert = 101
Private Const conMenu_Delete = 102
Private Const conMenu_Merge = 104
Private Const conMenu_Import = 105
Private Const conMenu_Save = 107
Private Const conMenu_Exit = 111
Private Const conMenu_MoveDown = 203
Private Const conMenu_MoveUp = 204

'执行时间示例
Private Const COL_按周执行 = _
    "每周三次 1/8-3/8-5/8 或 1/8:00-3/8:00-5/8:00" & vbCrLf & _
        vbTab & "表示在每周星期一的8:00,星期三的8:00,星期五的8:00这几个时间执行"
Private Const COL_按天执行 = _
    "每天三次 8-12-16 或 8:00-12:00-16:00" & vbCrLf & _
        vbTab & "表示在每天8:00,12:00,16:00这几个时间执行" & vbCrLf & _
    "两天一次 1/8 或 1/8:00" & vbCrLf & _
        vbTab & "表示在每两天中的第1天8:00这个时间执行"
Private Const COL_按时执行 = _
    "每小时两次 1:20-1:40" & vbCrLf & _
        vbTab & "表示在每小时内的20和40分钟这两个时间执行" & vbCrLf & _
    "两小时一次 2:30 或 1:30 或 1:00" & vbCrLf & _
        vbTab & "表示在每两小时内的第2的个小时的30分钟这个时间执行" & vbCrLf & _
        vbTab & "　或在每两小时内的第1的个小时的30分钟这个时间执行" & vbCrLf & _
        vbTab & "　或在每两小时内的第1的个小时这个时间执行"

Private Enum mvCol
    '可见列索引
    col_备选 = 0
    col_缺省 = 1
    COL_期效 = 2
    col_医嘱内容 = 3
    COL_总量 = 4
    COL_总量单位 = 5
    COL_单量 = 6
    COL_单量单位 = 7
    COL_天数 = 8
    COL_频率 = 9
    COL_用法 = 10
    COL_医生嘱托 = 11
    COL_执行时间 = 12
    
    '隐藏列索引
    COL_相关ID = 13
    COL_序号 = 14
    COL_类别 = 15
    COL_诊疗项目ID = 16
    COL_名称 = 17
    COL_标本部位 = 18
    COL_检查方法 = 19
        COL_中药形态 = 19 '0=散装，1=中药饮片，2=免煎剂
    COL_收费细目ID = 20
    COL_频率次数 = 21
    COL_频率间隔 = 22
    COL_间隔单位 = 23
    COL_执行科室ID = 24
    COL_执行性质 = 25 '病人医嘱记录.执行性质=诊疗项目目录.执行科室
    COL_执行标记 = 26
    
    COL_计算方式 = 27 '诊疗项目目录.计算方式
    COL_频率性质 = 28 '诊疗项目目录.执行频率
    COL_操作类型 = 29 '诊疗项目目录.操作类型
    COL_可否分零 = 30 '卫材用于存放是否跟踪在用
        COL_跟踪在用 = 30
    COL_剂量系数 = 31
    COL_包装单位 = 32
    COL_包装系数 = 33
    COL_毒理分类 = 34
    COL_药品剂型 = 35
    COL_配方ID = 36
    COL_临床自管药 = 37
    COL_组合项目ID = 38
    COL_适用证候 = 39
    COL_抗菌等级 = 40 '抗菌药物等级:0-非抗菌药,1-非限制级,2-限制级,3-特殊使用级
    COL_是否停用 = 41 '=1标识已停用，=0或NULL标识未停用
    COL_执行分类 = 42 '0-其他治疗类别,1-输液类,2-注射类,3-皮试,4-口服
End Enum

Public Function ShowMe(frmParent As Object, ByVal int范围 As Long, Optional rsScheme As ADODB.Recordset, _
    Optional ByVal blnView As Boolean, Optional ByVal bln显示缺省列 As Boolean, Optional ByVal str使用科室 As String, Optional ByVal byt场合 As Byte, _
    Optional ByVal str诊疗分类 As String, Optional ByVal str操作类型 As String, Optional ByVal str执行分类 As String) As ADODB.Recordset
'返回：与表"诊疗项目组合"相同结构的动态记录集,如果取消则返回Nothing
'参数：byt场合=1-临床路径项目定义调用，0-成套方案调用,2-临床路径项目批量调整
'   str诊疗分类:byt场合=2时传人
'   str操作类型:byt场合=2时传人
'   str执行分类:byt场合=2时传人

    mint范围 = int范围
    mbln显示缺省列 = bln显示缺省列
    Set mrsScheme = rsScheme
    mblnView = blnView
    mstr使用科室 = str使用科室
    mbyt场合 = byt场合
    
    mstr诊疗分类 = str诊疗分类
    mstr操作类型 = str操作类型
    mstr执行分类 = str执行分类
   
    On Error Resume Next
    Me.Show 1, frmParent
    
    If mblnOK Then
        Set ShowMe = mrsScheme
    End If
    Set mrsScheme = Nothing
    
End Function

Private Sub InitCommandBar()
'功能：初始化工具栏
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
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = frmIcons.imgMain.Icons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_New, "增加"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Insert, "插入")
        Set objControl = .Add(xtpControlButton, conMenu_Delete, "删除")
        Set objControl = .Add(xtpControlButton, conMenu_Merge, "一并给药"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_MoveUp, "上移")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_MoveDown, "下移")
        Set objControl = .Add(xtpControlButton, conMenu_Import, "导入")
        objControl.IconId = conMenu_Insert
        objControl.BeginGroup = True
        objControl.ToolTipText = "从病人医嘱导入"
        Set objControl = .Add(xtpControlButton, conMenu_Save, "保存")
        objControl.ToolTipText = "确认保存并退出"
        Set objControl = .Add(xtpControlButton, conMenu_Exit, "退出"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_New
        .Add FCONTROL, vbKeyI, conMenu_Insert
        .Add FCONTROL, vbKeyK, conMenu_Merge
        .Add FCONTROL, vbKeyT, conMenu_Import
        .Add FCONTROL, vbKeyS, conMenu_Save
        .Add FALT, vbKeyX, conMenu_Exit
    End With
End Sub

Private Sub InitAdviceTable()
'功能：初始化表格内容，用在窗体个性化设置恢复之前
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant
    
    strHead = _
        "备选,450,4;缺省,450,4;期效,500,4;医嘱内容,3500,1;总量,600,7;单位,450,1;单量,600,7;单位,450,1;天数,450,1;频率,1200,1;用法,1200,1;" & _
        "医生嘱托,1000,1;执行时间;相关ID;序号;类别;诊疗项目ID;名称;标本部位;检查方法;收费细目ID;频率次数;频率间隔;间隔单位;执行科室ID;" & _
        "执行性质;执行标记;计算方式;频率性质;操作类型;可否分零;剂量系数;包装单位;包装系数;毒理分类;药品剂型;配方ID;临床自管药;组合项目ID;" & _
        "适用证候,1000,1;抗菌等级;是否停用;执行分类"
        
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        If mbln显示缺省列 = False Then
            .ColHidden(col_缺省) = True
            .ColHidden(col_备选) = True
        Else
            .ColDataType(col_缺省) = flexDTBoolean
            .ColDataType(col_备选) = flexDTBoolean
            .ColHidden(col_缺省) = False
            .ColHidden(col_备选) = False
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub cbo滴速_Change()
     cbo滴速.Tag = "1"
End Sub

Private Sub cbo滴速_Click()
    cbo滴速.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cbo滴速_GotFocus()
    zlControl.TxtSelAll cbo滴速
End Sub

Private Sub cbo滴速_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cbo滴速_Validate(False)
    ElseIf InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbo滴速_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cbo滴速.Text) > 10 Then
        MsgBox "滴速输入内容过长，请检查输入是否正确。", vbInformation, gstrSysName
        Call cbo滴速_GotFocus: Cancel = True: Exit Sub
    End If
    '更新数据
    Call AdviceChange
End Sub

Private Sub cbo附加执行_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSql As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo附加执行.ListIndex = -1 Then Exit Sub
    
    If cbo附加执行.ItemData(cbo附加执行.ListIndex) = -1 Then
        strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
            " From 部门表 A,部门性质说明 B" & _
            " Where A.ID=B.部门ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([1],3)") & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " Order by A.编码"
        vRect = zlControl.GetControlRect(cbo附加执行.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lbl附加执行.Caption, False, "", "", False, False, True, vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, mint范围)
        If Not rsTmp Is Nothing Then
            intIdx = Cbo.FindIndex(cbo附加执行, rsTmp!ID)
            If intIdx <> -1 Then
                cbo附加执行.ListIndex = intIdx
            Else
                cbo附加执行.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo附加执行.ListCount - 1
                cbo附加执行.ItemData(cbo附加执行.NewIndex) = rsTmp!ID
                cbo附加执行.ListIndex = cbo附加执行.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
            '恢复成现有的科室(不引发Click)
            intIdx = Cbo.FindIndex(cbo附加执行, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_执行科室ID)))
            Call Cbo.SetIndex(cbo附加执行.Hwnd, intIdx)
        End If
    Else
        cbo附加执行.Tag = "1"
        lngRow = vsAdvice.Row
        
        '更新更改了的执行科室医嘱内容
       Call AdviceChange
    End If
End Sub

Private Sub cbo附加执行_GotFocus()
    Call zlControl.TxtSelAll(cbo附加执行)
End Sub

Private Sub cbo附加执行_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo附加执行.ListIndex = -1 Then
            Call cbo附加执行_Validate(blnCancel)
        End If
        If Not blnCancel Then
            If SeekNextControl Then Call cbo附加执行_Validate(False)
        End If
    End If
End Sub

Private Sub cbo附加执行_Validate(Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If cbo附加执行.ListIndex <> -1 Then Exit Sub '已选中
    If cbo附加执行.Text = "" Then '不输入
        cbo附加执行.Tag = "1"
        Call AdviceChange
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '是否可以任意或选择科室
    blnLimit = True
    If cbo附加执行.ListCount > 0 Then
        If cbo附加执行.ItemData(cbo附加执行.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    strInput = UCase(zlCommFun.GetNeedName(cbo附加执行.Text))
    strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([3],3)") & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And (A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2])" & _
        " Order by A.编码"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, strInput & "%", mstrLike & strInput & "%", mint范围)
        For i = 1 To rsTmp.RecordCount
            intIdx = Cbo.FindIndex(cbo附加执行, rsTmp!ID)
            If intIdx <> -1 Then cbo附加执行.ListIndex = intIdx: Exit For
            rsTmp.MoveNext
        Next
        If cbo附加执行.ListIndex = -1 Then
            MsgBox "未到对应的科室。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = zlControl.GetControlRect(cbo附加执行.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lbl附加执行.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = Cbo.FindIndex(cbo附加执行, rsTmp!ID)
            If intIdx <> -1 Then
                cbo附加执行.ListIndex = intIdx
            Else
                cbo附加执行.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo附加执行.ListCount - 1
                cbo附加执行.ItemData(cbo附加执行.NewIndex) = rsTmp!ID
                cbo附加执行.ListIndex = cbo附加执行.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "未到对应的科室。", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo期效_Click()
'功能：更改项目期效时,清空当前行的数据
    Dim lngRow As Long, i As Long
    
    With vsAdvice
        lngRow = .Row
        If .RowData(lngRow) = 0 Then Exit Sub
        
        If zlCommFun.GetNeedName(cbo期效.Text) = .TextMatrix(lngRow, COL_期效) Then Exit Sub
        
        '自由录入医嘱直接更改期效
        If Val(.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
            .TextMatrix(lngRow, COL_期效) = zlCommFun.GetNeedName(cbo期效.Text)
            mblnNoSave = True: Exit Sub
        End If
        
        If CanAlterType(lngRow) Then
            Call AdviceAlterType(lngRow)
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, col_医嘱内容)
        Else
            '一并给药中某一个不准改(因为规格原因),则当前行内容不能清除
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                If RowIn一并给药(lngRow) Then
                    MsgBox "一并给药的药品中存在未按规格下达的药品，不能更改为临嘱。", vbInformation, gstrSysName
                    Call Cbo.SetIndex(cbo期效.Hwnd, IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1))
                    Exit Sub
                End If
            End If
        
            If MsgBox("更改医嘱期效后需要重新输入医嘱内容,要更改吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Call Cbo.SetIndex(cbo期效.Hwnd, IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1))
                Exit Sub
            End If
            
            '清除医嘱数据行
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '西成药、中成药:只可能是单独给药的,删除给药途径行,并清除当前行
                i = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                Call DeleteRow(i)
                Call DeleteRow(lngRow, True)
            ElseIf InStr(",D,F,K,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '检查组合项目、手术项目、输血医嘱
                '删除部位行、手术附加行(附加手术,麻醉项目)、输血途径
                Call Delete检查手术输血(lngRow)
                '清除当前行
                Call DeleteRow(lngRow, True)
            ElseIf RowIn配方行(lngRow) Then
                '中药配方：顺序(序号)要求必须严格控制
                '删除组成味药及煎法行:删除之后重新定位的当前行
                lngRow = Delete中药配方(lngRow)
                '清除当前行(中药用法行)
                Call DeleteRow(lngRow, True)
            Else
                '其它项目直接清除当前行内容
                Call DeleteRow(lngRow, True)
            End If
            
            '重新进入行
            i = cbo期效.ListIndex '保留当前选择的期效
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, col_医嘱内容)
            cbo期效.ListIndex = i '就是需要再激活以设置开始时间值
        End If
    End With
End Sub

Private Sub cbo期效_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo期效.ListIndex <> -1 Then
            Call SeekNextControl
        End If
    ElseIf KeyAscii >= 32 Then
        lngIdx = Cbo.MatchIndex(cbo期效.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo期效.ListCount > 0 Then lngIdx = 0
        cbo期效.ListIndex = lngIdx
    End If
End Sub

Private Sub Set用法Input(rsInput As ADODB.Recordset, ByVal int类型 As Integer)
'功能：输入给药途径或中药用法后调用
'参数：rsInput=输入或选择的返回记录
'      int类型=2-给药途径,4-中药用法
'说明：如果可选频率,则配合给药途径处理可用执行时间方案的变化
    Dim rsTmp As New ADODB.Recordset
    Dim blnValid As Boolean, strSql As String, i As Long
    Dim str频率 As String, int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim vMsg As VbMsgBoxResult, strMsg As String
    
    On Error GoTo errH
    cmd用法.Tag = rsInput!ID
    txt用法.Text = rsInput!名称
    txt用法.Tag = "1"
    
    With vsAdvice
       
        If int类型 = 2 Then
            If NVL(rsInput!执行分类ID, 0) <> 1 And cbo滴速.Text <> "" Then
                '非输液类清除滴速
                cbo滴速.Text = ""
                cbo滴速.Tag = "1"
            End If
        End If
        '重新获取可用的缺省时间方案
        If cbo执行时间.Enabled Then '"可选频率"或药品时
            Call Get时间方案(cbo执行时间, Get频率范围(.Row), .TextMatrix(.Row, COL_频率), rsInput!ID)
            If cbo执行时间.ListCount > 0 Then
                Call Cbo.SetIndex(cbo执行时间.Hwnd, 0)
                cbo执行时间.Tag = "1"
            Else
                '判断当前执行时间是否合法
                If cbo执行时间.Text <> "" Then
                    blnValid = ExeTimeValid(cbo执行时间.Text, Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), .TextMatrix(.Row, COL_间隔单位))
                    If Not blnValid Then '如果不合法,则另取,否则保持
                        cbo执行时间.Text = ""
                        cbo执行时间.Tag = "1"
                    End If
                End If
            End If
        End If
        
        '根据诊疗用法用量作缺省设置
        If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
            If Val(.TextMatrix(.Row, COL_收费细目ID)) <> 0 Then
                strSql = "Select 频次,小儿剂量,成人剂量,医生嘱托,疗程" & _
                    " From 药品用法用量 Where  药品ID=[1] And 用法ID=[2] And 性质=1"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(.Row, COL_收费细目ID)), Val(rsInput!ID))
            Else
                strSql = "Select 频次,小儿剂量,成人剂量,医生嘱托,疗程 From 诊疗用法用量 Where 性质>0 And 项目ID=[1] And 用法ID=[2]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(.Row, COL_诊疗项目ID)), Val(rsInput!ID))
            End If
            
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!频次) And Val(.TextMatrix(.Row, COL_频率性质)) <> 1 Then '已为一次性时不管
                    Call Get频率信息_编码(rsTmp!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                    txt频率.Text = str频率
                    cmd频率.Tag = str频率
                    txt频率.Tag = "1"
                End If
                
                '根据新的频率重新设置执行时间
                If cbo执行时间.Enabled Then
                    Call Get时间方案(cbo执行时间, Get频率范围(.Row), str频率, rsInput!ID)
                    If cbo执行时间.ListCount > 0 Then
                        Call Cbo.SetIndex(cbo执行时间.Hwnd, 0)
                        cbo执行时间.Tag = "1"
                    Else
                        '判断当前执行时间是否合法
                        If cbo执行时间.Text <> "" Then
                            blnValid = ExeTimeValid(cbo执行时间.Text, int频率次数, int频率间隔, str间隔单位)
                            If Not blnValid Then '如果不合法,则另取,否则保持
                                cbo执行时间.Text = ""
                                cbo执行时间.Tag = "1"
                            End If
                        End If
                    End If
                End If

                '药品单量
                If NVL(rsTmp!成人剂量, 0) <> 0 Then
                    txt单量.Text = FormatEx(rsTmp!成人剂量, 5)
                    txt单量.Tag = "1"
                End If
                
                '医生嘱托
                If Not IsNull(rsTmp!医生嘱托) Then
                    cbo医生嘱托.Text = rsTmp!医生嘱托
                    cbo医生嘱托.Tag = "1"
                End If
            End If
        End If
    End With
    
    '处理当前医嘱给药途径/煎法的变化
    Call AdviceChange
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set频率Input(rsInput As ADODB.Recordset, ByVal int范围 As Integer, ByVal int项目频率 As Integer)
'功能：输入执行频率后调用
'参数：rsInput=输入或选择的返回记录
'      int范围=1-西医;2-中医;-1-一次性;-2-持续性
'      int项目频率=项目本身的执行频率属性
'说明：配合用法处理可用执行时间方案的变化
    Dim lng用法ID As Long, blnValid As Boolean
    Dim str原执行时间 As String, i As Long
    Dim sng天数 As Single
    
    str原执行时间 = cbo执行时间.Text
    With vsAdvice
        '备用医嘱的执行频率和已执行一致。
        .TextMatrix(.Row, COL_频率性质) = decode(int范围, 1, 0, 2, 0, -1, 1, -2, 2, -3, 1, -5, 1)
        If RowIn检验行(.Row) Or int范围 = -3 Or int范围 = -5 Then   '同步赋值,因为后续以检验项目的执行性质作判断
            For i = .Row - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(.Row) Then
                    .TextMatrix(i, COL_频率性质) = .TextMatrix(.Row, COL_频率性质)
                Else
                    Exit For
                End If
            Next
        End If
        cmd频率.Tag = rsInput!名称
        txt频率.Text = rsInput!名称
        txt频率.Tag = "1"
        
        '先设置临嘱药品天数的可用性
        If cbo期效.ListIndex = 1 And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
            If Val(.TextMatrix(.Row, COL_频率性质)) = 1 Then
                If txt天数.Enabled Then SetDayState -1, -1
            Else
                If Not txt天数.Enabled Then SetDayState 1, 1
            End If
        End If
        
        '先设置总量的可用性:临嘱"计次"可选频率的设置为一次性后不输入总量(除药品外)
        If cbo期效.ListIndex = 1 And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) = 0 And Not RowIn配方行(.Row) Then
            If Val(.TextMatrix(.Row, COL_计算方式)) = 3 And int项目频率 = 0 Then
                If txt总量.Enabled And Val(.TextMatrix(.Row, COL_频率性质)) = 1 Then
                    SetItemEditable , -1
                    txt总量.Text = "1"
                ElseIf Not txt总量.Enabled And Val(.TextMatrix(.Row, COL_频率性质)) = 0 Then
                    SetItemEditable , 1
                End If
                lbl总量单位.Caption = .TextMatrix(.Row, COL_总量单位)
            End If
        End If
        
        '先设置执行时间的可用性(临嘱可选频率项目可能在一次性之间切换,及分钟频率切换)
        If int项目频率 = 0 And decode(int范围, 1, 0, 2, 0, -1, 1, -2, 2, -3, 1, -5, 1) <> 1 Then
            If Not cbo执行时间.Enabled Then SetItemEditable , , , , 1
        Else
            If cbo执行时间.Enabled Then SetItemEditable , , , , -1
        End If
        If cbo执行时间.Enabled Then '"可选频率"或药品时
            If rsInput!间隔单位 & "" = "分钟" Then
                cbo执行时间.Text = ""
            Else
                '处理可用执行时间方案的变化
                If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                    '查找给药途径对应的行
                    lng用法ID = .FindRow(CLng(.TextMatrix(.Row, COL_相关ID)), .Row + 1)
                    If lng用法ID <> -1 Then '未找到给药途径的情况,应该不可能
                        lng用法ID = .TextMatrix(lng用法ID, COL_诊疗项目ID)
                    Else
                        lng用法ID = 0
                    End If
                ElseIf RowIn配方行(.Row) Then
                    '得到对应的中药用法ID
                    lng用法ID = Val(.TextMatrix(.Row, COL_诊疗项目ID))
                End If
                
                Call Get时间方案(cbo执行时间, int范围, txt频率.Text, lng用法ID)
                '取新的频率的默认执行时间
                If cbo执行时间.ListCount > 0 Then
                    Call Cbo.SetIndex(cbo执行时间.Hwnd, 0)
                    cbo执行时间.Tag = "1"
                Else
                    '判断当前执行时间是否合法
                    If cbo执行时间.Text <> "" Then
                        blnValid = ExeTimeValid(cbo执行时间.Text, Val(rsInput!频率次数 & ""), Val(rsInput!频率间隔 & ""), rsInput!间隔单位 & "")
                        If Not blnValid Then '如果不合法,则另取,否则保持
                            cbo执行时间.Text = ""
                            cbo执行时间.Tag = "1"
                        End If
                    End If
                End If
            End If
            
            '重新计算总量
            If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 _
                And .TextMatrix(.Row, COL_期效) = "临嘱" And Val(.TextMatrix(.Row, COL_频率性质)) <> 1 Then
                sng天数 = Val(txt天数.Text)
                If sng天数 = 0 Then sng天数 = 1
                
                If txt频率.Text <> "" And Val(txt单量.Text) <> 0 _
                    And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 _
                    And Val(.TextMatrix(.Row, COL_包装系数)) <> 0 Then
                    
                    txt总量.Text = FormatEx(Calc缺省药品总量( _
                        Val(txt单量.Text), sng天数, rsInput!频率次数, _
                        rsInput!频率间隔, rsInput!间隔单位 & "", cbo执行时间.Text, _
                        Val(.TextMatrix(.Row, COL_剂量系数)), _
                        Val(.TextMatrix(.Row, COL_包装系数)), _
                        Val(.TextMatrix(.Row, COL_可否分零))), 5)
                    txt总量.Tag = "1"
                End If
            End If
        End If
        If rsInput!间隔单位 & "" = "分钟" Then
            If cbo执行时间.Enabled Then SetItemEditable , , , , -1
        End If
    End With
    
    '检查是否变化
    If cbo执行时间.Text <> str原执行时间 Then cbo执行时间.Tag = "1"
    
    '处理当前医嘱执行频率的变化
    Call AdviceChange
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    vsAdvice.Left = lngLeft
    vsAdvice.Top = lngTop
    vsAdvice.Height = lngBottom - lngTop - (fraAdvice.Height - 80)
    vsAdvice.Width = lngRight - lngLeft
    
    fraAdvice.Left = lngLeft
    fraAdvice.Top = vsAdvice.Top + vsAdvice.Height - 80
    fraAdvice.Width = lngRight - lngLeft
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    If vsAdvice.Redraw = flexRDNone Then Exit Sub
    
    If mbyt场合 = 2 Then
        If Control.ID = conMenu_Delete Or Control.ID = conMenu_Save Or Control.ID = conMenu_Exit Then
            Control.Visible = True
            If Control.ID = conMenu_Save Then Control.Enabled = mblnNoSave
        Else
            Control.Visible = False
        End If
        Exit Sub
    End If
    
    Select Case Control.ID
        Case conMenu_New
            If mblnView Then Control.Visible = False
        Case conMenu_Insert
            If mblnView Then
                Control.Visible = False
            Else
                blnEnabled = True
                If Not fraAdvice.Enabled Then
                    If InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_类别)) > 0 _
                        And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID)) = Val(vsAdvice.TextMatrix(vsAdvice.Row - 1, COL_相关ID)) Then
                        blnEnabled = False
                    End If
                End If
                Control.Enabled = blnEnabled
            End If
        Case conMenu_Delete
            If mblnView Then
                Control.Visible = False
            Else
                With vsAdvice
                    blnEnabled = True
                    If .RowData(.Row) <> 0 Then
                        If Not fraAdvice.Enabled Then blnEnabled = False
                    End If
                    Control.Enabled = blnEnabled
                End With
            End If
        Case conMenu_Merge
            If mblnView Then
                Control.Visible = False
            Else
                Control.Checked = mblnRowMerge
                blnEnabled = True
                If Not fraAdvice.Enabled Then blnEnabled = False
                Control.Enabled = blnEnabled
            End If
        Case conMenu_Import
            If mblnView Then Control.Visible = False
        Case conMenu_Save
            If mblnView Then
                Control.Visible = False
            Else
                Control.Enabled = mblnNoSave
            End If
    End Select
    
End Sub

Private Sub chkMedicineVariety_Click()
    '取消按品种输入
    If chkMedicineVariety.Tag = "" And Trim(txt医嘱内容.Text) <> "" Then
        If MsgBox("你确定要清除当前医嘱内容重新输入吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            txt医嘱内容.Text = ""
            If txt医嘱内容.Enabled Then txt医嘱内容.SetFocus
        End If
    End If
End Sub

Private Sub chkMedicineVariety_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextControl
    End If
End Sub

Private Sub cmd频率_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str范围 As String, int频率 As Integer, vRect As RECT
    Dim lng诊疗项目ID As Long, lngFind As Long
        
    With vsAdvice
        If cbo期效.ListIndex = 1 Then
            int频率 = Get项目频率(.Row)
            If Not RowIn配方行(.Row) And int频率 = 0 Then
                str范围 = "1,-1" '临嘱可以为一次性
            Else
                str范围 = Get频率范围(.Row)
            End If
        Else
            str范围 = Get频率范围(.Row)
            int频率 = decode(str范围, "1", 0, "2", 0, "-1", 1, "-2", 2, "-3", 1, "-5", 1)
        End If
        
        '可选择频率的常用频率
        lng诊疗项目ID = Val(.TextMatrix(.Row, COL_诊疗项目ID))
        If RowIn检验行(.Row) Then
            lngFind = .FindRow(CStr(.RowData(.Row)), .FixedRows, COL_相关ID)
            If lngFind <> -1 Then
                lng诊疗项目ID = Val(.TextMatrix(lngFind, COL_诊疗项目ID))
            End If
        End If
        strSql = ""
        If InStr("," & str范围 & ",", ",1,") > 0 Then
            strSql = " And (Exists(Select 1 From 诊疗用法用量 Where 项目ID=[2] And 用法ID is NULL And 频次=A.编码 And A.适用范围=1)" & _
                " Or (Select Count(*) From 诊疗用法用量 Where 项目ID=[2] And 用法ID is NULL And 频次 Is Not NULL)<=1)"
        End If
        strSql = _
            " Select Rownum as ID,A.编码,A.名称,A.简码," & _
            " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位,A.适用范围 as 范围ID" & _
            " From 诊疗频率项目 A" & _
            " Where (Instr([1],','||A.适用范围||',')>0  Or a.适用范围=[3])" & strSql & _
            " Order by A.适用范围,A.编码"
        vRect = zlControl.GetControlRect(txt频率.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, "诊疗频率", False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt频率.Height, blnCancel, False, True, "," & str范围 & ",", lng诊疗项目ID, IIF(cbo期效.ListIndex = 1, -5, -3))
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有可用的诊疗频率项目，请先到医嘱频率管理中设置。", vbInformation, gstrSysName
            End If
            txt频率.Text = .TextMatrix(.Row, COL_频率)
            Call zlControl.TxtSelAll(txt频率)
            txt频率.SetFocus: Exit Sub
        End If
        Call Set频率Input(rsTmp, rsTmp!范围ID, int频率)
        txt频率.SetFocus
        Call SeekNextControl
    End With
End Sub

Private Sub cmd适用证候_Click()
    Dim strSql As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vPoint As PointAPI
    
    strSql = _
            " Select ID,ID as 项目ID,编码,附码,名称," & IIF(mint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
            " From 疾病编码目录" & _
            " Where 类别='Z' " & _
            " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by 编码"
                
    vPoint = zlControl.GetCoordPos(txt适用证候.Hwnd, 0, 0)
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, "中医证候", False, "", "", False, False, True, _
        vPoint.X, vPoint.Y, txt适用证候.Height, blnCancel, False, True)
    If Not blnCancel Then '无匹配输入时,按任意输入处理,取消不同
        '检查诊断输入方式
        If rsTmp Is Nothing Then
            MsgBox "没有找到中医证候疾病。", vbInformation, gstrSysName
        Else
            txt适用证候.Text = rsTmp!名称 & ""
            txt适用证候.Tag = rsTmp!ID & ""
        End If
    End If
    '更新数据
    Call AdviceChange
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strMsg As String
    
    If mblnNoSave Then
        strMsg = "当前成套方案内容编辑后尚未保存，确实要退出吗？"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Sub lbl滴速_Click()
    Call Load输液滴速(cbo滴速, lbl滴速单位, True)
    cbo滴速.Tag = "1"
    Call AdviceChange
End Sub

Private Sub tbrFree_ButtonClick(ByVal Button As MSComctlLib.Button)
    '强起时清除已有内容
    If Button.value = 0 Then
        If vsAdvice.RowData(vsAdvice.Row) <> 0 Then
            If MsgBox("取消自由录入状态将清除已录入的医嘱内容，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Button.value = 1
                Call zlControl.TxtSelAll(txt医嘱内容)
                txt医嘱内容.SetFocus: Exit Sub
            End If
            Call DeleteRow(vsAdvice.Row, True)
            mblnNoSave = True
            Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        End If
    End If
    
    txt医嘱内容.Text = ""
    txt医嘱内容.SetFocus
End Sub

Private Sub txt频率_GotFocus()
    Call zlControl.TxtSelAll(txt频率)
End Sub

Private Sub txt频率_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str范围 As String, int频率 As Integer, vRect As RECT
    Dim lng诊疗项目ID As Long, lngFind As Long
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If cmd频率.Tag <> "" And txt频率.Text = .TextMatrix(.Row, COL_频率) And txt频率.Text <> "" Then
                Call SeekNextControl
            ElseIf txt频率.Text = "" Then
                If cmd频率.Enabled And cmd频率.Visible Then cmd频率_Click
            Else
                If cbo期效.ListIndex = 1 Then
                    int频率 = Get项目频率(.Row)
                    If Not RowIn配方行(.Row) And int频率 = 0 Then
                        str范围 = "1,-1" '临嘱可以为一次性
                    Else
                        str范围 = Get频率范围(.Row)
                    End If
                Else
                    str范围 = Get频率范围(.Row)
                    int频率 = int频率 = decode(str范围, "1", 0, "2", 0, "-1", 1, "-2", 2, "-3", 1, "-5", 1)
                End If
                
                '可选择频率的常用频率
                lng诊疗项目ID = Val(.TextMatrix(.Row, COL_诊疗项目ID))
                If RowIn检验行(.Row) Then
                    lngFind = .FindRow(CStr(.RowData(.Row)), .FixedRows, COL_相关ID)
                    If lngFind <> -1 Then
                        lng诊疗项目ID = Val(.TextMatrix(lngFind, COL_诊疗项目ID))
                    End If
                End If
                strSql = ""
                If InStr("," & str范围 & ",", ",1,") > 0 Then
                    strSql = " And (Exists(Select 1 From 诊疗用法用量 Where 项目ID=[4] And 用法ID is NULL And 频次=A.编码 And A.适用范围=1)" & _
                        " Or (Select Count(*) From 诊疗用法用量 Where 项目ID=[4] And 用法ID is NULL And 频次 Is Not NULL)<=1)"
                End If
                strSql = _
                    " Select Rownum as ID,A.编码,A.名称,A.简码," & _
                    " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位,A.适用范围 as 范围ID" & _
                    " From 诊疗频率项目 A" & _
                    " Where (Instr([3],','||A.适用范围||',')>0   Or a.适用范围=[5])" & strSql & _
                    " And (A.编码 Like [1] Or Upper(A.名称) Like [2]" & _
                    " Or Upper(A.简码) Like [2] Or Upper(A.英文名称) Like [2])" & _
                    " Order by A.适用范围,A.编码"
                vRect = zlControl.GetControlRect(txt频率.Hwnd)
                Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, "诊疗频率", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt频率.Height, blnCancel, False, True, UCase(txt频率.Text) & "%", _
                    mstrLike & UCase(txt频率.Text) & "%", "," & str范围 & ",", lng诊疗项目ID, IIF(cbo期效.ListIndex = 1, -5, -3))
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "未找到匹配的诊疗频率项目。", vbInformation, gstrSysName
                    End If
                    txt频率.Text = .TextMatrix(.Row, COL_频率)
                    Call zlControl.TxtSelAll(txt频率)
                    txt频率.SetFocus: Exit Sub
                End If
                Call Set频率Input(rsTmp, rsTmp!范围ID, int频率)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Function GetBaseRow(ByVal lngRow As Long) As Long
'功能：由当前可见行获取主项目的行
    If RowIn配方行(lngRow) Then
        '获取中药配方第一味中药行
        GetBaseRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
    ElseIf RowIn检验行(lngRow) Then
        '获取一并采样的第一个项目行
        GetBaseRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
    Else
        GetBaseRow = lngRow
    End If
End Function

Private Function Get项目频率(ByVal lngRow As Long) As Integer
'功能：获取指定项目的原始执行频率属性
'参数：lngRow=当前可见行
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    lngRow = GetBaseRow(lngRow)
    strSql = "Select 执行频率 From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID)))
    If Not rsTmp.EOF Then Get项目频率 = NVL(rsTmp!执行频率, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmd用法_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim int类型 As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim lng项目id As Long
    Dim strWhere As String

    With vsAdvice
        If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
            int类型 = 2 '给药途径
        ElseIf RowIn检验行(vsAdvice.Row) Then
            int类型 = 6 '采集方法
        ElseIf .TextMatrix(.Row, COL_类别) = "K" Then
            If gbln血库系统 = True Then
                If Val(.TextMatrix(.Row, COL_检查方法)) = 0 Then
                    int类型 = 9 '采集输血途径
                Else
                    int类型 = 8 '输血途径
                    strWhere = " And nvl(A.执行分类,0)=1 "
                End If
            Else
                int类型 = 8 '输血途径
            End If
        Else
            int类型 = 4 '中药用法
        End If
        lng项目id = Val(.TextMatrix(.Row, COL_诊疗项目ID))
        If int类型 = 2 Then '只取有效范围的给药途径(无设置或仅一个时可任选)
            If Val(.TextMatrix(.Row, COL_收费细目ID)) = 0 Then
                strSql = " And (A.ID IN(Select 用法ID From 诊疗用法用量 Where 项目ID=[2] And 性质>0)" & _
                    " Or (Select Count(A.用法ID) From 诊疗用法用量 A,诊疗项目目录 B" & _
                        " Where A.用法ID=B.ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([3],3)") & " And A.项目ID=[2] And A.性质>0)<=1)"
            Else
                lng项目id = Val(.TextMatrix(.Row, COL_收费细目ID))
                strSql = " And (A.ID IN (Select 用法ID From 药品用法用量 Where 药品ID=[2] And 性质=1)" & _
                    " Or (Select Count(A.用法ID) From 药品用法用量 A,诊疗项目目录 B" & _
                        " Where A.用法ID=B.ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([3],3)") & " And A.药品ID=[2] And A.性质=1)<=1)"
            End If
        End If
        strSql = "Select Distinct A.ID,A.编码,A.名称,C.名称 as 分类,A.执行分类 as 执行分类ID " & _
            " From 诊疗项目别名 B,诊疗项目目录 A,诊疗分类目录 C" & _
            " Where A.ID=B.诊疗项目ID And A.分类ID=C.ID(+)" & _
            " And A.类别='E' And A.操作类型=[1] And " & IIF(mint范围 = 3, "Nvl(A.服务对象,0)<>0", "A.服务对象 IN([3],3)") & strWhere & strSQL & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " Order by A.编码"
        vRect = zlControl.GetControlRect(txt用法.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lbl用法.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, CStr(int类型), lng项目id, mint范围)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有可用的" & lbl用法.Caption & "，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
            txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法))
            Call zlControl.TxtSelAll(txt用法)
            txt用法.SetFocus: Exit Sub
        End If
        
        '对一并给药的其它药品的可用给药途径进行检查
        If int类型 = 2 Then
            Call Get一并给药范围(Val(.TextMatrix(.Row, COL_相关ID)), lngBegin, lngEnd)
            For i = lngBegin To lngEnd
                If i <> .Row And .RowData(i) <> 0 Then
                    If Not Check适用用法(rsTmp!ID, Val(.TextMatrix(i, COL_诊疗项目ID)), mint范围) Then
                        .Refresh
                        MsgBox """" & rsTmp!名称 & """不适用于与当前药品一并给药的""" & .TextMatrix(i, col_医嘱内容) & """。", vbInformation, gstrSysName
                        .Refresh
                        txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法))
                        Call zlControl.TxtSelAll(txt用法)
                        txt用法.SetFocus: Exit Sub
                    End If
                End If
            Next
        End If
        
        Call Set用法Input(rsTmp, int类型)
        txt用法.SetFocus
        Call SeekNextControl
    End With
End Sub

Private Sub txt适用证候_GotFocus()
    zlControl.TxtSelAll txt适用证候
End Sub

Private Sub txt适用证候_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextControl
    End If
End Sub



Private Sub txt适用证候_Validate(Cancel As Boolean)
    Dim strSql As String, rsTmp As Recordset
    Dim strInput As String, blnCancel As Boolean, vPoint As PointAPI
    
    strInput = UCase(txt适用证候.Text)
    If strInput = "" Then Exit Sub
    If zlCommFun.IsCharChinese(strInput) Then
        strSql = "名称 Like [2]" '输入汉字时只匹配名称
    Else
        strSql = "编码 Like [1] Or 名称 Like [2] Or " & IIF(mint简码 = 0, "简码", "五笔码") & " Like [2]"
    End If
    strSql = _
            " Select ID,ID as 项目ID,编码,附码,名称," & IIF(mint简码 = 0, "简码", "五笔码 as 简码") & ",说明" & _
            " From 疾病编码目录" & _
            " Where 类别='Z' And (" & strSql & ")" & _
            " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by 编码"
                
    vPoint = zlControl.GetCoordPos(txt适用证候.Hwnd, 0, 0)
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, "中医证候", False, "", "", False, False, True, _
        vPoint.X, vPoint.Y, txt适用证候.Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
        Cancel = True
    Else
        '检查诊断输入方式
        If rsTmp Is Nothing Then
            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
            Cancel = True
        Else
            txt适用证候.Text = rsTmp!名称 & ""
            txt适用证候.Tag = rsTmp!ID & ""
        End If
    End If
    '更新数据
    Call AdviceChange
End Sub

Private Sub txt天数_Change()
    txt天数.Tag = "1"
End Sub

Private Sub txt天数_GotFocus()
    Call zlControl.TxtSelAll(txt天数)
End Sub

Private Sub txt天数_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        '为临嘱
        If (IsNumeric(txt单量.Text) Or txt单量.Text = "") _
            And (IsNumeric(txt天数.Text) Or txt天数.Text = "") Then
            If SeekNextControl Then Call txt天数_Validate(False)
        End If
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt天数_Validate(Cancel As Boolean)
    Dim sng天数 As Single, i As Long
    Dim strSame As String, strMsg As String
        
    If txt天数.Text <> "" Then
        With vsAdvice
            If Val(txt天数.Text) = 0 Then
                txt天数.Text = 1: txt天数.Tag = "1"
            End If
            
            '天数至少需要一个频率同期的天数
            If Val(.TextMatrix(.Row, COL_频率间隔)) <> 0 Then
                If .TextMatrix(.Row, COL_间隔单位) = "周" Then
                    sng天数 = 7
                ElseIf .TextMatrix(.Row, COL_间隔单位) = "天" Then
                    sng天数 = Val(.TextMatrix(.Row, COL_频率间隔))
                ElseIf .TextMatrix(.Row, COL_间隔单位) = "小时" Then
                    sng天数 = Val(.TextMatrix(.Row, COL_频率间隔)) \ 24
                ElseIf .TextMatrix(.Row, COL_间隔单位) = "分钟" Then
                    sng天数 = Val(.TextMatrix(.Row, COL_频率间隔)) \ (24 * 60)
                End If
                If Val(txt天数.Text) < sng天数 Then
                    If MsgBox("按""" & .TextMatrix(.Row, COL_频率) & """执行时，至少需要 " & sng天数 & " 天的用药，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt天数_GotFocus: Exit Sub
                    End If
                End If
            End If
    
            '重新计算总量
            If .TextMatrix(.Row, COL_频率) <> "" _
                And Val(.TextMatrix(.Row, COL_单量)) <> 0 _
                And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 _
                And Val(.TextMatrix(.Row, COL_包装系数)) <> 0 Then
                
                txt总量.Text = FormatEx(Calc缺省药品总量( _
                    Val(.TextMatrix(.Row, COL_单量)), Val(txt天数.Text), _
                    Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), _
                    .TextMatrix(.Row, COL_间隔单位), .TextMatrix(.Row, COL_执行时间), _
                    Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_包装系数)), _
                    Val(.TextMatrix(.Row, COL_可否分零))), 5)
                txt总量.Tag = "1"
            End If
        End With
        
        '每次输入天数后，作为下次的缺省
        If txt天数.Tag = "1" Then
            msng天数 = Val(txt天数.Text)
        End If
    End If
    
    Call AdviceChange
End Sub

Private Sub txt用法_GotFocus()
    Call zlControl.TxtSelAll(txt用法)
End Sub

Private Sub txt用法_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim int类型 As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long
    Dim strLike As String, i As Long
    Dim lng项目id As Long
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Val(cmd用法.Tag) <> 0 And txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法)) And txt用法.Text <> "" Then
                Call SeekNextControl
            ElseIf txt用法.Text = "" Then
                If cmd用法.Enabled And cmd用法.Visible Then cmd用法_Click
            Else
                If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                    int类型 = 2 '给药途径
                ElseIf RowIn检验行(vsAdvice.Row) Then
                    int类型 = 6 '采集方法
                ElseIf .TextMatrix(.Row, COL_类别) = "K" Then
                    int类型 = 8 '输血途径
                Else
                    int类型 = 4 '中药用法
                End If
                lng项目id = Val(.TextMatrix(.Row, COL_诊疗项目ID))
                If int类型 = 2 Then '只取有效范围的给药途径(无设置或仅一个时可任选)
                    If Val(.TextMatrix(.Row, COL_收费细目ID)) = 0 Then
                        strSql = " And (A.ID IN(Select 用法ID From 诊疗用法用量 Where 项目ID=[4] And 性质>0)" & _
                            " Or (Select Count(A.用法ID) From 诊疗用法用量 A,诊疗项目目录 B" & _
                                " Where A.用法ID=B.ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([6],3)") & " And A.项目ID=[4] And A.性质>0)<=1)"
                    Else
                        lng项目id = Val(.TextMatrix(.Row, COL_收费细目ID))
                        strSql = " And (A.ID IN(Select 用法ID From 药品用法用量 Where 药品ID=[4] And 性质=1)" & _
                            " Or (Select Count(A.用法ID) From 药品用法用量 A,诊疗项目目录 B" & _
                                " Where A.用法ID=B.ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([6],3)") & " And A.药品ID=[4] And A.性质=1)<=1)"
                    End If
                End If
                
                '优化
                strLike = mstrLike
                If Len(txt用法.Text) < 2 Then strLike = ""
                
                strSql = "Select Distinct A.ID,A.编码,A.名称,A.执行分类 as 执行分类ID " & _
                    " From 诊疗项目目录 A,诊疗项目别名 B" & _
                    " Where A.ID=B.诊疗项目ID" & _
                    " And A.类别='E' And A.操作类型=[3] And " & IIF(mint范围 = 3, "Nvl(A.服务对象,0)<>0", "A.服务对象 IN([6],3)") & strSql & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2])" & _
                    decode(mint简码, 0, " And B.码类 IN([5],3)", 1, " And B.码类 IN([5],3)", "") & _
                    " Order by A.编码"
                vRect = zlControl.GetControlRect(txt用法.Hwnd)
                Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lbl用法.Caption, False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, UCase(txt用法.Text) & "%", _
                    strLike & UCase(txt用法.Text) & "%", CStr(int类型), lng项目id, mint简码 + 1, mint范围)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "未找到匹配的" & lbl用法.Caption & "。", vbInformation, gstrSysName
                    End If
                    txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法))
                    Call zlControl.TxtSelAll(txt用法)
                    txt用法.SetFocus: Exit Sub
                End If
                
                '对一并给药的其它药品的可用给药途径进行检查
                If int类型 = 2 Then
                    Call Get一并给药范围(Val(.TextMatrix(.Row, COL_相关ID)), lngBegin, lngEnd)
                    For i = lngBegin To lngEnd
                        If i <> .Row And .RowData(i) <> 0 Then
                            If Not Check适用用法(rsTmp!ID, Val(.TextMatrix(i, COL_诊疗项目ID)), mint范围) Then
                                .Refresh
                                MsgBox """" & rsTmp!名称 & """不适用于与当前药品一并给药的""" & .TextMatrix(i, col_医嘱内容) & """。", vbInformation, gstrSysName
                                .Refresh
                                txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法))
                                Call zlControl.TxtSelAll(txt用法)
                                txt用法.SetFocus: Exit Sub
                            End If
                        End If
                    Next
                End If
                
                Call Set用法Input(rsTmp, int类型)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Sub txt用法_Validate(Cancel As Boolean)
    With vsAdvice
        '恢复人为的清除
        If Val(cmd用法.Tag) <> 0 And txt用法.Text <> IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法)) Then
            txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法))
        End If
    End With
End Sub

Private Sub txt频率_Validate(Cancel As Boolean)
    With vsAdvice
        '恢复人为的清除
        If cmd频率.Tag <> "" And txt频率.Text <> .TextMatrix(.Row, COL_频率) Then
            txt频率.Text = .TextMatrix(.Row, COL_频率)
        End If
    End With
End Sub

Private Sub cbo执行科室_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSql As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
    Dim lng配制中心 As Long, str药房IDs As String
    Dim lngBegin As Long, lngEnd As Long
    
    If cbo执行科室.ListIndex = -1 Then Exit Sub
    
    If cbo执行科室.ItemData(cbo执行科室.ListIndex) = -1 Then
        strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
            " From 部门表 A,部门性质说明 B" & _
            " Where A.ID=B.部门ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([1],3)") & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " Order by A.编码"
        vRect = zlControl.GetControlRect(cbo执行科室.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lbl执行科室.Caption, False, "", "", False, False, True, vRect.Left, vRect.Top, cbo执行科室.Height, blnCancel, False, True, mint范围)
        If Not rsTmp Is Nothing Then
            intIdx = Cbo.FindIndex(cbo执行科室, rsTmp!ID)
            If intIdx <> -1 Then
                cbo执行科室.ListIndex = intIdx
            Else
                cbo执行科室.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo执行科室.ListCount - 1
                cbo执行科室.ItemData(cbo执行科室.NewIndex) = rsTmp!ID
                cbo执行科室.ListIndex = cbo执行科室.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
            '恢复成现有的科室(不引发Click)
            intIdx = Cbo.FindIndex(cbo执行科室, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_执行科室ID)))
            Call Cbo.SetIndex(cbo执行科室.Hwnd, intIdx)
        End If
    Else
        lngRow = vsAdvice.Row
        
        '检查一并给药的配制中心
        With vsAdvice
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 And RowIn一并给药(lngRow) Then
                Call Get一并给药范围(Val(.TextMatrix(lngRow, COL_相关ID)), lngBegin, lngEnd)
                
                '当前行由普通药房或其他配制中心改为配制中心
                If sys.DeptHaveProperty(cbo执行科室.ItemData(cbo执行科室.ListIndex), "配制中心") Then
                    lng配制中心 = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                End If
                '当前行由配置中心或改为普通药房
                If lng配制中心 = 0 Then
                    For i = lngBegin To lngEnd
                        If i <> lngRow And Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                            '自备药不管它
                            If Not (Val(.TextMatrix(i, COL_执行科室ID)) = 0 And Val(.TextMatrix(i, COL_执行性质)) = 5) Then
                                If sys.DeptHaveProperty(Val(.TextMatrix(i, COL_执行科室ID)), "配制中心") Then
                                    lng配制中心 = Val(.TextMatrix(i, COL_执行科室ID)): Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                '这两种情况所有药品都执行科室相同，检查存储设定
                If lng配制中心 <> 0 Then
                    For i = lngBegin To lngEnd
                        If i <> lngRow And Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                            '自备药不管它
                            If Not (Val(.TextMatrix(i, COL_执行科室ID)) = 0 And Val(.TextMatrix(i, COL_执行性质)) = 5) Then
                                str药房IDs = Get可用药房IDs(.TextMatrix(i, COL_类别), Val(.TextMatrix(i, COL_诊疗项目ID)), Val(.TextMatrix(i, COL_收费细目ID)), 0, mint范围)
                                If InStr("," & str药房IDs & ",", "," & cbo执行科室.ItemData(cbo执行科室.ListIndex) & ",") = 0 Then
                                    MsgBox "一并给药的药品中，""" & .TextMatrix(i, col_医嘱内容) & """在""" & zlCommFun.GetNeedName(cbo执行科室.Text) & """中没有存储。", vbInformation, gstrSysName
                                    '恢复成现有的科室(不引发Click)
                                    intIdx = Cbo.FindIndex(cbo执行科室, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_执行科室ID)))
                                    Call Cbo.SetIndex(cbo执行科室.Hwnd, intIdx)
                                    Exit Sub
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End With
        
        cbo执行科室.Tag = "1"
        
        '更新更改了的执行科室医嘱内容
        Call AdviceChange
    End If
End Sub

Private Sub cbo执行科室_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo执行科室.ListIndex = -1 Then
            Call cbo执行科室_Validate(blnCancel)
        End If
        If Not blnCancel Then
            If SeekNextControl Then Call cbo执行科室_Validate(False)
        End If
    End If
End Sub

Private Sub cbo执行科室_GotFocus()
    Call zlControl.TxtSelAll(cbo执行科室)
End Sub

Private Sub cbo执行科室_Validate(Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If cbo执行科室.ListIndex <> -1 Then Exit Sub '已选中
    If cbo执行科室.Text = "" Then '不输入
        cbo执行科室.Tag = "1"
        Call AdviceChange
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '是否可以任意或选择科室
    blnLimit = True
    If cbo执行科室.ListCount > 0 Then
        If cbo执行科室.ItemData(cbo执行科室.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    strInput = UCase(zlCommFun.GetNeedName(cbo执行科室.Text))
    strSql = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([3],3)") & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And (A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2])" & _
        " Order by A.编码"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, strInput & "%", mstrLike & strInput & "%", mint范围)
        For i = 1 To rsTmp.RecordCount
            intIdx = Cbo.FindIndex(cbo执行科室, rsTmp!ID)
            If intIdx <> -1 Then cbo执行科室.ListIndex = intIdx: Exit For
            rsTmp.MoveNext
        Next
        If cbo执行科室.ListIndex = -1 Then
            MsgBox "未到对应的科室。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = zlControl.GetControlRect(cbo执行科室.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lbl执行科室.Caption, False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", mint范围)
        If Not rsTmp Is Nothing Then
            intIdx = Cbo.FindIndex(cbo执行科室, rsTmp!ID)
            If intIdx <> -1 Then
                cbo执行科室.ListIndex = intIdx
            Else
                cbo执行科室.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo执行科室.ListCount - 1
                cbo执行科室.ItemData(cbo执行科室.NewIndex) = rsTmp!ID
                cbo执行科室.ListIndex = cbo执行科室.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "未找到对应的科室。", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo执行时间_Change()
    cbo执行时间.Tag = "1"
End Sub

Private Sub cbo执行时间_Click()
    'cbo执行时间_Change
    '更新数据
    cbo执行时间.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cbo执行时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cbo执行时间_Validate(False)
    Else
        If InStr("0123456789:-/" & Chr(8) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cbo执行时间_Validate(Cancel As Boolean)
    Dim blnValid As Boolean, lngRow As Long, strTmp As String
    
    lngRow = vsAdvice.Row
        
    With vsAdvice
        If cbo执行时间.Text <> "" Then
            '检查长度
            If Len(cbo执行时间.Text) > 50 Then
                MsgBox "输入内容不能超过 50 个字符。", vbInformation, gstrSysName
                Call cbo执行时间_GotFocus
                Cancel = True: Exit Sub
            End If
            '检查合法性
            If .RowData(lngRow) <> 0 Then
                blnValid = ExeTimeValid(cbo执行时间.Text, Val(.TextMatrix(lngRow, COL_频率次数)), Val(.TextMatrix(lngRow, COL_频率间隔)), .TextMatrix(lngRow, COL_间隔单位))
                If Not blnValid Then
                    If .TextMatrix(lngRow, COL_间隔单位) = "周" Then
                        strTmp = COL_按周执行
                    ElseIf .TextMatrix(lngRow, COL_间隔单位) = "天" Then
                        strTmp = COL_按天执行
                    ElseIf .TextMatrix(lngRow, COL_间隔单位) = "小时" Then
                        strTmp = COL_按时执行
                    End If
                    MsgBox "输入的执行时间方案格式不正确，请检查。" & vbCrLf & vbCrLf & "例：" & vbCrLf & strTmp, vbInformation, gstrSysName
                    Call cbo执行时间_GotFocus
                    Cancel = True: Exit Sub
                End If
            End If
        End If
    End With
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub cbo执行性质_Click()
    cbo执行性质.Tag = "1"
    '更新数据
    Call AdviceChange
End Sub

Private Sub cbo执行性质_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo执行性质.ListIndex <> -1 Then
            Call SeekNextControl
        End If
    ElseIf KeyAscii >= 32 Then
        lngIdx = Cbo.MatchIndex(cbo执行性质.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo执行性质.ListCount > 0 Then lngIdx = 0
        cbo执行性质.ListIndex = lngIdx
    End If
End Sub

Private Sub cmdExt_Click()
'功能：修改现有医嘱的扩充内容
    Dim rsCurr As New ADODB.Recordset
    Dim strExtData As String, strTmp As String
    Dim lngRow As Long, lngFirstRow As Long
    Dim lng诊疗项目ID As Long, lng用法ID As Long, str缺省 As String
    Dim strMsg As String, vMsg As VbMsgBoxResult
    Dim lngBegin As Long, lngEnd As Long, i As Long, blnRefresh As Boolean
    Dim lng配方ID As Long
    Dim intType As Integer, lng项目id As Long, blnOK As Boolean
    Dim t_Pati As TYPE_PatiInfoEx
    
    lngRow = vsAdvice.Row
        
    If vsAdvice.TextMatrix(lngRow, COL_类别) = "D" Then
        strExtData = Get检查部位方法(lngRow)
        If strExtData = "" Then
            MsgBox "该检查医嘱是系统升级以前下达的，与现有方式不兼容。请重新下达该检查医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
        intType = 0
    ElseIf vsAdvice.TextMatrix(lngRow, COL_类别) = "F" Then
        strExtData = Get手术附加IDs(lngRow)
        intType = 1
    ElseIf RowIn配方行(lngRow) Then
        strExtData = Get中药配方IDs(lngRow)
        intType = 2
    ElseIf RowIn检验行(lngRow) Then
        strExtData = Get检验组合IDs(lngRow)
        intType = 4
    Else
        Exit Sub '兼容以前的检验项目
    End If
    
    If intType = 4 Then
        lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
        lng项目id = Val(vsAdvice.TextMatrix(lngFirstRow, COL_诊疗项目ID))
    Else
        lng项目id = Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
    End If

    On Error Resume Next
    If intType = 2 Then
        blnOK = frmAdviceFormula.ShowMe(Me, Nothing, txt医嘱内容.Hwnd, t_Pati, 3, IIF(mbyt场合 <> 2, 0, 3), cbo期效.ListIndex, mint范围, , lng项目id, strExtData)
    Else
        blnOK = frmSchemeEditEx.ShowMe(Me, txt医嘱内容.Hwnd, intType, cbo期效.ListIndex, mint范围, mblnNewLIS, False, lng项目id, strExtData)
    End If
    On Error GoTo 0
    
    '重新设置相关内容
    If blnOK Then
        str缺省 = vsAdvice.TextMatrix(lngRow, col_缺省)
        
        If vsAdvice.TextMatrix(lngRow, COL_类别) = "D" Then
            '检查组合
            Call AdviceSet检查组合(lngRow, strExtData)
            vsAdvice.TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
            txt医嘱内容.Text = vsAdvice.TextMatrix(lngRow, col_医嘱内容)
        ElseIf vsAdvice.TextMatrix(lngRow, COL_类别) = "F" Then
            '一组手术
            Call AdviceSet手术组合(lngRow, strExtData)
            vsAdvice.TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
            txt医嘱内容.Text = vsAdvice.TextMatrix(lngRow, col_医嘱内容)
            
            '刷新处理手术麻醉的执行科室
            blnRefresh = True
        ElseIf RowIn检验行(lngRow) Then
            '检验组合
            lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
            lng用法ID = Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
            
            '先获取当前已经设置好值
            rsCurr.Fields.Append "医嘱ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "频率", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "频率次数", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "频率间隔", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "间隔单位", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "总量", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "执行时间", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "医生嘱托", adVarChar, 100, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
                        
            '采集方法的执行科室可能与检验项目不同
            If Val(vsAdvice.TextMatrix(lngFirstRow, COL_执行科室ID)) <> 0 Then
                rsCurr!执行科室ID = Val(vsAdvice.TextMatrix(lngFirstRow, COL_执行科室ID))
            End If
            If Val(vsAdvice.TextMatrix(lngRow, COL_总量)) <> 0 Then
                rsCurr!总量 = Val(vsAdvice.TextMatrix(lngRow, COL_总量))
            End If
            rsCurr!执行时间 = vsAdvice.TextMatrix(lngRow, COL_执行时间)
            rsCurr!频率 = vsAdvice.TextMatrix(lngRow, COL_频率)
            rsCurr!频率次数 = Val(vsAdvice.TextMatrix(lngRow, COL_频率次数))
            rsCurr!频率间隔 = Val(vsAdvice.TextMatrix(lngRow, COL_频率间隔))
            rsCurr!间隔单位 = vsAdvice.TextMatrix(lngRow, COL_间隔单位)
            rsCurr!医生嘱托 = vsAdvice.TextMatrix(lngRow, COL_医生嘱托)
            rsCurr!医嘱ID = vsAdvice.RowData(lngRow)
            rsCurr.Update
            
            '完全重新设置该检验组合
            '------------------------
            '删除检验项目行:删除之后重新定位的当前行
            lngRow = Delete检验组合(lngRow)
            '清除当前行(采集方法行)
            Call DeleteRow(lngRow, True, False)
            '重新产生:产生之后重新定位的当前行
            lngRow = AdviceSet检验组合(lngRow, lng用法ID, strExtData, rsCurr)
            
            '强行显示当前医嘱卡片
            blnRefresh = True
        ElseIf RowIn配方行(lngRow) Then
            '中药配方
            lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
            lng诊疗项目ID = Val(vsAdvice.TextMatrix(lngFirstRow, COL_诊疗项目ID))
            lng用法ID = Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
            
            '先获取当前已经设置好值
            rsCurr.Fields.Append "医嘱ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "执行性质", adVarChar, 10, adFldIsNullable
            rsCurr.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "频率", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "频率次数", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "频率间隔", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "间隔单位", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "总量", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "执行时间", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "医生嘱托", adVarChar, 100, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
            
            rsCurr!执行性质 = zlCommFun.GetNeedName(cbo执行性质.Text) '正常,自备药,离院带药
             '取配方界面选择的药房
            rsCurr!执行科室ID = Val(Split(strExtData, "|")(4))
            rsCurr!频率 = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
            rsCurr!频率次数 = Val(vsAdvice.TextMatrix(lngFirstRow, COL_频率次数))
            rsCurr!频率间隔 = Val(vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔))
            rsCurr!间隔单位 = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
            If Val(vsAdvice.TextMatrix(lngFirstRow, COL_总量)) <> 0 Then
                rsCurr!总量 = Val(vsAdvice.TextMatrix(lngFirstRow, COL_总量))
            End If
            rsCurr!执行时间 = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
            rsCurr!医生嘱托 = vsAdvice.TextMatrix(lngRow, COL_医生嘱托)
            rsCurr!医嘱ID = vsAdvice.RowData(lngRow)
            
            rsCurr.Update
            
            '完全重新设置该中药配方行
            '------------------------
            '删除组成味药及煎法行:删除之后重新定位的当前行
            lngRow = Delete中药配方(lngRow)
            '如果当前用法的配方ID不为空，则传入配方ID
            lng配方ID = Val(vsAdvice.TextMatrix(lngRow, COL_配方ID))
            '清除当前行(中药用法行)
            Call DeleteRow(lngRow, True, False)
            '产生配方:产生之后重新定位的当前行
            lngRow = AdviceSet中药配方(lng诊疗项目ID, lngRow, lng用法ID, strExtData, rsCurr, lng配方ID)
            
            blnRefresh = True
        End If
        
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            vsAdvice.TextMatrix(i, col_缺省) = str缺省
        Next
    
        '刷新医嘱卡片
        If blnRefresh Then Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        
        mblnNoSave = True '标记为未保存
    End If
    
    Call vsAdvice.AutoSize(col_医嘱内容)
    
    txt医嘱内容.SetFocus
End Sub

Private Sub ClinicSelecter(Optional ByVal lng分类ID As Long)
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = frmClinicSelect.ShowSelect(Me, -1, 0, 0, cbo期效.ListIndex, "", , , mint范围, lng分类ID, , , , mstr使用科室, , mstr诊疗分类, mstr操作类型, mstr执行分类)
    If rsTmp Is Nothing Then '取消或无数据
        zlControl.TxtSelAll txt医嘱内容
        txt医嘱内容.SetFocus: Exit Sub
    End If
        
    '根据选择项目设置缺省医嘱信息
    If AdviceInput(rsTmp, vsAdvice.Row) Then
        '显示已缺省设置的值
        Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_临床自管药)) = 1 Then
            cbo执行性质.Tag = "1"
            Call AdviceChange
        End If
        txt医嘱内容.SetFocus '必须先定位
        Call SeekNextControl
    Else
        '恢复原值(AdviceInput函数中可能处理了一下)
        txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容)
        txt医嘱内容.SetFocus
    End If
End Sub

Private Sub cmdSel_Click()
    ClinicSelecter
End Sub

Private Sub Form_Activate()
    If mblnRunFirst Then
        mblnRunFirst = False
        If cbo期效.Visible And cbo期效.Enabled Then
            cbo期效.SetFocus
        ElseIf txt医嘱内容.Enabled Then
            txt医嘱内容.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            If tbrFree.Visible And tbrFree.Enabled And tbrFree.Buttons(1).Enabled And tbrFree.Buttons(1).Visible Then
                tbrFree.Buttons(1).value = IIF(tbrFree.Buttons(1).value = 1, 0, 1)
                Call tbrFree_ButtonClick(tbrFree.Buttons(1))
            End If
        Case vbKeyF4
            If Me.ActiveControl Is txt用法 Then
                If cmd用法.Visible And cmd用法.Enabled Then cmd用法_Click
            ElseIf Me.ActiveControl Is txt频率 Then
                If cmd频率.Visible And cmd频率.Enabled Then cmd频率_Click
            End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeySpace Then
        If Me.ActiveControl Is txt用法 Then
            KeyAscii = 0
            If cmd用法.Visible And cmd用法.Enabled Then cmd用法_Click
        ElseIf Me.ActiveControl Is txt频率 Then
            KeyAscii = 0
            If cmd频率.Visible And cmd频率.Enabled Then cmd频率_Click
        ElseIf Me.ActiveControl Is cbo滴速 Then
            KeyAscii = 0
            If cbo滴速.Visible And cbo滴速.Enabled Then zlCommFun.PressKey (vbKeyF4)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lngRow As Long
    Dim strErr As String
    Dim strPre As String
    
    If gobjLIS Is Nothing Then
        '如果是临床路径或成套，可能未创建LIS部件，则先创建
        Call InitObjLis(IIF(mint范围 = 1, p门诊医生站, p住院医生站))
        If gobjLIS Is Nothing Then
            mblnNewLIS = False
        Else
            On Error Resume Next
            mblnNewLIS = gobjLIS.GetApplicationFormShowType
            err.Clear: On Error GoTo 0
        End If
    Else
        On Error Resume Next
        mblnNewLIS = gobjLIS.GetApplicationFormShowType
        err.Clear: On Error GoTo 0
    End If
    Call InitCommandBar
    Call InitAdviceTable
    Call RestoreWinState(Me, App.ProductName)
    
    Call Cbo.SetListHeight(cbo滴速, Me.Height)
    Call Cbo.SetListHeight(cbo执行科室, Me.Height)
    Call Cbo.SetListWidth(cbo执行科室.Hwnd, cbo执行科室.Width * 1.3)
    
    '图标
    tbrFree.HotImageList = frmIcons.img24
    tbrFree.ImageList = frmIcons.img24
    tbrFree.Buttons(1).Image = 1
    
    tbrFree.Top = 810 '初始位置
    tbrFree.Visible = Not (mint范围 = 1)  '门诊成套目前不支持自由录入医嘱
    
    If mbyt场合 = 0 Then
        Me.Caption = "成套医嘱"
    ElseIf mbyt场合 = 1 Then
        Me.Caption = "路径医嘱"
    ElseIf mbyt场合 = 2 Then
        Me.Caption = "替换医嘱"
        tbrFree.Visible = False '批量调整,禁止自由录入
    End If
    
    mblnOK = False
    mblnNoSave = False
    mblnRowMerge = False
    mblnRunFirst = True
    mblnRowChange = True
    mlngNextID = 0
        
    '输入匹配
    mstrLike = IIF(Val(zldatabase.GetPara("输入匹配")) = 0, "%", "")
    '简码匹配方式：0-拼音,1-五笔
    mint简码 = Val(zldatabase.GetPara("简码方式"))
    
    '临嘱缺省一次性
    mbln一次性 = Val(zldatabase.GetPara("临嘱缺省一次性", glngSys, p住院医嘱下达)) <> 0
    
    If mbyt场合 <> 2 Then
        '常用滴速
        strPre = cbo滴速.Text '加入后保持原有值
        Call Load输液滴速(cbo滴速, lbl滴速单位, False)
        cbo滴速.Text = strPre
    End If
    
    If Not mblnView Then
        '常用嘱托
        Call ReadEnjoin
        
        '医嘱内容定义
        If CreateScript(mobjVBA, mobjScript) Then
            Set mrsDefine = InitAdviceDefine
        End If
    End If
    
    If mint范围 = 1 Then
        lbl期效.Enabled = False
        cbo期效.ListIndex = 1
        cbo期效.Enabled = False
    End If
    
    '读取并显示成套内容
    If Not mrsScheme Is Nothing Then
        Call LoadAdvice(0, vsAdvice.FixedRows)
    End If
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
End Sub

Private Sub InitSchemeRecordset()
    Set mrsScheme = New ADODB.Recordset
    mrsScheme.Fields.Append "是否缺省", adSmallInt
    mrsScheme.Fields.Append "是否备选", adSmallInt
    mrsScheme.Fields.Append "序号", adBigInt
    mrsScheme.Fields.Append "相关序号", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "期效", adSmallInt
    mrsScheme.Fields.Append "诊疗项目ID", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "医嘱内容", adVarChar, 1000, adFldIsNullable
    mrsScheme.Fields.Append "天数", adSingle, , adFldIsNullable
    mrsScheme.Fields.Append "单次用量", adSingle, , adFldIsNullable
    mrsScheme.Fields.Append "总给予量", adSingle, , adFldIsNullable
    mrsScheme.Fields.Append "医生嘱托", adVarChar, 1000, adFldIsNullable
    mrsScheme.Fields.Append "执行频次", adVarChar, 100, adFldIsNullable
    mrsScheme.Fields.Append "频率次数", adSmallInt, , adFldIsNullable
    mrsScheme.Fields.Append "频率间隔", adSmallInt, , adFldIsNullable
    mrsScheme.Fields.Append "间隔单位", adVarChar, 10, adFldIsNullable
    mrsScheme.Fields.Append "时间方案", adVarChar, 100, adFldIsNullable
    mrsScheme.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "执行性质", adSmallInt
    mrsScheme.Fields.Append "标本部位", adVarChar, 100, adFldIsNullable
    mrsScheme.Fields.Append "检查方法", adVarChar, 100, adFldIsNullable
    mrsScheme.Fields.Append "配方ID", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "组合项目ID", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "执行标记", adSingle, , adFldIsNullable
    If mbyt场合 = 1 Then
        mrsScheme.Fields.Append "类别", adVarChar, 1, adFldIsNullable
        mrsScheme.Fields.Append "操作类型", adVarChar, 20, adFldIsNullable
    End If
    mrsScheme.CursorLocation = adUseClient
    mrsScheme.LockType = adLockOptimistic
    mrsScheme.CursorType = adOpenStatic
    mrsScheme.Open
End Sub

Private Function ReadEnjoin() As Boolean
'功能：读取并加入常用嘱托
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strPre As String
        
    On Error GoTo errH
    
    strPre = cbo医生嘱托.Text '加入后保持原有值
    cbo医生嘱托.Clear
    
    strSql = _
        " Select 名称 From 常用嘱托 Where 名称 is Not Null And 人员=[1]" & _
        " Union" & _
        " Select 名称 From 常用嘱托 Where 名称 is Not Null And 人员 is Null" & _
        " Order by 名称"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.姓名)
    Do While Not rsTmp.EOF
        AddComboItem cbo医生嘱托.Hwnd, CB_ADDSTRING, 0, rsTmp!名称
        rsTmp.MoveNext
    Loop
    cbo医生嘱托.Text = strPre
    ReadEnjoin = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    msng天数 = 0
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mrsDefine = Nothing
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Function RowCanMerge(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional strMsg As String) As Boolean
'功能：判断两行是否可以一并给药
'参数：lngRow1=前面一条已经输入的药品行
'      lngRow2=当前行(已输入或未输入)
'返回：如果不可以，则strMsg返回提示信息
    Dim lngFind As Long
    Dim lng配制中心 As Long
    Dim str药房IDs As String
    
    With vsAdvice
        strMsg = ""
        If Not Between(lngRow1, .FixedRows, .Rows - 1) Then Exit Function
        If Not Between(lngRow2, .FixedRows, .Rows - 1) Then Exit Function
        If .RowHidden(lngRow1) Or .RowHidden(lngRow2) Then Exit Function
        If .RowData(lngRow1) = 0 Then Exit Function
        
        If .RowData(lngRow2) = 0 Then
            '必须全部为成药且类别相同
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_类别)) = 0 Then
                strMsg = "一并给药的药品必须都为西成药或都为中成药。"
                Exit Function
            End If
        ElseIf .RowData(lngRow2) <> 0 Then
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_类别)) = 0 _
                Or InStr(",5,6,", .TextMatrix(lngRow2, COL_类别)) = 0 Then
                strMsg = "一并给药的药品必须都为西成药或都为中成药。"
                Exit Function
            End If
            
            '期效必须相同
            If .TextMatrix(lngRow1, COL_期效) <> .TextMatrix(lngRow2, COL_期效) Then
                strMsg = "一并给药的药品医嘱期效必须相同。"
                Exit Function
            End If
            
            '一并给药(前面药品)的给药途径是否适用于当前药品
            lngFind = .FindRow(CLng(.TextMatrix(lngRow1, COL_相关ID)), lngRow1 + 1)
            If lngFind <> -1 Then
                If Not Check适用用法(Val(.TextMatrix(lngFind, COL_诊疗项目ID)), Val(.TextMatrix(lngRow2, COL_诊疗项目ID)), mint范围) Then
                    strMsg = """" & .TextMatrix(lngRow2, col_医嘱内容) & """不能使用""" & .TextMatrix(lngFind, col_医嘱内容) & """给药途径，" & _
                    vbCrLf & "不能与""" & .TextMatrix(lngRow1, col_医嘱内容) & """设置为一并给药。"
                    Exit Function
                End If
            End If
            
            '检查如果有配制中心，是否都可以存储，自备药不管它
            If Not (Val(.TextMatrix(lngRow1, COL_执行科室ID)) = 0 And Val(.TextMatrix(lngRow1, COL_执行性质)) = 5) Then
                If sys.DeptHaveProperty(Val(.TextMatrix(lngRow1, COL_执行科室ID)), "配制中心") Then
                    lng配制中心 = Val(.TextMatrix(lngRow1, COL_执行科室ID))
                End If
            End If
            If lng配制中心 = 0 Then
                If Not (Val(.TextMatrix(lngRow2, COL_执行科室ID)) = 0 And Val(.TextMatrix(lngRow2, COL_执行性质)) = 5) Then
                    If sys.DeptHaveProperty(Val(.TextMatrix(lngRow2, COL_执行科室ID)), "配制中心") Then
                        lng配制中心 = Val(.TextMatrix(lngRow2, COL_执行科室ID))
                    End If
                End If
            End If
            If lng配制中心 <> 0 Then
                If Not (Val(.TextMatrix(lngRow1, COL_执行科室ID)) = 0 And Val(.TextMatrix(lngRow1, COL_执行性质)) = 5) Then
                    str药房IDs = Get可用药房IDs(.TextMatrix(lngRow1, COL_类别), Val(.TextMatrix(lngRow1, COL_诊疗项目ID)), Val(.TextMatrix(lngRow1, COL_收费细目ID)), 0, mint范围)
                    If InStr("," & str药房IDs & ",", "," & lng配制中心 & ",") = 0 Then
                        strMsg = "药品""" & .TextMatrix(lngRow1, col_医嘱内容) & """在配制中心""" & sys.RowValue("部门表", lng配制中心, "名称") & """没有存储。"
                        Exit Function
                    End If
                End If
                If Not (Val(.TextMatrix(lngRow2, COL_执行科室ID)) = 0 And Val(.TextMatrix(lngRow2, COL_执行性质)) = 5) Then
                    str药房IDs = Get可用药房IDs(.TextMatrix(lngRow2, COL_类别), Val(.TextMatrix(lngRow2, COL_诊疗项目ID)), Val(.TextMatrix(lngRow2, COL_收费细目ID)), 0, mint范围)
                    If InStr("," & str药房IDs & ",", "," & lng配制中心 & ",") = 0 Then
                        strMsg = "药品""" & .TextMatrix(lngRow2, col_医嘱内容) & """在配制中心""" & sys.RowValue("部门表", lng配制中心, "名称") & """没有存储。"
                        Exit Function
                    End If
                End If
            End If
        End If
    End With
    RowCanMerge = True
End Function

Private Sub MoveCurrRow(ByVal lngRow As Long, ByVal lngWay As Long)
'功能：将当前行上移或下移一行
'参数：lngRow=当前行
'      lngWay=1上移一行,-1下移一行(相当于下一行上移一行)
    Dim lngPreRow As Long, lngNextRow As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngUpBegin As Long, lngUpEnd As Long
    Dim lngDownBegin As Long, lngDownEnd As Long
    Dim i As Long, j As Long
    Dim lngMoveRows As Long, blnRedraw As Boolean
   
    With vsAdvice
        If .RowData(lngRow) = 0 Then Exit Sub   '空白行排除
        '当前行可能是一并给药中间的行
        Call GetRowScope(lngRow, lngBegin, lngEnd)
                
        If lngWay = 1 Then
            lngPreRow = GetPreRow(lngBegin)
            If lngPreRow = -1 Then Exit Sub
          
            lngDownBegin = lngBegin
            lngDownEnd = lngEnd
            Call GetRowScope(lngPreRow, lngUpBegin, lngUpEnd)
            lngMoveRows = lngDownBegin - lngUpBegin
        Else
            lngNextRow = GetNextRow(lngEnd)
            If lngNextRow = -1 Then Exit Sub
            
            lngUpBegin = lngBegin
            lngUpEnd = lngEnd
            Call GetRowScope(lngNextRow, lngDownBegin, lngDownEnd)
            lngMoveRows = lngDownEnd - lngUpEnd
        End If
        
        blnRedraw = .Redraw
        .Redraw = False
        
        j = 0
        For i = lngDownBegin To lngDownEnd
            .RowPosition(i) = lngUpBegin + j
            j = j + 1
        Next
               
        mblnRowChange = False
        lngRow = lngRow - lngWay * lngMoveRows
        .Row = lngRow
        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col): .TopRow = .Row
        mblnRowChange = True
         
        mblnNoSave = True '标记为未保存
        .Redraw = blnRedraw
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lng医嘱ID As Long, lng相关ID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngPreRow As Long, strMsg As String
    Dim str期效 As String, lng诊疗项目ID As Long
    Dim lngTmp As Long, i As Long, j As Long
    
    If Not Control.Visible Then Exit Sub 'Visible=False时通过热键居然也会执行
    
    Call AdviceChange '强制更新医嘱内容
    
    With vsAdvice
        Select Case Control.ID
            Case conMenu_MoveUp
                Call MoveCurrRow(.Row, 1)
            Case conMenu_MoveDown
                Call MoveCurrRow(.Row, -1)
            Case conMenu_New
                If .RowData(.Row) = 0 Then
                ElseIf .RowData(.Rows - 1) = 0 Then
                    .Row = .Rows - 1
                Else
                    '先删除中间间隔的空行
                    mblnRowChange = False
                    For i = .Rows - 1 To .FixedRows Step -1
                        If .RowData(i) = 0 Then .RemoveItem i
                    Next
                    mblnRowChange = True
                    
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    .Col = .FixedCols
                End If
                
                Call .ShowCell(.Row, .Col)
                If Visible Then
                    If cbo期效.Visible And cbo期效.Enabled Then
                        cbo期效.SetFocus
                    ElseIf txt医嘱内容.Enabled Then
                        txt医嘱内容.SetFocus
                    End If
                End If
            Case conMenu_Insert
                If .RowData(.Row) = 0 Then
                    MsgBox "当前行无内容，请先在当前行录入有效医嘱。", vbInformation, gstrSysName
                    Exit Sub
                End If
                            
                lngPreRow = GetPreRow(.Row)
                            
                '插入后成自动成为一并给药:插入在一并给药的中间才行
                If lngPreRow <> -1 Then
                    If Val(.TextMatrix(lngPreRow, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) _
                        And Val(.TextMatrix(lngPreRow, COL_相关ID)) <> 0 And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                        
                        lng相关ID = Val(.TextMatrix(lngPreRow, COL_相关ID))
                    End If
                End If
                
                '先删除中间间隔的空行
                mblnRowChange = False
                lng医嘱ID = .RowData(.Row)
                For i = .Rows - 1 To .FixedRows Step -1
                    If .RowData(i) = 0 Then .RemoveItem i
                Next
                .Row = .FindRow(lng医嘱ID)
                mblnRowChange = True
                            
                '当前行之前插入新行
                '--------------------------------------------------------------
                If RowIn配方行(.Row) Or RowIn检验行(.Row) Then
                    '中药配方及检验组合行是前面的行隐藏
                    lngBegin = .FindRow(CStr(.RowData(.Row)), , COL_相关ID)
                Else
                    lngBegin = .Row
                End If
                
                mblnRowChange = False
                .AddItem "", lngBegin
                .Row = lngBegin
                .Col = .FixedCols
                mblnRowChange = True
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
                Call .ShowCell(.Row, .Col)
                
                If cbo期效.Visible And cbo期效.Enabled Then
                    cbo期效.SetFocus
                ElseIf txt医嘱内容.Enabled Then
                    txt医嘱内容.SetFocus
                End If
            Case conMenu_Merge '一并给药
                If Not Control.Checked Then '想按下
                    lngBegin = GetPreRow(.Row)
                    '前面没有行
                    If lngBegin = -1 Then
                        MsgBox "前面没有可以一并给药的医嘱行。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    '两行不符合条件
                    If Not RowCanMerge(lngBegin, .Row, strMsg) Then
                        MsgBox strMsg, vbInformation, gstrSysName
                        Exit Sub
                    End If
                    If .RowData(.Row) = 0 Then
                        '当前行尚未输入内容的情况
                        cbo期效.ListIndex = IIF(.TextMatrix(lngBegin, COL_期效) = "临嘱", 1, 0)
                        mblnRowMerge = True: cbsMain.RecalcLayout '*允许按下
                        txt医嘱内容.SetFocus: Exit Sub
                    Else
                        '要把当前行与前面行一起一并给药
                        Call MergeRow(lngBegin, .Row, False)
                    End If
                Else '想弹起
                    If .RowData(.Row) = 0 Then
                        '是否当前行尚未输入内容的情况
                        If Not RowIn一并给药(.Row) Then
                            mblnRowMerge = False '*允许弹起
                            cbsMain.RecalcLayout
                        End If
                        Exit Sub
                    Else
                        '当前行是一并给药中的行
                        Call Get一并给药范围(Val(.TextMatrix(.Row, COL_相关ID)), lngBegin, lngEnd)
                                                
                        '先提示
                        If Not (.Row = lngEnd And lngEnd - lngBegin > 1) Then
                            '整个一并给药取消为单独给药
                            If MsgBox("要将该组一并给药的药品全部取消为单独给药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Exit Sub
                            End If
                        End If
                        
                        '删除中间的空行
                        lngTmp = .RowData(.Row)
                        For i = lngEnd To lngBegin Step -1
                            If .RowData(i) = 0 Then
                                .RemoveItem i
                                lngEnd = lngEnd - 1
                            End If
                        Next
                        .Row = .FindRow(lngTmp, lngBegin)
                        
                        If .Row = lngEnd And lngEnd - lngBegin > 1 Then
                            '从一并给药中分离该行
                            Call SplitRow(.Row)
                        Else
                            '取消一并给药
                            lngTmp = .RowData(.Row) '记录用于恢复行定位
                            Call AdviceSet单独给药(lngBegin, lngEnd)
                            .Row = .FindRow(lngTmp)
                        End If
                    End If
                End If
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
            Case conMenu_Delete
                If .RowSel <> .Row Then
                    MsgBox "一次只能删除一条医嘱，请选择要删除的医嘱行。", vbInformation, gstrSysName
                    Exit Sub
                End If
                If .RowData(.Row) <> 0 Then
                    If MsgBox("确实要删除医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                '删除当前行
                Call AdviceDelete(.Row)
                .SetFocus
            Case conMenu_Import '导入医嘱
                strMsg = frmSchemeImport.ShowMe(Me, mint范围, lngTmp)
                If strMsg <> "" And lngTmp <> 0 Then
                    Call LoadAdvice(0, 0, strMsg, lngTmp)
                    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
                    mblnNoSave = True
                End If
            Case conMenu_Save '保存医嘱
                If Not CheckAdvice Then Exit Sub '检查中处理了光标定位
                If Not SaveAdvice Then .SetFocus: Exit Sub
                Unload Me
            Case conMenu_Exit
                Unload Me
        End Select
    End With
End Sub

Private Sub Get一并给药范围(ByVal lng相关ID As Long, lngBegin As Long, lngEnd As Long)
'功能：根据相关的给药途径医嘱ID,确定一并给药的一组药品的起止行号
'说明：中间可能包含有空行
    Dim i As Long
    lngBegin = vsAdvice.FindRow(CStr(lng相关ID), , COL_相关ID)
    For i = lngBegin To vsAdvice.Rows - 1
        If Not vsAdvice.RowHidden(i) And vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                lngEnd = i
            Else
                Exit For
            End If
        End If
    Next
End Sub

Private Sub txt单量_Change()
    txt单量.Tag = "1"
End Sub

Private Sub txt单量_GotFocus()
    zlControl.TxtSelAll txt单量
End Sub

Private Sub txt单量_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt单量.Text) Or txt单量.Text = "" Then
            If SeekNextControl Then Call txt单量_Validate(False)
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt单量_Validate(Cancel As Boolean)
    Dim strMsg As String, dbl次数 As Double, sng天数 As Single
    
    With vsAdvice
        If Val(txt单量.Text) = 0 Then txt单量.Text = ""
        If Not IsNumeric(txt单量.Text) Then
            If txt单量.Text <> "" Then
                Cancel = True: txt单量_GotFocus: Exit Sub
            ElseIf .RowData(.Row) <> 0 And .TextMatrix(.Row, COL_期效) = "长嘱" Then
'                '恢复人为的清除
'                If IsNumeric(.TextMatrix(.Row, COL_单量)) Then
'                    txt单量.Text = .TextMatrix(.Row, COL_单量)
'                End If
            End If
        ElseIf CDbl(txt单量.Text) <= 0 Then
            Cancel = True: txt单量_GotFocus: Exit Sub
        ElseIf CDbl(txt单量.Text) > LONG_MAX Then
            Cancel = True: txt单量_GotFocus: Exit Sub
        Else
            '单量合法性检查
            If txt单量.Text <> "" And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 And Val(.TextMatrix(.Row, COL_收费细目ID)) <> 0 Then
                dbl次数 = IIF(Val(.TextMatrix(.Row, COL_总量)) = 0, 1, Val(.TextMatrix(.Row, COL_总量))) * _
                    Val(.TextMatrix(.Row, COL_包装系数)) * Val(.TextMatrix(.Row, COL_剂量系数)) / Val(txt单量.Text)
                If dbl次数 > 200 Then
                    If MsgBox("该药品按每次 " & FormatEx(txt单量.Text, 5) & .TextMatrix(.Row, COL_单量单位) & " 使用，" & _
                        IIF(Val(.TextMatrix(.Row, COL_总量)) = 0, "每", Val(.TextMatrix(.Row, COL_总量))) & _
                        .TextMatrix(.Row, COL_包装单位) & "可以使用 " & FormatEx(dbl次数, 5) & " 次。" & _
                        vbCrLf & vbCrLf & "你确认单量输入正确吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt单量_GotFocus: Exit Sub
                    End If
                End If
            End If
            
            txt单量.Text = FormatEx(txt单量.Text, 5)
            
            '重新计算药品总量(先输入单量时)
            If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 And .TextMatrix(.Row, COL_期效) = "临嘱" Then
                If .TextMatrix(.Row, COL_频率) <> "" And Val(.TextMatrix(.Row, COL_频率性质)) <> 1 _
                    And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 And Val(.TextMatrix(.Row, COL_包装系数)) <> 0 Then
                    
                    sng天数 = Val(.TextMatrix(.Row, COL_天数))
                    If sng天数 = 0 Then sng天数 = 1
                    
                    txt总量.Text = FormatEx(Calc缺省药品总量( _
                        Val(txt单量.Text), sng天数, _
                        Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), _
                        .TextMatrix(.Row, COL_间隔单位), .TextMatrix(.Row, COL_执行时间), _
                        Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_包装系数)), _
                        Val(.TextMatrix(.Row, COL_可否分零))), 5)
                    txt总量.Tag = "1"
                End If
            End If
        End If
        
        '更新数据
        Call AdviceChange
    End With
End Sub

Private Sub cbo医生嘱托_Change()
    cbo医生嘱托.Tag = "1"
End Sub

Private Sub cbo医生嘱托_Click()
    cbo医生嘱托.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cbo医生嘱托_GotFocus()
    zlControl.TxtSelAll cbo医生嘱托
End Sub

Private Sub cbo医生嘱托_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cbo医生嘱托_Validate(False)
    Else
        Call Cbo.AppendText(cbo医生嘱托, KeyAscii)
    End If
End Sub

Private Sub cbo医生嘱托_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cbo医生嘱托.Text) > 100 Then
        MsgBox "输入内容不过超过 50 个汉字或 100 个字符。", vbInformation, gstrSysName
        cbo医生嘱托_GotFocus
        Cancel = True: Exit Sub
    End If
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub txt医嘱内容_DblClick()
    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
End Sub

Private Sub txt医嘱内容_GotFocus()
    Call zlControl.TxtSelAll(txt医嘱内容)
End Sub

Private Sub txt医嘱内容_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txt医嘱内容)
    End If
End Sub

Private Sub txt医嘱内容_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim byt匹配 As Byte
    If KeyAscii = 13 Then
        
        KeyAscii = 0
        If txt医嘱内容.Text = "" Then Exit Sub
        If txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容) Then
            Call SeekNextControl
            Exit Sub
        End If
        
        If tbrFree.Buttons(1).value = 0 Then
            Set rsTmp = frmClinicSelect.ShowSelect(Me, -1, 0, 0, cbo期效.ListIndex, "", txt医嘱内容.Text, txt医嘱内容, mint范围, , , , , mstr使用科室, byt匹配, mstr诊疗分类, mstr操作类型, mstr执行分类)
            If rsTmp Is Nothing Then '取消或无数据
                '恢复原值
                txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容)
                zlControl.TxtSelAll txt医嘱内容
                txt医嘱内容.SetFocus: Exit Sub
            ElseIf byt匹配 = 1 Then
                Call Cbo.SetIndex(cbo期效.Hwnd, IIF(cbo期效.ListIndex = 0, 1, 0))
            End If
            '新项目的录入
            '成套项目中如果包含成药,则不能按规格下医嘱
            
            '根据选择项目设置缺省医嘱信息
            Me.Refresh
            If AdviceInput(rsTmp, vsAdvice.Row) Then
                '显示已缺省设置的值
                Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
                If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_临床自管药)) = 1 Then
                    cbo执行性质.Tag = "1"
                    Call AdviceChange
                End If
                Call SeekNextControl
            Else
                '恢复原值
                txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容)
                zlControl.TxtSelAll txt医嘱内容
                txt医嘱内容.SetFocus: Exit Sub
        End If
        ElseIf tbrFree.Buttons(1).value = 1 Then
            If txt医嘱内容.Text <> "" Then
                If zlCommFun.ActualLen(txt医嘱内容.Text) > txt医嘱内容.MaxLength Then
                    MsgBox "输入内容不过超过 " & txt医嘱内容.MaxLength \ 2 & " 个汉字或 " & txt医嘱内容.MaxLength & " 个字符。", vbInformation, gstrSysName
                    Call txt医嘱内容_GotFocus: Exit Sub
                End If
                Call AdviceInputFree(vsAdvice.Row)
                Call SeekNextControl
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        If cmdSel.Visible And cmdSel.Enabled Then Call cmdSel_Click
    End If
End Sub

Private Sub cbo执行时间_GotFocus()
    zlControl.TxtSelAll cbo执行时间
End Sub

Private Sub txt医嘱内容_Validate(Cancel As Boolean)
    If tbrFree.Buttons(1).value = 0 Then
        '恢复人为的改变
        If txt医嘱内容.Text <> vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容) Then
            txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容)
        End If
    ElseIf tbrFree.Buttons(1).value = 1 Then
        If vsAdvice.RowData(vsAdvice.Row) <> 0 And txt医嘱内容.Text = "" Then
            '因为必须录入,所以自动恢复
            txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容)
            Exit Sub
        End If
        
        If txt医嘱内容.Text <> "" Then
            If zlCommFun.ActualLen(txt医嘱内容.Text) > txt医嘱内容.MaxLength Then
                MsgBox "输入内容不过超过 " & txt医嘱内容.MaxLength \ 2 & " 个汉字或 " & txt医嘱内容.MaxLength & " 个字符。", vbInformation, gstrSysName
                Call txt医嘱内容_GotFocus: Cancel = True: Exit Sub
            End If
            Call AdviceInputFree(vsAdvice.Row)
        End If
    End If
End Sub

Private Sub txt总量_Change()
    txt总量.Tag = "1"
End Sub

Private Sub txt总量_GotFocus()
    zlControl.TxtSelAll txt总量
End Sub

Private Sub txt总量_KeyPress(KeyAscii As Integer)
    Dim strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt总量.Text) Or txt总量.Text = "" Then
            If SeekNextControl Then Call txt总量_Validate(False)
        End If
    Else
        If RowIn配方行(vsAdvice.Row) Then
            strMask = "0123456789" '中药配方只能输入整数
        Else
            strMask = "0123456789."
        End If
        If InStr(strMask & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt总量_Validate(Cancel As Boolean)
    Dim strMsg As String, sng天数 As Single
    Dim dbl总量 As Double, bln配方行 As Boolean
    
    With vsAdvice
        If Val(txt总量.Text) = 0 Then txt总量.Text = ""
        If Not IsNumeric(txt总量.Text) Then
            If txt总量.Text <> "" Then
                Cancel = True: txt总量_GotFocus: Exit Sub
            ElseIf .RowData(.Row) <> 0 Then
'                '恢复人为的清除
'                If IsNumeric(.TextMatrix(.Row, COL_总量)) Then
'                    txt总量.Text = .TextMatrix(.Row, COL_总量)
'                End If
            End If
        ElseIf CDbl(txt总量.Text) <= 0 Then
            Cancel = True: txt总量_GotFocus: Exit Sub
        ElseIf CDbl(txt总量.Text) > LONG_MAX Then
            Cancel = True: txt总量_GotFocus: Exit Sub
        Else
            txt总量.Text = FormatEx(txt总量.Text, 5)
        End If
        
        bln配方行 = RowIn配方行(.Row)
        
        If IsNumeric(txt总量.Text) Then
            If bln配方行 Then
                txt总量.Text = CInt(txt总量.Text)
            End If
        End If
        
        '检查总量够否
        If txt总量.Text <> "" And InStr(",4,5,6,", .TextMatrix(.Row, COL_类别)) > 0 And .TextMatrix(.Row, COL_期效) = "临嘱" Then
            If .TextMatrix(.Row, COL_频率) <> "" _
                And Val(.TextMatrix(.Row, COL_单量)) <> 0 _
                And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 _
                And Val(.TextMatrix(.Row, COL_包装系数)) <> 0 Then
                
                If Val(.TextMatrix(.Row, COL_频率性质)) = 1 Then
                    dbl总量 = FormatEx(Calc缺省药品总量( _
                        Val(.TextMatrix(.Row, COL_单量)), 1, 1, 1, "天", "", _
                        Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_包装系数)), _
                        Val(.TextMatrix(.Row, COL_可否分零))), 5)
                Else
                    sng天数 = Val(.TextMatrix(.Row, COL_天数))
                    If sng天数 = 0 Then sng天数 = 1
                    
                    dbl总量 = FormatEx(Calc缺省药品总量( _
                        Val(.TextMatrix(.Row, COL_单量)), sng天数, _
                        Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), _
                        .TextMatrix(.Row, COL_间隔单位), .TextMatrix(.Row, COL_执行时间), _
                        Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_包装系数)), _
                        Val(.TextMatrix(.Row, COL_可否分零))), 5)
                End If
                If Val(txt总量.Text) < dbl总量 Then
                    If MsgBox(.TextMatrix(.Row, COL_名称) & "按每次 " & .TextMatrix(.Row, COL_单量) & .TextMatrix(.Row, COL_单量单位) & "," & _
                        .TextMatrix(.Row, COL_频率) & IIF(Val(.TextMatrix(.Row, COL_频率性质)) <> 1 _
                            And Val(.TextMatrix(.Row, COL_天数)) > 0 And .TextMatrix(.Row, COL_类别) <> "4", ",用药 " & sng天数 & " 天", "") & _
                        "执行时,至少需要 " & FormatEx(dbl总量, 5) & .TextMatrix(.Row, COL_总量单位) & ",要继续吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt总量_GotFocus: Exit Sub
                    End If
                End If
            End If
        End If
        
        '更新数据
        Call AdviceChange
    End With
End Sub

Private Sub ClearAdviceCard()
'功能：清除医嘱显示卡片相关的内容
'参数：bln开始时间=是否清除开始时间
    Call SetCardEditable(True)
    
    txt医嘱内容.Text = ""
    cbo医生嘱托.Text = ""
    cbo执行科室.Clear
    cbo附加执行.Clear
    
    cmdExt.Enabled = False
    Call SetDayState(-1, -1)
    Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1)
End Sub

Private Sub SetCardEditable(ByVal Editable As Boolean)
'功能：用颜色标识当前医嘱是否可以编辑
    Dim obj As Object
    
    For Each obj In Controls
        If InStr("Label;TextBox;ComboBox;CheckBox", TypeName(obj)) > 0 Then
            If Not obj.Container Is Nothing Then
                If obj.Container Is fraAdvice Then
                    If Editable Then
                        obj.ForeColor = Me.ForeColor
                    Else
                        obj.ForeColor = &H808080
                    End If
                End If
            End If
        End If
    Next
    fraAdvice.Enabled = Editable
    cmdSel.Enabled = fraAdvice.Enabled
End Sub

Private Sub SetDayState(Optional ByVal intVisible As Integer, Optional ByVal intEnabled As Integer)
'功能：设置执行天数可用和或见状态
'参数：0-保持不变,-1-禁止,1-允许
    If intEnabled = -1 Then
        txt天数.Enabled = False
        txt天数.BackColor = Me.BackColor
        txt天数.Text = ""
    ElseIf intEnabled = 1 Then
        txt天数.TabStop = True
        txt天数.Enabled = True
        txt天数.BackColor = vsAdvice.BackColor
    End If
    
    If intVisible = -1 Then
        lbl天数.Visible = False
        txt天数.Visible = False
        txt天数.Text = ""
        
        lbl总量.Left = lbl用法.Left + lbl用法.Width - lbl总量.Width
        txt总量.Left = txt用法.Left
        txt总量.Width = txt用法.Width - cmd用法.Width - 15
        lbl总量单位.Left = txt总量.Left + txt总量.Width + 30
        
        lbl单量.Left = lbl频率.Left + lbl频率.Width - lbl单量.Width
        txt单量.Left = txt频率.Left
        txt单量.Width = txt频率.Width - cmd频率.Width - 15
        lbl单量单位.Left = txt单量.Left + txt单量.Width + 30
        
        txt总量.TabIndex = cmd频率.TabIndex + 1
        txt天数.TabIndex = txt总量.TabIndex + 1
        txt单量.TabIndex = txt天数.TabIndex + 1
    ElseIf intVisible = 1 Then
        lbl天数.Visible = True
        txt天数.Visible = True
        
        lbl单量.Left = lbl用法.Left + lbl用法.Width - lbl单量.Width
        txt单量.Left = txt用法.Left
        txt单量.Width = txt用法.Width - txt天数.Width - Me.TextWidth("三个字!") - 15
        lbl单量单位.Left = txt单量.Left + txt单量.Width + 30
        
        lbl总量.Left = lbl频率.Left + lbl频率.Width - lbl总量.Width
        txt总量.Left = txt频率.Left
        txt总量.Width = txt频率.Width - cmd频率.Width - 15
        lbl总量单位.Left = txt总量.Left + txt总量.Width + 30
        
        txt单量.TabIndex = cmd频率.TabIndex + 1
        txt天数.TabIndex = txt单量.TabIndex + 1
        txt总量.TabIndex = txt天数.TabIndex + 1
    End If
End Sub

Private Function Get频率范围(ByVal lngRow As Long) As Integer
    Dim lngFind As Long
    
    With vsAdvice
        If RowIn配方行(lngRow) Then
            Get频率范围 = 2 '中医
        Else
            If RowIn检验行(lngRow) Then '以检验项目行为准
                lngFind = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
                If lngFind <> -1 Then lngRow = lngFind
            End If
            If Val(.TextMatrix(lngRow, COL_频率性质)) = 0 Then
                Get频率范围 = 1 '可选频率的项目使用西医频率项目
            ElseIf Val(.TextMatrix(lngRow, COL_频率性质)) = 1 Then
                Get频率范围 = -1 '一次性
            ElseIf Val(.TextMatrix(lngRow, COL_频率性质)) = 2 Then
                Get频率范围 = -2 '持续性
            End If
        End If
    End With
End Function

Private Function SeekVisibleRow() As Boolean
'功能：当前行为隐藏行时，定位到它所属的可见行
    Dim lngRow As Long
    
    With vsAdvice
        If Not .RowHidden(.Row) Then Exit Function
        If InStr(",F,G,C,D,E,", .TextMatrix(.Row, COL_类别)) > 0 And Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_相关ID))))
        ElseIf .TextMatrix(.Row, COL_类别) = "7" Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_相关ID))))
        ElseIf .TextMatrix(.Row, COL_类别) = "E" And Val(.TextMatrix(.Row, COL_相关ID)) = 0 Then
            lngRow = .Row - 1
        End If
        If lngRow <> -1 Then
            If .RowData(lngRow) <> 0 Then
                .Row = lngRow: SeekVisibleRow = True
            End If
        End If
    End With
End Function

Private Sub SetCbo执行性质(ByVal bln临嘱 As Boolean, ByVal bln含自备药 As Boolean, ByVal bln临床自管药 As Boolean)
    cbo执行性质.Clear
    
    If bln临床自管药 Then
        cbo执行性质.AddItem "1-自备药"
    Else
        If bln临嘱 Then
            cbo执行性质.AddItem "0-正常"
            If bln含自备药 Then cbo执行性质.AddItem "1-自备药"
            cbo执行性质.AddItem "2-离院带药"
            cbo执行性质.AddItem "3-自取药"
            cbo执行性质.AddItem "4-不取药"
        Else
            cbo执行性质.AddItem "0-正常"
            If bln含自备药 Then cbo执行性质.AddItem "1-自备药"
            cbo执行性质.AddItem "4-不取药"
        End If
    End If
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'设置一组医嘱行的"缺省"列状态

    If Col = col_缺省 Or Col = col_备选 Then
        Dim i As Long, lng组ID As Long, lngThis组ID As Long
        Dim lngBegin As Long, lngEnd As Long
        
        With vsAdvice
            
            '一并给药的一起设置，其他情况找出起始行
            If Not RowIn一并给药(Row) Then
                Call GetRowScope(Row, lngBegin, lngEnd)
            Else
                Call Get一并给药范围(Val(.TextMatrix(Row, COL_相关ID)), lngBegin, lngEnd)
            End If
            
            For i = lngBegin To lngEnd
                If i <> Row Then
                    .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                End If
                If Col = col_备选 And .TextMatrix(Row, Col) = -1 And mbln显示缺省列 Then
                    .TextMatrix(i, col_缺省) = 0
                End If
                If Col = col_缺省 And .TextMatrix(Row, Col) = -1 And mbln显示缺省列 Then
                    .TextMatrix(i, col_备选) = 0
                End If
            Next
            mblnNoSave = True
        End With
    End If
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'功能：当行改变时，更新卡片内容
    Dim rsItem As New ADODB.Recordset
    Dim strSql As String, lngRow As Long
    Dim lng用法ID As Long, blnEditable As Boolean
    Dim lng药品ID As Long, lngBaseRow As Long    '中药配方的第一味组成药行
    Dim dblPrice As Double, strTmp As String, i As Long

    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, col_医嘱内容)
    End If

    If NewRow = OldRow Then Exit Sub
    If Not mblnRowChange Then Exit Sub
    If SeekVisibleRow Then Exit Sub

    lngRow = NewRow

    '当前行是空行时，如果前一行是一并给药行，则缺省按下“一并”按钮
    If vsAdvice.RowData(lngRow) = 0 Then
        i = GetPreRow(lngRow)
        If i = -1 Then
            mblnRowMerge = False
        Else
            mblnRowMerge = RowIn一并给药(i)
        End If
    Else
        mblnRowMerge = RowIn一并给药(lngRow)
    End If
    cbsMain.RecalcLayout    '*即时刷新

    Me.Refresh
    zlControl.FormLock Me.Hwnd

    On Error GoTo errH
    chkMedicineVariety.Visible = True

    With vsAdvice
        If .RowData(lngRow) = 0 Then
            '无效行清除卡片内容
            Call ClearAdviceCard

            '缺省为非自由录入
            tbrFree.Buttons(1).value = 0
            tbrFree.Buttons(1).Enabled = Not RowIn一并给药(lngRow)
            tbrFree.Buttons(1).Image = IIF(tbrFree.Buttons(1).Enabled, 1, 2)

            '缺省期效根据上一行的显示
            i = GetPreRow(lngRow)
            If i = -1 Or Not Visible Then
                Call Cbo.SetIndex(cbo期效.Hwnd, 1)    '缺省为临嘱
            Else
                Call Cbo.SetIndex(cbo期效.Hwnd, IIF(.TextMatrix(i, COL_期效) = "长嘱", 0, 1))
            End If
        ElseIf Val(.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
            '自由录入医嘱
            blnEditable = Not mblnView
            Call SetCardEditable(blnEditable)

            tbrFree.Buttons(1).value = 1
            tbrFree.Buttons(1).Enabled = blnEditable
            tbrFree.Buttons(1).Image = IIF(blnEditable, 1, 2)
            cmdExt.Enabled = False
            cmdSel.Enabled = False
            chkMedicineVariety.Visible = False

            '其它输入项禁用
            Call SetDayState(-1, -1)
            SetItemEditable -1, -1, -1, -1, -1, , -1, -1, -1

            '显示当前医嘱卡片内容
            '--------------------------------------------------------------------------------------------
            Call Cbo.SetIndex(cbo期效.Hwnd, IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1))

            '医嘱内容
            txt医嘱内容.Text = .TextMatrix(lngRow, col_医嘱内容)

            '医生嘱托
            cbo医生嘱托.Text = .TextMatrix(lngRow, COL_医生嘱托)

            '可选执行科室
            SetItemEditable , , , , , 1
            Call Get成套执行科室(cbo执行科室, "*", 0, 0, 4, Val(.TextMatrix(lngRow, COL_执行科室ID)), cbo期效.ListIndex, mint范围)
        Else
            '卡片编辑：已校对的医嘱不能修改,补录医嘱时不能更改非补录的内容
            blnEditable = Not mblnView
            Call SetCardEditable(blnEditable)

            '已有诊疗项目，不可变为自由录入
            tbrFree.Buttons(1).value = 0
            tbrFree.Buttons(1).Enabled = False
            tbrFree.Buttons(1).Image = 2


            '获取诊疗项目基本信息
            '---------------------
            chkMedicineVariety.Tag = "不清除"
            If InStr("4,5,6", Val(.TextMatrix(lngRow, COL_类别))) > 0 Then
                lng药品ID = Val(.TextMatrix(lngRow, COL_收费细目ID))
                chkMedicineVariety.value = IIF(lng药品ID = 0, 1, 0)
            Else
                chkMedicineVariety.Visible = False
            End If
            chkMedicineVariety.Tag = ""

            If RowIn配方行(lngRow) Then
                txt总量.MaxLength = 3
                '获取中药配方第一味中药行
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
                lng药品ID = Val(.TextMatrix(lngBaseRow, COL_收费细目ID))
            ElseIf RowIn检验行(lngRow) Then
                '获取一并采样的第一个项目行
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
                txt总量.MaxLength = txt单量.MaxLength
            Else
                lngBaseRow = lngRow
                txt总量.MaxLength = txt单量.MaxLength
            End If
            Set rsItem = Get诊疗项目记录(Val(.TextMatrix(lngBaseRow, COL_诊疗项目ID)))

            '扩展按钮可用状态(检查组合,检验组合,手术,中药配方)
            cmdExt.Enabled = InStr(",7,C,F,D,", rsItem!类别) > 0

            '显示当前医嘱卡片内容
            '--------------------------------------------------------------------------------------------
            Call Cbo.SetIndex(cbo期效.Hwnd, IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1))
            '医嘱内容
            txt医嘱内容.Text = .TextMatrix(lngRow, col_医嘱内容)

            '单量
            '----------------------
            If rsItem!类别 = "7" Then    '中药配方(中草药)虽然有单量,但不在这里填写
                SetItemEditable -1
            ElseIf cbo期效.ListIndex = 0 Then
                '长嘱：成药或计时,计量项目可以录入
                If InStr(",1,2,", NVL(rsItem!计算方式, 0)) > 0 Or InStr(",5,6,", rsItem!类别) > 0 Then
                    SetItemEditable 1
                    txt单量.Text = .TextMatrix(lngRow, COL_单量)
                    lbl单量单位.Caption = .TextMatrix(lngRow, COL_单量单位)
                Else
                    SetItemEditable -1
                End If
            ElseIf cbo期效.ListIndex = 1 Then
                '临嘱:成药或可选择频率的计时,计量项目可以录入(注意这是原始频率,当前可能已是一次性)
                If (NVL(rsItem!执行频率, 0) = 0 And InStr(",1,2,", NVL(rsItem!计算方式, 0)) > 0) _
                   Or InStr(",5,6,", rsItem!类别) > 0 Then
                    SetItemEditable 1
                    txt单量.Text = .TextMatrix(lngRow, COL_单量)
                    lbl单量单位.Caption = .TextMatrix(lngRow, COL_单量单位)
                Else
                    SetItemEditable -1
                End If
            End If

            '天数：西药，中成药临嘱才使用，用于计算总量
            '一般：临嘱的药品(非中药)或可选择频率的计时,计量项目可以使用天数来自动计算总量
            blnEditable = False
            If cbo期效.ListIndex = 1 And InStr(",5,6,", rsItem!类别) > 0 Then
                If Val(.TextMatrix(lngRow, COL_频率性质)) <> 1 Then blnEditable = True
            End If
            If blnEditable Then
                SetDayState 1, 1
            Else
                SetDayState -1, -1
            End If
            txt天数.Text = Val(.TextMatrix(lngRow, COL_天数))
            If Val(txt天数.Text) = 0 Then txt天数.Text = ""

            '总量
            '--------------------
            If rsItem!类别 = "7" Then
                '中药配方(中草药)填写为付数
                If cbo期效.ListIndex = 1 Then
                    SetItemEditable , 1
                Else
                    SetItemEditable , -1    '配方长嘱不能输入总量，但兼容已输入的数据(新的为总量作总单量，固定为1付不用输入)
                End If
                lbl总量单位.Caption = "付"
                txt总量.Text = .TextMatrix(lngRow, COL_总量)    '付数

            ElseIf cbo期效.ListIndex = 1 Then
                '临嘱都需要填写总量:临嘱发送以总量为准
                If rsItem!类别 = "Z" And NVL(rsItem!操作类型) <> "0" Then
                    SetItemEditable , -1    '特殊医嘱不允许修改总量(固定为1次)
                ElseIf InStr(",5,6,", rsItem!类别) = 0 And NVL(rsItem!计算方式, 0) = 3 _
                       And (NVL(rsItem!执行频率, 0) = 1 Or Val(.TextMatrix(lngRow, COL_频率性质)) = 1) Then
                    SetItemEditable , -1    '非药品一次性计次项目不输入总量(原始频率为一次性或当前设置为一次性)
                Else
                    SetItemEditable , 1
                End If
                lbl总量单位.Caption = .TextMatrix(lngRow, COL_总量单位)
                txt总量.Text = .TextMatrix(lngRow, COL_总量)
            Else
                '其它长嘱不允许填写总量
                SetItemEditable , -1
            End If

            '给药途径和中药用法
            '--------------
            If InStr(",5,6,", rsItem!类别) > 0 Then
                SetItemEditable , , 1
                lbl用法.Caption = "给药途径"
                '查找给药途径对应的行:查找的Rowdata(Variant)数据要转为Long型,才能精确匹配
                lng用法ID = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                lng用法ID = Val(.TextMatrix(lng用法ID, COL_诊疗项目ID))
                cmd用法.Tag = lng用法ID
                txt用法.Text = sys.RowValue("诊疗项目目录", lng用法ID, "名称")
            ElseIf rsItem!类别 = "K" Then
                '输血医嘱：要兼容以前没有输血途径的情况
                lng用法ID = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
                If lng用法ID <> -1 Then
                    SetItemEditable , , 1
                    If Val(.TextMatrix(lngRow, COL_检查方法)) = 0 And gbln血库系统 = True Then
                        lbl用法.Caption = "采集方法"
                    Else
                        lbl用法.Caption = "输血途径"
                    End If
                    lng用法ID = Val(.TextMatrix(lng用法ID, COL_诊疗项目ID))
                    cmd用法.Tag = lng用法ID
                    txt用法.Text = sys.RowValue("诊疗项目目录", lng用法ID, "名称")
                Else
                    SetItemEditable , , -1
                End If
            ElseIf rsItem!类别 = "7" Then
                SetItemEditable , , 1
                lbl用法.Caption = "中药用法"

                '中药配方显示行就是中药用法行
                lng用法ID = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
                cmd用法.Tag = lng用法ID
                txt适用证候.Text = .TextMatrix(lngRow, COL_适用证候)
                txt用法.Text = sys.RowValue("诊疗项目目录", lng用法ID, "名称")
            ElseIf RowIn检验行(lngRow) Then    '不用类别判断,兼容以前的检验
                '检验组合
                SetItemEditable , , 1
                lbl用法.Caption = "采集方法"

                '检验组合显示行就是采集方法行
                lng用法ID = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
                cmd用法.Tag = lng用法ID
                txt用法.Text = sys.RowValue("诊疗项目目录", lng用法ID, "名称")
            Else
                SetItemEditable , , -1
            End If

            If rsItem!类别 = "7" And mbyt场合 = 1 Then
                SetItemEditable , , , , , , , , 1
            Else
                SetItemEditable , , , , , , , , -1
            End If
            
            '滴速：输液类给药途径的药品可以输入
            If InStr(",5,6,", rsItem!类别) > 0 And mbyt场合 <> 2 Then
                i = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                If Val(.TextMatrix(i, COL_执行分类)) = 1 Then
                    SetItemEditable , , , , , , , , , 1
                    If InStr(.TextMatrix(i, COL_医生嘱托), "滴/分钟") > 0 Then
                        lbl滴速单位.Caption = "滴/分钟"
                    ElseIf InStr(.TextMatrix(i, COL_医生嘱托), "毫升/小时") > 0 Then
                        lbl滴速单位.Caption = "毫升/小时"
                    End If
                    Call Load输液滴速(cbo滴速, lbl滴速单位, False)
                    cbo滴速.Text = Replace(.TextMatrix(i, COL_医生嘱托), lbl滴速单位.Caption, "")
                Else
                    SetItemEditable , , , , , , , , , -1
                End If
            Else
                SetItemEditable , , , , , , , , , -1
            End If
       
            
            If mbyt场合 <> 2 Then
                '频率：都可以选择(临嘱输入用于指导使用)
                If True Then
                    SetItemEditable , , , 1
                    cmd频率.Tag = .TextMatrix(lngRow, COL_频率)
                    txt频率.Text = .TextMatrix(lngRow, COL_频率)
                Else
                    SetItemEditable , , , -1
                End If
    
                '执行时间："可选频率"或药品(当前未被设置为一次性)。非"分钟"间隔执行的
                If NVL(rsItem!执行频率, 0) = 0 And Val(.TextMatrix(lngBaseRow, COL_频率性质)) <> 1 And .TextMatrix(lngRow, COL_间隔单位) <> "分钟" Then
                    SetItemEditable , , , , 1
                    Call Get时间方案(cbo执行时间, Get频率范围(lngRow), .TextMatrix(lngRow, COL_频率), lng用法ID)
                    cbo执行时间.Text = .TextMatrix(lngRow, COL_执行时间)
                Else
                    SetItemEditable , , , , -1
                End If
    
                '医生嘱托
                cbo医生嘱托.Text = .TextMatrix(lngRow, COL_医生嘱托)
    
                '执行性质:长嘱目前可以使用"自备药"
                If InStr(",5,6,7,", rsItem!类别) > 0 Then
                    '如果是自管药则固定选择自备药
                    If Val(.TextMatrix(lngRow, COL_临床自管药)) = 1 Then
                        strTmp = "自备药"
                    Else
                        If rsItem!类别 = "7" Then
                            '对于中药配方,根据诊疗项目管理中限制及本程序处理,不可能用法和煎法一个为院外执行,一个不为
                            If Val(.TextMatrix(lngBaseRow, COL_执行性质)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质)) <> 5 Then
    
                                strTmp = IIF(Val(.TextMatrix(lngBaseRow, COL_执行标记)) = 2, "不取药", "自备药")
    
                            ElseIf Val(.TextMatrix(lngBaseRow, COL_执行性质)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                                strTmp = "离院带药"
                            Else
                                strTmp = IIF(Val(.TextMatrix(lngBaseRow, COL_执行标记)) = 0, "正常", "自取药")
                            End If
                        Else
                            i = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                            If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                If Val(.TextMatrix(lngRow, COL_执行标记)) = 2 Then
                                    strTmp = "不取药"
                                Else
                                    strTmp = "自备药"
                                End If
                            ElseIf Val(.TextMatrix(lngRow, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                strTmp = "离院带药"
                            Else
                                strTmp = IIF(Val(.TextMatrix(lngRow, COL_执行标记)) = 0, "正常", "自取药")
                            End If
                        End If
                    End If
    
                    Call SetCbo执行性质(cbo期效.ListIndex = 1, gbln抗菌药物使用自备药 Or Not gblnKSSStrict Or Val(.TextMatrix(lngRow, COL_抗菌等级)) = 0, Val(.TextMatrix(lngRow, COL_临床自管药)) = 1)
                    SetItemEditable , , , , , , 1
                    Call Cbo.SetIndex(cbo执行性质.Hwnd, Cbo.FindIndex(cbo执行性质, strTmp, True))
                Else
                    SetItemEditable , , , , , , -1
                End If
    
                lbl执行科室.Caption = "执行科室"
                '执行科室
                If rsItem!类别 = "Z" And NVL(rsItem!操作类型, 0) = 3 Then
                    '转科医嘱用临床科室
                    SetItemEditable , , , , , 1
                    lbl执行科室.Caption = "转入科室"
                    Call Get临床科室(mint范围, 0, Val(.TextMatrix(lngRow, COL_执行科室ID)), cbo执行科室, True)
                ElseIf rsItem!类别 = "Z" And NVL(rsItem!操作类型, 0) = 7 Then
                    '会诊医嘱用临床科室
                    SetItemEditable , , , , , 1
                    lbl执行科室.Caption = "会诊科室"
                    Call Get临床科室(mint范围, 0, Val(.TextMatrix(lngRow, COL_执行科室ID)), cbo执行科室)
                Else
                    '是药品则以药品行为准显示,检验组合以检验项目为准显示
                    i = lngRow
                    If rsItem!类别 = "7" Then
                        i = lngBaseRow
                    ElseIf RowIn检验行(lngRow) Then    '不用类别判断,兼容以前的检验
                        i = lngBaseRow
                    End If
    
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                        '非叮嘱和院外执行时才显示和可以选择(包括药品)
                        SetItemEditable , , , , , 1
                        Call Get成套执行科室(cbo执行科室, rsItem!类别, rsItem!ID, lng药品ID, NVL(rsItem!执行科室, 0), Val(.TextMatrix(i, COL_执行科室ID)), cbo期效.ListIndex, mint范围)
    
                        '非散装形态，只允许在配方界面选药房
                        If rsItem!类别 = "7" Then
                            If Val(.TextMatrix(lngRow, COL_中药形态)) <> 0 Then
                                cbo执行科室.Enabled = False
                                cbo执行科室.BackColor = Me.BackColor
                            End If
                        End If
    
                    ElseIf InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                        SetItemEditable , , , , , -1
                        If Val(.TextMatrix(i, COL_执行性质)) = 0 Then
                            cbo执行科室.AddItem "<无执行叮嘱>"
                        Else
                            cbo执行科室.AddItem "-"
                        End If
                        Call Cbo.SetIndex(cbo执行科室.Hwnd, 0)
                    End If
                    If InStr("5,6,7", rsItem!类别) > 0 Then lbl执行科室.Caption = "发药药房"
                End If
    
                '附加执行:指给药途径,中药用法,手术麻醉,采集方式的执行科室
                If Should附加执行(lngRow, i, strTmp) Then
                    SetItemEditable , , , , , , , 1
                    Call Get成套执行科室(cbo附加执行, .TextMatrix(i, COL_类别), Val(.TextMatrix(i, COL_诊疗项目ID)), lng药品ID, Val(.TextMatrix(i, COL_执行性质)), Val(.TextMatrix(i, COL_执行科室ID)), cbo期效.ListIndex, mint范围)
                Else
                    SetItemEditable , , , , , , , -1
                    If i <> -1 Then
                        If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                            If Val(.TextMatrix(i, COL_执行性质)) = 0 Then
                                cbo附加执行.AddItem "<无执行叮嘱>"
                            ElseIf Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                cbo附加执行.AddItem "-"
                            End If
                            Call Cbo.SetIndex(cbo附加执行.Hwnd, 0)
                        End If
                    End If
                End If
                lbl附加执行.Caption = strTmp
            Else
                SetItemEditable , , , 1, -1, -1, -1, -1, -1
            End If
        End If
    End With

    '清除编辑标志
    Call ClearItemTag

    cbsMain.RecalcLayout    '即时刷新,有Lock可不要
    zlControl.FormLock 0
    Exit Sub
errH:
    zlControl.FormLock 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Should附加执行(ByVal lngRow As Long, lngRow2 As Long, str执行科室 As String) As Boolean
'功能：判断指定的医嘱行(可见行)是否可以设置附加的执行科室
'参数：lngRow2=返回附加行的医嘱行号
'      str执行科室=附加执行科室类型
    Dim i As Long
    
    lngRow2 = -1
    str执行科室 = "附加执行"
    With vsAdvice
        If lngRow = 0 Or .RowData(lngRow) = 0 Then Exit Function

        If RowIn配方行(lngRow) Then
            '中药用法
            lngRow2 = lngRow
            str执行科室 = "用法执行"
            Should附加执行 = True
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
            '给药途径
            lngRow2 = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
            str执行科室 = "给药执行"
            Should附加执行 = True
        ElseIf .TextMatrix(lngRow, COL_类别) = "F" Then
            '手术麻醉
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "G" Then
                        lngRow2 = i: Exit For
                    End If
                Else
                    Exit For
                End If
            Next
            str执行科室 = "麻醉执行"
            If lngRow2 <> -1 Then Should附加执行 = True
        ElseIf .TextMatrix(lngRow, COL_类别) = "K" Then
            '输血途径
            If Val(.TextMatrix(lngRow, COL_检查方法)) = 0 And gbln血库系统 = True Then
                str执行科室 = "采集执行"
            Else
                str执行科室 = "输血执行"
            End If
            lngRow2 = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
            If lngRow2 <> -1 Then Should附加执行 = True
        ElseIf .TextMatrix(lngRow, COL_类别) = "E" _
            And .TextMatrix(lngRow - 1, COL_类别) = "C" _
            And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
            '采集方式
            lngRow2 = lngRow
            str执行科室 = "采集执行"
            Should附加执行 = True
        End If
        
        '叮嘱或院外执行
        If Should附加执行 Then
            If InStr(",0,5,", Val(.TextMatrix(lngRow2, COL_执行性质))) > 0 Then
                Should附加执行 = False
            End If
        End If
    End With
End Function


Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(0, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
        
        If Col = col_医嘱内容 Then Call vsAdvice.AutoSize(col_医嘱内容)
    End If
End Sub

Private Sub vsAdvice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> col_缺省 And Col <> col_备选)
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        End If
    End If
End Sub

Private Function RowIsLastVisible(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否最后一可见行
    Dim i As Long
    
    With vsAdvice
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then Exit For
        Next
        If i >= .FixedRows Then
            RowIsLastVisible = lngRow = i
        End If
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '擦除固定列中的表格线
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)

            '仅左边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅上边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅下边表格线
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If RowIsLastVisible(Row) Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅右边表格线
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            lngLeft = COL_期效: lngRight = COL_期效
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_天数: lngRight = COL_用法
                If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            End If
            
            If Not RowIn一并给药(Row) Then Exit Sub
            If .RowData(Row) = 0 Then
                Call Get一并给药范围(Val(.TextMatrix(Row - 1, COL_相关ID)), lngBegin, lngEnd)
            Else
                Call Get一并给药范围(Val(.TextMatrix(Row, COL_相关ID)), lngBegin, lngEnd)
            End If
            
            vRect.Left = Left '擦除左边表格线
            vRect.Right = Right - 1 '保留右边表格线
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '底行保留下边线
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        '执行Execute时,如果按钮"当前"不可用则不实际执行；不管是否可见
        cbsMain.FindControl(, conMenu_Delete, True, True).Execute
    End If
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim objEdit As Object
    
    If KeyAscii = 13 Then
        '定位到对应的编辑控件
        KeyAscii = 0
        Select Case vsAdvice.Col
            Case COL_期效
                Set objEdit = cbo期效
            Case col_医嘱内容
                Set objEdit = txt医嘱内容
            Case COL_单量
                Set objEdit = txt单量
            Case COL_总量
                Set objEdit = txt总量
            Case COL_用法
                Set objEdit = txt用法
            Case COL_频率
                Set objEdit = txt频率
            Case COL_执行时间
                Set objEdit = cbo执行时间
            Case COL_执行科室ID
                Set objEdit = cbo执行科室
            Case COL_医生嘱托
                Set objEdit = cbo医生嘱托
        End Select
        If Not objEdit Is Nothing Then
            If objEdit.Enabled And objEdit.Visible Then objEdit.SetFocus
        End If
    End If
End Sub

Private Sub ClearItemTag()
'功能：清除控件编辑标志
    txt天数.Tag = ""
    txt单量.Tag = ""
    txt总量.Tag = ""
    txt用法.Tag = ""
    txt频率.Tag = ""
    cbo执行时间.Tag = ""
    cbo医生嘱托.Tag = ""
    cbo执行科室.Tag = ""
    cbo执行性质.Tag = ""
    cbo附加执行.Tag = ""
    txt适用证候.Tag = ""
    cbo滴速.Tag = ""
End Sub

Private Sub SetItemEditable(Optional int单量 As Integer, Optional int总量 As Integer, _
    Optional int用法 As Integer, Optional int频率 As Integer, _
    Optional int执行时间 As Integer, Optional int执行科室 As Integer, _
    Optional int执行性质 As Integer, Optional int附加执行 As Integer, _
    Optional int适用证候 As Integer, Optional int滴速 As Integer)
'功能：设置指定编辑项的可用状态
'参数：0-保持不变,-1-禁止,1-允许,2-锁定
'说明：禁止时,同时清除该项目数据(不是全部)

    '依次设置为禁止时,会引发焦点改变,从而可能引发Validate事件,所以先禁止焦点顺序
    If int单量 = -1 Then txt单量.TabStop = False
    If int总量 = -1 Then txt总量.TabStop = False
    If int用法 = -1 Then txt用法.TabStop = False
    If int频率 = -1 Then txt频率.TabStop = False
    If int执行时间 = -1 Then cbo执行时间.TabStop = False
    If int执行科室 = -1 Then cbo执行科室.TabStop = False
    If int执行性质 = -1 Then cbo执行性质.TabStop = False
    If int附加执行 = -1 Then cbo附加执行.TabStop = False
    If int适用证候 = -1 Then txt适用证候.TabStop = False
    
    If int单量 = -1 Then
        txt单量.Enabled = False
        txt单量.BackColor = Me.BackColor
        txt单量.Text = ""
        lbl单量单位.Caption = "" '"单位"
    ElseIf int单量 = 1 Then
        txt单量.TabStop = True
        txt单量.Enabled = True
        txt单量.BackColor = vsAdvice.BackColor
    End If

    If int总量 = -1 Then
        txt总量.Enabled = False
        txt总量.BackColor = Me.BackColor
        txt总量.Text = ""
        lbl总量单位.Caption = "" '"单位"
    ElseIf int总量 = 1 Then
        txt总量.TabStop = True
        txt总量.Enabled = True
        txt总量.BackColor = vsAdvice.BackColor
    End If
    
    If int用法 = -1 Then
        txt用法.Enabled = False
        txt用法.BackColor = Me.BackColor
        txt用法.Text = ""
        cmd用法.Enabled = False
        lbl用法.Caption = "用法"
    ElseIf int用法 = 1 Then
        txt用法.TabStop = True
        txt用法.Enabled = True
        cmd用法.Enabled = True
        txt用法.BackColor = vsAdvice.BackColor
    End If

    If int频率 = -1 Then
        txt频率.Enabled = False
        cmd频率.Enabled = False
        txt频率.BackColor = Me.BackColor
        txt频率.Text = ""
    ElseIf int频率 = 1 Then
        txt频率.TabStop = True
        txt频率.Enabled = True
        cmd频率.Enabled = True
        txt频率.BackColor = vsAdvice.BackColor
    End If

    If int执行时间 = -1 Then
        cbo执行时间.Text = ""
        cbo执行时间.Enabled = False
        cbo执行时间.BackColor = Me.BackColor
        cbo执行时间.Clear
    ElseIf int执行时间 = 1 Then
        cbo执行时间.TabStop = True
        cbo执行时间.Enabled = True
        cbo执行时间.BackColor = vsAdvice.BackColor
    End If

    If int执行科室 = -1 Then
        lbl执行科室.Caption = "执行科室"
        cbo执行科室.Enabled = False
        cbo执行科室.BackColor = Me.BackColor
        cbo执行科室.Clear
    ElseIf int执行科室 = 1 Then
        lbl执行科室.Caption = "执行科室"
        cbo执行科室.TabStop = True
        cbo执行科室.Enabled = True
        cbo执行科室.BackColor = vsAdvice.BackColor
    End If

    If int执行性质 = -1 Then
        cbo执行性质.Enabled = False
        cbo执行性质.BackColor = Me.BackColor
        Call Cbo.SetIndex(cbo执行性质.Hwnd, -1) '不清除
    ElseIf int执行性质 = 1 Then
        cbo执行性质.TabStop = True
        cbo执行性质.Enabled = True
        cbo执行性质.BackColor = vsAdvice.BackColor
    End If
    
    If int附加执行 = -1 Then
        lbl附加执行.Caption = "附加执行"
        cbo附加执行.Enabled = False
        cbo附加执行.BackColor = Me.BackColor
        cbo附加执行.Clear
    ElseIf int附加执行 = 1 Then
        lbl附加执行.Caption = "附加执行"
        cbo附加执行.TabStop = True
        cbo附加执行.Enabled = True
        cbo附加执行.BackColor = vsAdvice.BackColor
    End If
    
    If int适用证候 = -1 Then
        lbl适用证候.Visible = False
        txt适用证候.Visible = False
        cmd适用证候.Visible = False
    ElseIf int适用证候 = 1 Then
        lbl适用证候.Visible = True
        txt适用证候.Visible = True
        cmd适用证候.Visible = True
        txt适用证候.TabStop = True
    End If
    
    If int滴速 = -1 Then
        cbo滴速.Text = ""
        lbl滴速.Visible = False
        cbo滴速.Visible = False
        lbl滴速单位.Visible = False
    ElseIf int滴速 = 1 Then
        lbl滴速.Visible = True
        cbo滴速.Visible = True
        lbl滴速单位.Visible = True
    End If
End Sub

Private Function GetPreRow(ByVal lngRow As Long) As Long
'功能：取上一最近有效可见行
'返回：无有效行时,返回-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow - 1 To vsAdvice.FixedRows Step -1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        
        End If
    Next
    GetPreRow = lngTmp
End Function

Private Function GetNextRow(ByVal lngRow As Long) As Long
'功能：取下一最近有效可见行
'返回：无有效行时,返回-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        End If
    Next
    GetNextRow = lngTmp
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
'功能：获取组ID相同的一组医嘱行号范围(注意考虑一并给药中的空行)
    Dim lngS组ID As Long, lngO组ID As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS组ID = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) = 0, .RowData(lngRow), Val(.TextMatrix(lngRow, COL_相关ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_相关ID)))
            If Not (.RowData(i) = 0 And i >= .FixedRows) Then '跳过空行
                If lngO组ID = lngS组ID Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_相关ID)))
            If Not (.RowData(i) = 0 And i >= .FixedRows) Then '跳过空行
                If lngO组ID = lngS组ID Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
End Sub

Private Function GetNextID() As Long
'功能：模拟获取下一个ID
    mlngNextID = mlngNextID + 1
    GetNextID = mlngNextID
End Function

Private Function GetCurRow序号(lngRow As Long) As Long
'功能：获取指定行可用的的序号
'参数：lngRow=要取序号的行
    Dim lng序号 As Long, i As Long
    Dim lng序号1 As Long, lng序号2 As Long
            
    '取之后最近一个有效序号,直接使用
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If IsNumeric(vsAdvice.TextMatrix(i, COL_序号)) Then
                lng序号 = Val(vsAdvice.TextMatrix(i, COL_序号))
                Exit For
            End If
        End If
    Next
    If lng序号 = 0 Then
        '后面没有,则取之前的最大序号+1
        For i = lngRow - 1 To vsAdvice.FixedRows Step -1
            If vsAdvice.RowData(i) <> 0 Then
                If IsNumeric(vsAdvice.TextMatrix(i, COL_序号)) Then
                    lng序号 = Val(vsAdvice.TextMatrix(i, COL_序号))
                    Exit For
                End If
            End If
        Next
        If lng序号 <> 0 Then lng序号 = lng序号 + 1
    End If
    If lng序号 = 0 Then lng序号 = 1
    GetCurRow序号 = lng序号
End Function

Private Sub AdviceSet医嘱序号(lngRow As Long, intStep As Integer)
'功能：将当前病人医嘱记录中序号前移或后移
'参数：lngRow=起始调整行,intStep=调整步长,如1或-1
    Dim i As Long
    
    For i = lngRow To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If IsNumeric(vsAdvice.TextMatrix(i, COL_序号)) Then
                vsAdvice.TextMatrix(i, COL_序号) = Val(vsAdvice.TextMatrix(i, COL_序号)) + intStep
            End If
        End If
    Next
End Sub

Private Sub AdviceDelete(ByVal lngRow As Long)
'功能：指定的医嘱删除处理
    Dim lngBegin As Long, lngEnd As Long
    Dim lng相关ID As Long, blnGroup As Boolean
    Dim lng医嘱ID As Long, i As Integer
    
    mblnRowChange = False
    vsAdvice.Redraw = flexRDNone
    
    If vsAdvice.RowData(lngRow) <> 0 Then
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_类别)) > 0 Then
            lng医嘱ID = vsAdvice.RowData(lngRow)
            lng相关ID = Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
            blnGroup = RowIn一并给药(lngRow)
            If blnGroup Then
                '先删除一并给药中的空行(一定要删)
                Call Get一并给药范围(lng相关ID, lngBegin, lngEnd)
                For i = lngEnd To lngBegin Step -1 '必须反向
                    If vsAdvice.RowData(i) = 0 Then Call DeleteRow(i)
                Next
                
                '删除之后当前行号可能变了
                lngRow = vsAdvice.FindRow(lng医嘱ID, lngBegin)
                
                '一并给药只删除当前行
                Call DeleteRow(lngRow)
            Else
                '单独的成药：删除给药途径行及当前行
                i = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                Call DeleteRow(i)
                Call DeleteRow(lngRow)
            End If
        ElseIf InStr(",D,F,K,", vsAdvice.TextMatrix(lngRow, COL_类别)) > 0 Then
            Call Delete检查手术输血(lngRow)
            Call DeleteRow(lngRow)
        ElseIf RowIn配方行(lngRow) Then
            '删除组成味药及煎法行:删除之后重新定位的当前行
            lngRow = Delete中药配方(lngRow)
            '删除当前行(中药用法行)
            Call DeleteRow(lngRow)
        ElseIf RowIn检验行(lngRow) Then
            lngRow = Delete检验组合(lngRow)
            Call DeleteRow(lngRow)
        Else
            Call DeleteRow(lngRow)
        End If
        
        mblnNoSave = True '标记为未保存
    Else
        '空行直接删除
        Call DeleteRow(lngRow)
    End If
    
    '重新定位行
    If vsAdvice.RowHidden(vsAdvice.Row) Then
        i = GetPreRow(vsAdvice.Row)
        If i = -1 Then i = GetNextRow(vsAdvice.Row)
        If i <> -1 Then vsAdvice.Row = i
    End If
    
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    mblnRowChange = True
    vsAdvice.Redraw = flexRDDirect
    Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
End Sub

Private Sub DeleteRow(ByVal lngRow As Long, Optional ByVal blnClear As Boolean, Optional blnDelID As Boolean = True)
'功能：删除表格中的一行,但不改变当前行
'参数：blnClear=是否仅清除该行内容,不删除
'      blnDelID=是否记录要删除的医嘱ID
    Dim lngCol As Long, blnDraw As Boolean, blnChange As Boolean
    
    With vsAdvice
        lngCol = .Col
        blnDraw = .Redraw
        blnChange = mblnRowChange
        
        mblnRowChange = False
        .Redraw = flexRDNone
        
        If .RowData(lngRow) <> 0 Then
            '调整序号
            Call AdviceSet医嘱序号(lngRow + 1, -1)
        End If
            
        '如果为行1且仅剩行1或仅清除,则保留
        If Not (lngRow = .FixedRows And .Rows = .FixedRows + 1) And Not blnClear Then
            .RemoveItem lngRow
        Else
            '清除该行数据
            .RowData(lngRow) = Empty
            .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "" '文字
            .Cell(flexcpData, lngRow, 0, lngRow, .Cols - 1) = Empty '数据
            .Cell(flexcpFontBold, lngRow, .FixedCols, lngRow, .Cols - 1) = False '粗体
            .Cell(flexcpForeColor, lngRow, .FixedCols, lngRow, .Cols - 1) = .ForeColor '文字色
            If .FixedCols > 0 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .FixedCols - 1) = .ForeColorFixed '固定列文字色
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .FixedCols - 1) = .BackColorFixed '固定列背景色
            End If
            Set .Cell(flexcpPicture, lngRow, 0, lngRow, .Cols - 1) = Nothing '单元图片
            
            '单元格边框
            .Select lngRow, .FixedCols, lngRow, COL_执行时间
            .CellBorder vbRed, 0, 0, 0, 0, 0, 0
        End If
        
        .Col = lngCol '因为有删除行,所以调用程序肯定有行定位,所以不必恢复行
        .Redraw = blnDraw
        mblnRowChange = blnChange
    End With
End Sub

Private Sub Delete检查手术输血(ByVal lngRow As Long)
'功能：1.删除检查组合项目的部位行
'      2.删除手术项目的附加手术行及麻醉项目行
'      3.删除输血项目的输血途径行
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_相关ID) '不一定有,所以用查找
    If i <> -1 Then
        lngBegin = i
        For i = lngBegin To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = vsAdvice.RowData(lngRow) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
        For i = lngEnd To lngBegin Step -1
            Call DeleteRow(i)
        Next
    End If
End Sub

Private Function Delete中药配方(ByVal lngRow As Long) As Long
'功能：删除中药配方的组成味药及煎法行
'参数：lngRow=中药配方用法行(可见)
'返回：删除之后重新定位的当前行(中药用法行)
    Dim lngBegin As Long, lngEnd As Long
    Dim lng医嘱ID As Long, i As Long
    
    lng医嘱ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng医嘱ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '因为是在前面删除,需要重新定位到中药用法行
    i = vsAdvice.FindRow(lng医嘱ID)
    vsAdvice.Row = i '不可能找不到
    
    mblnRowChange = True
    
    Delete中药配方 = vsAdvice.Row
End Function

Private Function Delete检验组合(ByVal lngRow As Long) As Long
'功能：删除一并采集的多个检验项目行
'参数：lngRow=采集方法行(可见)
'返回：删除之后重新定位的当前行(采集方法行)
    Dim lngBegin As Long, lngEnd As Long
    Dim lng医嘱ID As Long, i As Long
    
    lng医嘱ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng医嘱ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '因为是在前面删除,需要重新定位到采集方法行
    i = vsAdvice.FindRow(lng医嘱ID)
    vsAdvice.Row = i '不可能找不到
    
    mblnRowChange = True
    
    Delete检验组合 = vsAdvice.Row
End Function

Private Function Get检查部位方法(ByVal lngRow As Long) As String
'功能：获取指定行的检查部位方法串
'参数：lngRow=检查医嘱的可见行
'返回："部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
'      如果是老的检查组合方式，或者是以前的单部位检查，则返回空以便程序识别
    Dim str部位 As String, str部位Last As String
    Dim str方法 As String, i As Long
    
    With vsAdvice
        For i = lngRow + 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                If Val(.TextMatrix(i, COL_诊疗项目ID)) <> Val(.TextMatrix(lngRow, COL_诊疗项目ID)) Then Exit Function '老的方式
                
                If .TextMatrix(i, COL_标本部位) <> "" Then
                    If .TextMatrix(i, COL_标本部位) <> str部位Last And str部位Last <> "" Then
                        str部位 = str部位 & "|" & str部位Last & IIF(str方法 <> "", ";" & Mid(str方法, 2), "")
                        str方法 = ""
                    End If
                    If .TextMatrix(i, COL_检查方法) <> "" Then
                        str方法 = str方法 & "," & .TextMatrix(i, COL_检查方法)
                    End If
                    
                    str部位Last = .TextMatrix(i, COL_标本部位)
                End If
            Else
                Exit For
            End If
        Next
        If str部位Last <> "" Then
            str部位 = str部位 & "|" & str部位Last & IIF(str方法 <> "", ";" & Mid(str方法, 2), "")
        End If
        Get检查部位方法 = Mid(str部位, 2) & vbTab & 0
    End With
End Function

Private Function Get手术附加IDs(ByVal lngRow As Long) As String
'功能：获取指定手术行的附加手术及麻醉项目ID串
'返回："手术ID1,手术ID2,...;麻醉ID",其中可能没有附加手术和麻醉
    Dim strTmp As String, lng麻醉ID As Long, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_相关ID)
    If i <> -1 Then
        For i = i To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = vsAdvice.RowData(lngRow) Then
                If vsAdvice.TextMatrix(i, COL_类别) = "G" Then
                    lng麻醉ID = Val(vsAdvice.TextMatrix(i, COL_诊疗项目ID))
                Else
                    strTmp = strTmp & "," & Val(vsAdvice.TextMatrix(i, COL_诊疗项目ID))
                End If
            Else
                Exit For
            End If
        Next
    End If
    Get手术附加IDs = Mid(strTmp, 2) & ";" & IIF(lng麻醉ID = 0, "", lng麻醉ID)
End Function

Private Function Get中药配方IDs(ByVal lngRow As Long) As String
'功能：获取中药配方的组成味药及煎法ID串
'返回："中药规格ID1,单量1,脚注1;中药规格ID2,单量2,脚注2;...|煎法ID|中药形态|付数|药房ID"
    Dim lng煎法ID As Long, str中药IDs As String, i As Long, lng形态 As Long
    Dim lng付数 As Long, lng药房ID As Long
    Dim strTmp As String
    
    With vsAdvice
        lng形态 = Val(.TextMatrix(lngRow, COL_中药形态))    '用法行
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                If .TextMatrix(i, COL_类别) = "E" Then
                    lng煎法ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                    strTmp = .TextMatrix(i, COL_标本部位) '代表中药的 煎量
                ElseIf .TextMatrix(i, COL_类别) = "7" Then
                    str中药IDs = Val(.TextMatrix(i, COL_收费细目ID)) & "," & _
                        .TextMatrix(i, COL_单量) & "," & .TextMatrix(i, COL_医生嘱托) & _
                        ";" & str中药IDs
                    If lng药房ID = 0 Then
                        lng药房ID = Val(.TextMatrix(i, COL_执行科室ID))
                        lng付数 = Val(.TextMatrix(i, COL_总量))
                    End If
                End If
            Else
                Exit For
            End If
        Next
        Get中药配方IDs = Mid(str中药IDs, 1, Len(str中药IDs) - 1) & "|" & lng煎法ID & "|" & lng形态 & "|" & lng付数 & "|" & lng药房ID & "|" & strTmp
    End With
End Function

Private Function Get检验组合IDs(ByVal lngRow As Long) As String
'功能：获取一并采集的检验组合项目ID及标本
'返回："'      检验组合="项目ID1,项目ID2,...;检验标本" 如果是新版LIS的模式则是："项目ID1|指标1|指标2...,项目ID2|指标1|指标2...,...;检验标本""
    Dim str项目IDs As String, str标本 As String, i As Long
    Dim j As Long
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                If Val(.TextMatrix(i, COL_组合项目ID)) = 0 And mblnNewLIS Then
                    For j = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_相关ID)) = .RowData(lngRow) Then
                            If Val(.TextMatrix(j, COL_组合项目ID)) = Val(.TextMatrix(i, COL_诊疗项目ID)) And Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                                str项目IDs = "|" & Val(.TextMatrix(j, COL_诊疗项目ID)) & str项目IDs
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    str项目IDs = "," & Val(.TextMatrix(i, COL_诊疗项目ID)) & str项目IDs
                Else
                    If Not mblnNewLIS Then
                        str项目IDs = "," & Val(.TextMatrix(i, COL_诊疗项目ID)) & str项目IDs
                    End If
                End If
                str标本 = .TextMatrix(i, COL_标本部位)
            Else
                Exit For
            End If
        Next
    End With
    Get检验组合IDs = Right(str项目IDs, Len(str项目IDs) - 1) & ";" & str标本
End Function

Private Function RowIn检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于检验组合中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then
            '采集方法行
            If .TextMatrix(lngRow - 1, COL_类别) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
                RowIn检验行 = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, COL_类别) = "C" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            '检验项目行
            RowIn检验行 = True: Exit Function
        End If
    End With
End Function

Private Function RowIn配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于中药配方中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_类别) = "E" Then
            If Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then
                '用法行
                If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) _
                    And .TextMatrix(lngRow - 1, COL_类别) = "E" Then
                    RowIn配方行 = True: Exit Function
                End If
            Else
                '煎法行
                If .TextMatrix(lngRow - 1, COL_类别) = "7" _
                    And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    RowIn配方行 = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, COL_类别) = "7" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            '中药行
            RowIn配方行 = True: Exit Function
        End If
    End With
End Function

Private Function RowIn一并给药(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中
'参数：lngRow=可见的行,可能是空行
'说明：一并给药的范围中可能存在空行
    Dim lngPreRow As Long, lngNextRow As Long
    Dim lng相关ID As Long, blnGroup As Boolean, i As Long
    
    lngPreRow = GetPreRow(lngRow)
    lngNextRow = GetNextRow(lngRow)
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            If lngPreRow <> -1 And lngNextRow <> -1 Then
                If Val(.TextMatrix(lngPreRow, COL_相关ID)) = Val(.TextMatrix(lngNextRow, COL_相关ID)) _
                    And Val(.TextMatrix(lngPreRow, COL_相关ID)) <> 0 _
                    And InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 _
                    And InStr(",5,6,", .TextMatrix(lngNextRow, COL_类别)) > 0 Then
                    blnGroup = True
                End If
            End If
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 _
            And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            
            lng相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
            If lngPreRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 _
                    And Val(.TextMatrix(lngPreRow, COL_相关ID)) = lng相关ID Then blnGroup = True
            End If
            If Not blnGroup And lngNextRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngNextRow, COL_类别)) > 0 _
                    And Val(.TextMatrix(lngNextRow, COL_相关ID)) = lng相关ID Then blnGroup = True
            End If
        End If
    End With
    RowIn一并给药 = blnGroup
End Function

Private Function AdviceInput(rsInput As ADODB.Recordset, ByVal lngRow As Long) As Boolean
'功能：根据新输的诊疗项目(新增或更换)设置缺省的医嘱数据
'参数：rsInput=输入或选择返回的记录集,lngRow=当前输入行
'返回：本次录入是否有效
    Dim intType As Integer
    Dim str过敏 As String, blnGroup As Boolean
    Dim lng用法ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim strExtData As String, strAppend As String
    Dim strMsg As String, vMsg As VbMsgBoxResult
    Dim i As Long
    Dim objControl As CommandBarControl
    Dim lngBegin As Long, lngEnd As Long
    Dim blnOK As Boolean
    Dim lng药品ID As Long
    Dim t_Pati As TYPE_PatiInfoEx
    Dim bln备血 As Boolean '是否为备血医嘱 备血=0，用血=1,存于K类别医嘱行的 检查方法  字段;备血-采集方式 / 用血-输血途径
    Dim strWhere As String
    
    On Error GoTo errH
        
    lngPreRow = GetPreRow(lngRow) '取上一有效行,某些内容缺省与上一行相同
    lngNextRow = GetNextRow(lngRow) '取下一有效行
    
    '项目附加数据输入及输入合法性检查
    '---------------------------------------------------------------------------------------------------------------
    txt医嘱内容.Text = rsInput!名称 '暂时显示
    
    With vsAdvice
        '检验项目：采集方法判断
        If rsInput!类别ID = "C" Then
            '所有数据中取一个缺省的采集方法,同时判断是否有采集方法数据
            lng用法ID = Get缺省用法ID(6, mint范围)
            If lng用法ID = 0 Then
                .Refresh
                MsgBox "没有可用的标本采集方法,请先到诊疗项目管理中设置！", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '缺省与上一行相同
            If lngPreRow <> -1 Then
                If RowIn检验行(lngPreRow) Then
                    If Val(.TextMatrix(lngPreRow, COL_是否停用)) = 0 Then lng用法ID = Val(.TextMatrix(lngPreRow, COL_诊疗项目ID))
                End If
            End If
        End If
        
        '输血医嘱：输血途径判断
        If rsInput!类别ID = "K" Then
            If gbln血库系统 Then
                vMsg = frmMsgBox.ShowMsgBox("请选择输血医嘱类型。", Me, , 2)
                If vMsg = vbNo Then
                    bln备血 = True
                ElseIf vMsg = vbCancel Then
                    Exit Function
                End If
            Else
                bln备血 = True
            End If
            '所有数据中取一个缺省的输血途径
            strWhere = ""
            If bln备血 = False And gbln血库系统 = True Then
                strWhere = " And NVL(执行分类,0)=1 "
            End If
            lng用法ID = Get缺省用法ID(IIF(bln备血 And gbln血库系统, 9, 8), mint范围, strWhere)
            
            If lng用法ID = 0 Then
                .Refresh
                 MsgBox "没有可用的输血" & IIF(bln备血 And gbln血库系统, "采集方法", "途径") & ",请先到诊疗项目管理中设置！", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '缺省与上一行相同
            If lngPreRow <> -1 Then
                If .TextMatrix(lngPreRow, COL_类别) = "K" And Val(.TextMatrix(lngPreRow, COL_检查方法)) = IIF(bln备血, "0", "1") Then
                    i = .FindRow(CStr(.RowData(lngPreRow)), lngPreRow + 1, COL_相关ID)
                    If i <> -1 Then
                        If Val(.TextMatrix(i, COL_是否停用)) = 0 Then lng用法ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                    End If
                End If
            End If
        End If
        
        '中药配方：给成与中药用法判断
        If InStr(",7,8,", rsInput!类别ID) > 0 Then
            If rsInput!类别ID = "8" Then
                If GetGroupCount(rsInput!诊疗项目ID, mint范围, False) = 0 Then
                    .Refresh
                    MsgBox """" & rsInput!名称 & """是一个中药配方，但没有设置有效的组成中药。" & vbCrLf & "请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                    .Refresh: Exit Function
                End If
                
                '部份药无效的提示
                strMsg = GetGroupNone(rsInput!诊疗项目ID, mint范围)
                If strMsg <> "" Then
                    .Refresh
                    MsgBox "配方""" & rsInput!名称 & """中以下药品已撤档或服务对象不匹配：" & _
                        vbCrLf & vbCrLf & vbTab & strMsg & vbCrLf & vbCrLf & "这些药品将不会出现在配方中。", vbInformation, gstrSysName
                    .Refresh
                End If
            End If
        
            '所有数据中取一个缺省的中药用法,同时判断是否有中药用法数据
            lng用法ID = Get缺省用法ID(4, mint范围)
            If lng用法ID = 0 Then
                .Refresh
                MsgBox "没有可用的中药用(服)法,请先到诊疗项目管理中设置！", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '中药用法缺省与上一行相同
            If RowIn配方行(lngPreRow) Then
                If Val(.TextMatrix(lngPreRow, COL_是否停用)) = 0 Then lng用法ID = Val(.TextMatrix(lngPreRow, COL_诊疗项目ID))
            End If
        End If
        
        '中西成药：给药途径判断
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            '给药途径缺省与上一个行相同剂型的相同
            If lngPreRow <> -1 And Not IsNull(rsInput!药品剂型) Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 And .TextMatrix(lngPreRow, COL_药品剂型) = NVL(rsInput!药品剂型) Then
                    i = .FindRow(CLng(.TextMatrix(lngPreRow, COL_相关ID)), lngPreRow + 1)
                    If i <> -1 Then
                        If Val(.TextMatrix(i, COL_是否停用)) = 0 Then lng用法ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                    End If
                End If
            End If
        End If
        
        '中西成药：一并给药的判断
        blnGroup = RowIn一并给药(lngRow)
        If blnGroup Then
            If rsInput!类别ID = "9" Then
                .Refresh
                MsgBox "不能在一并给药的药品中直接输入成套方案。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
            If .RowData(lngRow) = 0 Then
                '一并给药中的待输入空行：只有插入在一并给药的中间,才能自动成为一并给药
                lngGroupRow = lngPreRow
            Else
                '一并给药中的药品行：可能是第一行或最后一行'取当前行的下一行，避免在操作已有医嘱时重选诊疗项目操作时，当前行的内容被删除，后续过程无法取到其中的值
                If lngPreRow = -1 Then
                    lngGroupRow = vsAdvice.FindRow(.TextMatrix(lngRow, COL_相关ID), lngRow + 1, COL_相关ID)
                Else
                    If InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 _
                        And Val(.TextMatrix(lngPreRow, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                        lngGroupRow = lngPreRow
                    Else
                        lngGroupRow = lngNextRow
                    End If
                End If
            End If
            
            '一并给药的,类别，期效必须相同
            If decode(rsInput!类别ID, "5", "Y", "6", "Y", "N") <> decode(.TextMatrix(lngGroupRow, COL_类别), "5", "Y", "6", "Y", "N") Then
                .Refresh
                MsgBox "该组一并给药的药品必须都为西成药或中成药。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            If zlCommFun.GetNeedName(cbo期效.Text) <> .TextMatrix(lngGroupRow, COL_期效) Then
                .Refresh
                MsgBox "该组一并给药的药品必须都为""" & .TextMatrix(lngGroupRow, COL_期效) & """。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
            i = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_相关ID)), lngGroupRow + 1)
            lng用法ID = Val(.TextMatrix(i, COL_诊疗项目ID)) '一并给药的给药途径相同
            
            '检查一并给药的的给药途径是否适合于当前输入药品(非一并给药的缺省用法在输入函数中作了判断处理)
            If Not Check适用用法(lng用法ID, rsInput!诊疗项目ID, mint范围) Then
                .Refresh
                MsgBox "一并的给药途径为""" & .TextMatrix(i, col_医嘱内容) & """，不适用于当前输入药品。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
        End If
    
        '成套项目
        If rsInput!类别ID = "9" Then
            If GetGroupCount(rsInput!诊疗项目ID, mint范围) = 0 Then
                .Refresh
                MsgBox """" & rsInput!名称 & """是一个成套方案，但没有设置有效的组成项目。" & vbCrLf & "请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            strExtData = frmSchemeSelect.ShowMe(Me, rsInput!诊疗项目ID, mint范围)
            If strExtData = "" Then .Refresh: Exit Function
        End If
    
        '需要输入更多数据的一些项目
        '---------------------------------------------------------------------------------------------------------------
        intType = -1
        If rsInput!类别ID <> "9" Then strExtData = ""
        If rsInput!类别ID = "D" Then
            '检查项目：都要扩展编辑了，不象以前还有单部位项目
            intType = 0
        ElseIf rsInput!类别ID = "F" Then
            '手术：需要输入麻醉项目，及可选择附加手术
            intType = 1
        ElseIf InStr(",7,8,", rsInput!类别ID) > 0 Then
            '中药配方(单味草药当配方处理)
            intType = 2
        ElseIf rsInput!类别ID = "C" Then
            '输入一并采集的多个检验项目及检验标本
            intType = 4
            strExtData = rsInput!诊疗项目ID & ";" & NVL(rsInput!规格) '项目;标本
        End If
        If intType <> -1 Then
            If intType = 2 Then
                lng药品ID = Val("" & rsInput!收费细目ID)   '一组配方时为空
            End If
            On Error Resume Next
            If intType = 2 Then
                blnOK = frmAdviceFormula.ShowMe(Me, Nothing, txt医嘱内容.Hwnd, t_Pati, 3, IIF(mbyt场合 <> 2, 0, 3), cbo期效.ListIndex, mint范围, _
                            , rsInput!诊疗项目ID, strExtData, , lng药品ID)
            Else
                blnOK = frmSchemeEditEx.ShowMe(Me, txt医嘱内容.Hwnd, intType, cbo期效.ListIndex, mint范围, mblnNewLIS, True, rsInput!诊疗项目ID, strExtData)
            End If
            On Error GoTo errH
            
            If Not blnOK Then Exit Function
        End If
    
        '修改已有项目时,先删除当前医嘱的内容
        '---------------------------------------------------------------------------------------------------------------
        If .RowData(lngRow) <> 0 Then
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '西成药、中成药
                If Not blnGroup Then
                    '单个成药删除给药途径行,并清除当前行
                    i = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    Call DeleteRow(i)
                    Call DeleteRow(lngRow, True)
                Else
                    '一组成药时,只清除当前行
                    Call DeleteRow(lngRow, True)
                End If
            ElseIf InStr(",D,F,K,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '检查组合项目、手术项目、输血医嘱
                '删除部位行、手术附加行(附加手术,麻醉项目)、输血途径
                Call Delete检查手术输血(lngRow)
                '清除当前行
                Call DeleteRow(lngRow, True)
            ElseIf RowIn配方行(lngRow) Then
                '中药配方：顺序(序号)要求必须严格控制
                '删除组成味药及煎法行:删除之后重新定位的当前行
                lngRow = Delete中药配方(lngRow)
                '清除当前行(中药用法行)
                Call DeleteRow(lngRow, True)
            ElseIf RowIn检验行(lngRow) Then
                '删除检验项目行:删除之后重新定位的当前行
                lngRow = Delete检验组合(lngRow)
                '清除当前行(采集方法行)
                Call DeleteRow(lngRow, True)
            Else
                '其它项目直接清除当前行内容
                Call DeleteRow(lngRow, True)
            End If
        End If
        
        '当前行新增医嘱
        '---------------------------------------------------------------------------------------------------------------
        If InStr(",7,8,", rsInput!类别ID) > 0 Then
            '中药配方(单味草药当配方处理):处理之后重新定位的当前行
            lngRow = AdviceSet中药配方(rsInput!诊疗项目ID, lngRow, lng用法ID, strExtData)
        ElseIf rsInput!类别ID = "9" Then
            '成套医嘱需要分解为多个项目加入
            Call LoadAdvice(rsInput!诊疗项目ID, lngRow, strExtData)
        ElseIf rsInput!类别ID = "C" Then
            '检验组合
            lngRow = AdviceSet检验组合(lngRow, lng用法ID, strExtData)
        Else
            '中、西成药，卫材，检查(组合)，手术(组合)，输血，及其它诊疗项目
            Call AdviceSet诊疗项目(rsInput, lngRow, lng用法ID, lngGroupRow, strExtData, bln备血)
            
            '自动设置一并给药
            If InStr(",5,6,", rsInput!类别ID) > 0 Then
                If Not RowIn一并给药(lngRow) Then
                    If mblnRowMerge Then
                        '手工使一并给药
                        Call MergeRow(lngPreRow, lngRow) '本来就是显示当前行的内容,不用再强行RowChange
                    ElseIf lngPreRow <> -1 Then
                        '自动使一并给药
                        Set objControl = cbsMain.FindControl(, conMenu_Merge, , True)
                        If objControl.Checked = True Then
                            If .TextMatrix(lngPreRow, COL_类别) = rsInput!类别ID Then
                                If RowIn一并给药(lngPreRow) And RowCanMerge(lngPreRow, lngRow) And GetNextRow(lngRow) = -1 Then
                                    mblnRowMerge = True: cbsMain.RecalcLayout '*即时刷新
                                    Call MergeRow(lngPreRow, lngRow, False)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            If i <> lngRow Then vsAdvice.TextMatrix(i, col_缺省) = vsAdvice.TextMatrix(lngRow, col_缺省)
        Next
        
        '重新自动调整行高
        Call .AutoSize(col_医嘱内容)
    End With
    mblnNoSave = True '标记为未保存
    
    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub MergeRow(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional ByVal blnCheck As Boolean = True)
'功能：将两行设置为一并给药
'参数：lngRow1=前面行,可能本来已经属于一并给药
'      lngRow2=当前行
'说明：设置完成后,表格仍定位在原lngRow2的当前行
    Dim lngBegin As Long, lngEnd As Long
    Dim blnDo As Boolean, lngTmp As Long
    
    With vsAdvice
        If blnCheck Then
            blnDo = RowCanMerge(lngRow1, lngRow2)
        Else
            blnDo = True
        End If
        If blnDo Then
            mblnRowChange = False: .Redraw = flexRDNone
            lngTmp = .RowData(lngRow2) '记录以再定位到当前行
            '先取消之前的一并给药
            If RowIn一并给药(lngRow1) Then
                Call Get一并给药范围(Val(.TextMatrix(lngRow1, COL_相关ID)), lngBegin, lngEnd)
                Call AdviceSet单独给药(lngBegin, lngEnd)
                lngRow1 = lngBegin
                lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            End If
            Call AdviceSet一并给药(lngRow1, lngRow2)
            lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            .Row = lngRow2
            mblnRowChange = True: .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub SplitRow(ByVal lngRow As Long)
'功能：将指定行从一并给药中独立出来(该组一并给药必须至少包含三行)
'参数：lngRow=当前行,且为一并给药中的最后一药品行
'说明：设置完成后,表格仍定位在原lngRow的当前行
    Dim lngBegin As Long, lngEnd As Long, lngTmp As Long
    
    With vsAdvice
        mblnRowChange = False: .Redraw = flexRDNone
        lngTmp = .RowData(lngRow) '记录用于恢复定位当前行
        Call Get一并给药范围(Val(.TextMatrix(lngRow, COL_相关ID)), lngBegin, lngEnd)
        
        '先取消整个的一并给药
        Call AdviceSet单独给药(lngBegin, lngEnd)
        
        '再设置除最后行外的行为一并给药
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        lngEnd = GetPreRow(lngRow)
        Call AdviceSet一并给药(lngBegin, lngEnd)
        
        '恢复当前行
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        .Row = lngRow
        mblnRowChange = True: .Redraw = flexRDDirect
    End With
End Sub

Private Function GetTableFromRecordSet() As String
'功能：根据传入的记录集产生一个虚拟表
    Dim strSql As String, i As Long
    Dim strValue As String, strFiled As String
    Dim blnHave As Boolean
    Dim blnHave备选 As Boolean
    Dim lng序号 As Long
    
    For i = 0 To mrsScheme.Fields.Count - 1
        If mrsScheme.Fields(i).Name = "是否缺省" Then blnHave = True
        If mrsScheme.Fields(i).Name = "是否备选" Then blnHave备选 = True
    Next
    
    If mrsScheme.RecordCount > 0 Then
        mrsScheme.MoveFirst
        Do While Not mrsScheme.EOF
            '门诊不允许自由录入医嘱
            If Not (mint范围 = 1 And IsNull(mrsScheme!诊疗项目ID)) Then
                lng序号 = lng序号 + 1
                strFiled = lng序号 & IIF(strSql = "", " as 顺序", "")
                If Not blnHave Then
                    strFiled = strFiled & ",1" & IIF(strSql = "", " as 是否缺省", "")
                End If
                If Not blnHave备选 Then
                    strFiled = strFiled & ",1" & IIF(strSql = "", " as 是否备选", "")
                End If
                
                For i = 0 To mrsScheme.Fields.Count - 1
                    If IsNull(mrsScheme.Fields(i).value) Then
                        Select Case mrsScheme.Fields(i).Type
                            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                strValue = "-Null"
                            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                strValue = "Null+Sysdate"
                            Case Else
                                strValue = "Null"
                        End Select
                    Else
                        Select Case mrsScheme.Fields(i).Type
                            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                strValue = "'" & Replace(Replace(mrsScheme.Fields(i).value, "[", "("), "]", ")") & "'"
                            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                strValue = mrsScheme.Fields(i).value
                            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                strValue = "To_Date('" & Format(mrsScheme.Fields(i).value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End Select
                    End If
                    
                    If strSql = "" Then
                        strFiled = strFiled & "," & strValue & " as " & mrsScheme.Fields(i).Name '首行加别名
                    Else
                        strFiled = strFiled & "," & strValue
                    End If
                Next
                
                strSql = strSql & " Union ALL Select " & strFiled & " From Dual"
            End If
            mrsScheme.MoveNext
        Loop
        mrsScheme.MoveFirst
        If strSql <> "" Then
            GetTableFromRecordSet = "(" & Mid(strSql, 12) & ")"
        End If
    End If
End Function

Private Function GetTableFromAdvice(ByVal str组IDs As String, ByVal lng病人ID As Long) As String
'功能：根据医嘱的病人及组ID，虚拟产生一个"诊疗项目组合"表
'注意：在主SQL中，病人ID条件的顺序是[3],组ID条件顺序为[4]
    Dim strSql As String
    
    '门诊不支持自由录入医嘱
    strSql = "Select /*+ Rule*/ 序号 as 顺序,1 as 是否缺省,0 as 是否备选,ID as 序号,相关ID as 相关序号,医嘱期效 as 期效,A.诊疗项目ID,A.医嘱内容,A.天数,A.单次用量,A.总给予量,A.医生嘱托," & _
        " A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.执行科室ID,A.执行时间方案 as 时间方案,A.执行性质,A.执行标记,A.收费细目ID,A.标本部位,A.检查方法,A.配方ID,a.组合项目ID" & _
        " From 病人医嘱记录 A,Table(Cast(f_Num2List([4]) As zlTools.t_NumList)) B" & _
        " Where Nvl(A.相关ID,A.ID)=B.Column_Value And A.病人ID=[3]" & IIF(mint范围 = 1, " And A.诊疗项目ID is Not NULL", "")
    
    GetTableFromAdvice = "(" & strSql & ")"
End Function

Private Sub LoadAdvice(ByVal lng成套ID As Long, ByVal lngRow As Long, Optional ByVal str序号 As String, Optional ByVal lng病人ID As Long)
'功能：输入成套项目(包括一并给药,检查组合,手术附加,中药配方)
'参数：lng成套ID=为0时表示从传入的记录集中或者从病人医嘱记录读取
'      lngRow=空的输入行(可能是插入的新行,但不位于一并给药中间)，如果为0表示清除当前内容
'      str序号=要读取的成套方案内容的明细序号，或者医嘱记录的组ID
'      lng病人ID=当不为0时，表示通过"str序号"为组ID串读取病人医嘱记录
    Dim rsItems As New ADODB.Recordset
    Dim rs规格 As New ADODB.Recordset
    Dim rs材料 As New ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    
    Dim lngCurRow As Long, intCount As Integer, lng序号 As Long
    Dim bln给药途径 As Boolean, bln采集方法 As Boolean, bln输血途径 As Boolean
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim bln中药用法 As Boolean, bln中药煎法 As Boolean, bln配方 As Boolean
    Dim lng倍数 As Long, vBookMark As Variant, str药房IDs As String
    Dim lng相关ID As Long, strSQL序号 As String, str记录 As String
    Dim int频率性质 As Integer, str适用范围 As String, str频率 As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Me.Refresh
    
    '下医嘱缺省的天数
    If msng天数 = 0 Then msng天数 = 1

    If lng成套ID <> 0 Then
        '门诊不支持自由录入医嘱，成套选择器中已限制
        str记录 = "(Select 1 as 顺序,1 as 是否缺省,0 as 是否备选,A.* From 诊疗项目组合 A Where A.诊疗组合ID=[1])"
        
        If str序号 <> "" Then
            If Left(str序号, 1) = "+" Then
                strSQL序号 = " And Instr([2],','||A.序号||',')>0"
            ElseIf Left(str序号, 1) = "-" Then
                strSQL序号 = " And Instr([2],','||A.序号||',')=0"
            End If
        End If
    ElseIf lng病人ID <> 0 Then
        str记录 = GetTableFromAdvice(str序号, lng病人ID)
    Else
        str记录 = GetTableFromRecordSet
        If str记录 = "" Then
            '门诊不支持自由录入医嘱
            str记录 = "(Select 1 as 顺序,1 as 是否缺省,0 as 是否备选,A.* From 诊疗项目组合 A Where A.诊疗组合ID=[1]" & IIF(mint范围 = 1, " And A.诊疗项目ID is Not NULL", "") & ")"
        End If
    End If
    
    '药品规格信息:虽然存了收费细目ID,但长嘱可能没存,以前的数据也没存
    strSql = "Select A.序号,B.药名ID,B.药品ID,B.剂量系数,B." & IIF(mint范围 = 1, "门诊", "住院") & "可否分零 As 可否分零,C.编码,Nvl(D.名称,C.名称) as 名称,C.规格,C.产地," & _
        decode(mint范围, 1, "B.门诊包装 as 包装系数,B.门诊单位 as 包装单位", 2, "B.住院包装 as 包装系数,B.住院单位 as 包装单位", "C.计算单位 as 包装单位,1 as 包装系数") & _
        " From " & str记录 & " A,药品规格 B,收费项目目录 C,收费项目别名 D" & _
        " Where A.诊疗项目ID=B.药名ID And B.药品ID=C.ID" & strSQL序号 & _
        " And C.ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=[5]" & _
        " Order by A.顺序,A.序号,C.编码"
    Set rs规格 = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng成套ID, "," & Mid(str序号, 2) & ",", lng病人ID, str序号, IIF(gbyt药品名称显示 = 0, 1, 3))
    
    '卫材信息
    strSql = "Select A.序号,B.材料ID,B.跟踪在用,C.名称,C.计算单位" & _
        " From 诊疗项目组合 A,材料特性 B,收费项目目录 C" & _
        " Where A.收费细目ID=B.材料ID And B.材料ID=C.ID" & _
        " And A.诊疗组合ID=[1]" & strSQL序号 & _
        " Order by A.序号,C.编码"
    Set rs材料 = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng成套ID, "," & Mid(str序号, 2) & ",")
    
    '按序号排列后应该与医嘱编辑时的次序一致
    strSql = "Select A.是否缺省,A.是否备选,A.期效,A.序号,A.相关序号,A.诊疗项目ID,A.收费细目ID,A.医嘱内容,A.天数,A.总给予量,A.单次用量," & _
        " A.医生嘱托,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.执行科室ID,B.类别,B.名称,B.计算单位,Decode(B.类别,'D',A.标本部位," & _
        " Nvl(A.标本部位,B.标本部位)) as 标本部位,A.检查方法,A.时间方案,Nvl(A.执行性质,B.执行科室) as 执行性质," & _
        " Nvl(A.执行标记,0) as 执行标记,B.操作类型,B.计算方式,B.执行频率,C.毒理分类,C.抗生素,C.药品剂型,C.品种医嘱,A.配方ID," & _
        " c.临床自管药,a.组合项目ID,d.名称 As 适用证候,b.撤档时间,b.执行分类 " & _
        " From " & str记录 & " A,诊疗项目目录 B,药品特性 C,疾病编码目录 D" & _
        " Where Nvl(A.诊疗项目ID,0)=B.ID(+) And Nvl(A.诊疗项目ID,0)=C.药名ID(+) And Nvl(a.组合项目ID,0)=d.ID(+) " & strSQL序号 & _
        " Order by A.顺序,A.序号"
    Set rsItems = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng成套ID, "," & Mid(str序号, 2) & ",", lng病人ID, str序号)
    With vsAdvice
        mblnRowChange = False
        .Redraw = flexRDNone
        If lngRow = 0 And lng病人ID = 0 Then
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
            lngRow = .FixedRows
        End If
        If lng病人ID <> 0 Then
            '找到输入行
            For i = 1 To .Rows - 1
                If .TextMatrix(i, col_医嘱内容) = "" Then
                    If i = .Rows - 1 Then
                        lngRow = i
                    Else
                        .RemoveItem i
                        .Rows = .Rows + 1
                        lngRow = .Rows - 1
                    End If
                    Exit For
                End If
            Next
            If lngRow = 0 Then
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
        End If
        intCount = 0 '已经设置的行数
        lng序号 = GetCurRow序号(lngRow) '起始序号
        
        For i = 1 To rsItems.RecordCount
            lngCurRow = lngRow + intCount
            If lngCurRow > lngRow Then .AddItem "", lngCurRow
             
            '记录相对ID
            .RowData(lngCurRow) = -1 * rsItems!序号
            If Not IsNull(rsItems!相关序号) Then
                .TextMatrix(lngCurRow, COL_相关ID) = -1 * rsItems!相关序号
            End If
            
            .TextMatrix(lngCurRow, col_缺省) = IIF(rsItems!是否缺省 = 1, -1, 0)
            .TextMatrix(lngCurRow, COL_序号) = lng序号 + intCount
            .TextMatrix(lngCurRow, COL_期效) = IIF(NVL(rsItems!期效, 0) = 0, "长嘱", "临嘱")
            .TextMatrix(lngCurRow, COL_类别) = NVL(rsItems!类别, "*") '自由录入医嘱特殊标记
            
            .TextMatrix(lngCurRow, COL_诊疗项目ID) = NVL(rsItems!诊疗项目ID)
            .TextMatrix(lngCurRow, COL_名称) = NVL(rsItems!名称)
            .TextMatrix(lngCurRow, COL_标本部位) = NVL(rsItems!标本部位)
            .TextMatrix(lngCurRow, COL_检查方法) = NVL(rsItems!检查方法)

            '其它
            .TextMatrix(lngCurRow, COL_计算方式) = NVL(rsItems!计算方式, 0)
            .TextMatrix(lngCurRow, COL_操作类型) = NVL(rsItems!操作类型)
            .TextMatrix(lngCurRow, COL_毒理分类) = NVL(rsItems!毒理分类)
            .TextMatrix(lngCurRow, COL_药品剂型) = NVL(rsItems!药品剂型)
            .TextMatrix(lngCurRow, col_备选) = IIF(rsItems!是否备选 = 1, -1, 0)
            .TextMatrix(lngCurRow, COL_配方ID) = NVL(rsItems!配方ID)
            .TextMatrix(lngCurRow, COL_临床自管药) = rsItems!临床自管药 & ""
            .TextMatrix(lngCurRow, COL_组合项目ID) = "" & rsItems!组合项目ID
            If .TextMatrix(lngCurRow, COL_组合项目ID) <> "" Then
                If .TextMatrix(lngCurRow, COL_类别) = "E" And .TextMatrix(lngCurRow, COL_操作类型) = "4" Then
                    .TextMatrix(lngCurRow, COL_适用证候) = rsItems!适用证候 & ""
                End If
            End If
            .TextMatrix(lngCurRow, COL_抗菌等级) = Val("" & rsItems!抗生素)
            .TextMatrix(lngCurRow, COL_执行标记) = Val("" & rsItems!执行标记)
            
            If Format(NVL(rsItems!撤档时间, "3000/1/1"), "yyyy-MM-dd") <> "3000-01-01" Then
                .TextMatrix(lngCurRow, COL_是否停用) = 1
            End If
            .TextMatrix(lngCurRow, COL_执行分类) = rsItems!执行分类 & ""
            
            '药品规格信息:中草药肯定有,成药按单量与剂量单位自动匹配
            lng倍数 = 0: vBookMark = 0
            '临床路径定义和成套允许不定规格
            If NVL(rsItems!类别) = "7" Or (InStr(",5,6,", NVL(rsItems!类别, "*")) > 0 _
                And (NVL(rsItems!期效, 0) = 1 Or gbln药品按规格下医嘱 And NVL(rsItems!品种医嘱, 0) = 0)) Then
                If Not IsNull(rsItems!收费细目ID) Then
                    rs规格.Filter = "药品ID=" & rsItems!收费细目ID
                Else
                    rs规格.Filter = "药名ID=0"
                End If
                If Not rs规格.EOF Then
                    If IsNull(rsItems!收费细目ID) Then
                        '取剂量系数为单量的最小整倍数的那一个规格
                        If CInt(NVL(rsItems!单次用量, 0)) <> 0 Then
                            Do While Not rs规格.EOF
                                If rs规格!剂量系数 / rsItems!单次用量 = Int(rs规格!剂量系数 / rsItems!单次用量) Then
                                    If rs规格!剂量系数 / rsItems!单次用量 < lng倍数 Or lng倍数 = 0 Then
                                        vBookMark = rs规格.Bookmark
                                        lng倍数 = rs规格!剂量系数 / rsItems!单次用量
                                    End If
                                End If
                                rs规格.MoveNext
                            Loop
                            If vBookMark <> 0 Then rs规格.Bookmark = vBookMark
                        End If
                        If rs规格.EOF Then rs规格.MoveFirst
                    End If
                    .TextMatrix(lngCurRow, COL_名称) = NVL(rs规格!名称)
                    .TextMatrix(lngCurRow, COL_收费细目ID) = rs规格!药品ID
                    .TextMatrix(lngCurRow, COL_剂量系数) = NVL(rs规格!剂量系数)
                    .TextMatrix(lngCurRow, COL_包装系数) = NVL(rs规格!包装系数)
                    .TextMatrix(lngCurRow, COL_包装单位) = NVL(rs规格!包装单位)
                    .TextMatrix(lngCurRow, COL_可否分零) = NVL(rs规格!可否分零, 0)
                End If
            ElseIf NVL(rsItems!类别) = "4" Then
                rs材料.Filter = "材料ID=" & NVL(rsItems!收费细目ID, 0)
                If Not rs材料.EOF Then
                    .TextMatrix(lngCurRow, COL_名称) = NVL(rs材料!名称)
                    .TextMatrix(lngCurRow, COL_包装单位) = NVL(rs材料!计算单位) '散装单位
                    .TextMatrix(lngCurRow, COL_跟踪在用) = NVL(rs材料!跟踪在用, 0)
                End If
                .TextMatrix(lngCurRow, COL_剂量系数) = 1
                .TextMatrix(lngCurRow, COL_包装系数) = 1
                .TextMatrix(lngCurRow, COL_收费细目ID) = NVL(rsItems!收费细目ID, 0)
            End If
                                
            '判断是否特定行
            bln给药途径 = False: bln采集方法 = False: bln输血途径 = False
            bln中药用法 = False: bln中药煎法 = False: bln配方 = False
            If rsItems!类别 = "E" Then
                If IsNull(rsItems!相关序号) Then
                    If Val(.TextMatrix(lngCurRow - 1, COL_相关ID)) = .RowData(lngCurRow) Then
                        If InStr(",5,6,", .TextMatrix(lngCurRow - 1, COL_类别)) > 0 Then
                            bln给药途径 = True
                        ElseIf .TextMatrix(lngCurRow - 1, COL_类别) = "C" Then
                            bln采集方法 = True
                        Else
                            bln中药用法 = True
                        End If
                    End If
                ElseIf .TextMatrix(lngCurRow - 1, COL_类别) = "K" And .RowData(lngCurRow - 1) = Val(.TextMatrix(lngCurRow, COL_相关ID)) Then
                    bln输血途径 = True
                Else
                    bln中药煎法 = True
                End If
            End If
            If rsItems!类别 = "7" Or bln中药煎法 Or bln中药用法 Then bln配方 = True
            
            '频率性质
            If bln采集方法 Then
                '采集方法以检验项目的为准
                j = .FindRow(CStr(.RowData(lngCurRow)), , COL_相关ID)
                int频率性质 = .TextMatrix(j, COL_频率性质)
            Else
                int频率性质 = NVL(rsItems!执行频率, 0)
            End If
            If bln配方 Then
                str适用范围 = 2 '中药配方(包括煎法,用法)用中医
            ElseIf int频率性质 = 1 Then
                str适用范围 = -1 '一次性
            ElseIf int频率性质 = 2 Then
                str适用范围 = -2 '持续性
            ElseIf int频率性质 = 0 Then '可选频率
                If NVL(rsItems!期效, 0) = 1 Then
                    str适用范围 = "1,-1" '临嘱可能为一次性(光名称不能唯一区分)
                Else
                    str适用范围 = 1
                End If
            End If
            If rsItems!执行频次 & "" = "必要时" Then
                str适用范围 = -3
            ElseIf rsItems!执行频次 & "" = "需要时" Then
                str适用范围 = -5
            End If
            
            '频率,频率次数,频率间隔,间隔单位
            .TextMatrix(lngCurRow, COL_频率性质) = int频率性质
            If Not IsNull(rsItems!执行频次) Then
                If Check频率可用(NVL(rsItems!诊疗项目ID, 0), Val(str适用范围), NVL(rsItems!执行频次)) Then 'Val(str适用范围)
                    If Get频率信息_名称(rsItems!执行频次, 0, 0, "", str适用范围) Then
                        .TextMatrix(lngCurRow, COL_频率) = rsItems!执行频次
                        .TextMatrix(lngCurRow, COL_频率次数) = NVL(rsItems!频率次数, 0)
                        .TextMatrix(lngCurRow, COL_频率间隔) = NVL(rsItems!频率间隔, 0)
                        .TextMatrix(lngCurRow, COL_间隔单位) = NVL(rsItems!间隔单位)
                        
                        '临嘱可选频率可能设置为了一次性
                        If NVL(rsItems!期效, 0) = 1 And int频率性质 = 0 And NVL(rsItems!频率次数, 0) = 0 And NVL(rsItems!频率间隔, 0) = 0 Then
                            .TextMatrix(lngCurRow, COL_频率性质) = 1
                        End If
                    End If
                End If
            End If
            If .TextMatrix(lngCurRow, COL_频率) = "" And Not IsNull(rsItems!诊疗项目ID) Then '取缺省的
                If NVL(rsItems!期效, 0) = 1 And int频率性质 = 0 Then
                    If mbln一次性 Then '临嘱缺省为一次性
                        str适用范围 = -1
                        .TextMatrix(lngCurRow, COL_频率性质) = 1
                    Else
                        str适用范围 = 1
                    End If
                End If
                Call Get缺省频率(NVL(rsItems!诊疗项目ID, 0), str适用范围, str频率, int频率次数, int频率间隔, str间隔单位)
                .TextMatrix(lngCurRow, COL_频率) = str频率
                .TextMatrix(lngCurRow, COL_频率次数) = int频率次数
                .TextMatrix(lngCurRow, COL_频率间隔) = int频率间隔
                .TextMatrix(lngCurRow, COL_间隔单位) = str间隔单位
            End If
            
            '天数
            .TextMatrix(lngCurRow, COL_天数) = NVL(rsItems!天数)
            If InStr(",5,6,", NVL(rsItems!类别, "*")) > 0 And NVL(rsItems!天数, 0) > 0 Then
                msng天数 = rsItems!天数 '最近的作为缺省
            End If
            
            '单量
            .TextMatrix(lngCurRow, COL_单量) = FormatEx(NVL(rsItems!单次用量), 5)
            If NVL(rsItems!类别) = "4" Then
                .TextMatrix(lngCurRow, COL_单量单位) = .TextMatrix(lngCurRow, COL_包装单位) '散装单位
            ElseIf bln中药用法 Then
                .TextMatrix(lngCurRow, COL_单量单位) = ""
            ElseIf NVL(rsItems!期效, 0) = 0 Then
                If InStr(",5,6,7,", NVL(rsItems!类别, "*")) > 0 Or InStr(",1,2,", NVL(rsItems!计算方式, 0)) > 0 Then
                    .TextMatrix(lngCurRow, COL_单量单位) = NVL(rsItems!计算单位)
                End If
            Else
                If InStr(",5,6,7,", NVL(rsItems!类别, "*")) > 0 Or (int频率性质 = 0 And InStr(",1,2,", NVL(rsItems!计算方式, 0)) > 0) Then
                    .TextMatrix(lngCurRow, COL_单量单位) = NVL(rsItems!计算单位)
                End If
            End If
            
            '总量
            If InStr(",5,6,", NVL(rsItems!类别, "*")) > 0 Then
                '成药临嘱有总量,以零售单位存放,包装单位显示
                If Not IsNull(rsItems!总给予量) And Val(.TextMatrix(lngCurRow, COL_包装系数)) <> 0 Then
                    .TextMatrix(lngCurRow, COL_总量) = FormatEx(rsItems!总给予量 / Val(.TextMatrix(lngCurRow, COL_包装系数)), 5)
                End If
                If NVL(rsItems!期效, 0) = 1 Then
                    .TextMatrix(lngCurRow, COL_总量单位) = .TextMatrix(lngCurRow, COL_包装单位)
                End If
            Else
                '其它情况有中药和其它临嘱
                If Not IsNull(rsItems!总给予量) Then
                    .TextMatrix(lngCurRow, COL_总量) = rsItems!总给予量
                End If
                If bln配方 Then
                    .TextMatrix(lngCurRow, COL_总量单位) = "付" '中药配方总量单位为"付"
                ElseIf NVL(rsItems!期效, 0) = 1 Then
                    If NVL(rsItems!类别) = "4" Then
                        .TextMatrix(lngCurRow, COL_总量单位) = .TextMatrix(lngCurRow, COL_包装单位) '散装单位
                    Else
                        .TextMatrix(lngCurRow, COL_总量单位) = NVL(rsItems!计算单位)
                    End If
                End If
            End If
            
            '执行时间
            If .TextMatrix(lngCurRow, COL_频率) <> "" And Val(.TextMatrix(lngCurRow, COL_频率性质)) <> 1 Then
                If Not IsNull(rsItems!时间方案) Then
                    If ExeTimeValid(rsItems!时间方案, Val(.TextMatrix(lngCurRow, COL_频率次数)), _
                        Val(.TextMatrix(lngCurRow, COL_频率间隔)), .TextMatrix(lngCurRow, COL_间隔单位)) Then
                        .TextMatrix(lngCurRow, COL_执行时间) = rsItems!时间方案
                    End If
                End If
            End If
            
            '用法的显示
            If bln采集方法 Then
                .TextMatrix(lngCurRow, COL_用法) = rsItems!名称
            ElseIf bln给药途径 Or bln中药用法 Then
                '成药和中药配方的用法,执行时间
                If bln中药用法 Then
                    .TextMatrix(lngCurRow, COL_用法) = rsItems!名称
                End If
                For j = lngCurRow - 1 To lngRow Step -1
                    If Val(.TextMatrix(j, COL_相关ID)) = .RowData(lngCurRow) Then
                        If bln给药途径 Then
                            .TextMatrix(j, COL_用法) = rsItems!名称 & rsItems!医生嘱托  '滴速
                        End If
                        .TextMatrix(j, COL_执行时间) = .TextMatrix(lngCurRow, COL_执行时间)
                    Else
                        Exit For
                    End If
                Next
            ElseIf bln输血途径 Then
                .TextMatrix(lngCurRow - 1, COL_用法) = rsItems!名称
            End If
                                
            '执行性质
            If InStr(",5,6,7,", NVL(rsItems!类别, "*")) > 0 Then
                If NVL(rsItems!执行性质, 0) = 5 Then
                    .TextMatrix(lngCurRow, COL_执行性质) = 5
                Else
                    .TextMatrix(lngCurRow, COL_执行性质) = 4
                End If
            ElseIf NVL(rsItems!类别) = "4" Then
                .TextMatrix(lngCurRow, COL_执行性质) = 4
            Else
                .TextMatrix(lngCurRow, COL_执行性质) = NVL(rsItems!执行性质, 0)
            End If
            
            '执行科室ID:为0-叮嘱,5-院外执行时取出为0
            If InStr(",0,5,", Val(.TextMatrix(lngCurRow, COL_执行性质))) = 0 And NVL(rsItems!执行科室ID, 0) <> 0 Then
                If InStr(",5,6,7,", NVL(rsItems!类别, "*")) > 0 Then
                    str药房IDs = Get可用药房IDs(rsItems!类别, rsItems!诊疗项目ID, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), 0, mint范围)
                    If InStr("," & str药房IDs & ",", "," & rsItems!执行科室ID & ",") > 0 Then
                        .TextMatrix(lngCurRow, COL_执行科室ID) = NVL(rsItems!执行科室ID, 0)
                    End If
                ElseIf NVL(rsItems!类别) = "4" Then
                    str药房IDs = Get可用发料部门IDs(Val(.TextMatrix(lngCurRow, COL_收费细目ID)), 0, mint范围, rsItems!诊疗项目ID)
                    If InStr("," & str药房IDs & ",", "," & rsItems!执行科室ID & ",") > 0 Then
                        .TextMatrix(lngCurRow, COL_执行科室ID) = NVL(rsItems!执行科室ID, 0)
                    End If
                Else
                    .TextMatrix(lngCurRow, COL_执行科室ID) = NVL(rsItems!执行科室ID, 0)
                End If
            End If
                        
            '医生嘱托
            .TextMatrix(lngCurRow, COL_医生嘱托) = NVL(rsItems!医生嘱托)
            
            '----------------------
            '毒麻精药品标识:中药配方及组成味中药不处理
            If InStr(",5,6,", .TextMatrix(lngCurRow, COL_类别)) > 0 And .TextMatrix(lngCurRow, COL_毒理分类) <> "" Then
                If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", .TextMatrix(lngCurRow, COL_毒理分类)) > 0 Then
                    .Cell(flexcpFontBold, lngCurRow, col_医嘱内容) = True
                End If
            End If
            
            '隐蔽一些附加行
            If (InStr(",F,G,D,7,E,C,", NVL(rsItems!类别, "*")) > 0 And Not IsNull(rsItems!相关序号)) Or bln给药途径 Then
                .RowHidden(lngCurRow) = True
            End If
            
            '医嘱内容
            If Not .RowHidden(lngCurRow) Then
                If IsNull(rsItems!诊疗项目ID) Then
                    .TextMatrix(lngCurRow, col_医嘱内容) = rsItems!医嘱内容 '自由录入医嘱
                ElseIf InStr(",F,D,", NVL(rsItems!类别, "*")) > 0 And IsNull(rsItems!相关序号) Then
                    .TextMatrix(lngCurRow, col_医嘱内容) = rsItems!名称 '临时
                Else
                    .TextMatrix(lngCurRow, col_医嘱内容) = AdviceTextMake(lngCurRow)
                End If
            Else
                .TextMatrix(lngCurRow, col_医嘱内容) = rsItems!名称
            End If
            
            '----------------------
            intCount = intCount + 1
            rsItems.MoveNext
        Next
        
        '--------------------------------------------------
        If intCount > 0 Then
            '再取检查和手术的医嘱内容
            For i = lngRow To lngCurRow
                If InStr(",F,D,", .TextMatrix(i, COL_类别)) > 0 And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    .TextMatrix(i, col_医嘱内容) = AdviceTextMake(i)
                End If
            Next
            
            '调整受影响行的序号
            Call AdviceSet医嘱序号(lngCurRow + 1, intCount)
            
            '产生真实的医嘱ID
            For i = lngRow To lngCurRow
                lng相关ID = .RowData(i)
                .RowData(i) = GetNextID
                For j = i - 1 To lngRow Step -1
                    If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                        .TextMatrix(j, COL_相关ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
                For j = i + 1 To lngCurRow
                    If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                        .TextMatrix(j, COL_相关ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
            Next
        End If
        
        '--------------------------------------------------
        If .RowHidden(lngRow) Then '寻找可见行(如配方和检验之后)
            For i = lngRow + 1 To .Rows - 1
                If Not .RowHidden(i) And .RowData(i) <> 0 Then
                    lngRow = i: Exit For
                End If
            Next
        End If
        
        '固定列图标对齐:设置为中对齐,不然擦边框时可能有问题
        Call .AutoSize(col_医嘱内容)
        .Row = lngRow: .Col = col_医嘱内容
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        mblnRowChange = True
    End With
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function AdviceSet中药配方(lng诊疗项目ID As Long, ByVal lngRow As Long, ByVal lng用法ID As Long, ByVal strExtData As String, Optional rsCurr As ADODB.Recordset, Optional ByVal lng配方ID As Long) As Long
'功能：(重新)处理中药配方的缺省医嘱数据
'参数：lng诊疗项目ID=输入的中药配方ID或单味中药ID
'      lngRow=当前输入行
'      lng用法ID=缺省中药用法ID
'      strExtData=包含配方组成味药及煎法数据:规格ID1,数量,脚注;规格ID2,数量,脚注...|中药煎法|中药形态|付数|药房ID|煎量"
'      rsCurr=如果是修改了配方内容后调用,则包含要保持的一些当前值
'返回：处理后的中药配方的当前显示行号
    Dim rsItems As New ADODB.Recordset '中药详细信息
    Dim rsUse As New ADODB.Recordset '中药用法信息
    Dim rs煎法 As New ADODB.Recordset '中药煎法项目信息
    Dim rs用法 As New ADODB.Recordset '中药用法项目信息
    Dim arr中药s As Variant, str中药IDs As String, lng相关ID As Long
    Dim lngCopyRow As Long '缺省参照行
    Dim lngDrugRow As Long '如果缺省参照行是中药配方,则为该配方的第一个中药行
    Dim lngFirstRow As Long '当前配方的第一个中药行
    Dim strSql As String, i As Long
    
    Dim str频率 As String, int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim lng煎法ID As Long, int疗程 As Integer
    Dim str医生 As String, lng医生ID As Long
    Dim lng形态 As Long
    Dim str煎量 As String
        
    On Error GoTo errH
    
    '取上一或下一有效行,某些内容缺省与该行相同
    lngDrugRow = -1
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    If lngCopyRow <> -1 Then
        If RowIn配方行(lngCopyRow) Then
            '如果上一有效行是中药配方的,则取它的第一中药行
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngCopyRow)), , COL_相关ID)
        End If
    End If
    
    '获取相关数据库信息
    '------------------
    arr中药s = Split(Split(strExtData, "|")(0), ";")
    For i = 0 To UBound(arr中药s)
        str中药IDs = str中药IDs & "," & CStr(Split(arr中药s(i), ",")(0))
    Next
    str中药IDs = Mid(str中药IDs, 2)
    lng煎法ID = Val(Split(strExtData, "|")(1))
    lng形态 = Val(Split(strExtData, "|")(2))
    str煎量 = Split(strExtData, "|")(5)
    
    '配方用法信息:直接输入配方时才有可能有,输入单味中药无
    strSql = "Select A.用法ID,A.频次,A.疗程,A.医生嘱托" & _
        " From 诊疗用法用量 A,诊疗项目目录 B" & _
        " Where A.用法ID=B.ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([2],3)") & _
        " And Nvl(A.性质,0)=0 And A.项目ID=[1] And (b.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or b.撤档时间 is NULL)"
    Set rsUse = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng诊疗项目ID, mint范围)
    If Not rsUse.EOF Then lng用法ID = rsUse!用法ID '缺省设置的中药配方用法优先
    
    '配方组成味中药信息:中药无规格概念,对应的的规格记录一定有且只有一条
    strSql = "Select A.计算规则,A.站点,A.类别,A.分类ID,A.ID,A.编码,A.名称,A.标本部位,A.计算单位,A.计算方式,A.执行频率," & _
        "A.适用性别,A.单独应用,A.组合项目,A.操作类型,A.执行安排,A.执行科室,A.服务对象,A.计价性质,A.参考目录ID,A.人员ID,A.建档时间,A.撤档时间,A.录入限量,A.试管编码,A.执行分类,A.执行标记," & _
        "B.药品ID,B.剂量系数,B." & IIF(mint范围 = 1, "门诊", "住院") & "可否分零 As 可否分零," & _
        decode(mint范围, 1, "B.门诊包装 as 包装系数,B.门诊单位 as 包装单位", 2, "B.住院包装 as 包装系数,B.住院单位 as 包装单位", "C.计算单位 as 包装单位,1 as 包装系数") & _
        " From 诊疗项目目录 A,药品规格 B,收费项目目录 C" & _
        " Where A.ID=B.药名ID And B.药品ID=C.ID And B.药品ID IN(Select Column_Value From Table(f_Num2list([1])))"
    Set rsItems = zldatabase.OpenSQLRecord(strSql, Me.Caption, str中药IDs) 'In
        
    '配方煎法项目信息
    Set rs煎法 = Get诊疗项目记录(lng煎法ID)
    
    '配方用法项目信息
    Set rs用法 = Get诊疗项目记录(lng用法ID)
    
    
    '加入配方组成味中药行:按照用户输入顺序
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    mblnRowChange = False
    
    '中药用法的医嘱ID,ID顺序与序号不一定一致
    If Not rsCurr Is Nothing Then
        '修改了配方中的内容,用法行标记为修改,医嘱ID不变
        lng相关ID = rsCurr!医嘱ID
    Else
        '新输入的中药配方
        lng相关ID = GetNextID
    End If
    
    For i = 0 To UBound(arr中药s)
        rsItems.Filter = "药品ID=" & CStr(Split(arr中药s(i), ",")(0)) '应该肯定有
        
        vsAdvice.AddItem "", lngRow
        
        vsAdvice.RowHidden(lngRow) = True
        vsAdvice.RowData(lngRow) = GetNextID
        vsAdvice.TextMatrix(lngRow, COL_相关ID) = lng相关ID '对应到后面的中药用法行
        vsAdvice.TextMatrix(lngRow, COL_期效) = zlCommFun.GetNeedName(cbo期效.Text)
        vsAdvice.TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
        Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
        
        vsAdvice.TextMatrix(lngRow, COL_类别) = rsItems!类别
        vsAdvice.TextMatrix(lngRow, col_医嘱内容) = rsItems!名称
        vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID) = rsItems!ID
        vsAdvice.TextMatrix(lngRow, COL_计算方式) = NVL(rsItems!计算方式, 0)
        vsAdvice.TextMatrix(lngRow, COL_频率性质) = NVL(rsItems!执行频率, 0)
        vsAdvice.TextMatrix(lngRow, COL_操作类型) = NVL(rsItems!操作类型)
        
        vsAdvice.TextMatrix(lngRow, COL_单量) = FormatEx(Val(Split(arr中药s(i), ",")(1)), 5) '单味药的单次用量
        vsAdvice.TextMatrix(lngRow, COL_单量单位) = NVL(rsItems!计算单位)
        vsAdvice.TextMatrix(lngRow, COL_医生嘱托) = CStr(Split(arr中药s(i), ",")(2)) '单味药的脚注
        
        '规格信息:中药不存在规格概念,一定有
        vsAdvice.TextMatrix(lngRow, COL_收费细目ID) = rsItems!药品ID
        vsAdvice.TextMatrix(lngRow, COL_剂量系数) = rsItems!剂量系数
        vsAdvice.TextMatrix(lngRow, COL_包装单位) = rsItems!包装单位
        vsAdvice.TextMatrix(lngRow, COL_包装系数) = rsItems!包装系数
        vsAdvice.TextMatrix(lngRow, COL_可否分零) = NVL(rsItems!可否分零, 0) '对中药实际上无用
        
        If lngFirstRow <> 0 Then
            '与上一行已设置的组成中药相同
            vsAdvice.TextMatrix(lngRow, COL_执行性质) = vsAdvice.TextMatrix(lngFirstRow, COL_执行性质)
            vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = vsAdvice.TextMatrix(lngFirstRow, COL_执行科室ID)
            vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
            vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngFirstRow, COL_频率次数)
            vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔)
            vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
            vsAdvice.TextMatrix(lngRow, COL_总量) = vsAdvice.TextMatrix(lngFirstRow, COL_总量)
            vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
        ElseIf Not rsCurr Is Nothing Then
            '修改了配方内容后重新设置,保持与当前的值
            
            '执行性质:修改时根据当前界面设置决定
            vsAdvice.TextMatrix(lngRow, COL_执行性质) = decode(NVL(rsCurr!执行性质), "自备药", 5, 4)
            '执行科室
            vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = NVL(rsCurr!执行科室ID)
            
            vsAdvice.TextMatrix(lngRow, COL_频率) = NVL(rsCurr!频率)
            vsAdvice.TextMatrix(lngRow, COL_频率次数) = NVL(rsCurr!频率次数)
            vsAdvice.TextMatrix(lngRow, COL_频率间隔) = NVL(rsCurr!频率间隔)
            vsAdvice.TextMatrix(lngRow, COL_间隔单位) = NVL(rsCurr!间隔单位)
            vsAdvice.TextMatrix(lngRow, COL_总量) = NVL(rsCurr!总量)
            vsAdvice.TextMatrix(lngRow, COL_执行时间) = NVL(rsCurr!执行时间)
        Else
            '执行性质:中药配方组成中药相同,缺省=4-指定科室
            vsAdvice.TextMatrix(lngRow, COL_执行性质) = 4
                        
            '执行科室(先在配方界面选择)
            vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = Val(Split(strExtData, "|")(4))
                        
            '执行频率
            '根据用法里面设置的优先
            If Not rsUse.EOF Then
                If Not IsNull(rsUse!频次) Then
                    Call Get频率信息_编码(rsUse!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                    vsAdvice.TextMatrix(lngRow, COL_频率) = str频率
                    vsAdvice.TextMatrix(lngRow, COL_频率次数) = int频率次数
                    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                End If
            End If
            '或缺省与上一行相同
            If vsAdvice.TextMatrix(lngRow, COL_频率) = "" And lngDrugRow <> -1 Then
                If vsAdvice.TextMatrix(lngDrugRow, COL_频率) <> "" Then
                    vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngDrugRow, COL_频率)
                    vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngDrugRow, COL_频率次数)
                    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngDrugRow, COL_频率间隔)
                    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngDrugRow, COL_间隔单位)
                End If
            End If
            '或取缺省值
            If vsAdvice.TextMatrix(lngRow, COL_频率) = "" Then
                Call Get缺省频率(NVL(rsItems!ID, 0), 2, str频率, int频率次数, int频率间隔, str间隔单位)
                vsAdvice.TextMatrix(lngRow, COL_频率) = str频率
                vsAdvice.TextMatrix(lngRow, COL_频率次数) = int频率次数
                vsAdvice.TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                vsAdvice.TextMatrix(lngRow, COL_间隔单位) = str间隔单位
            End If
            
            '总量(付数):临嘱才需要,非散装形态已确定付数
            If Val(Split(strExtData, "|")(3)) > 1 Or lng形态 <> 0 Then
                vsAdvice.TextMatrix(lngRow, COL_总量) = Val(Split(strExtData, "|")(3))
            Else
                If vsAdvice.TextMatrix(lngRow, COL_期效) = "临嘱" And vsAdvice.TextMatrix(lngRow, COL_频率) <> "" Then
                    int疗程 = 1
                    If Not rsUse.EOF Then int疗程 = NVL(rsUse!疗程, 1)
                    '配方付数
                    vsAdvice.TextMatrix(lngRow, COL_总量) = Calc缺省药品总量(1, int疗程, _
                            Val(vsAdvice.TextMatrix(lngRow, COL_频率次数)), _
                            Val(vsAdvice.TextMatrix(lngRow, COL_频率间隔)), _
                            vsAdvice.TextMatrix(lngRow, COL_间隔单位))
                End If
            End If
            
            '执行时间
            If lngDrugRow <> -1 Then '缺省与上一行相同
                If vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngDrugRow, COL_频率) Then
                    vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngDrugRow, COL_执行时间)
                End If
            End If
            If vsAdvice.TextMatrix(lngRow, COL_执行时间) = "" Then '缺省时间方案
                vsAdvice.TextMatrix(lngRow, COL_执行时间) = Get缺省时间(2, vsAdvice.TextMatrix(lngRow, COL_频率), lng用法ID)
            End If
        End If
        
        '---------------------------------------
        If lngFirstRow = 0 Then lngFirstRow = lngRow '该中药配方的第一个组成中药行
        lngRow = lngRow + 1 '保持当前输入行位置
    Next
    
    '加入中药配方煎法行
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.AddItem "", lngRow
    vsAdvice.RowHidden(lngRow) = True
    vsAdvice.RowData(lngRow) = GetNextID
    vsAdvice.TextMatrix(lngRow, COL_相关ID) = lng相关ID
    vsAdvice.TextMatrix(lngRow, COL_期效) = vsAdvice.TextMatrix(lngFirstRow, COL_期效)
    vsAdvice.TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
    Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
    vsAdvice.TextMatrix(lngRow, COL_类别) = rs煎法!类别
    vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID) = lng煎法ID
    vsAdvice.TextMatrix(lngRow, COL_标本部位) = str煎量
    vsAdvice.TextMatrix(lngRow, COL_计算方式) = NVL(rs煎法!计算方式, 0)
    vsAdvice.TextMatrix(lngRow, COL_操作类型) = NVL(rs煎法!操作类型)
    
    '!中药煎法中也存放中药的付数
    vsAdvice.TextMatrix(lngRow, COL_总量) = vsAdvice.TextMatrix(lngFirstRow, COL_总量)
    
    vsAdvice.TextMatrix(lngRow, col_医嘱内容) = rs煎法!名称
    
    vsAdvice.TextMatrix(lngRow, COL_频率性质) = vsAdvice.TextMatrix(lngFirstRow, COL_频率性质) '以药品的为准
    vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
    vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngFirstRow, COL_频率次数)
    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔)
    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
    vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
    
    '执行性质:缺省根据项目设置(不可能为院外执行),修改时根据当前界面设置
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = NVL(rs煎法!执行科室, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = decode(NVL(rsCurr!执行性质), "离院带药", 5, NVL(rs煎法!执行科室, 0))
    End If
    
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_执行性质))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID(rs煎法!类别, lng煎法ID, 0, NVL(rs煎法!执行科室, 0), cbo期效.ListIndex, mint范围)
    End If
    
    '保持当前输入行位置
    lngRow = lngRow + 1
    
    '设置中药配方用法行:中药配方的显示行
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.RowData(lngRow) = lng相关ID
    
    If Get诊疗项目记录(lng诊疗项目ID)!类别 & "" = "8" Then
        vsAdvice.TextMatrix(lngRow, COL_配方ID) = lng诊疗项目ID
    End If
    If lng配方ID <> 0 Then
        vsAdvice.TextMatrix(lngRow, COL_配方ID) = lng配方ID
    End If
    vsAdvice.TextMatrix(lngRow, COL_期效) = vsAdvice.TextMatrix(lngFirstRow, COL_期效)
    vsAdvice.TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
    Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
    vsAdvice.TextMatrix(lngRow, COL_类别) = rs用法!类别
    vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID) = lng用法ID
    vsAdvice.TextMatrix(lngRow, COL_计算方式) = NVL(rs用法!计算方式, 0)
    vsAdvice.TextMatrix(lngRow, COL_操作类型) = NVL(rs用法!操作类型)
    
    '!中药用法中也存放中药的付数
    vsAdvice.TextMatrix(lngRow, COL_总量) = vsAdvice.TextMatrix(lngFirstRow, COL_总量)
    vsAdvice.TextMatrix(lngRow, COL_总量单位) = "付"
    
    vsAdvice.TextMatrix(lngRow, COL_名称) = rs用法!名称
    vsAdvice.TextMatrix(lngRow, COL_用法) = rs用法!名称
    vsAdvice.TextMatrix(lngRow, COL_频率性质) = vsAdvice.TextMatrix(lngFirstRow, COL_频率性质)
    vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
    vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngFirstRow, COL_频率次数)
    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔)
    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
    vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
    
    '执行性质:缺省根据项目设置(不可能为院外执行),修改时根据当前界面设置
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = NVL(rs用法!执行科室, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = decode(NVL(rsCurr!执行性质), "离院带药", 5, NVL(rs用法!执行科室, 0))
    End If
    
    '中药用法如果未设置执行科室,则缺省为病人所在科室
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_执行性质))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID(rs用法!类别, lng用法ID, 0, NVL(rs用法!执行科室, 0), cbo期效.ListIndex, mint范围)
    End If
    
    If Not rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_医生嘱托) = NVL(rsCurr!医生嘱托)
    ElseIf Not rsUse.EOF Then
        vsAdvice.TextMatrix(lngRow, COL_医生嘱托) = NVL(rsUse!医生嘱托)
    End If
    
    '中药形态(用于AdviceTextMake中)
    vsAdvice.TextMatrix(lngRow, COL_中药形态) = lng形态
    
    '中药配方医嘱内容
    vsAdvice.TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
    
    '-------------------
    vsAdvice.Row = lngRow
    mblnRowChange = True
        
    AdviceSet中药配方 = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet检验组合(ByVal lngRow As Long, ByVal lng采集方法ID As Long, ByVal strExtData As String, Optional rsCurr As ADODB.Recordset) As Long
'功能：处理新增的检验(组合)
'参数：rsItems=输入或选择返回的记录集
'      lngRow=当前输入行
'      lng采集方法ID=缺省的采集方法
'      strExtData=检查:"'      检验组合="项目ID1,项目ID2,...;检验标本" 如果是新版LIS的模式则是："项目ID1|指标1|指标2...,项目ID2|指标1|指标2...,...;检验标本"
'      rsCurr=修改检验项目时用
'返回：处理之后的当前显示行号
    Dim rsMore As New ADODB.Recordset '采集方法信息
    Dim rsItems As New ADODB.Recordset '检验项目信息
    Dim arrItems As Variant, strItems As String
    Dim str医生 As String, lng医生ID As Long
    Dim str频率 As String, int频率次数 As Integer
    Dim int频率间隔 As Integer, str间隔单位 As String
    Dim lng相关ID As Long, str医嘱内容 As String
    Dim lngCopyRow As Long, lngFirstRow As Long
    Dim strSql As String, i As Long
    Dim rsLIS As New ADODB.Recordset
    Dim strTmp As String
    Dim Y As Long
    Dim blnLis As Boolean
    
    On Error GoTo errH
    
    '取上一或下一有效行,某些内容缺省与该行相同
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    
    '检验项目信息
    '----------------------------------------------------------------------------
    '各个检验项目信息:按输入顺序
    arrItems = Split(Split(strExtData, ";")(0), ",")
    For i = UBound(arrItems) To 0 Step -1
        If mblnNewLIS Then
            strTmp = arrItems(i)
            If InStr(strTmp, "|") > 0 Then
                For Y = 0 To UBound(Split(strTmp, "|"))
                    strItems = strItems & "," & Val(Split(strTmp, "|")(Y))
                    If Y > 0 Then
                        strSql = strSql & " Union All " & " Select '" & Val(Split(strTmp, "|")(Y)) & "' as 子项,'" & Val(Split(strTmp, "|")(0)) & "' as 父项 From Dual "
                    End If
                Next
            Else
                strItems = strItems & "," & Val(strTmp)
            End If
        Else
            strItems = strItems & "," & Val(arrItems(i))
        End If
    Next
    Set rsItems = Get诊疗项目记录(0, Mid(strItems, 2))
    If strSql <> "" Then
        Set rsLIS = zldatabase.OpenSQLRecord(Mid(strSql, 11), Me.Caption)
        blnLis = True
    End If
    
    '取某个检验项目的采集方法
    strSql = "Select A.项目ID,Nvl(A.性质,0) as 序号,A.用法ID" & _
        " From 诊疗用法用量 A,诊疗项目目录 B" & _
        " Where A.用法ID=B.ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([2],3)") & _
        " And A.项目ID IN(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
        " And (b.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or b.撤档时间 is NULL)" & _
        " Order by A.项目ID,Nvl(A.性质,0)"
    Set rsMore = zldatabase.OpenSQLRecord(strSql, Me.Caption, Mid(strItems, 2), mint范围)
    If Not rsMore.EOF Then
        If rsCurr Is Nothing Or lng采集方法ID = 0 Then
            lng采集方法ID = rsMore!用法ID '修改时不变
        End If
    End If
    Set rsMore = Get诊疗项目记录(lng采集方法ID)
    
    mblnRowChange = False
    
    '设置各行检验项目
    '----------------------------------------------------------------------------
    '采集方法医嘱ID,ID顺序与序号不一定一致
    If Not rsCurr Is Nothing Then
        '修改了检验组合中的内容,采集方法行标记为修改,医嘱ID不变
        lng相关ID = rsCurr!医嘱ID
    Else
        '新输入的中药配方
        lng相关ID = GetNextID
    End If
    
    With vsAdvice
        For i = 1 To rsItems.RecordCount
            .AddItem "", lngRow
            
            .RowHidden(lngRow) = True
            .RowData(lngRow) = GetNextID
            .TextMatrix(lngRow, COL_相关ID) = lng相关ID '对应到采集方法行
            .TextMatrix(lngRow, COL_期效) = zlCommFun.GetNeedName(cbo期效.Text)
            
            .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
            Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
            
            .TextMatrix(lngRow, COL_类别) = rsItems!类别
            .TextMatrix(lngRow, col_医嘱内容) = rsItems!名称
            .TextMatrix(lngRow, COL_诊疗项目ID) = rsItems!ID
            .TextMatrix(lngRow, COL_计算方式) = NVL(rsItems!计算方式, 0)
            If .TextMatrix(lngRow, COL_期效) = "临嘱" And NVL(rsItems!执行频率, 0) = 0 And mbln一次性 Then
                .TextMatrix(lngRow, COL_频率性质) = 1 '可选择频率的临嘱缺省为一次性
            Else
                .TextMatrix(lngRow, COL_频率性质) = NVL(rsItems!执行频率, 0)
            End If
            .TextMatrix(lngRow, COL_操作类型) = NVL(rsItems!操作类型)
            .TextMatrix(lngRow, COL_执行性质) = NVL(rsItems!执行科室, 0)
            '检验标本
            .TextMatrix(lngRow, COL_标本部位) = Split(strExtData, ";")(1)
            If mblnNewLIS And rsItems!ID & "" <> "" And blnLis Then
                rsLIS.Filter = "子项=" & rsItems!ID
                If rsLIS.EOF = False Then
                    .TextMatrix(lngRow, COL_组合项目ID) = rsLIS!父项 & ""
                End If
            End If
            
            '部份内容一并采集的检验项目相同
            If lngFirstRow <> 0 Then
                .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngFirstRow, COL_总量)
                
                '一并采集的检验项目应该相同
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = .TextMatrix(lngFirstRow, COL_执行科室ID)
                End If
                .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngFirstRow, COL_频率)
                .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngFirstRow, COL_频率次数)
                .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngFirstRow, COL_频率间隔)
                .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngFirstRow, COL_间隔单位)
                .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngFirstRow, COL_执行时间)
            ElseIf Not rsCurr Is Nothing Then
                If cbo期效.ListIndex = 1 Then
                    .TextMatrix(lngRow, COL_总量) = NVL(rsCurr!总量, 1)
                End If
                
                '执行科室:执行性质为(0-叮嘱,5-院外执行)无执行科室
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                    If NVL(rsCurr!执行科室ID, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = rsCurr!执行科室ID
                    Else
                        .TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID(rsItems!类别, rsItems!ID, 0, NVL(rsItems!执行科室, 0), cbo期效.ListIndex, mint范围)
                    End If
                End If
                
                '执行频率
                .TextMatrix(lngRow, COL_频率) = NVL(rsCurr!频率)
                .TextMatrix(lngRow, COL_频率次数) = NVL(rsCurr!频率次数)
                .TextMatrix(lngRow, COL_频率间隔) = NVL(rsCurr!频率间隔)
                .TextMatrix(lngRow, COL_间隔单位) = NVL(rsCurr!间隔单位)
                .TextMatrix(lngRow, COL_执行时间) = NVL(rsCurr!执行时间)
            Else
                If cbo期效.ListIndex = 1 Then
                    .TextMatrix(lngRow, COL_总量) = 1
                End If
                
                '执行科室:执行性质为(0-叮嘱,5-院外执行)无执行科室
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                    '之前要求出开嘱科室ID
                    .TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID(rsItems!类别, rsItems!ID, 0, NVL(rsItems!执行科室, 0), cbo期效.ListIndex, mint范围)
                End If
                
                '执行频率
                Call Get缺省频率(NVL(rsItems!ID, 0), Get频率范围(lngRow), str频率, int频率次数, int频率间隔, str间隔单位)
                .TextMatrix(lngRow, COL_频率) = str频率
                .TextMatrix(lngRow, COL_频率次数) = int频率次数
                .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
            
                '执行时间:"可选频率"(药品是可选频率,但可能设置为一次性)
                If Val(.TextMatrix(lngRow, COL_频率性质)) = 0 Then
                    If lngCopyRow <> -1 Then '与上一行相同
                        If .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngCopyRow, COL_频率) Then
                            .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngCopyRow, COL_执行时间)
                        End If
                    End If
                    If .TextMatrix(lngRow, COL_执行时间) = "" Then  '缺省时间方案
                        .TextMatrix(lngRow, COL_执行时间) = Get缺省时间(1, .TextMatrix(lngRow, COL_频率))
                    End If
                End If
            End If
            
            str医嘱内容 = str医嘱内容 & "," & rsItems!名称 '医嘱内容
            If lngFirstRow = 0 Then lngFirstRow = lngRow '第一项目行
            lngRow = lngRow + 1 '保持当前输入行位置
            
            rsItems.MoveNext
        Next
        
        '设置标本的采集方法
        '----------------------------------------------------------------------------
        rsItems.MoveFirst
        .RowData(lngRow) = lng相关ID
        
        .TextMatrix(lngRow, COL_期效) = zlCommFun.GetNeedName(cbo期效.Text)
        
        .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
        Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
        
        .TextMatrix(lngRow, COL_类别) = rsMore!类别
        .TextMatrix(lngRow, COL_名称) = rsMore!名称
        .TextMatrix(lngRow, COL_用法) = rsMore!名称
        .TextMatrix(lngRow, COL_诊疗项目ID) = rsMore!ID
        .TextMatrix(lngRow, COL_计算方式) = NVL(rsMore!计算方式, 0)
        .TextMatrix(lngRow, COL_操作类型) = NVL(rsMore!操作类型)
        .TextMatrix(lngRow, COL_标本部位) = .TextMatrix(lngFirstRow, COL_标本部位)
        
        '总量为检验项目的,与检验项目相同
        .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngFirstRow, COL_总量)
        If cbo期效.ListIndex = 1 Then
            .TextMatrix(lngRow, COL_总量单位) = NVL(rsMore!计算单位)
        End If
        
        '执行频率
        .TextMatrix(lngRow, COL_频率性质) = .TextMatrix(lngFirstRow, COL_频率性质) '以检验的为准
        .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngFirstRow, COL_频率)
        .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngFirstRow, COL_频率次数)
        .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngFirstRow, COL_频率间隔)
        .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngFirstRow, COL_间隔单位)
        .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngFirstRow, COL_执行时间)
        .TextMatrix(lngRow, COL_执行性质) = NVL(rsMore!执行科室, 0)
        
        '执行科室:执行性质为(0-叮嘱,5-院外执行)无执行科室
        If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
            .TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID(rsMore!类别, rsMore!ID, 0, NVL(rsMore!执行科室, 0), cbo期效.ListIndex, mint范围)
        End If
        
        If Not rsCurr Is Nothing Then
            .TextMatrix(lngRow, COL_医生嘱托) = NVL(rsCurr!医生嘱托)
        End If
        
        '医嘱内容:检验1,检验2(标本 采集方法)
        .TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
        
        .Row = lngRow
    End With
    mblnRowChange = True
    AdviceSet检验组合 = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdviceSet诊疗项目(rsInput As ADODB.Recordset, ByVal lngRow As Long, ByVal lng给药途径ID As Long, ByVal lngGroupRow As Long, ByVal strExtData As String, Optional ByVal bln备血 As Boolean = True)
'功能：处理新增(插入)的中、西成药，检查(组合)，手术(组合)，卫材，输血，及其它诊疗项目的缺省医嘱数据
'参数：rsInput=输入或选择返回的记录集
'      lngRow=当前输入行
'      lng给药途径ID=缺省给药途径ID,或一并给药时的给药途径ID
'      lngGroupRow=在一并给药的一组成药中插入新的成药行时,对应一并给药的一行行号
'      strExtData=检查:包含检查部位方法等信息,手术:包含附加手术及麻醉的信息,可能无附加手术
'      bln备血 当前的输血医嘱为备血医嘱，仅对类别为K的诊疗项目
    Dim rsTmp As New ADODB.Recordset
    Dim rsMore As New ADODB.Recordset '诊疗项目详细信息
    Dim strSql As String, lngCopyRow As Long
    Dim lngTmp As Long, i As Long
    Dim str医生 As String, lng医生ID As Long
    Dim str药房IDs As String, sng天数 As Single
    Dim str频率 As String, int频率次数 As Integer
    Dim int频率间隔 As Integer, str间隔单位 As String
    Dim lng收费项目ID As Long, bln品种 As Boolean
        
    On Error GoTo errH
    
    '取上一或下一有效行,某些内容缺省与该行相同
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
            
    With vsAdvice
        '开始设置医嘱缺省内容
        .RowData(lngRow) = GetNextID
        .TextMatrix(lngRow, COL_期效) = zlCommFun.GetNeedName(cbo期效.Text)
        
        '序号:保持连续,当前行占用新序号后,后面的序号向后移
        .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
        Call AdviceSet医嘱序号(lngRow + 1, 1)
        
        .TextMatrix(lngRow, COL_类别) = rsInput!类别ID
        .TextMatrix(lngRow, COL_名称) = rsInput!名称 '该名称可能是别名
        .TextMatrix(lngRow, COL_诊疗项目ID) = rsInput!诊疗项目ID
        
        '药品特性
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            strSql = "Select 毒理分类,药品剂型,品种医嘱,临床自管药,抗生素 From 药品特性 Where 药名ID=[1]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(rsInput!诊疗项目ID))
            If Not rsTmp.EOF Then
                .TextMatrix(lngRow, COL_毒理分类) = NVL(rsTmp!毒理分类)
                .TextMatrix(lngRow, COL_药品剂型) = NVL(rsTmp!药品剂型)
                .TextMatrix(lngRow, COL_临床自管药) = rsTmp!临床自管药 & ""
                .TextMatrix(lngRow, COL_抗菌等级) = Val("" & rsTmp!抗生素)
                If chkMedicineVariety.value = 1 Then
                    bln品种 = True
                Else
                    '是否长嘱药品固定按品种下达
                    bln品种 = NVL(rsTmp!品种医嘱, 0) <> 0 And cbo期效.ListIndex = 0
                End If
            End If
        End If
        
        If NVL(rsInput!类别ID) = "4" And mbyt场合 = 1 Then
            If chkMedicineVariety.value = 1 Then
                bln品种 = True
            Else
                bln品种 = False
            End If
        End If
        
        '是否长嘱药品固定按品种下达
        lng收费项目ID = NVL(rsInput!收费细目ID, 0)
        If bln品种 Then lng收费项目ID = 0
        
        '药品、卫材的规格信息
        .TextMatrix(lngRow, COL_收费细目ID) = lng收费项目ID
        If lng收费项目ID <> 0 Then
            If InStr(",5,6,", rsInput!类别ID) > 0 Then
                strSql = "Select Nvl(C.名称,A.名称) as 名称,B.剂量系数,B." & IIF(mint范围 = 1, "门诊", "住院") & "可否分零 As 可否分零," & _
                    decode(mint范围, 1, "B.门诊包装 as 包装系数,B.门诊单位 as 包装单位", 2, "B.住院包装 as 包装系数,B.住院单位 as 包装单位", "A.计算单位 as 包装单位,1 as 包装系数") & _
                    " From 收费项目目录 A,药品规格 B,收费项目别名 C" & _
                    " Where A.ID=B.药品ID And A.ID=[1]" & _
                    " And A.ID=C.收费细目ID(+) And C.码类(+)=1 And C.性质(+)=[2]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng收费项目ID, IIF(gbyt药品名称显示 = 0, 1, 3))
                .TextMatrix(lngRow, COL_名称) = rsTmp!名称 '将别名换成正式规格名称
                .TextMatrix(lngRow, COL_剂量系数) = rsTmp!剂量系数
                .TextMatrix(lngRow, COL_包装单位) = rsTmp!包装单位
                .TextMatrix(lngRow, COL_包装系数) = rsTmp!包装系数
                .TextMatrix(lngRow, COL_可否分零) = NVL(rsTmp!可否分零, 0)
            ElseIf rsInput!类别ID = "4" Then
                strSql = "Select A.跟踪在用,B.名称,B.计算单位 From 材料特性 A,收费项目目录 B Where A.材料ID=B.ID And A.材料ID=[1]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng收费项目ID)
                .TextMatrix(lngRow, COL_名称) = rsTmp!名称 '将别名换成正式规格名称
                .TextMatrix(lngRow, COL_剂量系数) = 1
                .TextMatrix(lngRow, COL_包装系数) = 1
                .TextMatrix(lngRow, COL_包装单位) = NVL(rsTmp!计算单位) '散装单位
                .TextMatrix(lngRow, COL_跟踪在用) = NVL(rsTmp!跟踪在用, 0)
            End If
        End If
        
        '获取更多诊疗项目信息
        '----------------------------------------------------------------------------
        If InStr(",5,6,", rsInput!类别ID) > 0 And lng收费项目ID <> 0 Then
            strSql = "Select a.用法id,a.频次,a.成人剂量,a.小儿剂量,a.医生嘱托,a.疗程,c.药名id as 项目ID " & _
                " From 药品用法用量 A,诊疗项目目录 B,药品规格 C " & _
                " Where A.用法ID=B.ID and a.药品ID=c.药品id And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([2],3)") & _
                " And A.药品ID=[3] And A.性质=1"
            strSql = "Select A.*,1 as 性质,B.用法ID," & _
                " B.频次,B.成人剂量,B.小儿剂量,B.医生嘱托,B.疗程" & _
                " From 诊疗项目目录 A,(" & strSql & ") B" & _
                " Where A.ID=b.项目id(+) And A.ID=[1]"
        Else
            strSql = "Select A.*" & _
                " From 诊疗用法用量 A,诊疗项目目录 B" & _
                " Where A.用法ID=B.ID And (Nvl(A.性质,0)=0 Or " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([2],3)") & ")" & _
                " And A.项目ID=[1]"
            strSql = "Select A.*,Nvl(B.性质,0) as 性质,B.用法ID," & _
                " B.频次,B.成人剂量,B.小儿剂量,B.医生嘱托,B.疗程" & _
                " From 诊疗项目目录 A,(" & strSql & ") B" & _
                " Where A.ID=B.项目ID(+) And A.ID=[1]" & _
                " Order by 性质"
        End If
        
        Set rsMore = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(rsInput!诊疗项目ID), mint范围, lng收费项目ID)
        
        If lng收费项目ID = 0 Then '将别名换成正式诊疗名称
            .TextMatrix(lngRow, COL_名称) = rsMore!名称
        End If
        
        '单量单位
        If rsInput!类别ID = "4" Then
            .TextMatrix(lngRow, COL_单量单位) = .TextMatrix(lngRow, COL_包装单位) '散装单位
        Else
            If cbo期效.ListIndex = 0 Then
                If InStr(",5,6,", rsInput!类别ID) > 0 Or InStr(",1,2,", NVL(rsMore!计算方式, 0)) > 0 Then
                    .TextMatrix(lngRow, COL_单量单位) = NVL(rsMore!计算单位) '药品为剂量单位
                End If
            Else
                If InStr(",5,6,", rsInput!类别ID) > 0 Or (NVL(rsMore!执行频率, 0) = 0 And InStr(",1,2,", NVL(rsMore!计算方式, 0)) > 0) Then
                    .TextMatrix(lngRow, COL_单量单位) = NVL(rsMore!计算单位) '药品为剂量单位
                End If
            End If
        End If
        
        If cbo期效.ListIndex = 1 Then
            If InStr(",5,6,", rsInput!类别ID) > 0 Then
                '中、西成药临嘱的总量单位就是包装单位
                .TextMatrix(lngRow, COL_总量单位) = .TextMatrix(lngRow, COL_包装单位)
            ElseIf rsInput!类别ID = "4" Then
                .TextMatrix(lngRow, COL_总量单位) = .TextMatrix(lngRow, COL_包装单位) '散装单位
            Else
                '其它临嘱要输入总量
                '如果为一次性或计次临嘱缺省总量为1
                If NVL(rsMore!执行频率, 0) = 1 Or NVL(rsMore!计算方式, 0) = 3 Then
                    .TextMatrix(lngRow, COL_总量) = 1
                End If
                .TextMatrix(lngRow, COL_总量单位) = NVL(rsMore!计算单位)
            End If
        End If
        
        .TextMatrix(lngRow, COL_计算方式) = NVL(rsMore!计算方式, 0)
        If .TextMatrix(lngRow, COL_期效) = "临嘱" And NVL(rsMore!执行频率, 0) = 0 And mbln一次性 Then
            .TextMatrix(lngRow, COL_频率性质) = 1 '可选择频率的临嘱缺省为一次性
        Else
            .TextMatrix(lngRow, COL_频率性质) = NVL(rsMore!执行频率, 0)
        End If
        .TextMatrix(lngRow, COL_操作类型) = NVL(rsMore!操作类型)
        
        '标本部位
        If InStr(",4,5,6,", rsInput!类别ID) > 0 Then
            .TextMatrix(lngRow, COL_标本部位) = rsInput!名称 '记录药品、卫材输入时选择名称
        ElseIf rsInput!类别ID <> "D" Then
            .TextMatrix(lngRow, COL_标本部位) = NVL(rsMore!标本部位)
        End If
        
        '执行性质:新增项目时根据项目设置,药品、卫材=4-指定科室,一并给药的相同
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            If lngGroupRow <> 0 Then
                .TextMatrix(lngRow, COL_执行性质) = .TextMatrix(lngGroupRow, COL_执行性质)
            Else
                .TextMatrix(lngRow, COL_执行性质) = 4
            End If
        ElseIf rsInput!类别ID = "4" Then
            .TextMatrix(lngRow, COL_执行性质) = 4
        Else
            .TextMatrix(lngRow, COL_执行性质) = NVL(rsMore!执行科室, 0)
        End If
        
        '执行科室:药品缺省与上一行相同,一并给药的相同
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            If lngGroupRow <> 0 Then
                str药房IDs = Get可用药房IDs(rsInput!类别ID, rsInput!诊疗项目ID, lng收费项目ID, 0, mint范围)
                If InStr("," & str药房IDs & ",", "," & .TextMatrix(lngGroupRow, COL_执行科室ID) & ",") > 0 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = .TextMatrix(lngGroupRow, COL_执行科室ID)
                End If
            ElseIf lngCopyRow <> -1 Then
                If rsInput!类别ID = .TextMatrix(lngCopyRow, COL_类别) Then
                    str药房IDs = Get可用药房IDs(rsInput!类别ID, rsInput!诊疗项目ID, lng收费项目ID, 0, mint范围)
                    If InStr("," & str药房IDs & ",", "," & .TextMatrix(lngCopyRow, COL_执行科室ID) & ",") > 0 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = .TextMatrix(lngCopyRow, COL_执行科室ID)
                    End If
                End If
            End If
        End If
        If Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
            If rsInput!类别ID = "Z" And (NVL(rsMore!操作类型, 0) = 3 Or NVL(rsMore!操作类型, 0) = 2 Or NVL(rsMore!操作类型, 0) = 1) Then
                '转科,入院，留观医嘱，缺省执行科室为空
            ElseIf rsInput!类别ID = "Z" And NVL(rsMore!操作类型, 0) = 7 Then
                '会诊医嘱
            ElseIf InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                '执行性质为(0-叮嘱,5-院外执行)无执行科室
                '先要求出开嘱科室ID
                .TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID(rsInput!类别ID, rsInput!诊疗项目ID, lng收费项目ID, NVL(rsMore!执行科室, 0), cbo期效.ListIndex, mint范围)
            End If
        End If
        
        '执行频率:可选频率,一次性或持续性
        If True Then 'If Nvl(rsMore!执行频率, 0) = 0 Then
            '缺省与上一新增行相同
            If lngCopyRow <> -1 Then
                If .TextMatrix(lngRow, COL_期效) = .TextMatrix(lngCopyRow, COL_期效) And Get频率范围(lngRow) = Get频率范围(lngCopyRow) Then
                    If .TextMatrix(lngCopyRow, COL_频率) <> "" _
                        And Not (.TextMatrix(lngRow, COL_类别) = "7" And Not RowIn配方行(lngCopyRow)) _
                        And Not (.TextMatrix(lngRow, COL_类别) <> "7" And RowIn配方行(lngCopyRow)) _
                        And Check频率可用(NVL(rsInput!诊疗项目ID, 0), Get频率范围(lngRow), .TextMatrix(lngCopyRow, COL_频率)) Then
                        .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngCopyRow, COL_频率)
                        .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngCopyRow, COL_频率次数)
                        .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngCopyRow, COL_频率间隔)
                        .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngCopyRow, COL_间隔单位)
                    End If
                End If
            End If
            '或取缺省频率
            If .TextMatrix(lngRow, COL_频率) = "" Then
                Call Get缺省频率(NVL(rsInput!诊疗项目ID, 0), Get频率范围(lngRow), str频率, int频率次数, int频率间隔, str间隔单位)
                .TextMatrix(lngRow, COL_频率) = str频率
                .TextMatrix(lngRow, COL_频率次数) = int频率次数
                .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
            End If
        End If
        
        '中，西成药的一些缺省信息
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            '执行频率
            If lngGroupRow <> 0 Then
                '一并给药的相同
                .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngGroupRow, COL_频率)
                .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngGroupRow, COL_频率次数)
                .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngGroupRow, COL_频率间隔)
                .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngGroupRow, COL_间隔单位)
                .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngGroupRow, COL_执行时间)
                '频率性质也要相同,如已强制设置为一次性
                .TextMatrix(lngRow, COL_频率性质) = .TextMatrix(lngGroupRow, COL_频率性质)
            End If
            
            '确定临嘱用药天数：
            '1.最少为一个频率周期天数
            '2-有疗程则为疗程天数(应大于一个频率周期天数)
            If cbo期效.ListIndex = 1 Then
                sng天数 = msng天数

                If .TextMatrix(lngRow, COL_间隔单位) = "周" Then
                    If 7 > sng天数 Then sng天数 = 7
                ElseIf .TextMatrix(lngRow, COL_间隔单位) = "天" Then
                    If Val(.TextMatrix(lngRow, COL_频率间隔)) > sng天数 Then
                        sng天数 = Val(.TextMatrix(lngRow, COL_频率间隔))
                    End If
                ElseIf .TextMatrix(lngRow, COL_间隔单位) = "小时" Then
                    If Val(.TextMatrix(lngRow, COL_频率间隔)) \ 24 > sng天数 Then
                        sng天数 = Val(.TextMatrix(lngRow, COL_频率间隔)) \ 24
                    End If
                ElseIf .TextMatrix(lngRow, COL_间隔单位) = "分钟" Then
                    If sng天数 = 0 Then sng天数 = 1
                End If
                If sng天数 = 0 Then sng天数 = 1
            End If

            rsMore.Filter = "性质>0" '取第一种给药途径用为缺省设置
            If Not rsMore.EOF Then
                '不是一并给药时,设置的缺省用法频率优先
                If lngGroupRow = 0 Then
                    If Not IsNull(rsMore!用法ID) Then lng给药途径ID = rsMore!用法ID
                    If Not IsNull(rsMore!频次) And Val(.TextMatrix(lngRow, COL_频率性质)) <> 1 Then '缺省为一次性优先
                        Call Get频率信息_编码(rsMore!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                        .TextMatrix(lngRow, COL_频率) = str频率
                        .TextMatrix(lngRow, COL_频率次数) = int频率次数
                        .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                        .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                    End If
                End If
                
                '医生嘱托
                .TextMatrix(lngRow, COL_医生嘱托) = NVL(rsMore!医生嘱托) '一般为给药途径的说明
                
                '药品单量
                If NVL(rsMore!成人剂量, 0) <> 0 Then
                    .TextMatrix(lngRow, COL_单量) = FormatEx(rsMore!成人剂量, 5)
                End If
                If Val(.TextMatrix(lngRow, COL_单量)) = 0 Then .TextMatrix(lngRow, COL_单量) = ""
                
                '药品临嘱总量:包装单位
                If cbo期效.ListIndex = 1 Then
                    If NVL(rsMore!疗程, 1) > sng天数 Then sng天数 = NVL(rsMore!疗程, 1)
                    If .TextMatrix(lngRow, COL_频率) <> "" And Val(.TextMatrix(lngRow, COL_单量)) <> 0 _
                        And Val(.TextMatrix(lngRow, COL_剂量系数)) <> 0 And Val(.TextMatrix(lngRow, COL_包装系数)) <> 0 Then
                        If Val(.TextMatrix(lngRow, COL_频率性质)) = 1 Then '临嘱药品可能缺省为一次性
                            '仅按疗程算改为按最少用药天数算
                            .TextMatrix(lngRow, COL_总量) = FormatEx(Calc缺省药品总量( _
                                    Val(.TextMatrix(lngRow, COL_单量)), 1, 1, 1, "天", "", _
                                    Val(.TextMatrix(lngRow, COL_剂量系数)), _
                                    Val(.TextMatrix(lngRow, COL_包装系数)), _
                                    Val(.TextMatrix(lngRow, COL_可否分零))), 5)
                        Else
                            '仅按疗程算改为按最少用药天数算
                            .TextMatrix(lngRow, COL_总量) = FormatEx(Calc缺省药品总量( _
                                    Val(.TextMatrix(lngRow, COL_单量)), sng天数, _
                                    Val(.TextMatrix(lngRow, COL_频率次数)), _
                                    Val(.TextMatrix(lngRow, COL_频率间隔)), _
                                    .TextMatrix(lngRow, COL_间隔单位), _
                                    .TextMatrix(lngRow, COL_执行时间), _
                                    Val(.TextMatrix(lngRow, COL_剂量系数)), _
                                    Val(.TextMatrix(lngRow, COL_包装系数)), _
                                    Val(.TextMatrix(lngRow, COL_可否分零))), 5)
                        End If
                    End If
                End If
            End If
            
            '记录缺省天数
            If cbo期效.ListIndex = 1 And Val(.TextMatrix(lngRow, COL_频率性质)) <> 1 Then
                .TextMatrix(lngRow, COL_天数) = IIF(sng天数 = 0, "", sng天数)
            End If
        End If
        
        If rsMore.Filter <> 0 Then rsMore.Filter = 0
        
        '执行时间:"可选频率"(药品是可选频率,但可能设置为一次性)
        If Val(.TextMatrix(lngRow, COL_频率性质)) = 0 Then
            If .TextMatrix(lngRow, COL_执行时间) = "" Then
                If lngCopyRow <> -1 Then '与上一行相同
                    If .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngCopyRow, COL_频率) Then
                        .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngCopyRow, COL_执行时间)
                    End If
                End If
                If .TextMatrix(lngRow, COL_执行时间) = "" Then  '缺省时间方案
                    .TextMatrix(lngRow, COL_执行时间) = Get缺省时间(1, .TextMatrix(lngRow, COL_频率), lng给药途径ID)
                End If
            End If
        End If
        
        '在主行处理完成之后处理附加行,并组合医嘱内容
        '-------------------------------------------------------------------------
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            '新增一个给药途径项目,并设置相关
            If lng给药途径ID <> 0 Then
                .TextMatrix(lngRow, COL_用法) = sys.RowValue("诊疗项目目录", lng给药途径ID, "名称")
            End If
            If lngGroupRow <> 0 Then
                '一并给药的关联相同的给药途径行
                lngTmp = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_相关ID)), lngGroupRow + 1)
                If lngTmp > lngRow Then
                    .TextMatrix(lngRow, COL_相关ID) = .TextMatrix(lngGroupRow, COL_相关ID)
                Else
                    '这种情况是仅为了使用一并给药的相同设置
                    .TextMatrix(lngRow, COL_相关ID) = AdviceSet给药途径(lngRow, lng给药途径ID)
                End If
            Else '独立新增的成药关联独立的给药途径行
                .TextMatrix(lngRow, COL_相关ID) = AdviceSet给药途径(lngRow, lng给药途径ID)
            End If
            
            '毒麻精的颜色标识
            If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", .TextMatrix(lngRow, COL_毒理分类)) > 0 _
                And .TextMatrix(lngRow, COL_毒理分类) <> "" Then
                .Cell(flexcpFontBold, lngRow, col_医嘱内容) = True
            End If
        ElseIf rsInput!类别ID = "D" And strExtData <> "" Then
            '检查的组合部位行
            Call AdviceSet检查组合(lngRow, strExtData)
        ElseIf rsInput!类别ID = "F" And strExtData <> "" Then
            '手术的附加手术及麻醉项目行
            Call AdviceSet手术组合(lngRow, strExtData)
        ElseIf rsInput!类别ID = "K" Then
            '输血的途径行
            If lng给药途径ID <> 0 Then
                If gbln血库系统 = True Then
                    strSQL = "Select a.名称,a.操作类型,a.执行分类 From 诊疗项目目录 A where a.id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng给药途径ID)
                    .TextMatrix(lngRow, COL_用法) = rsTmp!名称 & ""
                
                    If Val(rsTmp!操作类型 & "") = 8 And Val(rsTmp!执行分类 & "") = 1 Then '如果是编辑界面用申请单时需要重设一次
                        .TextMatrix(lngRow, COL_检查方法) = 1
                    Else
                        .TextMatrix(lngRow, COL_检查方法) = ""
                    End If
                Else
                    .TextMatrix(lngRow, COL_用法) = Sys.RowValue("诊疗项目目录", lng给药途径ID, "名称")
                End If
                Call AdviceSet输血途径(lngRow, lng给药途径ID)
            End If
        End If
        
        '医嘱内容
        .TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AdviceInputFree(ByVal lngRow As Long)
'功能：处理新增自由输入医嘱
    Dim str医生 As String, lng医生ID As Long
    Dim lngCopyRow As Long
    
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    
    With vsAdvice
        If .RowData(lngRow) <> 0 Then
            If txt医嘱内容.Text <> .TextMatrix(lngRow, col_医嘱内容) Then
                .TextMatrix(lngRow, col_医嘱内容) = txt医嘱内容.Text
                mblnNoSave = True '标记为未保存
            End If
        Else
            .RowData(lngRow) = GetNextID
            .TextMatrix(lngRow, COL_期效) = zlCommFun.GetNeedName(cbo期效.Text)
            
            '序号:保持连续,当前行占用新序号后,后面的序号向后移
            .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
            Call AdviceSet医嘱序号(lngRow + 1, 1)
                            
            .TextMatrix(lngRow, col_医嘱内容) = txt医嘱内容.Text
            .TextMatrix(lngRow, COL_类别) = "*" '特殊标记,为程序处理需要
            .TextMatrix(lngRow, COL_诊疗项目ID) = 0
            
            .TextMatrix(lngRow, COL_执行性质) = 4 '按可选执行科室处理，缺省为无
            .TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID("*", 0, 0, 4, cbo期效.ListIndex, mint范围)
            mblnNoSave = True '标记为未保存
            
            Call vsAdvice_AfterRowColChange(-1, -1, lngRow, .Col)
        End If
    End With
End Sub

Private Sub AdviceSet检查组合(ByVal lngRow As Long, ByVal strExData As String)
'功能：重新设置指定检查组合项目的部位方法行,用于新输入检查组合项目或修改部位方法
'参数：lngRow=当前输入行
'      strExData=包含检查部位方法等信息,格式为:"部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
    Dim arrItems As Variant, arrMethod As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim str检查部位 As String
    
    '删除现有的检查部位方法行
    Call Delete检查手术输血(lngRow)
    
    '重新加入部位方法行
    If strExData <> "" Then
        arrItems = Split(Split(strExData, vbTab)(0), "|")
        For i = 0 To UBound(arrItems)
            str检查部位 = Split(arrItems(i), ";")(0)
            arrMethod = Split(Split(arrItems(i), ";")(1), ",")
            For j = 0 To UBound(arrMethod)
                k = k + 1
                With vsAdvice
                    .AddItem "", lngRow + k
                    .RowHidden(lngRow + k) = True
                    
                    .RowData(lngRow + k) = GetNextID
                    .TextMatrix(lngRow + k, COL_相关ID) = .RowData(lngRow)
                    
                    .TextMatrix(lngRow + k, COL_序号) = Val(.TextMatrix(lngRow, COL_序号)) + k
                    .TextMatrix(lngRow + k, COL_期效) = .TextMatrix(lngRow, COL_期效)
                    
                    .TextMatrix(lngRow + k, COL_类别) = .TextMatrix(lngRow, COL_类别)
                    .TextMatrix(lngRow + k, COL_诊疗项目ID) = .TextMatrix(lngRow, COL_诊疗项目ID) '为同一个检查项目
                    
                    .TextMatrix(lngRow + k, COL_计算方式) = .TextMatrix(lngRow, COL_计算方式)
                    .TextMatrix(lngRow + k, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质)
                    .TextMatrix(lngRow + k, COL_操作类型) = .TextMatrix(lngRow, COL_操作类型)
                    
                    .TextMatrix(lngRow + k, col_医嘱内容) = .TextMatrix(lngRow, COL_名称) '记录为检查项目名称
                    .TextMatrix(lngRow + k, COL_标本部位) = str检查部位
                    .TextMatrix(lngRow + k, COL_检查方法) = arrMethod(j)
                    
                    .TextMatrix(lngRow + k, COL_单量) = .TextMatrix(lngRow, COL_单量)
                    .TextMatrix(lngRow + k, COL_总量) = .TextMatrix(lngRow, COL_总量)
                    
                    .TextMatrix(lngRow + k, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                    .TextMatrix(lngRow + k, COL_频率) = .TextMatrix(lngRow, COL_频率)
                    .TextMatrix(lngRow + k, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                    .TextMatrix(lngRow + k, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                    .TextMatrix(lngRow + k, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                    
                    .TextMatrix(lngRow + k, COL_执行性质) = .TextMatrix(lngRow, COL_执行性质)
                    .TextMatrix(lngRow + k, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                End With
            Next
        Next
                
        '调整后面医嘱的序号
        Call AdviceSet医嘱序号(lngRow + k + 1, k)
    End If
End Sub

Private Sub AdviceSet手术组合(ByVal lngRow As Long, ByVal strDataIDs As String)
'功能：重新设置指定手术项目的附加手术及麻醉项目行,用于新输入手术项目或手术项目的附加手术及麻醉项目
'参数：lngRow=当前输入行
'      strDataIDs=包含附加手术及麻醉项目信息,其中可能没有附加手术和麻醉
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '删除现有的附加手术及麻醉项目行
    Call Delete检查手术输血(lngRow)
    
    '重新加入附加手术行及麻醉项目行
    strDataIDs = Trim(Replace(strDataIDs, ";", ","))
    If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
    If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    
    If strDataIDs <> "" Then
        Set rsTmp = Get诊疗项目记录(0, strDataIDs)
        If Not rsTmp.EOF Then
            arrIDs = Split(strDataIDs, ",")
            For i = 0 To UBound(arrIDs) '按用户输入项目顺序
                rsTmp.Filter = "ID=" & CStr(arrIDs(i)) '不可能EOF
                
                With vsAdvice
                    .AddItem "", lngRow + i + 1
                    .RowHidden(lngRow + i + 1) = True
                    
                    .RowData(lngRow + i + 1) = GetNextID
                    .TextMatrix(lngRow + i + 1, COL_相关ID) = .RowData(lngRow)
                    
                    .TextMatrix(lngRow + i + 1, COL_序号) = Val(.TextMatrix(lngRow, COL_序号)) + i + 1
                    .TextMatrix(lngRow + i + 1, COL_期效) = .TextMatrix(lngRow, COL_期效)
                    
                    .TextMatrix(lngRow + i + 1, COL_类别) = rsTmp!类别
                    .TextMatrix(lngRow + i + 1, COL_诊疗项目ID) = rsTmp!ID
                    .TextMatrix(lngRow + i + 1, COL_计算方式) = NVL(rsTmp!计算方式, 0)
                    .TextMatrix(lngRow + i + 1, COL_频率性质) = NVL(rsTmp!执行频率, 0)
                    .TextMatrix(lngRow + i + 1, COL_操作类型) = NVL(rsTmp!操作类型)
                    
                    .TextMatrix(lngRow + i + 1, COL_标本部位) = NVL(rsTmp!标本部位)
                    .TextMatrix(lngRow + i + 1, col_医嘱内容) = rsTmp!名称
                    
                    .TextMatrix(lngRow + i + 1, COL_单量) = .TextMatrix(lngRow, COL_单量)
                    .TextMatrix(lngRow + i + 1, COL_总量) = .TextMatrix(lngRow, COL_总量)
                    
                    .TextMatrix(lngRow + i + 1, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                    .TextMatrix(lngRow + i + 1, COL_频率) = .TextMatrix(lngRow, COL_频率)
                    .TextMatrix(lngRow + i + 1, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                    .TextMatrix(lngRow + i + 1, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                    .TextMatrix(lngRow + i + 1, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                    
                    '执行性质:根据项目自身设置
                    .TextMatrix(lngRow + i + 1, COL_执行性质) = NVL(rsTmp!执行科室, 0)
                    
                    '叮嘱和院外执行无执行科室,手术麻醉单独执行科室
                    '否则不管其执行科室设置,一个手术组合应该相同
                    If InStr(",0,5,", NVL(rsTmp!执行科室, 0)) > 0 Then
                        .TextMatrix(lngRow + i + 1, COL_执行科室ID) = 0
                    Else
                        If rsTmp!类别 = "G" Then
                            .TextMatrix(lngRow + i + 1, COL_执行科室ID) = Get成套执行科室ID(rsTmp!类别, rsTmp!ID, 0, NVL(rsTmp!执行科室, 0), IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1), mint范围)
                        Else
                            .TextMatrix(lngRow + i + 1, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                        End If
                    End If
                End With
            Next
                
            '调整序号
            Call AdviceSet医嘱序号(lngRow + UBound(arrIDs) + 2, UBound(arrIDs) + 1)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSet给药途径(ByVal lngRow As Long, ByVal lng给药途径ID As Long, Optional str执行性质 As String, Optional ByVal str滴速 As String) As Long
'功能：为录入的中，西成药设置对应的给药途径行(新增或修改)
'参数：lngRow=要处理给药途径的药品行
'      lng给药途径ID=给药途径ID
'      str执行性质=修改给药途径时,当前界面设置的执行性质
'      str滴速=修改给药途径时,当前界面设置的滴速
'返回：被设置的给药途径行的医嘱ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lngNewRow As Long
    Dim blnNew As Boolean
    
    On Error GoTo errH
    Set rsTmp = Get诊疗项目记录(lng给药途径ID)
    If rsTmp.EOF Then lng给药途径ID = 0 '没有数据，先设置以保持关系
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then '未设置"相关ID"时
            blnNew = True
            lngNewRow = lngRow + 1
            .AddItem "", lngNewRow
            .RowHidden(lngNewRow) = True
        Else
            '修改医嘱的内容时重新设置给药途径内容(不是更换诊疗项目)
            blnNew = False
            lngNewRow = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
        End If
        
        '无效内容：名称,收费细目ID,剂量系数,包装单位,包装系数,标本部位,医生嘱托,单量,总量,用法
        If blnNew Then
            .RowData(lngNewRow) = GetNextID
            .TextMatrix(lngNewRow, COL_序号) = Val(.TextMatrix(lngRow, COL_序号)) + 1
            .TextMatrix(lngNewRow, col_缺省) = .TextMatrix(lngRow, col_缺省)
        End If
        
        .TextMatrix(lngNewRow, COL_期效) = .TextMatrix(lngRow, COL_期效)
        
        .TextMatrix(lngNewRow, COL_类别) = "E" '给药途径属于治疗
        .TextMatrix(lngNewRow, COL_诊疗项目ID) = lng给药途径ID
        
        '如果没有确定给药途径，暂时不设置的内容
        If Not rsTmp.EOF Then
            .TextMatrix(lngNewRow, COL_计算方式) = NVL(rsTmp!计算方式, 0)
            .TextMatrix(lngNewRow, COL_操作类型) = NVL(rsTmp!操作类型)
            .TextMatrix(lngNewRow, COL_执行分类) = NVL(rsTmp!执行分类, 0)
            .TextMatrix(lngNewRow, col_医嘱内容) = rsTmp!名称
            
            '滴速
            If str滴速 <> "" Then
                .TextMatrix(lngNewRow, COL_医生嘱托) = str滴速
            End If
            '执行性质:缺省根据项目设置,修改时根据当前界面设置
            If str执行性质 = "" Then
                .TextMatrix(lngNewRow, COL_执行性质) = NVL(rsTmp!执行科室, 0)
            Else
                .TextMatrix(lngNewRow, COL_执行性质) = decode(str执行性质, "离院带药", 5, NVL(rsTmp!执行科室, 0))
            End If
            
            If InStr(",0,5,", Val(.TextMatrix(lngNewRow, COL_执行性质))) = 0 Then
                .TextMatrix(lngNewRow, COL_执行科室ID) = Get成套执行科室ID("E", lng给药途径ID, 0, NVL(rsTmp!执行科室, 0), IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1), mint范围)
            Else
                .TextMatrix(lngNewRow, COL_执行科室ID) = 0
            End If
        End If
        
        .TextMatrix(lngNewRow, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质) '以药品的为准
        .TextMatrix(lngNewRow, COL_频率) = .TextMatrix(lngRow, COL_频率)
        .TextMatrix(lngNewRow, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
        .TextMatrix(lngNewRow, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
        .TextMatrix(lngNewRow, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
        .TextMatrix(lngNewRow, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
        
        '往后调整序号
        If blnNew Then Call AdviceSet医嘱序号(lngNewRow + 1, 1)
        
        AdviceSet给药途径 = .RowData(lngNewRow)
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet输血途径(ByVal lngRow As Long, ByVal lng输血途径ID As Long) As Long
'功能：为录入的中，西成药设置对应的给药途径行(新增或修改)
'参数：lngRow=要处理输血途径的输血医嘱行
'      lng输血途径ID=输血途径ID
'返回：被设置的输血途径行的医嘱ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lngNewRow As Long
    Dim blnNew As Boolean
    
    On Error GoTo errH
    Set rsTmp = Get诊疗项目记录(lng输血途径ID)
    
    With vsAdvice
        lngNewRow = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
        If lngNewRow = -1 Then '尚未设置输血途径时
            blnNew = True
            lngNewRow = lngRow + 1
            .AddItem "", lngNewRow
            .RowHidden(lngNewRow) = True
        End If
        
        '无效内容：名称,收费细目ID,剂量系数,包装单位,包装系数,标本部位,医生嘱托,单量,总量,用法
        If blnNew Then
            .RowData(lngNewRow) = GetNextID
            .TextMatrix(lngNewRow, COL_相关ID) = .RowData(lngRow)
            .TextMatrix(lngNewRow, COL_序号) = Val(.TextMatrix(lngRow, COL_序号)) + 1
            .TextMatrix(lngNewRow, col_缺省) = .TextMatrix(lngRow, col_缺省)
        End If
        
        .TextMatrix(lngNewRow, COL_期效) = .TextMatrix(lngRow, COL_期效)
        
        .TextMatrix(lngNewRow, COL_类别) = "E" '输血途径属于治疗
        .TextMatrix(lngNewRow, COL_诊疗项目ID) = lng输血途径ID
        
        .TextMatrix(lngNewRow, COL_计算方式) = NVL(rsTmp!计算方式, 0)
        .TextMatrix(lngNewRow, COL_操作类型) = NVL(rsTmp!操作类型)
        .TextMatrix(lngNewRow, col_医嘱内容) = rsTmp!名称
        .TextMatrix(lngNewRow, COL_执行性质) = NVL(rsTmp!执行科室, 0)
        
        If InStr(",0,5,", Val(.TextMatrix(lngNewRow, COL_执行性质))) = 0 Then
            .TextMatrix(lngNewRow, COL_执行科室ID) = Get成套执行科室ID("E", lng输血途径ID, 0, NVL(rsTmp!执行科室, 0), IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", 0, 1), mint范围)
        Else
            .TextMatrix(lngNewRow, COL_执行科室ID) = 0
        End If
        
        .TextMatrix(lngNewRow, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质) '以药品的为准
        .TextMatrix(lngNewRow, COL_频率) = .TextMatrix(lngRow, COL_频率)
        .TextMatrix(lngNewRow, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
        .TextMatrix(lngNewRow, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
        .TextMatrix(lngNewRow, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
        .TextMatrix(lngNewRow, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
        
        '往后调整序号
        If blnNew Then Call AdviceSet医嘱序号(lngNewRow + 1, 1)
        
        AdviceSet输血途径 = .RowData(lngNewRow)
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceChange()
'功能：根据当前医嘱卡片中的内容，更新当前医嘱内容
'说明：对于ListIndex=-1而对应医嘱项又有内容的，保持原内容不更新
    Dim lngRow As Long, lngBeginRow As Long
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim blnCurDo As Boolean, blnTmp As Boolean
    Dim lngTmp As Long, strTmp As String
    Dim lng执行科室ID As Long, lng开嘱科室ID As Long
    Dim blnReInRow As Boolean, i As Long, j As Long

    With vsAdvice
        lngRow = .Row

        If .RowData(lngRow) = 0 Then Call ClearItemTag: Exit Sub    '清除编辑标志

        If RowIn配方行(lngRow) Then
            '中药配方
            lngBeginRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            For i = lngBeginRow To lngRow
                '修改处理配方的所有行内容(包括煎法和用法)
                If txt总量.Enabled And IsNumeric(txt总量.Text) And txt总量.Tag <> "" Then
                    .TextMatrix(i, COL_总量) = FormatEx(Val(txt总量.Text), 5)
                    blnCurDo = True
                End If
                If txt频率.Enabled And cmd频率.Tag <> "" And txt频率.Tag <> "" Then
                    .TextMatrix(i, COL_频率) = txt频率.Text
                    Call Get频率信息_名称(txt频率.Text, int频率次数, int频率间隔, str间隔单位, 2)    '中医范围
                    .TextMatrix(i, COL_频率次数) = int频率次数
                    .TextMatrix(i, COL_频率间隔) = int频率间隔
                    .TextMatrix(i, COL_间隔单位) = str间隔单位
                    blnCurDo = True
                End If
                If cbo执行时间.Tag <> "" Then
                    .TextMatrix(i, COL_执行时间) = cbo执行时间.Text
                    blnCurDo = True
                End If

                '适用证候
                If txt适用证候.Tag <> "" Then
                    .TextMatrix(i, COL_组合项目ID) = txt适用证候.Tag
                    .TextMatrix(i, COL_适用证候) = txt适用证候.Text
                    blnCurDo = True
                End If

                If .TextMatrix(i, COL_类别) = "7" Then
                    '更改的是组成中药的执行科室(用法煎法的改不到)
                    If cbo执行科室.Tag <> "" Then
                        If cbo执行科室.ListIndex <> -1 Then
                            .TextMatrix(i, COL_执行科室ID) = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                        Else
                            .TextMatrix(i, COL_执行科室ID) = ""
                        End If
                        blnCurDo = True
                    End If

                    '执行性质:配方中所有组成的中药相同
                    If cbo执行性质.Tag <> "" Then
                        .TextMatrix(i, COL_执行性质) = decode(zlCommFun.GetNeedName(cbo执行性质.Text), "自备药", 5, "不取药", 5, 4)
                        .TextMatrix(i, COL_执行标记) = decode(zlCommFun.GetNeedName(cbo执行性质.Text), "自取药", 1, "不取药", 2, 0)
                        If Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                            .TextMatrix(i, COL_执行科室ID) = 0
                        ElseIf Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                            '恢复缺省执行科室,缺省与前面相同
                            If i = lngBeginRow Then
                                For j = i - 1 To .FixedRows Step -1
                                    If .TextMatrix(j, COL_类别) = "7" And Val(.TextMatrix(j, COL_执行科室ID)) <> 0 Then
                                        .TextMatrix(i, COL_执行科室ID) = .TextMatrix(j, COL_执行科室ID)
                                        Exit For
                                    End If
                                Next
                                If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                                    .TextMatrix(i, COL_执行科室ID) = Get成套执行科室ID(.TextMatrix(i, COL_类别), Val(.TextMatrix(i, COL_诊疗项目ID)), Val(.TextMatrix(i, COL_收费细目ID)), 4, cbo期效.ListIndex, mint范围)
                                End If
                            Else
                                .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngBeginRow, COL_执行科室ID)
                            End If
                        End If
                        blnReInRow = True    '界面执行科室编辑性变化
                        blnCurDo = True
                    End If
                End If

                '修改时自动更新部份内容
                blnTmp = False
                If cbo医生嘱托.Tag <> "" Or cbo执行性质.Tag <> "" _
                   Or (Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "") Then
                    blnTmp = True
                End If

                If .TextMatrix(i, COL_类别) = "E" And i <> lngRow Then lngTmp = i    '煎法行号

                '---------------
                If blnCurDo Then mblnNoSave = True    '标记为未保存
            Next

            '涉及中药用法行的内容:直接更改当前行的内容(煎法行在配方编辑中才能改)
            '-----------------------------------------------------------
            blnCurDo = False

            '医生嘱托:是放在中药用法行(显示行)中的
            If cbo医生嘱托.Tag <> "" Then
                .TextMatrix(lngRow, COL_医生嘱托) = cbo医生嘱托.Text
                blnCurDo = True
            End If

            '中药用法
            If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                .TextMatrix(lngRow, COL_诊疗项目ID) = Val(cmd用法.Tag)
                .TextMatrix(lngRow, COL_用法) = txt用法.Text

                '同时更改执行性质
                i = NVL(sys.RowValue("诊疗项目目录", Val(cmd用法.Tag), "执行科室"), 0)
                .TextMatrix(lngRow, COL_执行性质) = decode(zlCommFun.GetNeedName(cbo执行性质.Text), "离院带药", 5, i)
                If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = 0
                Else
                    .TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID("E", Val(cmd用法.Tag), 0, Val(.TextMatrix(lngRow, COL_执行性质)), cbo期效.ListIndex, mint范围)
                End If

                blnReInRow = True    '需要刷新中药用法执行科室
                blnCurDo = True
            End If

            '用法和煎法的执行性质
            If cbo执行性质.Tag <> "" Then
                '用法
                i = NVL(sys.RowValue("诊疗项目目录", Val(.TextMatrix(lngRow, COL_诊疗项目ID)), "执行科室"), 0)
                .TextMatrix(lngRow, COL_执行性质) = decode(zlCommFun.GetNeedName(cbo执行性质.Text), "离院带药", 5, i)
                If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = 0
                Else
                    .TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID(.TextMatrix(lngRow, COL_类别), Val(.TextMatrix(lngRow, COL_诊疗项目ID)), 0, Val(.TextMatrix(lngRow, COL_执行性质)), cbo期效.ListIndex, mint范围)
                End If

                '煎法
                i = NVL(sys.RowValue("诊疗项目目录", Val(.TextMatrix(lngTmp, COL_诊疗项目ID)), "执行科室"), 0)
                .TextMatrix(lngTmp, COL_执行性质) = decode(zlCommFun.GetNeedName(cbo执行性质.Text), "离院带药", 5, i)
                If Val(.TextMatrix(lngTmp, COL_执行性质)) = 5 Then
                    .TextMatrix(lngTmp, COL_执行科室ID) = 0
                Else
                    .TextMatrix(lngTmp, COL_执行科室ID) = Get成套执行科室ID(.TextMatrix(lngTmp, COL_类别), Val(.TextMatrix(lngTmp, COL_诊疗项目ID)), 0, Val(.TextMatrix(lngTmp, COL_执行性质)), cbo期效.ListIndex, mint范围)
                End If

                mblnNoSave = True    '标记为未保存

                blnCurDo = True
            End If

            '中药用法执行科室:即配方当前显示行的执行科室
            If cbo附加执行.Tag <> "" Then
                If cbo附加执行.ListIndex <> -1 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = cbo附加执行.ItemData(cbo附加执行.ListIndex)
                Else
                    .TextMatrix(lngRow, COL_执行科室ID) = ""
                End If
                blnCurDo = True
            End If

            '---------------
            If blnCurDo Then mblnNoSave = True    '标记为未保存
        Else    '其它诊疗项目
            If txt单量.Enabled And (IsNumeric(txt单量.Text) Or txt单量.Text = "") And txt单量.Tag <> "" Then
                .TextMatrix(lngRow, COL_单量) = FormatEx(txt单量.Text, 5)
                blnCurDo = True
            End If

            If txt天数.Tag <> "" Then
                .TextMatrix(lngRow, COL_天数) = txt天数.Text
                blnCurDo = True
            End If

            If txt总量.Enabled And (IsNumeric(txt总量.Text) Or txt总量.Text = "") And txt总量.Tag <> "" Then
                .TextMatrix(lngRow, COL_总量) = FormatEx(txt总量.Text, 5)
                blnCurDo = True
            End If

            If txt频率.Enabled And cmd频率.Tag <> "" And txt频率.Tag <> "" Then
                '频率性质已经在设置时确定(临嘱可能在一次性之间切换)
                .TextMatrix(lngRow, COL_频率) = txt频率.Text
                Call Get频率信息_名称(txt频率.Text, int频率次数, int频率间隔, str间隔单位, Get频率范围(lngRow))
                .TextMatrix(lngRow, COL_频率次数) = int频率次数
                .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                blnCurDo = True
            End If

            If cbo执行时间.Tag <> "" Then
                .TextMatrix(lngRow, COL_执行时间) = cbo执行时间.Text
                blnCurDo = True
            End If
            If cbo医生嘱托.Tag <> "" Then
                .TextMatrix(lngRow, COL_医生嘱托) = cbo医生嘱托.Text
                blnCurDo = True
            End If

            If cbo执行科室.Tag <> "" Then
                If Not RowIn检验行(lngRow) Then    '采集方法的执行科室不同
                    If cbo执行科室.ListIndex <> -1 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                    Else
                        .TextMatrix(lngRow, COL_执行科室ID) = ""
                    End If
                End If
                blnCurDo = True
            End If
            
            '滴速：输液药品
            If cbo滴速.Tag <> "" Then
                lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                If lngTmp <> -1 Then
                    If cbo滴速.Text <> "" Then
                        .TextMatrix(lngTmp, COL_医生嘱托) = cbo滴速.Text & lbl滴速单位.Caption
                    Else
                        .TextMatrix(lngTmp, COL_医生嘱托) = ""
                    End If
                    blnCurDo = True
                End If
                If cbo滴速.Text <> "" Then
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text & cbo滴速.Text & lbl滴速单位.Caption
                Else
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text
                End If
            End If
            
            '附加执行科室：给药途径,手术麻醉,采集方法
            If cbo附加执行.Tag <> "" Then
                lngTmp = -1
                If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                ElseIf .TextMatrix(lngRow, COL_类别) = "F" Then
                    For i = lngRow + 1 To .Rows - 1
                        If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                            If .TextMatrix(i, COL_类别) = "G" Then
                                lngTmp = i: Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf .TextMatrix(lngRow, COL_类别) = "E" _
                       And .TextMatrix(lngRow - 1, COL_类别) = "C" _
                       And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
                    lngTmp = lngRow
                ElseIf .TextMatrix(lngRow, COL_类别) = "K" _
                    And .TextMatrix(lngRow + 1, COL_类别) = "E" _
                    And Val(.TextMatrix(lngRow + 1, COL_相关ID)) = .RowData(lngRow) Then
                    lngTmp = lngRow + 1
                End If

                '只更新对应行,不影响其它行
                If lngTmp <> -1 Then
                    If cbo附加执行.ListIndex <> -1 Then
                        .TextMatrix(lngTmp, COL_执行科室ID) = cbo附加执行.ItemData(cbo附加执行.ListIndex)
                    Else
                        .TextMatrix(lngTmp, COL_执行科室ID) = ""
                    End If
                    mblnNoSave = True    '标记为未保存
                End If
            End If

            '执行性质,给药途径:为更新开嘱时间(包括给药途径的同步更改),先判断是否改变
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                If cbo执行性质.Tag <> "" Then blnCurDo = True
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then blnCurDo = True
            End If

            '修改时自动更新部份内容
            blnTmp = False
            If cbo执行性质.Tag <> "" Or (Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "") Then
                blnReInRow = True    '需要刷新给药途径,采集方式的执行科室
                blnTmp = True
            End If

            '其它需要同步处理的关联行
            '----------------------------------------------------------------
            If RowIn检验行(lngRow) Then
                '采集方法
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                    .TextMatrix(lngRow, COL_诊疗项目ID) = Val(cmd用法.Tag)
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text
                    .TextMatrix(lngRow, COL_名称) = txt用法.Text

                    '同时更改执行性质
                    .TextMatrix(lngRow, COL_执行性质) = NVL(sys.RowValue("诊疗项目目录", Val(cmd用法.Tag), "执行科室"), 0)
                    If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID("E", Val(cmd用法.Tag), 0, Val(.TextMatrix(lngRow, COL_执行性质)), cbo期效.ListIndex, mint范围)
                    Else
                        .TextMatrix(lngRow, COL_执行科室ID) = 0
                    End If
                    
                    blnCurDo = True
                End If

                '设置一并采集的各个检验项目
                If blnCurDo Then
                    For i = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                            If txt总量.Tag <> "" Then
                                .TextMatrix(i, COL_总量) = .TextMatrix(lngRow, COL_总量)
                            End If
                            If txt频率.Tag <> "" Then
                                .TextMatrix(i, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质)
                                .TextMatrix(i, COL_频率) = .TextMatrix(lngRow, COL_频率)
                                .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                                .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                                .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                            End If
                            If cbo执行科室.Tag <> "" Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Or cbo执行科室.ListIndex = -1 Then
                                    .TextMatrix(i, COL_执行科室ID) = 0
                                Else
                                    .TextMatrix(i, COL_执行科室ID) = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                                End If
                                
                            End If
                            If cbo执行时间.Tag <> "" Then
                                .TextMatrix(i, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                                
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '中、西成药处理给药途径及一并给药的情况

                '执行性质
                If cbo执行性质.Tag <> "" Then
                    .TextMatrix(lngRow, COL_执行性质) = decode(zlCommFun.GetNeedName(cbo执行性质.Text), "自备药", 5, "不取药", 5, 4)
                    .TextMatrix(lngRow, COL_执行标记) = decode(zlCommFun.GetNeedName(cbo执行性质.Text), "自取药", 1, "不取药", 2, 0)
                    If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = 0
                    ElseIf Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
                        '恢复缺省药房,缺省与前面的成药相同
                        strTmp = Get可用药房IDs(.TextMatrix(lngRow, COL_类别), Val(.TextMatrix(lngRow, COL_诊疗项目ID)), Val(.TextMatrix(lngRow, COL_收费细目ID)), 0, mint范围)
                        For i = lngRow - 1 To .FixedRows Step -1
                            '西成药和中成药的药房可能不同,所以类别要相同
                            If .TextMatrix(i, COL_类别) = .TextMatrix(lngRow, COL_类别) And Val(.TextMatrix(i, COL_执行科室ID)) <> 0 Then
                                If InStr("," & strTmp & ",", "," & Val(.TextMatrix(i, COL_执行科室ID)) & ",") > 0 Then
                                    .TextMatrix(lngRow, COL_执行科室ID) = Val(.TextMatrix(i, COL_执行科室ID))
                                    Exit For
                                End If
                            End If
                        Next
                        If Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
                            .TextMatrix(lngRow, COL_执行科室ID) = Get成套执行科室ID(.TextMatrix(lngRow, COL_类别), Val(.TextMatrix(lngRow, COL_诊疗项目ID)), Val(.TextMatrix(lngRow, COL_收费细目ID)), 4, cbo期效.ListIndex, mint范围)
                        End If
                    End If

                    cbo执行科室.Tag = "1"    '标明执行科室一并给药的要同步变
                    blnReInRow = True    '界面执行科室编辑性变化
                End If

                '给药途径本身及其它相关数据同步更改
                strTmp = ""
                If Trim(cbo滴速.Text) <> "" Then
                    strTmp = cbo滴速.Text & lbl滴速单位.Caption
                End If
                
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text & strTmp
                    Call AdviceSet给药途径(lngRow, Val(cmd用法.Tag), zlCommFun.GetNeedName(cbo执行性质.Text), strTmp)
                ElseIf blnCurDo Then    'cbo执行性质.Tag <> "" Then
                    '如果执行性质更改了,需要强行修改对应的给药途径的执行性质和执行科室
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    Call AdviceSet给药途径(lngRow, Val(.TextMatrix(lngTmp, COL_诊疗项目ID)), zlCommFun.GetNeedName(cbo执行性质.Text), strTmp)
                End If

                '一并给药:不处理给药途径,前面已单独设置
                If blnCurDo Then
                    lngBeginRow = .FindRow(.TextMatrix(lngRow, COL_相关ID), , COL_相关ID)
                    For i = lngBeginRow To .Rows - 1
                        If i <> lngRow And .RowData(i) <> 0 Then    '可能现在中间有空行
                            If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                                If txt用法.Tag <> "" Then
                                    .TextMatrix(i, COL_用法) = .TextMatrix(lngRow, COL_用法)
                                    '将滴速补填到一并给药的其他行中
                                    If cbo滴速.Tag <> "" Then
                                        .TextMatrix(i, COL_用法) = txt用法.Text & strTmp
                                    End If
                                End If
                                
                                
                                If txt频率.Tag <> "" Then
                                    .TextMatrix(i, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质)    '需要同步设置,因为临嘱可能在一次性之间切换
                                    .TextMatrix(i, COL_频率) = .TextMatrix(lngRow, COL_频率)
                                    .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                                    .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                                    .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                                End If

                                '一并给药的,天数相同变化,总量重新计算
                                If txt天数.Tag <> "" Then
                                    .TextMatrix(i, COL_天数) = .TextMatrix(lngRow, COL_天数)
                                    If txt天数.Text <> "" And .TextMatrix(i, COL_频率) <> "" _
                                       And Val(.TextMatrix(i, COL_频率性质)) <> 1 And Val(.TextMatrix(i, COL_单量)) <> 0 _
                                       And Val(.TextMatrix(i, COL_剂量系数)) <> 0 And Val(.TextMatrix(i, COL_包装系数)) <> 0 Then

                                        .TextMatrix(i, COL_总量) = FormatEx(Calc缺省药品总量( _
                                                                          Val(.TextMatrix(i, COL_单量)), Val(.TextMatrix(i, COL_天数)), _
                                                                          Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), _
                                                                          .TextMatrix(i, COL_间隔单位), .TextMatrix(i, COL_执行时间), _
                                                                          Val(.TextMatrix(i, COL_剂量系数)), Val(.TextMatrix(i, COL_包装系数)), _
                                                                          Val(.TextMatrix(i, COL_可否分零))), 5)
                                    End If
                                End If

                                If cbo执行时间.Tag <> "" Then
                                    .TextMatrix(i, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                                End If

                                '执行性质、执行标记:离院带药、自取药在一并给药中需一致，其它可单独设置
                                If cbo执行性质.Tag <> "" And zlCommFun.GetNeedName(cbo执行性质.Text) = "离院带药" Or zlCommFun.GetNeedName(cbo执行性质.Text) = "自取药" Then
                                    .TextMatrix(i, COL_执行性质) = .TextMatrix(lngRow, COL_执行性质)
                                    .TextMatrix(i, COL_执行标记) = .TextMatrix(lngRow, COL_执行标记)
                                    '由自备药转过来时需要重新设置执行科室
                                    If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                                        .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                                    End If
                                End If

                                '执行科室:执行科室(药房)可以不同,除非是配制中心
                                If cbo执行科室.Tag <> "" Then
                                    '输入行改为自备药，或某行为自备药的情况不管它
                                    If Not (Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 And Val(.TextMatrix(lngRow, COL_执行性质)) = 5) _
                                       And Not (Val(.TextMatrix(i, COL_执行科室ID)) = 0 And Val(.TextMatrix(i, COL_执行性质)) = 5) Then
                                        If sys.DeptHaveProperty(Val(.TextMatrix(lngRow, COL_执行科室ID)), "配制中心") Then
                                            '输入行药品由普通药房或其他配制中心改为新的配制中心,则该组药都改为该配制中心
                                            .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                                        ElseIf sys.DeptHaveProperty(Val(.TextMatrix(i, COL_执行科室ID)), "配制中心") Then
                                            '输入行药品由配制中心改成普通药房,则该组药都改为该普通药房
                                            .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                                            
                                        End If
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        End If
                    Next
                End If
            ElseIf .TextMatrix(lngRow, COL_类别) = "K" Then
                '输血医嘱的处理
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text
                    Call AdviceSet输血途径(lngRow, Val(cmd用法.Tag))
                ElseIf blnCurDo Then
                    lngTmp = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
                    If lngTmp <> -1 Then
                        Call AdviceSet输血途径(lngRow, Val(.TextMatrix(lngTmp, COL_诊疗项目ID)))
                    End If
                End If
            ElseIf InStr(",D,F,", .TextMatrix(lngRow, COL_类别)) > 0 And blnCurDo Then
                '检查组合项目行或手术附加行
                lngBeginRow = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
                If lngBeginRow <> -1 Then
                    For i = lngBeginRow To .Rows - 1
                        If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                            If txt单量.Tag <> "" Then
                                .TextMatrix(i, COL_单量) = .TextMatrix(lngRow, COL_单量)
                            End If
                            If txt总量.Tag <> "" Then
                                .TextMatrix(i, COL_总量) = .TextMatrix(lngRow, COL_总量)
                            End If

                            If cbo执行时间.Tag <> "" Then
                                .TextMatrix(i, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                            End If
                            If txt频率.Tag <> "" Then
                                .TextMatrix(i, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质)
                                .TextMatrix(i, COL_频率) = .TextMatrix(lngRow, COL_频率)
                                .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                                .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                                .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                            End If
                            If cbo执行科室.Tag <> "" Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                                    .TextMatrix(i, COL_执行科室ID) = 0
                                ElseIf .TextMatrix(i, COL_类别) <> "G" Then    '手术麻醉的执行科室为单独
                                    .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If

            If blnCurDo Then mblnNoSave = True    '标记为未保存
        End If

        '更新医嘱内容
        If AdviceTextChange(lngRow) Then
            .TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
            txt医嘱内容.Text = .TextMatrix(lngRow, col_医嘱内容)
        End If
    End With

    '清除编辑标志
    Call ClearItemTag

    '某些情况下需要重新设置卡片的项目编辑性(如修改了执行性质时)
    If blnReInRow Then
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub AdviceSet一并给药(ByVal lngBegin As Long, ByVal lngEnd As Long)
'功能：将选择范围内的药品设置为一并给药
'参数：起止行号,中间不包含空行,不包含最后一行药品的给药途径行
'说明：以第一行药品的给药途径为准,但位置放在最后一行药品之后
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lngRow1 As Long, lngRow2 As Long
    Dim lng相关ID As Long, i As Long
    Dim strStart As String, lng配制中心 As Long
    
    lngRow1 = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngBegin, COL_相关ID)), lngBegin + 1) '第一给药途径行
    lngRow2 = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngEnd, COL_相关ID)), lngEnd + 1) '最后给药途径行
    
    '删除给药途径行之前记录执行性质,以便后面作判断
    For i = lngRow2 To lngRow1 Step -1
        If vsAdvice.RowHidden(i) Then
            vsAdvice.Cell(flexcpData, i - 1, COL_执行性质) = Val(vsAdvice.TextMatrix(i, COL_执行性质))
        End If
    Next
    
    '复制第一行的给药途径到最后一行的给药途径
    For i = vsAdvice.FixedCols To vsAdvice.Cols - 1
        If i <> COL_相关ID And i <> COL_序号 Then
            vsAdvice.TextMatrix(lngRow2, i) = vsAdvice.TextMatrix(lngRow1, i)
        End If
    Next
    lng相关ID = vsAdvice.RowData(lngRow2)
    
    varTmp1 = mblnRowChange: varTmp2 = vsAdvice.Redraw
    mblnRowChange = False: vsAdvice.Redraw = flexRDNone
    
    '删除除最后一行给药途径外的其它给药途径
    For i = lngEnd To lngBegin Step -1
        If vsAdvice.RowHidden(i) Then
            Call DeleteRow(i)
        Else
            vsAdvice.TextMatrix(i, COL_相关ID) = lng相关ID
        End If
    Next
    
    '行号已变更
    lngRow1 = lngBegin '开始一并给药行
    
    '处理一并给药其他行的相同信息
    For i = lngRow1 + 1 To vsAdvice.Rows - 1
        If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng相关ID Then
            lngRow2 = i '记录新的结束行号
            
            '一并给药的部分信息相同
            vsAdvice.TextMatrix(i, col_缺省) = vsAdvice.TextMatrix(lngRow1, col_缺省)
            vsAdvice.TextMatrix(i, COL_天数) = vsAdvice.TextMatrix(lngRow1, COL_天数)
            vsAdvice.TextMatrix(i, COL_用法) = vsAdvice.TextMatrix(lngRow1, COL_用法)
            
            vsAdvice.TextMatrix(i, COL_频率性质) = vsAdvice.TextMatrix(lngRow1, COL_频率性质)
            vsAdvice.TextMatrix(i, COL_频率) = vsAdvice.TextMatrix(lngRow1, COL_频率)
            vsAdvice.TextMatrix(i, COL_频率次数) = vsAdvice.TextMatrix(lngRow1, COL_频率次数)
            vsAdvice.TextMatrix(i, COL_频率间隔) = vsAdvice.TextMatrix(lngRow1, COL_频率间隔)
            vsAdvice.TextMatrix(i, COL_间隔单位) = vsAdvice.TextMatrix(lngRow1, COL_间隔单位)
            vsAdvice.TextMatrix(i, COL_执行时间) = vsAdvice.TextMatrix(lngRow1, COL_执行时间)
            
            '离院带药一组相同
            If Val(vsAdvice.TextMatrix(lngRow1, COL_执行性质)) <> 5 And Val(vsAdvice.Cell(flexcpData, lngRow1, COL_执行性质)) = 5 Then
                '第一行是离院带药,全部设置为离院带药
                vsAdvice.TextMatrix(i, COL_执行性质) = vsAdvice.TextMatrix(lngRow1, COL_执行性质)
                If Val(vsAdvice.TextMatrix(i, COL_执行科室ID)) = 0 Then '执行科室可以不同,没有时才缺省相同
                    vsAdvice.TextMatrix(i, COL_执行科室ID) = vsAdvice.TextMatrix(lngRow1, COL_执行科室ID)
                End If
            ElseIf Val(vsAdvice.TextMatrix(i, COL_执行性质)) <> 5 And Val(vsAdvice.Cell(flexcpData, i, COL_执行性质)) = 5 Then
                '当前行是离院带药,则设置为与第一行相同
                vsAdvice.TextMatrix(i, COL_执行性质) = vsAdvice.TextMatrix(lngRow1, COL_执行性质)
                If Val(vsAdvice.TextMatrix(i, COL_执行科室ID)) = 0 Then
                    vsAdvice.TextMatrix(i, COL_执行科室ID) = vsAdvice.TextMatrix(lngRow1, COL_执行科室ID)
                End If
            Else
                '否则保持不变
            End If
        Else
            Exit For
        End If
    Next
    
    '检查这些药品中是否存在配制中心拿药的，以第一个为准
    For i = lngRow1 To vsAdvice.Rows - 1
        If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng相关ID Then
            '自备药的情况不管它
            If Not (Val(vsAdvice.TextMatrix(i, COL_执行科室ID)) = 0 And Val(vsAdvice.TextMatrix(i, COL_执行性质)) = 5) Then
                If sys.DeptHaveProperty(Val(vsAdvice.TextMatrix(i, COL_执行科室ID)), "配制中心") Then
                    lng配制中心 = Val(vsAdvice.TextMatrix(i, COL_执行科室ID)): Exit For
                End If
            End If
        Else
            Exit For
        End If
    Next
    '配制中心一组相同
    If lng配制中心 <> 0 Then
        For i = lngRow1 To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                '自备药的情况不管它
                If Not (Val(vsAdvice.TextMatrix(i, COL_执行科室ID)) = 0 And Val(vsAdvice.TextMatrix(i, COL_执行性质)) = 5) Then
                    vsAdvice.TextMatrix(i, COL_执行科室ID) = lng配制中心
                End If
            Else
                Exit For
            End If
        Next
    End If
    
    mblnRowChange = varTmp1: vsAdvice.Redraw = varTmp2
    mblnNoSave = True '标记为未保存
End Sub

Private Sub AdviceSet单独给药(ByVal lngBegin As Long, ByVal lngEnd As Long)
'功能：取消一组药品的一并给药
'参数：起止行号,中间不包含空行,不包含最后一行药品的给药途径行
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lng给药途径ID As Long, i As Long
    Dim int执行性质 As Integer, str执行性质 As String, str滴速 As String
    Dim lngRow As Long, curDate As Date
    
    With vsAdvice
        varTmp1 = mblnRowChange: varTmp2 = .Redraw
        mblnRowChange = False: .Redraw = flexRDNone
        
        '一并给药途径
        lngRow = .FindRow(CLng(.TextMatrix(lngEnd, COL_相关ID)), lngEnd + 1)
        lng给药途径ID = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
        int执行性质 = Val(.TextMatrix(lngRow, COL_执行性质))
        str滴速 = .TextMatrix(lngRow, COL_医生嘱托)
        
        For i = lngEnd - 1 To lngBegin Step -1 '必须反向
            '设置给药途径行
            If Val(.TextMatrix(i, COL_执行性质)) = 5 And int执行性质 <> 5 Then
                str执行性质 = "自备药"
            ElseIf Val(.TextMatrix(i, COL_执行性质)) <> 5 And int执行性质 = 5 Then
                str执行性质 = "离院带药"
            Else
                str执行性质 = ""
            End If
            .TextMatrix(i, COL_相关ID) = "" '必须清除作为标志
            .TextMatrix(i, COL_相关ID) = AdviceSet给药途径(i, lng给药途径ID, str执行性质, str滴速)
        Next
        
        mblnRowChange = varTmp1: .Redraw = varTmp2
        mblnNoSave = True '标记为未保存
    End With
End Sub

Private Function SaveAdvice() As Boolean
'功能：保存当前病人的医嘱记录
    Dim dbl总量 As Double, lng相关ID
    Dim i As Long, j As Long
    
    With vsAdvice
        mlngNextID = 0
        Call InitSchemeRecordset

        '调整序号为顺序增加
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                .RowData(i) = -1 * .RowData(i)
                If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                    .TextMatrix(i, COL_相关ID) = -1 * Val(.TextMatrix(i, COL_相关ID))
                End If
            End If
        Next
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                lng相关ID = .RowData(i)
                .RowData(i) = GetNextID
                For j = i - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                        .TextMatrix(j, COL_相关ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
                For j = i + 1 To .Rows - 1
                    If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                        .TextMatrix(j, COL_相关ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
            End If
        Next
        
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                '总量转换
                dbl总量 = 0
                If Val(.TextMatrix(i, COL_总量)) <> 0 Then
                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                        '成药转换成零售单位
                        dbl总量 = Format(Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_包装系数)), "0.00000")
                    Else
                        '中药配方付数或非药临嘱总量,不转换
                        dbl总量 = Val(.TextMatrix(i, COL_总量))
                    End If
                End If
                
                mrsScheme.AddNew
                mrsScheme!序号 = Val(.RowData(i))
                mrsScheme!相关序号 = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Null, Val(.TextMatrix(i, COL_相关ID)))
                mrsScheme!期效 = IIF(.TextMatrix(i, COL_期效) = "长嘱", 0, 1)
                mrsScheme!诊疗项目ID = IIF(Val(.TextMatrix(i, COL_诊疗项目ID)) = 0, Null, Val(.TextMatrix(i, COL_诊疗项目ID)))
                mrsScheme!收费细目ID = IIF(Val(.TextMatrix(i, COL_收费细目ID)) = 0, Null, Val(.TextMatrix(i, COL_收费细目ID)))
                mrsScheme!医嘱内容 = IIF(.TextMatrix(i, col_医嘱内容) = "", Null, .TextMatrix(i, col_医嘱内容))
                mrsScheme!天数 = IIF(Val(.TextMatrix(i, COL_天数)) = 0, Null, Val(.TextMatrix(i, COL_天数)))
                mrsScheme!单次用量 = IIF(Val(.TextMatrix(i, COL_单量)) = 0, Null, Val(.TextMatrix(i, COL_单量)))
                mrsScheme!总给予量 = IIF(dbl总量 = 0, Null, dbl总量)
                mrsScheme!医生嘱托 = IIF(.TextMatrix(i, COL_医生嘱托) = "", Null, .TextMatrix(i, COL_医生嘱托))
                mrsScheme!执行频次 = IIF(.TextMatrix(i, COL_频率) = "", Null, .TextMatrix(i, COL_频率))
                mrsScheme!频率次数 = Val(.TextMatrix(i, COL_频率次数))
                mrsScheme!频率间隔 = Val(.TextMatrix(i, COL_频率间隔))
                mrsScheme!间隔单位 = IIF(.TextMatrix(i, COL_间隔单位) = "", Null, .TextMatrix(i, COL_间隔单位))
                mrsScheme!时间方案 = IIF(.TextMatrix(i, COL_执行时间) = "", Null, .TextMatrix(i, COL_执行时间))
                mrsScheme!执行科室ID = IIF(Val(.TextMatrix(i, COL_执行科室ID)) = 0, Null, Val(.TextMatrix(i, COL_执行科室ID)))
                mrsScheme!执行性质 = Val(.TextMatrix(i, COL_执行性质))
                mrsScheme!标本部位 = IIF(.TextMatrix(i, COL_标本部位) = "", Null, .TextMatrix(i, COL_标本部位))
                mrsScheme!检查方法 = IIF(.TextMatrix(i, COL_检查方法) = "", Null, .TextMatrix(i, COL_检查方法))
                mrsScheme!是否缺省 = IIF(Val(.TextMatrix(i, col_缺省)) = -1, 1, 0)
                mrsScheme!是否备选 = IIF(Val(.TextMatrix(i, col_备选)) = -1, 1, 0)
                mrsScheme!配方ID = .TextMatrix(i, COL_配方ID)
                mrsScheme!组合项目ID = .TextMatrix(i, COL_组合项目ID)
                mrsScheme!执行标记 = Val(.TextMatrix(i, COL_执行标记))
                If mbyt场合 = 1 Then
                    mrsScheme!类别 = .TextMatrix(i, COL_类别)
                    mrsScheme!操作类型 = .TextMatrix(i, COL_操作类型)
                End If
                mrsScheme.Update
            End If
        Next
        
        If mrsScheme.RecordCount > 0 Then mrsScheme.MoveFirst
    End With
    
    mblnNoSave = False
    SaveAdvice = True
    mblnOK = True
End Function

Private Function CheckAdvice() As Boolean
'功能：检查当前病人(婴儿)的医嘱输入是否合法
'说明：如果有不合法的地方，在本函数中提示及定位
    Dim blnValid As Boolean
    Dim bln配方行 As Boolean, bln检验行 As Boolean
    Dim dbl总量 As Double, strMsg As String
    Dim blnSkipTotal As Boolean, lngRow As Long, i As Long, j As Long
    Dim vMsg As VbMsgBoxResult, sng天数 As Single
    Dim lngBegin As Long, lngEnd As Long
    
    With vsAdvice
        '为避免意外，将一组医嘱的“缺省”列设置为相同的值
        For i = .FixedRows To .Rows - 1
            Call GetRowScope(i, lngBegin, lngEnd)
            For j = lngBegin To lngEnd
                If .TextMatrix(j, col_缺省) <> .TextMatrix(i, col_缺省) Then
                    .TextMatrix(j, col_缺省) = .TextMatrix(i, col_缺省)
                End If
            Next
            i = lngEnd + 1
        Next
    
    
        For i = .FixedRows To .Rows - 1
            '其它输入合法性检查
            If .RowData(i) <> 0 And Not .RowHidden(i) Then
                bln配方行 = RowIn配方行(i)
                bln检验行 = RowIn检验行(i)
                lngRow = i
                If bln配方行 Then '得到配方的第一药品行
                    lngRow = .FindRow(CStr(.RowData(i)), , COL_相关ID)
                ElseIf bln检验行 Then '得到检验医嘱行
                    lngRow = .FindRow(CStr(.RowData(i)), , COL_相关ID)
                End If
                
                If Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                    '临嘱规格判断(临床路径定义，成套方案定义，允许先不确定规格，只确定到品种)
'                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
'                        If .TextMatrix(i, COL_期效) = "临嘱" And Val(.TextMatrix(i, COL_收费细目ID)) = 0 Then
'                            strMsg = "没有对应的药品规格信息。"
'                            .Col = COL_医嘱内容: Exit For
'                        End If
'                    End If
                    
                    '单量录入合法性
                    If .TextMatrix(i, COL_单量) <> "" Then
                        If .TextMatrix(i, COL_期效) = "长嘱" Then
                            '长嘱：成药或计时,计量项目需要录入
                            If InStr(",1,2,", Val(.TextMatrix(i, COL_计算方式))) > 0 Or InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                                If Not IsNumeric(.TextMatrix(i, COL_单量)) Or Val(.TextMatrix(i, COL_单量)) <= 0 Then
                                    strMsg = "没有录入正确的单次用量。"
                                    .Col = COL_单量: Exit For
                                End If
                            End If
                        Else
                            '临嘱:成药或可选择频率的计时,计量项目可以录入(也可不录)
                            If Val(.TextMatrix(i, COL_频率性质)) = 0 And InStr(",1,2,", Val(.TextMatrix(i, COL_计算方式))) > 0 Then
                                If .TextMatrix(i, COL_单量) <> "" Then
                                    If Not IsNumeric(.TextMatrix(i, COL_单量)) Or Val(.TextMatrix(i, COL_单量)) <= 0 Then
                                        strMsg = "没有录入正确的单次用量。"
                                        .Col = COL_单量: Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    '总量录入合法性:配方,临嘱(药品或其它)
                    If .TextMatrix(i, COL_总量) <> "" Then
                        If .TextMatrix(i, COL_期效) = "临嘱" Then
                            If Not IsNumeric(.TextMatrix(i, COL_总量)) Or Val(.TextMatrix(i, COL_总量)) <= 0 Then
                                If bln配方行 Then
                                    strMsg = "没有录入正确的中药配方付数。"
                                ElseIf InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                                    strMsg = "没有录入正确的药品总给予量。"
                                Else
                                    strMsg = "没有录入正确的总量。"
                                End If
                                .Col = COL_总量: Exit For
                            End If
                        End If
                    End If
                End If
                
                '本次新增或修改的行
                '---------------------------------------------------
                If Val(.TextMatrix(i, COL_诊疗项目ID)) = 0 Then
                    If .TextMatrix(i, col_医嘱内容) = "" Then
                        strMsg = "没有录入医嘱内容。"
                        .Col = COL_用法: Exit For
                    End If
                Else
                    '给药途径，中药用法，采集方法设置检查
                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) = .RowData(i + 1) And Val(.TextMatrix(i + 1, COL_诊疗项目ID)) = 0 Then
                            strMsg = "没有设置对应的给药途径。"
                            .Col = COL_用法: Exit For
                        End If
                    End If
                    If .TextMatrix(i, COL_类别) = "E" And Val(.TextMatrix(i, COL_诊疗项目ID)) = 0 Then
                        If .RowData(i) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            If InStr(",7,E,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                                strMsg = "中药配方没有设置对应的用法。"
                            ElseIf .TextMatrix(i - 1, COL_类别) = "C" Then
                                strMsg = "没有设置对应的标本采集方法。"
                            End If
                            .Col = COL_用法: Exit For
                        End If
                    End If
                    
                    '最少总量检查:至少要满足一个频次周期的用量
                    If Val(.TextMatrix(i, COL_总量)) <> 0 And .TextMatrix(i, COL_期效) = "临嘱" And (InStr(",4,5,6,", .TextMatrix(i, COL_类别)) > 0 Or bln配方行) Then
                        If Not blnSkipTotal And .TextMatrix(i, COL_频率) <> "" Then
                            strMsg = ""
                            If bln配方行 Then '判断
                                dbl总量 = Calc缺省药品总量(1, 1, Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位))
                                If Val(.TextMatrix(i, COL_总量)) < dbl总量 Then
                                    strMsg = .TextMatrix(i, col_医嘱内容) & vbCrLf & vbCrLf & _
                                        "在按""" & .TextMatrix(i, COL_频率) & """执行时,至少需要 " & dbl总量 & "付。"
                                End If
                            ElseIf Val(.TextMatrix(i, COL_剂量系数)) <> 0 And Val(.TextMatrix(i, COL_单量)) <> 0 Then
                                If Val(.TextMatrix(i, COL_频率性质)) = 1 Then '临嘱成药可能为一次性
                                    dbl总量 = Calc缺省药品总量(Val(.TextMatrix(i, COL_单量)), 1, 1, 1, "天", "", Val(.TextMatrix(i, COL_剂量系数)), Val(.TextMatrix(i, COL_包装系数)), Val(.TextMatrix(i, COL_可否分零)))
                                Else
                                    sng天数 = Val(.TextMatrix(i, COL_天数))
                                    If sng天数 = 0 Then sng天数 = 1
                                    dbl总量 = Calc缺省药品总量(Val(.TextMatrix(i, COL_单量)), sng天数, Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位), .TextMatrix(i, COL_执行时间), Val(.TextMatrix(i, COL_剂量系数)), Val(.TextMatrix(i, COL_包装系数)), Val(.TextMatrix(i, COL_可否分零)))
                                End If
                                If Val(.TextMatrix(i, COL_总量)) < dbl总量 Then
                                    strMsg = .TextMatrix(i, col_医嘱内容) & vbCrLf & vbCrLf & _
                                        "在按每次 " & .TextMatrix(i, COL_单量) & .TextMatrix(i, COL_单量单位) & "," & .TextMatrix(i, COL_频率) & _
                                        IIF(Val(.TextMatrix(i, COL_频率性质)) <> 1 And Val(.TextMatrix(i, COL_天数)) > 0 And .TextMatrix(i, COL_类别) <> "4", ",用药 " & sng天数 & " 天", "") & _
                                        "执行时,至少需要 " & dbl总量 & .TextMatrix(i, COL_总量单位) & "。"
                                End If
                            End If
                            If strMsg <> "" Then '提示
                                .Row = i: .Col = COL_总量: Call .ShowCell(.Row, .Col)
                                vMsg = frmMsgBox.ShowMsgBox(strMsg & "^^要继续吗？", Me)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If txt总量.Enabled And txt总量.Visible Then txt总量.SetFocus
                                    Exit Function
                                ElseIf vMsg = vbIgnore Then
                                    blnSkipTotal = True
                                End If
                            End If
                        End If
                    End If
                    
                    '执行时间合法性检查
                    If .TextMatrix(i, COL_执行时间) <> "" And .TextMatrix(i, COL_频率) <> "" Then
                        blnValid = ExeTimeValid(.TextMatrix(i, COL_执行时间), Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位))
                        If Not blnValid Then
                            If .TextMatrix(i, COL_间隔单位) = "周" Then
                                strMsg = COL_按周执行
                            ElseIf .TextMatrix(i, COL_间隔单位) = "天" Then
                                strMsg = COL_按天执行
                            ElseIf .TextMatrix(i, COL_间隔单位) = "小时" Then
                                strMsg = COL_按时执行
                            End If
                            strMsg = "录入的执行时间方案格式不正确，请检查。" & vbCrLf & vbCrLf & "例：" & vbCrLf & strMsg
                            .Col = COL_执行时间: Exit For
                        End If
                    End If
                End If
            End If
        Next
        
        '--------------------------------------------------------------------------
        '中间退出的错误提示
        If i <= .Rows - 1 Then
            .Row = i: Call .ShowCell(.Row, .Col)
            If strMsg <> "" Then
                If bln配方行 Then
                    strMsg = "该中药配方" & strMsg
                Else
                    strMsg = """" & .TextMatrix(i, col_医嘱内容) & """" & strMsg
                End If
                MsgBox strMsg, vbInformation, gstrSysName
                .Refresh
            End If
            If .Col = col_医嘱内容 Then
                If txt医嘱内容.Enabled Then txt医嘱内容.SetFocus
            Else
                Call vsAdvice_KeyPress(13)
            End If
            Exit Function
        End If
        
        '没有数据
        If lngRow = 0 Then
            MsgBox "成套方案中没有内容，请先录入成套方案内容！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckAdvice = True
End Function

Private Function SeekNextControl() As Boolean
'功能：定位到下一个焦点的控件上,并根据情况决定是否自动新增一行医嘱
'返回：如果通过SetFocus强制定位的,则返回True
    Dim objActive As Object, objNext As Object
    Dim blnDo As Boolean, i As Long
    Dim strSkip As String
    
    Set objActive = Me.ActiveControl
    
    If Not objActive Is Nothing Then
        If TypeName(objActive) = "TextBox" Or TypeName(objActive) = "ComboBox" Then
            If objActive.Container Is fraAdvice Then
                strSkip = GetInputSkip(vsAdvice.Row)
                Set objNext = zlControl.GetNextControl(objActive.TabIndex, Me, strSkip)
                If Not objNext Is Nothing Then
                    If objNext Is vsAdvice Then
                        For i = vsAdvice.Row + 1 To vsAdvice.Rows - 1
                            If Not vsAdvice.RowHidden(i) Then
                                Call AdviceChange '强制更新医嘱内容
                                vsAdvice.Row = i
                                Call zlCommFun.PressKey(vbKeyTab)
                                blnDo = vsAdvice.RowData(i) <> 0 '无内容则再跳入编辑
                                Exit For
                            End If
                        Next
                        If i > vsAdvice.Rows - 1 Then
                            blnDo = True
                            If mbyt场合 = 2 Then Exit Function '批量替换时只允许输入一条替换的项目。
                            cbsMain.FindControl(, conMenu_New, True, True).Execute
                        End If
                    ElseIf strSkip <> "" And InStr(";" & strSkip & ";", objNext.Name) = 0 Then
                        blnDo = True: objNext.SetFocus
                    End If
                End If
            End If
        End If
    End If
    If Not blnDo Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        SeekNextControl = True
    End If
End Function

Private Function GetInputSkip(ByVal lngRow As Long) As String
'功能：获取输入医嘱过程中，回车光标应跳过的控件
    Dim strSkip As String, lngFind As Long
    
    With vsAdvice
        '一并给药中的药品输入时应跳过的内容
        If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 And .RowData(lngRow) <> 0 Then
            If Val(.TextMatrix(lngRow, COL_相关ID)) = Val(.TextMatrix(lngRow - 1, COL_相关ID)) Then
                '给药途径,附加执行
                If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                    lngFind = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    If lngFind <> -1 Then
                        If Val(.TextMatrix(lngFind, COL_诊疗项目ID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.txt用法.Name
                        End If
                        If Val(.TextMatrix(lngFind, COL_执行科室ID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.cbo附加执行.Name
                        End If
                    End If
                End If
                '频率
                If .TextMatrix(lngRow, COL_频率) <> "" Then strSkip = strSkip & ";" & Me.txt频率.Name
                '执行时间
                If .TextMatrix(lngRow, COL_执行时间) <> "" Then strSkip = strSkip & ";" & Me.cbo执行时间.Name
            End If
        End If
    End With
    GetInputSkip = Mid(strSkip, 2)
End Function

Private Function AdviceTextChange(ByVal lngRow As Long) As Boolean
'功能：当医嘱卡片输入内容变化时，判断医嘱内容文本是否应该重新组织
    Dim str类别 As String, strText As String, blnDefine As Boolean
    
    With vsAdvice
        '确定医嘱类别
        str类别 = .TextMatrix(lngRow, COL_类别)
        If str类别 = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then '中药配方或一组检验
            lngRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            If lngRow <> -1 Then str类别 = .TextMatrix(lngRow, COL_类别)
        End If
        If str类别 = "7" Then str类别 = "8"
                
        '确定是否定义
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "诊疗类别='" & str类别 & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(NVL(mrsDefine!医嘱内容)) = "" Then
                blnDefine = False
            End If
        End If
        If blnDefine Then strText = mrsDefine!医嘱内容
        
        '检查内容变动
        If blnDefine Then '公共字段部份或可以公共处理的部份
            If cbo医生嘱托.Tag <> "" And InStr(strText, "[医生嘱托]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If cmd频率.Tag <> "" And txt频率.Tag <> "" Then
                If InStr(strText, "[中文频率]") > 0 Or InStr(strText, "[英文频率]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
            If cbo执行时间.Tag <> "" And InStr(strText, "[执行时间]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If (IsNumeric(txt单量.Text) Or txt单量.Text = "") And txt单量.Tag <> "" And InStr(strText, "[单量]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If IsNumeric(txt总量.Text) And txt总量.Tag <> "" And InStr(strText, "[总量]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
        End If
        
        Select Case str类别 '不同的类别检查
        Case "5", "6" '中西成药
            If Not blnDefine Then
                
            Else
                '[输入名][通用名][商品名][英文名][规格][产地]是输入或修改整个药品时变化
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[给药途径]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "8" '中药配方
            If Not blnDefine Then
                If IsNumeric(txt总量.Text) And txt总量.Tag <> "" Then AdviceTextChange = True: Exit Function
                If cmd频率.Tag <> "" And txt频率.Tag <> "" Then AdviceTextChange = True: Exit Function
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[配方组成][煎法]是输入或修改整个配方时变化
                If IsNumeric(txt总量.Text) And txt总量.Tag <> "" And InStr(strText, "[付数]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[用法]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "C" '检验
            If Not blnDefine Then
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[检验项目][检验标本]是输入或修改整个项目时变化
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[采集方法]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "D" '检查
            If Not blnDefine Then
                
            Else
                '[检查项目][检查部位]是输入或修改整个项目时变化
            End If
        Case "F" '手术
            If Not blnDefine Then
            Else
                '[主要手术][附加手术][麻醉方法]是输入或修改整个项目时变化
            End If
        Case "K" '输血
            If Not blnDefine Then
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[输血途径]
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[输血途径]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case Else '其他
            If Not blnDefine Then
                
            Else
                '[诊疗项目]是输入或修改整个项目时变化
            End If
        End Select
    End With
End Function

Private Function AdviceTextMake(ByVal lngRow As Long) As String
'功能：获取医嘱内容文本
'参数：lngRow=已有医嘱数据的可见行
    Dim rsTmp As New ADODB.Recordset
    Dim blnDefine As Boolean, str类别 As String
    Dim strText As String, strSql As String
    Dim strField As String, int频率范围 As Integer
    Dim i As Long, k As Long
    
    Dim str中药 As String, str煎法 As String, str形态 As String
    Dim str麻醉 As String, str附术 As String
    Dim str检验 As String, str标本 As String
    Dim str部位 As String, str部位Last As String, str方法 As String
    Dim dbl数量 As Double
    Dim blnDo As Boolean
    Dim str中药名称 As String
    
    On Error GoTo errH
    
    With vsAdvice
        '确定医嘱类别
        str类别 = .TextMatrix(lngRow, COL_类别)
        If str类别 = "E" Then '中药配方或一组检验
            k = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            If k <> -1 Then str类别 = .TextMatrix(k, COL_类别)
        End If
        If str类别 = "7" Then str类别 = "8"
                
        '确定是否定义
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "诊疗类别='" & str类别 & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(NVL(mrsDefine!医嘱内容)) = "" Then
                blnDefine = False
            End If
        End If
        
ReDoDefault: '用于按定义公式计算失败，重新按缺省规则进行组织
        strText = ""
        If blnDefine Then strText = mrsDefine!医嘱内容
        
        '产生医嘱内容
        Select Case str类别
        Case "C" '检验-------------------------------------------------------------
            str检验 = "": str标本 = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_组合项目ID)) = 0 And mblnNewLIS Or Not mblnNewLIS Then
                        str检验 = .TextMatrix(i, col_医嘱内容) & "," & str检验
                    End If
                    str标本 = .TextMatrix(i, COL_标本部位)
                Else
                    Exit For
                End If
            Next
            If str检验 = "" Then '老的方式
                str检验 = .TextMatrix(lngRow, COL_名称)
            Else
                str检验 = Left(str检验, Len(str检验) - 1)
            End If
            
            If Not blnDefine Then
                strText = str检验 & IIF(str标本 <> "", "(" & str标本 & ")", "")
            Else
                If InStr(strText, "[检验项目]") > 0 Then
                    strField = str检验
                    strText = Replace(strText, "[检验项目]", """" & strField & """")
                End If
                If InStr(strText, "[检验标本]") > 0 Then
                    strField = str标本
                    strText = Replace(strText, "[检验标本]", """" & strField & """")
                End If
                If InStr(strText, "[采集方法]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[采集方法]", """" & strField & """")
                End If
            End If
        Case "D" '检查-------------------------------------------------------------
            str部位 = "": str方法 = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_标本部位) <> "" Then
                        If .TextMatrix(i, COL_标本部位) <> str部位Last And str部位Last <> "" Then
                            str部位 = str部位 & "," & str部位Last & IIF(str方法 <> "", "(" & Mid(str方法, 2) & ")", "")
                            str方法 = ""
                        End If
                        If .TextMatrix(i, COL_检查方法) <> "" Then
                            str方法 = str方法 & "," & .TextMatrix(i, COL_检查方法)
                        End If
                        
                        str部位Last = .TextMatrix(i, COL_标本部位)
                    End If
                Else
                    Exit For
                End If
            Next
            If str部位Last <> "" Then
                str部位 = str部位 & "," & str部位Last & IIF(str方法 <> "", "(" & Mid(str方法, 2) & ")", "")
            End If
            str部位 = Mid(str部位, 2) '检查组合项目的部位
            
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_名称) & IIF(str部位 <> "", ":" & str部位, "")
            Else
                If InStr(strText, "[检查项目]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[检查项目]", """" & strField & """")
                End If
                If InStr(strText, "[检查部位]") > 0 Then
                    strField = str部位
                    strText = Replace(strText, "[检查部位]", """" & strField & """")
                End If
            End If
        Case "F" '手术-------------------------------------------------------------
            str麻醉 = "": str附术 = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "G" Then
                        str麻醉 = .TextMatrix(i, col_医嘱内容)
                    Else
                        str附术 = str附术 & "," & .TextMatrix(i, col_医嘱内容)
                    End If
                Else
                    Exit For
                End If
            Next
            str附术 = Mid(str附术, 2)
            
            If Not blnDefine Then
                strText = ""
                If str麻醉 <> "" Then
                    strText = strText & IIF(str麻醉 <> "", " 在 " & str麻醉 & " 下行 ", " 行 ")
                End If
                strText = strText & .TextMatrix(lngRow, COL_名称)
                If str附术 <> "" Then
                    strText = strText & " 及 " & str附术
                End If
            Else
                If InStr(strText, "[主要手术]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[主要手术]", """" & strField & """")
                End If
                If InStr(strText, "[附加手术]") > 0 Then
                    strField = str附术
                    strText = Replace(strText, "[附加手术]", """" & strField & """")
                End If
                If InStr(strText, "[麻醉方法]") > 0 Then
                    strField = str麻醉
                    strText = Replace(strText, "[麻醉方法]", """" & strField & """")
                End If
            End If
        Case "8" '中药配方---------------------------------------------------------
            str中药 = "": str煎法 = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "7" Then
                        dbl数量 = dbl数量 + Val(.TextMatrix(i, COL_单量))
                        If Val(.TextMatrix(lngRow, COL_中药形态)) = 0 Then
                            blnDo = .TextMatrix(i, COL_收费细目ID) <> .TextMatrix(i - 1, COL_收费细目ID)
                        Else
                            blnDo = .TextMatrix(i, COL_诊疗项目ID) <> .TextMatrix(i - 1, COL_诊疗项目ID)
                        End If
                        
                        If blnDo Then
                            str中药名称 = .TextMatrix(i, col_医嘱内容)
                            
                            If Val(.TextMatrix(lngRow, COL_中药形态)) = 0 Then
                                strSql = "Select 规格 as 名称 From 收费项目目录 Where ID=[1] And Exists(Select 1 From 药品规格 Where 药品ID<>[1] And 药名ID=[2])"
                                Set rsTmp = New ADODB.Recordset '清除Filter
                                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(i, COL_收费细目ID)), Val(.TextMatrix(i, COL_诊疗项目ID)))
                                If rsTmp.RecordCount > 0 Then
                                    If Not IsNull(rsTmp!名称) Then str中药名称 = str中药名称 & "(" & rsTmp!名称 & ")"
                                End If
                            End If
                        
                            str中药 = RTrim(str中药名称 & _
                                " " & FormatEx(dbl数量, 5) & .TextMatrix(i, COL_单量单位) & _
                                " " & .TextMatrix(i, COL_医生嘱托)) & "," & str中药
                            dbl数量 = 0
                        End If
                    ElseIf .TextMatrix(i, COL_类别) = "E" Then
                        str煎法 = .TextMatrix(i, col_医嘱内容) & .TextMatrix(i, COL_标本部位)
                    End If
                Else
                    Exit For
                End If
            Next
            If str中药 <> "" Then
                str中药 = Mid(str中药, 1, Len(str中药) - 1)
            End If
            If Not blnDefine Or .TextMatrix(lngRow, COL_期效) = "长嘱" Then
                If .TextMatrix(lngRow, COL_中药形态) = "1" Then
                    str形态 = "[饮片]"
                ElseIf .TextMatrix(lngRow, COL_中药形态) = "2" Then
                    str形态 = "[免煎剂]"
                End If
                '数字后加了空格在文本框中会自动换行
                If .TextMatrix(lngRow, COL_期效) = "长嘱" Then
                    '长嘱配方内容付数不好处理，暂用固定规则
                    strText = "中药配方" & str形态 & "," & _
                        .TextMatrix(lngRow, COL_频率) & "," & str煎法 & "," & _
                        .TextMatrix(lngRow, COL_用法) & ":" & str中药
                Else
                    strText = "中药" & str形态 & .TextMatrix(lngRow, COL_总量) & "付," & _
                        .TextMatrix(lngRow, COL_频率) & "," & str煎法 & "," & _
                        .TextMatrix(lngRow, COL_用法) & ":" & str中药
                End If
            Else
                If InStr(strText, "[付数]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_总量)
                    strText = Replace(strText, "[付数]", """" & strField & """")
                End If
                If InStr(strText, "[配方组成]") > 0 Then
                    strField = str中药
                    strText = Replace(strText, "[配方组成]", """" & strField & """")
                End If
                If InStr(strText, "[用法]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[用法]", """" & strField & """")
                End If
                If InStr(strText, "[煎法]") > 0 Then
                    strField = str煎法
                    strText = Replace(strText, "[煎法]", """" & strField & """")
                End If
            End If
        Case "4" '卫材------------------------------------------------------------
            If Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                strSql = "Select 名称,规格,产地 From 收费项目目录 Where ID=[1]"
                Set rsTmp = New ADODB.Recordset '清除Filter
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(lngRow, COL_收费细目ID)))
            ElseIf blnDefine Then
                strSql = "Select 名称,NULL As 规格,NULL As 产地 From 诊疗项目目录 Where ID=[1]"
                Set rsTmp = New ADODB.Recordset '清除Filter
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(lngRow, COL_诊疗项目ID)))
            End If
                If Not blnDefine Then
                    strText = .TextMatrix(lngRow, COL_名称)
                    If Not IsNull(rsTmp!规格) Then
                        strText = strText & " " & rsTmp!规格
                    End If
                Else
                    If InStr(strText, "[卫生材料]") > 0 Then
                        strField = rsTmp!名称
                        strText = Replace(strText, "[卫生材料]", """" & strField & """")
                    End If
                    If InStr(strText, "[规格]") > 0 Then
                        strField = NVL(rsTmp!规格)
                        strText = Replace(strText, "[规格]", """" & strField & """")
                    End If
                    If InStr(strText, "[产地]") > 0 Then
                        strField = NVL(rsTmp!产地)
                        strText = Replace(strText, "[产地]", """" & strField & """")
                    End If
                End If
        Case "5", "6" '西成药，中成药---------------------------------------------
            If Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                '性质:0-正名,1-英文名,3-商品名
                strSql = "Select Nvl(B.名称,A.名称) as 名称,A.规格,A.产地,B.性质" & _
                    " From 收费项目目录 A,收费项目别名 B Where A.ID=B.收费细目ID(+) And A.ID=[1] Order by B.性质,B.码类"
                Set rsTmp = New ADODB.Recordset '清除Filter
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(lngRow, COL_收费细目ID)))
            ElseIf blnDefine Then
                '性质:0-正名,1-英文名
                strSql = "Select Nvl(B.名称,A.名称) as 名称,Null as 规格,Null as 产地,B.性质" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B Where A.ID=B.诊疗项目ID(+) And A.ID=[1] Order by B.性质,B.码类"
                Set rsTmp = New ADODB.Recordset '清除Filter
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(lngRow, COL_诊疗项目ID)))
            End If
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_标本部位)
                If Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                    If strText = "" Then
                        If gbyt药品名称显示 <> 0 Then rsTmp.Filter = "性质=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strText = rsTmp!名称
                    End If
                    If Not IsNull(rsTmp!产地) Then
                        strText = strText & "(" & rsTmp!产地 & ")"
                    End If
                    If Not IsNull(rsTmp!规格) Then
                        strText = strText & " " & rsTmp!规格
                    End If
                Else
                    If strText = "" Then
                        strText = .TextMatrix(lngRow, COL_名称)
                    End If
                End If
            Else
                If InStr(strText, "[输入名]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_标本部位)
                    If strField = "" Then
                        If gbyt药品名称显示 <> 0 Then rsTmp.Filter = "性质=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strField = rsTmp!名称
                    End If
                    strText = Replace(strText, "[输入名]", """" & strField & """")
                End If
                If InStr(strText, "[通用名]") > 0 Then
                    rsTmp.Filter = 0
                    strField = rsTmp!名称
                    strText = Replace(strText, "[通用名]", """" & strField & """")
                End If
                If InStr(strText, "[商品名]") > 0 Then
                    rsTmp.Filter = "性质=3"
                    If rsTmp.EOF Then
                        strField = ""
                    Else
                        strField = rsTmp!名称
                    End If
                    strText = Replace(strText, "[商品名]", """" & strField & """")
                End If
                If InStr(strText, "[英文名]") > 0 Then
                    rsTmp.Filter = "性质=2"
                    If rsTmp.EOF Then
                        strField = ""
                    Else
                        strField = rsTmp!名称
                    End If
                    strText = Replace(strText, "[英文名]", """" & strField & """")
                End If
                If InStr(strText, "[规格]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = NVL(rsTmp!规格)
                    strText = Replace(strText, "[规格]", """" & strField & """")
                End If
                If InStr(strText, "[产地]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = NVL(rsTmp!产地)
                    strText = Replace(strText, "[产地]", """" & strField & """")
                End If
                If InStr(strText, "[给药途径]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[给药途径]", """" & strField & """")
                End If
            End If
        Case "K" '输血医嘱
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_名称)
                If .TextMatrix(lngRow, COL_用法) <> "" Then
                    strText = strText & "(" & .TextMatrix(lngRow, COL_用法) & ")"
                End If
            Else
                If InStr(strText, "[诊疗项目]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[诊疗项目]", """" & strField & """")
                End If
                If InStr(strText, "[输血项目]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[输血项目]", """" & strField & """")
                End If
                If InStr(strText, "[输血途径]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[输血途径]", """" & strField & """")
                End If
            End If
        Case Else '其它所有类别-----------------------------------------------------
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_名称)
            Else
                If InStr(strText, "[诊疗项目]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[诊疗项目]", """" & strField & """")
                End If
            End If
            '术后医嘱特殊显示
            If .TextMatrix(lngRow, COL_类别) = "Z" And (Val(.TextMatrix(lngRow, COL_操作类型)) = 4 Or Val(.TextMatrix(lngRow, COL_操作类型)) = 14) Then
                strText = "━━━" & strText & "━━━"
            End If
            '转科医嘱特殊显示
            If .TextMatrix(lngRow, COL_类别) = "Z" And Val(.TextMatrix(lngRow, COL_操作类型)) = 3 Then
                strText = "━━━" & strText & "━━━"
            End If
        End Select
        
        '公共字段或可以公共处理的字段-------------------------------------------
        If blnDefine Then
            If InStr(strText, "[医生嘱托]") > 0 Then
                strField = .Cell(flexcpData, lngRow, COL_医生嘱托)
                strText = Replace(strText, "[医生嘱托]", """" & strField & """")
            End If
            If InStr(strText, "[中文频率]") > 0 Then
                strField = .TextMatrix(lngRow, COL_频率)
                strText = Replace(strText, "[中文频率]", """" & strField & """")
            End If
            If InStr(strText, "[英文频率]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_频率) <> "" Then
                    int频率范围 = Get频率范围(lngRow)
                    strSql = "Select 英文名称 From 诊疗频率项目 Where 名称=[1] And 适用范围=[2]"
                    Set rsTmp = New ADODB.Recordset '清除Filter
                    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, .TextMatrix(lngRow, COL_频率), int频率范围)
                    If Not rsTmp.EOF Then strField = NVL(rsTmp!英文名称)
                End If
                strText = Replace(strText, "[英文频率]", """" & strField & """")
            End If
            If InStr(strText, "[单量]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_单量) <> "" Then
                    strField = .TextMatrix(lngRow, COL_单量) & .TextMatrix(lngRow, COL_单量单位)
                End If
                strText = Replace(strText, "[单量]", """" & strField & """")
            End If
            If InStr(strText, "[总量]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_总量) <> "" Then
                    strField = .TextMatrix(lngRow, COL_总量) & .TextMatrix(lngRow, COL_总量单位)
                End If
                strText = Replace(strText, "[总量]", """" & strField & """")
            End If
            If InStr(strText, "[执行时间]") > 0 Then
                strField = .TextMatrix(lngRow, COL_执行时间)
                strText = Replace(strText, "[执行时间]", """" & strField & """")
            End If
        End If
                
        '计算医嘱内容
        If blnDefine Then
            On Error Resume Next
            strText = mobjVBA.Eval(strText)
            If mobjVBA.Error.Number <> 0 Then
                err.Clear: On Error GoTo errH
                blnDefine = False: GoTo ReDoDefault
            End If
        End If
    End With
    AdviceTextMake = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CanAlterType(ByVal lngRow As Long) As Boolean
'功能：判断指定的医嘱是否可以切换期效
'参数：lngRow=可见的医嘱行
'说明：允许切换期效的条件：
'   1.成长嘱：执行频率=0(可选频率),2(持续性)
'   2.成临嘱：执行频率=0(可选频率),1(一次性);药品必须指定了规格
    Dim rsMore As New ADODB.Recordset
    Dim strSql As String, strType As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            CanAlterType = True: Exit Function
        ElseIf Val(.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
            '自由输入的可以切换
            CanAlterType = True: Exit Function
        ElseIf RowIn配方行(lngRow) Then
            '中药配方固定可以切换
            CanAlterType = True: Exit Function
        ElseIf RowIn检验行(lngRow) Then
            '检验以检验行为准判断
            lngRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            If lngRow = -1 Then Exit Function
        End If
    
        strType = IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", "临嘱", "长嘱")
        
        '以原始频率为准判断:因为可选择频率的可能已缺成一次性
        strSql = "Select 执行频率 From 诊疗项目目录 Where ID=[1]"
        On Error GoTo errH
        Set rsMore = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(lngRow, COL_诊疗项目ID)))
        
        If strType = "长嘱" Then
            If InStr(",0,2,", NVL(rsMore!执行频率, 0)) = 0 Then Exit Function
        Else
            If InStr(",0,1,", NVL(rsMore!执行频率, 0)) = 0 Then Exit Function
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                Call GetRowScope(lngRow, lngBegin, lngEnd)
                For i = lngBegin To lngEnd
                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_收费细目ID)) = 0 Then Exit Function
                    End If
                Next
            End If
        End If
    End With
    CanAlterType = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceAlterType(ByVal lngRow As Long)
'功能：在尽量保持数据的情况下，切换指定行医嘱的期效(长期<->临时)
'参数：lngRow=可见的医嘱行
'说明：执行该函数时应保证已用CanAlterType函数进行了判断
    Dim rsMore As New ADODB.Recordset
    Dim strType As String, strSql As String
    Dim int频率性质 As Integer, sng天数 As Single
    Dim str频率 As String, int频率次数 As Integer
    Dim int频率间隔 As Integer, str间隔单位 As String
    Dim lng用法ID As Long, blnToNormal As Boolean
    Dim lngBegin As Long, lngEnd As Long
    Dim lngCopyRow As Long, i As Long
    
    On Error GoTo errH
    With vsAdvice
        '最终要转换为的期效
        strType = IIF(.TextMatrix(lngRow, COL_期效) = "长嘱", "临嘱", "长嘱")
        
        If Val(.TextMatrix(lngRow, COL_诊疗项目ID)) <> 0 Then
            '取上一或下一有效行,某些内容缺省与该行相同
            lngCopyRow = GetPreRow(lngRow)
            If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
            
            '获取一组医嘱的操作行范围
            Call GetRowScope(lngRow, lngBegin, lngEnd)
        End If
        
        '针对不同类别的医嘱进行转换-----------------------------------------
        If Val(.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
            '自由录入的医嘱直接处理
            .TextMatrix(lngRow, COL_期效) = strType
        ElseIf RowIn配方行(lngRow) Then '中药配方
            '药品长嘱不能为离院带药
            If strType = "长嘱" And .TextMatrix(lngEnd, COL_类别) = "E" _
                And .RowData(lngEnd) = Val(.TextMatrix(lngBegin, COL_相关ID)) Then
                If Val(.TextMatrix(lngBegin, COL_执行性质)) <> 5 And Val(.TextMatrix(lngEnd, COL_执行性质)) = 5 Then
                    lng用法ID = Val(.TextMatrix(lngEnd, COL_诊疗项目ID))
                    blnToNormal = True '表示给药执行应恢复成正常值
                End If
            End If
            
            For i = lngBegin To lngEnd
                '期效值
                .TextMatrix(i, COL_期效) = strType

                '总量
                If strType = "长嘱" Then
                    .TextMatrix(i, COL_总量) = ""
                End If
                '备用医嘱频率
                If .TextMatrix(i, COL_频率) = "必要时" Then
                    .TextMatrix(i, COL_频率) = "需要时"
                    txt频率.Text = "需要时"
                    cmd频率.Tag = "需要时"
                ElseIf .TextMatrix(i, COL_频率) = "需要时" Then
                    .TextMatrix(i, COL_频率) = "必要时"
                    txt频率.Text = "必要时"
                    cmd频率.Tag = "必要时"
                End If
                
                '执行性质:药品长嘱不能为"离院带药"
                If i = lngEnd And blnToNormal Then
                    strSql = "Select 执行科室 From 诊疗项目目录 Where ID=[1]"
                    Set rsMore = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng用法ID)
                    
                    .TextMatrix(i, COL_执行性质) = NVL(rsMore!执行科室, 0)
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                        .TextMatrix(i, COL_执行科室ID) = Get成套执行科室ID("E", lng用法ID, 0, NVL(rsMore!执行科室, 0), IIF(strType = "长嘱", 0, 1), mint范围)
                    Else
                        .TextMatrix(i, COL_执行科室ID) = 0
                    End If
                End If
            Next
        Else '其它诊断项目,包括药品,卫材,检查(组合),手术(组合)；检验组合因代码处理部份相同,因此一起处理
            '获取给药途径ID
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 _
                And .TextMatrix(lngEnd, COL_类别) = "E" And .RowData(lngEnd) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                lng用法ID = Val(.TextMatrix(lngEnd, COL_诊疗项目ID))
                
                '药品长嘱不能为离院带药
                If strType = "长嘱" Then
                    If Val(.TextMatrix(lngRow, COL_执行性质)) <> 5 And Val(.TextMatrix(lngEnd, COL_执行性质)) = 5 Then
                        blnToNormal = True '表示给药执行应恢复成正常值
                    End If
                End If
            End If
            
            '------------------------------------------------------------------------------------------------------
            '同时处理一组医嘱的相关行
            For i = lngBegin To lngEnd
                '期效值
                .TextMatrix(i, COL_期效) = strType
                
                '由长期药嘱切换为临时药嘱
                If .Cell(flexcpData, i, COL_可否分零) = 1 Then
                    .Cell(flexcpData, i, COL_可否分零) = Empty '按理还应还原本身的分零方式
                End If
                
                '获取当前项目的附加信息
                If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 And i = lngBegin Then
                    '第一药品行才取这些信息
                    strSql = "Select 项目ID,频次,疗程 From 诊疗用法用量 Where Nvl(性质,0)>0 And 项目ID=[1] And 用法ID=[2]"
                    strSql = "Select A.执行科室,A.执行频率,A.计算方式,A.计算单位,B.频次,B.疗程" & _
                        " From 诊疗项目目录 A,(" & strSql & ") B Where A.ID=B.项目ID(+) And A.ID=[1]"
                Else
                    strSql = "Select 执行科室,执行频率,计算方式,计算单位,Null as 频次,Null as 疗程 From 诊疗项目目录 Where ID=[1]"
                End If
                Set rsMore = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(i, COL_诊疗项目ID)), lng用法ID)
                If Not rsMore.EOF Then '给药途径没有指定的情况
                    '总量(单位)
                    If strType = "临嘱" Then
                        If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                            '中、西成药临嘱的总量单位就是包装单位
                            .TextMatrix(i, COL_总量单位) = .TextMatrix(i, COL_包装单位)
                        ElseIf .TextMatrix(i, COL_类别) = "4" Then
                            .TextMatrix(i, COL_总量单位) = .TextMatrix(i, COL_包装单位) '散装单位
                        Else
                            '其它临嘱要输入总量
                            .TextMatrix(i, COL_总量单位) = NVL(rsMore!计算单位)
                            
                            '如果为一次性或计次临嘱缺省总量为1
                            If i = lngBegin Then
                                If NVL(rsMore!执行频率, 0) = 1 Or NVL(rsMore!计算方式, 0) = 3 Then
                                    .TextMatrix(i, COL_总量) = 1
                                End If
                            ElseIf Not (lng用法ID = Val(.TextMatrix(i, COL_诊疗项目ID))) Then
                                .TextMatrix(i, COL_总量) = .TextMatrix(lngBegin, COL_总量)
                            End If
                        End If
                    Else
                        .TextMatrix(i, COL_总量) = ""
                        .TextMatrix(i, COL_总量单位) = ""
                    End If
                    
                    '备用医嘱频率
                    If .TextMatrix(i, COL_频率) = "必要时" Then
                        .TextMatrix(i, COL_频率) = "需要时"
                        txt频率.Text = "需要时"
                        cmd频率.Tag = "需要时"
                    ElseIf .TextMatrix(i, COL_频率) = "需要时" Then
                        .TextMatrix(i, COL_频率) = "必要时"
                        txt频率.Text = "必要时"
                        cmd频率.Tag = "必要时"
                    Else
                
                        '频率性质,执行频率,执行时间
                        If i = lngBegin Then '以第一行为准
                            int频率性质 = Val(.TextMatrix(i, COL_频率性质))
                            If strType = "临嘱" And NVL(rsMore!执行频率, 0) = 0 And mbln一次性 Then
                                .TextMatrix(i, COL_频率性质) = 1 '可选择频率的临嘱缺省为一次性
                            Else
                                .TextMatrix(i, COL_频率性质) = NVL(rsMore!执行频率, 0)
                            End If
            
                            '执行频率:当适用范围有所变化时
                            If Val(.TextMatrix(i, COL_频率性质)) <> int频率性质 Then
                                '标记为重新取
                                .TextMatrix(i, COL_频率) = ""
                                .TextMatrix(i, COL_执行时间) = ""
                                
                                '药品设置的缺省频率优先
                                If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 _
                                    And Not IsNull(rsMore!频次) And Val(.TextMatrix(i, COL_频率性质)) <> 1 Then
                                    Call Get频率信息_编码(rsMore!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                                    .TextMatrix(i, COL_频率) = str频率
                                    .TextMatrix(i, COL_频率次数) = int频率次数
                                    .TextMatrix(i, COL_频率间隔) = int频率间隔
                                    .TextMatrix(i, COL_间隔单位) = str间隔单位
                                End If
                                '缺省与上一新增行相同
                                If .TextMatrix(i, COL_频率) = "" And lngCopyRow <> -1 Then
                                    If .TextMatrix(i, COL_期效) = .TextMatrix(lngCopyRow, COL_期效) _
                                        And Val(.TextMatrix(i, COL_频率性质)) = Val(.TextMatrix(lngCopyRow, COL_频率性质)) Then
                                        If .TextMatrix(lngCopyRow, COL_频率) <> "" _
                                            And Not (.TextMatrix(i, COL_类别) = "7" And Not RowIn配方行(lngCopyRow)) _
                                            And Not (.TextMatrix(i, COL_类别) <> "7" And RowIn配方行(lngCopyRow)) _
                                            And Check频率可用(Val(.TextMatrix(i, COL_诊疗项目ID)), Get频率范围(i), .TextMatrix(lngCopyRow, COL_频率)) Then
                                            .TextMatrix(i, COL_频率) = .TextMatrix(lngCopyRow, COL_频率)
                                            .TextMatrix(i, COL_频率次数) = .TextMatrix(lngCopyRow, COL_频率次数)
                                            .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngCopyRow, COL_频率间隔)
                                            .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngCopyRow, COL_间隔单位)
                                        End If
                                    End If
                                End If
                                '或取缺省频率
                                If .TextMatrix(i, COL_频率) = "" Then
                                    Call Get缺省频率(Val(.TextMatrix(i, COL_诊疗项目ID)), Get频率范围(i), str频率, int频率次数, int频率间隔, str间隔单位)
                                    .TextMatrix(i, COL_频率) = str频率
                                    .TextMatrix(i, COL_频率次数) = int频率次数
                                    .TextMatrix(i, COL_频率间隔) = int频率间隔
                                    .TextMatrix(i, COL_间隔单位) = str间隔单位
                                End If
                                
                                '执行时间:可选频率的项目
                                If Val(.TextMatrix(i, COL_频率性质)) = 0 Then
                                    If lngCopyRow <> -1 Then '与上一行相同
                                        If .TextMatrix(i, COL_频率) = .TextMatrix(lngCopyRow, COL_频率) Then
                                            .TextMatrix(i, COL_执行时间) = .TextMatrix(lngCopyRow, COL_执行时间)
                                        End If
                                    End If
                                    If .TextMatrix(i, COL_执行时间) = "" Then  '缺省时间方案
                                        .TextMatrix(i, COL_执行时间) = Get缺省时间(1, .TextMatrix(i, COL_频率), lng用法ID)
                                    End If
                                End If
                            End If
                        Else
                            .TextMatrix(i, COL_频率) = .TextMatrix(lngBegin, COL_频率)
                            .TextMatrix(i, COL_频率次数) = .TextMatrix(lngBegin, COL_频率次数)
                            .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngBegin, COL_频率间隔)
                            .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngBegin, COL_间隔单位)
                            .TextMatrix(i, COL_频率性质) = .TextMatrix(lngBegin, COL_频率性质)
                            .TextMatrix(i, COL_执行时间) = .TextMatrix(lngBegin, COL_执行时间)
                        End If
                    End If
                    
                    '药品临嘱天数和总量
                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 And strType = "临嘱" Then
                        '确定临嘱用药天数：
                        '1.最少为一个频率周期天数
                        '2-有疗程则为疗程天数(应大于一个频率周期天数)
                        If i = lngBegin Then '以第一行为准
                            sng天数 = Val(.TextMatrix(i, COL_天数)) '如果以前有则保持
                            If sng天数 = 0 Then sng天数 = msng天数
                            
                            If .TextMatrix(i, COL_间隔单位) = "周" Then
                                If 7 > sng天数 Then sng天数 = 7
                            ElseIf .TextMatrix(i, COL_间隔单位) = "天" Then
                                If Val(.TextMatrix(i, COL_频率间隔)) > sng天数 Then
                                    sng天数 = Val(.TextMatrix(i, COL_频率间隔))
                                End If
                            ElseIf .TextMatrix(i, COL_间隔单位) = "小时" Then
                                If Val(.TextMatrix(i, COL_频率间隔)) \ 24 > sng天数 Then
                                    sng天数 = Val(.TextMatrix(i, COL_频率间隔)) \ 24
                                End If
                            ElseIf .TextMatrix(i, COL_间隔单位) = "分钟" Then
                                If sng天数 = 0 Then sng天数 = 1
                            End If

                            If NVL(rsMore!疗程, 1) > sng天数 Then sng天数 = NVL(rsMore!疗程, 1)
                            If sng天数 = 0 Then sng天数 = 1
                        End If
                        
                        '天数
                        If Val(.TextMatrix(i, COL_频率性质)) <> 1 Then
                            .TextMatrix(i, COL_天数) = IIF(sng天数 = 0, "", sng天数)
                        End If
                        
                        '总量
                        If .TextMatrix(i, COL_频率) <> "" And Val(.TextMatrix(i, COL_单量)) <> 0 _
                            And Val(.TextMatrix(i, COL_剂量系数)) <> 0 And Val(.TextMatrix(i, COL_包装系数)) <> 0 Then
                            If Val(.TextMatrix(i, COL_频率性质)) = 1 Then '临嘱药品可能缺省为一次性
                                '仅按疗程算改为按最少用药天数算
                                .TextMatrix(i, COL_总量) = FormatEx(Calc缺省药品总量( _
                                        Val(.TextMatrix(i, COL_单量)), 1, 1, 1, "天", "", Val(.TextMatrix(i, COL_剂量系数)), _
                                        Val(.TextMatrix(i, COL_包装系数)), Val(.TextMatrix(i, COL_可否分零))), 5)
                            Else
                                '仅按疗程算改为按最少用药天数算
                                .TextMatrix(i, COL_总量) = FormatEx(Calc缺省药品总量( _
                                        Val(.TextMatrix(i, COL_单量)), sng天数, Val(.TextMatrix(i, COL_频率次数)), _
                                        Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位), _
                                        .TextMatrix(i, COL_执行时间), Val(.TextMatrix(i, COL_剂量系数)), _
                                        Val(.TextMatrix(i, COL_包装系数)), Val(.TextMatrix(i, COL_可否分零))), 5)
                            End If
                        End If
                    End If
                    
                    '执行性质:药品长嘱不能为"离院带药"
                    If i = lngEnd And blnToNormal Then
                        .TextMatrix(i, COL_执行性质) = NVL(rsMore!执行科室, 0)
                        If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                            .TextMatrix(i, COL_执行科室ID) = Get成套执行科室ID("E", lng用法ID, 0, NVL(rsMore!执行科室, 0), IIF(strType = "长嘱", 0, 1), mint范围)
                        Else
                            .TextMatrix(i, COL_执行科室ID) = 0
                        End If
                    End If
                End If
            Next
        End If
    End With
    
    mblnNoSave = True '标记为未保存
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get成套执行科室(objCbo As Object, ByVal str类别 As String, ByVal lng项目id As Long, ByVal lng药品ID As Long, _
    ByVal int执行科室 As Integer, ByVal lng当前执行ID As Long, ByVal int期效 As Integer, ByVal int范围 As Integer) As Boolean
'功能：根据诊疗项目执行科室信息返回可用的执行科室在指定下拉框中
'参数：int执行科室=项目执行科室标志
'      lng当前执行ID=医嘱当前的执行科室ID
'      int期效=0-长嘱,1-临嘱,临嘱可能要判断上班时间
'      int范围=1-门诊,2-住院,3-门诊和住院
'说明：对非药医嘱,当前的执行科室可能是强行选择出来的,需要显示在选择框中;另选择框中增加一个其它供选择
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, str药房 As String
    Dim bln规格 As Boolean, i As Long
    
    If str类别 = "4" Then
        strSql = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象" & _
            " From " & IIF(lng药品ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            " And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([2],3)") & " And B.部门ID=C.ID " & IIF(lng药品ID <> 0, " And A.收费细目ID=[3]", " And A.诊疗项目ID=[4]") & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " Order by B.服务对象,C.编码"
    ElseIf InStr(",5,6,7,", str类别) > 0 Then
        bln规格 = ((int期效 = 1 Or gbln药品按规格下医嘱) And lng药品ID <> 0) Or lng药品ID <> 0
        
        '系统可以指定药品执行科室,这里提取所有可选的供再选择
        If str类别 = "5" Then
            str药房 = "西药房"
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
        End If
            
        '药品从系统指定的储备药房中找
        strSql = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象" & _
            " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
            " And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([2],3)") & " And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            IIF(bln规格, " And A.收费细目ID=[3]", " And A.诊疗项目ID=[4]") & _
            " Order by B.服务对象,C.编码"
    Else
        Select Case int执行科室
            Case 0, 5 '0-无执行的叮嘱,5-院外执行
                Get成套执行科室 = True: Exit Function
            Case 1, 2, 3, 6 '1-病人所在科室/2-病人所在病区/3-操作员所在科室/6-开单人所在科室
                strSql = _
                    " Select ID,编码,简码,名称 From 部门表 Where ID=[5]" & _
                    " Union " & _
                    " Select Distinct A.ID,A.编码,A.简码,A.名称" & _
                    " From 部门表 A,部门人员 B,部门性质说明 C" & _
                    " Where A.ID=B.部门ID And A.ID=C.部门ID" & _
                    " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And " & IIF(mint范围 = 3, "Nvl(C.服务对象,0)<>0", "C.服务对象 IN([2],3)") & " And B.人员ID=[6]" & _
                    " Order by 编码"
            Case 4 '4-指定科室
                strSql = _
                    " Select Distinct A.ID,A.编码,A.简码,A.名称" & _
                    " From 部门表 A,诊疗执行科室 B,部门性质说明 C" & _
                    " Where A.ID=B.执行科室ID And A.ID=C.部门ID" & _
                    " And " & IIF(mint范围 = 3, "Nvl(C.服务对象,0)<>0", "C.服务对象 IN([2],3)") & " And B.诊疗项目ID=[4]" & _
                    " Union Select ID,编码,简码,名称 From 部门表 Where ID=[5]" & _
                    " Order by 编码"
        End Select
    End If
        
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", str药房, int范围, lng药品ID, lng项目id, lng当前执行ID, UserInfo.部门ID)
    objCbo.Clear
    For i = 1 To rsTmp.RecordCount
        '使用API快速加入,不然可能有点慢
        AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, rsTmp!编码 & "-" & rsTmp!名称
        SetComboData objCbo.Hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
        If lng当前执行ID = rsTmp!ID Then
            Call Cbo.SetIndex(objCbo.Hwnd, i - 1)
        End If
        rsTmp.MoveNext
    Next
    
    '仅非药、非卫材医嘱可以选择
    If InStr(",4,5,6,7,", str类别) = 0 And objCbo.ListCount = 0 Then
        AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, "[其它...]"
        SetComboData objCbo.Hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get成套执行科室ID(ByVal str类别 As String, ByVal lng项目id As Long, ByVal lng药品ID As Long, _
    ByVal int执行科室 As Integer, ByVal int期效 As Integer, ByVal int范围 As Integer) As Long
'功能：根据诊疗项目执行科室信息返回缺省的执行科室ID
'参数：lng药品ID=药品ID,确定到规格时要用
'      int执行科室=项目执行科室标志
'      int期效=0-长嘱,1-临嘱,临嘱可能要判断上班时间
'      int范围=1-门诊,2-住院,3-门诊和住院
'      blnBy缺省=获取缺省药房时，如果本地有指定，是否按本地缺省指定的药房来，没有则不返回
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim str药房 As String, lng药房 As Long
    Dim bln规格 As Boolean
    
    On Error GoTo errH
    
    If str类别 = "4" Then
        lng药房 = Val(zldatabase.GetPara(decode(int范围, 1, "门诊", 2, "住院", "") & "缺省发料部门", glngSys, decode(int范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0)))
        strSql = _
            " Select Distinct B.服务对象,C.编码,A.执行科室ID" & _
            " From " & IIF(lng药品ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            " And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([1],3)") & " And B.部门ID=C.ID " & IIF(lng药品ID <> 0, " And A.收费细目ID=[2]", " And A.诊疗项目ID=[3]") & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " Order by B.服务对象,C.编码"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", int范围, lng药品ID, lng项目id)
        If Not rsTmp.EOF Then
            If rsTmp.RecordCount = 1 Then
                Get成套执行科室ID = rsTmp!执行科室ID
            ElseIf lng药房 <> 0 Then
                rsTmp.Filter = "执行科室ID=" & lng药房
                If Not rsTmp.EOF Then Get成套执行科室ID = rsTmp!执行科室ID
            End If
        End If
    ElseIf InStr(",5,6,7,", str类别) > 0 Then
        bln规格 = ((int期效 = 1 Or gbln药品按规格下医嘱) And lng药品ID <> 0) Or lng药品ID <> 0
        
        If str类别 = "5" Then
            str药房 = "西药房"
            lng药房 = Val(zldatabase.GetPara(decode(int范围, 1, "门诊", 2, "住院", "") & "缺省西药房", glngSys, decode(int范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0)))
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            lng药房 = Val(zldatabase.GetPara(decode(int范围, 1, "门诊", 2, "住院", "") & "缺省成药房", glngSys, decode(int范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0)))
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            lng药房 = Val(zldatabase.GetPara(decode(int范围, 1, "门诊", 2, "住院", "") & "缺省中药房", glngSys, decode(int范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0)))
        End If
        
        '药品从系统指定的储备药房中找
        strSql = _
            " Select Distinct B.服务对象,C.编码,A.执行科室ID" & _
            " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
            " And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([2],3)") & " And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
             IIF(bln规格, " And A.收费细目ID=[3]", " And A.诊疗项目ID=[4]") & _
            " Order by B.服务对象,C.编码"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", str药房, int范围, lng药品ID, lng项目id)
        If Not rsTmp.EOF Then
            If rsTmp.RecordCount = 1 Then
                Get成套执行科室ID = rsTmp!执行科室ID
            ElseIf lng药房 <> 0 Then
                rsTmp.Filter = "执行科室ID=" & lng药房
                If Not rsTmp.EOF Then Get成套执行科室ID = rsTmp!执行科室ID
            End If
        End If
    Else
        Select Case int执行科室
            Case 0, 5 '0-无执行的叮嘱/5-院外执行
                Exit Function
            Case 1, 2, 3, 6 '1-病人所在科室/2-病人所在病区/3-操作员所在科室/6-开单人所在科室
                Get成套执行科室ID = UserInfo.部门ID
            Case 4 '4-指定科室
                strSql = "Select Distinct A.执行科室ID From 诊疗执行科室 A,部门性质说明 B" & _
                    " Where A.执行科室ID=B.部门ID And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([2],3)") & " And A.诊疗项目ID=[1]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng项目id, int范围)
                If Not rsTmp.EOF Then
                    If rsTmp.RecordCount = 1 Then
                        Get成套执行科室ID = rsTmp!执行科室ID
                    End If
                End If
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get可用药房IDs(ByVal str类别 As String, ByVal lng项目id As Long, _
    ByVal lng药品ID As Long, ByVal lng科室id As Long, Optional ByVal int范围 As Integer = 2) As String
'功能：获取药品的有效诊疗执行科室ID串,用于判断缺省执行科室
'参数：lng科室ID=病人科室ID
'      int范围=1-门诊,2-住院(缺省)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, str药房 As String, str可用药房 As String
    Dim str药房IDs As String
    
    '系统可以指定药品执行科室,这里提取所有可选的供再选择
    If str类别 = "5" Then
        str药房 = "西药房"
        str可用药房 = zldatabase.GetPara(decode(int范围, 1, "门诊", 2, "住院", "") & "可用西药房", glngSys, decode(int范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0))
    ElseIf str类别 = "6" Then
        str药房 = "成药房"
        str可用药房 = zldatabase.GetPara(decode(int范围, 1, "门诊", 2, "住院", "") & "可用成药房", glngSys, decode(int范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0))
    ElseIf str类别 = "7" Then
        str药房 = "中药房"
        str可用药房 = zldatabase.GetPara(decode(int范围, 1, "门诊", 2, "住院", "") & "可用中药房", glngSys, decode(int范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0))
    End If
        
    '药品从系统指定的储备药房中找
    strSql = _
        " Select Distinct C.ID" & _
        " From " & IIF(lng药品ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
        " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
        " And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([2],3)") & " And B.部门ID=C.ID" & _
        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
        IIF(int范围 <> 3, " And (A.病人来源 is NULL Or A.病人来源=[2])", "") & _
        IIF(lng科室id <> 0, " And (A.开单科室ID is NULL Or A.开单科室ID=[3])", "") & _
        IIF(lng药品ID <> 0, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]")
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", str药房, int范围, lng科室id, lng药品ID, lng项目id)
    Do While Not rsTmp.EOF
        If str可用药房 = "" Then
            str药房IDs = str药房IDs & "," & rsTmp!ID
        ElseIf InStr("," & str可用药房 & ",", "," & rsTmp!ID & ",") > 0 Then
            str药房IDs = str药房IDs & "," & rsTmp!ID
        End If
        rsTmp.MoveNext
    Loop
    Get可用药房IDs = Mid(str药房IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get可用发料部门IDs(ByVal lng材料ID As Long, ByVal lng科室id As Long, Optional ByVal int范围 As Integer = 2, Optional ByVal lng项目id As Long) As String
'功能：获取卫材的有效诊疗执行科室ID串,用于判断缺省执行科室
'参数：lng科室ID=病人科室ID
'      int范围=1-门诊,2-住院(缺省)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, str发料部门IDs As String
    
    strSql = _
        " Select Distinct C.ID" & _
        " From " & IIF(lng材料ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
        " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
        " And " & IIF(mint范围 = 3, "Nvl(B.服务对象,0)<>0", "B.服务对象 IN([1],3)") & " And B.部门ID=C.ID " & IIF(lng材料ID <> 0, " And A.收费细目ID=[3]", " And A.诊疗项目ID=[4]") & _
        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
        IIF(int范围 <> 3, " And (A.病人来源 is NULL Or A.病人来源=[1])", "") & _
        IIF(lng科室id <> 0, " And (A.开单科室ID is NULL Or A.开单科室ID=[2])", "")
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", int范围, lng科室id, lng材料ID, lng项目id)
    Do While Not rsTmp.EOF
        str发料部门IDs = str发料部门IDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    Get可用发料部门IDs = Mid(str发料部门IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
