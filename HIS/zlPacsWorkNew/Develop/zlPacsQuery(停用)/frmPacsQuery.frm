VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "*\A..\idking\zlIDKind.vbp"
Begin VB.Form frmPacsQuery 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer TimFlicker 
      Interval        =   500
      Left            =   120
      Top             =   1560
   End
   Begin VB.PictureBox PicLine 
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   5775
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6000
      Width           =   5775
   End
   Begin VB.TextBox txtDetail 
      BackColor       =   &H00FDD6C6&
      Height          =   615
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "frmPacsQuery.frx":0000
      Top             =   7560
      Width           =   5775
   End
   Begin VB.PictureBox picListRowInfo 
      Height          =   615
      Left            =   720
      ScaleHeight     =   555
      ScaleWidth      =   5715
      TabIndex        =   10
      Top             =   6840
      Width           =   5775
      Begin VB.Label labPatientInfoName 
         AutoSize        =   -1  'True
         Caption         =   "??? ? ???"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   375
         Index           =   0
         Left            =   4920
         Top             =   120
         Width           =   495
      End
      Begin VB.Label labPatientInfoNo 
         AutoSize        =   -1  'True
         Caption         =   "检查号：99999999 标识号：12345678"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   11
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.PictureBox picHistory 
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   5715
      TabIndex        =   7
      Top             =   6240
      Width           =   5775
      Begin VB.ComboBox cboHistory 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "历史检查："
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.PictureBox picVsf 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   720
      ScaleHeight     =   3975
      ScaleWidth      =   5775
      TabIndex        =   4
      Top             =   1920
      Width           =   5775
      Begin VB.PictureBox picGroup 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   5115
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Bindings        =   "frmPacsQuery.frx":000F
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   5175
         _cx             =   9128
         _cy             =   5318
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
         OwnerDraw       =   4
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
   Begin VB.PictureBox picFilter 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   720
      ScaleHeight     =   615
      ScaleWidth      =   5775
      TabIndex        =   3
      Top             =   1200
      Width           =   5775
      Begin XtremeCommandBars.CommandBars cbrFilter 
         Left            =   0
         Top             =   120
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picSearch 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   5775
      TabIndex        =   1
      Top             =   600
      Width           =   5775
      Begin zlIDKind.PatiIdentify patiSearch 
         Bindings        =   "frmPacsQuery.frx":0023
         Height          =   300
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPacsQuery.frx":0037
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   0
         ShowSortName    =   -1  'True
         DefaultCardType =   "就诊卡"
         IDKindWidth     =   1200
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
      Begin XtremeCommandBars.CommandBars cbrBaseFilter 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picTag2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin XtremeSuiteControls.TabControl tabQuery 
      Bindings        =   "frmPacsQuery.frx":00EA
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   0
      Width           =   4485
      _Version        =   589884
      _ExtentX        =   7911
      _ExtentY        =   873
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":00FE
            Key             =   "复选留空"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":0698
            Key             =   "复选选中"
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":0C32
            Key             =   "定位"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":0FC4
            Key             =   "查找"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":1356
            Key             =   "单选留空"
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":1A68
            Key             =   "单选选中"
            Object.Tag             =   "90003"
         EndProperty
      EndProperty
   End
   Begin VB.Label labHint 
      AutoSize        =   -1  'True
      Caption         =   "----"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   15
      Top             =   8400
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmPacsQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const C_LAYOUT_BASEHEIGHTOFDETAILINFO As Long = 1200 ' 详细信息基准高度1200
Private Const C_LAYOUT_LISTLEFT As Long = 30 ' 列表左侧设计空出宽度 30
Private Const C_ICON_FIND As Long = 4
Private Const C_ICON_LOCATE As Long = 3
Private Const C_ICON_MENUCHOOSE As Long = 90001
Private Const C_ICON_MENUNOCHOOSE As Long = 90000
            
'某些功能必须参数变量
Private mblnRelatingPatient  As Boolean '是否启用关联病人
Private mblnAssignment As Boolean
Private mblSearching As Boolean 'DataSource操作中，会触发 selchange 要避免进入selchange 流程



'公共信息/系统常用变量参数
Private mcnOracle As ADODB.Connection
Private mlngUserId As Long
Private mlngModule As Long              '模块号
Private mstrCurRoom As String          '科室ID
Private mlngSys As Long
Private mstrDBUser As String                 '当前数据库用户
Private mbytFontSize As Byte                '字号
Private mfrmParent As Object

'当前查询相关

Private mTPatientBaseInfo As TPatientBaseInfo '病人信息（用于左下角病人信息、历史检查）
Private mrsData As ADODB.Recordset  '数据库查询出的记录集（不能做任何修改）
Private mrsDataShow As ADODB.Recordset 'mrsData经过一些转换后的记录集
Private mDTStart As Date
Private mDTEnd As Date
Private mPatiName  'pati控件参数名(用于pati控件的查询)
Private mTqueryType As TqueryType  '查询类型，用于查询条件赋值（LSQ待优化整理）
Private mintShowType As Integer '显示类型0：pacsMain    1：测试
Private mlngSortCol As Long               '当前进行排序的列
Private mintSortOrder As Integer         '当前进行排序的方式

Private WithEvents mobjSqlParse As clsSqlParse  '用于快速过滤参数值的获取
Attribute mobjSqlParse.VB_VarHelpID = -1

''方案信息
Private mstrListKeyCol As String '列表关键列  比如"医嘱ID"
Private mstrCachePath As String

Private mTColSort As TColSort   '排序信息

Private mTPatiIdentifyInfo As TPatiIdentifyInfo

Private mPicDictionary As Scripting.Dictionary    '图标缓存
Private mTQuickFilterState As TQuickFilterState   '用于快速过滤处理（选中状态缓存）
Private mlngSchemeNo As Long '当前使用的方案号
Private mColCfgInfo() As Integer    '列配置信息（只要列表改变列顺序后应该更新这个变量，用于快速根据当前列表列序号找到对应的列配置）

Private mSqlScheme As clsSqlScheme '当前方案
Private mstrSchemeCfg As TSchemeCfg '方案配置参数（查询条件、快速过滤、列头等）

Private WithEvents mObjQuery As clsPacsQuery
Attribute mObjQuery.VB_VarHelpID = -1


''当前列信息
Private mTStudyInfo As TStudyInfo '检查信息
Private mlngAdviceID As Long '当前医嘱ID(选中医嘱ID)

''''''''''''''''''''窗体自身控件、界面展示相关变量
Private mDataGrid As VSFlexGrid
Private mlngMove As Long '界面布局位置相关
Private mTLayout As TLayout

'必要引用
Private mobjSquareCard As Object

'''''''''''''事件
Public Event OnListRowSelClear() '列表选中项目重置（例如经过快速过滤后导致列表不显示数据，此时需要同步更新一些状态）
Public Event OnColStatistics(ByVal strStatisticsInfo As String)   '进行列统计
Public Event OnDblClick() '双击
Public Event OnRefreshSelectTab(ByVal lngAdviveID As Long)  '保持原来的功能
Public Event OnSelectScheme(ByVal strName As String)
Public Event OnSelChange()
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Type TPatiIdentifyInfo
    blFind As Boolean '是否查找   true:  查找    false:定位
    blHaveLoad As Boolean '是否已经加载过
    blIsFinding As Boolean '正在进行查找功能
    blShowPatiIdentify As Boolean '是否显示刷卡控件
    
    strFindItems As String '查找项目串
    strLocateItems As String '定位项目串
    strFindItem As String '当前查找项目
    strLocateItem As String '当前定位项目
    strDefault As String '默认查找项
    
    
End Type

'界面布局相关
Private Type TLayout
    blShowTimeSelect As Boolean '是否显示时间选择菜单
    blShowBaseFilter As Boolean '基本过滤（时间，Pati控件）
    blShowQuickFilter As Boolean '快速过滤
    blShowHistory As Boolean '历史检查下拉框
End Type


'行颜色信息
Private Type TColSort
    LngSchemeNo As Long '当前方案号
    dictSortInfo As Dictionary
End Type

'行颜色信息
Private Type TRowColorInfo
    LngSchemeNo As Long '当前方案号
    intRowColorIndex As Long '涉及行颜色的列
    blHaveRowColor As Boolean
End Type

'闪烁超时信息
Private Type TFlickerInfo
    LngSchemeNo As Long '当前方案号
    strName As String '闪烁字段名 如： "检查过程"
    strInfo As String '详细信息 如"已登记,申请时间,30|已报到,采样时间,40|"
End Type

'列统计信息
Private Type TColTotalInfo
    strName As String
    intCount As Integer
End Type

Private Type TPatientBaseInfo
    lngAdviceID As Long '医嘱ID
    lngPatientID As Long '患者ID
    lngLinkId As Long        ' 关联ID
    lngBaby As Long  ' 婴儿
    lngMarkNum As Long '标志号
    
    strName As String '姓名
    strAge As String '年龄
    strSex As String '性别
End Type
                
Private Type TSchemeCfg
    strSearchCfg  As String '快速查询配置
    strFilterCfg As String  '快速过滤配置
    strListCfg As String  '列表配置(顺序、宽度、是否可见)
    strListCfgDefault As String '列表初始配置(顺序、宽度、是否可见)
    strListCfgDefaultColOrder As String '列表初始配置(只有列顺序)
End Type

Private Type TQuickFilterCmdItem
    intItemIndex As Integer '序号（1,2,3,4...,99）
    blChoose As Boolean '是否选中
    strName As String '项目名称
    strFilterValue As String '项目过滤内容 "a,b,c,d"
End Type
    
Private Type TQuickFilterCmdState
'关联过滤定义 比如 "影像类别"->"部位"
'对影像类别来说 intRelation=1 strRelationName="部位"
'对部位来说 intRelation =2
    cmdItem() As TQuickFilterCmdItem '子项目信息
    
    intMenuIndex As Integer '序号
    intItemCount As Integer '子项目数量（这一个快速过滤菜单包含的可选项目）
    intRelation As Integer ' 1:关联中前者    2:关联中后者
    lngID As Long '主菜单ID
    
    strName As String '自身名称(字段名称)
    strRelationName As String '关联过滤名称
    strRelationChooseMenu As String '动态过滤的选用菜单"头;上肢;五官" 要使用名字
    strRelationValueForVBSFilter As String '组合好的用于VBS过滤的条件"1,2,3,a,b,c,d,e",在菜单选中项改变时变化，直接用于VBS过滤
    strCustomScript As String
    strMenuSQL As String
    
    blSimpleFilter As Boolean '是否简单过滤（是：显示的字段是参与过滤的字段   否：显示字段不参与过滤，菜单Category属性参与过滤 ）
    blSingleChoose As Boolean '过滤项是否单选，默认多选

End Type

Private Type TQuickFilterState
    intQuickFilterMenuCount As Integer '过滤项目数量
    TCmdState() As TQuickFilterCmdState '子项目信息
End Type

Private Enum TqueryType
    更新一行 = 0
    过滤 = 1
    刷新 = 2
    查找 = 3
End Enum

Property Get rsDataShow() As ADODB.Recordset
    If Not mrsDataShow Is Nothing Then Set rsDataShow = mrsDataShow
End Property

Property Get rsData() As ADODB.Recordset
    If Not mrsData Is Nothing Then Set rsData = mrsData
End Property

Property Get objSqlScheme() As clsSqlScheme
    Set objSqlScheme = mSqlScheme
End Property

Property Get objQuery() As clsPacsQuery
    Set objQuery = mObjQuery
End Property

Property Get DataGrid() As VSFlexGrid
    Set DataGrid = mDataGrid
End Property

Public Sub SetVars(ByVal strVarName As String, ByVal Value As Variant)
'需要的变量和参数赋值

    Select Case strVarName
        Case varName_数据库连接
            Set mcnOracle = Value
        Case varName_模块号
            mlngModule = Value
        Case varName_用户ID
            mlngUserId = Value
        Case varName_科室ID
            mstrCurRoom = Value
        Case varName_查询方案ID
            mlngSchemeNo = Value
        Case varName_查询界面类型
            mintShowType = Value
        Case varName_列表关键字
            mstrListKeyCol = Value
        Case varName_系统号
            mlngSys = Value
        Case varName_数据库用户名
            mstrDBUser = Value
        Case varName_字号
            mbytFontSize = Value
        Case varName_父窗体
            Set mfrmParent = Value
        Case varName_是否启用关联病人
            mblnRelatingPatient = Value
        Case Else
            MsgBox "[SetVars]" & vbLf & "参数:[" & strVarName & "]不存在", vbOKOnly, "异常"
    End Select

End Sub

'Public Function SetVars( _
'    ByVal cnOracle As ADODB.Connection, _
'    Optional ByVal lngModule As Long = 0, _
'    Optional ByVal lngUserId As Long = 0, _
'    Optional ByVal strCurRoom As String = "0", _
'    Optional ByVal lngSchemeId As Long = 0, _
'    Optional ByVal intShowType As Integer = 0, _
'    Optional ByVal lngSys As Long = 0, _
'    Optional ByVal strDBUser As String = 0, _
'    Optional ByVal bytFontSize As String = 0)
''需要的变量和参数赋值
'
'    Set mcnOracle = cnOracle
'    mlngModule = lngModule
'    mlngUserId = lngUserId
'    mstrCurRoom = strCurRoom
'    mlngSchemeNo = lngSchemeId
'    mintShowType = intShowType
'    mstrListKeyCol = "医嘱ID"
'
'    mlngSys = lngSys
'    mstrDBUser = strDBUser
'    mbytFontSize = bytFontSize
'End Function

Public Function init() As Boolean
'首先应该调用的函数
'查询界面初始化
On Error GoTo errHandle
    Dim i As Integer
    Dim intIndex As Integer
    
    Set mObjQuery = New clsPacsQuery
    
    mObjQuery.init mcnOracle, mlngUserId, mstrCachePath
    mObjQuery.LoadQueryScheme mlngModule
    

    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    
    mobjSquareCard.zlInitComponents Me, mlngModule, mlngSys, mstrDBUser, gcnOracle
    patiSearch.zlInit Me, mlngSys, mlngModule, gcnOracle, mstrDBUser, mobjSquareCard, InitCardType("姓名;")
    
    With tabQuery
        .RemoveAll
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll

        For i = 1 To mObjQuery.SchemeCount
            If mObjQuery.SchemeInfo(i).IsDefault Or mObjQuery.SchemeInfo(i).IsOften Then
                Call .InsertItem(i, mObjQuery.SchemeInfo(i).Name, picTag2.hwnd, 0)
                .Item(.ItemCount - 1).Tag = mObjQuery.SchemeInfo(i).SchemeId
                If mObjQuery.SchemeInfo(i).IsDefault Then
                    intIndex = .ItemCount - 1
                End If
            End If
        Next i

        If .ItemCount >= 1 Then
            If intIndex > 0 Then
                .Item(intIndex).Selected = True
            Else
                .Item(0).Selected = True
            End If
        Else
            Call HaveNoScheme
        End If

    End With
    
    Call ReSetFormFontSize
    
    Exit Function
errHandle:
    Err.Raise -1, "frmPacsQuery", "[init]" & vbCrLf & Err.Description
End Function

Private Sub HaveNoScheme()
'没有任何方案：一些控件不可见，一些控件enable=false
On Error GoTo errHandle
    picSearch.Visible = False
    picFilter.Visible = False
    picVsf.Visible = False
    picHistory.Visible = False
    picListRowInfo.Visible = False
    txtDetail.Visible = False
    
    labHint = "未找到有效的查询方案，请先配置方案"
    labHint.Visible = True
    Call labHint.Move(0.5 * (Me.Width - labHint.Width), 0.5 * (Me.Height - labHint.Height))
errHandle:
End Sub

Private Sub cboHistory_DropDown()
On Error GoTo errHandle
    Call SendMessage(cboHistory.hwnd, &H160, 500, 0)
errHandle:
End Sub

Private Sub cboHistory_Click()
On Error GoTo errHandle
    Dim lngAdviceID As Long
    
    If cboHistory.ListCount <= 1 Then Exit Sub
    If cboHistory.Tag = "" Then Exit Sub '此时 cboHistory 项目未增加完成，属listindex 赋值触发
    
    lngAdviceID = cboHistory.ItemData(cboHistory.ListIndex)
    
    If lngAdviceID = mTStudyInfo.lngAdviceID Then
        Call vsfList_SelChange
        Exit Sub  '当次与当前选中医嘱ID相同时不由本函数控制
    End If
    
    '根据病人提供不同选项卡
    RaiseEvent OnRefreshSelectTab(lngAdviceID)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbrBaseFilter_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim objControl As CommandBarControl
    Dim objCboControl As CommandBarComboBox
    Dim objfrmTimeSet As frmTimeSet
    Dim DTStartNew As Date
    Dim DTEndNew As Date
    
    Select Case control.Id
        Case conMenu_PacsQuery_TimeLab
        Case conMenu_PacsQuery_TimeCbo

            Set objCboControl = cbrBaseFilter.FindControl(xtpControlComboBox, conMenu_PacsQuery_TimeCbo)
            
            If objCboControl.Text = "自定义" Then   '自定义弹出时间选择窗
                '弹出自定义时间处理
                If objfrmTimeSet Is Nothing Then Set objfrmTimeSet = New frmTimeSet
                
                Call objfrmTimeSet.zlShowMe(mfrmParent, mDTStart, mDTEnd, mSqlScheme.dateRange)
                Call objfrmTimeSet.GetTimeSet(DTStartNew, DTEndNew)
                
                mDTStart = DTStartNew
                mDTEnd = DTEndNew
                
                objCboControl.ToolTipText = "自定义时间范围:" & DTStartNew & "到" & DTEndNew
                
                mstrSchemeCfg.strSearchCfg = Split(mstrSchemeCfg.strSearchCfg, ",")(0) & "," & mDTStart & "," & _
                mDTEnd & "," & Split(mstrSchemeCfg.strSearchCfg, ",")(3)
            Else
                objCboControl.ToolTipText = "时间选择"
            End If
            
            If objCboControl.Text <> "无限制" Then  '不是无限制则更新保存的时间选择
                mstrSchemeCfg.strSearchCfg = objCboControl.Text & "," & Split(mstrSchemeCfg.strSearchCfg, ",")(1) & "," & _
                Split(mstrSchemeCfg.strSearchCfg, ",")(2) & "," & Split(mstrSchemeCfg.strSearchCfg, ",")(3)
            End If
            
        
        Case conMenu_PacsQuery_FindWay
            Dim blFindWayOld As Boolean
            
            blFindWayOld = mTPatiIdentifyInfo.blFind

            mTPatiIdentifyInfo.blFind = Not mTPatiIdentifyInfo.blFind
            
            If blFindWayOld <> mTPatiIdentifyInfo.blFind Then
                Call DoPatiIdentify
            End If
            
            Call SaveLocalPara_PatiIdentify
        Case conMenu_PacsQuery_PatiControl
        Case conMenu_PacsQuery_Do
            If mTPatiIdentifyInfo.blFind Then
                Call ExecuteQuery("查找")
            Else
                Call SeekNextPati(patiSearch.Tag <> patiSearch.Text, patiSearch.GetCurCard.名称, patiSearch.Text)
            End If
    End Select
    Exit Sub
errHandle:
    MsgBox "[cbrBaseFilter_Execute]" & vbCrLf & Err.Description, vbOKOnly, "异常"
End Sub

Private Sub cbrBaseFilter_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCboControl As CommandBarComboBox
    Dim objCusControl As CommandBarControlCustom
    
    Select Case control.Id
        Case conMenu_PacsQuery_TimeLab
        Case conMenu_PacsQuery_TimeCbo
        Case conMenu_PacsQuery_FindWay
            control.IconId = IIf(mTPatiIdentifyInfo.blFind, C_ICON_FIND, C_ICON_LOCATE)
        Case conMenu_PacsQuery_PatiControl
        Case conMenu_PacsQuery_Do
            control.Caption = IIf(mTPatiIdentifyInfo.blFind, "执行查找", "执行定位")
        Case Else
    End Select
    
    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[cbrBaseFilter_Update]" & vbCrLf & Err.Description
End Sub

Private Sub Form_Initialize()
    Set mDataGrid = vsfList
    Set mPicDictionary = New Dictionary
    
    mDataGrid.AllowUserResizing = flexResizeColumns '用于改遍列宽
    mDataGrid.ExplorerBar = 7 '用于列头拖动和排序
    mDataGrid.SelectionMode = flexSelectionListBox '用于选择整行
    mDataGrid.AllowSelection = False '用于选择整行
    mDataGrid.ScrollTrack = True '滚动条随时更新
    mDataGrid.FixedCols = 1
    mDataGrid.BackColorSel = &HFEE0E2      '&HFECFD2

    picHistory.BorderStyle = 0
    picListRowInfo.BorderStyle = 0
    
End Sub

Private Sub Label1_Click()
On Error GoTo errH
'    'LSQ 测试功能
'    Dim t1 As Long
'    Dim t2 As Long
'    Dim strTMp As String
'    Dim i As Long
'    Dim j As Long
'
'
'
'    Debug.Print ""
'    Debug.Print ""
'
   MsgBox "vsfList.MouseRow:" & vsfList.MouseRow
    MsgBox "pati名称2:" & patiSearch.Name
'patiSearch.Text = ""  '切换Item时，要将输入框清空
'    mPatiName = objCard.名称
    Exit Sub
errH:
    MsgBox "测试功能" & Err.Description
End Sub

Private Sub mobjSqlParse_OnGetParameterValue(ByVal strParName As String, Value As Variant)
'获取快速过滤的参数值
On Error GoTo errHandle
    Dim i As Integer
    Dim strValue As String
    Dim strValueAll As String
    Dim j As Integer
    Dim blChooseOne As Boolean
    
    blChooseOne = False
    
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        If mTQuickFilterState.TCmdState(i).strName = strParName Then
            Exit For
        End If
    Next
    
    For j = 1 To mTQuickFilterState.TCmdState(i).intItemCount
        If mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose Then
            strValue = IIf(Len(strValue) = 0, strValue, strValue & ",")
            strValue = strValue & mTQuickFilterState.TCmdState(i).cmdItem(j).strName
            blChooseOne = True
        End If
        strValueAll = IIf(Len(strValueAll) = 0, strValueAll, strValueAll & ",")
        strValueAll = strValueAll & mTQuickFilterState.TCmdState(i).cmdItem(j).strName
    Next
    
    If Not blChooseOne Then
        Value = strValueAll
    Else
        Value = strValue
    End If
    
    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[mobjSqlParse_OnGetParameterValue]" & vbCrLf & Err.Description
End Sub

Private Sub patiSearch_ItemClick(Index As Integer, objCard As zlIDKind.Card)
On Error GoTo errHandle
    
    If mblnAssignment Then Exit Sub
    patiSearch.Text = ""  '切换Item时，要将输入框清空
    mPatiName = objCard.名称
    
    If mTPatiIdentifyInfo.blFind Then
        mTPatiIdentifyInfo.strFindItem = mPatiName
    Else
        mTPatiIdentifyInfo.strLocateItem = mPatiName
    End If
    
    Call SaveLocalPara_PatiIdentify
    Exit Sub
errHandle:
    MsgBox "[patiSearch_ItemClick]" & vbCrLf & Err.Description, vbOKOnly, "异常"
End Sub


Public Sub StartReadCard()
On Error GoTo errHandle
'开始读卡
    Dim lngPatientID As Long
    Dim strCurCardName As String

    If mTPatiIdentifyInfo.blFind Then
        strCurCardName = mTPatiIdentifyInfo.strFindItem
    Else
        strCurCardName = mTPatiIdentifyInfo.strLocateItem
    End If

    If patiSearch.GetCurCard.接口序号 > 0 Then
        Call mobjSquareCard.zlGetPatiID(patiSearch.GetCurCard.接口序号, patiSearch.Text, , lngPatientID)

        Call OnFilterRead(strCurCardName, patiSearch.Text, IIf(lngPatientID > 0, lngPatientID, ""))
    Else
        Call OnFilterRead(strCurCardName, patiSearch.Text, "")
    End If
    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[StartReadCard]" & vbCrLf & Err.Description
End Sub

Private Sub OnFilterRead(ByVal strCardName As String, ByVal strFilter As String, ByVal strPatientId As String)
'开始查找数据
On Error GoTo errHandle

    If mTPatiIdentifyInfo.blFind Then
        '查找
        mTPatiIdentifyInfo.blIsFinding = True
        Call ExecuteQuery("查找")
        mTPatiIdentifyInfo.blIsFinding = False
        
'        If mrsData.RecordCount < 1 Then
'            Call MsgBoxD(Me, "未找到任何数据,请注意时间范围是否正确" & vbCrLf & "  查找类别:" & strCardName & vbCrLf & "  查找数据:" & strFilter, vbOKOnly, "提示")
'        Else
'            If vsfList.Rows <= 1 Then
'                Call MsgBoxD(Me, "查询到数据但未显示到列表中,请注意快速过滤条件", vbOKOnly, "提示")
'            End If
'        End If
    Else
        '定位
        Call SeekNextPati(patiSearch.Tag <> patiSearch.Text, patiSearch.GetCurCard.名称, patiSearch.Text)
    End If

    Call patiSearch.SetFocus
    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[OnFilterRead]" & vbCrLf & Err.Description
End Sub


Private Sub patiSearch_KeyPress(KeyAscii As Integer)
'录入事件
On Error GoTo errHandle
    Dim blnCard As Boolean
    Dim lngPatientID As Long

    If KeyAscii = 13 Then
        Call StartReadCard

        Exit Sub
    End If

'    If patiSearch.GetCurCard.是否刷卡 Then
'        blnCard = patiSearch.zlIsBrushCard(patiSearch.objTxtInput, KeyAscii)
'
'        If blnCard And Len(patiSearch.Text) = patiSearch.GetCardNoLen - 1 And KeyAscii <> 8 Then  '刷卡完毕处理
'            patiSearch.Text = patiSearch.Text & Chr(KeyAscii)
'
'            KeyAscii = 0
'
'            If patiSearch.GetCurCard.接口序号 > 0 Then
'                Call mobjSquareCard.zlGetPatiID(patiSearch.GetCurCard.接口序号, patiSearch.Text, , lngPatientID)
'
'                Call OnFilterRead(patiSearch.GetCurCard.名称, patiSearch.Text, IIf(lngPatientID > 0, lngPatientID, ""))
'            Else
'                Call OnFilterRead(patiSearch.GetCurCard.名称, patiSearch.Text, "")
'            End If
'        End If
'    End If
    
Exit Sub
errHandle:
    MsgBox "[cbrBaseFilter_Execute]" & vbCrLf & Err.Description, vbOKOnly, "异常"
End Sub

Private Sub PicLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'左下方详细信息高度可以改变
On Error GoTo errHandle
    
    If Button = 1 Then
        '当值达到一定范围就退出函数
        
        If PicLine.Top + Y < Me.Top + 3000 Or PicLine.Top + Y > Me.Height - picHistory.Height - picListRowInfo.Height Then
            Exit Sub
        End If

        picVsf.Height = picVsf.Height + Y
        PicLine.Top = PicLine.Top + Y
        picHistory.Top = picHistory.Top + Y
        picListRowInfo.Top = picListRowInfo.Top + Y
        txtDetail.Top = txtDetail.Top + Y
        txtDetail.Height = txtDetail.Height - Y
        mlngMove = txtDetail.Height - C_LAYOUT_BASEHEIGHTOFDETAILINFO
        
    End If
    
errHandle:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errH
    SaveLocalPara

    If mlngSchemeNo > 0 And mintShowType = 0 Then Call SaveShemeCustomCfg(mlngSchemeNo)
    
    Set mObjQuery = Nothing
    Set mrsData = Nothing
    Set mobjSqlParse = Nothing
    Set mTColSort.dictSortInfo = Nothing
    Set mPicDictionary = Nothing
    Set mDataGrid = Nothing
    Set mfrmParent = Nothing
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[Form_Unload]" & vbCrLf & Err.Description
End Sub

Private Sub mObjQuery_OnGetParameterValue(ByVal strParName As String, Value As Variant)
On Error GoTo errH
    Dim strTest As String
    Dim strValue As String
    

    Select Case strParName
        Case "系统.科室ID"
            Value = mstrCurRoom
'            Value = "123,123,23,24,25,26"
        Case "系统.医嘱ID"
            If mTqueryType = TqueryType.更新一行 Then
            ElseIf mTqueryType = TqueryType.刷新 Then
                Value = ""
            Else
                Value = ""
            End If
        Case "系统.开始日期"
            If mTqueryType = TqueryType.更新一行 Then
                Value = ""
            ElseIf mTqueryType = TqueryType.刷新 Then
            Else
            End If
        Case "系统.结束日期"
            If mTqueryType = TqueryType.更新一行 Then
                Value = ""
            ElseIf mTqueryType = TqueryType.刷新 Then
            Else
            End If
        Case Else
            If mTqueryType = 查找 Then
                If mPatiName = strParName Then
                    strValue = patiSearch.Text
                    Value = IIf(IsNumeric(strValue), Val(strValue), strValue)
                End If
            End If
            
    End Select
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[mObjQuery_OnGetParameterValue]" & vbCrLf & Err.Description
End Sub

Private Sub picHistory_Resize()
On Error Resume Next
    Dim lngLeft As Long
    
    Label1.Move 120, 80
    lngLeft = Label1.Left + Label1.Width + 120
    cboHistory.Move lngLeft, 30, picHistory.Width - lngLeft - 60
    
End Sub

Private Sub picListRowInfo_Resize()
On Error Resume Next
    Dim i As Integer, j As Integer
    Dim lngLeft As Long
    
    labPatientInfoName.Move C_LAYOUT_LISTLEFT, C_LAYOUT_LISTLEFT, labPatientInfoName.Width, labPatientInfoName.Height
    labPatientInfoNo.Move labPatientInfoName.Left + labPatientInfoName.Width + 2 * C_LAYOUT_LISTLEFT, C_LAYOUT_LISTLEFT, labPatientInfoNo.Width, labPatientInfoNo.Height

    For i = 0 To imgState.Count - 1
        '重新设置位置
        lngLeft = picListRowInfo.Width
        For j = 0 To i
            lngLeft = lngLeft - imgState(i).Width
        Next
        Call imgState(i).Move(lngLeft, 0)
    Next
    
End Sub

Private Sub picVsf_Resize()
On Error Resume Next
    Call vsfList.Move(0, 0, picVsf.Width, picVsf.Height)
End Sub

Private Sub tabQuery_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errH
'本过程大多数处理有先后顺序，改动需谨慎
'顺序 ： 保存参数，更新方案号，加载参数 ，设置列表，界面布局刷新
    Dim i As Long
    
    
    Call SaveLocalPara
    
    If mlngSchemeNo > 0 And mintShowType = 0 Then Call SaveShemeCustomCfg(mlngSchemeNo)
    
    '变量初始化
    mstrSchemeCfg.strListCfgDefault = ""
    mstrSchemeCfg.strListCfgDefaultColOrder = ""
    mTPatiIdentifyInfo.blHaveLoad = False
    mTPatiIdentifyInfo.strFindItems = ""
    mTPatiIdentifyInfo.strLocateItems = ""
    mTLayout.blShowBaseFilter = False
    mTLayout.blShowTimeSelect = False
    mTPatiIdentifyInfo.blShowPatiIdentify = False
    
    
    
    ReDim mColCfgInfo(0)
    
    mlngSchemeNo = Item.Tag
    
    Call GetLocalPara
    
    Call mObjQuery.ChangeCurScheme(mlngSchemeNo)
    Set mSqlScheme = mObjQuery.GetSqlScheme(mlngSchemeNo)
    
    Call GetSchemePara
    

    Call LoadShemeCustomCfg(mlngSchemeNo)
    Call RefreshQueryWindow(mlngSchemeNo)
    
    
    
    Call ReSetFormFontSize(mbytFontSize)
    
    Call Form_Resize
    
    RaiseEvent OnSelectScheme(Item.Caption)
    Call ExecuteQuery("刷新", 1)
    
    Exit Sub
errH:
    MsgBox Err.Description & "tabQuery_SelectedChanged"
End Sub

Private Sub TimFlicker_Timer()
On Error GoTo errH
'   超时闪烁的处理
    Dim i As Integer, j As Integer
    Dim lngCol As Long, lngColContrast As Long
    Dim strTmp As String
    Dim lngStateColor As Long, lngNextStateColor As Long, lngPreStateColor As Long
    Dim objRowRelation As New clsScRowRelation
    
    Static intSta As Integer
    Static TPFlickerInfo As TFlickerInfo '超时闪烁配置
    
    '方案第一次加载时获取超时闪烁相关信息
    If TPFlickerInfo.LngSchemeNo <> mlngSchemeNo Then
        TPFlickerInfo.strName = ""
        TPFlickerInfo.strInfo = ""
    
        If mSqlScheme Is Nothing Then Exit Sub
        TPFlickerInfo.LngSchemeNo = mlngSchemeNo
        
        For i = 1 To mSqlScheme.ShowCfgCount
            For j = 1 To mSqlScheme.ShowCfg(i).RowRelationCount
                Set objRowRelation = mSqlScheme.ShowCfg(i).RowRelation(j)
                
                If objRowRelation.FlickerTimeOut > 0 Then
                    TPFlickerInfo.strName = mSqlScheme.ShowCfg(i).Name
                    TPFlickerInfo.strInfo = TPFlickerInfo.strInfo & objRowRelation.TiggerData & "," & objRowRelation.TimeOutReferCol & "," & objRowRelation.FlickerTimeOut & "|"

                End If
            Next
        Next
        
        intSta = 0
        Exit Sub
        
    End If
    
    intSta = intSta + 1
    If intSta = 4 Then intSta = 1

    lngCol = vsfList.ColIndex(TPFlickerInfo.strName)
    If vsfList.TopRow = vsfList.BottomRow Then Exit Sub
    For i = vsfList.TopRow To vsfList.BottomRow   '遍历可见行  For 1
        For j = 0 To UBound(Split(TPFlickerInfo.strInfo, "|")) - 1 '判断是否满足超时条件 For 2
            strTmp = Split(TPFlickerInfo.strInfo, "|")(j)
            If Split(strTmp, ",")(0) = vsfList.TextMatrix(i, lngCol) Then
                lngColContrast = vsfList.ColIndex(Split(strTmp, ",")(1))
                
                If IsDate(vsfList.TextMatrix(i, lngColContrast)) Then
                
                    If DateDiff("N", vsfList.TextMatrix(i, lngColContrast), Now) >= Val(Split(strTmp, ",")(2)) Then    '若满足设置的超时时间
                    
                        '首先测试闪烁功能
                        lngStateColor = vsfList.Cell(flexcpBackColor, i, 0)
                        lngNextStateColor = RGB(200, 0, 0)
                        lngPreStateColor = RGB(0, 0, 0)
    
                        If intSta = 1 Then
                            vsfList.Cell(flexcpBackColor, i, 0) = lngPreStateColor
                        ElseIf intSta = 2 Then
                            vsfList.Cell(flexcpBackColor, i, 0) = lngStateColor
                        Else
                            vsfList.Cell(flexcpBackColor, i, 0) = lngNextStateColor
                        End If
                    End If
                End If
                
                Exit For   '若满足超时条件 退出For 2
            End If
        Next
    Next
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[TimFlicher_Timer]" & vbCrLf & Err.Description
End Sub

Private Sub vsfList_AfterSort(ByVal Col As Long, Order As Integer)
'排序后需要更新列表可见行行关联设置
On Error GoTo errH
    Dim RowIndex As Long
    
    mlngSortCol = Col
    mintSortOrder = Order
    
    If vsfList.TopRow = vsfList.BottomRow Then Exit Sub
    For RowIndex = vsfList.TopRow To vsfList.BottomRow
        Call RefreshRowRelation(RowIndex)
    Next
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[vsfList_AfterSort]" & vbCrLf & Err.Description
End Sub

Private Sub vsfList_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
On Error GoTo errH
'获取之后的列表显示范围，用于控制执行RefreshRowRelation
    Dim lngHeight As Long
    Dim RowIndex As Long
    Dim LngListBottom As Long
    
    
'    Debug.Print "当前方案号" & mlngSchemeNo
'    Debug.Print "当前方案号" & mSqlScheme.SchemeId
'    Debug.Print "当前方案名称" & mSqlScheme.SchemeName
'    Debug.Print "当前数量1：" & mSqlScheme
'    Debug.Print "当前相关缓存"
'    Debug.Print "实际数量1：" & mSqlScheme.SchemeName
    
    
    lngHeight = vsfList.BottomRow - vsfList.TopRow
    
    LngListBottom = NewTopRow + lngHeight
    If LngListBottom > vsfList.Rows - 1 Then LngListBottom = vsfList.Rows - 1
    
    For RowIndex = NewTopRow To LngListBottom
        Call RefreshRowRelation(RowIndex)
    Next
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[vsfList_BeforeScroll]" & vbCrLf & Err.Description
End Sub

Private Sub vsfList_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo errH
    Dim lngOrder As Long
    
    If Col <> vsfList.ColIndex(GetColSort(vsfList.ColKey(Col))) Then
        
        '排序列也可以有两种方向的排序
        If mintSortOrder = 2 Or mintSortOrder = 4 Or mintSortOrder = 6 Or mintSortOrder = 8 Then
            lngOrder = 3
        Else
            lngOrder = 4
        End If
        
        '使用了排序列需要单独排序，后面需要将Order 设置为0 避免执行自带的排序
        Call SetOrder(vsfList.ColIndex(GetColSort(vsfList.ColKey(Col))), lngOrder)
        
        mlngSortCol = vsfList.ColIndex(GetColSort(vsfList.ColKey(Col)))
        mintSortOrder = lngOrder
    
        Order = 0
    End If
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[vsfList_BeforeSort]" & vbCrLf & Err.Description
End Sub

Private Sub vsfList_DblClick()
    RaiseEvent OnDblClick
End Sub

Private Sub vsfList_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim rc As Rect
    
    If Col = 0 And Row > 0 Then
        rc.Bottom = Bottom
        rc.Left = Left
        rc.Right = Right
        rc.Top = Top
        
        Call DrawText(hDC, Row, Len("" & Row), rc, 0)
         
    End If
End Sub

Private Sub vsfList_SelChange()
'列表选中行改变
'1 根据医嘱ID查询基本信息并且保存到列表中
'2 刷新列表下方控件显示内容
'3 更新当前列表结构数据
On Error GoTo errH
    Dim intCol As Integer
    Dim lngAdviceID As Long
    Static lngListSelectRow As Long

    If mblSearching Then Exit Sub 'datasource 操作避免进入后面流程
    
    If mSqlScheme Is Nothing Then Exit Sub

    If vsfList.MouseRow < 1 And vsfList.Row < 1 Then Exit Sub

        
    If vsfList.MouseRow > 0 Then
        '手动点击进入selchange
        If lngListSelectRow = vsfList.MouseRow Then
            Exit Sub
        Else
            lngListSelectRow = vsfList.MouseRow
        End If
    Else
        '过滤、刷新等操作进入selchange
        lngListSelectRow = vsfList.RowSel
    End If

    '状态图
    Call DoStateImage(lngListSelectRow)
    intCol = vsfList.ColIndex(mstrListKeyCol)
    If intCol = -1 Then Exit Sub

    mlngAdviceID = Val(vsfList.TextMatrix(lngListSelectRow, intCol))
    mTStudyInfo.lngAdviceID = mlngAdviceID

'    根据是否已经有检查数据判断是否需要重新查询
    If IsEmpty(vsfList.Cell(flexcpData, lngListSelectRow)) Then
        Call GetStudyInfo
        vsfList.Cell(flexcpData, lngListSelectRow) = mTStudyInfo
    Else
        mTStudyInfo = vsfList.Cell(flexcpData, lngListSelectRow)
        mTStudyInfo.lngAdviceID = mlngAdviceID
    End If
    RaiseEvent OnSelChange

    Call FillHistoryStudy
    Call FillCurAdviceTxtInfor
    Call FillCurAdviceAppend(lngListSelectRow)
    Call SetSelectRowFont
    
    Exit Sub
errH:
    MsgBox "[vsfList_SelChange]" & vbCrLf & Err.Description, vbOKOnly, "异常"
End Sub

Private Sub SetSelectRowFont()
'选中行设置字体加粗，其他取消字体加粗
On Error GoTo errH
    
    With vsfList
        
        If .RowSel < 0 Then Exit Sub
    
        If .Cols > 2 And .Rows > 1 Then
            .Cell(flexcpFontBold, .TopRow, 1, .BottomRow, .Cols - 1) = False
            .Cell(flexcpFontBold, .RowSel, 1, .RowSel, .Cols - 1) = True
        End If
    End With
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[SetSelectRowFont]" & vbCrLf & Err.Description
End Sub

Private Sub GetRGB(ByVal lngColor As Long, lngR As Long, lngG As Long, lngB As Long)
On Error GoTo errH
    Dim lngMinVal As Long
    Dim lngMaxVal As Long
    
    lngMinVal = 80
    lngMaxVal = 225
    
    lngR = lngColor Mod 256
    
    If lngR <= lngMinVal Then
        lngR = lngMinVal
    ElseIf lngR > lngMaxVal Then
        lngR = lngMaxVal
    End If
    
    lngG = (Fix(lngColor \ 256)) Mod 256
 
    If lngG <= lngMinVal Then
        lngG = lngMinVal
    ElseIf lngG > lngMaxVal Then
        lngG = lngMaxVal
    End If
    
    lngB = Fix(lngColor \ 256 \ 256)
 
    If lngB <= lngMinVal Then
        lngB = lngMinVal
    ElseIf lngB > lngMaxVal Then
        lngB = lngMaxVal
    End If
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[GetRGB]" & vbCrLf & Err.Description
End Sub

Private Sub FillCurAdviceAppend(ByVal lngListSelectRow As Long, Optional blIsClear As Boolean = False)
'填充左下角详细信息，需要增加一些处理，例如判断信息显示的条件
On Error GoTo errHandle
    Dim i As Integer
      
    txtDetail = ""
    If blIsClear Then Exit Sub
    
    If vsfList.Rows = 1 Or vsfList.Cols < 2 Then Exit Sub
    For i = 2 To vsfList.Cols
        txtDetail = txtDetail & vsfList.TextMatrix(0, i - 1) & ":  " & LTrim(vsfList.TextMatrix(lngListSelectRow, i - 1)) & vbNewLine
    Next

    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[FillCurAdviceAppend]" & vbCrLf & Err.Description
End Sub

Private Sub GetStudyInfo()
'获取病人基本信息
On Error GoTo errHandle

    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strTemp As String
    
    If mlngModule <> G_LNG_PATHSTATION_MODULE Then
        strSql = "select A.ID 医嘱ID,A.病人ID,A.开嘱时间,A.医嘱内容,A.姓名,A.性别,A.年龄,B.执行状态,B.执行过程,C.关联ID,C.检查号" & _
               " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 C" & _
               " Where A.ID = [1] And A.相关id Is Null And B.医嘱ID=A.ID " & _
               " AND A.ID=C.医嘱ID(+)"
    Else
        strSql = "select A.ID 医嘱ID,A.病人ID,A.开嘱时间,A.医嘱内容,A.姓名,A.性别,A.年龄,B.执行状态,B.执行过程,C.关联ID,D.病理号 " & _
               " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 C,病理检查信息 D" & _
               " Where A.ID = [1] And A.相关id Is Null And B.医嘱ID=A.ID " & _
               " AND A.ID=C.医嘱ID(+) and C.医嘱ID=D.医嘱ID(+)"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询Pacs病人基本信息", mTStudyInfo.lngAdviceID)

    With mTStudyInfo
        .strPatientAge = NVL(rsTemp!年龄)
        .strPatientName = NVL(rsTemp!姓名)
        .strPatientSex = NVL(rsTemp!性别)
        .strStudyNum = NVL(rsTemp(GetStudyNumberDisplayName))
        .lngLinkId = NVL(rsTemp!关联ID, 0)
        .lngPatId = rsTemp!病人ID
    End With

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function GetStudyNumberDisplayName() As String
'获取检查号码显示名称
    GetStudyNumberDisplayName = IIf(mlngModule = G_LNG_PATHSTATION_MODULE, "病理号", "检查号")
End Function

Private Sub FillHistoryStudy()
'填充历史检查记录
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strTemp As String
    
    If mTStudyInfo.lngAdviceID = 0 Then
        cboHistory.Clear
        Exit Sub
    End If

    cboHistory.Tag = "" 'cboHistory，用于区别是"增加项目"时触发还是"cboHistory"触发
    
    If mlngModule <> G_LNG_PATHSTATION_MODULE Then
        strSql = "select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
               " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 C" & _
               " Where A.病人id = [1] And A.相关id Is Null And B.医嘱ID=A.ID " & _
               " AND A.ID=C.医嘱ID And Instr([2],A.执行科室id ) >0"
    Else
        strSql = "select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
               " From 病人医嘱记录 A,病人医嘱发送 B,病理检查信息 C" & _
               " Where A.病人id = [1] And A.相关id Is Null And B.医嘱ID=A.ID " & _
               " AND A.ID=C.医嘱ID And Instr([2],A.执行科室id ) >0 "
    End If
              
    '启用关联病人，才查询关联ID
    If mblnRelatingPatient = True And mTStudyInfo.lngLinkId <> 0 Then
        If mTStudyInfo.lngLinkId <> 0 Then
            If mlngModule <> G_LNG_PATHSTATION_MODULE Then
                strSql = strSql & " union select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
                    " From 病人医嘱记录 A " & _
                    " Where A.id in (select 医嘱ID from 影像检查记录 Where 关联ID =[3]) "
            Else
                strSql = strSql & " union select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
                    " From 病人医嘱记录 A, 病理检查信息 B " & _
                    " Where  A.id in (select 医嘱ID from 影像检查记录 Where 关联ID =[3]) and a.id=b.医嘱ID "
            End If
        End If
    End If
    
    strTemp = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
    strTemp = Replace(strTemp, "病人医嘱发送", "H病人医嘱发送")
    strTemp = Replace(strTemp, "影像检查记录", "H影像检查记录")
    strSql = strSql & vbNewLine & " Union ALL " & vbNewLine & strTemp
    strSql = "select * From (" & vbNewLine & strSql & vbNewLine & ") Order By 开嘱时间 Asc"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", mTStudyInfo.lngPatId, mstrCurRoom, mTStudyInfo.lngLinkId)

    cboHistory.Clear
    
    If rsTemp.RecordCount > 50 Then
        If MsgBox("检测到本条医嘱的历史记录超过50，选[是]继续加载，选[否]将不会加载历史检查", vbYesNo + vbDefaultButton2, "提示") = vbNo Then Exit Sub
    End If
    
    Do Until rsTemp.EOF
        If rsTemp!医嘱ID = mTStudyInfo.lngAdviceID Then
            '当前
            cboHistory.AddItem "●第" & rsTemp.AbsolutePosition & "次/共" & rsTemp.RecordCount & "次(" & Format(rsTemp!开嘱时间, "yyyy-mm-dd") & ")  " & _
            Trim(rsTemp!医嘱内容)
        Else
            cboHistory.AddItem "  第" & rsTemp.AbsolutePosition & "次/共" & rsTemp.RecordCount & "次(" & Format(rsTemp!开嘱时间, "yyyy-mm-dd") & ")  " & _
            Trim(rsTemp!医嘱内容)
        End If
        
        cboHistory.ItemData(cboHistory.NewIndex) = rsTemp!医嘱ID
       
        If rsTemp!医嘱ID = mTStudyInfo.lngAdviceID Then cboHistory.ListIndex = cboHistory.NewIndex
        
        rsTemp.MoveNext
    Loop
    
    If cboHistory.ListCount > 1 Then
        cboHistory.ForeColor = &HC0&
    Else
        cboHistory.ForeColor = &H80000008
    End If
    
    cboHistory.Tag = "完成加载"

Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub FillCurAdviceTxtInfor()
'填充右上方病人基本信息
On Error GoTo errHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim intChargeState As Integer
    Dim intColIndex As Integer
    Dim blnQueryMoneyState As Boolean

    If mTStudyInfo.lngAdviceID <= 0 Then
        labPatientInfoName = "姓名:  性别:  年龄:"
        labPatientInfoNo = "[" & GetStudyNumberDisplayName & ":--- ]"
        Call picListRowInfo_Resize
        Exit Sub
    End If
    
    labPatientInfoName = mTStudyInfo.strPatientName & " " & mTStudyInfo.strPatientSex & " " & mTStudyInfo.strPatientAge

    If mTStudyInfo.lngAdviceID > 0 Then
        labPatientInfoNo.Caption = "[" & GetStudyNumberDisplayName & ":" & IIf(mTStudyInfo.strStudyNum <> "-1", mTStudyInfo.strStudyNum, "--- ") & "]"


'lsq 待增加婴儿病人的处理
'            If mcurAdviceInf.lngBaby <> 0 Then
'
'                strSql = "select Nvl(A.婴儿姓名, B.姓名 || '之子' || Trim(To_Char(A.序号, '9'))) As 婴儿姓名, 婴儿性别, 出生时间" & vbNewLine & _
'                        "From 病人新生儿记录 A, 病人信息 B" & vbNewLine & _
'                        "Where A.病人id = [1] And A.主页id = [2] And A.病人id = B.病人id And A.序号 = [3]"
'
'                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取婴儿信息", mcurAdviceInf.lngPatId, mcurAdviceInf.lngPageID, mcurAdviceInf.lngBaby)
'
'                If Not rsTemp.EOF Then
'                    labPatientInfoName.Caption = "姓名:" & NVL(rsTemp!婴儿姓名) & "  性别:" & NVL(rsTemp!婴儿性别) & _
'                                        "  年龄:" & NVL(rsTemp!出生时间)
'                End If
'            End If

    Else
        labPatientInfoNo.Caption = "[" & GetStudyNumberDisplayName & ":--- ]"
    End If
    
    Call picListRowInfo_Resize
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Resize()
    If mbytFontSize > 0 Then
        Call AdjustFace(mbytFontSize)
    Else
        Call AdjustFace(9)
    End If
End Sub

Private Sub LoadShemeCustomCfg(ByVal LngSchemeNo As Long)
'根据用户ID/方案ID加载 个性化配置,检查配置是否符合要求，不符合则对应配置自动设置为空
On Error GoTo errH
    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strTmp As String
    
    strSql = "select 条件配置,过滤配置,列表配置 from 影像查询特性  where 用户ID=[1] and 查询方案ID =[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "加载查询个性化配置", mlngUserId, LngSchemeNo)

    '初始化查询参数
    
    On Error Resume Next
    
    If rsTemp.RecordCount = 1 Then
        mstrSchemeCfg.strSearchCfg = Split(rsTemp!过滤配置, "%")(0)
        mstrSchemeCfg.strFilterCfg = Split(rsTemp!过滤配置, "%")(1)
        mstrSchemeCfg.strListCfg = rsTemp!列表配置
    End If
    
    '检查strSearchCfg
    If UBound(Split(mstrSchemeCfg.strSearchCfg, ",")) < 3 Then mstrSchemeCfg.strSearchCfg = "当天," & Date & "," & Date & ","
    
    On Error GoTo errH
    
    mDTStart = CDate(Split(mstrSchemeCfg.strSearchCfg, ",")(1))
    mDTEnd = CDate(Split(mstrSchemeCfg.strSearchCfg, ",")(2))
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[LoadShemeCustomCfg]" & vbCrLf & Err.Description
    
End Sub

Private Function RefreshQueryWindow(ByVal LngSchemeNo As Long) As Boolean
'根据方案改变快速过滤界面,加载快速过滤菜单，根据个性化参数加载菜单选中项
On Error GoTo errH
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim rsTemp As ADODB.Recordset
    
    Dim i As Integer, j As Integer, lngID As Long
    Dim intMenuCount As Integer '快速过滤菜单数
    Dim intItemCount As Integer '快速过滤菜单子项目数
    Dim lngCount As Long
    
    Dim strSql As String
    Dim strTemp As String
    Dim strItems As String '保存数据库查询出来的快速查询项目
    Dim strName As String, strValue As String, strItemValue As String, strValueTmp As String
    
    Dim blNeedCreat As Boolean '基本过滤工具栏需要创建
    Dim blDynamicFilter As Boolean '是否动态过滤
    
    RefreshQueryWindow = False
    blNeedCreat = True
    
    '''''''清除已有快速过滤菜单
    Call LockWindowUpdate(Me.hwnd)
    For lngCount = cbrFilter.Count To 2 Step -1
        cbrFilter(lngCount).Delete
    Next
    
    '判断是已经创建，若有特定菜单说明已经创建过 blNeedCreat 设置为T
    Set objControl = cbrBaseFilter.FindControl(xtpControlLabel, conMenu_PacsQuery_TimeLab)
    If Not objControl Is Nothing Then
        blNeedCreat = False
    Else
        blNeedCreat = True
    End If
    
    Call LockWindowUpdate(0)
    
    '''''''创建基本过滤菜单
    Call InitCbrBaseFilter(LngSchemeNo, blNeedCreat)
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrFilter.VisualTheme = xtpThemeOfficeXP
    With cbrFilter.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    cbrFilter.AddImageList img16 '以VB.ImageList的Tag与ID进行关联
    cbrFilter.EnableCustomization False
    cbrFilter.ActiveMenuBar.Visible = False
    
    Set objBar = cbrFilter.Add("快速过滤", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False

    ReDim mTQuickFilterState.TCmdState(mSqlScheme.FilterCfgCount)
    mTQuickFilterState.intQuickFilterMenuCount = 0


    '创建快速过滤菜单 分为三种类型，
    '1:  普通快速过滤，可选项固定：  如配置：已登记;已报到;已报告
    '2:  普通快速过滤，可选项通过查询得到： 如配置： select distinct 编码 as 影像类别 from 影像检查类别
    '3:  自定义快速过滤，可选项通过前面的快速过滤选择情况得到，：如配置：[影像类别]#(这里是快速过滤可选项的Select查询语句)

    For i = 1 To mSqlScheme.FilterCfgCount
        strTemp = mSqlScheme.FilterCfg(i).DataFrom
        
        If InStr(UCase(strTemp), "SELECT") > 0 And Len(mSqlScheme.FilterCfg(i).CustomScript) = 0 Then
        '过滤条件为SQL语句，目前的条件是有"SELECT"
            strSql = mSqlScheme.FilterCfg(i).DataFrom
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "快速过滤获取项目")
    
            strItems = ""
            strTemp = mSqlScheme.FilterCfg(i).Name
            If rsTemp.RecordCount > 0 Then
                While rsTemp.EOF = False
                    If strItems = "" Then
                        strItems = strItems & rsTemp(strTemp)
                    Else
                        strItems = strItems & ";" & rsTemp(strTemp)
                    End If
                    rsTemp.MoveNext
                Wend
            End If
            
            Call cbrListAdd(objBar.Controls, mSqlScheme.FilterCfg(i).Name, strItems, i, , , , mSqlScheme.FilterCfg(i).SelectWay = swSingle)
        
        ElseIf InStr(UCase(strTemp), "SELECT") = 0 And Len(mSqlScheme.FilterCfg(i).CustomScript) = 0 Then
        '过滤条件为提前配置，用";"隔开这种
            Call cbrListAdd(objBar.Controls, mSqlScheme.FilterCfg(i).Name, mSqlScheme.FilterCfg(i).DataFrom, i, , , , mSqlScheme.FilterCfg(i).SelectWay = swSingle)
        ElseIf Len(mSqlScheme.FilterCfg(i).CustomScript) > 0 Then
        '过滤条件为前面快速过滤条件的结果
            strTemp = Split(mSqlScheme.FilterCfg(i).DataFrom, "#")(0)
            strTemp = Replace(strTemp, "[", "")
            strTemp = Replace(strTemp, "]", "")
            strTemp = Trim(strTemp)
            mTQuickFilterState.TCmdState(i).strMenuSQL = Split(mSqlScheme.FilterCfg(i).DataFrom, "#")(1)
            
            Call cbrListAdd(objBar.Controls, mSqlScheme.FilterCfg(i).Name, "", i, True, strTemp, mSqlScheme.FilterCfg(i).CustomScript, mSqlScheme.FilterCfg(i).SelectWay = swSingle)
        End If
    Next
    
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next

    cbrFilter.RecalcLayout
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''以下功能：根据参数恢复以前快速过滤菜单的选中情况
    '根据参数设置快速过滤选中情况
    'strValue 所有的快速过滤信息
    'strValueTmp 一组菜单的信息
    'strItemValue 一组菜单的信息

    If mintShowType = 1 Then Exit Function
    strValue = mstrSchemeCfg.strFilterCfg
    If Len(strValue) = 0 Then Exit Function
    
    intMenuCount = mTQuickFilterState.intQuickFilterMenuCount
    
    '处理每个快速过滤菜单
    For i = 1 To intMenuCount
    
        intItemCount = mTQuickFilterState.TCmdState(i).intItemCount
        lngID = 100 * i
        Set objControl = cbrFilter.FindControl(, lngID)
        strName = objControl.Parameter
        
        If UBound(Split(strValue, "|")) <> intMenuCount Then  '保存的菜单数量与数据库保存的数量不同
            On Error Resume Next
        Else
            On Error GoTo errH
        End If
        
        '找到对应的快速过滤菜单,根据菜单名称获取配置,需要判断是否动态过滤
        For j = 0 To UBound(Split(strValue, "|")) - 1
            strValueTmp = Split(strValue, "|")(j) '一项菜单的信息
            
            blDynamicFilter = False
            
            If Split(strValueTmp, ",")(0) = strName Then
                blDynamicFilter = (Split(strValueTmp, ",")(1) = "1")
                strValueTmp = Split(strValueTmp, ",")(2)
                
                '若是动态过滤 获取选中菜单名
                If blDynamicFilter Then mTQuickFilterState.TCmdState(i).strRelationChooseMenu = strValueTmp
                Exit For
            End If
        Next
        
        If UBound(Split(strValueTmp, ";")) + 1 <> intItemCount Then '保存的菜单数量与数据库保存的数量不同
            On Error Resume Next
        Else
            On Error GoTo errH
        End If
        
        '处理快速过滤菜单子项选中情况
        If Not blDynamicFilter Then
            '非动态过滤根据 0,1判断
            For j = 1 To intItemCount
                mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose = IIf(Val(Split(strValueTmp, ";")(j - 1)) = 1, True, False)
            Next
        End If
        
    Next
    
    Call RefreshCbrQuickFilterALL
    
    RefreshQueryWindow = True
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[RefreshQueryWindow]" & vbCrLf & Err.Description
End Function

Private Function DoStateImage(ByVal lngRow As Long) As Boolean
'处理状态图
On Error GoTo errH
    Dim i As Integer, j As Integer, k1 As Integer, k2 As Integer
    Dim objClsRelation As New clsScRowRelation
    Dim intImgCount As Integer
    Dim lngLeft As Long
    
    '首先清空状态图
    For i = imgState.Count - 1 To 0 Step -1
        imgState(i).Visible = False
    Next
    intImgCount = 0
    
    If mSqlScheme.ShowCfgCount < 1 Then Exit Function
    
    With vsfList
        
        For i = 1 To mSqlScheme.ShowCfgCount 'i 遍历列显示配置
            If mSqlScheme.ShowCfg(i).RowRelationCount > 0 Then
                
                For j = 1 To mSqlScheme.ShowCfg(i).RowRelationCount 'j遍历行关联
                    If Len(mSqlScheme.ShowCfg(i).RowRelation(j).Icon) > 0 Then '首先判断是否配置了显示图标
                        If .Cell(flexcpText, lngRow, .ColIndex(mSqlScheme.ShowCfg(i).Name)) = mSqlScheme.ShowCfg(i).RowRelation(j).TiggerData Then '判断是否符合触发数据
                            '添加状态图
                            If intImgCount = 0 Then
                                Set imgState(0).Picture = GetIcon(mSqlScheme.ShowCfg(i).RowRelation(j).Icon)
                                Call imgState(0).Move(labPatientInfoNo.Left + labPatientInfoNo.Width, 0)
                                imgState(0).Visible = True
                                
                                intImgCount = 1
                            Else
                                If imgState.Count <= intImgCount Then Load imgState(intImgCount)

                                Set imgState(intImgCount).Picture = GetIcon(mSqlScheme.ShowCfg(i).RowRelation(j).Icon)
                            
                                '重新设置位置
                                lngLeft = picListRowInfo.Width
                                lngLeft = intImgCount * imgState(0).Width
                                Call imgState(intImgCount).Move(lngLeft, 0)
                                imgState(intImgCount).Visible = True
                                
                                intImgCount = intImgCount + 1
                            End If
                            
                        End If
                    End If
                    
                Next  ' for j
            End If
        Next 'for i
    End With
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[DoStateImage]" & vbCrLf & Err.Description
End Function

Private Sub InitCbrBaseFilter(ByVal LngSchemeNo As Long, ByVal blNeedCreat As Boolean)
'初始化基本过滤控件，blNeedCreat是否需要创建，首次加载需要、切换方案只需要更新控件
On Error GoTo errH
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    Dim objPopbar As CommandBarPopup, objCusControl As CommandBarControlCustom
    Dim objCboControl As CommandBarComboBox
    Dim i As Integer
    Dim strSql As String
    
    Call LoadPatiIdentifyInfo
    
    If blNeedCreat Then
        '新建工具栏的处理
        CommandBarsGlobalSettings.App = App
        CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
        CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
        cbrBaseFilter.VisualTheme = xtpThemeOfficeXP
        With cbrBaseFilter.Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True '放在VisualTheme后有效
            .UseDisabledIcons = True
            .LargeIcons = False
            .SetIconSize False, 16, 16
            .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
        End With
        cbrBaseFilter.AddImageList img16 '以VB.ImageList的Tag与ID进行关联
        cbrBaseFilter.EnableCustomization False
        cbrBaseFilter.ActiveMenuBar.Visible = False
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set objBar = cbrBaseFilter.Add("基本过滤", xtpBarTop)
        objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        objBar.ContextMenuPresent = False
    
    ''''''''''''''''''''''''''''''''''''''''''时间菜单的处理
        If mTLayout.blShowTimeSelect Then
            '查找方式
            Set objControl = objBar.Controls.Add(xtpControlLabel, conMenu_PacsQuery_TimeLab, "时间范围：")
            
            Set objCboControl = objBar.Controls.Add(xtpControlComboBox, conMenu_PacsQuery_TimeCbo, "时间选择")
            Call objCboControl.AddItem("当天")
            Call objCboControl.AddItem("三天")
            Call objCboControl.AddItem("一周")
            Call objCboControl.AddItem("半个月")
            Call objCboControl.AddItem("一个月")
            Call objCboControl.AddItem("三个月")
            Call objCboControl.AddItem("半年")
            If mSqlScheme.dateRange = 0 Then Call objCboControl.AddItem("无限制")
            Call objCboControl.AddItem("自定义")
            Call SeekIndexSimple(objCboControl, Split(mstrSchemeCfg.strSearchCfg, ",")(0), False)
            
        End If
        
        Set objControl = objBar.Controls.Add(xtpControlButton, conMenu_PacsQuery_FindWay, "查询")
        objControl.Style = xtpButtonIcon
        objControl.IconId = IIf(mTPatiIdentifyInfo.blFind, C_ICON_FIND, C_ICON_LOCATE)
        
        
    
        Set objCusControl = objBar.Controls.Add(xtpControlCustom, conMenu_PacsQuery_PatiControl, "查找值")
        objCusControl.Handle = patiSearch.hwnd
        objControl.Visible = mTPatiIdentifyInfo.blShowPatiIdentify
        
        If mTPatiIdentifyInfo.blShowPatiIdentify Then
        '加载pati控件
            Call DoPatiIdentify
        End If
        
        Set objControl = objBar.Controls.Add(xtpControlButton, conMenu_PacsQuery_Do, "执行")
        objControl.Style = xtpButtonCaption
    Else
    '更新工具栏的处里
        
        Set objControl = cbrBaseFilter.FindControl(xtpControlLabel, conMenu_PacsQuery_TimeLab)
        objControl.Visible = mTLayout.blShowTimeSelect
        
        Set objCboControl = cbrBaseFilter.FindControl(xtpControlComboBox, conMenu_PacsQuery_TimeCbo)
        objCboControl.Clear
        Call objCboControl.AddItem("当天")
        Call objCboControl.AddItem("三天")
        Call objCboControl.AddItem("一周")
        Call objCboControl.AddItem("半个月")
        Call objCboControl.AddItem("一个月")
        Call objCboControl.AddItem("三个月")
        Call objCboControl.AddItem("半年")
        If mSqlScheme.dateRange = 0 Then Call objCboControl.AddItem("无限制")
        Call objCboControl.AddItem("自定义")
        Call SeekIndexSimple(objCboControl, Split(mstrSchemeCfg.strSearchCfg, ",")(0), False)
        objCboControl.Visible = mTLayout.blShowTimeSelect
        
        Set objControl = cbrBaseFilter.FindControl(xtpControlButton, conMenu_PacsQuery_FindWay)
        objControl.Visible = mTPatiIdentifyInfo.blShowPatiIdentify
            
        '注意使用这种类型 xtpControlButton
        Set objControl = cbrBaseFilter.FindControl(xtpControlButton, conMenu_PacsQuery_PatiControl)
        objControl.Visible = mTPatiIdentifyInfo.blShowPatiIdentify
        
        If mTPatiIdentifyInfo.blShowPatiIdentify Then
        '加载pati控件
            Call DoPatiIdentify
        End If
        
        Set objControl = cbrBaseFilter.FindControl(xtpControlButton, conMenu_PacsQuery_Do)
        objControl.Visible = mTPatiIdentifyInfo.blShowPatiIdentify
    End If
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[InitCbrBaseFilter]" & vbCrLf & Err.Description
End Sub

Private Sub SaveShemeCustomCfg(ByVal LngSchemeNo As Long)
'保存个性化配置
On Error GoTo errH
    Dim strSql As String
    Dim objCboControl As CommandBarComboBox
    
    strSql = "Zl_影像查询_个性化配置(" & mlngUserId & "," & LngSchemeNo & ",'" & mstrSchemeCfg.strSearchCfg & "%" & mstrSchemeCfg.strFilterCfg & "','" & mstrSchemeCfg.strListCfg & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "保存个性化配置")
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[SaveShemeCustomCfg]" & vbCrLf & Err.Description
End Sub

Private Sub cbrFilter_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
'快速过滤 执行
On Error GoTo errHandle
    Dim i As Integer
    Dim strTemp As String
    Dim lngAdviceID As Long
    Dim intIndex As Integer '菜单序号
    Dim intItemIndex As Integer '项目序号
    Dim intTMP As Integer
    
    intTMP = control.Id
    If (intTMP Mod 100) <> 0 Then
        intIndex = Int(intTMP / 100) + 1
        intItemIndex = (intTMP Mod 100)
        
        '根据选择情况更新缓存中的值
        If mTQuickFilterState.TCmdState(intIndex).blSingleChoose Then
            For i = 1 To mTQuickFilterState.TCmdState(intIndex).intItemCount
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose = False
            Next
            mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose = True
        Else
            mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose = Not mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose
        End If
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''自定义快速过滤处理开始
'        '刷新动态菜单设置
        If mTQuickFilterState.TCmdState(intIndex).intRelation = 1 Then
        '若是关联过滤前者，需要更新后者显示项
            Call RefreshCbrQuickFilter(intIndex, False)
            
        ElseIf mTQuickFilterState.TCmdState(intIndex).intRelation = 2 Then
        '若是关联过滤后者，需要更新选择项
            control.Checked = mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose
            
            If control.Checked Then
                If InStr(";" & mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu & ";", ";" & control.Parameter & ";") = 0 Then
                    mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu = mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu & control.Parameter & ";"
                End If
            Else
                If InStr(";" & mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu & ";", ";" & control.Parameter & ";") > 0 Then
                    mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu = Replace(mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu, control.Parameter & ";", "")
                End If
            End If
            
            Call GetQuickFilterSQLPar(intIndex)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''自定义快速过滤处理结束
        
        Call SaveFilterCfg

        ''''''''''''''点击快速过滤后立刻执行过滤并且呈现到列表中
        If Not mrsData Is Nothing Then
            Set mrsDataShow = GetFilterFromQuickFilter
            
            lngAdviceID = GetSelectRowAdviceID
            
            mblSearching = True
            Set vsfList.DataSource = mrsDataShow
            mblSearching = False
            
            '列统计
            Call ColStatistics(mrsDataShow)
            
            '列顺序重排
            Call LoadListHeadCfg(mlngAdviceID)
            
            '行关联
            If vsfList.TopRow <> vsfList.BottomRow Then
                For i = vsfList.TopRow To vsfList.BottomRow
                    Call RefreshRowRelation(i)
                Next
            End If
            
            '重新排序
            Call ResetSort(mlngSortCol, mintSortOrder)
        Else
            Call ColStatistics(mrsData)
        End If
        
    End If
    
    Exit Sub
errHandle:
    MsgBox "[cbrFilter_Execute]" & vbCrLf & Err.Description, vbOKOnly, "异常"
End Sub

Private Sub cbrFilter_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
'On Error GoTo errHandle ID9400 下标越界 原因未知
On Error Resume Next
    Dim i As Integer
    Dim strTemp As String
    Dim intTMP As Integer
    Dim intIndex As Integer
    Dim intItemIndex As Integer
    
    Static blRun As Boolean
    
    If blRun Then Exit Sub
    blRun = True
    strTemp = ""
    intTMP = control.Id
    
    intItemIndex = (intTMP Mod 100)

    If (intTMP Mod 100) = 0 Then
        intIndex = Int(intTMP / 100)
        '根据子项目选择情况决定菜单名称和使用的图标
        For i = 1 To mTQuickFilterState.TCmdState(intIndex).intItemCount
            
            If mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose Then
                If Len(strTemp) = 0 Then
                    strTemp = mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName
                Else
                    strTemp = strTemp & "," & mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName
                End If
            End If
            
        Next

        If Len(strTemp) = 0 Then
            control.ToolTipText = "根据[" & control.Parameter & "]进行过滤"
            control.Caption = control.Parameter
            control.IconId = C_ICON_MENUNOCHOOSE
        Else
            control.ToolTipText = "显示[" & control.Parameter & "]为[" & strTemp & "]的检查"
            control.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
            control.IconId = C_ICON_MENUCHOOSE
        End If
    
        If mTQuickFilterState.TCmdState(intIndex).intItemCount = 0 Then
            control.ToolTipText = "根据[" & control.Parameter & "]进行过滤"
            control.Caption = control.Parameter
            control.IconId = C_ICON_MENUCHOOSE
            control.Enabled = True
        Else
            control.Enabled = True
        End If
        
    Else
        '改变图标

        intIndex = Int(intTMP / 100) + 1
        
        If mTQuickFilterState.TCmdState(intIndex).intRelation <> 2 Then
            control.IconId = IIf(mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose, C_ICON_MENUCHOOSE, C_ICON_MENUNOCHOOSE)
        Else
            If mTQuickFilterState.TCmdState(intIndex).intItemCount < 1 Then
                control.Enabled = False
            End If
            
            control.IconId = IIf(mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose, C_ICON_MENUCHOOSE, C_ICON_MENUNOCHOOSE)
        End If
    End If
    
    blRun = False
    
    Exit Sub
'errHandle:
'    Err.Raise -1, "frmPacsQuery", "[cbrFilter_Update]" & vbCrLf & Err.Description
End Sub

Private Sub RefreshCbrQuickFilterALL()
'刷新所有动态快速过滤选中情况，也就是初始化
On Error GoTo errH
    Dim i As Integer
    Dim intNeedDoCount As Integer
    Dim intNeedDoIndex() As Integer
    
    intNeedDoCount = 0
    ReDim intNeedDoIndex(0)
    
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        If mTQuickFilterState.TCmdState(i).intRelation = 2 Then
            
            intNeedDoCount = intNeedDoCount + 1
            ReDim Preserve intNeedDoIndex(intNeedDoCount)
            intNeedDoIndex(intNeedDoCount) = i
        End If
    Next
    
    For i = intNeedDoCount To 1 Step -1
        Call RefreshCbrQuickFilter(intNeedDoIndex(i), True)
    Next
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[RefreshCbrQuickFilterALL]" & vbCrLf & Err.Description
End Sub

Private Sub RefreshCbrQuickFilter(ByVal lngIndex As Long, ByVal blInit As Boolean)
'更新自定义及快速过滤菜单，参数： 菜单ID，触发菜单选中项"123,456,789"这种形式
'根据触发菜单信息查找自定义菜单信息，删除原来的自定义菜单子项，然后重新生成。
'blInit 是否初始化 是：lngIndex表示需要改变的菜单   否：表示关联菜单前项
On Error GoTo errH
    Dim lngIndexRelationMenu ' 自定义菜单索引
    Dim lngRelationID As Long '自定义菜单ID
    Dim strRelationName As String '自定义菜单名称
    Dim i As Long, j As Long
    Dim ObjPopMenu As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    Dim rsTemp As Recordset
    Dim blHaveSameMenu As Boolean
    Dim strMenuName As String
    Dim intMenuCount As Integer
    Dim strSql As String, strTmp As String
    
    '''获取需要改变的菜单在缓存中的ID开始
    If Not blInit Then
        strRelationName = mTQuickFilterState.TCmdState(lngIndex).strRelationName
        
        '寻找关联菜单ID 和 名称
        For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
            If mTQuickFilterState.TCmdState(i).strName = strRelationName Then
                lngRelationID = mTQuickFilterState.TCmdState(i).lngID
                lngIndexRelationMenu = i
                Exit For
            End If
        Next
    Else
        lngRelationID = mTQuickFilterState.TCmdState(lngIndex).lngID
        lngIndexRelationMenu = lngIndex
    End If
    '''获取需要改变的菜单在缓存中的ID结束

    '''清除已有菜单子项开始
    Call LockWindowUpdate(Me.hwnd)
    
    Set ObjPopMenu = cbrFilter.FindControl(, lngRelationID)
    
    If Not ObjPopMenu Is Nothing Then
        For i = 1 To ObjPopMenu.CommandBar.Controls.Count
            ObjPopMenu.CommandBar.Controls(1).Delete
        Next
    End If
    '''清除已有菜单子项结束
    
    Call LockWindowUpdate(0)
    
    '创建菜单
    strSql = mTQuickFilterState.TCmdState(lngIndexRelationMenu).strMenuSQL
            
    If mobjSqlParse Is Nothing Then Set mobjSqlParse = New clsSqlParse
    Call mobjSqlParse.init(strSql)
    strSql = mobjSqlParse.GetQuerySql
    
    Set rsTemp = ExecuteCore(strSql, "快速过滤获取项目", mobjSqlParse.ParValues)
    
    ''''判断是否子菜单显示数据是过滤数据,跟 rsTemp.Fields.Count 的值有关
    ''是，距离 影像类别-部位分组     对于部位分组来说，显示的是分组，参与过滤的是部位。rsTemp.Fields.Count 应该是 2
    ''否，rsTemp.Fields.Count 应该是1 菜单显示的内容就是参与过滤的内容
    If rsTemp.RecordCount = 0 Then
        ReDim mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(0)
        mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount = 0
        ObjPopMenu.Enabled = False
        Exit Sub
    End If
    
    '获取选择菜单配置（用于恢复选择情况）
    strTmp = mTQuickFilterState.TCmdState(lngIndexRelationMenu).strRelationChooseMenu
    
    If rsTemp.Fields.Count = 1 Then
    '''1 表示过滤内容就是显示内容
        mTQuickFilterState.TCmdState(lngIndexRelationMenu).blSimpleFilter = True
        ReDim mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(rsTemp.RecordCount)
        
        For i = 1 To rsTemp.RecordCount
            strMenuName = rsTemp.Fields(0).Value
            Set cbrControl = ObjPopMenu.CommandBar.Controls.Add(xtpControlButton, lngRelationID - 100 + i, strMenuName)
            
            cbrControl.Parameter = strMenuName
            
            mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(i).blChoose = IIf(InStr(";" & strTmp & ";", ";" & strMenuName & ";") > 0, True, False)
            mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(i).intItemIndex = i
            mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(i).strName = strMenuName
            cbrControl.CloseSubMenuOnClick = False
            If i <> rsTemp.RecordCount Then rsTemp.MoveNext
        Next
        mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount = rsTemp.RecordCount
    Else
    '''rsTemp.Fields.Count <> 1(=2) 表示第一个字段是显示内容 第二个字段是实际过滤内容
        mTQuickFilterState.TCmdState(lngIndexRelationMenu).blSimpleFilter = False
        rsTemp.MoveFirst

        intMenuCount = 1
        
        ReDim mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(0)
        mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount = 0
               
        For i = 1 To rsTemp.RecordCount
            strMenuName = rsTemp.Fields(0).Value
            blHaveSameMenu = False
            
            '首先判断是否有重复的分组，若有，把后者与前者不同的部位加到前面的菜单中
            If mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount > 0 Then
                For j = 1 To mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount
                    If rsTemp.Fields(0).Value = mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(j).strName Then
    '                        '有重复分组
                        Set cbrControl = cbrFilter.FindControl(, 100 * (lngIndex - 1) + j, , True)
    
                        '这个地方做个简单的处理，只有之前没有的部位才需要额外增加
                        cbrControl.Category = cbrControl.Category & CbrFilterDeal(cbrControl.Category, rsTemp.Fields(1).Value)
                        cbrControl.Category = Replace(cbrControl.Category, ",,", ",")
                        mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(j).strFilterValue = cbrControl.Category
    
                        '增加后重复的项目
                        blHaveSameMenu = True
                    End If
                Next
            End If

            If Not blHaveSameMenu Then
            '没有重复，直接增加
                ReDim Preserve mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount + 1)

                Set cbrControl = ObjPopMenu.CommandBar.Controls.Add(xtpControlButton, lngRelationID - 100 + intMenuCount, strMenuName)

                cbrControl.Parameter = strMenuName
                cbrControl.Category = rsTemp.Fields(1).Value
                
                mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(intMenuCount).blChoose = IIf(InStr(";" & strTmp & ";", ";" & strMenuName & ";") > 0, True, False)
                
                mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(intMenuCount).intItemIndex = i
                mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(intMenuCount).strName = strMenuName
                mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(intMenuCount).strFilterValue = cbrControl.Category
                mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount = intMenuCount

                intMenuCount = intMenuCount + 1

                cbrControl.CloseSubMenuOnClick = False

            End If

             If i <> rsTemp.RecordCount Then rsTemp.MoveNext
        Next
               
        Call GetQuickFilterSQLPar(lngIndex)
    End If
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[RefreshCbrQuickFilter]" & vbCrLf & Err.Description
End Sub

Private Sub cbrListAdd(ByVal Mycontrol As CommandBarControls, ByVal strName As String, ByVal strItems As String, ByVal intIndex As Integer, _
Optional blDynamic As Boolean = False, Optional strRelationName As String = "", Optional strCustomScript As String = "", Optional blSingleChoose As Boolean = False)
'每个过滤菜单，占用100个ID 例如1号 1~99  2号100~199
'向cbrList中增加快速过滤菜单
'strName：菜单名 如:检查类型
'strItems：菜单项目名，如： 已登记;已报到;已检查 用";"分开，非动态快速过滤 由此获取菜单数量
'intIndex:第几个过滤项
'blDynamic是否动态快速过滤  若是，需要在关联菜单中进行一些设置
' strRelationName 关联项目名称
'rsData: 菜单记录集

On Error GoTo errH
    Dim objControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim i As Integer
    Dim intCount As Integer
    Dim rsTemp As Recordset
    
    '快速过滤菜单数 1
    mTQuickFilterState.intQuickFilterMenuCount = mTQuickFilterState.intQuickFilterMenuCount + 1
    
    intCount = UBound(Split(strItems, ";")) + 1
    
    Set objControl = Mycontrol.Add(xtpControlButtonPopup, 100 * intIndex, strName)
    objControl.ToolTipText = "根据" & strName & "过滤"
    
    objControl.Parameter = strName
    mTQuickFilterState.TCmdState(intIndex).intMenuIndex = mTQuickFilterState.intQuickFilterMenuCount
    mTQuickFilterState.TCmdState(intIndex).intItemCount = intCount
    mTQuickFilterState.TCmdState(intIndex).strName = strName
    mTQuickFilterState.TCmdState(intIndex).lngID = 100 * intIndex
    mTQuickFilterState.TCmdState(intIndex).blSingleChoose = blSingleChoose
    
    '100 * (intIndex - 1) + i分配给这个快速过滤项的ID
    If blDynamic Then
        '关联过滤设置类型2
        mTQuickFilterState.TCmdState(intIndex).intRelation = 2
        mTQuickFilterState.TCmdState(intIndex).strCustomScript = strCustomScript
        mTQuickFilterState.TCmdState(intIndex).strRelationName = strRelationName
        '根据名称寻找被关联过滤
        For i = 1 To intIndex - 1
            If mTQuickFilterState.TCmdState(i).strName = strRelationName Then
                
                '被关联过滤设置类型1
                mTQuickFilterState.TCmdState(i).intRelation = 1
                '被关联过滤设置关联过滤名称
                mTQuickFilterState.TCmdState(i).strRelationName = strName
                '找到后退出
                Exit For
            End If
        Next
        
    Else
        ReDim mTQuickFilterState.TCmdState(intIndex).cmdItem(intCount)
        '非自定义快速过滤
        For i = 1 To intCount
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, 100 * (intIndex - 1) + i, Split(strItems, ";")(i - 1))
            cbrPopControl.Parameter = Split(strItems, ";")(i - 1)
            
            mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose = False
            mTQuickFilterState.TCmdState(intIndex).cmdItem(i).intItemIndex = i
            mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName = Split(strItems, ";")(i - 1)
            cbrPopControl.CloseSubMenuOnClick = False
        Next
    
    End If
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[cbrListAdd]" & vbCrLf & Err.Description
End Sub

Private Sub SaveFilterCfg()
'保存快速过滤参数
On Error GoTo errH
    Dim objControl As CommandBarControl
    Dim strName As String
    Dim strValue As String
    Dim lngID As Long
    Dim i As Integer
    Dim j As Integer
    Dim intMenuCount As Integer '菜单数
    Dim intItemCount As Integer '菜单子菜单数
    Dim strValueAll As String
    
    strValueAll = ""
    
    intMenuCount = mTQuickFilterState.intQuickFilterMenuCount
    For i = 1 To intMenuCount
        intItemCount = mTQuickFilterState.TCmdState(i).intItemCount
        
        lngID = 100 * i
        Set objControl = cbrFilter.FindControl(, lngID)
        strName = objControl.Parameter
        strValue = ""
        
        If mTQuickFilterState.TCmdState(i).intRelation <> 2 Then
            '非动态快速过滤的保存，按顺序用0/1表示是否选中
            If intItemCount > 0 Then
                For j = 1 To intItemCount
                    
                    If Len(strValue) = 0 Then
                        strValue = strValue & IIf(mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose, "1", "0")
                    Else
                        strValue = strValue & ";" & IIf(mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose, "1", "0")
                    End If
                Next
            End If
            '注意此处的"0"用来表示非动态快速过滤
            strValue = strName & ",0," & strValue
            strValueAll = strValueAll & strValue & "|"
        Else
            '动态快速过滤的保存，保存的菜单名称
            If intItemCount > 0 Then
                For j = 1 To intItemCount
                    If mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose Then strValue = strValue & mTQuickFilterState.TCmdState(i).cmdItem(j).strName & ";"
                Next
            End If
            '注意此处的"1"用来表示动态快速过滤
            strValue = strName & ",1," & strValue
            strValueAll = strValueAll & strValue & "|"
        End If

    Next
    
    mstrSchemeCfg.strFilterCfg = strValueAll
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[SaveFilterCfg]" & vbCrLf & Err.Description
End Sub
Private Function GetFilterFromQuickFilter() As Recordset
'获取快速过滤条件、并且进行过滤,返回过滤后的记录集
'对mrsData 操作，运行后不会改变mrsData,返回经过过滤后的全新记录集
'获得快速过滤条件
On Error GoTo errH
    Dim objControl As CommandBarControl
    Dim strFilter As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim blChooseOne As Boolean '某个菜单是否有过滤项被选中，若没有相当于不过滤
    Dim strFilterTmp As String
    Dim blIsStr As Boolean '快速过滤类型是否是字符，决定是否增加 “' 这个符号
    Dim strFilterField As String
    Dim intCustomeFilter As Integer   '是否有自定义快速过滤
    
    intCustomeFilter = 0
    strFilter = ""
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        blChooseOne = False
        strFilterTmp = ""
        Set objControl = cbrFilter.FindControl(, i * 100)
        blIsStr = True
        
        For j = 0 To mrsData.Fields.Count - 1
            
            If objControl.Parameter = mrsData.Fields(j).Name Then
                If mrsData.Fields(j).Type = adVarNumeric Or mrsData.Fields(j).Type = adNumeric Then
                    blIsStr = False
                    Exit For
                End If
            End If
        Next
        
        '首先进行固定快速过滤的处理
        If mTQuickFilterState.TCmdState(i).intRelation <> 2 Then
            For j = 1 To mTQuickFilterState.TCmdState(i).intItemCount
                strFilterField = mTQuickFilterState.TCmdState(i).cmdItem(j).strName
            
                If InStr(strFilterField, "-") > 0 Then strFilterField = Mid(strFilterField, 1, InStr(strFilterField, "-") - 1)
    
                If mTQuickFilterState.TCmdState(i).intRelation <> 2 Then
                    If Not mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose Then
                    '未被选中
                        If Not blIsStr Then
                            If Len(strFilterTmp) = 0 Then
                                strFilterTmp = strFilterTmp & objControl.Parameter & " <> " & strFilterField & " "
                            Else
                                strFilterTmp = strFilterTmp & " and " & objControl.Parameter & " <> " & strFilterField & " "
                            End If
                        Else
                            If Len(strFilterTmp) = 0 Then
                                strFilterTmp = strFilterTmp & objControl.Parameter & " <> '" & strFilterField & "' "
                            Else
                                strFilterTmp = strFilterTmp & " and " & objControl.Parameter & " <> '" & strFilterField & "' "
                            End If
                        End If
                    Else
                        blChooseOne = True
                    End If
                    
                End If
            Next
        Else
            intCustomeFilter = intCustomeFilter + 1
        End If
        '只有被选中了过滤项，才加入到过滤条件中
        If blChooseOne And Len(strFilterTmp) > 0 Then
            If Len(strFilter) = 0 Then
                strFilter = strFilter & strFilterTmp
            Else
                strFilter = strFilter & " and " & strFilterTmp
            End If
        End If

    Next
    mrsData.Filter = strFilter

    '没有自定义快速过滤可以现在退出
    If intCustomeFilter = 0 Then
        Set GetFilterFromQuickFilter = CopyRecordSet(mrsData)
        Exit Function
    End If
    
    Dim strVBS As String
    Dim rstVBS As Recordset
    Dim rsTmp() As Recordset
    Dim objGlobal As clsGlobal

    Set objGlobal = New clsGlobal
    '遍历每个自定义快速过滤条件,根据菜单选中情况执行VBS脚本
    ReDim rsTmp(intCustomeFilter)
    j = 0
    
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        If mTQuickFilterState.TCmdState(i).intRelation = 2 Then
            j = j + 1
            If j = 1 Then
                Set rsTmp(0) = CopyRecordSet(mrsData)
            Else
                Set rsTmp(j - 1) = CopyRecordSet(rsTmp(j - 2))
            End If

            Call objGlobal.ExecuteScript(mTQuickFilterState.TCmdState(i).strCustomScript, rsTmp(j - 1), mTQuickFilterState.TCmdState(i).strRelationValueForVBSFilter)

        End If
    Next
    
    Set GetFilterFromQuickFilter = rsTmp(intCustomeFilter - 1)
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetFilterFromQuickFilter]" & vbCrLf & Err.Description
End Function

Private Function GetListHeadString() As String
'得到列名参数: 名称,宽度,是否显示  例如  "类别,1000,1|执行过程,2000,0|"
On Error GoTo errH
    Dim i As Integer
    Dim strValue As String
    Dim strTemp As String
    
    strTemp = ""
    
    For i = 1 To vsfList.Cols - 1
        If Len(strTemp) > 0 Then
            strTemp = strTemp & "|"
        End If
        
        strTemp = strTemp & vsfList.TextMatrix(0, i)
        strTemp = strTemp & "," & vsfList.ColWidth(i)
        strTemp = strTemp & "," & IIf(vsfList.ColHidden(i), "0", "1")
        
    Next

    GetListHeadString = strTemp
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetListHeadString]" & vbCrLf & Err.Description
End Function

Private Sub SaveListHeadCfg()
'保存加载列头参数
On Error GoTo errH
    Dim strValue As String
    
    mstrSchemeCfg.strListCfg = GetListHeadString()
    Call DoBeforeFilterData(True)
    
    Exit Sub
errH:
    MsgBox "SaveListHeadCfg" & Err.Description
End Sub

Private Function LoadListHeadCfg(ByVal lngAdviceIDOld As Long) As Boolean
'LoadListHeadCfg：之前的选中检查医嘱ID，若为0，直接选中第一行,若为 -1 不需要改变选中
'根据个性化参数刷新列表配置(排序、宽度、是否显示)
'注意i和j的初始值

'该函数需要处理之前选中状态的保留：若之前选择了检查，则任何操作后应该保持之前的检查，若之前的检查已经消失，则选中第一个。
On Error GoTo errH
    Dim strTmp As String
    Dim strValue As String
    Dim strColName As String
    Dim intWidth As Integer
    Dim lngAdviceIDNew As Long
    
    Dim blShow As Boolean
    Dim blMatch As Boolean '已经保存的列表是否与配置 匹配，若不匹配（方案已经使用，后来修改了配置、比如增加字段）
    Dim i As Integer
    Dim j As Integer
    
    
    blMatch = True
'    If mSqlScheme.ShowCfgCount < 1 Then Exit Function
    If mintShowType = 1 Then Exit Function
    
    strValue = mstrSchemeCfg.strListCfg
    If Len(strValue) = 0 Then Exit Function
    
    '判断保存的列表配置是否与当前列表匹配（列表字段，数量），因为调整方案后可能导致旧配置不可用
    If UBound(Split(strValue, "|")) <> vsfList.Cols - 2 Then blMatch = False
    For i = 1 To vsfList.Cols - 1
        If InStr(strValue, vsfList.TextMatrix(0, i)) = 0 Then blMatch = False
    Next
    
    If blMatch Then
    '"LoadListHeadCfg匹配"
        For i = 1 To vsfList.Cols - 1

            strTmp = Split(strValue, "|")(i - 1)
            strColName = Split(strTmp, ",")(0)
            intWidth = Val(Split(strTmp, ",")(1))
            blShow = Val(Split(strTmp, ",")(2))

            If vsfList.TextMatrix(0, i) <> strColName Then
                For j = 1 To vsfList.Cols - 1
                    If vsfList.TextMatrix(0, j) = strColName Then

                        vsfList.ColPosition(j) = i

                        Exit For
                    End If
                Next
            End If

            vsfList.ColWidth(i) = intWidth
            vsfList.ColHidden(i) = Not blShow

        Next
    Else
        '"LoadListHeadCfg不匹配"
        '列表根据配置初始化
        On Error Resume Next
        strValue = mstrSchemeCfg.strListCfgDefault
        
        For i = 1 To vsfList.Cols - 1
    
            strTmp = Split(strValue, "|")(i - 1)
            strColName = Split(strTmp, ",")(0)
            intWidth = Val(Split(strTmp, ",")(1))
            blShow = Val(Split(strTmp, ",")(2))
    
            If vsfList.TextMatrix(0, i) <> strColName Then
                For j = 1 To vsfList.Cols - 1
                    If vsfList.TextMatrix(0, j) = strColName Then
    
                        vsfList.ColPosition(j) = i
    
                        Exit For
                    End If
                Next
            End If
    
            vsfList.ColWidth(i) = intWidth
            vsfList.ColHidden(i) = Not blShow
    
        Next
        On Error GoTo errH
    End If

    Call DoBeforeFilterData(True)
     
    '根据最大行号修改第一列的宽度
    If vsfList.Rows < 11 Then
        vsfList.ColWidth(0) = TextWidth("XX")
    ElseIf 10 < vsfList.Rows And vsfList.Rows < 101 Then
        vsfList.ColWidth(0) = TextWidth("XXX")
    ElseIf 100 < vsfList.Rows And vsfList.Rows < 1001 Then
        vsfList.ColWidth(0) = TextWidth("XXXX")
    Else
        vsfList.ColWidth(0) = TextWidth("XXXXX")
    End If
    
    If lngAdviceIDOld = -1 Then Exit Function
    
    If vsfList.Rows > 1 Then
        
        '之前的检查列表选中行的医嘱ID
        If lngAdviceIDOld = 0 Then
            GoTo ClearListFace
        Else
            lngAdviceIDNew = vsfList.FindRow(lngAdviceIDOld, 1, vsfList.ColIndex(mstrListKeyCol), False, False)
            
            If lngAdviceIDNew > 0 Then
                vsfList.Row = lngAdviceIDNew
    
                If vsfList.TopRow > vsfList.Row Then vsfList.TopRow = vsfList.Row
                
                If vsfList.BottomRow - 1 < vsfList.Row Then
                    vsfList.TopRow = vsfList.TopRow + (vsfList.Row - vsfList.BottomRow) + 1
                End If
            Else
                GoTo ClearListFace
            End If
        
        End If
    Else
        GoTo ClearListFace
    End If
    Exit Function
    
ClearListFace:
'ClearListFace 检查列表不选中任何检查，右边Tab也需要设置为未选择检查的状态
    mlngAdviceID = 0
    For i = imgState.Count - 1 To 0 Step -1
        imgState(i).Visible = False
    Next

    GetNullStudyInfo
    cboHistory.Clear
    Call FillCurAdviceTxtInfor
    Call FillCurAdviceAppend(0, True)
    
    RaiseEvent OnListRowSelClear
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[LoadListHeadCfg]" & vbCrLf & Err.Description
End Function


Private Sub vsfList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    SaveListHeadCfg
End Sub

Private Sub vsfList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    SaveListHeadCfg
End Sub
Public Function ExecuteQuery(ByVal strExecuteType As String, Optional ByVal LngSetRow As Long = 0) As Boolean
'执行过滤、刷新、查找功能
On Error GoTo errH
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim objCboControl As CommandBarComboBox
    Dim strTmp As String
    Dim i As Integer, j As Integer
    Dim lngAdviceID As Long
    
    Dim bllimitTime As Boolean '是否限制时间  选择“无限制”为False  其他为  true
    
    bllimitTime = True
    
    Set objCboControl = cbrBaseFilter.FindControl(xtpControlComboBox, conMenu_PacsQuery_TimeCbo)
    If objCboControl.Text = "自定义" Then
        dtStart = mDTStart
        dtEnd = mDTEnd
    Else
        dtEnd = zlDatabase.Currentdate()
        Select Case objCboControl.Text
            Case "当天"
                dtStart = DateAdd("d", -1, dtEnd)
            Case "三天"
                dtStart = DateAdd("d", -3, dtEnd)
            Case "一周"
                dtStart = DateAdd("ww", -1, dtEnd)
            Case "半个月"
                dtStart = DateAdd("ww", -2, dtEnd)
            Case "一个月"
                dtStart = DateAdd("m", -1, dtEnd)
            Case "三个月"
                dtStart = DateAdd("m", -3, dtEnd)
            Case "半年"
                dtStart = DateAdd("m", -6, dtEnd)
            Case "无限制"
                bllimitTime = False
'                '无限制怎么取时间？ 很大的时间？100年？
'                dtStart = DateAdd("yyyy", -50, dtEnd)
        End Select
    End If

    If bllimitTime Then
        Call mObjQuery.SetFilterValue("系统.开始日期", dtStart)
        Call mObjQuery.SetFilterValue("系统.结束日期", dtEnd)
'    Else
'        Call mObjQuery.SetFilterValue("系统.开始日期", Null)
'        Call mObjQuery.SetFilterValue("系统.结束日期", Null)
    End If

    If strExecuteType = "过滤" Then
        mTqueryType = 过滤
        Set mrsData = mObjQuery.ExecuteWithFilter(dtStart, dtEnd, Me)
    ElseIf strExecuteType = "刷新" Then
        mTqueryType = 刷新
        Set mrsData = mObjQuery.Execute(dtStart, dtEnd, False)
    ElseIf strExecuteType = "查找" Then
        mTqueryType = 查找
        Set mrsData = mObjQuery.Execute(dtStart, dtEnd, False)
    End If
    
    Call DoBeforeFilterData(False)
    
    If Not mrsData Is Nothing Then
        '获取所有字段信息
        If Len(mstrSchemeCfg.strListCfgDefaultColOrder) = 0 Then
            For i = 0 To mrsData.Fields.Count - 1
                mstrSchemeCfg.strListCfgDefaultColOrder = mstrSchemeCfg.strListCfgDefaultColOrder & mrsData.Fields(i).Name & "|"
            Next
        End If
             
        If mrsData.RecordCount > 0 Then

            Set mrsDataShow = GetFilterFromQuickFilter
            Set mrsDataShow = mObjQuery.DataConvert(mrsDataShow, mlngSchemeNo)
        
            If Not mrsDataShow Is Nothing Then
                
                lngAdviceID = GetSelectRowAdviceID
                mblSearching = True
                Set vsfList.DataSource = mrsDataShow
                mblSearching = False
                
                If vsfList.TopRow <> vsfList.BottomRow Then
                    For i = vsfList.TopRow To vsfList.BottomRow
                        Call RefreshRowRelation(i)
                    Next
                End If
                
                '列统计
                Call ColStatistics(mrsDataShow)
                
                Call DoListCfg
                
                Call LoadListHeadCfg(mlngAdviceID)
            Else
                Call ColStatistics(mrsData)
                Call LoadListHeadCfg(0)
            End If
    
            Call ResetSort(mlngSortCol, mintSortOrder)
        Else
            lngAdviceID = GetSelectRowAdviceID
            mblSearching = True
            Set vsfList.DataSource = mrsData
            mblSearching = False
            '列统计
            Call ColStatistics(mrsData)
            Call LoadListHeadCfg(0)
        End If
    End If

    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[ExecuteQuery]" & vbCrLf & Err.Description
End Function

Public Function ExecuteWithLink(ByVal strSql As String) As Boolean
'收藏功能
On Error GoTo errH
    Set mrsData = mObjQuery.ExecuteWithLink(strSql)
    
    If Not mrsData Is Nothing Then
        If mrsData.RecordCount > 0 Then
            Set mrsDataShow = GetFilterFromQuickFilter
            Set mrsDataShow = mObjQuery.DataConvert(mrsDataShow, mlngSchemeNo)
            If Not mrsDataShow Is Nothing Then
                mblSearching = True
                Set vsfList.DataSource = mrsDataShow
                mblSearching = False
                '列统计
                Call ColStatistics(mrsDataShow)
                Call LoadListHeadCfg(0)
            Else
                Call LoadListHeadCfg(0)
            End If
        End If
    Else
        Call ColStatistics(mrsData)
        Call LoadListHeadCfg(0)
    End If
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[ExecuteWithLink]" & vbCrLf & Err.Description
End Function

Private Sub vsfList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errH
    Dim strCfgNew As String
    Dim blRaiseEvent As Boolean '比如用于pacsMain弹出右键菜单
    
    blRaiseEvent = True
    
    If mintShowType = 1 Then Exit Sub
    If vsfList.MouseRow = 0 And Button = 2 Then
        '获得初始配置串 和 个性化配置串
        Call frmVsfColsList.ShowVsfColsListWindow(mstrSchemeCfg.strListCfgDefault, mstrSchemeCfg.strListCfg, mfrmParent)
        strCfgNew = frmVsfColsList.GetListCfg
        
        If Len(strCfgNew) > 0 Then
            mstrSchemeCfg.strListCfg = strCfgNew
            Call LoadListHeadCfg(-1)
        End If
        
        '此处需要调用这个，保证行关联的正确
        Call DoBeforeFilterData(True)
        
        blRaiseEvent = False
    End If
    

    If blRaiseEvent Then RaiseEvent OnMouseUp(Button, Shift, X, Y)

    Exit Sub
errH:
    MsgBox "[vsfList_MouseUp]" & vbCrLf & Err.Description, vbOKOnly, "异常"
End Sub

Private Function InitCardType(ByVal strCardNames As String) As String
'按指定格式初始化卡类型
On Error GoTo errH
    Dim i As Integer
    Dim aryKindInfo() As String
    Dim strKinds As String
    
    aryKindInfo = Split(strCardNames, ";")
    
    strKinds = ""
    For i = 0 To UBound(aryKindInfo) - 1
        If strKinds <> "" Then strKinds = strKinds & ";"
        strKinds = strKinds & aryKindInfo(i) & "|" & aryKindInfo(i) & "|-1"
    Next i
    
    InitCardType = strKinds & ";"
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[InitCardType]" & vbCrLf & Err.Description
End Function

Private Function GetFilter() As String
'按指定格式初始化卡类型
'类似这种格式   ：    "姓名|时间|系统时间|日期|类别"
On Error GoTo errH
    Dim i As Integer
    GetFilter = ""
    If mSqlScheme Is Nothing Then Exit Function
    
    For i = 1 To mSqlScheme.SerachCfgCount
        
        If mSqlScheme.SerachCfg(i).InputType = itPopup Or mSqlScheme.SerachCfg(i).InputType = itBoth Then
            GetFilter = GetFilter & mSqlScheme.SerachCfg(i).Name & ";"
        End If
    Next

    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetFilter]" & vbCrLf & Err.Description
End Function

Private Sub SeekNextPati(ByVal blnFirst As Boolean, ByVal strName As String, _
    ByVal strFilter As String, Optional blnIsReSeek As Boolean = False)
On Error GoTo errH
'------------------------------------------------
'功能：在病人列表中定位指定的记录
'参数： blnFirst -- 是否第一次查找
'返回：无，直接在病人列表中定位
'------------------------------------------------
    Dim i As Long
    Dim intB As Integer
    Dim lngEndRow As Long
    Dim lngSelRow As Long
    Dim strTemp As String
    Dim lngRowIndex As Long

    
    '如果没有记录，则退出
    If mDataGrid.Rows - 1 <= 0 Then Exit Sub

    intB = 0
    lngRowIndex = -1

    If Not blnFirst Then
        intB = mDataGrid.Row + 1
        If intB >= mDataGrid.Rows Then intB = 1
    End If
'
    lngSelRow = mDataGrid.Row
    lngEndRow = mDataGrid.Rows - 1

continue1:
   '根据字段名定位的处理
    If mDataGrid.ColIndex(strName) > 0 Then
        lngRowIndex = mDataGrid.FindRow(strFilter, intB, mDataGrid.ColIndex(strName), False, False)
    End If

    If lngRowIndex > 0 Then
        patiSearch.Tag = patiSearch.Text
        mDataGrid.Row = lngRowIndex

        If mDataGrid.TopRow > mDataGrid.Row Then mDataGrid.TopRow = mDataGrid.Row
        If mDataGrid.BottomRow - 1 < mDataGrid.Row Then
            mDataGrid.TopRow = mDataGrid.TopRow + (mDataGrid.Row - mDataGrid.BottomRow) + 1
        End If
    Else
        If intB > 1 Then
            intB = 0
            GoTo continue1:
        End If
        
    End If

    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[SeekNextPati]" & vbCrLf & Err.Description
End Sub

Public Function UpdateRow(ByVal blIsAdd As Boolean, ByVal lngAdviceID As Long) As Boolean
'根据医嘱ID,重新查询一行数据并且刷新列表那一行
'strCol: 列名称 VarValue值  blIsAdd 是否新增列  医嘱ID,lngAdviceID只有新增列需要
'首先判断列表是否包含了被更新的列，若包含，需要刷新列显示
'更新记录集对应行数据
On Error GoTo errH
    Dim rsTemp As ADODB.Recordset
    Dim rsTempShow As ADODB.Recordset
    Dim strColName As String
    Dim i As Integer
    Dim lngAdviceIDOld As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strTmp As String
    
    UpdateRow = False
    If mObjQuery Is Nothing Then Exit Function

    With mObjQuery
'        mTqueryType = 更新一行
        Set rsTemp = .ExecuteWithAttach("[系统.医嘱ID]", lngAdviceID)
        
        Set rsTempShow = CopyRecordSet(rsTemp)
        Set rsTempShow = mObjQuery.DataConvert(rsTempShow, mlngSchemeNo)
        Call rsTemp.MoveFirst
        
        If Not blIsAdd Then

            mrsData.MoveFirst

            While Not mrsData.EOF
                If Val(mrsData!医嘱ID) = lngAdviceID Then

                    For i = 0 To rsTemp.Fields.Count - 1
                        mrsData.Fields(i).Value = rsTemp.Fields(i).Value
                    Next
                    
                    mrsDataShow.MoveFirst
                    While Not mrsDataShow.EOF
                        If Val(mrsDataShow!医嘱ID) = lngAdviceID Then

                            For i = 0 To rsTempShow.Fields.Count - 1
                                mrsDataShow.Fields(i).Value = rsTempShow.Fields(i).Value
                            Next

                            GoTo Refresh
                        End If
                        mrsDataShow.MoveNext
                    Wend

                End If
                mrsData.MoveNext
            Wend
  
        Else
            'lsq待验证有效性
            mrsData.AddNew

            For i = 0 To rsTemp.Fields.Count - 1
                mrsData.Fields(i) = rsTemp.Fields(i)
            Next

        End If

    End With

Refresh:
    '检查列表显示更新
    If blIsAdd Then
        '直接刷新整个列表
        If Not mrsData Is Nothing Then
            If mrsData.RecordCount > 0 Then
                Set mrsDataShow = GetFilterFromQuickFilter
                Set mrsDataShow = mObjQuery.DataConvert(mrsDataShow, mlngSchemeNo)
                
                If Not mrsDataShow Is Nothing Then
                    lngAdviceIDOld = GetSelectRowAdviceID
                    mblSearching = True
                    Set vsfList.DataSource = mrsDataShow
                    mblSearching = False
                    '列统计
                    Call ColStatistics(mrsDataShow)
                    
                    lngRow = vsfList.FindRow(lngAdviceID, 1, vsfList.ColIndex(mstrListKeyCol))
                    If lngRow = -1 Then lngRow = 0
                    Call LoadListHeadCfg(lngRow)
                Else
                    Call LoadListHeadCfg(0)
                End If

            End If
        End If
    Else
        lngRow = vsfList.FindRow(lngAdviceID, 1, vsfList.ColIndex(mstrListKeyCol))
        If lngRow = -1 Then
            UpdateRow = True
            Exit Function
        End If
                    
        '更新一行的数据
        For i = 1 To vsfList.Cols - 1
            strColName = vsfList.TextMatrix(0, i)
            vsfList.TextMatrix(lngRow, i) = NVL(rsTempShow.Fields(strColName).Value)
        Next

        Call RefreshRowRelation(lngRow)
    End If
    
    UpdateRow = True
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[UpdateRow]" & vbCrLf & Err.Description
End Function

Private Function GetQuickFilterSQLPar(ByVal lngIndex As Long) As String
'获取与组合自定义快速过滤条件，类似"头,双眼,甲状腺"这种，用于过滤
'参数： 过滤信息index，菜单ID
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngID As Long, lngIDEnd As Long
    
    Dim strTmp As String
    Dim cbrPopControl As CommandBarControl
    Dim objControl As CommandBarControl
    Dim blSimpleFilter As Boolean
    Dim blChooseOne As Boolean
    Dim strAry() As String
    
    blSimpleFilter = mTQuickFilterState.TCmdState(lngIndex).blSimpleFilter
    lngID = 100 * lngIndex
    lngIDEnd = mTQuickFilterState.TCmdState(lngIndex).intItemCount
    blChooseOne = False
    
    For i = 1 To lngIDEnd
        Set objControl = cbrFilter.FindControl(, lngID - 100 + i, , True)
        
        If mTQuickFilterState.TCmdState(lngIndex).cmdItem(i).blChoose Then
            blChooseOne = True
            If blSimpleFilter Then
                If Len(strTmp) > 0 Then strTmp = strTmp & ","
                strTmp = strTmp & objControl.Caption
            Else
                If Len(strTmp) > 0 Then strTmp = strTmp & ","
                strTmp = strTmp & objControl.Category
            End If
        End If
        
    Next
    
    If blChooseOne = False Then
        For i = 1 To lngIDEnd
            Set objControl = cbrFilter.FindControl(, lngID - 100 + i, , True)
            If blSimpleFilter Then
                If Len(strTmp) > 0 Then strTmp = strTmp & ","
                strTmp = strTmp & objControl.Caption
            Else
                If Len(strTmp) > 0 Then strTmp = strTmp & ","
                strTmp = strTmp & objControl.Category
            End If
        Next
    End If
    
    strAry = Split(strTmp, ",")
    strTmp = ""
    For i = 0 To UBound(strAry)
        For j = i + 1 To UBound(strAry)
            If strAry(i) = strAry(j) Then strAry(j) = ""
        Next
        
        If strAry(i) <> "" Then
            If Len(strTmp) > 0 Then strTmp = strTmp & ","
            strTmp = strTmp & strAry(i)
        End If
    Next
    
    strTmp = Replace(strTmp, ",,,,,,", ",")
    strTmp = Replace(strTmp, ",,,,,", ",")
    strTmp = Replace(strTmp, ",,,,", ",")
    strTmp = Replace(strTmp, ",,,", ",")
    strTmp = Replace(strTmp, ",,", ",")
    
    mTQuickFilterState.TCmdState(lngIndex).strRelationValueForVBSFilter = Trim(strTmp)

    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetQuickFilterSQLPar]" & vbCrLf & Err.Description
End Function

Private Function CbrFilterDeal(ByVal strFilterOld As String, ByVal strFilterNew As String) As String
On Error GoTo errH
    Dim i As Integer
    Dim strNew() As String
    Dim strOld As String
    
    strOld = "," & strFilterOld & ","
    
    strNew = Split(strFilterNew, ",")
    
    For i = 0 To UBound(strNew)
        If InStr(strOld, "," & strNew(i) & ",") = 0 Then
            If Len(CbrFilterDeal) > 0 Then CbrFilterDeal = CbrFilterDeal & ","
            CbrFilterDeal = CbrFilterDeal & strNew(i)
        End If
    Next
    
    CbrFilterDeal = "," & CbrFilterDeal
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[CbrFilterDeal]" & vbCrLf & Err.Description
End Function

Private Function DoBeforeFilterData(ByVal blOnlyChangeCol As Boolean) As Boolean
'查询之后或者列头结构改变后，DataSource之前的处理。
'blChangeCol 是否只是改变Col宽度、顺序这种操作。若是，不需要清空list
On Error GoTo errH
    Dim i As Integer, j As Integer
    
    If mrsData Is Nothing Or mSqlScheme Is Nothing Then Exit Function
    
    If blOnlyChangeCol Then
        With vsfList
            ReDim mColCfgInfo(.Cols - 1)
            
            For i = 1 To .Cols - 1
                .ColKey(i) = .TextMatrix(0, i)
                
                For j = 1 To mSqlScheme.ShowCfgCount
                    mColCfgInfo(i) = 0
                    If mSqlScheme.ShowCfg(j).Name = .ColKey(i) Then
                        mColCfgInfo(i) = j
                        Exit For
                    End If
                Next
            Next
        End With
    Else
        ReDim mColCfgInfo(mrsData.Fields.Count)
        
        With vsfList
            If Not blOnlyChangeCol Then
                .Clear
                .Cols = mrsData.Fields.Count + 1
            End If
            
            For i = 1 To mrsData.Fields.Count
                .ColKey(i) = mrsData.Fields(i - 1).Name
                
                For j = 1 To mSqlScheme.ShowCfgCount
                    mColCfgInfo(i) = 0
                    If mSqlScheme.ShowCfg(j).Name = .ColKey(i) Then
                        mColCfgInfo(i) = j
                        Exit For
                    End If
                Next
            Next
        End With
    End If
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[DoBeforeFilterData]" & vbCrLf & Err.Description
End Function

Private Function GetDefaultColCfg() As Boolean
'获取初始列配置信息（名称，是否隐藏）用于恢复列状态
On Error GoTo errH
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim strCol() As String
    Dim picLoad As StdPicture
    Dim blHaveCfg As Boolean

    If mSqlScheme Is Nothing Or Len(mstrSchemeCfg.strListCfgDefault) > 0 Then Exit Function

    '最后一个没有内容
    strCol = Split(mstrSchemeCfg.strListCfgDefaultColOrder, "|")
    If mSqlScheme.ShowCfgCount > 0 Then
        For i = 0 To UBound(strCol) - 1   'for 1
            For j = 1 To mSqlScheme.ShowCfgCount   'for 2
                blHaveCfg = False
                If strCol(i) = mSqlScheme.ShowCfg(j).Name Then
                    blHaveCfg = True
    '                    隐藏列
                    If mSqlScheme.ShowCfg(j).HiddenCol Then
                        strCol(i) = strCol(i) & "," & TextWidth(strCol(i)) & "," & "0"
                    Else
                        strCol(i) = strCol(i) & "," & TextWidth(strCol(i)) & "," & "1"
                    End If
                    
                    Exit For  '跳出for 2
                End If
                            
            Next ' for 2
            
            '能执行到这里说明没有对应的列配置设置
            If Not blHaveCfg Then
                If Len(strCol(i)) > 0 Then strCol(i) = strCol(i) & "," & TextWidth(strCol(i)) & "," & "1"
            End If
    
        Next ' for 1
    Else
        For i = 0 To UBound(strCol) - 1   'for 1
            strCol(i) = strCol(i) & "," & TextWidth(strCol(i)) & "," & "1"
        Next ' for 1
    End If
    

    For i = 0 To UBound(strCol) - 1
        If Len(mstrSchemeCfg.strListCfgDefault) > 0 Then mstrSchemeCfg.strListCfgDefault = mstrSchemeCfg.strListCfgDefault & "|"
        mstrSchemeCfg.strListCfgDefault = mstrSchemeCfg.strListCfgDefault & strCol(i)
    Next
        
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetDefaultColCfg]" & vbCrLf & Err.Description
End Function


Private Function DoListCfg() As Boolean
'处理列表配置
'此时list中已经有数据，需要进行列头配置，顺序改变等操作
'根据列配置，处理列显示信息（图标，是否隐藏等）
On Error GoTo errH
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim strCol() As String
    Dim picLoad As StdPicture
    Dim ObjScShowCfg As New clsScShowCfg
    Dim blHaveCfg As Boolean

    If mSqlScheme Is Nothing Then Exit Function
    
    If Len(mstrSchemeCfg.strListCfgDefault) = 0 Then
        Call GetDefaultColCfg
    End If
        
    strCol = Split(mstrSchemeCfg.strListCfgDefaultColOrder, "|")
    
    'i:初始配置中字段序号  j:列显示配置序号
    With vsfList
        For i = 1 To UBound(strCol) - 1   'for 1
            
            '根据配置进行一些调整
            If mColCfgInfo(i) > 0 Then
                Set ObjScShowCfg = mSqlScheme.ShowCfg(mColCfgInfo(i))
                           
                If Len(ObjScShowCfg.Icon) > 0 Then
                    '列图标操作
                    Set picLoad = GetIcon(ObjScShowCfg.Icon)
                    Set .Cell(flexcpPicture, 0, i) = GetIcon(ObjScShowCfg.Icon)
        '
        '                '需要有效的缩小图片的方式[LSQB2]
        ''                If imgList16.ImageHeight > .RowHeight(0) Then .RowHeight(0) = imgList16.ImageHeight
        '
        '                If picLoad.Height > .RowHeight(0) Then
        '                    .RowHeight(0) = picLoad.Height
        '                End If
                End If
                
                '隐藏数据显示 统一处理 在前方增加空格
                If ObjScShowCfg.HiddenData Then
                
                    For j = 1 To .Rows - 1
                        .Cell(flexcpText, j, i) = "                                " & .Cell(flexcpText, j, i)
                        .Cell(flexcpAlignment, j, i) = flexAlignLeftTop
                    
                    Next
                End If
                
            End If
          
        Next ' for 1
    End With
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[DoListCfg]" & vbCrLf & Err.Description
End Function

Private Function GetIcon(ByVal strID As String) As StdPicture
'通过ID获取图标，判断字典中是否已经存在该图标，若存在，使用字典中对象，若不存在。调用mObjQuery.GetIconRes并且添加到字典
On Error GoTo errH
    Dim stdPic As StdPicture
    
    If mPicDictionary.Exists(strID) Then
        Set GetIcon = mPicDictionary.Item(strID)
    Else
        Set stdPic = mObjQuery.GetIconRes(strID)
        Call mPicDictionary.Add(strID, stdPic)
        Set GetIcon = stdPic
    End If
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetIcon]" & vbCrLf & Err.Description
End Function

Private Sub ResetSort(ByVal lngCol As Long, ByVal lngWay As Long)
'重置排序
On Error GoTo errH
    Dim RowIndex As Long
    
    If vsfList.Rows <= 1 Then Exit Sub
    
    If vsfList.Col <> vsfList.ColIndex(GetColSort(vsfList.ColKey(lngCol))) Then
        vsfList.Col = vsfList.ColIndex(GetColSort(vsfList.ColKey(lngCol)))
        
        '奇数  和  偶数 是排序的两个方向
        If lngWay = 2 Or lngWay = 4 Or lngWay = 6 Or lngWay = 8 Then
            vsfList.Sort = 4
        Else
            vsfList.Sort = 3
        End If
    Else
        vsfList.Col = lngCol
        vsfList.Sort = lngWay
    End If
    
    If vsfList.TopRow = vsfList.BottomRow Then Exit Sub
    For RowIndex = vsfList.TopRow To vsfList.BottomRow
        Call RefreshRowRelation(RowIndex)
    Next
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[ResetSort]" & vbCrLf & Err.Description
End Sub

Private Sub GetSchemePara()
'获取方案参数（配置中影像界面显示的设置）
    mTLayout.blShowHistory = mSqlScheme.ShowHistory
    mTLayout.blShowQuickFilter = mSqlScheme.FilterCfgCount > 0
End Sub

Private Sub GetLocalPara()
'获取本地参数（注册表参数）
    mlngSortCol = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "排序列", 0))
    mintSortOrder = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "排序方向", 0))
    mlngMove = Val(GetSetting("ZLSOFT", "私有模块\" & mstrDBUser & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "列表检查信息高度设置", 0))
    
    mTPatiIdentifyInfo.strFindItem = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "查找项目")
    mTPatiIdentifyInfo.strLocateItem = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "定位项目")
    
    mTPatiIdentifyInfo.blFind = (GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "是否查找", "1") = "1")
End Sub

Private Sub SaveLocalPara()
'设置本地参数
    Call SaveSetting("ZLSOFT", "私有模块\" & mstrDBUser & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "列表检查信息高度设置", mlngMove)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "排序列", mlngSortCol)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "排序方向", mintSortOrder)
    
    Call SaveLocalPara_PatiIdentify
End Sub

Private Sub SaveLocalPara_PatiIdentify()
'设置Pati控件相关参数
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "查找项目", mTPatiIdentifyInfo.strFindItem)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "定位项目", mTPatiIdentifyInfo.strLocateItem)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "是否查找", IIf(mTPatiIdentifyInfo.blFind, 1, 0))
End Sub



Private Sub LoadPatiIdentifyInfo()
'第一次加载方案是需要 获取查找/定位选项，后面可以直接使用
'查找功能：默认数据+过滤数据
'定位功能：列表配置
'同时判断是否需要显示Pati控件
On Error GoTo errH
    Dim i As Integer
    Dim blHaveFind As Boolean '是否有查找项目
    Dim blHaveLocate As Boolean '是否有定位项目
    Dim strSql As String
    
    blHaveFind = False
    blHaveLocate = False
    
    '若还未加载过就先加载
    If Not mTPatiIdentifyInfo.blHaveLoad Then
        
        For i = 1 To mSqlScheme.SerachCfgCount
        
            If mSqlScheme.SerachCfg(i).InputType = itFast Or mSqlScheme.SerachCfg(i).InputType = itBoth Then
                mTPatiIdentifyInfo.strFindItems = mTPatiIdentifyInfo.strFindItems & mSqlScheme.SerachCfg(i).Name & ";"
                blHaveFind = True
            End If
            
        Next
        
        If mTPatiIdentifyInfo.strFindItems = "" Then mTPatiIdentifyInfo.strFindItems = "姓名;"
        
        For i = 1 To mSqlScheme.ShowCfgCount
            If mSqlScheme.ShowCfg(i).UseListLocate Then
                mTPatiIdentifyInfo.strLocateItems = mTPatiIdentifyInfo.strLocateItems & mSqlScheme.ShowCfg(i).Name & ";"
                blHaveLocate = True
            End If
        Next
        
        If mTPatiIdentifyInfo.strLocateItems = "" Then mTPatiIdentifyInfo.strLocateItems = "姓名;"
        
        mTPatiIdentifyInfo.blHaveLoad = True
    
        If blHaveFind Or blHaveLocate Then mTPatiIdentifyInfo.blShowPatiIdentify = True
        
        '判断是否有时间范围条件
        strSql = mSqlScheme.GetScheme
        If InStr(strSql, "[系统.开始日期]") > 0 And InStr(strSql, "[系统.结束日期]") Then
            mTLayout.blShowTimeSelect = True
        Else
            mTLayout.blShowTimeSelect = False
        End If
        
        mTLayout.blShowBaseFilter = mTPatiIdentifyInfo.blShowPatiIdentify Or mTLayout.blShowTimeSelect
    End If

    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[LoadPatiIdentifyInfo]" & vbCrLf & Err.Description
End Sub

Public Sub RefreshRowRelation(ByVal lngRow As Long)
'单独刷新某一行，同时触发行关联处理
On Error GoTo errH
    Dim Value As String
    Dim i As Integer
    
    For i = 1 To vsfList.Cols - 1
        Value = vsfList.TextMatrix(lngRow, i)
        Call RowRelationConvert(lngRow, i, Value)
    Next
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[RefreshRowRelation]" & vbCrLf & Err.Description
    
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    mbytFontSize = bytFontSize
End Sub

Public Sub ReSetFormFontSize(Optional bytFontSize As Byte = 0)
On Error Resume Next
    
    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType As String
    
    If bytFontSize > 0 Then mbytFontSize = bytFontSize
    
    Me.FontSize = mbytFontSize
    Set CtlFont = New StdFont
    strFontType = IIf(IsUseClearType = True, "微软雅黑", "宋体")
    CtlFont.Name = strFontType
    CtlFont.Size = mbytFontSize
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") '页面控件
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = mbytFontSize
        Case UCase("Label")
            If objCtrl.Name <> "lblCash" Then
                objCtrl.Font.Name = strFontType
                objCtrl.FontSize = mbytFontSize
                objCtrl.Height = TextHeight("罗") + 60
            End If
        Case UCase("vsFlexGrid")
            objCtrl.Font = CtlFont
        Case UCase("ComboBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("罗冠" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("罗冠" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("罗") * 1.5
        Case UCase("textBox")
          objCtrl.FontName = strFontType
          objCtrl.FontSize = mbytFontSize
        Case UCase("ReportControl")
            
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            
            Set objCtrl.Options.Font = CtlFont
            
        Case UCase("TabControl")
            Set objCtrl.PaintManager.Font = CtlFont
            
        Case UCase("CommandButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
            
        Case UCase("PatiIdentify")
            objCtrl.CardNoShowFont.Size = mbytFontSize
            objCtrl.Font.Size = mbytFontSize
            objCtrl.IDKindFont.Size = mbytFontSize
            If mbytFontSize = 9 Then
                objCtrl.Height = 330
            ElseIf mbytFontSize = 12 Then
                objCtrl.Height = 360
            ElseIf mbytFontSize = 15 Then
                objCtrl.Height = 390
            End If
            objCtrl.Refrash
        
        End Select
    Next
    
    Call AdjustFace(mbytFontSize)
    
End Sub

Private Sub AdjustFace(ByVal bytFontSize As Byte)
'字号 目前工作站支持9,12,15三种
On Error Resume Next
    Dim lngHeight页签 As Long
    Dim lngHeight基本过滤 As Long
    Dim lngHeight快速过滤 As Long
    Dim lngHeight列表 As Long
    Dim lngHeight历史检查 As Long
    Dim lngHeight病人信息 As Long
    Dim lngHeight详细信息 As Long
    Dim lngHeight分割线 As Long
    
    '这里的 10000  1000是大概规定的分割线有效移动范围
    If mlngMove > 10000 Then mlngMove = 10000
    If mlngMove < -1000 Then mlngMove = -1000
    
    lngHeight分割线 = 50
    
    If bytFontSize = 9 Then
        lngHeight页签 = 300
        
        lngHeight基本过滤 = IIf(mTLayout.blShowBaseFilter, 350, 0)
        lngHeight快速过滤 = IIf(mTLayout.blShowQuickFilter, 400, 0)
        
        lngHeight历史检查 = IIf(Label1.Height > cboHistory.Height, Label1.Height, cboHistory.Height) + 60
        lngHeight历史检查 = IIf(mTLayout.blShowHistory, lngHeight历史检查, 0)
        
        lngHeight病人信息 = labPatientInfoName.Height + 90
    
        lngHeight详细信息 = C_LAYOUT_BASEHEIGHTOFDETAILINFO + mlngMove
        
        lngHeight列表 = Me.ScaleHeight - lngHeight页签 - lngHeight基本过滤 - lngHeight快速过滤 - lngHeight历史检查 - lngHeight病人信息 - lngHeight详细信息
    
        Call tabQuery.Move(C_LAYOUT_LISTLEFT, 0, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight页签)
    
        Call picSearch.Move(C_LAYOUT_LISTLEFT, tabQuery.Top + tabQuery.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight基本过滤)
    
        Call picFilter.Move(C_LAYOUT_LISTLEFT, picSearch.Top + picSearch.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight快速过滤)
    
        Call picVsf.Move(C_LAYOUT_LISTLEFT, picFilter.Top + picFilter.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight列表)
        
        Call PicLine.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight分割线)
    
        Call picHistory.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight历史检查)
    
        Call picListRowInfo.Move(C_LAYOUT_LISTLEFT, picHistory.Top + picHistory.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight病人信息)
        Call txtDetail.Move(C_LAYOUT_LISTLEFT, picListRowInfo.Top + picListRowInfo.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight详细信息)
    
    ElseIf bytFontSize = 12 Then
        lngHeight页签 = 360
        
        lngHeight基本过滤 = IIf(mTLayout.blShowBaseFilter, 390, 0)
        lngHeight快速过滤 = IIf(mTLayout.blShowQuickFilter, 420, 0)
        
        lngHeight历史检查 = IIf(Label1.Height > cboHistory.Height, Label1.Height, cboHistory.Height) + 60
        lngHeight历史检查 = IIf(mTLayout.blShowHistory, lngHeight历史检查, 0)
        
        lngHeight病人信息 = labPatientInfoName.Height + 90
    
        lngHeight详细信息 = C_LAYOUT_BASEHEIGHTOFDETAILINFO + mlngMove
        
        lngHeight列表 = Me.ScaleHeight - lngHeight页签 - lngHeight基本过滤 - lngHeight快速过滤 - lngHeight历史检查 - lngHeight病人信息 - lngHeight详细信息
    
        Call tabQuery.Move(C_LAYOUT_LISTLEFT, 0, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight页签)
    
        Call picSearch.Move(C_LAYOUT_LISTLEFT, tabQuery.Top + tabQuery.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight基本过滤)
    
        Call picFilter.Move(C_LAYOUT_LISTLEFT, picSearch.Top + picSearch.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight快速过滤)
    
        Call picVsf.Move(C_LAYOUT_LISTLEFT, picFilter.Top + picFilter.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight列表)
        
        Call PicLine.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight分割线)
    
        Call picHistory.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight历史检查)
    
        Call picListRowInfo.Move(C_LAYOUT_LISTLEFT, picHistory.Top + picHistory.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight病人信息)
        Call txtDetail.Move(C_LAYOUT_LISTLEFT, picListRowInfo.Top + picListRowInfo.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight详细信息)
    
    ElseIf bytFontSize = 15 Then
        lngHeight页签 = 420
        
        lngHeight基本过滤 = IIf(mTLayout.blShowBaseFilter, 430, 0)
        lngHeight快速过滤 = IIf(mTLayout.blShowQuickFilter, 440, 0)
        
        lngHeight历史检查 = IIf(Label1.Height > cboHistory.Height, Label1.Height, cboHistory.Height) + 60
        lngHeight历史检查 = IIf(mTLayout.blShowHistory, lngHeight历史检查, 0)
        
        lngHeight病人信息 = labPatientInfoName.Height + 90
    
        lngHeight详细信息 = C_LAYOUT_BASEHEIGHTOFDETAILINFO + mlngMove
        
        lngHeight列表 = Me.ScaleHeight - lngHeight页签 - lngHeight基本过滤 - lngHeight快速过滤 - lngHeight历史检查 - lngHeight病人信息 - lngHeight详细信息
    
        Call tabQuery.Move(C_LAYOUT_LISTLEFT, 0, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight页签)
    
        Call picSearch.Move(C_LAYOUT_LISTLEFT, tabQuery.Top + tabQuery.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight基本过滤)
    
        Call picFilter.Move(C_LAYOUT_LISTLEFT, picSearch.Top + picSearch.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight快速过滤)
    
        Call picVsf.Move(C_LAYOUT_LISTLEFT, picFilter.Top + picFilter.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight列表)
        
        Call PicLine.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight分割线)
    
        Call picHistory.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight历史检查)
    
        Call picListRowInfo.Move(C_LAYOUT_LISTLEFT, picHistory.Top + picHistory.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight病人信息)
        Call txtDetail.Move(C_LAYOUT_LISTLEFT, picListRowInfo.Top + picListRowInfo.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight详细信息)
    End If
    
End Sub

Private Sub ColStatistics(rsData As Recordset)
'列统计处理
On Error GoTo errHandle
    Dim i As Long, j As Long, k As Long
    Dim lngColICount As Long '待统计的数据序号
    Dim lngColIndex() As Integer
    
    Dim strColName As String '"需要处理跟踪列的列" 类似： "影像质量;检查过程;是否阳性"
    Dim strColNameAll As String
    Dim strStateBarInfo As String '最终通过事件传递的列统计信息
    Dim strInfoTmp As String
    
    Dim objTColTotalInfo() As TColTotalInfo
    Dim DictColTotal() As Dictionary
    
    strColNameAll = ""
    strColName = ""
    lngColICount = 0
    
    If mSqlScheme Is Nothing Or rsData Is Nothing Then
        RaiseEvent OnColStatistics("")
        Exit Sub
    End If
    
    If rsData.RecordCount = 0 Then
        RaiseEvent OnColStatistics("")
        Exit Sub
    End If
     
    With mSqlScheme
        For i = 1 To .ShowCfgCount
            If .ShowCfg(i).IsTotal Then
                lngColICount = lngColICount + 1
                strColNameAll = strColNameAll & .ShowCfg(i).Name & ";"
                ReDim Preserve DictColTotal(lngColICount)
                Set DictColTotal(lngColICount) = New Dictionary
            End If
        Next
    End With
        
    With rsData
        
        strStateBarInfo = ""
        strColName = ""
        
        For k = 1 To lngColICount
        
            '获取字段名
            strColName = Split(strColNameAll, ";")(k - 1)

            .MoveFirst
            While Not .EOF
                If Not IsNull(.Fields(strColName).Value) Then
                    If Len(Trim(.Fields(strColName).Value)) > 0 Then
                        If DictColTotal(k).Exists(.Fields(strColName).Value) Then
                             DictColTotal(k).Item(.Fields(strColName).Value) = DictColTotal(k).Item(.Fields(strColName).Value) + 1
                        Else
                            Call DictColTotal(k).Add(.Fields(strColName).Value, 1)
                        End If
                        
                    End If
                End If
            .MoveNext
            Wend
            
            strInfoTmp = "[" & strColName & "]:"
            For j = 1 To DictColTotal(k).Count
                strInfoTmp = strInfoTmp & DictColTotal(k).Keys(j - 1) & ":" & DictColTotal(k).Item(DictColTotal(k).Keys(j - 1)) & " "
            Next
            
            If Len(strStateBarInfo) > 0 Then strStateBarInfo = strStateBarInfo & "|"
            strStateBarInfo = strStateBarInfo & strInfoTmp
            
            RaiseEvent OnColStatistics(strStateBarInfo)
                
        Next
   
        .MoveFirst
    End With
    
    RaiseEvent OnColStatistics(strStateBarInfo)
    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[ColStatistics]" & vbCrLf & Err.Description
End Sub

Private Sub RowRelationConvert(ByVal Row As Long, ByVal Col As Long, Value As String)
'行关联处理 主要处理 颜色，图标 每行只需要处理一次行颜色
On Error GoTo errH
    Dim i As Integer, j As Integer, k As Integer
    Dim lngRelationIndex As Long
    Dim objClsRelation As New clsScRowRelation
    Dim blContinue As Boolean '是否继续
    Dim lngColColor As Long '行颜色列
    Dim strColorColValue As String '行颜色列内容 如 "已登记"
    Static TpRowColorInfo As TRowColorInfo
       
    lngRelationIndex = mColCfgInfo(Col)
    If UBound(mColCfgInfo) < Col - 1 Then Exit Sub
    
    If lngRelationIndex < 1 Then
        If Col <> 1 Then
            Exit Sub
        End If
    Else
        If mSqlScheme.ShowCfg(lngRelationIndex).RowRelationCount < 1 Then
            If Col <> 1 Then
                Exit Sub
            End If
        End If
    End If

    If Col = 1 Then
        '首次加载方案后获取行颜色相关列
        
        If TpRowColorInfo.LngSchemeNo <> mlngSchemeNo Then
            TpRowColorInfo.blHaveRowColor = False
            TpRowColorInfo.LngSchemeNo = mlngSchemeNo
            
            For i = 1 To mSqlScheme.ShowCfgCount
            
                For j = 1 To mSqlScheme.ShowCfg(i).RowRelationCount
                    If mSqlScheme.ShowCfg(i).RowRelation(j).RowBackColor > 0 Or mSqlScheme.ShowCfg(i).RowRelation(j).RowFontColor > 0 Then
                        TpRowColorInfo.intRowColorIndex = i
                        TpRowColorInfo.blHaveRowColor = True
                        Exit For
                    End If
                Next
                
            Next
    
        End If
        
        If TpRowColorInfo.blHaveRowColor Then
        '存在行颜色才需要进行下面的处理
            If TpRowColorInfo.intRowColorIndex > 0 Then
                With vsfList
                    '首先处理行关联行颜色
                    For i = 1 To mSqlScheme.ShowCfg(TpRowColorInfo.intRowColorIndex).RowRelationCount
                        Set objClsRelation = mSqlScheme.ShowCfg(TpRowColorInfo.intRowColorIndex).RowRelation(i)
    
                        lngColColor = .ColIndex(mSqlScheme.ShowCfg(TpRowColorInfo.intRowColorIndex).Name)
                        strColorColValue = .TextMatrix(Row, lngColColor)
                        If strColorColValue = objClsRelation.TiggerData Then
                            '行背景色
                            If objClsRelation.RowBackColor > 0 Then .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = objClsRelation.RowBackColor
                            '行前景色
                            If objClsRelation.RowFontColor > 0 Then .Cell(flexcpForeColor, Row, 1, Row, .Cols - 1) = objClsRelation.RowFontColor
                        End If
                    Next
                End With
                
                mrsDataShow.MoveFirst
            End If
        End If
        
    End If
    
    'lngRelationIndex <1 说明不需要后面的处理
    If lngRelationIndex < 1 Then Exit Sub

    With vsfList
        '首先处理行关联
        If mSqlScheme.ShowCfg(lngRelationIndex).RowRelationCount > 0 Then
            For i = 1 To mSqlScheme.ShowCfg(lngRelationIndex).RowRelationCount
                Set objClsRelation = mSqlScheme.ShowCfg(lngRelationIndex).RowRelation(i)

                If Value = objClsRelation.TiggerData Then
                 '图标显示及 指定列显示
                    If Val(objClsRelation.Icon) > 0 Then
                        '图标显示列
                        If Len(objClsRelation.IconPerformCol) > 0 Then
                            Set .Cell(flexcpPicture, Row, .ColIndex(objClsRelation.IconPerformCol)) = GetIcon(objClsRelation.Icon)
                        Else
                        '图标显示
                            Set .Cell(flexcpPicture, Row, Col) = GetIcon(objClsRelation.Icon)
                        End If
                    End If

                    If Len(objClsRelation.ColorPerformCol) > 0 Then
                        If objClsRelation.CellBackColor > 0 Or objClsRelation.CellFontColor > 0 Then
                            For k = 1 To .Cols - 1
                                If .Cell(flexcpText, 0, k) = objClsRelation.ColorPerformCol Then
                                    '颜色显示列
                                    If objClsRelation.RowBackColor > 0 Then .Cell(flexcpBackColor, Row, k) = objClsRelation.CellBackColor
                                    '颜色显示列
                                    If objClsRelation.RowFontColor > 0 Then .Cell(flexcpForeColor, Row, k) = objClsRelation.CellFontColor
        
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        '当前cell背景色
                        If objClsRelation.CellBackColor > 0 Then .Cell(flexcpBackColor, Row, Col) = objClsRelation.CellBackColor
                        '当前cell前景色
                        If objClsRelation.CellFontColor > 0 Then .Cell(flexcpForeColor, Row, Col) = objClsRelation.CellFontColor
                    End If

                End If
            Next
        End If
    End With

    Exit Sub

errH:
    Err.Raise -1, "frmPacsQuery", "[RowRelationConvert]" & vbCrLf & Err.Description
'    MsgBox Err.Description
End Sub


Private Function GetColSort(ByVal strColName As String) As String
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
'获取排序列
    GetColSort = strColName
    
    If mTColSort.LngSchemeNo <> mlngSchemeNo Then
        '首先获取一次排序信息
        mTColSort.LngSchemeNo = mlngSchemeNo
        Set mTColSort.dictSortInfo = New Dictionary
        
        If mSqlScheme Is Nothing Then Exit Function
        
        For i = 1 To mSqlScheme.ShowCfgCount
            If Len(mSqlScheme.ShowCfg(i).SortContrastCol) > 0 Then
                Call mTColSort.dictSortInfo.Add(mSqlScheme.ShowCfg(i).Name, mSqlScheme.ShowCfg(i).SortContrastCol)
            End If
        Next
    
    End If
    
    If mTColSort.dictSortInfo.Exists(strColName) Then
        GetColSort = mTColSort.dictSortInfo.Item(strColName)
    End If
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetColSort]" & vbCrLf & Err.Description
End Function

Public Function SetOrder(ByVal lngCurSortCol As Long, ByVal lngCurOrder As Long) As Long
'单独排序，替代控件自带的排序（参考vsflexgrid的排序demo）
On Error GoTo errH
    SetOrder = lngCurOrder
    
     '没有数据时退出排序
    If vsfList.Rows = 1 Then Exit Function
    
    With vsfList
        Dim R&, c&, RS&, cs&
        .GetSelection R, c, RS, cs
        .Redraw = flexRDNone
    
        ' apply sort to non-empty range
        Dim Row%
        
        For Row = .Rows - 1 To .FixedRows Step -1
            '整行数据为空时，不参与排序
            If Len(.TextMatrix(Row, lngCurSortCol)) Or Not Trim(.TextMatrix(Row, .ColIndex(mstrListKeyCol))) = "" Then Exit For
        Next
        
        If Row > .FixedRows Then
            .Select .FixedRows, lngCurSortCol, Row, lngCurSortCol
            .Sort = lngCurOrder
        End If
        
        ' restore selection
        .Select R, c, RS, cs
        .Redraw = flexRDDirect
        
        ' cancel default sort
        SetOrder = 0
    End With
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[SetOrder]" & vbCrLf & Err.Description
End Function

Private Sub GetNullStudyInfo()
    With mTStudyInfo
        .strPatientAge = ""
        .strPatientName = ""
        .strPatientSex = ""
        .strStudyNum = ""
        .lngLinkId = 0
        .lngPatId = 0
        .lngAdviceID = 0
    End With
End Sub

Private Function GetSelectRowAdviceID() As Long
'根据当前列选中行获取医嘱ID
On Error GoTo errH
    GetSelectRowAdviceID = 0
    If vsfList.Rows < 1 Or vsfList.RowSel < 1 Then Exit Function
    
    GetSelectRowAdviceID = Val(vsfList.TextMatrix(vsfList.RowSel, vsfList.ColIndex(mstrListKeyCol)))
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetSelectRowAdviceID]" & vbCrLf & Err.Description
End Function

Private Sub DoPatiIdentify()
'处理Pati控件下拉项目和当前选中项目
On Error GoTo errH
    mblnAssignment = True
    If mTPatiIdentifyInfo.blFind Then
        patiSearch.IDKindStr = InitCardType(mTPatiIdentifyInfo.strFindItems)
        If mTPatiIdentifyInfo.strFindItem <> "" Then
            patiSearch.IDKindIDX = patiSearch.GetKindIndex(mTPatiIdentifyInfo.strFindItem)
        Else
            mTPatiIdentifyInfo.strFindItem = Split(mTPatiIdentifyInfo.strFindItems, ";")(0)
            patiSearch.IDKindIDX = patiSearch.GetKindIndex(mTPatiIdentifyInfo.strFindItem)
        End If
    Else
        patiSearch.IDKindStr = InitCardType(mTPatiIdentifyInfo.strLocateItems)
        If mTPatiIdentifyInfo.strLocateItems <> "" Then
            patiSearch.IDKindIDX = patiSearch.GetKindIndex(mTPatiIdentifyInfo.strLocateItem)
        Else
            mTPatiIdentifyInfo.strLocateItem = Split(mTPatiIdentifyInfo.strLocateItems, ";")(0)
            patiSearch.IDKindIDX = patiSearch.GetKindIndex(mTPatiIdentifyInfo.strLocateItem)
        End If
    End If
    mblnAssignment = False
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[GetSelectRowAdviceID]" & vbCrLf & Err.Description
End Sub

