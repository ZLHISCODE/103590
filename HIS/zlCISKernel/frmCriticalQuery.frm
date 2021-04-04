VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCriticalQuery 
   Caption         =   "危急值查询"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18090
   Icon            =   "frmCriticalQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   18090
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPatiC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   4050
      ScaleHeight     =   540
      ScaleWidth      =   1575
      TabIndex        =   34
      Top             =   5550
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   150
      ScaleHeight     =   2535
      ScaleWidth      =   4020
      TabIndex        =   2
      Top             =   2490
      Width           =   4020
      Begin VB.ComboBox cboState 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1275
         Width           =   2040
      End
      Begin VB.PictureBox picLX 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1245
         ScaleHeight     =   285
         ScaleWidth      =   2295
         TabIndex        =   30
         Top             =   1560
         Width           =   2295
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            Caption         =   "外来"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1560
            TabIndex        =   35
            Top             =   45
            Width           =   795
         End
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            Caption         =   "住院"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   750
            TabIndex        =   32
            Top             =   45
            Width           =   795
         End
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            Caption         =   "门诊"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   31
            Top             =   45
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "刷新"
         Height          =   300
         Left            =   3075
         TabIndex        =   24
         Top             =   930
         Width           =   615
      End
      Begin VB.ComboBox cboSelectTime 
         Height          =   300
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   915
         Width           =   2055
      End
      Begin VB.ComboBox cboRegDept 
         Height          =   300
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   510
         Width           =   2055
      End
      Begin VB.ComboBox cboPatiDept 
         Height          =   300
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   135
         Width           =   2055
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "记录状态"
         Height          =   180
         Left            =   165
         TabIndex        =   36
         Top             =   1305
         Width           =   720
      End
      Begin VB.Label lblLX 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病人类型"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   33
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lblRegTime 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "登记时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   975
         Width           =   720
      End
      Begin VB.Label lblRegDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "登记科室"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   5
         Top             =   555
         Width           =   720
      End
      Begin VB.Label lblPatiDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "确认科室"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   3
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.Timer timeRefreshCard 
      Interval        =   1000
      Left            =   3495
      Top             =   6390
   End
   Begin VB.PictureBox picCItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2190
      Index           =   0
      Left            =   11010
      ScaleHeight     =   2190
      ScaleWidth      =   1800
      TabIndex        =   21
      Top             =   5445
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Label lblAge 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   28
         Top             =   810
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblSex 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   0
         Left            =   255
         TabIndex        =   27
         Top             =   810
         Width           =   360
      End
      Begin VB.Label lblTime 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "报告时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   26
         Top             =   165
         Width           =   720
      End
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   25
         Top             =   510
         Width           =   450
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "内容描述"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   22
         Top             =   1095
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblSelect 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   435
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.PictureBox picCardFra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   11805
      ScaleHeight     =   2730
      ScaleWidth      =   4275
      TabIndex        =   19
      Top             =   4260
      Width           =   4275
      Begin VB.PictureBox picCardCon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   1470
         ScaleHeight     =   1500
         ScaleWidth      =   2160
         TabIndex        =   29
         Top             =   375
         Width           =   2160
      End
      Begin VB.VScrollBar vscH 
         Height          =   2625
         LargeChange     =   10
         Left            =   3975
         Max             =   100
         SmallChange     =   5
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   250
      End
   End
   Begin VB.Frame fraPati 
      Height          =   1590
      Left            =   4695
      TabIndex        =   11
      Top             =   1215
      Width           =   8265
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1245
         Left            =   210
         ScaleHeight     =   1245
         ScaleWidth      =   7395
         TabIndex        =   12
         Top             =   150
         Width           =   7395
         Begin VB.Image imgPatient 
            Height          =   705
            Left            =   75
            Picture         =   "frmCriticalQuery.frx":6852
            Stretch         =   -1  'True
            Top             =   210
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "姓名"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1365
            TabIndex        =   17
            Top             =   195
            Width           =   600
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "年龄"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   5445
            TabIndex        =   16
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "性别"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   3375
            TabIndex        =   15
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "标识号"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   1485
            TabIndex        =   14
            Top             =   855
            Width           =   540
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "科室"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   3300
            TabIndex        =   13
            Top             =   810
            Width           =   360
         End
      End
   End
   Begin VB.PictureBox picCritical 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2310
      Left            =   6030
      ScaleHeight     =   2310
      ScaleWidth      =   4935
      TabIndex        =   1
      Top             =   2865
      Width           =   4935
      Begin VSFlex8Ctl.VSFlexGrid vsCritical 
         Bindings        =   "frmCriticalQuery.frx":771C
         Height          =   1395
         Left            =   435
         TabIndex        =   10
         Top             =   390
         Width           =   4000
         _cx             =   7056
         _cy             =   2461
         Appearance      =   2
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
         BackColorSel    =   16444122
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
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   1935
      Left            =   4665
      TabIndex        =   0
      Top             =   2835
      Width           =   1395
      _Version        =   589884
      _ExtentX        =   2461
      _ExtentY        =   3413
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   7785
      Width           =   18090
      _ExtentX        =   31909
      _ExtentY        =   635
      SimpleText      =   $"frmCriticalQuery.frx":7730
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCriticalQuery.frx":7777
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   26829
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
   Begin VB.Image imgWJ 
      Height          =   240
      Index           =   0
      Left            =   9945
      Picture         =   "frmCriticalQuery.frx":800B
      Top             =   5475
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCL 
      Height          =   240
      Index           =   0
      Left            =   9570
      Picture         =   "frmCriticalQuery.frx":E85D
      Top             =   5475
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCardBack 
      Height          =   2190
      Index           =   0
      Left            =   7050
      Picture         =   "frmCriticalQuery.frx":150AF
      Top             =   5385
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image imgCardBack 
      Height          =   2190
      Index           =   1
      Left            =   8745
      Picture         =   "frmCriticalQuery.frx":192C8
      Top             =   6045
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image imgDefual 
      Height          =   705
      Left            =   1770
      Picture         =   "frmCriticalQuery.frx":1CC8F
      Stretch         =   -1  'True
      Top             =   150
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLoad 
      Height          =   705
      Left            =   1425
      Picture         =   "frmCriticalQuery.frx":1DB59
      Stretch         =   -1  'True
      Top             =   675
      Visible         =   0   'False
      Width           =   975
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   405
      Top             =   195
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmCriticalQuery.frx":1EA23
      Left            =   405
      Top             =   1125
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCriticalQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PatiCol
    COL_病人ID = 0
    COL_主页ID
    COL_挂号ID
    COL_挂号单
    COL_姓名
    COL_性别
    COL_门诊号
    COL_住院号
    COL_床号
    COL_年龄
    COL_科室
End Enum

Private Enum AdviceCol
    col危急值描述
    col报告时间
    col报告人
    col处理情况
    col确认时间
    col确认人
    col确认科室
    col结果
    
    
    '隐藏列
    colID
    col状态
    col医嘱ID
End Enum

Private Enum e_Ctrl
    e姓名 = 0
    e性别
    e年龄
    e标识号
    e科室
End Enum

Private Const conMenu_View_AppCritical = 200
Private Const clngX = 100 '左上角第一张卡片位置
Private mobjCISJob As Object
Private mclsMipModule As zl9ComLib.clsMipModule '消息对象
Private mlngModul As Long
Private mstrPrivs As String
Private mfrmParent As Object
Private mblnModal As Boolean '显示方式，模态，非模态
Private mint方式 As Integer '0-单个病人查询，1-分类查询
Private mint类型  As Integer '0-门诊，1-住院，2-门诊和住院，3-导航台独立查询功能
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstr挂号单 As String
Private mlng就诊ID As Long
Private mlng科室ID As Long '界面科室ID
Private mlng病区ID As Long '病区ID
Private mint场合 As Integer '0-医生站,1-医技站
Private mblnOK As Boolean

Private mrsCard As ADODB.Recordset '数据信息
Private mlngCntCard As Long '总记录数
Private mblnRefreshCard As Boolean
Private mintCurIndex As Integer '当前选择的卡片下标
Private mlngPreRowCnt As Long '前一界面上一行中的卡片数量
Private mstrPrePati As String '前一个病人
Private mlngPreCardID As Long

Private mintPreTim As Integer
Private mdatB登记 As Date
Private mdatE登记 As Date

Private mint显示方式 As Integer '0-正常查询，1-卡片选择器
Private mlng记录ID As Long

Public Function ShowMe(frmParent As Object, ByVal blnModal As Boolean, ByVal int类型 As Integer, ByVal int场合 As Integer, ByVal lng科室id As Long, ByVal lng病区ID As Long, ByRef objMip As Object) As Boolean
'功能：显示窗体
'参数：frmParent 父窗体 ，blnModal 窗体显示模式，false-非模态，true-模态
'      int类型 病人类型 0-门诊，1-住院，2-门诊和住院，3-导航台独立查询功能
'      int场合 0-医生站,1-医技站
'      lng科室ID 医技站调用时医技科室，医生站调用时，病人科室ID,
'      lng病区ID 住院医生站按病区显示时，界面选择的病区ID
'      objMip 用于发送消息的对象 zl9ComLib.clsMipModule
    Set mfrmParent = frmParent
    mint显示方式 = 0
    mblnModal = blnModal
    mint类型 = int类型
    mlng科室ID = lng科室id
    mint场合 = int场合
    mlng病区ID = lng病区ID
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Me.Show IIF(blnModal, 1, 0), frmParent
    ShowMe = mblnOK
End Function

Public Function ShowMeQuery(ByVal lngSys As Long, ByVal lngModul As Long, ByRef frmParent As Object, ByVal strPrivs As String)
'功能：独立查询功能
    mlngModul = lngModul
    mstrPrivs = strPrivs
    mint显示方式 = 0
    mint类型 = 3
    Set mfrmParent = frmParent
    Me.Show , frmParent
    ShowMeQuery = mblnOK
End Function

Public Function ShowMeSelCard(frmParent As Object, ByVal rsIn As ADODB.Recordset) As Long
'功能：卡片选择器模式
'参数：rsIn 要加载的记录，
'返回：危急值记录ID
    Set frmParent = frmParent
    mint显示方式 = 1
    mlng记录ID = 0
    Set mrsCard = zldatabase.CopyNewRec(rsIn)
    Me.Show 1, frmParent
    ShowMeSelCard = mlng记录ID
End Function

Private Sub cboRegDept_Click()
'功能：切换科室
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long
    Dim objControl As CommandBarControl
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIF(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
        cbsMain_Resize
    Case conMenu_View_AppCritical '查看单子
        Call EditData(2)
    Case conMenu_Edit_Modify '修改
        Call EditData(1)
    Case conMenu_Edit_Delete '删除记录
        Call DeleteData
    Case conMenu_Edit_Send
        Call FunAffirm
    Case conMenu_View_Refresh
        Call LoadPatients
    Case conMenu_Tool_Archive
        If mlng就诊ID <> 0 Then
            Call mobjCISJob.ShowArchive(Me, mlng病人ID, mlng就诊ID)
        End If
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
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
    Case conMenu_Edit_Modify, conMenu_Edit_Delete '删除记录
        Control.Enabled = CanEdit()
    Case conMenu_Tool_Archive '电子病案查阅
        If GetInsidePrivs(1259) = "" Then
            Control.Visible = False
        Else
            Control.Enabled = mlng病人ID <> 0
        End If
    Case conMenu_Edit_Send '危急值确认，只有外来病人可用
        Control.Enabled = False
        If mintCurIndex > 0 Then
            mrsCard.Filter = "ID=" & Val(lblName(mintCurIndex).Tag)
            If Not mrsCard.EOF Then
                If Val(mrsCard!主页ID & "") = 0 And mrsCard!挂号单 & "" = "" Then
                    Control.Enabled = True
                End If
            End If
        End If
    End Select
End Sub

Private Function CanEdit() As Boolean
'功能：当前选择的危急值记录是否可以编辑
    Dim strTmp As String
    Dim strTag As String
    Dim blnEdit  As Boolean
    
    strTag = tbcSub.Selected.Tag
    Select Case strTag
    Case "病人"
    Case "危急值"
        With vsCritical
            If Val(.TextMatrix(.Row, col状态)) = 1 And Val(.TextMatrix(.Row, colID)) <> 0 Then
                blnEdit = True
            Else
                blnEdit = False
            End If
        End With
    Case "明细卡"
        If mintCurIndex > 0 Then
            blnEdit = Not imgCL(mintCurIndex).Visible
        End If
    End Select
    CanEdit = blnEdit
End Function

Private Sub cmdOK_Click()
    Call LoadPatients
End Sub

Private Sub Form_Load()
    Dim intIdx As Integer
    Dim objPane As Pane
        
    If mint显示方式 = 0 Then
        Call RestoreWinState(Me, App.ProductName)
        'CommandBars
        '-----------------------------------------------------
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
        If mint类型 <> 2 Then
            'DockingPane
            '-----------------------------------------------------
            Me.dkpMain.SetCommandBars Me.cbsMain
            Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
            Me.dkpMain.Options.ThemedFloatingFrames = True
            Me.dkpMain.Options.AlphaDockingContext = True
            Set objPane = Me.dkpMain.CreatePane(1, 350, 400, DockLeftOf, Nothing)
            objPane.Title = "过滤条件"
            objPane.Options = PaneNoCloseable Or PaneNoFloatable
        End If
        
        With Me.tbcSub
            With .PaintManager
                .Appearance = xtpTabAppearancePropertyPage2003
                .ClientFrame = xtpTabFrameSingleLine
                .BoldSelected = True
                .OneNoteColors = True
                .ShowIcons = True
            End With
            
            .InsertItem(intIdx, "列表", picCritical.Hwnd, 0).Tag = "危急值": intIdx = intIdx + 1
            .InsertItem(intIdx, "卡片", picCardFra.Hwnd, 0).Tag = "明细卡": intIdx = intIdx + 1
            .Item(1).Selected = True

            If mint类型 = 2 Then
                .Item(0).Visible = False
            End If
             
        End With
        
        If mint类型 = 3 Then
            Call Init登记科室
            Call Init确认科室
            
            Set mobjCISJob = CreateObject("zl9CISJob.clsCISJob")
        End If
        mintPreTim = -1
        With cboSelectTime
            .Clear
            .AddItem "今天内"
            .ItemData(.NewIndex) = 0
            .AddItem "昨天内"
            .ItemData(.NewIndex) = 1
            .AddItem "前天内"
            .ItemData(.NewIndex) = 2
            .AddItem "一周内"
            .ItemData(.NewIndex) = 7
            .AddItem "30天内"
            .ItemData(.NewIndex) = 30
            .AddItem "60天内"
            .ItemData(.NewIndex) = 60
            .AddItem "[指定...]"
            .ItemData(.NewIndex) = -1
        End With
        cboSelectTime.ListIndex = 0
        
        
        With cboState
            .Clear
            .AddItem "全部状态"
            .AddItem "未确认"
            .AddItem "确认为非危急值"
            .AddItem "确认为是危急值"
            .ListIndex = 0
        End With
        
        
        mblnOK = False
        mintCurIndex = -1
        Call SetFaceCtrl
        Call SetFilterInfo
        Call InitTable
        Call MainDefCommandBar
         
        Call LoadPatients
    ElseIf mint显示方式 = 1 Then
        Me.BorderStyle = 3
        Me.Caption = "危急值选择(双击选择)"
        Me.Width = 5800
        Me.Height = 4900
        
        Call ShowAllCard
        
        '关于滚动条的显示
        If picCardCon.Height < picCItem(mlngCntCard).Top + picCItem(mlngCntCard).Height + 100 Then
            vscH.Visible = True
            vscH.value = 0
        Else
            vscH.Visible = False
        End If
        
        picCardCon.Height = picCItem(mlngCntCard).Top + picCItem(mlngCntCard).Height + 100
    
    End If
End Sub

Private Sub cboSelectTime_Click()
 
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zldatabase.Currentdate, "yyyy-MM-dd")
    If cboSelectTime.ListIndex = mintPreTim And intDateCount <> -1 Then Exit Sub
    If intDateCount = -1 Then
        If Not frmSelectTime.ShowMe(Me, mdatB登记, mdatE登记, cboSelectTime) Then
            '取消时恢复原来的选择
            Call Cbo.SetIndex(cboSelectTime.Hwnd, mintPreTim)
            Exit Sub
        End If
    Else
        mdatE登记 = datCurr
        mdatB登记 = mdatE登记 - intDateCount
    End If
    If mdatB登记 = CDate(0) Or mdatE登记 = CDate(0) Then
        cboSelectTime.ToolTipText = ""
    Else
        cboSelectTime.ToolTipText = "范围：" & Format(mdatB登记, "yyyy-MM-dd") & " 至 " & Format(mdatE登记, "yyyy-MM-dd")
    End If
    mintPreTim = cboSelectTime.ListIndex
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picFilter.Hwnd
    ElseIf Item.ID = 2 Then
'        Item.Handle = picPati.Hwnd
    End If
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim strFunName As String

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False) '固有
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_AppCritical, "查看危急值单(&D)")
            objControl.IconId = 3031
        If mint类型 = 2 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "危急值确认")
        End If
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
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
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_AppCritical, "查看危急值单(&D)")
            objControl.IconId = 3031
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True '固有
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_AppCritical, "查看危急值单") '固有
            objControl.IconId = 3031
            objControl.Style = xtpButtonIconAndCaption
            
        If mint类型 = 2 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "危急值确认")
                objControl.Style = xtpButtonIconAndCaption
        End If
        
        If mint类型 = 3 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
                objControl.BeginGroup = True
                objControl.Style = xtpButtonIconAndCaption
        End If
            
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
            objControl.Style = xtpButtonIconAndCaption
            objControl.IconId = 191
            objControl.BeginGroup = True
    End With
     
    objControl.Style = xtpButtonIconAndCaption
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
    End With
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngH As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    
    If mint类型 = 2 Then
        picFilter.Top = lngTop
        picFilter.Height = 400
        picFilter.Width = lngRight - lngLeft
        picFilter.Left = lngLeft
        With Me.fraPati
            .Left = lngLeft: .Top = lngTop - 60 + 400
            .Width = lngRight - lngLeft
        End With
        With Me.tbcSub
            .Left = lngLeft: .Width = lngRight - lngLeft
            .Top = lngTop + fraPati.Height: .Height = lngBottom - lngTop - fraPati.Height - IIF(Me.stbThis.Visible, stbThis.Height, 0)
        End With
    ElseIf mint类型 = 3 Then
        lngH = 1150
        picPatiC.Move lngLeft, lngTop, lngRight - lngLeft, lngH
        fraPati.Move 0, -60, picPatiC.Width, lngH + 60
        
        
        With Me.tbcSub
            .Left = lngLeft: .Width = lngRight - lngLeft
            .Top = lngTop + picPatiC.Height: .Height = lngBottom - lngTop - picPatiC.Height - IIF(Me.stbThis.Visible, stbThis.Height, 0)
        End With
        
    End If
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    imgPatient.Top = 10
    imgPatient.Left = 10
    imgPatient.Height = 1000
    
    picInfo.Left = 30
    picInfo.Top = 100
    picInfo.Width = fraPati.Width - 130
    picInfo.Height = imgPatient.Height + 30
    fraPati.Height = 1200
    
    lblInfo(e标识号).Left = lblInfo(e姓名).Left
    lblInfo(e科室).Left = lblInfo(e性别).Left
    lblInfo(e标识号).Top = 800
    lblInfo(e科室).Top = 800
    
    If mint显示方式 = 1 Then
        picCardFra.Move 0, 0, Me.Width - 80, Me.Height - 430
        picCardFra.ZOrder 0
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng病人ID = 0
    mlng主页ID = 0
    mstr挂号单 = ""
    mlng就诊ID = 0
    Call UnloadControls
    Set mrsCard = Nothing
    mstrPrePati = ""
    Set mobjCISJob = Nothing
    If mint显示方式 = 0 Then
        Call SaveWinState(Me, App.ProductName)
    End If
End Sub

Private Sub lblText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCItem(Index).Hwnd, lblText(Index).Caption, True
End Sub

Private Sub imgCL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCItem(Index).Hwnd, imgCL(Index).Tag, True
End Sub

Private Sub imgWJ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCItem(Index).Hwnd, imgWJ(Index).Tag, True
End Sub

Private Sub lblName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblAge_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblSex_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblTime_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub imgCL_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub imgWJ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCItem_MouseDown(Index, Button, Shift, X, Y)
End Sub


Private Sub lblName_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub lblAge_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub lblSex_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub lblTime_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub lblText_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub lblSelect_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub imgCL_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub imgWJ_DblClick(Index As Integer)
    Call picCItem_DblClick(Index)
End Sub

Private Sub optInfo_Click(Index As Integer)
    If Index = 0 Then '门诊
    
    ElseIf Index = 1 Then '住院
    
    ElseIf Index = 2 Then '外来
    
    End If
End Sub

Private Sub picCardFra_Resize()
    On Error Resume Next
    picCardCon.Move 0, 0, picCardFra.Width - vscH.Width, picCardFra.Height
    vscH.Left = picCardCon.Width
    vscH.Top = 0
    vscH.Height = picCardFra.Height
 
    '调用界面卡片适应
    Call ReSetCardPos
End Sub

Private Sub picCItem_DblClick(Index As Integer)
'功能：显示卡片
    Call ShowCardPop
End Sub

Private Sub picCItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
    If mintCurIndex > 0 Then
        '清除上一个的选择
        lblSelect(mintCurIndex).Visible = False
    End If
    mintCurIndex = Index
    Call ShowSelect
End Sub

Private Sub picCItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCItem(Index).Hwnd, ""
End Sub

Private Sub picCritical_Resize()
    On Error Resume Next
    vsCritical.Move 0, 0, picCritical.Width, picCritical.Height
End Sub

Private Sub picFilter_Resize()
    Dim lngTmp As Long
    
    On Error Resume Next
    
    If mint类型 = 2 Then
        '医技站调用
        lblRegDept.Top = 120
        lblRegDept.Left = 60
        
        lblRegTime.Left = 60
        lblRegTime.Top = 450
        
        Call zlControl.SetPubCtrlPos(False, 0, lblRegDept, 1200, lblRegTime, 60, cboSelectTime, 80, cmdOK)
    ElseIf mint类型 = 3 Then
        lblLX.Top = 120
        lblLX.Left = 60
        
        lngTmp = 200
        
        Call zlControl.SetPubCtrlPos(True, 0, lblLX, lngTmp, lblPatiDept, lngTmp, lblRegDept, lngTmp, lblState, lngTmp, lblRegTime)
 
        Call zlControl.SetPubCtrlPos(False, 0, lblLX, 100, picLX)
        
        Call zlControl.SetPubCtrlPos(False, 0, lblPatiDept, 100, cboPatiDept)
        
        Call zlControl.SetPubCtrlPos(False, 0, lblRegDept, 100, cboRegDept)
        
        Call zlControl.SetPubCtrlPos(False, 0, lblState, 100, cboState)
        
        Call zlControl.SetPubCtrlPos(False, 0, lblRegTime, 100, cboSelectTime, 80, cmdOK)
        
    End If
End Sub
 
Private Function LoadPatients() As Boolean
'功能：加载病人列表
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngSelectRow As Long
    Dim i As Long
    Dim datETmp As Date
    Dim strWhere As String
    Dim lng确认科室ID As Long
    Dim lng登记科室ID As Long
    
    On Error GoTo errH
    
    datETmp = Format(mdatE登记, "yyyy-MM-dd 23:59:59")
    
    If mint类型 = 2 Then
        strSql = "select rownum as 序号,a.id,a.病人id,a.主页id,a.挂号单,a.医嘱ID,a.状态,a.是否危急值,a.姓名,a.危急值描述,a.性别,a.年龄,a.报告时间 from 病人危急值记录 a" & _
            " where a.报告科室id = [1] And a.报告时间 Between [2] And [3] order by a.报告时间 desc "
        Set mrsCard = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlng科室ID, mdatB登记, datETmp)
        mblnRefreshCard = True
    End If
    
    If mint类型 = 3 Then
        If optInfo(0).value Then
            strWhere = " and a.挂号单 is not null "
        ElseIf optInfo(1).value Then
            strWhere = " and nvl(a.主页id,0)>0 "
        ElseIf optInfo(2).value Then
            strWhere = " and nvl(a.主页id,0)=0 and  a.挂号单  is  null"
        End If
        
        If cboPatiDept.ListIndex >= 0 Then
            '确认科室
            If cboPatiDept.ItemData(cboPatiDept.ListIndex) <> 0 Then
                lng确认科室ID = cboPatiDept.ItemData(cboPatiDept.ListIndex)
                strWhere = strWhere & " and a.确认科室ID =[1] "
            End If
        End If
        
        If cboRegDept.ListIndex >= 0 Then
            '登记科室
            If cboRegDept.ItemData(cboRegDept.ListIndex) <> 0 Then
                lng登记科室ID = cboRegDept.ItemData(cboRegDept.ListIndex)
                strWhere = strWhere & " and a.报告科室ID =[2] "
            End If
        End If
        
        
        If cboState.ListIndex >= 0 Then
            Select Case cboState.ListIndex
            Case 0
            Case 1
                strWhere = strWhere & " and a.状态=1 "
            Case 2
                strWhere = strWhere & " and a.状态=2 and nvl(a.是否危急值,0)=0 "
            Case 3
                strWhere = strWhere & " and a.状态=2 and nvl(a.是否危急值,0)=1 "
            End Select
        End If
        
        
        strSql = "select rownum as 序号,a.id,a.病人id,a.主页id,a.挂号单,a.医嘱ID,a.状态,a.是否危急值,a.姓名,a.危急值描述,a.性别,a.年龄,a.报告时间 from 病人危急值记录 a" & _
            " where a.报告时间 Between [3] And [4] " & strWhere & " order by a.报告时间 desc "
        Set mrsCard = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng确认科室ID, lng登记科室ID, mdatB登记, datETmp)
        mblnRefreshCard = True
        
        Call LoadCritical
    End If
     
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadCritical() As Boolean
'功能：加载危急值列表
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim str结果 As String
    Dim lng确认科室ID As Long
    Dim lng登记科室ID As Long
    Dim datETmp As Date
    Dim strWhere As String
    
    On Error GoTo errH
    If mint类型 <> 3 Then Exit Function
    
    datETmp = Format(mdatE登记, "yyyy-MM-dd 23:59:59")
    
    If optInfo(0).value Then
        strWhere = " and a.挂号单 is not null "
    ElseIf optInfo(1).value Then
        strWhere = " and nvl(a.主页id,0)>0 "
    ElseIf optInfo(2).value Then
        strWhere = " and nvl(a.主页id,0)=0 and  a.挂号单  is  null"
    End If
    
    If cboPatiDept.ListIndex >= 0 Then
        '确认科室
        If cboPatiDept.ItemData(cboPatiDept.ListIndex) <> 0 Then
            lng确认科室ID = cboPatiDept.ItemData(cboPatiDept.ListIndex)
            strWhere = strWhere & " and a.确认科室ID =[1] "
        End If
    End If
    
    If cboRegDept.ListIndex >= 0 Then
        '登记科室
        If cboRegDept.ItemData(cboRegDept.ListIndex) <> 0 Then
            lng登记科室ID = cboRegDept.ItemData(cboRegDept.ListIndex)
            strWhere = strWhere & " and a.报告科室ID =[2] "
        End If
    End If
    
    If cboState.ListIndex >= 0 Then
        Select Case cboState.ListIndex
        Case 0
        Case 1
            strWhere = strWhere & " and a.状态=1 "
        Case 2
            strWhere = strWhere & " and a.状态=2 and nvl(a.是否危急值,0)=0 "
        Case 3
            strWhere = strWhere & " and a.状态=2 and nvl(a.是否危急值,0)=1 "
        End Select
    End If
    
    strSql = "select  a.id,a.危急值描述,a.报告时间,a.报告人,a.处理情况,a.确认时间,a.确认人,a.确认科室id,b.名称 as 确认科室,a.状态,a.医嘱id,a.是否危急值  from 病人危急值记录 a,部门表 b" & _
        " where a.确认科室id=b.id(+) and  a.报告时间 Between [3] And [4] " & strWhere & " order by a.报告时间 desc "
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng确认科室ID, lng登记科室ID, mdatB登记, datETmp)
 
    With vsCritical
        .Redraw = flexRDNone
        .Rows = 1
        .ExplorerBar = 7
        If rsTmp.RecordCount > 0 Then
            i = 1
            Do While Not rsTmp.EOF
                .AddItem ""
                .TextMatrix(i, col危急值描述) = rsTmp!危急值描述 & ""
                .TextMatrix(i, col报告时间) = Format(rsTmp!报告时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, col报告人) = rsTmp!报告人 & ""
                .TextMatrix(i, col处理情况) = rsTmp!处理情况 & ""
                
                If Not IsNull(rsTmp!确认时间) Then
                    .TextMatrix(i, col确认时间) = Format(rsTmp!确认时间, "yyyy-MM-dd HH:mm")
                End If
                
                .TextMatrix(i, col确认人) = rsTmp!确认人 & ""
                .TextMatrix(i, col确认科室) = rsTmp!确认科室 & ""
                
                If Val(rsTmp!状态 & "") = 2 Then
                    If Val(rsTmp!是否危急值 & "") = 1 Then
                        .TextMatrix(i, col结果) = "是危急值"
                    Else
                        .TextMatrix(i, col结果) = "不是危急值"
                    End If
                End If
                    
                .TextMatrix(i, colID) = Val(rsTmp!ID & "")
                .TextMatrix(i, col状态) = Val(rsTmp!状态 & "")
                .TextMatrix(i, col医嘱ID) = Val(rsTmp!医嘱ID & "")
                i = i + 1
                rsTmp.MoveNext
            Loop
        Else
            .AddItem ""
        End If
        .Redraw = flexRDDirect
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitTable()
    Dim arrHead As Variant, i As Long
    Dim strHead As String
    strHead = "危急值描述,2500,1;报告时间,1800,1;报告人,700,1;处理情况,2000,1;确认时间,1800,1;确认人,700,1;确认科室,800,1;结果,800,1;ID;状态;医嘱ID"
    arrHead = Split(strHead, ";")
    With vsCritical
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub
 
Private Function ReadPatPricture(ByVal lng病人ID As Long, ByRef imgPatient As Image, Optional ByRef strFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人照片
    '参数：lng病人ID=读取指定病人的照片
    '           imgPatient=照片加载位置
    '           strFile=照片的本地路径
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    imgPatient.Picture = Nothing
    strFile = ""
    strFile = sys.Readlob(glngSys, 27, lng病人ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = LoadPicture(strFile)
        ReadPatPricture = True
        Kill strFile
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ClearPatiInfo()
'功能：清除病人信息
    lblInfo(e姓名).Caption = "姓名"
    lblInfo(e性别).Caption = "性别"
    lblInfo(e年龄).Caption = "年龄"
    lblInfo(e标识号).Caption = "标识号"
    lblInfo(e科室).Caption = "科室"
    imgPatient.Picture = imgDefual.Picture
End Sub

Private Sub vscH_Change()
'
    Dim lngMove As Long
    Dim lngY As Long
    If Not vscH.Visible Then Exit Sub
    '计算单步步长
    lngMove = CLng((picCItem(mlngCntCard).Top + picCItem(mlngCntCard).Height + 100 - picCardFra.Height) / 100)

    If lngMove < 0 Then lngMove = 0
    lngY = -1 * vscH.value * lngMove
    If lngY >= 0 And lngY < 100 Then lngY = 0
    
    picCardCon.Top = lngY
    
End Sub
 
Private Sub vsCritical_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    
    With vsCritical
        If Val(.TextMatrix(NewRow, colID)) <> 0 Then
            For i = 1 To mlngCntCard
                If Val(.TextMatrix(NewRow, colID)) = Val(lblName(i).Tag) Then
                    
                    If mintCurIndex > 0 Then
                        '清除上一个的选择
                        lblSelect(mintCurIndex).Visible = False
                    End If
                    mintCurIndex = i
                    Call ShowSelect
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsCritical_DblClick()
'功能：双击为查看危急值单
    Dim i As Long
    
    With vsCritical
        If Val(.TextMatrix(.Row, colID)) <> 0 Then
            For i = 1 To mlngCntCard
                If Val(.TextMatrix(.Row, colID)) = Val(lblName(i).Tag) Then
                    
                    If mintCurIndex > 0 Then
                        '清除上一个的选择
                        lblSelect(mintCurIndex).Visible = False
                    End If
                    mintCurIndex = i
                    Call ShowSelect
                    Exit For
                End If
            Next
        End If
    End With
    
    Call ShowCardPop
    
'    Dim lng记录ID As Long
'    Dim lng医嘱ID As Long
'    Dim int调用类型 As Integer
'    Dim lng病人ID As Long
'    Dim lng主页ID As Long
'    Dim str挂号单 As String
'    Dim str危急指标 As String
'    Dim str危急结果 As String
'
'    If rptPati.SelectedRows.Count = 0 Then Exit Sub          '非正常情况
'    With rptPati.SelectedRows(0)
'        If Not .GroupRow Then
'            lng病人ID = Val(.Record(COL_病人ID).value)
'            lng主页ID = Val(.Record(COL_主页ID).value)
'            str挂号单 = .Record(COL_挂号单).value
'        End If
'    End With
'
'    If lng病人ID = 0 Then
'        MsgBox "请选择一个病人。", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'    With vsCritical
'        lng记录ID = Val(.TextMatrix(.Row, colID))
'        lng医嘱ID = Val(.TextMatrix(.Row, col医嘱ID))
'    End With
'    If lng记录ID = 0 Then
'        MsgBox "请选择一条危急值记录。", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'
'    If str挂号单 = "" Then
'        int调用类型 = 2
'    Else
'        int调用类型 = 1
'    End If
'
'    Call frmCriticalEdit.ShowMe(Me, True, 2, int调用类型, lng病人ID, lng主页ID, str挂号单, 0, lng记录ID, lng医嘱ID)
End Sub

Private Sub DeleteData()
'功能：删除危急值记录
    Dim strSql As String
    Dim lngID As Long
    
    Select Case tbcSub.Selected.Tag
    Case "危急值"
    
    lngID = Val(vsCritical.TextMatrix(vsCritical.Row, colID))
    
    strSql = "zl_病人危急值记录_delete(" & lngID & ")"
    Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
    Call vsCritical.RemoveItem(vsCritical.Row)
    
    Case "明细卡"
        lngID = Val(lblName(mintCurIndex).Tag)
        strSql = "zl_病人危急值记录_delete(" & lngID & ")"
        Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
        Call LoadPatients
        Call ShowAllCard
    End Select
    mblnOK = True
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub EditData(ByVal intType As Integer)
'功能：修改或者查看记录
'参数：intType 1-修改，2-查看
    If tbcSub.Selected.Tag = "明细卡" Or tbcSub.Selected.Tag = "危急值" Then
        Call ShowCardBybnt(intType)
        Exit Sub
    End If
End Sub

Private Sub ShowAllCard()
'功能：显示卡片
    Dim i As Long
    
    mintCurIndex = -1
    mlngCntCard = mrsCard.RecordCount
    
    Call LoadAllCard
    Call LocatePati
  
    stbThis.Panels(2).Text = "一共" & mlngCntCard & "条危值信息！"
End Sub

Private Sub UnloadControls()
'功能：卸载控件
    Dim i As Long
    Dim lngCnt As Long
    
    lngCnt = picCItem.Count - 1
    
    For i = lngCnt To 1 Step -1
        Unload imgWJ(i)
        Unload imgCL(i)
        
        Unload lblName(i)
        Unload lblAge(i)
        Unload lblSex(i)
        Unload lblTime(i)
        Unload lblSelect(i)
        Unload lblText(i)
        Unload picCItem(i)
    Next
    
End Sub

Private Sub LoadOneCard(ByVal lngIdx As Long, ByVal lngX As Long, lngY As Long)
'功能：新增一张卡片

    Load picCItem(lngIdx)
    Set picCItem(lngIdx).Container = picCardCon
    picCItem(lngIdx).Visible = True
    picCItem(lngIdx).Picture = imgCardBack(1).Picture
    picCItem(lngIdx).Top = lngY
    picCItem(lngIdx).Left = lngX
    
    Load lblText(lngIdx)
    Set lblText(lngIdx).Container = picCItem(lngIdx)
    lblText(lngIdx).Visible = True
    
    Load lblName(lngIdx)
    Set lblName(lngIdx).Container = picCItem(lngIdx)
    lblName(lngIdx).Visible = True
    
    
    Load lblAge(lngIdx)
    Set lblAge(lngIdx).Container = picCItem(lngIdx)
    lblAge(lngIdx).Visible = True
    
    Load lblSex(lngIdx)
    Set lblSex(lngIdx).Container = picCItem(lngIdx)
    lblSex(lngIdx).Visible = True
    
    
    Load lblTime(lngIdx)
    Set lblTime(lngIdx).Container = picCItem(lngIdx)
    lblTime(lngIdx).Visible = True
    
     
    Load lblSelect(lngIdx)
    Set lblSelect(lngIdx).Container = picCItem(lngIdx)
    lblSelect(lngIdx).Visible = False
        
    Load imgWJ(lngIdx)
    Set imgWJ(lngIdx).Container = picCItem(lngIdx)
    imgWJ(lngIdx).Visible = False
    
    Load imgCL(lngIdx)
    Set imgCL(lngIdx).Container = picCItem(lngIdx)
    imgCL(lngIdx).Visible = False
    
End Sub

Private Sub ResiceCard(ByVal lngIdx As Long)
'功能：设置卡片内部控件位置
    
    lblSelect(lngIdx).Left = 140
    lblSelect(lngIdx).Width = 1510
    lblSelect(lngIdx).Top = 600
    lblName(lngIdx).Top = 660
    lblName(lngIdx).Left = 160
    
    
    lblSex(lngIdx).Left = lblName(lngIdx).Left
    lblSex(lngIdx).Top = lblName(lngIdx).Top + lblName(lngIdx).Height + 120
    
    lblAge(lngIdx).Left = lblSex(lngIdx).Left + lblSex(lngIdx).Width + 300
    lblAge(lngIdx).Top = lblSex(lngIdx).Top
    
    
    lblText(lngIdx).Left = lblName(lngIdx).Left
    lblText(lngIdx).Width = lblSelect(lngIdx).Width
    lblText(lngIdx).Top = lblAge(lngIdx).Top + lblAge(lngIdx).Height + 120
    
    lblTime(lngIdx).Left = 750
    
    lblTime(lngIdx).Top = lblText(lngIdx).Top + lblText(lngIdx).Height + 300
    
    imgCL(lngIdx).Left = lblName(lngIdx).Left
    imgCL(lngIdx).Top = 300
    
    imgWJ(lngIdx).Left = imgCL(lngIdx).Left + imgCL(lngIdx).Width + 10
    imgWJ(lngIdx).Top = imgCL(lngIdx).Top
    
End Sub

Private Sub LoadAllCard()
'功能：显示所有卡片
    Dim lngX As Long, lngY As Long
    Dim i As Long
    Dim lngRowCount As Long
    
    lngX = clngX
    lngY = clngX
    
    lngRowCount = (picCardCon.Width) \ (picCItem(0).Width)
    mlngPreRowCnt = lngRowCount
    Call UnloadControls
    mrsCard.Filter = 0
    If mlngCntCard = 0 Then Exit Sub
    mrsCard.MoveFirst
    For i = 1 To mrsCard.RecordCount
        Call LoadOneCard(i, lngX, lngY)
        Call SetCardData(i, mrsCard)
        Call ResiceCard(i)
        '计算下一张卡片的坐标
        lngX = lngX + picCItem(i).Width
        
        If i Mod lngRowCount = 0 Then
            lngX = clngX
            lngY = lngY + picCItem(i).Height
        End If
        mrsCard.MoveNext
    Next
    
End Sub

Private Sub ReSetCardPos()
'功能：重新排列卡片的位置
    Dim lngX As Long, lngY As Long
    Dim i As Long
    Dim lngRowCount As Long
    
    lngX = clngX
    lngY = clngX
    
    '如果无卡片则退出
    If mlngCntCard = 0 Then
        Exit Sub
    End If
    
    lngRowCount = (picCardCon.Width) \ (picCItem(0).Width)
    
    If mlngPreRowCnt = lngRowCount Then
        '如果每一行的卡片数据不变则不用调位置
        Exit Sub
    End If
    
    lngX = clngX
    lngY = clngX
    
    For i = 1 To mlngCntCard
        
        picCItem(i).Top = lngY
        picCItem(i).Left = lngX
    
        '计算下一张卡片的坐标
        lngX = lngX + picCItem(0).Width
        If i Mod lngRowCount = 0 Then
            lngX = clngX
            lngY = lngY + picCItem(0).Height
        End If
    Next
    
    '关于滚动条的显示
    If picCardCon.Height < picCItem(mlngCntCard).Top + picCItem(mlngCntCard).Height + 100 Then
        vscH.Visible = True
        vscH.value = 0
    Else
        vscH.Visible = False
    End If
    
    picCardCon.Height = picCItem(mlngCntCard).Top + picCItem(mlngCntCard).Height + 100
End Sub

Private Sub timeRefreshCard_Timer()
  
    If Not mblnRefreshCard Then Exit Sub
    mblnRefreshCard = False
    timeRefreshCard.Enabled = False
    Call ShowAllCard
    timeRefreshCard.Enabled = True
End Sub

Private Sub ShowSelect()
'功能：选中卡片
    lblSelect(mintCurIndex).Visible = True
    
    If mint显示方式 = 0 Then
        mlngPreCardID = Val(lblName(mintCurIndex).Tag)
        Call LoadPatiInfobyCard
    End If
End Sub

Private Sub SetCardData(ByVal lngIdx As Long, ByVal rsData As ADODB.Recordset)
'功能：加载卡片上的信息
    lblText(lngIdx).Caption = rsData!危急值描述 & ""
    lblName(lngIdx).Caption = rsData!姓名 & ""
    lblName(lngIdx).Tag = rsData!ID & "" '---关键信息
    lblSex(lngIdx).Caption = rsData!性别 & ""
    lblAge(lngIdx).Caption = rsData!年龄 & ""
    lblTime(lngIdx).Caption = Format(rsData!报告时间, "yyyy/MM/dd")
    If Val(rsData!状态 & "") = 2 Then
        imgCL(lngIdx).Visible = True
        Set imgCL(lngIdx).Picture = imgCL(0).Picture
        imgCL(lngIdx).Tag = "已处理"
        
        If Val(rsData!是否危急值 & "") = 1 Then
            imgWJ(lngIdx).Visible = True
            Set imgWJ(lngIdx).Picture = imgWJ(0).Picture
            imgWJ(lngIdx).Tag = "确认是危急值"
        End If
    End If
    '--------以上部分为卡片选择器中提供的字段 mint显示方式=1
    If mint显示方式 = 0 Then
        '其它信息，
    End If
End Sub

Private Sub ShowCardPop()
'功能：弹出登记单
    
    If mint显示方式 = 1 Then
        mlng记录ID = Val(lblName(mintCurIndex).Tag)
        
        Unload Me
    Else
        If mintCurIndex > 0 Then
            mrsCard.Filter = "ID=" & Val(lblName(mintCurIndex).Tag)
            If Not mrsCard.EOF Then
                Call frmCriticalEdit.ShowMe(Me, True, 2, IIF(IsNull(mrsCard!挂号单), 2, 1), _
                    Val(mrsCard!病人ID & ""), Val(mrsCard!主页ID & ""), mrsCard!挂号单 & "", 0, Val(mrsCard!ID & ""), Val(mrsCard!医嘱ID & ""))
            End If
        End If
    End If
End Sub

Private Sub ShowCardBybnt(ByVal intType As Integer)
'功能：查看登记单
    Dim blnOK As Boolean
    If mintCurIndex > 0 Then
        mrsCard.Filter = "ID=" & Val(lblName(mintCurIndex).Tag)
        If Not mrsCard.EOF Then
            blnOK = frmCriticalEdit.ShowMe(Me, True, intType, IIF(IsNull(mrsCard!挂号单), 2, 1), _
                Val(mrsCard!病人ID & ""), Val(mrsCard!主页ID & ""), mrsCard!挂号单 & "", 0, Val(mrsCard!ID & ""), Val(mrsCard!医嘱ID & ""))
            If blnOK Then
                mblnOK = True
                If intType = 1 Then
                    Call LoadPatients
                    mblnRefreshCard = True
                End If
            End If
        End If
    End If
End Sub

Private Sub LoadPatiInfobyCard()
'功能：切换卡片时显示病人信息

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    On Error GoTo errH
    If mintCurIndex > 0 Then
        mrsCard.Filter = "ID=" & Val(lblName(mintCurIndex).Tag)
        If Not mrsCard.EOF Then
            If mstrPrePati = Val(mrsCard!病人ID & "") & "," & Val(mrsCard!主页ID & "") & "," & mrsCard!挂号单 Then
                Exit Sub
            End If
            mstrPrePati = Val(mrsCard!病人ID & "") & "," & Val(mrsCard!主页ID & "") & "," & mrsCard!挂号单
            If mrsCard!挂号单 & "" <> "" Then
                strSql = "select a.id as 就诊ID, a.门诊号 as 标识号,b.名称 as 科室 from  病人挂号记录 a,部门表 b where a.执行部门id=b.id and a.no=[1]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mrsCard!挂号单 & "")
            ElseIf Val(mrsCard!主页ID & "") <> 0 Then
                strSql = "select a.主页id as 就诊ID, a.住院号 as 标识号,b.名称 as 科室  from 病案主页 a,部门表 b where a.出院科室id=b.id and a.病人id=[1] and a.主页id=[2]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsCard!病人ID & ""), Val(mrsCard!主页ID & ""))
            Else
                '外来病人
                strSql = "select 0 as 就诊ID, null as 标识号,b.名称 as 科室 from 病人医嘱记录 a,部门表 b where a.病人科室ID=b.id and a.id=[1]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsCard!医嘱ID & ""))
            End If
            
            mlng就诊ID = Val(rsTmp!就诊ID & "")
            
            lblInfo(e姓名).Caption = lblName(mintCurIndex).Caption
            lblInfo(e性别).Caption = lblSex(mintCurIndex).Caption
            lblInfo(e年龄).Caption = lblAge(mintCurIndex).Caption
            
            If Val(mrsCard!主页ID & "") = 0 Then
                strTmp = "门诊号:" & rsTmp!标识号
            Else
                strTmp = "住院号:" & rsTmp!标识号
            End If
            lblInfo(e标识号).Caption = strTmp
            lblInfo(e科室).Caption = "科室:" & rsTmp!科室
            
            Call ReadPatPricture(Val(mrsCard!病人ID & ""), imgLoad)
            If imgLoad.Picture = 0 Then
                imgPatient.Picture = imgDefual.Picture
            Else
                imgPatient.Picture = imgLoad.Picture
            End If
            mlng病人ID = Val(mrsCard!病人ID & "")
            
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LocatePati()
'功能：缺省定位病人
    Dim i As Long
    
    If mlngPreCardID = 0 Then
        Exit Sub
    End If
    Call ClearPatiInfo
    For i = 1 To mlngCntCard
        If Val(lblName(i).Tag) = mlngPreCardID Then
            mintCurIndex = i
            mstrPrePati = ""
            Call ShowSelect
            Exit For
        End If
    Next
End Sub

Private Sub SetFaceCtrl()
'功能：设置界面的控件可见性
    If mint类型 = 2 Then
        lblPatiDept.Visible = False
        cboPatiDept.Visible = False
        cboRegDept.Visible = False
    ElseIf mint类型 = 3 Then
        picPatiC.Visible = True
        Set fraPati.Container = picPatiC
    End If
End Sub

Private Sub SetFilterInfo()
'功能：初始化过滤条件
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errH
    If mint类型 = 2 Then
        strTmp = sys.RowValue("部门表", mlng科室ID, "名称")
        lblRegDept.Caption = "登记科室:" & strTmp
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
 
    lngCur = vscH.value
    lngMin = vscH.Min
    lngMax = vscH.Max

    If KeyCode = vbKeyPageDown Then '下
        If Between(lngCur + (lngMax - lngMin) / 100, lngMin, lngMax) Then
            vscH.value = lngCur + (lngMax - lngMin) / 100
        Else
            vscH.value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '上
        If Between(lngCur - (lngMax - lngMin) / 100, lngMin, lngMax) Then
            vscH.value = lngCur - (lngMax - lngMin) / 100
        Else
            vscH.value = lngMin
        End If
    End If
 
End Sub

Private Function Init确认科室() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim lngPreDept As Long
    
    If cboPatiDept.ListIndex <> -1 Then
        lngPreDept = cboPatiDept.ItemData(cboPatiDept.ListIndex)
    End If
    cboPatiDept.Clear
    cboPatiDept.AddItem "所有科室"
    cboPatiDept.ItemData(cboPatiDept.NewIndex) = 0
    On Error GoTo errH
    Set rsTmp = GetDataToDepts
    
    For i = 1 To rsTmp.RecordCount
        cboPatiDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboPatiDept.ItemData(cboPatiDept.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPreDept Then '保留原有定位
            Call Cbo.SetIndex(cboPatiDept.Hwnd, cboPatiDept.NewIndex)
        ElseIf InStr(mstrPrivs, "全院病人") > 0 Then
            If UserInfo.部门ID = rsTmp!ID And (lngPreDept = 0 Or cboPatiDept.ListIndex = -1) Then '直接所属优先
                Call Cbo.SetIndex(cboPatiDept.Hwnd, cboPatiDept.NewIndex)
            End If
        Else
            '所属缺省病区包含的可能有多个
            If rsTmp!缺省 = 1 And cboPatiDept.ListIndex = -1 Then
                Call Cbo.SetIndex(cboPatiDept.Hwnd, cboPatiDept.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboPatiDept.ListIndex = -1 And cboPatiDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboPatiDept.Hwnd, 0)
    End If
    Init确认科室 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetDataToDepts() As ADODB.Recordset
'功能：获取科室病区列表数据记录集
'参数：strIn 过滤条件
    Dim strSql As String
    Dim strDeptIDs As String
    If optInfo(1).value Then
        '按科室读取显示
        '包含门急诊观察室的病人还没有上床，不加只显床上有病人的科室的限制
        If InStr(mstrPrivs, "全院病人") > 0 Then
            strSql = _
                " Select Distinct A.ID,A.编码,A.名称" & _
                " From 部门表 A,部门性质说明 B" & _
                " Where B.部门ID=A.ID And B.工作性质='临床'" & _
                " And ((B.服务对象 IN(2,3) " & _
                ")Or (B.服务对象=1 And Exists(Select 1 From 床位状况记录 C Where B.部门ID = C.科室ID)))" & _
                " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编码"
        Else
            '求有权限的科室：本身所在科室+所属病区包含的科室
            strSql = _
                " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
                " From 部门表 A,部门性质说明 B,部门人员 C" & _
                " Where B.部门ID=A.ID And A.ID=C.部门ID And C.人员ID=[1]" & _
                " And (B.服务对象 IN(2,3) Or (B.服务对象=1 And Exists(Select 1 From 床位状况记录 C Where B.部门ID = C.科室ID)))" & _
                " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And B.工作性质='临床'"
            strSql = strSql & " Union " & _
                " Select C.ID,C.编码,C.名称,Nvl(A.缺省,0) As 缺省" & _
                " From 部门人员 A,病区科室对应 B,部门表 C" & _
                " Where A.部门ID=B.病区ID And B.科室ID=C.ID And A.人员ID=[1]" & _
                " And Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=B.病区ID)" & _
                " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=B.病区ID)" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)"
            If InStr(mstrPrivs, "ICU病人") > 0 Then
                strSql = strSql & " Union " & _
                    " Select A.ID,A.编码,A.名称,0 As 缺省" & _
                    " From 部门表 A" & _
                    " Where Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='ICU')" & _
                    " And Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='临床')" & _
                    " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
            End If
            strSql = "Select ID,编码,名称,Max(缺省) As 缺省 From (" & strSql & ") Group By ID,编码,名称 Order by 编码"
        End If
    End If
    
    If Not optInfo(1).value Then
        strSql = "Select Distinct B.ID,B.编码,B.名称,A.缺省" & _
            " From 部门人员 A,部门表 B,部门性质说明 C" & _
            " Where A.部门ID=B.ID And B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
            " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
            " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And A.人员ID=[1]" & _
            " Order by B.编码"
    End If
    
    On Error GoTo errH
    
    Set GetDataToDepts = zldatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init登记科室() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim str科室IDs As String, str来源 As String
    
    On Error GoTo errH
    
    '包含门诊/住院医技科室
    str来源 = "3"
    If InStr(mstrPrivs, "门诊病人") > 0 And InStr(mstrPrivs, "住院病人") > 0 Then
        str来源 = "1,2,3"
    ElseIf InStr(mstrPrivs, "门诊病人") > 0 Then
        str来源 = "1,3"
    ElseIf InStr(mstrPrivs, "住院病人") > 0 Then
        str来源 = "2,3"
    End If
    If InStr(mstrPrivs, "所有科室") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And B.服务对象 IN(" & str来源 & ") And B.工作性质 IN('检查','检验','手术','治疗','营养')" & _
            " Order by A.编码"
    Else
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And B.服务对象 IN(" & str来源 & ") And B.工作性质 IN('检查','检验','手术','治疗','营养')" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
    End If
    
    cboRegDept.Clear
    cboRegDept.AddItem "所有科室"
    cboRegDept.ItemData(cboRegDept.NewIndex) = 0
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    str科室IDs = GetUser科室IDs
    For i = 1 To rsTmp.RecordCount
        cboRegDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboRegDept.ItemData(cboRegDept.NewIndex) = rsTmp!ID
        
        If rsTmp!ID = UserInfo.部门ID Then
            Call Cbo.SetIndex(cboRegDept.Hwnd, cboRegDept.NewIndex) '直接所属优先
        End If
        If InStr("," & str科室IDs & ",", "," & rsTmp!ID & ",") > 0 And cboRegDept.ListIndex = -1 Then
            Call Cbo.SetIndex(cboRegDept.Hwnd, cboRegDept.NewIndex)
        End If
        
        rsTmp.MoveNext
    Next
    If cboRegDept.ListIndex = -1 And cboRegDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboRegDept.Hwnd, 0)
    End If
        
    If cboRegDept.ListIndex <> -1 Then
        Call cboRegDept_Click  '同时对mstrDeptNode赋值
    End If
    Init登记科室 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FunAffirm()
'功能：危急值确认
    Dim lng危急值ID As Long
    Dim lng医嘱ID As Long
    Dim lng病人ID As Long
    Dim blnOK As Boolean
    
    On Error GoTo errH
    
    If mintCurIndex > 0 Then
        mrsCard.Filter = "ID=" & Val(lblName(mintCurIndex).Tag)
        If Not mrsCard.EOF Then
            lng危急值ID = Val(lblName(mintCurIndex).Tag)
            lng医嘱ID = Val(mrsCard!医嘱ID & "")
            lng病人ID = Val(mrsCard!病人ID & "")
            blnOK = frmCriticalEdit.ShowMe(Me, True, 3, 3, lng病人ID, 0, "", 0, lng危急值ID, lng医嘱ID)
            If blnOK Then
                Call LoadPatients
                Call ShowAllCard
                mblnOK = True
            End If
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
